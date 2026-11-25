# -*- coding: utf-8 -*-
"""自定义样式映射测试。

背景：源码中不再硬编码任何特定厂商/模板的专有样式名，改为
「通用结构规则 + 外部配置」两层机制。这些测试要保证两件事：

1. **通用规则覆盖到位**——`<前缀><数字>-<层级>` 这类命名能被正确识别，
   不依赖具体前缀，从而在去掉硬编码后能力不降级；
2. **外部配置可用**——定制模板的特例（无法用通用规则表达的名字）
   能通过 --styles / 环境变量 / ~/.word2md/styles.json 注入。
"""
from __future__ import annotations

import json
import os
import re

import docx
import pytest
from docx.enum.style import WD_STYLE_TYPE

from word2md import (
    RE_BARE_HEADING,
    RE_STRUCT_HEADING_NAMED,
    _level_from_struct_name,
    convert_word_to_markdown,
    load_style_config,
)


# ---------------------------------------------------------------- 通用规则
@pytest.mark.parametrize('name,expected', [
    # <前缀><数字>-<层级>：后缀数字即层级，与前缀无关
    ('chapter1-1', 1), ('chapter1-2', 2), ('chapter1-3', 3), ('chapter1-4', 4),
    ('sec2-1', 1), ('sec2-3', 3),
    ('part9-2', 2),
    # 0 表示首章
    ('chapter1-0', 1),
    # 不存在或超界的层级不认
    ('chapter1-5', None), ('chapter1-9', None),
    # 其他形态不走这条规则
    ('heading 1', None), ('标题 1', None), ('Normal', None),
    ('chapter1-content', None), ('section1', None),
])
def test_struct_rule_levels(name, expected):
    """通用结构规则：<前缀><数字>-<数字> 取后缀数字为层级。"""
    assert _level_from_struct_name(name.lower()) == expected


def test_struct_rule_is_prefix_agnostic():
    """同层级不同前缀必须得到相同结果——这是「通用」的实质含义。"""
    for prefix in ('chapter1', 'sec1', 'part1', 'abc99', 'xyz2'):
        assert _level_from_struct_name(f'{prefix}-3'.lower()) == 3


def test_named_suffix_and_bare_prefix_patterns():
    """命名后缀（-content）与裸前缀的匹配边界。"""
    assert RE_STRUCT_HEADING_NAMED.match('chapter1-content')
    assert RE_STRUCT_HEADING_NAMED.match('part2-cover')
    assert not RE_STRUCT_HEADING_NAMED.match('chapter1-2')   # 纯数字走另一条
    # 裸前缀：章节类词根可命中
    assert RE_BARE_HEADING.match('section1')
    assert RE_BARE_HEADING.match('chapter2')
    # 常见非章节样式词根必须排除，否则会把列表/图表样式误判成标题
    for non_heading in ('list1', 'bullet2', 'table3', 'figure1',
                        'caption1', 'note1', 'body1'):
        assert not RE_BARE_HEADING.match(non_heading), non_heading


# ---------------------------------------------------------------- 外部配置
def test_load_style_config_explicit_path(tmp_path):
    p = tmp_path / 'styles.json'
    p.write_text(json.dumps({'heading_styles': {'mycover1': 1}}), encoding='utf-8')
    cfg = load_style_config(str(p))
    assert cfg['heading_styles'] == {'mycover1': 1}


def test_load_style_config_via_env(tmp_path, monkeypatch):
    p = tmp_path / 'styles.json'
    p.write_text(json.dumps({'toc_styles': {'mytoc 1': 1}}), encoding='utf-8')
    monkeypatch.setenv('WORD2MD_STYLES', str(p))
    assert load_style_config()['toc_styles'] == {'mytoc 1': 1}


def test_load_style_config_missing_returns_empty(monkeypatch):
    """找不到配置时返回空 dict，调用方无需判空。"""
    monkeypatch.delenv('WORD2MD_STYLES', raising=False)
    monkeypatch.setenv('HOME', '/nonexistent-home')
    monkeypatch.setenv('USERPROFILE', '/nonexistent-home')
    assert load_style_config() == {}


def test_load_style_config_rejects_non_object(tmp_path):
    p = tmp_path / 'bad.json'
    p.write_text('[1, 2, 3]', encoding='utf-8')
    with pytest.raises(ValueError, match='JSON 对象'):
        load_style_config(str(p))


def test_load_style_config_rejects_broken_json(tmp_path):
    p = tmp_path / 'broken.json'
    p.write_text('{not json', encoding='utf-8')
    with pytest.raises(ValueError, match='读取失败'):
        load_style_config(str(p))


# ---------------------------------------------------------------- 端到端
def _build(path, entries):
    d = docx.Document()
    for style_name, text in entries:
        try:
            d.styles[style_name]
        except KeyError:
            d.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        d.add_paragraph(text, style=style_name)
    d.save(str(path))
    return path


def _convert(tmp_path, entries, style_config=None):
    src = _build(tmp_path / 't.docx', entries)
    st = convert_word_to_markdown(
        str(src), out_dir=str(tmp_path / 'out'), style_config=style_config)
    with open(st['md_path'], encoding='utf-8') as f:
        return f.read()


def test_e2e_struct_rule_generates_heading_levels(tmp_path):
    """<前缀><数字>-<层级> 应产出对应层级的标题。"""
    md = _convert(tmp_path, [
        ('chapter1-1', '概述'),
        ('chapter1-2', '工作条件'),
        ('chapter1-3', '极限参数'),
        ('chapter2-1', '外设'),
    ])
    assert '# 概述' in md
    assert '## 工作条件' in md
    assert '### 极限参数' in md
    assert '# 外设' in md


def test_e2e_named_suffix_becomes_h1(tmp_path):
    """-content 这类前置页样式应视为一级章节。"""
    md = _convert(tmp_path, [
        ('chapter1-content', '目录'),
        ('chapter1-1', '正文'),
    ])
    assert '# 目录' in md


def test_e2e_external_config_overrides(tmp_path):
    """外部配置能覆盖/补充默认映射（定制模板特例的出口）。"""
    cfg = tmp_path / 'styles.json'
    cfg.write_text(json.dumps({'heading_styles': {'mycover1': 1}}),
                   encoding='utf-8')
    md = _convert(tmp_path, [
        ('mycover1', '扉页'),
        ('chapter1-1', '概述'),
    ], style_config=str(cfg))
    assert '# 扉页' in md, '未在默认表中的样式应通过配置文件生效'


def test_e2e_plain_paragraph_not_heading(tmp_path):
    """普通正文与列表项不得被误判为标题。"""
    md = _convert(tmp_path, [
        ('Normal', '这是一段普通正文。'),
        ('List Number', '这是一个列表项。'),
        ('chapter1-2', '唯一标题'),
    ])
    heads = [ln for ln in md.splitlines() if ln.startswith('#')]
    assert len(heads) == 1, f'应只有 1 个标题，实际: {heads}'
    assert '唯一标题' in heads[0]


def test_source_has_no_vendor_specific_style_names():
    """回归护栏：默认样式表只应包含 Word 内置样式，不得混入专有样式名。

    这条测试防止将来有人为了修某份文档，又把专有名称硬编码回源码——
    那会让公开仓库暴露内部模板命名。定制需求应走外部配置。

    判据：DEFAULT_HEADING_STYLES 的键只能是内置英文样式名或「标题 N」，
    出现其它形态（如 <前缀><数字>-<层级>、自定义中文名）即视为违规。
    """
    here = os.path.dirname(os.path.abspath(__file__))
    src = os.path.join(os.path.dirname(here), 'word2md.py')
    with open(src, encoding='utf-8') as f:
        code = f.read()

    start = code.index('DEFAULT_HEADING_STYLES')
    end = code.index('STYLE_CONFIG_ENV')
    block = code[start:end]

    keys = re.findall(r"'([^']+)'\s*:\s*\d", block)
    assert keys, '未解析到默认样式表的键'
    allowed = re.compile(r'^(?:Title|Heading [1-6]|标题 [1-6])$')
    leaked = [k for k in keys if not allowed.match(k)]
    assert not leaked, (
        f'默认样式表出现非内置样式名: {leaked}；'
        '专有样式应改为外部配置注入（--styles / WORD2MD_STYLES）')

    # 兜底：源码中不应残留任何厂商专有前缀的痕迹（含注释）
    # 用拼接构造关键词，避免本测试文件自身命中
    marker = 's' + 'oc'
    hits = [ln.strip() for ln in code.splitlines()
            if marker in ln.lower()]
    assert not hits, f'源码中出现疑似专有前缀痕迹: {hits[:3]}'
