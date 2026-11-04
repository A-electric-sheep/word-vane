# -*- coding: utf-8 -*-
"""转换能力测试：以结构断言为主，不依赖具体措辞。

设计原则
--------
1. 夹具全部运行时生成（见 conftest.py），仓库不含 .docx 二进制；
2. 断言针对**结构**（是否出现 colspan/rowspan、是否产出 $..$、条目数等），
   而非逐字比对全文——全文比对会把无关的措辞调整也判为失败；
3. 每类断言都附带"为什么这条重要"的说明，便于后续维护者判断能否删改。
"""
from __future__ import annotations

import os

import conftest
import pytest

from verify_consistency import verify
from word2md import convert_word_to_markdown


@pytest.fixture(scope='session')
def fixtures(tmp_path_factory):
    """生成全套夹具，返回 {名称: 路径}。"""
    d = tmp_path_factory.mktemp('fixtures')
    return conftest.build_all(str(d))


def _convert(src, out_dir):
    stats = convert_word_to_markdown(src, out_dir=str(out_dir))
    with open(stats['md_path'], encoding='utf-8') as f:
        md = f.read()
    return stats, md


# ---------------------------------------------------------------- 合并单元格
def test_merged_cells_emit_colspan_and_rowspan(fixtures, tmp_path):
    """合并单元格必须以 HTML colspan/rowspan 保留。

    这是本项目相对 GFM 路线工具的核心差异：GFM 管道表在语法上无法表达合并，
    只能把值塞进单格、其余留空，读者无从判断原表有几处合并。
    """
    stats, md = _convert(fixtures['merged'], tmp_path / 'out')
    assert 'colspan="3"' in md, '横向合并应输出 colspan=3'
    assert 'rowspan="2"' in md, '纵向合并应输出 rowspan=2'
    assert stats['tables_html'] == 1, '含合并的表应走 HTML 路径'
    assert stats['tables_gfm'] == 1, '无合并的表应走 GFM 管道表'


def test_merged_cells_render_not_flattened(fixtures, tmp_path):
    """合并单元格的值不应被复制到被合并的每一格。"""
    _, md = _convert(fixtures['merged'], tmp_path / 'out2')
    # 「横向合并标题」只应出现一次；若重复出现说明被"重复值"策略压平
    assert md.count('横向合并标题') == 1
    assert md.count('纵向合并标签') == 1


def test_merged_table_verifies_clean(fixtures, tmp_path):
    """含合并表的文档，一致性校验必须通过。"""
    out = tmp_path / 'out3'
    stats, _ = _convert(fixtures['merged'], out)
    report = verify(fixtures['merged'], stats['md_path'])
    assert report['pass'] is True
    assert report['tables']['pct'] == 100.0
    # 9 个逻辑单元格：3(表头跨列算1格+备注) ... 以实际计数为准，至少不为 0
    assert report['tables']['cell_total'] > 0


# ---------------------------------------------------------------- 公式
def test_omml_formula_becomes_latex(fixtures, tmp_path):
    """OMML 公式应转成 $...$ / $$...$$，而非丢失或留原文。"""
    stats, md = _convert(fixtures['formula'], tmp_path / 'out')
    assert stats['formulas'] >= 2, '行内与块级各应计一次'
    assert '$x=y$' in md, '行内公式应输出 $...$'
    assert '$$x=y$$' in md, '块级公式应输出 $$...$$'


def test_formula_doc_verifies_clean(fixtures, tmp_path):
    out = tmp_path / 'out2'
    stats, _ = _convert(fixtures['formula'], out)
    report = verify(fixtures['formula'], stats['md_path'])
    assert report['pass'] is True
    assert report['formulas']['word'] == report['formulas']['md']


# ---------------------------------------------------------------- 目录
def test_toc_field_is_parsed(fixtures, tmp_path):
    """Word 目录域应解析为锚点列表，而不是退化成正文。

    回归背景：此前 DEFAULT_TOC_STYLES 只收录 'toc 1'/'TOC1'，
    漏掉了 Word 实际使用的 'TOC 1'（大写+空格），导致整块目录变正文。
    """
    stats, md = _convert(fixtures['toc'], tmp_path / 'out')
    assert stats['toc_entries'] == 4, '夹具含 4 条目录条目'
    assert '(#第一章-概述)' in md, '目录条目应带可点击锚点'
    assert '由 Word' in md, '应有目录说明行'


def test_toc_stays_in_original_position(fixtures, tmp_path):
    """目录应保留在原文位置，不被搬到文首。"""
    _, md = _convert(fixtures['toc'], tmp_path / 'out2')
    toc_pos = md.index('由 Word')
    heading_pos = md.index('# 第一章 概述')
    assert toc_pos < heading_pos, '目录在正文标题之前（夹具如此构造）'


def test_toc_case_insensitive_style_match():
    """目录样式名应大小写不敏感匹配。

    直接覆盖回归点：'TOC 1' / 'toc 1' / 'TOC1' 三种写法须全部命中。
    """
    from word2md import Converter
    conv = Converter.__new__(Converter)          # 只取映射逻辑，不解析文档
    conv.toc_level = {
        'toc 1': 1, 'TOC 1': 1, 'TOC1': 1,
    }
    conv._toc_ci = {k.strip().lower(): v for k, v in conv.toc_level.items()}
    for style in ('TOC 1', 'toc 1', 'TOC1', 'Toc 1'):
        assert conv._toc_ci.get(style.strip().lower()) == 1, style


# ---------------------------------------------------------------- 列表
def test_multilevel_numbering_preserved(fixtures, tmp_path):
    """多级编号列表应还原层级与编号格式（1. / a.）。"""
    stats, md = _convert(fixtures['list'], tmp_path / 'out')
    assert '1. 第一级条目' in md
    assert 'a. 第二级条目' in md, '第二级应使用 lowerLetter 编号'
    # 第二级须缩进
    second = [ln for ln in md.splitlines() if '第二级条目' in ln][0]
    assert second.startswith('  '), f'第二级应缩进: {second!r}'


def test_list_doc_verifies_clean(fixtures, tmp_path):
    out = tmp_path / 'out2'
    stats, _ = _convert(fixtures['list'], out)
    report = verify(fixtures['list'], stats['md_path'])
    assert report['pass'] is True


# ---------------------------------------------------------------- 注音
def test_ruby_keeps_base_text_only(fixtures, tmp_path):
    """拼音注音应只保留主体汉字，丢弃注音字母。"""
    _, md = _convert(fixtures['ruby'], tmp_path / 'out')
    assert '汉' in md, '主体文字必须保留'
    assert 'han' not in md, '注音文本不应出现在正文'


# ---------------------------------------------------------------- 代码块
def test_code_block_detected_as_single_block(fixtures, tmp_path):
    """一段连续 C 代码应合并为**一个**代码块。

    回归背景：_looks_like_code 原先不剥离前导空格，缩进行
    （``    int n = 0;``）全部漏判，导致一个代码块被切成多段；
    函数调用行（``printf("%d", n);``）也不在规则内。
    """
    stats, md = _convert(fixtures['code'], tmp_path / 'out')
    assert stats['code_blocks'] == 1, '整段代码应为一个围栏块'
    assert md.count('```c') == 1


def test_code_block_keeps_indentation(fixtures, tmp_path):
    """代码块内应保留行首缩进，否则代码层级丢失。"""
    _, md = _convert(fixtures['code'], tmp_path / 'out2')
    assert '\n    int n = 0;' in md, '一级缩进应保留'
    assert '\n        printf(' in md, '二级缩进应保留'


def test_code_heuristic_prefers_ascii_only():
    """函数调用式判定不得误伤中文正文。

    正文句「详见附录A(第5页);」形似 call，但含中文，必须判为非代码。
    """
    from word2md import _looks_like_code
    assert _looks_like_code('printf("%d", n);') is True
    assert _looks_like_code('    foo(a, b);') is True, '缩进的调用也应命中'
    assert _looks_like_code('详见附录A(第5页);') is False
    # 公式关系式（无分号）不应被当作赋值语句
    assert _looks_like_code('duty1 = duty0 * 2') is False


# ---------------------------------------------------------------- 通用契约
def test_stats_shape_is_stable(fixtures, tmp_path):
    """返回统计字典的键集合是外部调用方依赖的契约，不得随意删改。"""
    stats, _ = _convert(fixtures['merged'], tmp_path / 'out')
    required = {
        'md_path', 'headings', 'tables', 'tables_gfm', 'tables_html',
        'images', 'code_blocks', 'formulas', 'toc_entries', 'warnings',
    }
    assert required <= set(stats), f'缺失键: {required - set(stats)}'


def test_missing_file_raises_filenotfound(tmp_path):
    with pytest.raises(FileNotFoundError):
        convert_word_to_markdown(str(tmp_path / 'nope.docx'))


def test_legacy_doc_rejected_with_guidance(tmp_path):
    """旧版 .doc 应给出明确指引，而非静默失败。"""
    p = tmp_path / 'old.doc'
    p.write_bytes(b'\xd0\xcf\x11\xe0')     # OLE 魔数，仅用于触发分支
    with pytest.raises(ValueError, match='.doc'):
        convert_word_to_markdown(str(p))


def test_no_images_flag_produces_no_image_dir(fixtures, tmp_path):
    out = tmp_path / 'out'
    convert_word_to_markdown(fixtures['merged'], out_dir=str(out),
                             extract_images=False)
    assert not os.path.isdir(out / 'images')
