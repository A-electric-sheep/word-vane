# -*- coding: utf-8 -*-
"""校验器自身的测试。

重点不是"转换对了没有"（那是 test_convert.py 的事），而是：
**校验器能不能抓住错误**。一个永远返回 pass=True 的校验器比没有校验器更危险，
因此这里刻意构造篡改样本，要求校验器必须判为不通过。
"""
from __future__ import annotations

import conftest

from verify_consistency import verify
from word2md import convert_word_to_markdown


def _prepare(tmp_path, name):
    src = tmp_path / f'{name}.docx'
    conftest.BUILDERS[name](str(src))
    out = tmp_path / f'{name}_out'
    stats = convert_word_to_markdown(str(src), out_dir=str(out))
    return src, stats['md_path']


def test_verify_passes_on_faithful_conversion(tmp_path):
    src, md = _prepare(tmp_path, 'merged')
    report = verify(str(src), str(md))
    assert report['pass'] is True
    assert report['paragraphs']['pct'] == 100.0


def test_verify_detects_deleted_table(tmp_path):
    """删掉整张表，校验器必须发现。"""
    src, md = _prepare(tmp_path, 'merged')
    text = open(md, encoding='utf-8').read()

    # 移除 HTML 表格块
    start = text.index('<table')
    end = text.index('</table>') + len('</table>')
    tampered = text[:start] + text[end:]
    bad = tmp_path / 'tampered.md'
    bad.write_text(tampered, encoding='utf-8')

    report = verify(str(src), str(bad))
    assert report['pass'] is False, '表格被删除后不应通过'
    assert report['tables']['n_total'] >= 2
    assert report['tables']['pct'] < 100.0


def test_verify_detects_deleted_paragraph(tmp_path):
    """删掉一段正文，段落流一致率必须下降。"""
    src, md = _prepare(tmp_path, 'merged')
    text = open(md, encoding='utf-8').read()
    tampered = text.replace('两张表之间的说明段。', '')
    assert tampered != text, '夹具前提：目标段落必须存在'

    bad = tmp_path / 'no_para.md'
    bad.write_text(tampered, encoding='utf-8')
    report = verify(str(src), str(bad))
    assert report['pass'] is False


def test_verify_detects_missing_image_file(tmp_path):
    """图片引用存在但文件被删，应报缺失。"""
    import docx
    from PIL import Image

    img = tmp_path / 'pic.png'
    Image.new('RGB', (24, 24), (200, 200, 200)).save(img)

    src = tmp_path / 'with_img.docx'
    d = docx.Document()
    d.add_heading('图片文档', 1)
    d.add_paragraph('正文。')
    d.add_picture(str(img))
    d.save(str(src))

    out = tmp_path / 'img_out'
    stats = convert_word_to_markdown(str(src), out_dir=str(out))

    # 删除提取出来的图片，制造"引用悬空"
    img_dir = out / 'images'
    assert img_dir.is_dir(), '应提取出 images 目录'
    for f in img_dir.iterdir():
        f.unlink()

    report = verify(str(src), stats['md_path'])
    assert report['images']['missing'], '应报告缺失图片'


def test_verify_formula_count_mismatch(tmp_path):
    """公式占位数量不一致时必须不通过。"""
    src, md = _prepare(tmp_path, 'formula')
    text = open(md, encoding='utf-8').read()
    tampered = text.replace('$$x=y$$', 'x=y')     # 抹掉块级公式标记
    assert tampered != text

    bad = tmp_path / 'no_formula.md'
    bad.write_text(tampered, encoding='utf-8')
    report = verify(str(src), str(bad))
    assert report['pass'] is False
    assert report['formulas']['word'] != report['formulas']['md']


def test_verify_cli_returns_nonzero_on_failure(tmp_path):
    """CLI 退出码必须能表达失败，供 CI 直接使用。"""
    from verify_consistency import main

    src, md = _prepare(tmp_path, 'merged')
    assert main([str(src), str(md)]) == 0, '正确转换应返回 0'

    bad = tmp_path / 'bad.md'
    bad.write_text('# 完全不相干的内容\n', encoding='utf-8')
    assert main([str(src), str(bad)]) == 1, '不一致应返回 1'
