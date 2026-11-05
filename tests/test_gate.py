# -*- coding: utf-8 -*-
"""质量门禁测试。

门禁的价值在于"能失败"。一个永远返回 0 的 --fail-under 比没有更糟——它会让 CI
绿灯而内容已经损坏。因此这里对每条门禁都同时验证**通过路径**与**失败路径**。
"""
from __future__ import annotations

import conftest

from review_probe import main as probe_main
from review_probe import probe
from word2md import convert_word_to_markdown


def _low_quality_docx(path):
    """构造低可读性文档：全文无标题 -> 结构可导航性扣 40 分。"""
    import docx
    d = docx.Document()
    for i in range(3):
        d.add_paragraph(f'纯正文段落 {i}，全文没有任何标题。')
    d.save(str(path))
    return path


def test_probe_returns_stable_shape(tmp_path):
    """probe() 是 MCP 与门禁共用的口径，返回键不得随意改动。"""
    src = tmp_path / 'doc.docx'
    conftest.build_merged_table(str(src))
    out = tmp_path / 'out'
    stats = convert_word_to_markdown(str(src), out_dir=str(out))

    result = probe(str(src), stats['md_path'], sample=2)
    for key in ('structure', 'defects', 'readability', 'paragraph_check',
                'table_check', 'samples'):
        assert key in result, key
    assert 0 <= result['readability']['score'] <= 100
    assert result['readability']['grade'] in ('A', 'B', 'C', 'D')


def test_probe_fail_under_passes_above_threshold(tmp_path):
    src = tmp_path / 'doc.docx'
    conftest.build_merged_table(str(src))
    out = tmp_path / 'out'
    stats = convert_word_to_markdown(str(src), out_dir=str(out))
    assert probe_main([str(src), stats['md_path'], '--fail-under', '0']) == 0


def test_probe_fail_under_fails_below_threshold(tmp_path):
    src = _low_quality_docx(tmp_path / 'low.docx')
    out = tmp_path / 'out'
    stats = convert_word_to_markdown(str(src), out_dir=str(out))
    assert probe_main([str(src), stats['md_path'], '--fail-under', '95']) == 1


def test_probe_fail_under_absent_always_zero(tmp_path):
    """不给阈值时，即使分数很低也不应阻断——避免影响仅想查看诊断的用户。"""
    src = _low_quality_docx(tmp_path / 'low2.docx')
    out = tmp_path / 'out'
    stats = convert_word_to_markdown(str(src), out_dir=str(out))
    assert probe_main([str(src), stats['md_path']]) == 0


def test_word2md_cli_verify_flag_passes(tmp_path):
    """--verify 应一步完成转换+校验，正确文档返回 0。"""
    from word2md import main
    src = tmp_path / 'doc.docx'
    conftest.build_merged_table(str(src))
    out = tmp_path / 'out'
    assert main([str(src), '-o', str(out), '--verify']) == 0


def test_word2md_cli_fail_under_blocks(tmp_path):
    """低可读性文档配高阈值必须返回 1。"""
    from word2md import main
    src = _low_quality_docx(tmp_path / 'low.docx')
    out = tmp_path / 'out'
    assert main([str(src), '-o', str(out), '--verify',
                 '--probe', '--fail-under', '95']) == 1


def test_word2md_cli_without_verify_skips_check(tmp_path):
    """不带任何门禁参数时应保持原有行为（返回 0，不解析校验模块）。"""
    from word2md import main
    src = _low_quality_docx(tmp_path / 'low2.docx')
    out = tmp_path / 'out'
    assert main([str(src), '-o', str(out)]) == 0
