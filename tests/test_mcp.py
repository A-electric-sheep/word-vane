# -*- coding: utf-8 -*-
"""MCP server 的业务逻辑测试。

注意：这里**不启动 MCP 传输层**（stdio 需要真实客户端握手），只测
`convert_and_verify` 这个核心函数——传输层是 mcp 库的职责，业务口径才是
本项目要负责的部分。这样测试无需安装 mcp 也能跑。
"""
from __future__ import annotations

import conftest

import mcp_server


def test_convert_and_verify_ok(tmp_path):
    src = tmp_path / 'doc.docx'
    conftest.build_merged_table(str(src))
    out = tmp_path / 'out'

    result = mcp_server.convert_and_verify(str(src), out_dir=str(out))
    assert result['status'] == 'ok'
    assert result['converted'] is True
    assert result['consistency']['paragraph_pct'] == 100.0
    assert result['consistency']['table_pct'] == 100.0
    assert result['consistency']['images_missing'] == 0
    # 合并表应走 HTML 路径
    assert result['structure']['tables_html'] == 1
    # 产物路径必须真实存在，Agent 会据此读文件
    assert result['artifacts']['md_path'] and \
        result['artifacts']['md_path'].endswith('.md')


def test_default_out_dir_is_not_source_dir(tmp_path):
    """缺省输出必须落在临时目录，不得污染源文档所在目录。"""
    src = tmp_path / 'doc.docx'
    conftest.build_merged_table(str(src))
    before = set(p.name for p in tmp_path.iterdir())

    result = mcp_server.convert_and_verify(str(src))
    assert result['status'] == 'ok'
    out_dir = result['artifacts']['output_dir']
    assert str(tmp_path) not in out_dir, '默认输出不应写回源目录'

    after = set(p.name for p in tmp_path.iterdir())
    assert after == before, '源目录不应新增任何文件'


def test_rejects_url_input():
    """不接受 URL，避免服务端被诱导发起网络请求。"""
    r = mcp_server.convert_and_verify('https://example.com/a.docx')
    assert r['status'] == 'error'
    assert 'URL' in r['error'] or '本地' in r['error']


def test_rejects_data_uri():
    r = mcp_server.convert_and_verify('data:application/vnd.openxmlformats')
    assert r['status'] == 'error'


def test_missing_file(tmp_path):
    r = mcp_server.convert_and_verify(str(tmp_path / 'nope.docx'))
    assert r['status'] == 'error'
    assert '找不到' in r['error']


def test_legacy_doc_message(tmp_path):
    p = tmp_path / 'old.doc'
    p.write_bytes(b'\xd0\xcf\x11\xe0')
    r = mcp_server.convert_and_verify(str(p))
    assert r['status'] == 'error'
    assert '.docx' in r['error']


def test_unsupported_extension(tmp_path):
    p = tmp_path / 'a.txt'
    p.write_text('x', encoding='utf-8')
    r = mcp_server.convert_and_verify(str(p))
    assert r['status'] == 'error'
    assert '不支持' in r['error']


def test_fail_under_marks_below_threshold(tmp_path):
    """低可读性文档 + 高阈值 -> status 应为 below_threshold 而非 ok。"""
    import docx
    src = tmp_path / 'low.docx'
    d = docx.Document()
    for i in range(3):
        d.add_paragraph(f'纯正文段落 {i}，全文没有标题。')
    d.save(str(src))

    r = mcp_server.convert_and_verify(str(src), out_dir=str(tmp_path / 'o'),
                                      fail_under=95)
    assert r['status'] == 'below_threshold'
    assert r['readability']['score'] < 95
    # 内容本身没丢，converted 仍应为 True —— 两个维度分开表达
    assert r['converted'] is True


def test_result_keys_are_stable(tmp_path):
    """MCP 工具返回值是 Agent 的契约，键名不得随意改动。"""
    src = tmp_path / 'doc.docx'
    conftest.build_merged_table(str(src))
    r = mcp_server.convert_and_verify(str(src), out_dir=str(tmp_path / 'o'))
    for k in ('status', 'converted', 'readability', 'consistency',
              'structure', 'artifacts', 'warnings'):
        assert k in r, k
    for k in ('score', 'grade', 'verdict'):
        assert k in r['readability'], k
    for k in ('paragraph_pct', 'table_pct', 'images_missing'):
        assert k in r['consistency'], k
