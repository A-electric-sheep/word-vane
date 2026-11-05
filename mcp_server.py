# -*- coding: utf-8 -*-
"""MCP Server：把「转换 + 校验 + 可读性评分」封装成单一工具，供 Agent 调用。

设计取向
--------
**只暴露一个工具，且一次调用完成"转换→校验→评分"三件事。**

这不是为了省事，而是刻意贴合 Agent 的实际用法：Agent 拿到一份 Word，真正关心的
不是"先转换、再另起一次调用校验、第三次调用算分"，而是"这份文档能不能放心用"。
拆成三个工具会让每次调用都产生一轮模型往返与上下文开销，且中间产物路径需要
在对话里传递——徒增失败面。

安全边界
--------
1. 只接受**本地文件路径**，不接受 URL / data URI。避免服务端被诱导发起网络请求
   （SSRF），也避免把外部内容当成可信文档处理。
2. 拒绝非 .docx 后缀；.doc 走明确报错提示。
3. 输出目录由调用方指定或落在系统临时目录，不写入源文档所在目录（源目录常是
   只读挂载点）。
4. 不加载任何第三方 MCP 插件、不执行文档内嵌宏或 OLE 对象。

依赖
----
`mcp` 为可选依赖：`pip install "word-vane[mcp]"`。未安装时 `main()` 给出明确指引，
而不是抛 ImportError 堆栈。
"""
from __future__ import annotations

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

SERVER_NAME = 'word2md'


def _err(message, **extra):
    """统一错误返回结构。Agent 侧可据 status 字段决定是否重试。"""
    payload = {'status': 'error', 'error': message}
    payload.update(extra)
    return payload


def convert_and_verify(path, out_dir=None, fail_under=None, keep_images=True):
    """转换一份 .docx，并返回校验与可读性结论。

    参数:
        path:        本地 .docx 绝对路径
        out_dir:     输出目录；缺省落到系统临时目录下的唯一子目录
        fail_under:  可读性分数阈值；低于则 status 置为 'below_threshold'
        keep_images: 是否提取图片

    返回 dict:
        status:      'ok' | 'below_threshold' | 'error'
        converted:   是否内容级校验通过
        readability: {score, grade, verdict}
        consistency: {paragraph_pct, table_pct, formula_word, formula_md,
                      images_missing}
        artifacts:   {md_path, images_dir, output_dir}
        warnings:    转换器告警
    """
    # ---- 输入校验：先挡掉明显不合规的输入，避免半途失败 ----
    if not path:
        return _err('缺少 path 参数')
    if isinstance(path, str) and ('://' in path or path.startswith('data:')):
        return _err('只接受本地文件路径，不接受 URL 或 data URI')
    src = os.path.abspath(os.path.expanduser(str(path)))
    if not os.path.isfile(src):
        return _err(f'找不到文件: {src}')
    if src.lower().endswith('.doc'):
        return _err('旧版 .doc 不被直接支持，请先另存为 .docx '
                    '(或 soffice --headless --convert-to docx 输入.doc)')
    if not src.lower().endswith('.docx'):
        return _err(f'不支持的输入格式: {os.path.basename(src)}，仅支持 .docx')

    # ---- 输出目录: 缺省不写源目录, 避免污染或写入只读挂载点 ----
    if out_dir:
        target = os.path.abspath(os.path.expanduser(str(out_dir)))
        os.makedirs(target, exist_ok=True)
    else:
        target = tempfile.mkdtemp(prefix='word2md_mcp_')

    # ---- 延迟导入: 让未装 mcp 的用户仍能用 CLI ----
    from review_probe import probe
    from verify_consistency import verify
    from word2md import convert_word_to_markdown

    try:
        stats = convert_word_to_markdown(
            src, out_dir=target, extract_images=bool(keep_images))
    except Exception as e:                                   # noqa: BLE001
        return _err(f'转换失败: {type(e).__name__}: {e}')

    md_path = stats['md_path']
    try:
        report = verify(src, md_path)
        pr = probe(src, md_path, sample=0)     # sample=0: 不要抽样正文, 省 token
    except Exception as e:                                   # noqa: BLE001
        return _err(f'校验失败: {type(e).__name__}: {e}',
                    artifacts={'md_path': md_path, 'output_dir': target})

    rd = pr['readability']
    score = rd['score']
    status = 'ok'
    if not report['pass']:
        status = 'mismatch'
    elif fail_under is not None and score < fail_under:
        status = 'below_threshold'

    images_dir = os.path.join(target, 'images')
    return {
        'status': status,
        'converted': report['pass'],
        'readability': {
            'score': score,
            'grade': rd['grade'],
            'verdict': rd['verdict'],
            'defects': pr['defects'][:10],
        },
        'consistency': {
            'paragraph_pct': report['paragraphs']['pct'],
            'paragraph_matched': report['paragraphs']['matched'],
            'paragraph_total': report['paragraphs']['word_n'],
            'table_pct': report['tables']['pct'],
            'cell_matched': report['tables']['cell_matched'],
            'cell_total': report['tables']['cell_total'],
            'formula_word': report['formulas']['word'],
            'formula_md': report['formulas']['md'],
            'images_missing': len(report['images'].get('missing') or []),
        },
        'structure': {
            'headings': stats['headings'],
            'tables': stats['tables'],
            'tables_html': stats['tables_html'],
            'tables_gfm': stats['tables_gfm'],
            'images': stats['images'],
            'formulas': stats['formulas'],
            'code_blocks': stats['code_blocks'],
            'toc_entries': stats['toc_entries'],
        },
        'artifacts': {
            'md_path': md_path,
            'images_dir': images_dir if os.path.isdir(images_dir) else None,
            'output_dir': target,
        },
        'warnings': stats['warnings'],
    }


def _load_server_class():
    """兼容 mcp SDK 1.x 与 2.x 的 server 类加载。

    - mcp >= 2：`FastMCP` 已重命名为 `MCPServer`（`mcp.server.mcpserver`）
    - mcp 1.x：`from mcp.server.fastmcp import FastMCP`

    先试 2.x，再退 1.x。两代的 `tool()` 装饰器与 `run(transport='stdio')`
    用法一致，因此上层代码无需分支。
    """
    try:
        from mcp.server.mcpserver import MCPServer
        return MCPServer
    except ImportError:
        from mcp.server.fastmcp import FastMCP
        return FastMCP


def build_server():
    """构造 MCP server 实例并注册工具。单独成函数便于测试与复用。"""
    server_cls = _load_server_class()
    mcp = server_cls(SERVER_NAME)

    @mcp.tool()
    def word2md_convert_and_verify(
        path: str,
        out_dir: str | None = None,
        fail_under: float | None = None,
        keep_images: bool = True,
    ) -> dict:
        """把一份本地 Word(.docx) 转成 Markdown，并立即校验内容无丢失。

        一次调用完成三件事：转换 → 逐段/逐格一致性校验 → AI 可读性评分。
        适合在把文档喂给 LLM 之前做质量把关。

        Args:
            path: 本地 .docx 文件绝对路径（不接受 URL）。
            out_dir: 输出目录。缺省落在系统临时目录，不写源文件目录。
            fail_under: 可读性分数阈值(0-100)。低于该值时 status 为
                'below_threshold'，便于调用方决定是否采用结果。
            keep_images: 是否提取内嵌图片，默认 True。

        Returns:
            dict，含 status(ok/mismatch/below_threshold/error)、converted、
            readability{score,grade,verdict}、consistency{...}、
            structure{...}、artifacts{md_path,images_dir}、warnings。
        """
        return convert_and_verify(path, out_dir=out_dir,
                                  fail_under=fail_under,
                                  keep_images=keep_images)

    return mcp


def main(argv=None):
    try:
        mcp = build_server()
    except ImportError:
        print('未安装 MCP 依赖。请先执行:\n'
              '    pip install "word-vane[mcp]"\n'
              '然后再运行:\n'
              '    word2md-mcp   (或 python mcp_server.py)',
              file=sys.stderr)
        return 2
    # 默认走 stdio，由 MCP 客户端负责拉起进程
    mcp.run(transport='stdio')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
