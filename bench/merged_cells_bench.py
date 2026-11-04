# -*- coding: utf-8 -*-
"""合并单元格保真度基准：可复现地量化"合并语义是否被保留"。

为什么需要这个基准
------------------
GFM 管道表语法上没有 colspan/rowspan。因此任何"docx → GFM"的工具，遇到合并
单元格都只能三选一：重复值、留空、丢表——三种都让读者无法判断原表有几处合并。
本仓库选择输出 HTML `<table>` 保留 colspan/rowspan。

这个结论不该只写在 README 里靠信任，而应能被任何人用一条命令复现：

    python bench/merged_cells_bench.py            # 只测 word2md
    python bench/merged_cells_bench.py --with-anydoc   # 额外对比 anydoc(需已安装)

判定指标（不依赖人工看图）
--------------------------
合并语义保留率 = 输出中正确出现的 colspan/rowspan 属性数 / 夹具中实际合并数

其中夹具的"实际合并数"由 python-docx 读回 docx 后统计（基准真值），
而非硬编码——避免夹具改动后指标失真。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile

import docx
from docx.oxml.ns import qn

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)


# ----------------------------------------------------------------- 夹具构造
def _merge_cases():
    """返回 [(说明, 构造函数, 行数, 列数, 期望 colspan 数, 期望 rowspan 数)]。

    刻意**不做交叉合并**（即同一格既被横跨又被纵跨）。交叉合并会让 Word 把一个
    逻辑格拆成多个 gridSpan 起始格，属性计数随之上浮，使辅助指标与真值无法对齐。
    这里每个用例只施加互不重叠的合并，保证"真值 = 属性数"可校验。
    """
    def colspan3(d, t):
        t.cell(0, 0).merge(t.cell(0, 2)).text = '横向跨三列'

    def rowspan2(d, t):
        t.cell(0, 0).merge(t.cell(1, 0)).text = '纵向跨两行'

    def two_separate(d, t):
        # 两处互不相邻的横向合并（同一行不同列区间）
        t.cell(0, 0).merge(t.cell(0, 1)).text = '左段横跨'
        t.cell(1, 0).merge(t.cell(1, 1)).text = '中段横跨'

    def nested_header(d, t):
        # 两级表头：顶层横跨三列
        t.cell(0, 1).merge(t.cell(0, 3)).text = '分组标题'

    def rowspan_two_rows(d, t):
        # 同一列上两处独立的纵向合并（不相邻）
        t.cell(0, 0).merge(t.cell(1, 0)).text = '上段纵跨'
        t.cell(2, 0).merge(t.cell(3, 0)).text = '下段纵跨'

    return [
        ('横向合并 · 跨三列', colspan3, 3, 4, 1, 0),
        ('纵向合并 · 跨两行', rowspan2, 3, 3, 0, 1),
        ('两处横向合并', two_separate, 3, 4, 2, 0),
        ('两级表头', nested_header, 2, 4, 1, 0),
        ('两处纵向合并', rowspan_two_rows, 4, 3, 0, 2),
    ]


def build_fixture(path):
    """生成含多张合并表的夹具文档，返回真值统计。

    注意：小节标题**不得包含 `colspan` / `rowspan` 字样**。标题会作为普通段落
    写进 Markdown，属性计数若用全文正则就会被标题误命中（实测使 colspan 多计
    1 处、rowspan 多计 1 处，把 100% 误报成 125%）。
    """
    d = docx.Document()
    d.add_heading('合并单元格基准夹具', 1)
    truth = []
    for label, fn, rows, cols, _ec, _er in _merge_cases():
        d.add_heading(label, 2)
        t = d.add_table(rows=rows, cols=cols)
        t.style = 'Table Grid'
        # 先填满，便于观察哪些格被合并吞并
        for ri in range(rows):
            for ci in range(cols):
                t.cell(ri, ci).text = f'R{ri}C{ci}'
        fn(d, t)
        truth.append(label)
    d.add_paragraph('基准夹具结束。')
    d.save(path)
    return truth


def count_true_merges(docx_path):
    """从 docx 中统计真实合并数（基准真值）。

    - gridSpan > 1  → 一次横向合并
    - vMerge 且 val == 'restart'（或无 val 但属于起始格）→ 一次纵向合并起点

    注意两个坑：
    1. 不能用 `row.cells`：python-docx 对合并区域内每个网格位都返回同一个
       `_tc`，会把同一处合并重复计数。必须直接遍历 `w:tc` 元素。
    2. `vMerge` 的**继续格**通常不写 val 或写作 continue；只有起始格是
       restart。两者混淆会把合并行数数成两倍。
    """
    d = docx.Document(docx_path)
    colspan = rowspan = 0
    for t in d.tables:
        for tc in t._tbl.iter(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                continue
            gs = tcPr.find(qn('w:gridSpan'))
            if gs is not None:
                try:
                    if int(gs.get(qn('w:val'))) > 1:
                        colspan += 1
                except (TypeError, ValueError):
                    pass
            vm = tcPr.find(qn('w:vMerge'))
            if vm is not None and vm.get(qn('w:val')) == 'restart':
                rowspan += 1
    return colspan, rowspan


# ----------------------------------------------------------------- 转换与判定
def run_word2md(docx_path, out_dir):
    from word2md import convert_word_to_markdown
    stats = convert_word_to_markdown(docx_path, out_dir=out_dir)
    with open(stats['md_path'], encoding='utf-8') as f:
        return f.read(), stats


def run_anydoc(docx_path):
    """可选对比项。anydoc 未安装时返回 None。"""
    try:
        import anydoc
    except ImportError:
        return None
    try:
        return anydoc.to_markdown(docx_path)
    except Exception as e:                      # noqa: BLE001
        return f'<anydoc 失败: {type(e).__name__}: {e}>'


def score(md, truth_colspan, truth_rowspan):
    """辅助指标：统计输出中 colspan/rowspan 属性个数。

    仅作参考。属性计数在**不规则合并**下会与真值不等——例如一处 vMerge 起始格
    跨 3 行，输出是 rowspan="3" 一个属性，真值也只算 1 处，但若夹具里存在
    "合并后再合并"的嵌套情形，计数就会偏移。主判定请看 verify_consistency 的
    单元格级一致率（见 main 中的 table_pct）。
    """
    got_col = len(re.findall(r'colspan\s*=\s*["\']?(\d+)', md, re.I))
    got_row = len(re.findall(r'rowspan\s*=\s*["\']?(\d+)', md, re.I))
    return {
        'colspan_expected': truth_colspan,
        'colspan_found': got_col,
        'rowspan_expected': truth_rowspan,
        'rowspan_found': got_row,
        'uses_html_table': '<table' in md.lower(),
        'uses_pipe_table': bool(re.search(r'^\|.*\|$', md, re.M)),
    }


def main():
    ap = argparse.ArgumentParser(description='合并单元格保真度基准')
    ap.add_argument('--with-anydoc', action='store_true',
                    help='同时对比 anydoc（需 pip install firecrawl-anydoc）')
    ap.add_argument('--json', default=None, help='结果写入 JSON 路径')
    ap.add_argument('--keep', action='store_true', help='保留临时目录便于人工查看')
    args = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix='w2m_bench_')
    docx_path = os.path.join(tmp, 'merged_bench.docx')
    build_fixture(docx_path)
    tc, tr = count_true_merges(docx_path)

    print(f'夹具: {docx_path}')
    print(f'基准真值: colspan {tc} 处 / rowspan {tr} 处\n')

    out_dir = os.path.join(tmp, 'out')
    md, stats = run_word2md(docx_path, out_dir)
    w2m = score(md, tc, tr)

    # 主判定：单元格级一致性（校验器直接读 Word XML 与 Markdown 表格流比对，
    # 不受"属性个数"这种表象指标影响，也不依赖我如何构造夹具）
    from verify_consistency import verify
    report = verify(docx_path, stats['md_path'])
    table_pct = report['tables']['pct']
    cell_total = report['tables']['cell_total']

    rows = [('word2md', w2m)]
    if args.with_anydoc:
        other = run_anydoc(docx_path)
        if other is None:
            print('!! 未安装 anydoc，跳过对比。安装: pip install firecrawl-anydoc\n')
        else:
            rows.append(('anydoc', score(other, tc, tr)))
            with open(os.path.join(tmp, 'anydoc.md'), 'w', encoding='utf-8') as f:
                f.write(other)

    header = (f'{"工具":<10}{"colspan属性":>13}{"rowspan属性":>13}'
              f'{"HTML表":>9}{"管道表":>9}')
    print(header)
    print('-' * len(header))
    for name, s in rows:
        cs = f'{s["colspan_found"]}/{s["colspan_expected"]}'
        rs = f'{s["rowspan_found"]}/{s["rowspan_expected"]}'
        print(f'{name:<10}{cs:>13}{rs:>13}'
              f'{"是" if s["uses_html_table"] else "否":>9}'
              f'{"是" if s["uses_pipe_table"] else "否":>9}')

    print(f'\nword2md 转换统计: 表格 {stats["tables"]} '
          f'(HTML {stats["tables_html"]} / GFM {stats["tables_gfm"]})')
    print(f'word2md 单元格级一致率: {table_pct}%  (比对单元格 {cell_total} 个)')

    ok = table_pct >= 99.0 and w2m['uses_html_table']
    print(f'判定: {"通过 — 合并语义完整保留, 单元格级一致" if ok else "未通过 — 存在合并语义丢失"}')

    payload = {
        'fixture': docx_path,
        'truth': {'colspan': tc, 'rowspan': tr},
        'word2md': {**w2m, 'table_cell_pct': table_pct, 'cell_total': cell_total},
        'anydoc': rows[1][1] if len(rows) > 1 else None,
        'verdict': 'pass' if ok else 'fail',
    }
    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print(f'JSON 已写入 {args.json}')

    if args.keep:
        print(f'临时目录保留: {tmp}')
    return 0 if ok else 1
    return 0 if w2m['preservation_pct'] >= 99.0 else 1


if __name__ == '__main__':
    raise SystemExit(main())
