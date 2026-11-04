# -*- coding: utf-8 -*-
"""批量 Word->Markdown 转换驱动。

用法:
    python batch_convert.py <根目录> [--out 输出根目录] [--workers N]

行为:
  - 递归扫描根目录下全部 .docx(跳过 ~$ 临时文件与输出目录自身)
  - 按原相对目录结构镜像输出到 <输出根目录>/<相对路径去扩展名>/<同名>.md
  - 每份文档的统计/耗时/异常写入 batch_convert_result.json
"""
import argparse
import json
import os
import sys
import time
import traceback
from concurrent.futures import ProcessPoolExecutor, as_completed

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from word2md import convert_word_to_markdown  # noqa: E402


def collect(root, exclude_dirs):
    jobs = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(
            d for d in dirnames if os.path.join(dirpath, d) not in exclude_dirs
        )
        for fn in sorted(filenames):
            if not fn.lower().endswith('.docx') or fn.startswith('~$'):
                continue
            src = os.path.join(dirpath, fn)
            jobs.append(src)
    return jobs


def one(args_tuple):
    src, root, out_root = args_tuple
    t0 = time.time()
    rec = {'src': os.path.basename(src), 'status': 'ok'}
    try:
        if out_root:
            rel = os.path.relpath(os.path.dirname(src), root)
            # Windows 会静默剥除路径段末尾空格(如 "文档A v1.0 .docx"),
            # 导致后续 makedirs 路径与预期不一致 -> 显式 strip 后再拼接
            stem = os.path.splitext(os.path.basename(src))[0].strip()
            target = os.path.join(out_root, rel, stem)
        else:
            target = None
        stats = convert_word_to_markdown(src, out_dir=target)
        rec.update({
            'md_path': stats['md_path'],
            'headings': stats['headings'],
            'tables': stats['tables'],
            'tables_gfm': stats['tables_gfm'],
            'tables_html': stats['tables_html'],
            'images': stats['images'],
            'code_blocks': stats['code_blocks'],
            'formulas': stats['formulas'],
            'toc_entries': stats['toc_entries'],
            'warnings': stats['warnings'],
        })
    except Exception as e:  # noqa: BLE001
        rec['status'] = 'failed'
        rec['error'] = f'{type(e).__name__}: {e}'
        rec['traceback'] = traceback.format_exc()
    rec['seconds'] = round(time.time() - t0, 1)
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('root')
    ap.add_argument('--out', default=None,
                    help='Markdown 输出根目录(镜像原目录结构); 缺省写在源文件旁')
    ap.add_argument('--workers', type=int, default=4)
    ap.add_argument('--json', default=None, help='结果 JSON 路径')
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    out_root = os.path.abspath(args.out) if args.out else None
    jobs = collect(root, {os.path.join(root, '_converted')})
    print(f'发现 {len(jobs)} 份 .docx, 并发 {args.workers}', flush=True)

    results = []
    with ProcessPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(one, (j, root, out_root)): j for j in jobs}
        for i, fut in enumerate(as_completed(futs), 1):
            rec = fut.result()
            src = futs[fut]
            rec['rel_dir'] = os.path.relpath(os.path.dirname(src), root).replace('\\', '/')
            rec['root'] = rec['rel_dir'].split('/')[0]
            results.append(rec)
            mark = 'OK ' if rec['status'] == 'ok' else 'ERR'
            print(f'[{i}/{len(jobs)}] {mark} {rec["rel_dir"]}/{rec["src"]} '
                  f'({rec["seconds"]}s)', flush=True)

    results.sort(key=lambda r: (r['rel_dir'], r['src']))
    ok = sum(1 for r in results if r['status'] == 'ok')
    print(f'\n完成: 成功 {ok} / 失败 {len(results) - ok} / 总计 {len(results)}')

    # 结果 JSON 默认写到当前工作目录, 而非输入目录 ——
    # 输入目录常是只读的文档库/网盘挂载点, 写入会失败或污染源目录结构
    json_path = args.json or os.path.join(os.getcwd(), 'batch_convert_result.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'结果已写入 {json_path}')


if __name__ == '__main__':
    main()