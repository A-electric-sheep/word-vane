# -*- coding: utf-8 -*-
"""批量一致性校验驱动: 对每份 docx 与其生成 md 运行 verify_consistency。"""
import argparse
import json
import os
import sys
import time
import traceback
from concurrent.futures import ProcessPoolExecutor, as_completed

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from verify_consistency import verify  # noqa: E402


def one(pair):
    docx_path, md_path = pair
    t0 = time.time()
    rec = {'docx': docx_path, 'md': md_path, 'status': 'ok'}
    try:
        rep = verify(docx_path, md_path)
        rec.update(rep)
    except Exception as e:  # noqa: BLE001
        rec['status'] = 'failed'
        rec['error'] = f'{type(e).__name__}: {e}'
        rec['traceback'] = traceback.format_exc()
    rec['seconds'] = round(time.time() - t0, 1)
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('root')
    ap.add_argument('--converted', required=True, help='Markdown 输出根目录(镜像原目录结构)')
    ap.add_argument('--workers', type=int, default=4)
    ap.add_argument('--json', default=None)
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    conv_root = os.path.abspath(args.converted)
    pairs = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.endswith('_md')]
        for fn in sorted(filenames):
            if not fn.lower().endswith('.docx') or fn.startswith('~$'):
                continue
            src = os.path.join(dirpath, fn)
            rel = os.path.relpath(os.path.dirname(src), root)
            stem = os.path.splitext(fn)[0]
            cands = [
                os.path.join(conv_root, rel, stem + '.md'),
                os.path.join(conv_root, rel, stem + '_md', stem + '.md'),
                os.path.join(conv_root, rel, stem, stem + '.md'),
                os.path.join(conv_root, rel, stem, stem.strip() + '.md'),
                os.path.join(conv_root, rel, stem.strip(), stem + '.md'),
                os.path.join(conv_root, rel, stem.strip(), stem.strip() + '.md'),
            ]
            md = next((c for c in cands if os.path.isfile(c)), cands[0])
            pairs.append((src, md))

    print(f'待校验 {len(pairs)} 份', flush=True)
    results = []
    with ProcessPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(one, p): p for p in pairs}
        for i, fut in enumerate(as_completed(futs), 1):
            rec = fut.result()
            results.append(rec)
            par = rec.get('paragraphs', {})
            tab = rec.get('tables', {})
            print(f'[{i}/{len(pairs)}] {os.path.basename(rec["docx"])[:44]:<46} '
                  f'段落 {par.get("pct", "-")}% 表格 {tab.get("pct", "-")}% '
                  f'pass={rec.get("pass")}', flush=True)

    results.sort(key=lambda r: r['docx'])
    npass = sum(1 for r in results if r.get('pass'))
    print(f'\n校验完成: pass {npass} / {len(results)}')

    # 同 batch_convert: 结果默认写当前工作目录, 不污染输入目录
    json_path = args.json or os.path.join(os.getcwd(), 'batch_verify_result.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'结果已写入 {json_path}')


if __name__ == '__main__':
    main()