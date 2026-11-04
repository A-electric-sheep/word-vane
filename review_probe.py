# -*- coding: utf-8 -*-
"""单文档复核探针: 输出 AI 可读性与结构缺陷的客观指标。

用法:
    python review_probe.py <docx> <md> [--sample N]

输出 JSON, 字段说明:
  md.*          : Markdown 侧结构指标(体量/标题/表格/图片/公式/代码块)
  defects       : 结构缺陷清单(空标题/围栏未闭合/表格列数不齐/裸占位符等)
  readability   : 面向"AI 能否读懂"的可读性判断与扣分项
  samples       : 与 Word 段落对位的内容抽样(供人工/子代理比对)
"""
import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from verify_consistency import word_parse, md_parse, norm  # noqa: E402


def md_structure(md_path):
    txt = open(md_path, encoding='utf-8').read()
    lines = txt.split('\n')
    heads = []
    for ln in lines:
        m = re.match(r'^(#{1,6})\s+(.*)$', ln)
        if m:
            heads.append((len(m.group(1)), m.group(2).strip()))
    fences = sum(1 for ln in lines if ln.strip().startswith('```'))
    imgs = re.findall(r'!\[[^\]]*\]\(([^)]+)\)', txt)
    links = re.findall(r'(?<!!)\[[^\]]*\]\(#([^)]*)\)', txt)
    # GFM 表格列数一致性
    bad_tables = []
    i = 0
    tidx = 0
    while i < len(lines):
        if lines[i].strip().startswith('|'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                rows.append(lines[i].strip())
                i += 1
            widths = {r.strip('|').count('|') + 1 for r in rows}
            if len(widths) > 1:
                bad_tables.append({'table_index': tidx, 'col_counts': sorted(widths),
                                   'first_row': rows[0][:80]})
            tidx += 1
        else:
            i += 1
    html_tables = txt.count('<table')
    return {
        'chars': len(txt),
        'lines': len(lines),
        'headings': len(heads),
        'heading_levels': {f'h{i}': sum(1 for lv, _ in heads if lv == i) for i in range(1, 7)},
        'max_heading_depth': max((lv for lv, _ in heads), default=0),
        'images': len(imgs),
        'internal_links': len(links),
        'gfM_tables': tidx,
        'html_tables': html_tables,
        'code_fences': fences,
        'formula_inline': txt.count('$') // 2 if txt.count('$') else 0,
        'heads': heads,
        'imgs': imgs,
        'bad_tables': bad_tables,
        'text': txt,
        'lines_list': lines,
    }


def detect_defects(st, md_dir):
    d = []
    # 1. 代码围栏未闭合
    if st['code_fences'] % 2:
        d.append({'type': 'code_fence_unclosed', 'detail': '围栏数为奇数, 存在未闭合代码块'})
    # 2. 空标题
    for lv, t in st['heads']:
        if not t.strip('*_ '):
            d.append({'type': 'empty_heading', 'detail': f'{"#" * lv} 标题内容为空'})
    # 3. 表格列数不齐
    for b in st['bad_tables']:
        d.append({'type': 'table_column_mismatch', **b})
    # 4. 图片引用但文件缺失
    missing = []
    for p in st['imgs']:
        if p.startswith('http'):
            continue
        fp = os.path.join(md_dir, p)
        if not os.path.isfile(fp):
            missing.append(p)
    if missing:
        d.append({'type': 'missing_image', 'detail': missing[:10], 'count': len(missing)})
    # 5. 残留内部占位符
    for token in ('<<TOC_PLACEHOLDER>>', '\x00F\x00', '[[IMG:'):
        if token in st['text']:
            d.append({'type': 'placeholder_leak', 'detail': token})
    # 6. 标题层级跳级(如 h1 直接到 h3)
    prev = 0
    jumps = []
    for lv, t in st['heads']:
        if prev and lv > prev + 1:
            jumps.append(f'h{prev}->h{lv} @ {t[:40]}')
        prev = lv
    if jumps:
        d.append({'type': 'heading_level_jump', 'detail': jumps[:8], 'count': len(jumps)})
    # 7. 强调碎片(转换噪声): 连续星号簇夹在词中
    frag = len(re.findall(r'\*\*\*\*', st['text']))
    if frag:
        d.append({'type': 'emphasis_fragment', 'detail': f'{frag} 处 **** 碎片',
                  'count': frag})
    return d


def readability(st, defects):
    """AI 可读性评分(0-100) + 扣分说明。
    评分维度(权重): 结构可导航性 40 / 表格完整性 25 / 资源完整性 20 / 文本洁净度 15
    """
    notes = []
    score = 100.0

    # 结构可导航性: 有标题且层级合理
    if st['headings'] == 0:
        score -= 40
        notes.append('全文无 Markdown 标题, AI 无法建立章节导航(-40)')
    else:
        md_tables = st['gfM_tables'] + st['html_tables']
        if md_tables > 20 and st['headings'] < md_tables / 10:
            score -= 15
            notes.append(f'标题数({st["headings"]})相对表格数({md_tables})明显偏少(-15)')
        jump = [x for x in defects if x['type'] == 'heading_level_jump']
        if jump:
            score -= 8
            notes.append(f'标题层级跳级 {jump[0]["count"]} 处(-8)')

    # 表格完整性
    colbad = [x for x in defects if x['type'] == 'table_column_mismatch']
    if colbad:
        score -= 25
        notes.append(f'{len(colbad)} 个 GFM 表格列数不齐, AI 解析会错列(-25)')

    # 资源完整性
    miss = [x for x in defects if x['type'] == 'missing_image']
    if miss:
        score -= 20
        notes.append(f'图片缺失 {miss[0]["count"]} 张(-20)')

    # 文本洁净度
    leak = [x for x in defects if x['type'] == 'placeholder_leak']
    if leak:
        score -= 10
        notes.append('残留占位符 ' + '、'.join(str(x['detail']) for x in leak) + '(-10)')
    fence = [x for x in defects if x['type'] == 'code_fence_unclosed']
    if fence:
        score -= 10
        notes.append('代码围栏未闭合, 其后内容会被误判为代码(-10)')
    empty = [x for x in defects if x['type'] == 'empty_heading']
    if empty:
        score -= 5
        notes.append(f'空标题 {len(empty)} 处(-5)')
    frag = [x for x in defects if x['type'] == 'emphasis_fragment']
    if frag:
        n = frag[0]['count']
        penalty = min(8, n / max(1, st['chars'] / 20000))
        score -= penalty
        notes.append(f'加粗碎片 {n} 处, 轻微干扰阅读节奏(-{penalty:.0f})')

    score = max(0.0, round(score, 1))
    grade = 'A' if score >= 90 else 'B' if score >= 75 else 'C' if score >= 60 else 'D'
    return {'score': score, 'grade': grade, 'notes': notes,
            'verdict': ('结构完整, AI 可直接阅读理解' if score >= 90 else
                        '可读, 存在局部结构瑕疵' if score >= 75 else
                        '可读性受影响, 需修正缺陷' if score >= 60 else
                        '可读性差, 建议重新转换')}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('md')
    ap.add_argument('--sample', type=int, default=6)
    args = ap.parse_args()

    st = md_structure(args.md)
    defects = detect_defects(st, os.path.dirname(args.md))
    rd = readability(st, defects)

    wpar, wtabs = word_parse(args.docx)
    mpar, mtabs = md_parse(args.md)

    # 对位抽样: 取 Word 段落流的等间距样本, 输出两侧内容便于比对
    samples = []
    n = len(wpar)
    if n:
        step = max(1, n // args.sample)
        for k in range(0, n, step):
            if len(samples) >= args.sample:
                break
            wv = wpar[k]
            mv = mpar[k] if k < len(mpar) else '<缺>'
            samples.append({'index': k, 'match': wv == mv,
                            'word': wv[:160], 'md': mv[:160]})

    out = {
        'docx': args.docx,
        'md': args.md,
        'md_path_rel': args.md,
        'structure': {k: v for k, v in st.items() if k not in ('text', 'lines_list', 'heads', 'imgs')},
        'headings_preview': st['heads'][:15],
        'defects': defects,
        'readability': rd,
        'paragraph_check': {
            'word_n': len(wpar), 'md_n': len(mpar),
            'equal': sum(1 for a, b in zip(wpar, mpar) if a == b),
        },
        'table_check': {
            'word_n': len(wtabs), 'md_n': len(mtabs),
            'equal': sum(1 for a, b in zip(wtabs, mtabs) if a == b),
        },
        'samples': samples,
    }
    print(json.dumps(out, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()