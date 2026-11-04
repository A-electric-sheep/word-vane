# -*- coding: utf-8 -*-
"""
内容级一致性校验 v2: Word 原文 vs 生成 Markdown
分层比对:
  A. 正文段落流: Word 非表格段落(去 TOC/图片/空) vs MD 非表格块(剥 markdown 语法)
  B. 表格流: 顶层表格逐个比对, 每表内单元格文本逐格对比(嵌套表文本并入所在格)
口径要点:
  - 两侧统一 norm: 去全部空白
  - 公式: 统一占位符 F0 处理
  - MD 侧剥离: 标题#、列表前缀(-/1./a./缩进)、强调(*single*/**double**)、HTML标签、$公式$、GFM管道
"""
import re
import zipfile
from difflib import SequenceMatcher
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W}
# 目录条目样式: 统一小写比对, 覆盖 'TOC 1' / 'toc 1' / 'TOC1' 三种 Word 写法
TOC_STYLES = {'toc 1', 'toc 2', 'toc 3', 'toc 4', 'toc 5',
              'toc1', 'toc2', 'toc3', 'toc4', 'toc5'}


def _is_toc_style_name(name):
    """按样式名判定目录样式(与转换器 blk.style.name in toc_level 口径对称)。
    部分模板目录样式用数字 styleId(如 '59' -> name 'toc 1'), 仅凭 ID 无法识别。"""
    return bool(re.match(r'^toc\s*\d+$', (name or '').strip(), re.I))


HEADING_STYLE_IDS = {'1', '2', '3', '4', '5', '6', 'Heading1', 'Heading2', 'Heading3',
                     'Heading4', 'Heading5', 'Heading6', 'Title', '标题1', '标题2', '标题3',
                     'soc1-1', 'soc1-2', 'soc1-3', 'soc1-4'}
F0 = '\x00F\x00'


def norm(s):
    """归一化: 去全部空白 + 剥装饰星号(Word 字面 * 与 MD 强调 ** 在内容层等值)"""
    t = re.sub(r'\s+', '', s)
    t = t.replace('*', '')
    return t


# ---------- Word 侧 ----------
def word_w_text(ch):
    """提取 w:t 文本, 跳过 ruby 注音 rt 与隐藏 run(webHidden) 的文本——与转换器口径一致"""
    parts = []
    for t in ch.iter('{%s}t' % W):
        # 注: ElementTree 无 getparent, 祖先过滤统一在下方 walk() 中完成
        parts.append(('t', t))
    # 重建: 按结构过滤
    out = []

    def walk(el):
        for child in el:
            full_tag = child.tag
            tag = full_tag.split('}')[-1]
            if tag == 'rt':
                continue  # 注音丢弃(与转换器一致)
            if tag == 'r':
                rPr = child.find('w:rPr', NS)
                if rPr is not None and rPr.find('w:webHidden', NS) is not None:
                    continue
            if full_tag == '{%s}t' % W:
                if child.text:
                    out.append(child.text)
            elif full_tag == '{%s}t' % M:
                continue  # 数学文本跳过(公式另以占位统计)
            else:
                walk(child)

    walk(ch)
    return norm(''.join(out))


def word_cell_text(tc):
    """单元格完整文本; 公式就地插占位(非追加末尾); 与转换器口径一致"""
    parts = []

    def run_text(r_el):
        ruby = r_el.find('{%s}ruby' % W)
        if ruby is not None:
            rb = ruby.find('{%s}rubyBase' % W)
            return ''.join(t.text or '' for t in rb.iter('{%s}t' % W)) if rb is not None else ''
        return ''.join(t.text or '' for t in r_el.iter('{%s}t' % W))

    def walk(el):
        for child in el:
            ftag = child.tag
            tag = ftag.split('}')[-1]
            if ftag == '{%s}oMath' % M or ftag == '{%s}oMathPara' % M:
                parts.append(F0)
            elif tag == 'rt':
                continue
            elif tag == 'r':
                rPr = child.find('w:rPr', NS)
                if rPr is not None and rPr.find('w:webHidden', NS) is not None:
                    continue
                parts.append(run_text(child))
            elif tag == 'hyperlink':
                parts.append(''.join(t.text or '' for t in child.iter('{%s}t' % W)))
            elif full_tag_w(child, 'pPr') or full_tag_w(child, 'rPr'):
                continue
            else:
                walk(child)

    walk(tc)
    return norm(''.join(parts))


def full_tag_w(el, name):
    return el.tag == '{%s}%s' % (W, name)


def word_parse(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
        # 建立 styleId -> 样式名映射: 目录样式在部分模板里用数字 ID(如 '59'->'toc 1'),
        # 需按样式名判定, 与转换器口径对称
        id2name = {}
        try:
            sroot = ET.fromstring(z.read('word/styles.xml'))
            for st in sroot.iter('{%s}style' % W):
                sid = st.get('{%s}styleId' % W)
                nm = st.find('{%s}name' % W)
                if sid is not None and nm is not None:
                    id2name[sid] = nm.get('{%s}val' % W, '')
        except KeyError:
            pass
    body = root.find('w:body', NS)
    paras = []
    tables = []

    def _handle(ch):
        nonlocal paras, tables
        tag = ch.tag.split('}')[-1]
        if tag == 'p':
            pPr = ch.find('w:pPr', NS)
            style = ''
            if pPr is not None:
                ps = pPr.find('w:pStyle', NS)
                if ps is not None:
                    style = ps.get('{%s}val' % W, '')
            if style in TOC_STYLES or _is_toc_style_name(id2name.get(style, '')):
                return
            has_math = ch.find('.//{%s}oMath' % M) is not None
            # 按 XML 元素顺序分段提取(F0 占位精确插位)
            def walk_p(el2):
                parts = []

                def run_text(r_el):
                    # run 内文本, 排除 ruby rt 注音
                    ruby = r_el.find('{%s}ruby' % W)
                    if ruby is not None:
                        rb = ruby.find('{%s}rubyBase' % W)
                        return ''.join(t.text or '' for t in rb.iter('{%s}t' % W)) if rb is not None else ''
                    return ''.join(t.text or '' for t in r_el.iter('{%s}t' % W))

                def w(el3):
                    for child in el3:
                        ftag = child.tag
                        tag = ftag.split('}')[-1]
                        if ftag == '{%s}oMath' % M or ftag == '{%s}oMathPara' % M:
                            parts.append(F0)
                        elif tag == 'rt':
                            continue
                        elif tag == 'r':
                            rPr = child.find('w:rPr', NS)
                            if rPr is not None and rPr.find('w:webHidden', NS) is not None:
                                continue
                            parts.append(run_text(child))
                        elif tag == 'hyperlink':
                            parts.append(''.join(t.text or '' for t in child.iter('{%s}t' % W)))
                        else:
                            w(child)
                w(el2)
                return norm(''.join(parts))
            txt = walk_p(ch)
            # 手动编号段落(无 numPr 的 "1. xxx")与 MD 列表前缀剥离对称
            pPr2 = ch.find('w:pPr', NS)
            has_numpr = pPr2 is not None and pPr2.find('w:numPr', NS) is not None
            if not has_numpr and style not in HEADING_STYLE_IDS:
                # 手动编号 "1. xxx"/"1) xxx" 与 MD 列表前缀剥离对称;
                # 标题与数值行(32.768k 等)不剥
                txt = re.sub(r'^\d+[.)]\s*(?!\d)', '', txt)
            # 图片段落
            if ch.find('.//{%s}pict' % W) is not None and not has_math and not norm(re.sub(F0, '', txt)):
                return
            if not txt and not has_math:
                return
            paras.append(txt)
        elif tag == 'tbl':
            # 顶层单元格: 直接子 tr 的直接 tc; 跳过 vMerge='cont' 的占位格与空文本格
            # (MD 侧 rowspan 合并后无 cont 格; GFM 表空单元格不输出, 口径统一跳过)
            cells = []
            for tr in ch.findall('{%s}tr' % W):
                for tc in tr.findall('{%s}tc' % W):
                    tcPr = tc.find('w:tcPr', NS)
                    if tcPr is not None:
                        v = tcPr.find('w:vMerge', NS)
                        if v is not None and (v.get('{%s}val' % W) in (None, '', 'continue')):
                            continue
                    text = word_cell_text(tc)
                    # 注意: 不能用 text.strip(F0)——strip 参数是字符集合, F0 含 'F' 会把 "F" 单元格剥空
                    if text != F0 and text.replace(F0, '').strip():
                        cells.append(text)
            tables.append(cells)

    for ch in body:
        if ch.tag.split('}')[-1] == 'sdt':
            sdt_content = ch.find('{%s}sdtContent' % W)
            if sdt_content is not None:
                for sub in sdt_content:
                    _handle(sub)
            continue
        _handle(ch)
    return paras, tables


# ---------- MD 侧 ----------
def strip_md_inline(s):
    """剥 markdown 行内语法: 强调/删除/HTML标签/公式/图片/代码"""
    t = s
    t = re.sub(r'\$\$[^$]*\$\$', F0, t)
    t = re.sub(r'\$([^$\n]+)\$', F0, t)
    # Markdown 链接 -> 仅标签文本。确定性括号配对(非正则):
    # 从右向左找 ')(' 结合点, 再向左配平嵌套 '['/']' 取标签段。
    # 覆盖两种真实场景: [REG_A1[1:0]...](#slug) 内嵌方括号标签,
    # 以及 "REG_B2[31:0]...可见[锚点](#slug)" 字面方括号后接链接。
    search_from = 0
    while True:
        m_close = t.find('](', search_from)
        if m_close == -1:
            break
        i = t.find(')', m_close + 2)
        if i == -1:
            break
        # 仅把 "(#锚点)" 视为链接目标。转换器生成的链接目标一律是 #slug;
        # 而寄存器位域紧跟括号(如 "REG_A1[31:0](n=0~3)")会被误判为链接,
        # 导致 "[31:0]" 方括号被吞掉, 必须排除(跳过该位置继续向后找)。
        if not t[m_close + 2:i].lstrip().startswith('#'):
            search_from = m_close + 2
            continue
        j = m_close
        depth = 0
        while j >= 0:
            ch = t[j]
            if ch == ']':
                depth += 1
            elif ch == '[':
                depth -= 1
                if depth == 0:
                    break
            j -= 1
        if j < 0 or depth != 0:
            # 无配平 '[', 放弃该处(如孤立的 "[31:0]" 字面量)
            break
        t = t[:j] + t[j + 1:m_close] + t[i + 1:]
        search_from = 0
    # 强调: 连续星号簇统一剥; 内容含中文的除外(源文档正文自带 * 字面星号, 非强调)
    def _em(m):
        inner = m.group(1)
        if re.search(r'[\u4e00-\u9fff]', inner):
            return '*' + inner + '*'
        return inner
    t = re.sub(r'\*{4,}([^*\n]+?)\*{4,}', _em, t)
    t = re.sub(r'\*\*\*([^*\n]+?)\*\*\*', _em, t)
    t = re.sub(r'\*\*([^*\n]+?)\*\*', _em, t)
    t = re.sub(r'(?<!\*)\*([^*\n]+?)\*(?!\*)', _em, t)
    t = re.sub(r'~~([^~\n]+?)~~', r'\1', t)
    t = re.sub(r'`([^`\n]+?)`', r'\1', t)
    # 仅剥离转换器实际输出的行内标签(sub/sup/s/br)。切勿用贪婪 <[^>]+>:
    # 正文常见 "2 <= N <=192"、"大于192" 等文案会被误当标签吞掉合法文字。
    t = re.sub(r'</?(?:sub|sup|s|br)\s*/?>', '', t)
    t = t.replace('\\|', '|')
    # 残留孤立 * 清除(仅非中文语境)
    t = re.sub(r'(?<![\u4e00-\u9fff])\*(?![\u4e00-\u9fff])', '', t)
    return t


def md_table_cells(table_html):
    """HTML 表 -> 顶层 td 序列(仅直接 td, 嵌套表文本并入所在格)"""
    from html.parser import HTMLParser

    class TopTdParser(HTMLParser):
        def __init__(self):
            super().__init__(convert_charrefs=True)
            self.table_depth = 0
            self.td_depth = 0
            self.cells = []
            self.buf = []

        def handle_starttag(self, tag, attrs):
            if tag == 'table':
                self.table_depth += 1
            if tag == 'td':
                if self.table_depth == 1 and self.td_depth == 0:
                    self.buf = []
                self.td_depth += 1

        def handle_endtag(self, tag):
            if tag == 'td':
                self.td_depth -= 1
                if self.table_depth == 1 and self.td_depth == 0:
                    # HTML 表格格内也可能含行内公式/强调/上下标, 需与段落同样
                    # 走 strip_md_inline, 否则 $$LaTeX$$ 无法归一为 F0 占位
                    n = norm(strip_md_inline(''.join(self.buf)))
                    if n.strip():
                        self.cells.append(n)
                    self.buf = []
            if tag == 'table':
                self.table_depth -= 1

        def handle_data(self, data):
            if self.table_depth >= 1 and self.td_depth >= 1:
                self.buf.append(data)

    p = TopTdParser()
    p.feed(table_html)
    return p.cells


def md_parse(md_path, return_alt=False):
    content = open(md_path, encoding='utf-8').read()
    lines = content.split('\n')
    paras = []
    tables = []
    in_code = False
    code_buf = []
    in_toc = False
    # 段落索引 -> 等价候选文本(未剥离列表标记的形态), 供宽松比对
    alt_map = {}

    def flush_code():
        nonlocal code_buf
        if code_buf:
            # 逐行不清洗(保持与 Word 逐段对齐), 每行一个块
            paras.extend(code_buf)
            code_buf = []

    i = 0
    toc_line = False   # 目录态内是否已遇到过清单条目(用于判定目录自然结束)
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if s.startswith('```'):
            if in_code:
                flush_code()
            in_code = not in_code
            i += 1
            continue
        if in_code:
            code_buf.append(norm(s))
            i += 1
            continue
        # 目录说明行(裸目录块, 无标题): 原文只有目录域、没有目录标题时,
        # 转换器直接输出清单, 首行是本说明(中英文两版); 从该行起进入目录态
        if (s.startswith('> 由 Word「目录」域转换的静态目录')
                or s.startswith('> Static table of contents converted from the Word TOC field')):
            in_toc = True
            i += 1
            continue
        # 目录块(标题可能被加粗包裹, 如 "# **目录**"; 英文版标题为 Content/Contents)
        if re.match(r'^#{1,4}\s*(?:目录|目\s*录|content|contents|table\s+of\s+contents)\s*$',
                    re.sub(r'[*_]+', '', s).strip(), re.I):
            in_toc = True
            paras.append(norm(s.lstrip('#').strip()))
            i += 1
            continue
        if in_toc:
            if s.startswith('>'):
                i += 1
                continue
            # 目录清单条目: "- xxx" / "- [xxx](#slug)" 及其缩进子项
            if re.match(r'^\s*[-*+]\s', s):
                toc_line = True
                i += 1
                continue
            # 遇到标题即目录结束; 该标题必须继续走下面的常规处理(不能被吞掉,
            # 否则终止目录的第一个正文标题会从 MD 侧缺失, 造成伪差异)
            if re.match(r'^#{1,6} ', s):
                in_toc = False
            elif toc_line:
                # 目录清单后出现的第一个非列表行 = 正文首段(正文标题可能没有
                # # 前缀, 如英文 Datasheet 的 "General description"), 目录到此结束
                in_toc = False
            else:
                i += 1
                continue
        # 图片行
        if s.startswith('!['):
            i += 1
            continue
        # HTML 表格: 跨行块(以 <table 起始, 按 table 开闭配对收集)
        if s.startswith('<table'):
            buf = s
            i += 1
            opens = buf.count('<table')
            closes = buf.count('</table>')
            while opens > closes and i < len(lines):
                nxt = lines[i]
                opens += nxt.count('<table')
                closes += nxt.count('</table>')
                buf += nxt
                i += 1
            tables.append(md_table_cells(buf))
            continue
        # GFM 表格: 收集连续管道行
        if s.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                rows.append(lines[i].strip())
                i += 1
            cells = []
            for r in rows:
                cs = [c.strip() for c in r.strip('|').split('|')]
                # 分隔符行必须是 --- 等 3+ 连字符(转换器生成格式); 单个 '-' 是
                # 位号表"保留位"数据格, 不能当分隔行丢弃
                if all(re.match(r'^:?-{3,}:?$', c) for c in cs):
                    continue
                for c in cs:
                    n = norm(strip_md_inline(c))
                    if n:
                        cells.append(n)
            tables.append(cells)
            continue
        # 标题(H1-H5)
        m_head = re.match(r'^(#{1,5})\s+(.+)$', s)
        if m_head:
            flush_code()
            paras.append(norm(strip_md_inline(m_head.group(2))))
            i += 1
            continue
        # 列表(剥离前缀)
        # 注意: Word 自动编号转出的列表标记有两种真实形态:
        #   "**1.**** ****用...**"(编号与点号同段加粗)
        #   "**2****. ****用...**"(编号与点号被加粗边界拆开)
        # 仅剥离星号与空白, 不可剥离下划线 —— 否则 "_nop_();" 这类标识符
        # 会被误伤成 "nop_();"。
        # 编号后的点号不可紧跟数字 —— 否则正文 "1.5Kbytes" 会被误剥成 "5Kbytes"。
        # 兜底剥离行首 # 标题标记: 转换器会把某些 Word 段落(如寄存器名)按
        # outlineLvl 提升为标题, 而 Word 侧该文本是普通段落, 若不剥 "###### "
        # 会造成纯前缀伪差异。
        s = re.sub(r'^#{1,6}\s*', '', s)
        s2 = re.sub(r'^(?:[*\s]*\d+[*\s]*\.(?!\d)[*\s]*|[*\s]+)', '', s)
        s2 = re.sub(r'^(?:[-*+]|[a-zA-Z]\.)\s*', '', s2)
        t = norm(strip_md_inline(s2))
        if t:
            flush_code()
            paras.append(t)
            # 记录"未剥标记"形态作为等价候选, 供 cmp_stream 做宽松匹配。
            # Word 侧编号可能来自 numPr(文本无 "1."), 也可能手写在正文里
            # (numId=0 取消自动编号), 后者若被剥离就与 Word 对不上。
            t_keep = norm(strip_md_inline(s))
            if t_keep and t_keep != t:
                alt_map.setdefault(len(paras) - 1, t_keep)
        i += 1
    flush_code()
    if return_alt:
        return paras, tables, alt_map
    return paras, tables


def _is_same(x, y, j, alt_b):
    """b 侧等价候选判定: 完全相等, 或该位置存在与 a 相等的等价文本。

    原为内联 lambda, 改为具名函数以通过 E731 并便于测试。
    """
    return (x == y) or bool(alt_b and alt_b.get(j) == x)


def cmp_stream(a, b, name_a, name_b, alt_b=None):
    """逐块比对两个段落流。

    alt_b: {索引: 等价文本} —— 同一块的可替代归一化形态(如"未剥离列表标记"
    版本)。当主形态与 Word 不等而替代形态相等时, 视为一致, 避免因纯标记
    差异产生伪差异。
    """
    ok = 0
    problems = []
    # 先把 b 的等价候选映射成可查表, 用于逐个修正 opcode 内的明细
    sm = SequenceMatcher(None, a, b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            ok += i2 - i1
            continue
        # replace/insert 区间里做逐对宽松匹配, 救回仅标记不同的块
        if tag in ('replace', 'insert') and (i2 - i1) == (j2 - j1):
            good = 0
            for k in range(i2 - i1):
                if _is_same(a[i1 + k], b[j1 + k], j1 + k, alt_b):
                    good += 1
            if good == (i2 - i1):
                ok += good
                continue
        problems.append((tag, i1, a[i1:i2], j1, b[j1:j2]))
    return ok, len(a) if a else 1, problems


def verify(docx_path=None, md_path=None) -> dict:
    """校验入口: Word 原文 vs 生成 Markdown 的内容级一致性。
    返回 dict:
      {paragraphs: {word_n, md_n, matched, pct, differences: [...]},
       tables:     {n_same, n_diff, cell_matched, cell_total, pct, diff_indices:[...]},
       formulas:   {word, md},
       images:     {word_refs, md_refs, ok},
       pass:       bool}
    口径要点(与转换器一致):
      - 空白归一 + 星号装饰剥离(Word 字面 * 与 MD 强调 ** 等值)
      - 链接/HTML标签剥离, 代码块逐行对齐
      - ruby rt 注音两侧丢弃; vMerge cont 占位格与空单元格两侧跳过
      - 嵌套表文本并入所在格; 公式统一 F0 占位精确插位
    """
    if not docx_path or not md_path:
        raise SystemExit('用法: python verify_consistency.py <输入.docx> <输出.md>')
    wpar, wtabs = word_parse(docx_path)
    mpar, mtabs, alt_map = md_parse(md_path, return_alt=True)

    ok, total, probs = cmp_stream(wpar, mpar, 'Word', 'MD', alt_b=alt_map)
    problems = [p for p in probs if p[0] != 'equal']

    n_same = n_diff = 0
    cell_same = cell_total = 0
    diff_indices = []
    for ti in range(max(len(wtabs), len(mtabs))):
        w = wtabs[ti] if ti < len(wtabs) else []
        m = mtabs[ti] if ti < len(mtabs) else []
        if w == m:
            n_same += 1
            cell_same += len(w)
            cell_total += max(len(w), len(m))
        else:
            n_diff += 1
            sm = SequenceMatcher(None, w, m, autojunk=False)
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag == 'equal':
                    cell_same += i2 - i1
            cell_total += max(len(w), len(m))
            diff_indices.append(ti)

    # 图片引用核对
    import os as _os
    content = open(md_path, encoding='utf-8').read()
    md_imgs = re.findall(r'!\[[^\]]*\]\(([^)]+)\)', content)
    missing = [p for p in md_imgs
               if not _os.path.isfile(_os.path.join(_os.path.dirname(md_path), p.replace('/', _os.sep)))]
    # Word 侧图片数(media 目录或 v:imagedata)
    w_img = sum(1 for t in wtabs for c in t if False)  # 表格无图片口径, 用 media 计数
    media_dir = _os.path.join(_os.path.dirname(md_path), 'images')
    md_files = len(_os.listdir(media_dir)) if _os.path.isdir(media_dir) else 0

    report = {
        'paragraphs': {
            'word_n': len(wpar), 'md_n': len(mpar),
            'matched': ok, 'pct': 100 * ok / max(len(wpar), 1),
            'differences': problems[:10],
        },
        'tables': {
            'n_total': len(wtabs), 'n_same': n_same, 'n_diff': n_diff,
            'cell_matched': cell_same, 'cell_total': cell_total,
            'pct': 100 * cell_same / max(cell_total, 1),
            'diff_indices': diff_indices,
        },
        'formulas': {
            'word': sum(1 for p in wpar if F0 in p),
            'md': sum(1 for p in mpar if F0 in p),
        },
        'images': {
            'word_refs': len(w_img) if w_img else None,
            'md_refs': len(md_imgs),
            'missing': missing,
            'files_on_disk': md_files,
        },
        'pass': (not problems and n_diff == 0
                 and sum(1 for p in wpar if F0 in p) == sum(1 for p in mpar if F0 in p)),
    }
    return report


def _print_report(r):
    print(f"Word 段落块: {r['paragraphs']['word_n']} | 表格: {r['tables']['n_total']} 个")
    print(f"MD   段落块: {r['paragraphs']['md_n']}")
    print()
    p = r['paragraphs']
    print(f"【A. 正文段落流】完全一致 {p['matched']}/{p['word_n']} = {p['pct']:.2f}% ({len(p['differences'])} 处差异)")
    for i, (tag, i1, wa, j1, ma) in enumerate(p['differences'][:10], 1):
        print(f"  [{i}] {tag} Word@{i1}: {[x[:45] for x in wa[:2]]}  MD@{j1}: {[x[:45] for x in ma[:2]]}")
    print()
    t = r['tables']
    print(f"【B. 表格流】{t['n_same']}/{t['n_total']} 个表格内容完全一致; {t['n_diff']} 个有差异")
    print(f"  单元格级一致率: {t['cell_matched']}/{t['cell_total']} = {t['pct']:.2f}%")
    if t['diff_indices']:
        print(f"  差异表索引: {t['diff_indices'][:12]}")
    print()
    f = r['formulas']
    print(f"【公式】Word 侧占位: {f['word']} | MD 侧占位: {f['md']} {'✓' if f['word'] == f['md'] else '✗'}")
    i = r['images']
    print(f"【图片】MD 引用 {i['md_refs']} 个, 缺失 {len(i['missing'])} 个, 磁盘文件 {i['files_on_disk']} 个")
    print()
    print("【总判定】", "通过 ✓" if r['pass'] else "存在差异 ✗")


def main(argv=None):
    """CLI 入口: 校验 Word 原文与生成 Markdown 的一致性。

    返回码: 0 = 全部通过, 1 = 存在差异(便于 CI/构建流程直接判定)。
    """
    import argparse

    from word2md import __version__
    ap = argparse.ArgumentParser(description='Word 与生成 Markdown 内容级一致性校验')
    ap.add_argument('--version', action='version', version=f'verify_consistency {__version__}')
    ap.add_argument('docx', nargs='?', help='输入 .docx 路径')
    ap.add_argument('md', nargs='?', help='生成 .md 路径')
    args = ap.parse_args(argv)
    if not args.docx or not args.md:
        ap.error('需要同时提供 .docx 与 .md 路径')
    report = verify(args.docx, args.md)
    _print_report(report)
    return 0 if report.get('pass') else 1


if __name__ == '__main__':
    raise SystemExit(main())
