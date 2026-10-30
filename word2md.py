# -*- coding: utf-8 -*-
"""
通用 Word(.docx) -> Markdown 转换库 + CLI
==========================================

将任意 Word 文档转换为 Markdown, 覆盖常见元素:
  标题(H1-H4+) / 段落 / 富文本(粗斜体/上下标/删除线/拼音注音) / 表格(含合并单元格
  colspan/rowspan 与嵌套表) / 图片(自动提取, EMF/WMF 光栅化为 PNG) / 有序与无序列表 /
  代码块(启发式识别) / 内部与外部超链接 / 数学公式(OMML->LaTeX) / 目录(原位还原)

用法
----
CLI:
    python word2md.py 输入.docx [-o 输出目录] [--no-images] [--no-code-heuristics] [--no-toc]

库调用:
    from word2md import convert_word_to_markdown
    stats = convert_word_to_markdown('input.docx', out_dir='out')
    stats['md_path']   # 输出 .md 路径
    stats['warnings']  # 转换警告列表
    stats['images']    # 提取图片数

适用与限制见 README.md
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys
import zipfile
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from math_latex import omml2latex  # OMML->LaTeX (单元级移植自 eqword2llm, MIT)

NS_VML = 'urn:schemas-microsoft-com:vml'

DEFAULT_HEADING_STYLES = {'Title': 1, 'Heading 1': 1, 'Heading 2': 2, 'Heading 3': 3,
                          'Heading 4': 4, 'Heading 5': 5, 'Heading 6': 6,
                          '标题 1': 1, '标题 2': 2, '标题 3': 3, '标题 4': 4, '标题 5': 5,
                          # 某匿名化 MCU 规格书自定义样式(内部项目 Datasheet 实测,
                          # [DESENSITIZED] 原样式族名为厂商私有命名, 此处保留通用样式名示例)
                          'soc1-1': 1, 'soc1-2': 2, 'soc1-3': 3, 'soc1-4': 4,
                          # 同系列 TRM 一级章节标题样式(实测: 简介/文档约定/资源框图等 32 处)
                          'SOC1': 1}
DEFAULT_TOC_STYLES = {'toc 1': 1, 'toc 2': 2, 'toc 3': 3, 'toc 4': 4, 'toc 5': 1, 'TOC1': 1,
                      'TOC2': 2, 'TOC3': 3, 'TOC4': 4}

CODE_PREFIXES = ('#include', '#define', 'unsigned', 'void ', 'static', 'typedef',
                 'int ', 'char ', 'long ', 'short ', 'float ', '/*', '//')
RE_CODE_ASSIGN = re.compile(
    r"^[A-Za-z_$][\w$]*\s*(?:\[[^\]]*\])?\s*=\s*[0-9A-Za-z_x$'\s\.\*\+\-\/\%\|&\[\]\(\)~]+"
    r"(?:\s*;\s*(?://.*)?)?$"
)
RE_CODE_CTRL = re.compile(r'^(?:\b(?:for|if|while|switch)\s*\(.*|\breturn\b.*|[\{\}])$')


def _looks_like_code(t: str) -> bool:
    if not t:
        return False
    for p in CODE_PREFIXES:
        if t.startswith(p):
            return True
    return bool(RE_CODE_ASSIGN.match(t) or RE_CODE_CTRL.match(t))


def _slugify(text: str) -> str:
    t = re.sub(r'[^\w\u4e00-\u9fff\- ]', '', text, flags=re.UNICODE)
    return t.strip().lower().replace(' ', '-')


class Converter:
    """单个 docx 文档的转换器实例。"""

    def __init__(self, src_path, *, code_heuristics=True, include_toc=True,
                 heading_styles=None, toc_styles=None, code_lang='c',
                 image_dir_name='images', strip_emphasis=False):
        self.src = str(src_path)
        self.code_heuristics = code_heuristics
        self.include_toc = include_toc
        self.code_lang = code_lang
        self.img_sub = image_dir_name
        # 剥离加粗/粗斜体(为 AI 阅读链路降噪); 不动文本内容与 sub/sup/s/br 功能标签
        self.strip_emphasis = strip_emphasis
        self.heading_level = dict(DEFAULT_HEADING_STYLES)
        if heading_styles:
            self.heading_level.update(heading_styles)
        self.toc_level = dict(DEFAULT_TOC_STYLES)
        if toc_styles:
            self.toc_level.update(toc_styles)

        self.warnings: list[str] = []
        self.doc = docx.Document(self.src)
        self.z = zipfile.ZipFile(self.src)

        # 编号系统(键务必为 int — 历史 bug: str 键导致编号列表全部退化为 bullet)
        self.num2abs, self.ABSTRACT = self._load_numbering()

        # 预扫描: 标题 bookmarks -> slug
        self.anchor2slug: dict[str, str] = {}
        self.anchor2title: dict[str, str] = {}
        self._prescan_headings()

        # 统计计数
        self.img_seq = 0
        self.heading_count = 0
        self.tables_gfm = 0
        self.tables_html = 0
        self.code_blocks = 0
        self.formula_blocks = 0

    # ---------- 辅助 ----------
    def _load_numbering(self):
        try:
            num_xml = self.z.read('word/numbering.xml').decode('utf8', 'ignore')
        except KeyError:
            return {}, {}
        num2abs = {int(k): int(v) for k, v in
                   re.findall(r'<w:num w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="(\d+)"', num_xml)}
        abstract = {}
        for aid, body in re.findall(r'<w:abstractNum w:abstractNumId="(\d+)"[^>]*>(.*?)</w:abstractNum>', num_xml, re.S):
            lvls = {}
            for m in re.finditer(r'<w:lvl w:ilvl="(\d+)"[^>]*>(.*?)</w:lvl>', body, re.S):
                ilvl = int(m.group(1))
                fmt = re.search(r'<w:numFmt w:val="([^"]+)"', m.group(2))
                lvls[ilvl] = fmt.group(1) if fmt else 'bullet'
            abstract[int(aid)] = lvls
        return num2abs, abstract

    def _num_fmt_of(self, p_el):
        pPr = p_el.find(qn('w:pPr'))
        if pPr is None:
            return None
        numPr = pPr.find(qn('w:numPr'))
        if numPr is None:
            return None
        nid = numPr.find(qn('w:numId'))
        ilvl = numPr.find(qn('w:ilvl'))
        if nid is None or nid.get(qn('w:val')) is None:
            return None
        return (int(nid.get(qn('w:val'))), int(ilvl.get(qn('w:val'))) if ilvl is not None else 0)

    def _blocks(self):
        """body 直接子元素块序列; w:sdt(Word 自动目录等结构化容器)下钻取其内容"""
        for child in self.doc.element.body.iterchildren():
            tag = child.tag
            if tag == qn('w:p'):
                yield Paragraph(child, self.doc)
            elif tag == qn('w:tbl'):
                yield Table(child, self.doc)
            elif tag == qn('w:sdt'):
                sdt_content = child.find(qn('w:sdtContent'))
                if sdt_content is not None:
                    for sub in sdt_content.iterchildren():
                        if sub.tag == qn('w:p'):
                            yield Paragraph(sub, self.doc)
                        elif sub.tag == qn('w:tbl'):
                            yield Table(sub, self.doc)

    def _prescan_headings(self):
        used = set()
        for blk in self._blocks():
            if not isinstance(blk, Paragraph):
                continue
            if blk.style.name not in self.heading_level:
                continue
            text = ''.join(t.text or '' for t in blk._p.findall('.//' + qn('w:t'))).strip()
            slug = _slugify(text) or 'sec'
            base, n = slug, 1
            while slug in used:
                slug = f'{base}-{n}'
                n += 1
            used.add(slug)
            for b in blk._p.findall('.//' + qn('w:bookmarkStart')):
                self.anchor2slug[b.get(qn('w:name'))] = slug
                self.anchor2title[b.get(qn('w:name'))] = text
            self.anchor2slug.setdefault('TEXT:' + text, slug)
            self.anchor2title.setdefault('TEXT:' + text, text)

    # ---------- 图片: 从 docx 内部提取, EMF/WMF 光栅化为 PNG ----------
    def extract_images(self, out_dir: str) -> int:
        """把文档内嵌图片落到 out_dir/<img_sub>/, 返回图片数。"""
        media_dir = os.path.join(out_dir, self.img_sub)
        os.makedirs(media_dir, exist_ok=True)
        n = 0
        for rId, rel in self.doc.part.rels.items():
            if 'image' not in rel.reltype:
                continue
            blob = rel.target_part.blob
            base = os.path.basename(rel.target_ref)
            stem, ext = os.path.splitext(base)
            ext_l = ext.lower()
            if ext_l == '':
                ext_l = '.png'
            if ext_l in ('.emf', '.wmf'):
                out_name = stem + '.png'
                try:
                    from PIL import Image
                    with Image.open(io.BytesIO(blob)) as im:
                        im2 = im.convert('RGB')
                        w, h = im2.size
                        im2 = im2.resize((w * 3, h * 3), Image.LANCZOS)
                        im2.save(os.path.join(media_dir, out_name))
                except Exception as e:
                    self.warnings.append(f'矢量图光栅化失败 {out_name}: {e}')
                    with open(os.path.join(media_dir, out_name), 'wb') as f:
                        f.write(blob)
                    out_name = stem + '.png'
            else:
                out_name = stem + ext_l
                with open(os.path.join(media_dir, out_name), 'wb') as f:
                    f.write(blob)
            n += 1
        return n

    def _img_rel_path(self, rId):
        rel = self.doc.part.rels[rId]
        base = os.path.basename(rel.target_ref)
        stem, ext = os.path.splitext(base)
        if ext.lower() in ('.emf', '.wmf'):
            return self.img_sub + '/' + stem + '.png'
        return self.img_sub + '/' + stem + (ext if ext else '.png')

    # ---------- 段落富文本(GFM 路径) ----------
    def _run_fmt(self, r):
        pr = r.find(qn('w:rPr'))
        bold = italic = False
        va = None
        if pr is not None:
            b = pr.find(qn('w:b'))
            bold = b is not None and b.get(qn('w:val')) not in ('0', 'false', 'none')
            i = pr.find(qn('w:i'))
            italic = i is not None and i.get(qn('w:val')) not in ('0', 'false', 'none')
            v = pr.find(qn('w:vertAlign'))
            if v is not None:
                va = v.get(qn('w:val'))
        return bold, italic, va

    def _wrap(self, txt, bold, italic, sub):
        if sub == 'subscript':
            txt = '<sub>' + txt + '</sub>'
        elif sub == 'superscript':
            txt = '<sup>' + txt + '</sup>'
        if self.strip_emphasis:
            return txt
        if bold and italic:
            txt = '***' + txt + '***'
        elif bold:
            txt = '**' + txt + '**'
        elif italic:
            txt = '*' + txt + '*'
        return txt

    def rich_runs(self, el):
        parts = []

        def walk(node):
            for child in node:
                tag = child.tag
                if tag == qn('w:r'):
                    rPr = child.find(qn('w:rPr'))
                    hidden = rPr is not None and rPr.find(qn('w:webHidden')) is not None
                    if hidden:
                        continue
                    b, i, v = self._run_fmt(child)
                    ruby = child.find(qn('w:ruby'))
                    if ruby is not None:
                        rb = ruby.find(qn('w:rubyBase'))
                        if rb is not None:
                            txt = ''.join(t.text or '' for t in rb.findall('.//' + qn('w:t')))
                            if txt:
                                parts.append(self._wrap(txt, b, i, v))
                        continue
                    txt = ''.join(t.text or '' for t in child.findall(qn('w:t')))
                    if txt:
                        parts.append(self._wrap(txt, b, i, v))
                elif tag == qn('w:hyperlink'):
                    anchor = child.get(qn('w:anchor'))
                    inner = []
                    for r in child.findall(qn('w:r')):
                        rPr = r.find(qn('w:rPr'))
                        hidden = rPr is not None and rPr.find(qn('w:webHidden')) is not None
                        if hidden:
                            continue
                        b, i, v = self._run_fmt(r)
                        t = ''.join(x.text or '' for x in r.findall(qn('w:t')))
                        if t:
                            inner.append(self._wrap(t, b, i, v))
                    label = ''.join(inner)
                    if anchor and anchor in self.anchor2slug:
                        parts.append(f'[{label}](#{self.anchor2slug[anchor]})')
                    else:
                        parts.append(label)
                elif tag == qn('w:tab'):
                    parts.append('\t')
                elif tag == qn('w:br'):
                    parts.append('\n')
                elif tag == qn('w:pict'):
                    for im in child.findall('.//{' + NS_VML + '}imagedata'):
                        rid = im.get(qn('r:id'))
                        if rid:
                            parts.append('[[IMG:' + rid + ']]')
                elif tag == qn('w:ruby'):
                    rb = child.find(qn('w:rubyBase'))
                    if rb is not None:
                        for r in rb.findall(qn('w:r')):
                            rPr = r.find(qn('w:rPr'))
                            hidden = rPr is not None and rPr.find(qn('w:webHidden')) is not None
                            if hidden:
                                continue
                            b, i, v = self._run_fmt(r)
                            t = ''.join(x.text or '' for x in r.findall(qn('w:t')))
                            if t:
                                parts.append(self._wrap(t, b, i, v))
                elif tag == qn('w:instrText') or tag == qn('w:fldChar'):
                    continue
                elif tag == qn('m:oMathPara'):
                    parts.append('$$' + omml2latex(child) + '$$')
                    self.formula_blocks += 1
                elif tag == qn('m:oMath'):
                    parts.append('$' + omml2latex(child) + '$')
                    self.formula_blocks += 1
                elif tag in (qn('w:sdt'), qn('w:sdtContent'), qn('w:ins'),
                             qn('w:smartTag'), qn('w:dir'), qn('w:noProof')):
                    walk(child)
                else:
                    walk(child)

        walk(el)
        return ''.join(parts)

    # ---------- 目录解析 ----------
    def _parse_toc_entry(self, p_el):
        anchor = None
        label = None
        # 路径 A: Word 标准目录, 条目由 <w:hyperlink w:anchor="_Toc…"> 包裹
        for hp in p_el.findall(qn('w:hyperlink')):
            a = hp.get(qn('w:anchor'))
            if a:
                anchor = a
                real = self.anchor2title.get(a, '')
                if real:
                    label = real
                break
        # 路径 B: 域代码目录(WPS 及部分 Word 生成), anchor 藏在
        #   <w:instrText> HYPERLINK \l "_Toc…" </w:instrText>, 无 hyperlink 元素
        if anchor is None:
            instrs = ''.join(x.text or '' for x in p_el.iter(qn('w:instrText')))
            m = re.search(r'_Toc\d+', instrs)
            if m:
                anchor = m.group(0)
                real = self.anchor2title.get(anchor, '')
                if real:
                    label = real
        # 兜底 1: hyperlink 元素内可见文本(去隐藏 run), 编号前缀剥离
        if label is None or not label.strip():
            texts = []
            for hp in p_el.findall(qn('w:hyperlink')):
                for run in hp.findall(qn('w:r')):
                    rPr = run.find(qn('w:rPr'))
                    hidden = rPr is not None and rPr.find(qn('w:webHidden')) is not None
                    if hidden:
                        continue
                    texts.append(''.join(x.text or '' for x in run.findall(qn('w:t'))))
            raw = ''.join(texts).replace('\t', '').strip()
            if raw:
                m = re.match(r'^(\d+(?:\.\d+)*)\s*(.*)$', raw, re.S)
                label = m.group(2).strip() if m else raw
        # 兜底 2: 域代码目录的段落文本, 形如 "编号\t标题\t页码"
        #   (w:tab 单独成元素, 需按 XML 顺序重建, 否则编号/标题/页码会粘连)
        if label is None or not label.strip():
            segs = []
            for node in p_el.iter():
                if node.tag == qn('w:t'):
                    segs.append(node.text or '')
                elif node.tag == qn('w:tab'):
                    segs.append('\t')
            flat = ''.join(segs).strip()
            parts = [p.strip() for p in flat.split('\t') if p.strip()]
            # 去首段编号(^\d+(\.\d+)*$) 与尾段页码(^\d+$)
            if parts and re.match(r'^\d+(?:\.\d+)*$', parts[0]):
                parts = parts[1:]
            if parts and re.match(r'^\d+$', parts[-1]):
                parts = parts[:-1]
            label = ' '.join(parts).strip()
        if not label or not label.strip() or label.strip() in ('目录', '目录（Contents）'):
            return None
        return label.strip(), anchor

    # ---------- 单元格 HTML 路径(合并表/嵌套表) ----------
    @staticmethod
    def _has_merge_cells(tbl_el):
        # return False
        return (tbl_el.find('.//' + qn('w:gridSpan')) is not None or
                tbl_el.find('.//' + qn('w:vMerge')) is not None or
                tbl_el.find('.//' + qn('w:tbl')) is not None)

    def _run_to_html(self, r_el, inherit=None):
        inherit = inherit or {}
        rPr = r_el.find(qn('w:rPr'))
        b = inherit.get('b', False)
        i = inherit.get('i', False)
        s = inherit.get('s', False)
        va = inherit.get('va')
        if rPr is not None:
            bb = rPr.find(qn('w:b'))
            if bb is not None:
                b = bb.get(qn('w:val')) not in ('0', 'false', 'none')
            ii = rPr.find(qn('w:i'))
            if ii is not None:
                i = ii.get(qn('w:val')) not in ('0', 'false', 'none')
            ss = rPr.find(qn('w:strike'))
            if ss is not None:
                s = ss.get(qn('w:val')) not in ('0', 'false', 'none')
            vv = rPr.find(qn('w:vertAlign'))
            if vv is not None:
                va = vv.get(qn('w:val'))
        ruby = r_el.find(qn('w:ruby'))
        if ruby is not None:
            rb = ruby.find(qn('w:rubyBase'))
            if rb is not None:
                txt = ''.join(t.text or '' for t in rb.findall('.//' + qn('w:t')))
                if txt:
                    txt = txt.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    if self.strip_emphasis:
                        return txt
                    return ('<strong>' + txt + '</strong>') if b else txt
            return ''
        txt = ''.join(t.text or '' for t in r_el.findall(qn('w:t')))
        if not txt:
            return ''
        txt = txt.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        if va == 'subscript':
            txt = '<sub>' + txt + '</sub>'
        elif va == 'superscript':
            txt = '<sup>' + txt + '</sup>'
        if self.strip_emphasis:
            pass
        elif b and i:
            txt = '<strong><em>' + txt + '</em></strong>'
        elif b:
            txt = '<strong>' + txt + '</strong>'
        elif i:
            txt = '<em>' + txt + '</em>'
        if s:
            txt = '<s>' + txt + '</s>'
        return txt

    def _para_run_inherit(self, p_el):
        pPr = p_el.find(qn('w:pPr'))
        if pPr is None:
            return {}
        mPr = pPr.find(qn('w:rPr'))
        if mPr is None:
            return {}
        inherit = {}
        for tag, key in ((qn('w:b'), 'b'), (qn('w:i'), 'i'), (qn('w:strike'), 's')):
            e = mPr.find(tag)
            if e is not None:
                inherit[key] = e.get(qn('w:val')) not in ('0', 'false', 'none')
        vv = mPr.find(qn('w:vertAlign'))
        if vv is not None:
            inherit['va'] = vv.get(qn('w:val'))
        return inherit

    def _para_content_html(self, p_el):
        parts = []
        inherit = self._para_run_inherit(p_el)

        def walk(el):
            for child in el:
                tag = child.tag
                if tag == qn('w:r'):
                    parts.append(self._run_to_html(child, inherit))
                elif tag == qn('w:hyperlink'):
                    anchor = child.get(qn('w:anchor'))
                    label = ''.join(self._run_to_html(r, inherit) for r in child.findall(qn('w:r')))
                    if anchor and anchor in self.anchor2slug:
                        parts.append(f'<a href="#{self.anchor2slug[anchor]}">{label}</a>')
                    else:
                        parts.append(label)
                elif tag == qn('w:ruby'):
                    rb = child.find(qn('w:rubyBase'))
                    if rb is not None:
                        for r in rb.findall(qn('w:r')):
                            parts.append(self._run_to_html(r, dict(inherit)))
                elif tag == qn('w:bookmarkStart'):
                    name = child.get(qn('w:name'))
                    if name:
                        parts.append(f'<a id="{name}"></a>')
                elif tag == qn('w:tab'):
                    parts.append(' ')
                elif tag == qn('w:br'):
                    parts.append('<br>')
                elif tag == qn('m:oMathPara'):
                    parts.append('$$' + omml2latex(child) + '$$')
                elif tag == qn('m:oMath'):
                    parts.append('$' + omml2latex(child) + '$')
                elif tag in (qn('w:sdt'), qn('w:sdtContent'), qn('w:ins'),
                             qn('w:smartTag'), qn('w:dir')):
                    walk(child)
                elif tag == qn('w:pPr') or tag == qn('w:rPr'):
                    continue

        walk(p_el)
        return ''.join(parts)

    def _cell_html_blocks(self, tc_el):
        out = []
        pending = []

        def flush():
            if pending:
                fmt = pending[0][0]
                tag = 'ul' if fmt == 'bullet' else 'ol'
                items = ''.join('<li>' + h + '</li>' for _, h in pending)
                out.append(f'<{tag}>{items}</{tag}>')
                pending.clear()

        for direct in tc_el:
            if direct.tag == qn('w:tbl'):
                flush()
                ig = direct.find(qn('w:tblGrid'))
                icols = max(1, len(ig.findall(qn('w:gridCol'))) if ig is not None else 1)
                out.append('<table border="1">' + self._html_table_rows(direct, icols) + '</table>')
            elif direct.tag == qn('w:p'):
                inner = self._para_content_html(direct)
                nf = self._num_fmt_of(direct)
                if nf:
                    fmt = self.ABSTRACT.get(self.num2abs.get(nf[0], -1), {}).get(nf[1], 'bullet')
                    pending.append((fmt, inner))
                else:
                    flush()
                    if inner.strip():
                        out.append('<p>' + inner + '</p>')
        flush()
        return ''.join(out)

    def _html_table_rows(self, tbl_el, grid_cols):
        trs = tbl_el.findall(qn('w:tr'))
        n_rows = len(trs)
        if n_rows == 0:
            return ''

        cellinfo = []
        for tr in trs:
            info = []
            for tc in tr.findall(qn('w:tc')):
                tcPr = tc.find(qn('w:tcPr'))
                span = 1
                vm = None
                if tcPr is not None:
                    gs = tcPr.find(qn('w:gridSpan'))
                    if gs is not None:
                        try:
                            span = max(1, int(gs.get(qn('w:val'))))
                        except (TypeError, ValueError):
                            span = 1
                    v = tcPr.find(qn('w:vMerge'))
                    if v is not None:
                        vm = v.get(qn('w:val')) or 'cont'
                content = self._cell_html_blocks(tc)
                info.append({'span': span, 'vm': vm, 'html': content})
            cellinfo.append(info)

        occ = [[None] * grid_cols for _ in range(n_rows)]
        starts = []
        for r_i, info in enumerate(cellinfo):
            col = 0
            row_starts = [0] * len(info)
            for c_i, rec in enumerate(info):
                while col < grid_cols and occ[r_i][col] is not None:
                    col += 1
                row_starts[c_i] = col
                is_cont = rec['vm'] == 'cont'
                for k in range(rec['span']):
                    if col + k < grid_cols:
                        occ[r_i][col + k] = ('cont' if is_cont else 'main', c_i)
                col += rec['span']
            starts.append(row_starts)

        rowspan = []
        emit = []
        for r_i, info in enumerate(cellinfo):
            rs_row = []
            em_row = []
            for c_i, rec in enumerate(info):
                if rec['vm'] == 'restart':
                    sc = starts[r_i][c_i]
                    rs = 1
                    r2 = r_i + 1
                    while r2 < n_rows and sc < grid_cols:
                        owner = occ[r2][sc]
                        if owner is not None and owner[0] == 'cont' and starts[r2][owner[1]] == sc:
                            rs += 1
                            r2 += 1
                        else:
                            break
                    rs_row.append(rs)
                    em_row.append(True)
                elif rec['vm'] == 'cont':
                    rs_row.append(1)
                    em_row.append(False)
                else:
                    rs_row.append(1)
                    em_row.append(True)
            rowspan.append(rs_row)
            emit.append(em_row)

        out = []
        for r_i, info in enumerate(cellinfo):
            tds = []
            for c_i, rec in enumerate(info):
                if not emit[r_i][c_i]:
                    continue
                attrs = ''
                if rec['span'] > 1:
                    attrs += f' colspan="{rec["span"]}"'
                rs = rowspan[r_i][c_i]
                if rs > 1:
                    attrs += f' rowspan="{rs}"'
                tds.append('<td' + attrs + '>' + rec['html'] + '</td>')
            out.append('<tr>' + ''.join(tds) + '</tr>')
        return '\n'.join(out)

    # ---------- 表格入口 ----------
    def _render_inner_table(self, tc_el):
        inner = tc_el.find(qn('w:tbl'))
        if inner is None:
            return None
        ig = inner.find(qn('w:tblGrid'))
        icols = max(1, len(ig.findall(qn('w:gridCol'))) if ig is not None else 1)
        return '<table border="1">' + self._html_table_rows(inner, icols) + '</table>'

    def _cell_text(self, tc_el):
        parts = []
        for direct in tc_el:
            if direct.tag == qn('w:p'):
                parts.append(self.rich_runs(direct))
        # 多段落以换行连接(输出管线将 \n 转 <br>), 避免多段说明(如某寄存器分频配置逐项)
        # 粘成一行; 空段跳过防止产生孤立 <br>
        return '\n'.join(p.strip() for p in parts if p.strip()).strip()

    def _render_table(self, tbl_el):
        if self._has_merge_cells(tbl_el):
            self.tables_html += 1
            grid_el = tbl_el.find(qn('w:tblGrid'))
            grid_cols = len(grid_el.findall(qn('w:gridCol'))) if grid_el is not None else 1
            grid_cols = max(grid_cols, 1)
            inner = self._html_table_rows(tbl_el, grid_cols)
            return ['<table border="1">' + inner + '</table>']
        self.tables_gfm += 1
        grid_el = tbl_el.find(qn('w:tblGrid'))
        grid_cols = len(grid_el.findall(qn('w:gridCol'))) if grid_el is not None else 1
        grid_cols = max(grid_cols, 1)
        lines = []
        rows = []
        for tr in tbl_el.findall(qn('w:tr')):
            rowline = [None] * grid_cols
            col = 0
            for tc in tr.findall(qn('w:tc')):
                tcPr = tc.find(qn('w:tcPr'))
                span = 1
                vmerge = None
                if tcPr is not None:
                    gs = tcPr.find(qn('w:gridSpan'))
                    if gs is not None:
                        try:
                            span = max(1, int(gs.get(qn('w:val'))))
                        except (TypeError, ValueError):
                            span = 1
                    vm = tcPr.find(qn('w:vMerge'))
                    if vm is not None:
                        vmerge = vm.get(qn('w:val')) or 'cont'
                while col < grid_cols and rowline[col] is not None:
                    col += 1
                if col >= grid_cols:
                    break
                html = self._render_inner_table(tc)
                if html is not None:
                    content = ('HTML', html)
                else:
                    content = ('TXT', self._cell_text(tc))
                rec = {'vmerge': vmerge, 'content': content}
                rowline[col] = rec
                for k in range(1, span):
                    if col + k < grid_cols:
                        rowline[col + k] = ('SPAN',)
                col += span
            rows.append(rowline)

        used = [False] * grid_cols
        for rowline in rows:
            for c in range(grid_cols):
                cell = rowline[c]
                if cell is None or isinstance(cell, tuple):
                    continue
                if cell['vmerge'] == 'cont':
                    continue
                ktype, val = cell['content']
                if ktype == 'HTML' or val.strip():
                    used[c] = True
        keep_right = grid_cols
        while keep_right > 1 and not used[keep_right - 1]:
            keep_right -= 1

        for r_i, rowline in enumerate(rows):
            cells = []
            for c in range(keep_right):
                cell = rowline[c]
                if cell is None or isinstance(cell, tuple) or cell['vmerge'] == 'cont':
                    cells.append('')
                    continue
                ktype, val = cell['content']
                cells.append(val)
            cells = [c.replace('|', '\\|').replace('\n', '<br>').replace('\t', ' ') for c in cells]
            lines.append('| ' + ' | '.join(cells) + ' |')
            if r_i == 0:
                lines.append('|' + '---|' * keep_right)
        return lines

    def _para_imagedata(self, p_el):
        return p_el.findall('.//{' + NS_VML + '}imagedata')

    # ---------- 主转换 ----------
    def convert(self, out_dir: str, extract_images: bool = True) -> dict:
        os.makedirs(out_dir, exist_ok=True)
        n_images = self.extract_images(out_dir) if extract_images else 0
        if extract_images and n_images == 0:
            self.warnings.append('文档内未发现任何图片')

        lines: list[str] = []
        toc_entries = []
        list_counters: dict = {}
        code_buf: list[str] = []
        pending_blank = False

        def flush_code():
            nonlocal code_buf
            if len(code_buf) >= 1:
                lines.append('```' + self.code_lang)
                for c in code_buf:
                    lines.append(c)
                lines.append('```')
                lines.append('')
                self.code_blocks += 1
            code_buf = []

        for blk in self._blocks():
            if isinstance(blk, Paragraph):
                style = blk.style.name
                p_el = blk._p
                pure = ''.join(t.text or '' for t in p_el.findall('.//' + qn('w:t'))).strip()
                txt = self.rich_runs(p_el).strip()
                imgs = self._para_imagedata(p_el)

                if imgs:
                    flush_code()
                    if txt and '[[IMG:' not in txt:
                        lines.append(txt)
                        lines.append('')
                    for im in imgs:
                        rId = im.get(qn('r:id'))
                        if not rId:
                            continue
                        self.img_seq += 1
                        lines.append(f'![图 {self.img_seq}]({self._img_rel_path(rId)})')
                        lines.append('')
                    continue

                if not pure:
                    if code_buf:
                        pending_blank = True
                        continue
                    if txt:
                        flush_code()
                        lines.append(txt)
                        lines.append('')
                        continue
                    flush_code()
                    lines.append('')
                    continue

                if style in self.toc_level and self.include_toc:
                    entry = self._parse_toc_entry(p_el)
                    if entry:
                        label, anchor = entry
                        target = ''
                        if anchor:
                            target = self.anchor2slug.get(anchor, '')
                        if not target:
                            target = self.anchor2slug.get('TEXT:' + label, '')
                        if target == '#':
                            target = ''
                        toc_entries.append((self.toc_level[style], label, target))
                    continue

                if style in self.heading_level:
                    flush_code()
                    lv = self.heading_level[style]
                    lines.append('#' * lv + ' ' + txt)
                    lines.append('')
                    self.heading_count += 1
                    if pure == '目录' and self.include_toc:
                        lines.append('<<TOC_PLACEHOLDER>>')
                        lines.append('')
                    continue

                nf = self._num_fmt_of(p_el)
                if nf:
                    flush_code()
                    nid, ilvl = nf
                    fmt = self.ABSTRACT.get(self.num2abs.get(nid, -1), {}).get(ilvl, 'bullet')
                    key = (nid, ilvl)
                    list_counters[key] = list_counters.get(key, 0) + 1
                    idx = list_counters[key]
                    indent = '  ' * ilvl
                    if fmt == 'bullet':
                        marker = '- '
                    elif fmt == 'decimal':
                        marker = f'{idx}. '
                    elif fmt == 'lowerLetter':
                        marker = chr(ord('a') + idx - 1) + '. '
                    elif fmt == 'upperLetter':
                        marker = chr(ord('A') + idx - 1) + '. '
                    else:
                        marker = f'{idx}. '
                    lines.append(indent + marker + txt)
                    continue

                if self.code_heuristics and _looks_like_code(pure) and 'IMG:' not in txt:
                    if pending_blank:
                        code_buf.append('')
                        pending_blank = False
                    code_buf.append(txt)
                    continue
                else:
                    pending_blank = False
                    flush_code()

                lines.append(txt)
                lines.append('')
            else:
                flush_code()
                lines.extend(self._render_table(blk._tbl))
                lines.append('')

        flush_code()

        body = '\n'.join(lines)

        # 目录块原位注入
        toc_note = '> 由 Word「目录」域转换的静态目录，点击条目可跳转到对应章节。'
        toc_body_lines = [toc_note, '']
        for lv, label, target in toc_entries:
            indent = '  ' * max(0, lv - 1)
            if target:
                toc_body_lines.append(f'{indent}- [{label}](#{target})')
            else:
                toc_body_lines.append(f'{indent}- {label}')
        toc_body = '\n'.join(toc_body_lines)
        placeholder = '<<TOC_PLACEHOLDER>>'
        if self.include_toc:
            if placeholder in body:
                body = body.replace(placeholder, toc_body, 1)
            elif not toc_entries and '目录' not in body:
                # 无目录域且文档里没有「目录」标题: 完全跳过目录输出
                pass
            elif '目录' not in body:
                body = '## 目录\n\n' + toc_body + '\n\n' + body
        else:
            body = body.replace(placeholder, '', 1) if placeholder in body else body

        stem = os.path.splitext(os.path.basename(self.src))[0]
        md_path = os.path.join(out_dir, stem + '.md')
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(body)

        self.z.close()
        return {
            'src': self.src,
            'md_path': md_path,
            'out_dir': out_dir,
            'images': n_images,
            'headings': self.heading_count,
            'tables': self.tables_gfm + self.tables_html,
            'tables_gfm': self.tables_gfm,
            'tables_html': self.tables_html,
            'code_blocks': self.code_blocks,
            'formulas': self.formula_blocks,
            'toc_entries': len(toc_entries),
            'warnings': self.warnings,
        }


def convert_word_to_markdown(src_path, out_dir=None, **kwargs) -> dict:
    """入口: 转换任意 .docx 文档为 Markdown。
    参数:
        src_path: 输入 .docx 路径
        out_dir:  输出目录(缺省 = 输入文件旁 <名>_md 目录)
        extract_images:   提取内嵌图片(默认 True)
        code_heuristics:  代码块启发式识别(默认 True)
        include_toc:      目录域处理(默认 True)
        heading_styles:   额外标题样式映射 {样式名: 级别}
        code_lang:        代码块语言标记(默认 'c')
    返回: 统计 dict(含 md_path / warnings / 各类计数)"""
    src_path = str(src_path)
    if not os.path.isfile(src_path):
        raise FileNotFoundError(src_path)
    if not src_path.lower().endswith(('.docx', '.doc')):
        # .doc 需先转 docx, 给出明确提示
        if src_path.lower().endswith('.doc'):
            raise ValueError(
                '旧版 .doc 二进制格式不被本工具直接支持。请先用 Word/WPS 另存为 .docx, '
                '或使用 LibreOffice: soffice --headless --convert-to docx 输入.doc')
        raise ValueError(f'不支持的输入格式: {src_path}, 仅支持 .docx')

    conv = Converter(src_path,
                     code_heuristics=kwargs.get('code_heuristics', True),
                     include_toc=kwargs.get('include_toc', True),
                     heading_styles=kwargs.get('heading_styles'),
                     toc_styles=kwargs.get('toc_styles'),
                     code_lang=kwargs.get('code_lang', 'c'),
                     image_dir_name=kwargs.get('image_dir_name', 'images'),
                     strip_emphasis=kwargs.get('strip_emphasis', False))
    if out_dir is None:
        out_dir = os.path.join(os.path.dirname(src_path),
                               os.path.splitext(os.path.basename(src_path))[0] + '_md')
    return conv.convert(str(out_dir), extract_images=kwargs.get('extract_images', True))


def main(argv=None):
    ap = argparse.ArgumentParser(description='通用 Word(.docx) -> Markdown 转换器')
    ap.add_argument('input', help='输入 .docx 文件')
    ap.add_argument('-o', '--output-dir', default=None, help='输出目录(缺省: 输入旁 <名>_md)')
    ap.add_argument('--no-images', action='store_true', help='不提取图片')
    ap.add_argument('--no-code-heuristics', action='store_true', help='关闭代码块启发式')
    ap.add_argument('--no-toc', action='store_true', help='关闭目录域处理')
    ap.add_argument('--no-bold', action='store_true', help='剥离加粗/粗斜体强调语法(为 AI 阅读降噪, 文本内容与 sub/sup/br 功能标签保留)')
    ap.add_argument('--code-lang', default='c', help='代码块语言标记(默认 c)')
    args = ap.parse_args(argv)

    stats = convert_word_to_markdown(
        args.input, out_dir=args.output_dir,
        extract_images=not args.no_images,
        code_heuristics=not args.no_code_heuristics,
        include_toc=not args.no_toc,
        code_lang=args.code_lang,
        strip_emphasis=args.no_bold)

    print(f"转换完成: {stats['md_path']}")
    print(f"  标题 {stats['headings']} | 表格 {stats['tables']} (GFM {stats['tables_gfm']}/HTML {stats['tables_html']})"
          f" | 图片 {stats['images']} | 代码块 {stats['code_blocks']} | 公式 {stats['formulas']}"
          f" | 目录条目 {stats['toc_entries']}")
    for w in stats['warnings']:
        print(f"  ⚠ {w}")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())