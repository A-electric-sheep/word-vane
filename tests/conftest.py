# -*- coding: utf-8 -*-
"""测试夹具构造工具。

夹具一律**在运行时用 python-docx 动态生成**，仓库中不保存任何 .docx 二进制
文件。这样做的原因有两条：

1. 二进制夹具在 Git 中不可审阅，也无法体现"它测的是什么"；
2. 公开仓库不应携带任何真实文档，动态生成的合成夹具天然不含敏感信息。

每个 builder 只组合"结构性能力"，不含任何真实业务文案。
"""
from __future__ import annotations

import os

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


def _new_doc():
    """返回一个空白文档，正文字体设为常见中文字体，避免字体缺失告警。"""
    d = docx.Document()
    style = d.styles['Normal']
    style.font.name = 'SimSun'
    style.font.size = Pt(10.5)
    return d


def build_merged_table(path):
    """夹具：含横跨(colspan)、纵跨(rowspan)与无合并三种表格。

    这是本项目区别于 GFM 路线工具的核心能力——合并语义必须在 Markdown 中
    以 colspan/rowspan 形式可恢复，而非被压成留空的普通网格。
    """
    d = _new_doc()
    d.add_heading('合并单元格夹具', 1)

    # 表1：colspan + rowspan 混合
    t = d.add_table(rows=3, cols=4)
    t.style = 'Table Grid'
    head = t.cell(0, 0).merge(t.cell(0, 2))
    head.text = '横向合并标题'
    t.cell(0, 3).text = '备注'
    v = t.cell(1, 0).merge(t.cell(2, 0))
    v.text = '纵向合并标签'
    t.cell(1, 1).text = 'A1'
    t.cell(1, 2).text = 'A2'
    t.cell(1, 3).text = 'A3'
    t.cell(2, 1).text = 'B1'
    t.cell(2, 2).text = 'B2'
    t.cell(2, 3).text = 'B3'

    d.add_paragraph('两张表之间的说明段。')

    # 表2：无合并，应走 GFM 管道表
    t2 = d.add_table(rows=2, cols=2)
    t2.style = 'Table Grid'
    t2.cell(0, 0).text = '参数'
    t2.cell(0, 1).text = '取值'
    t2.cell(1, 0).text = '工作电压'
    t2.cell(1, 1).text = '3.3V'

    d.save(path)
    return path


def build_omml_formula(path):
    """夹具：行内与块级 OMML 公式。

    OMML 由 Office Math 命名空间下的 m:oMath 承载，python-docx 无高层 API，
    因此直接注入 XML。转换器应产出 $...$ 与 $$...$$。
    """
    d = _new_doc()
    d.add_heading('公式夹具', 1)
    d.add_paragraph('行内公式示例：')
    p = d.add_paragraph()

    # 构造一个最简 OMML: x = y
    def make_omml(display: bool):
        xml = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            '<m:r><m:t>x</m:t></m:r>'
            '<m:r><m:t>=</m:t></m:r>'
            '<m:r><m:t>y</m:t></m:r>'
            '</m:oMath>'
        )
        from lxml import etree
        el = etree.fromstring(xml)
        if display:
            wrapper = OxmlElement('m:oMathPara')
            wrapper.append(el)
            return wrapper
        return el

    p._p.append(make_omml(display=False))

    d.add_paragraph('块级公式示例：')
    pd = d.add_paragraph()
    pd._p.append(make_omml(display=True))

    d.save(path)
    return path


def build_toc(path):
    """夹具：Word 目录域（TOC field）+ 若干标题。

    转换器应把目录解析为可点击锚点列表，且**保留在原位置**而非搬到文首。

    注意：真实 Word 的 TOC 域在 separate 与 end 之间必然有若干条**结果条目**
    （形如「1  第一章 概述\t1」）。没有这些条目时，域是空的、无从解析——所以
    夹具必须补上，否则测的就不是"目录解析"而是"空域处理"。
    """
    d = _new_doc()
    d.add_heading('目录夹具', 1)

    def fld_run(tag):
        r = OxmlElement('w:r')
        el = OxmlElement('w:fldChar')
        el.set(qn('w:fldCharType'), tag)
        r.append(el)
        return r

    # python-docx 默认模板不含 TOC 1/TOC 2 样式，需先注册，
    # 否则转换器无处识别"这是目录条目"
    from docx.enum.style import WD_STYLE_TYPE
    for lv in (1, 2):
        name = f'TOC {lv}'
        try:
            d.styles[name]
        except KeyError:
            st = d.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            st.base_style = d.styles['Normal']

    def toc_entry(text, num, page, level):
        """构造一条目录结果段落：TOC 样式 + 编号/标题/页码以 tab 分隔。"""
        p = d.add_paragraph(style=f'TOC {level}')
        p.add_run(f'{num}\t{text}\t{page}')
        return p

    # 域开始
    p_begin = d.add_paragraph()
    p_begin._p.append(fld_run('begin'))
    r_instr = OxmlElement('w:r')
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = 'TOC \\o "1-2" \\h \\z \\u'
    r_instr.append(it)
    p_begin._p.append(r_instr)
    p_begin._p.append(fld_run('separate'))

    # 域内实际条目（Word 生成的目录结果）
    toc_entry('目录夹具', '1', '1', 1)
    toc_entry('第一章 概述', '2', '2', 1)
    toc_entry('1.1 小节', '3', '3', 2)
    toc_entry('1.2 小节', '4', '3', 2)

    # 域结束
    p_end = d.add_paragraph()
    p_end._p.append(fld_run('end'))

    # 正文标题（目录条目应指向它们）
    d.add_heading('第一章 概述', 1)
    d.add_paragraph('第一章正文。')
    d.add_heading('1.1 小节', 2)
    d.add_paragraph('小节正文。')

    d.save(path)
    return path


def build_numbered_list(path):
    """夹具：真正的多级编号列表（同一 numId，ilvl 0/1 交替）。

    注意：Word 默认模板里的 `List Number` / `List Number 2` 是**各自独立的
    单层编号定义**（每个 abstractNum 只有 ilvl=0），并非同一列表的第 1/2 级。
    因此这里手工构造一份含 ilvl 0 与 ilvl 1 的 abstractNum，再把段落按
    ilvl 挂到同一 numId 上——只有这样才能测出"层级还原"这一能力。
    """
    d = _new_doc()
    d.add_heading('列表夹具', 1)

    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    numbering = d.part.numbering_part.element
    from lxml import etree

    # 追加一个两级 abstractNum
    max_aid = max([int(a.get(qn('w:abstractNumId')))
                   for a in numbering.findall(qn('w:abstractNum'))] or [-1])
    aid = max_aid + 1
    abs_xml = (
        f'<w:abstractNum xmlns:w="{W}" w:abstractNumId="{aid}">'
        '<w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>'
        '<w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl>'
        '</w:abstractNum>'
    )
    numbering.append(etree.fromstring(abs_xml))

    # 追加一个 num 指向它
    max_nid = max([int(n.get(qn('w:numId')))
                   for n in numbering.findall(qn('w:num'))] or [0])
    nid = max_nid + 1
    num_xml = (
        f'<w:num xmlns:w="{W}" w:numId="{nid}">'
        f'<w:abstractNumId w:val="{aid}"/></w:num>'
    )
    numbering.append(etree.fromstring(num_xml))

    def add_item(text, ilvl):
        p = d.add_paragraph(text)
        pPr = p._p.get_or_add_pPr()
        numPr = OxmlElement('w:numPr')
        il = OxmlElement('w:ilvl')
        il.set(qn('w:val'), str(ilvl))
        ni = OxmlElement('w:numId')
        ni.set(qn('w:val'), str(nid))
        numPr.append(il)
        numPr.append(ni)
        pPr.append(numPr)
        return p

    add_item('第一级条目', 0)
    add_item('第二级条目', 1)
    add_item('第二级条目二', 1)
    add_item('第一级条目二', 0)

    d.save(path)
    return path


def build_ruby(path):
    """夹具：拼音注音(ruby)。

    Word 的中文注音以 w:ruby 承载，转换器应只输出主体文本、丢弃注音。
    """
    d = _new_doc()
    d.add_heading('注音夹具', 1)
    d.add_paragraph('带注音的词语：')
    p2 = d.add_paragraph()

    xml = (
        '<w:ruby xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:rubyPr><w:rubyAlign w:val="distributeSpace"/>'
        '<w:hps w:val="10"/><w:hpsRaise w:val="20"/><w:hpsBaseText w:val="20"/>'
        '<w:lid w:val="zh-CN"/></w:rubyPr>'
        '<w:rt><w:r><w:t>han</w:t></w:r></w:rt>'
        '<w:rubyBase><w:r><w:t>汉</w:t></w:r></w:rubyBase>'
        '</w:ruby>'
    )
    from lxml import etree
    el = etree.fromstring(xml)
    p2._p.append(el)

    d.save(path)
    return path


def build_code_block(path):
    """夹具：可被启发式命中的 C 代码段（含缩进）。

    注意：这里必须真的写入前导空格。Word 里"看起来缩进"若只靠段落首行缩进
    属性（w:ind），w:t 文本本身没有空格，转换器无从还原缩进——那属于源文档
    写法差异，不是转换器缺陷。
    """
    d = _new_doc()
    d.add_heading('代码夹具', 1)
    for line in [
        '#include <stdio.h>',
        'int main(void) {',
        '    int n = 0;',
        '    for (n = 0; n < 3; n++) {',
        '        printf("%d", n);',
        '    }',
        '    return 0;',
        '}',
    ]:
        p = d.add_paragraph()
        # 用不换行空格会被当普通空格，这里直接用 ASCII 空格
        run = p.add_run(line)
        run.font.name = 'Consolas'
    d.add_paragraph('代码之后的正文段落。')
    d.save(path)
    return path


def build_image_doc(path, image_path):
    """夹具：内嵌图片，用于验证图片提取与引用。"""
    d = _new_doc()
    d.add_heading('图片夹具', 1)
    d.add_paragraph('下方应嵌入一张图片。')
    d.add_picture(image_path)
    d.save(path)
    return path


BUILDERS = {
    'merged': build_merged_table,
    'formula': build_omml_formula,
    'toc': build_toc,
    'list': build_numbered_list,
    'ruby': build_ruby,
    'code': build_code_block,
}


def build_all(out_dir):
    """生成全部文本类夹具，返回 {名称: 路径}。"""
    os.makedirs(out_dir, exist_ok=True)
    made = {}
    for name, fn in BUILDERS.items():
        p = os.path.join(out_dir, f'{name}.docx')
        made[name] = fn(p)
    return made
