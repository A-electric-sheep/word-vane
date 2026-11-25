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

# 版本号: 分发后用于区分行为差异, 修改转换逻辑时请同步递增
__version__ = '1.0.0'

import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

# 段落样式类型值(python-docx 1.x 为枚举 WD_STYLE_TYPE.PARAGRAPH)
_PARA_STYLE = WD_STYLE_TYPE.PARAGRAPH
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from math_latex import omml2latex  # OMML->LaTeX (单元级移植自 eqword2llm, MIT)

NS_VML = 'urn:schemas-microsoft-com:vml'

DEFAULT_HEADING_STYLES = {'Title': 1, 'Heading 1': 1, 'Heading 2': 2, 'Heading 3': 3,
                          'Heading 4': 4, 'Heading 5': 5, 'Heading 6': 6,
                          '标题 1': 1, '标题 2': 2, '标题 3': 3, '标题 4': 4, '标题 5': 5}

# 仓库只保留 Word 内置样式与**通用结构规则**，不含任何特定厂商/模板的专有样式名。
# 定制样式请通过外部配置注入（见 load_style_config），不要在源码里硬编码：
#   1) --styles 参数指定
#   2) 环境变量 WORD2MD_STYLES 指向的 JSON
#   3) ~/.word2md/styles.json
# 配置文件格式见 README「自定义样式映射」。
STYLE_CONFIG_ENV = 'WORD2MD_STYLES'
STYLE_CONFIG_HOME = os.path.join('~', '.word2md', 'styles.json')

# 通用「子标题样式」模式：样式名本身不代表章节层级，但正文具备小标题特征时，
# 可提升为 h3。默认只覆盖语言/行业中立的命名习惯，厂商专有前缀一律走外部配置。
DEFAULT_SUBHEADING_PATTERNS = [
    r'.*多栏$',          # 中文模板的多栏排版小标题
    r'.*multicol$',      # 英文对应写法
]

# 通用「章节样式」结构规则：Word 模板中自定义章节样式常写作
#   <前缀><数字>-<层级>      如 chapter1-2、sec2-3、part1-4
# 其中「-数字」后缀就是章节层级，形态足够明确，可直接采信。
# 这条规则替代了原先针对个别模板的硬编码样式表，对任意前缀同样生效。
RE_STRUCT_HEADING = re.compile(r'^[a-z]+(\d+)-(\d+)$', re.I)
# <前缀><数字>-<非数字名>（如 chapter1-content、part2-cover）：模板常把「目录/扉页」
# 这类前置页定义成一级章节，故统一按 h1 处理；正文特征校验通过才采信。
RE_STRUCT_HEADING_NAMED = re.compile(r'^[a-z]+\d+-[a-z]+$', re.I)
# <前缀><数字> 结尾且无连字符（如 section1 / part2）视为一级章节，
# 但该形态歧义较大，故排除常见的非章节样式词根，且仍需正文特征校验。
RE_BARE_HEADING = re.compile(
    r'^(?!list|bullet|table|figure|image|caption|note|code|normal|body'
    r'|enum|index|toc|quote|footer|header)[a-z]*\d+$', re.I)


def _level_from_struct_name(lower_name):
    """按通用结构规则从样式名推断章节层级(1-4)，不匹配返回 None。

    仅处理 `<前缀><数字>-<层级>` 这种明确形态；裸前缀+数字
    （section1 等）因与其他样式名歧义较大，交由调用方结合正文特征判定。
    """
    m = RE_STRUCT_HEADING.match(lower_name)
    if not m:
        return None
    # 取连字符后的数字为层级；模型里 N 通常从 1 起，个别模板用 0 表示首章
    lv = int(m.group(2))
    if lv == 0:
        lv = 1
    return lv if 1 <= lv <= 4 else None


def load_style_config(path=None):
    """加载外部样式配置（用于注定制模板，源码中不保留任何专有名称）。

    查找顺序:
        1) 显式传入的 path
        2) 环境变量 WORD2MD_STYLES 指向的 JSON
        3) ~/.word2md/styles.json

    文件格式（各键均可选）:
        {
          "heading_styles":    {"my-chapter-1": 1, "my-chapter-2": 2},
          "toc_styles":        {"mytoc 1": 1},
          "subheading_patterns": ["^mybody\\d*$"]
        }

    返回 dict，未找到任何配置时返回空 dict（调用方无需判空）。
    """
    import json

    candidates = []
    if path:
        candidates.append(os.path.expanduser(str(path)))
    env_path = os.environ.get(STYLE_CONFIG_ENV)
    if env_path:
        candidates.append(os.path.expanduser(env_path))
    candidates.append(os.path.expanduser(STYLE_CONFIG_HOME))

    for p in candidates:
        if not p or not os.path.isfile(p):
            continue
        try:
            with open(p, encoding='utf-8') as f:
                data = json.load(f)
        except (OSError, ValueError) as e:
            raise ValueError(f'样式配置读取失败: {p}: {e}') from e
        if not isinstance(data, dict):
            raise ValueError(f'样式配置必须是 JSON 对象: {p}')
        return data
    return {}
DEFAULT_TOC_STYLES = {'toc 1': 1, 'toc 2': 2, 'toc 3': 3, 'toc 4': 4, 'toc 5': 1, 'TOC1': 1,
                      'TOC2': 2, 'TOC3': 3, 'TOC4': 4,
                      # Word 内置目录样式的标准写法是「TOC 1」~「TOC 9」(带空格、
                      # 大写); 仅收录全小写/无空格形态会漏判, 致整块目录退化为正文
                      'TOC 1': 1, 'TOC 2': 2, 'TOC 3': 3, 'TOC 4': 4, 'TOC 5': 5,
                      'TOC 6': 6, 'TOC 7': 7, 'TOC 8': 8, 'TOC 9': 9}

CODE_PREFIXES = ('#include', '#define', 'unsigned', 'void ', 'static', 'typedef',
                 'int ', 'char ', 'long ', 'short ', 'float ', '/*', '//')
# 左值: 标识符 + 可选的 .成员/->成员 与 [下标] 后缀。
# 下标两侧允许空格 —— 规格书代码里 "REG_X [i]= *( POINT+i);" 这类写法
# 极常见(实测 10 份文档), 旧式 [\w$\[\]\.\->]* 要求 "[" 紧贴标识符, 会漏判。
RE_CODE_LVALUE = r"[A-Za-z_$][\w$\.\->]*(?:\s*\[\s*[^\]\s]*\s*\])*[\w$\.\->]*"
RE_CODE_ASSIGN = re.compile(
    r"^" + RE_CODE_LVALUE + r"\s*"
    r"(?:<<=|>>=|\|=|&=|\+=|-=|\*=|/=|%=|\^=|=(?!=))\s*"
    r"[0-9A-Za-z_x$'\s\.\*\+\-\/\%\|&\[\]\(\)~<>!]+"
    r"(?:\s*;\s*(?://.*)?)?$"
)
RE_CODE_CTRL = re.compile(r'^(?:\b(?:for|if|while|switch)\s*\(.*|\breturn\b.*|[\{\}])$')

# 函数调用语句: 标识符(可带 . -> 成员) + 括号实参 + 分号收尾。
# 必须要求 ";" —— 这是与规格书正文的关键区分线:
#   代码      "_nop_();"、"I2C_START();"、"I2C_WRITE(addr);"  → 命中
#   正文说明  "REG_A1 (CBH)"、"SOP16L(150mil)"、"REG_B (F0H)"
#             "PWMn周期寄存器高8位(读/写)"  → 全部不命中(无分号)
# 实测 73 份规格书: 正文中形如 "标识符(参数)" 的段落 16 种, 其中仅 _nop_() 系列
# 共 95 次是真代码; 其余为寄存器/封装说明, 去掉分号要求会把它们误包进代码块。
RE_CODE_CALL = re.compile(
    r"^[A-Za-z_]\w*(?:\s*(?:\.|->)\s*[A-Za-z_]\w*)*\s*"
    r"\([^()]*\)\s*;\s*(?://.*)?$"
)

# 自增/自减语句: "counter++;"、"++i;"、"a[0]--;"。
# 这类行既不是赋值(无 =)也不是函数调用(无括号), 旧版三者全不匹配,
# 会把含自增的循环体从中间劈开。
RE_CODE_INCDEC = re.compile(
    r"^(?:(?:" + RE_CODE_LVALUE + r"\s*(?:\+\+|--))"
    r"|(?:(?:\+\+|--)\s*" + RE_CODE_LVALUE + r"))\s*;\s*(?://.*)?$"
)

# 数学关系式排除: 形如 "duty1 = duty0"、"Baud = clock/(((SEG1+2)+(SEG2+1)))"。
# 特征: 等号右侧不含数字/十六进制/函数调用/分号, 且整体是纯标识符算式 ——
# 规格书用它表达"两量相等"的推导结论, 包进 ```c 会被误读成可执行赋值。
RE_MATH_RELATION = re.compile(r"^\s*[A-Za-z_]\w*(?:\s*[\*\+\-/]\s*[A-Za-z_]\w*)*\s*$")

# ---------------- 8051 汇编识别 ----------------
# 规格书里的「汇编例程」与「指令说明」形状相近, 必须区分:
#   汇编代码  "MOV R1,#83H\t; 将 REG_A1 的地址写入 R1"   → 入块
#   指令说明  "MOV  A, #50H    （这条指令是将立即数50H送到累加器A中）" → 不入块
#             "INC  R0      是指（R0）+1→R0"
# 判据优先级:
#   1. 含分号 → 代码(汇编注释一律以 ; 起头)
#   2. 含全角括号 / 大间距后半角括号 → 说明文字
#   3. 逐个校验操作数是否合法 8051 操作数形态 → 合法才算代码
ASM_MNEMONICS = frozenset({
    'NOP', 'MOV', 'MOVC', 'MOVX', 'ORL', 'ANL', 'XRL', 'INC', 'DEC', 'CLR', 'CPL',
    'SETB', 'ADD', 'ADDC', 'SUBB', 'MUL', 'DIV', 'JMP', 'SJMP', 'LJMP', 'AJMP',
    'CALL', 'LCALL', 'ACALL', 'DJNZ', 'CJNE', 'JZ', 'JNZ', 'JC', 'JNC', 'JB',
    'JNB', 'JBC', 'PUSH', 'POP', 'RET', 'RETI', 'XCH', 'XCHD', 'DA', 'SWAP',
    'RL', 'RR', 'RLC', 'RRC'})

# 合法 8051 操作数: 寄存器/立即数/直接地址/@间接/位地址/标号/偏移表达式
RE_ASM_OPERAND = re.compile(
    r'^(?:'
    r'[ARCDPB]+\d*'                                  # A, R0-R7, C, DPTR
    r'|@(?:R[01]|A\+DPTR|DPTR)'                      # @R0 @R1 @A+DPTR @DPTR
    r'|#[\w$]+'                                      # #50H #83H #0x01
    r'|[A-Za-z_$][\w$]*(?:\s*[+\-]\s*[\w$]+)*'       # 标号 / 偏移表达式
    r'|[\w$]+[Hh]?'                                  # 50H / 直接地址
    r'|/[\w]+'                                       # 位取反 /bit
    r')$')


def _looks_like_asm(t: str) -> bool:
    """判断是否为 8051 汇编代码行（区别于「指令说明」文字）。"""
    s = t.strip()
    if not s:
        return False
    # 规格书里偶用全角美元符号表示当前地址, 归一化后再解析
    s = s.replace('\uff04', '$')
    # 先剥离注释再取助记符: "NOP; 至少需要12个NOP指令" 的助记符是 NOP 而非 "NOP;"
    has_comment = ';' in s
    code_part = s.split(';', 1)[0].strip() if has_comment else s
    # 纯注释行("; xxx"): 剥完为空, 无从取助记符。规格书里这种行紧贴代码块,
    # 判为代码可让它留在块内; 但正文也可能有分号起头的说明, 保守起见判 False。
    if not code_part:
        return False
    parts = code_part.split(None, 1)
    head = parts[0]
    # 助记符必须全大写: 排除 "Push the PC value onto the Stack" 这类英文说明
    if head != head.upper() or head not in ASM_MNEMONICS:
        return False
    rest = parts[1].strip() if len(parts) > 1 else ''
    # 1. 汇编注释以 ';' 起头, 有分号即代码
    if has_comment:
        return True
    # 2. 无操作数的纯指令: 仅 NOP 无歧义; 其余助记符缺操作数多为说明文字
    if not rest:
        return head == 'NOP'
    # 3. 说明文字特征: 全角括号, 或操作数后大间距接半角括号
    if re.search(r'[（]', rest):
        return False
    if re.search(r'\s{2,}[\(\[]', rest):
        return False
    # 4. 逐个校验操作数合法性（最多两个）
    ops = [o.strip() for o in rest.split(',')]
    if len(ops) > 2:
        return False
    return all(o and RE_ASM_OPERAND.match(o) for o in ops)


def _looks_like_code(t: str) -> bool:
    if not t:
        return False
    # 先剥离前导空白再判定。C 代码在 Word 里常以缩进形式出现
    # (``    int n = 0;``), 若拿原串做 startswith / 正则 ^ 锚点,
    # 缩进行会全部漏判, 导致单个代码块被切成多段甚至整段不被识别。
    t = t.lstrip()
    if not t:
        return False
    for p in CODE_PREFIXES:
        if t.startswith(p):
            return True
    # 函数调用语句(以分号收尾): _nop_(); I2C_START(); I2C_WRITE(addr);
    # 这类行旧版完全漏判, 会把一个连续代码块从中间劈开(实测某份规格书
    # 的 _nop_() x12 与寄存器循环体被甩到围栏外)。
    if RE_CODE_CALL.match(t):
        return True
    # 自增/自减语句: counter++; ++i; a[0]--;
    if RE_CODE_INCDEC.match(t):
        return True
    # 8051 汇编: MOV R1,#83H / NOP / ORL PCON,#02H
    # 「汇编例程」在 29 份规格书中出现, 这些行若不入块会散成独立段落。
    if _looks_like_asm(t):
        return True
    # 纯标识符关系式(无分号): "duty1 = duty0" —— 规格书用它表达推导结论,
    # 不是可执行赋值, 必须在下面的分号判定前拦掉。
    if '=' in t and not re.search(r';\s*(?://.*)?$', t.rstrip()):
        rhs = t.split('=', 1)[1]
        if RE_MATH_RELATION.match(rhs):
            return False
    if not (RE_CODE_ASSIGN.match(t) or RE_CODE_CTRL.match(t)):
        return False
    # 排除"数学关系式"被误判为 C 赋值语句: 形如
    #   "Baud = clock/(((SEG1+2)+(SEG2+1))*(PRESC+1))"
    #   "duty1 = duty0 * 2"
    # 特征是右侧含算术运算且整行无分号结尾 —— C 赋值几乎总以 ';' 收尾,
    # 而规格书里的公式关系式不会。这类文本更应保留为正文(或公式),
    # 否则会被包进 ```c 代码块, 让 AI 误认为是可执行代码。
    body = t.rstrip()
    # 若以 ';' 或 '; //...' 收尾, 视为 C 语句(赋值/位运算), 保留代码块
    if re.search(r';\s*(?://.*)?$', body):
        return True
    # 无分号时: 右侧出现 2 个以上英文单词(如 "CPHA = 1 data transfer diagram")
    # 说明是说明性文字/图注, 不是表达式
    rhs = body.split('=', 1)[1] if '=' in body else ''
    if len(re.findall(r'[A-Za-z]{2,}', rhs)) >= 2:
        return False
    if re.search(r'[*/|&]', body):
        # 无分号且含算术/位运算符 -> 数学关系式, 不是可执行语句
        return False
    return True


def _slugify(text: str) -> str:
    t = re.sub(r'[^\w\u4e00-\u9fff\- ]', '', text, flags=re.UNICODE)
    return t.strip().lower().replace(' ', '-')


class Converter:
    """单个 docx 文档的转换器实例。"""

    def __init__(self, src_path, *, code_heuristics=True, include_toc=True,
                 heading_styles=None, toc_styles=None, code_lang='c',
                 image_dir_name='images', strip_emphasis=False,
                 style_config=None, subheading_patterns=None):
        self.src = str(src_path)
        self.code_heuristics = code_heuristics
        self.include_toc = include_toc
        self.code_lang = code_lang
        self.img_sub = image_dir_name
        # 剥离加粗/粗斜体(为 AI 阅读链路降噪); 不动文本内容与 sub/sup/s/br 功能标签
        self.strip_emphasis = strip_emphasis

        # 外部样式配置(定制模板)：源码不含专有样式名，一律从这里注入
        cfg = load_style_config(style_config)

        self.heading_level = dict(DEFAULT_HEADING_STYLES)
        self.heading_level.update(cfg.get('heading_styles') or {})
        if heading_styles:
            self.heading_level.update(heading_styles)
        # 样式名大小写不敏感索引: Word 文档中大小写混用极常见,
        # 且个别模板把首页设为分页空段落、次样式才是一级标题,
        # 故统一用小写键查找, 空标题段落由调用处跳过。
        self._heading_ci = {k.strip().lower(): v for k, v in self.heading_level.items()}

        self.toc_level = dict(DEFAULT_TOC_STYLES)
        self.toc_level.update(cfg.get('toc_styles') or {})
        if toc_styles:
            self.toc_level.update(toc_styles)
        # 与标题映射保持一致: 样式名大小写不敏感查找。
        # Word 中「TOC 1」/「toc 1」/「TOC1」三种写法并存, 精确匹配会漏判。
        self._toc_ci = {k.strip().lower(): v for k, v in self.toc_level.items()}

        # 子标题样式模式：通用默认 + 外部配置追加
        pats = list(DEFAULT_SUBHEADING_PATTERNS)
        pats += [p for p in (cfg.get('subheading_patterns') or []) if p]
        if subheading_patterns:
            pats += list(subheading_patterns)
        self._subheading_re = (re.compile('|'.join(f'(?:{p})' for p in pats), re.I)
                               if pats else None)

        self.warnings: list[str] = []
        self.doc = docx.Document(self.src)
        self.z = zipfile.ZipFile(self.src)

        # 编号系统(键务必为 int — 历史 bug: str 键导致编号列表全部退化为 bullet)
        self.num2abs, self.ABSTRACT = self._load_numbering()

        # 样式查表缓存(需在预扫描前就绪: _heading_level_of 会用到)
        self._style_name_cache = {}
        self._style_obj_cache = {}
        try:
            self._style_default_name = self.doc.styles.default(_PARA_STYLE).name or ''
        except Exception:
            self._style_default_name = 'Normal'

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
        # 目录域自身的标题文字(如 "目录" / "Content"), 供无正文目录标题时复用
        self.toc_title = ''
        self._cjk_cache = None
        # 标题判定上下文: 上一块是否为图片(用于排除图片说明文字被当成标题)
        self._prev_block_had_image = False

    # ---------- 辅助 ----------
    def _is_cjk_doc(self):
        """按正文 CJK 字符占比判断文档语种(用于目录标题/图注文案)"""
        if self._cjk_cache is not None:
            return self._cjk_cache
        sample = ''
        for p in self.doc.paragraphs[:200]:
            sample += p.text or ''
        if not sample:
            self._cjk_cache = True
            return True
        cjk = sum(1 for ch in sample if '\u4e00' <= ch <= '\u9fff')
        self._cjk_cache = cjk * 20 >= len(sample)
        return self._cjk_cache

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

    def _num_fmt_of(self, p_el, para=None):
        """返回 (numId, ilvl) 或 None。

        Word 的列表编号既可以写在段落 pPr/numPr 上, 也可以由**样式**携带
        (pStyle -> style/pPr/numPr)。中文规格书大量使用后者, 只读段落会漏掉
        绝大部分列表(实测单批 3493 段样式级编号中 97% 被压平为普通段落)。
        因此需沿样式链继承。
        """
        pPr = p_el.find(qn('w:pPr'))
        numPr = pPr.find(qn('w:numPr')) if pPr is not None else None

        # 段落自身优先; 若仅指定 ilvl 未指定 numId, 仍需回落到样式取 numId
        if numPr is not None:
            nid = numPr.find(qn('w:numId'))
            ilvl = numPr.find(qn('w:ilvl'))
            nv = nid.get(qn('w:val')) if nid is not None else None
            iv = ilvl.get(qn('w:val')) if ilvl is not None else None
            if nv is not None and nv != '0':
                return (int(nv), int(iv) if iv is not None else 0)

        # 沿样式链继承(按 styleId 走缓存, 避免 python-docx 全表查找)
        style = self._style_obj_of(p_el) if para is not None else None
        cur = style
        hops = 0
        while cur is not None and hops < 6:
            hops += 1
            el = getattr(cur, 'element', None)
            if el is not None:
                spPr = el.find(qn('w:pPr'))
                sNum = spPr.find(qn('w:numPr')) if spPr is not None else None
                if sNum is not None:
                    nid = sNum.find(qn('w:numId'))
                    nv = nid.get(qn('w:val')) if nid is not None else None
                    if nv is not None and nv != '0':
                        ilvl = sNum.find(qn('w:ilvl'))
                        iv = ilvl.get(qn('w:val')) if ilvl is not None else None
                        # 段落若显式给了 ilvl, 以段落为准
                        if numPr is not None:
                            pi = numPr.find(qn('w:ilvl'))
                            if pi is not None and pi.get(qn('w:val')) is not None:
                                iv = pi.get(qn('w:val'))
                        return (int(nv), int(iv) if iv is not None else 0)
            cur = getattr(cur, 'base_style', None)
        return None

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

    def _img_caption(self, seq):
        """图片 alt 文案按文档语种生成, 避免英文文档出现中文图注"""
        return f'图 {seq}' if self._is_cjk_doc() else f'Figure {seq}'

    def _looks_like_subheading(self, text):
        """判断文本是否像"小标题"而非正文/条目/图注。

        用于区分子标题样式(小标题)与正文样式(正文段或图注)、条目样式。
        小标题特征: 短、不以标点收尾、无句号式陈述。
        """
        t = (text or '').strip()
        if not t or len(t) > 20:
            return False
        # 句末标点 = 陈述句, 不是标题
        if t[-1] in '。；;，,：:！!？?、.':
            return False
        # 句中含句读 = 正文长句被折断, 不是标题
        if re.search(r'[。；;，,]', t):
            return False
        if re.match(r'^[A-Za-z0-9_\-\.]+\s*[:：]', t):
            return False
        # 图注/表注("图 3 xxx"/"表 5-1 xxx"/"Figure 3 xxx")不是章节标题
        if re.match(r'^(图|表|Figure|Table)\s*[\d一二三四五六七八九十]', t, re.I):
            return False
        # 含数值+单位的条目(如 "1.8V~5.5V")不是标题
        if re.search(r'\d\s*(?:V|A|mA|Hz|MHz|KHz|℃|%|ms|us|ns|bit|byte)', t, re.I):
            return False
        return True

    def _style_obj_of(self, p_el):
        """按 pStyle styleId 取样式对象(带缓存), 供编号/标题判定使用"""
        pPr = p_el.find(qn('w:pPr'))
        sid = None
        if pPr is not None:
            ps = pPr.find(qn('w:pStyle'))
            if ps is not None:
                sid = ps.get(qn('w:val'))
        if sid is None:
            return None
        if sid in self._style_obj_cache:
            return self._style_obj_cache[sid]
        obj = None
        try:
            obj = self.doc.styles.get_by_id(sid, _PARA_STYLE)
        except Exception:
            obj = None
        self._style_obj_cache[sid] = obj
        return obj

    def _style_name_of(self, p_el):
        """按段落 pStyle 的 styleId 取样式名(带缓存)。

        python-docx 的 Paragraph.style 属性每次都调用 styles.get_by_id ->
        default_for 遍历整张样式表, 大文档上会退化成 O(n*m)。这里直接读
        pPr/pStyle 的 val(styleId), 用字典缓存 styleId -> 样式名。
        """
        pPr = p_el.find(qn('w:pPr'))
        sid = None
        if pPr is not None:
            ps = pPr.find(qn('w:pStyle'))
            if ps is not None:
                sid = ps.get(qn('w:val'))
        if sid is None:
            return self._style_default_name
        cached = self._style_name_cache.get(sid)
        if cached is not None:
            return cached
        name = ''
        try:
            st = self.doc.styles.get_by_id(sid, _PARA_STYLE)
            name = st.name or ''
        except Exception:
            name = ''
        self._style_name_cache[sid] = name
        return name

    def _heading_level_of(self, para_or_el):
        """返回段落的大纲级别(1-6), 非标题返回 None。

        判定优先级:
          1) 样式名命中映射表(大小写不敏感) —— 高置信, 直接采用
          2) 段落/样式自身的 w:outlineLvl —— 兜底, 但需通过合理性校验
        注意:
          - **不使用 basedOn 继承链**。实测某些正文样式 basedOn 为
            "Heading 5", 但语义上只是三级小标题, 按继承链会误判大量段落为 h5。
          - outlineLvl 兜底限 1~4 级且排除"寄存器名/位域名"式文本。Word 里
            寄存器小节标题常被设 outlineLvl=5, 强行提升会产出 100+ 个 h5/h6
            标题并造成 30+ 处层级跳级, 反而破坏 AI 的章节导航。
        """
        p_el = getattr(para_or_el, '_p', para_or_el)
        # 性能: python-docx 的 paragraph.style 每次都会遍历样式表(实测单文档
        # 3378 次调用耗 41s), 这里改为按 pStyle 的 styleId 走缓存查表
        sname = self._style_name_of(p_el)

        lv = self._heading_ci.get(sname.strip().lower())
        if lv:
            return lv

        text = ''.join(t.text or '' for t in p_el.findall('.//' + qn('w:t'))).strip()
        if not text:
            return None

        def plausible(level):
            """outlineLvl / 样式名的合理性校验"""
            # 兜底最多产出 h4: 更深层级的 outlineLvl 通常对应"寄存器小节/模式
            # 列举"等表项式文本, 提升为 h5/h6 会与现有 h2/h3 骨架错位, 产生
            # 大量层级跳级, 反而降低 AI 的章节导航质量
            if level > 4:
                return None
            # 寄存器/位域条目形如 "REG_A1 (83H@FFH) System Clock ..." /
            # "REG_B2 (F7H) ..." —— 语义是表项而非章节, 不提升为标题
            if re.match(r'^[A-Z][A-Z0-9_]{1,15}\s*[（(]\s*[0-9A-Fa-f]{1,4}[Hh]', text):
                return None
            if len(text) > 90:
                return None
            return level

        # 样式名启发式(仅针对 outlineLvl 兜底之前): 此类文档里
        #   子标题样式 -> 小标题("工作条件"、"封装类型")
        #   条目样式   -> 内容条目(非标题)
        # 二者仅靠样式名难以区分, 用"是否像标题"的文本特征判定。
        lower = sname.strip().lower()
        # 1) 通用结构规则: <前缀><数字>-<层级>, 取后缀数字为层级(形态明确, 直接采信)
        lv_struct = _level_from_struct_name(lower)
        if lv_struct:
            return lv_struct
        # 2) <前缀><数字>-<名字>(如 chapter1-content): 前置页章节, 按 h1, 需正文像标题
        if RE_STRUCT_HEADING_NAMED.match(lower) and len(text) <= 90 \
                and not getattr(self, '_prev_block_had_image', False) \
                and self._looks_like_subheading(text):
            return 1
        # 3) <前缀><数字> 形态(如 section1): 歧义较大, 要求正文也像标题
        if RE_BARE_HEADING.match(lower) and len(text) <= 90 \
                and not getattr(self, '_prev_block_had_image', False) \
                and self._looks_like_subheading(text):
            return 1
        # 4) 子标题模式(通用默认 + 外部配置): 仅当正文像小标题时提升为 h3
        if (self._subheading_re is not None and self._subheading_re.match(lower)
                and not getattr(self, '_prev_block_had_image', False)
                and self._looks_like_subheading(text)):
            # 依附于该文档已出现的最深 h1/h2 之下, 取 h3
            return 3

        # 段落直接的 outlineLvl
        pPr = p_el.find(qn('w:pPr'))
        if pPr is not None:
            ol = pPr.find(qn('w:outlineLvl'))
            if ol is not None:
                try:
                    v = int(ol.get(qn('w:val')))
                except (TypeError, ValueError):
                    v = None
                if v is not None and 0 <= v <= 8:
                    hit = plausible(min(v + 1, 4))
                    if hit:
                        return hit

        # 样式自身定义的 outlineLvl(不用 basedOn 链, 见 docstring)
        style = self._style_obj_of(p_el)
        if style is not None:
            el = getattr(style, 'element', None)
            if el is not None:
                ol = el.find(qn('w:pPr') + '/' + qn('w:outlineLvl'))
                if ol is not None:
                    try:
                        v = int(ol.get(qn('w:val')))
                    except (TypeError, ValueError):
                        v = None
                    if v is not None and 0 <= v <= 8:
                        hit = plausible(min(v + 1, 4))
                        if hit:
                            return hit
        return None

    def _prescan_headings(self):
        used = set()
        prev_img = False
        for blk in self._blocks():
            if not isinstance(blk, Paragraph):
                prev_img = False
                continue
            # 保持与渲染阶段一致的上下文(上一块是否为图), 否则两阶段的
            # 标题集合会不一致, 导致目录锚点与正文标题对不上
            self._prev_block_had_image = prev_img
            prev_img = bool(self._para_imagedata(blk._p))
            if self._heading_level_of(blk) is None:
                continue
            text = ''.join(t.text or '' for t in blk._p.findall('.//' + qn('w:t'))).strip()
            if not text:
                continue
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

    @staticmethod
    def _run_inline_text(r_el):
        """按 XML 顺序重建 run 的内联文本, 保留 w:tab / w:br / w:noBreakHyphen。

        只收集 w:t 会整体丢弃 run 内的制表符与软回车(缺陷 D1):
          - C 代码 "int\\ta = 1;" -> "inta = 1;"(标识符被篡改)
          - 汇编 "MOV A,#50H\\t; 注释" -> "MOV A,#50H; 注释"(注释符粘连)
          - 正文 "字段A\\t字段B" -> 列分隔丢失, 文本粘连
        这三类都是静默错误: 内容看起来完整, 但语义已变。
        """
        out = []
        for ch in r_el:
            tag = ch.tag
            if tag == qn('w:t'):
                out.append(ch.text or '')
            elif tag == qn('w:tab'):
                out.append('\t')
            elif tag == qn('w:cr'):
                out.append('\n')
            elif tag == qn('w:br'):
                # 仅文本环绕型换行产生软回车; 分页/分栏符不产生可见字符
                if ch.get(qn('w:type')) in (None, 'textWrapping'):
                    out.append('\n')
            elif tag == qn('w:noBreakHyphen'):
                out.append('-')
        return ''.join(out)

    @staticmethod
    def _to_prose_inline(s):
        """把内联文本转为正文形态。

        制表符 -> 单空格: 保留列间分隔, 同时避免行首 tab 被 Markdown
        解析成缩进代码块(4 空格/制表符起始即代码块), 破坏段落边界。
        软回车 -> <br>: 保留 Word 内的手工换行。

        首尾的软回车要丢弃: 段落仅由一个 w:br 构成的空段落(排版留白)
        若转成 <br> 会产出孤立空行, 是纯噪声; 且这类段落本应走空段落分支。
        """
        s = s.replace('\t', ' ')
        s = s.strip(' \r\n\t')
        # 去掉首尾残留的 <br>(可能由连续的 w:br 产生)
        while True:
            t = re.sub(r'^(?:<br>\s*)+', '', s)
            t = re.sub(r'(?:\s*<br>)+$', '', t)
            if t == s:
                break
            s = t
        return s.replace('\n', '<br>')

    def rich_runs(self, el, plain=False):
        """渲染段落内联内容。

        内部先收集 [(kind, payload, bold, italic, script)] 片段, 末尾把**相邻且
        格式完全相同**的文本片段合并后再统一包裹强调语法。Word 把一个视觉连续的
        加粗短语拆成多个 run 时(中文规格书极常见), 逐 run 包裹会产出大量
        `****` 接缝(实测单文档可达 1500+ 处), 严重干扰 AI 阅读与检索。

        plain=True 时输出纯文本: 不加 ** 强调、不生成 <sub>/<sup> 标签,
        供代码块使用。代码块内的 `*` 是 C 指针/解引用符, 用正则事后剔除
        会把 "unsigned char code * POINT" 误伤成 "unsigned char code  POINT"
        (语义破坏), 故必须在片段层面就抑制标记生成, 而非事后清洗。
        """
        segs = []   # ('text', str, b, i, v) 或 ('raw', str)

        def add_text(txt, b, i, v):
            if not txt:
                return
            # 上下标先就地内联成标签, 使合并键只含 (bold, italic) ——
            # 否则 "<sub>SYS</sub>" 与前后同粗体文本因 vertAlign 不同而无法合并,
            # 残留 "**f****<sub>SYS</sub>****）：**" 这类接缝
            if not plain:
                if v == 'subscript':
                    txt = '<sub>' + txt + '</sub>'
                elif v == 'superscript':
                    txt = '<sup>' + txt + '</sup>'
            if segs and segs[-1][0] == 'text':
                _, prev, pb, pi_ = segs[-1]
                if (pb, pi_) == (b, i):
                    segs[-1] = ('text', prev + txt, b, i)
                    return
            segs.append(('text', txt, b, i))

        def add_raw(s):
            if s:
                segs.append(('raw', s))

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
                            add_text(txt, b, i, v)
                        continue
                    txt = self._run_inline_text(child)
                    add_text(txt, b, i, v)
                elif tag == qn('w:hyperlink'):
                    anchor = child.get(qn('w:anchor'))
                    # 超链接内同样要把"相邻同格式"的 run 合并后再包裹强调,
                    # 否则会产出 "[**2****7.10ADC****电气特性**](#...)" 这类
                    # 碎片(实测多份 Datasheet 的交叉引用处)
                    inner_segs = []

                    def add_inner(txt, b, i, v):
                        if not txt:
                            return
                        if not plain:
                            if v == 'subscript':
                                txt = '<sub>' + txt + '</sub>'
                            elif v == 'superscript':
                                txt = '<sup>' + txt + '</sup>'
                        if inner_segs and inner_segs[-1][0] == 'text':
                            _, prev, pb, pi_ = inner_segs[-1]
                            if (pb, pi_) == (b, i):
                                inner_segs[-1] = ('text', prev + txt, b, i)
                                return
                        inner_segs.append(('text', txt, b, i))

                    for r in child.findall(qn('w:r')):
                        rPr = r.find(qn('w:rPr'))
                        hidden = rPr is not None and rPr.find(qn('w:webHidden')) is not None
                        if hidden:
                            continue
                        b, i, v = self._run_fmt(r)
                        t = self._run_inline_text(r)
                        add_inner(t, b, i, v)
                    # plain 模式只取纯文本标签, 不生成链接语法(代码块内不适用)
                    if plain:
                        add_raw(''.join(s[1] for s in inner_segs if s[0] == 'text'))
                        continue
                    label = ''.join(self._wrap(s[1], s[2], s[3], None)
                                    for s in inner_segs if s[0] == 'text')
                    if anchor and anchor in self.anchor2slug:
                        add_raw(f'[{label}](#{self.anchor2slug[anchor]})')
                    else:
                        add_raw(label)
                elif tag == qn('w:tab'):
                    add_raw('\t')
                elif tag == qn('w:br'):
                    add_raw('\n')
                elif tag == qn('w:pict'):
                    for im in child.findall('.//{' + NS_VML + '}imagedata'):
                        rid = im.get(qn('r:id'))
                        if rid:
                            add_raw('[[IMG:' + rid + ']]')
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
                            add_text(t, b, i, v)
                elif tag == qn('w:instrText') or tag == qn('w:fldChar'):
                    continue
                elif tag == qn('m:oMathPara'):
                    add_raw('$$' + omml2latex(child) + '$$')
                    # plain 通道仅用于取代码块文本, 不计入公式统计, 否则
                    # 同一段落被渲染两次会导致统计翻倍
                    if not plain:
                        self.formula_blocks += 1
                elif tag == qn('m:oMath'):
                    add_raw('$' + omml2latex(child) + '$')
                    if not plain:
                        self.formula_blocks += 1
                else:
                    walk(child)

        walk(el)
        out = []
        for seg in segs:
            if seg[0] == 'raw':
                out.append(seg[1])
            else:
                _, txt, b, i = seg
                # plain(代码块) 模式: 不包 ** 强调, 保留文本里字面的 * 与 _
                if plain:
                    out.append(txt)
                else:
                    out.append(self._wrap(txt, b, i, None))
        if plain:
            return ''.join(out)
        return self._seal_emphasis(''.join(out))

    def _seal_emphasis(self, s):
        """消除强调接缝 `****`。

        合并同类相邻 run 之后仍有几种残留接缝:
          - 超链接元素夹在加粗文本中间: "**...的**[**ADC**](#x)**。**"
            会输出 "...的**[" 形式, 形成 "...**[" 与 "]**。**" 的相邻 `****`
          - 表格 HTML 单元格内跨 <p> 的加粗
        这里做纯文本层面的安全收口: 把 "****" 折叠为 "**", 并清理
        "**[" / "]**" 边界处的多余标记。仅在 strip_emphasis 关闭时执行。
        """
        if self.strip_emphasis or '****' not in s:
            return s
        # 接缝的本质: "…**" 紧跟 "**…" 时, 闭合与开启彼此抵消, 只剩一个开区间。
        # 反复消除 "**" + "**" 的边界, 直到稳定; 再把残留 >=3 的星号簇压成 2。
        prev = None
        while prev != s:
            prev = s
            s = s.replace('****', '')
        s = re.sub(r'\*{3,}', '**', s)
        return s

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

        # 兜底 3: 页码直接粘在标题尾部(无 w:tab 分隔), 如
        #   "TWI0/1 register mapping122" -> "TWI0/1 register mapping"
        # 仅当尾部数字 ≥2 位且剥离后标题仍非空时处理, 避免误删
        # "Operating Mode 0" / "Timer0" 这类本身含数字的标题。
        if label and label.strip():
            s = label.strip()
            # 尾部数字前若是 "/" "~" "-" "." 等连接符, 说明它属于标题本身
            # (如 "SPI0/1"、"INT0~15"、"TIMER0-3"), 不是页码, 不可剥离
            m_tail = re.match(r'^(.*?[^\d\s])\s*(\d{2,4})$', s)
            if m_tail:
                head = m_tail.group(1)
                if len(head) >= 4 and not re.search(r'[/\-–—\.\~～]$', head):
                    # 再排除"数字范围"整体: 形如 "INT0~15"/"P0.0-P0.7",
                    # 剥离后会留下悬空连接符, 视为标题自带内容
                    if not re.search(r'\d\s*[\~\-\u2013\u2014]\s*$', head):
                        label = head.strip()
        if not label or not label.strip():
            return None
        # 过滤目录自引用条目: 目录域里常把"目录"/"Content"自身也列进去,
        # 页码可能是 1~2 位(如 "目录9"、"Content2"), 故先剥尾部页码再判等
        bare = re.sub(r'\s*\d{1,4}$', '', label.strip()).strip()
        if bare in ('目录', '目录（Contents）', 'Content', 'Contents',
                    'Table of Contents', '目 录'):
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
        txt = self._run_inline_text(r_el)
        if not txt:
            return ''
        # run 内可能含制表符/软回车: 表格单元格内制表符作空格(HTML 会折叠
        # 连续空白), 软回车转 <br> 以保留手工换行。
        # 转义必须在插入 <br> 之前完成, 否则自家标签会被转义成实体。
        txt = txt.replace('\t', ' ')
        txt = txt.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        txt = txt.replace('\n', '<br>')
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
        """取段落级默认文字格式。

        **只继承 vertAlign(上下标), 不继承 b/i/strike。**

        原因: OOXML 里 pPr/rPr 描述的是「段落标记」(pilcrow ¶)自身的格式,
        不是段落内文字的格式。Word 里改「段落标记」字体时只影响那个看不见的
        回车符, 不影响文字显示。

        实测某份规格书: 表43 的 REG_C 单元格 pPr/rPr 带 <w:strike/>, 但
        run 自身无 strike, 同表的其它单元格也无此标记 —— 三者在 Word 中
        显示完全一致(都没删除线)。旧版把它当文字格式继承, 于是输出
        <s>REG_C</s>, 凭空多出删除线。同类误伤: 全库 37 处 <strong> + 14 处
        <s> 是段落标记污染, 与实际可见格式不符。

        vertAlign 保留继承: 部分文档把整段设为上下标时只写在 pPr/rPr 里,
        且该属性针对文字本身, 无 pilcrow 语义歧义。
        """
        pPr = p_el.find(qn('w:pPr'))
        if pPr is None:
            return {}
        mPr = pPr.find(qn('w:rPr'))
        if mPr is None:
            return {}
        inherit = {}
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
        # 多段落以换行连接(输出管线将 \n 转 <br>), 避免多段说明(QTWCK 分频逐项)
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
        """收集段落内所有图片引用 rId(保序、去重)。

        必须同时覆盖两种图片容器, 否则会整段丢图:
          - VML:       <w:pict>/<v:imagedata r:id="...">       (老式/兼容模式)
          - DrawingML: <w:drawing>/<a:blip r:embed="...">      (Word 2007+ 默认)
        历史缺陷: 早期实现只扫 VML, 导致全部 DrawingML 图片丢失, 而"产品命名规则"
        等章节的唯一内容就是一张图, 整节因此变空。
        """
        rids = []
        seen = set()

        def push(rid):
            if rid and rid not in seen:
                seen.add(rid)
                rids.append(rid)

        # VML
        for im in p_el.findall('.//{' + NS_VML + '}imagedata'):
            push(im.get(qn('r:id')))
        # DrawingML (a:blip 用 r:embed; 少数用 r:link)
        for blip in p_el.findall('.//' + qn('a:blip')):
            push(blip.get(qn('r:embed')) or blip.get(qn('r:link')))
        return rids

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
        in_list = False   # 是否处于列表块中(用于列表结束后补空行)
        prev_heading_lv = 0   # 上一个已输出标题的层级(用于压平跳级)

        def break_list():
            """列表块结束: 补空行, 避免后续内容被 GFM 并入最后一个列表项"""
            nonlocal in_list
            if in_list:
                if lines and lines[-1] != '':
                    lines.append('')
                in_list = False

        def flush_code():
            nonlocal code_buf
            if len(code_buf) >= 1:
                # 语言标记按块内容判定: 块内过半行为汇编指令则标 asm,
                # 否则用用户指定值(默认 c)。C 代码块里混一两行汇编注释时
                # 仍应标 c, 故用「过半」而非「任一」。
                lang = self.code_lang
                body = [c for c in code_buf if c.strip()]
                if body:
                    n_asm = sum(1 for c in body if _looks_like_asm(c))
                    if n_asm * 2 >= len(body):
                        lang = 'asm'
                lines.append('```' + lang)
                for c in code_buf:
                    lines.append(c)
                lines.append('```')
                lines.append('')
                self.code_blocks += 1
            code_buf = []

        prev_had_image = False
        for blk in self._blocks():
            if isinstance(blk, Paragraph):
                p_el = blk._p
                style = self._style_name_of(p_el)   # 走缓存, 避免全表遍历
                pure = ''.join(t.text or '' for t in p_el.findall('.//' + qn('w:t'))).strip()
                # 内联原文(保留 \t 与 \n)与正文形态(制表符->空格, 软回车-><br>)
                # 必须分开: 代码块内制表符是有意义的缩进/对齐, 转空格会破坏
                # 对齐; 而正文里的裸制表符行首会被 Markdown 当作缩进代码块。
                txt_raw = self.rich_runs(p_el)
                # 代码块走纯文本模式: 不生成 ** 强调与 <sub>/<sup> 标签。
                # 不可用正则事后剔除星号 —— C 的指针/解引用符 "*" 会被误删,
                # "unsigned char code * POINT" 会变成 "unsigned char code  POINT"
                code_txt = self.rich_runs(p_el, plain=True).strip('\r\n')
                txt = self._to_prose_inline(txt_raw)
                imgs = self._para_imagedata(p_el)
                # 上一块是图片、本块是短文本 -> 这是图片说明文字(图注),
                # 不应当作章节标题; 记入供标题判定排除
                self._prev_block_had_image = prev_had_image
                prev_had_image = bool(imgs)

                if imgs:
                    break_list()
                    flush_code()
                    if txt and '[[IMG:' not in txt:
                        lines.append(txt)
                        lines.append('')
                    for rId in imgs:
                        self.img_seq += 1
                        lines.append(f'![{self._img_caption(self.img_seq)}]({self._img_rel_path(rId)})')
                        lines.append('')
                    continue

                if not pure:
                    if code_buf:
                        pending_blank = True
                        continue
                    if txt:
                        break_list()
                        flush_code()
                        lines.append(txt)
                        lines.append('')
                        continue
                    flush_code()
                    lines.append('')
                    continue

                _toc_lv = self._toc_ci.get(style.strip().lower())
                if _toc_lv and self.include_toc:
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
                        toc_entries.append((_toc_lv, label, target))
                    continue

                hlv = self._heading_level_of(blk)
                if hlv and txt:
                    break_list()
                    flush_code()
                    # 层级连续性约束: 若当前级别比上一标题深超过 1 级(如 h1 后
                    # 直接出现 Heading 5 的"工作条件"), 压到 prev+1。
                    # 源文档常把正文小标题设成 Heading 5, 直接照搬会产生
                    # h1->h5 的断裂层级, 破坏 AI 的章节树导航。
                    lv = hlv
                    if prev_heading_lv and lv > prev_heading_lv + 1:
                        lv = prev_heading_lv + 1
                    prev_heading_lv = lv
                    # 标题本身已是结构强调, 内部再保留 ** 只会产生
                    # "# **主要功能**" 这类噪声(实测 7 份文档 26 处), 故剥离
                    head_txt = re.sub(r'\*{1,3}', '', txt).strip()
                    lines.append('#' * lv + ' ' + head_txt)
                    lines.append('')
                    self.heading_count += 1
                    txt = head_txt
                    if not self.toc_title and pure:
                        # 记录正文中的目录标题(中文"目录"/英文"Content"等),
                        # 供目录域无原位标题时复用, 避免凭空造标题
                        if re.fullmatch(r'\s*(目录|目\s*录|content|contents|table of contents)\s*',
                                        pure, re.I):
                            self.toc_title = pure
                    # 目录标题命中时原位放占位符: 中英文版本都能把目录块渲染在
                    # 原文目录标题之下, 而不是一律塞到文首再造一个标题
                    if self.include_toc and re.fullmatch(
                            r'\s*(目录|目\s*录|content|contents|table of contents)\s*', pure, re.I):
                        lines.append('<<TOC_PLACEHOLDER>>')
                        lines.append('')
                    continue

                nf = self._num_fmt_of(p_el, blk)
                if nf and re.match(r'^\s*(?:\d+\.|[a-zA-Z]\.|[（(]\d+[）)]|[-–—•·▪])\s', txt):
                    # 正文自带编号前缀(如 "2. Set ADCIS...")且样式又挂了 numPr 时,
                    # 再套一层自动编号会变成 "56. 2. Set ..." 双编号, 故不生成标记
                    nf = None
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
                    in_list = True
                    continue

                if self.code_heuristics and _looks_like_code(pure) and 'IMG:' not in txt:
                    if pending_blank:
                        code_buf.append('')
                        pending_blank = False
                    # 代码块内保留原始制表符: 它是缩进与注释对齐的语义载体,
                    # 转空格会改变列位置(缺陷 D1 的代码侧表现)
                    code_buf.append(code_txt)
                    continue
                else:
                    pending_blank = False
                    flush_code()

                # 列表块结束时补空行: 否则紧随其后的正文会被 GFM 解析器并入
                # 最后一个列表项(实测 21 处), 破坏段落边界与 AI 的分段理解
                break_list()
                lines.append(txt)
                lines.append('')
            else:
                break_list()
                flush_code()
                lines.extend(self._render_table(blk._tbl))
                lines.append('')

        flush_code()

        body = '\n'.join(lines)

        # 目录块原位注入
        toc_note = ('> 由 Word「目录」域转换的静态目录，点击条目可跳转到对应章节。'
                    if self._is_cjk_doc() else
                    '> Static table of contents converted from the Word TOC field; '
                    'each entry links to its section.')
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
            elif not toc_entries:
                # 无目录域: 完全跳过目录输出
                pass
            elif self.toc_title:
                # 原文正文里有目录标题(中文"目录"/英文"Content"), 但目录域未紧邻
                # 标题: 用原文标题兜底, 保证标题文字来自原文档而非杜撰
                body = f'## {self.toc_title}\n\n' + toc_body + '\n\n' + body
            else:
                # 原文没有目录标题(仅裸目录域, 多见于英文 Datasheet): 直接输出
                # 目录清单, 不杜撰标题 —— 否则会多出一个原文不存在的块,
                # 破坏与 Word 的逐段一致性
                body = toc_body + '\n\n' + body
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
        style_config:     外部样式配置 JSON 路径(定制模板用, 见 load_style_config)
        subheading_patterns: 追加的子标题样式名正则列表
        code_lang:        代码块语言标记(默认 'c')
    返回: 统计 dict(含 md_path / warnings / 各类计数)"""
    src_path = str(src_path)
    if not os.path.isfile(src_path):
        raise FileNotFoundError(
            f'找不到文件: {src_path}\n'
            f'请检查路径是否正确(当前目录: {os.getcwd()})')
    # 旧版 .doc 是 OLE 二进制, python-docx 无法读取。必须在这里拦下并给出
    # 可操作指引; 否则会直接抛底层 PackageNotFoundError, 用户无从判断原因。
    # 注意: 不要写成 endswith(('.docx', '.doc')) 再在内层判断 —— 那样外层条件
    # 对 .doc 恒为 False, 内层提示永不可达, 会退化为底层异常。
    if src_path.lower().endswith('.doc'):
        raise ValueError(
            '旧版 .doc 二进制格式不被本工具直接支持。请先用 Word/WPS 另存为 .docx, '
            '或使用 LibreOffice: soffice --headless --convert-to docx 输入.doc')
    if not src_path.lower().endswith('.docx'):
        raise ValueError(f'不支持的输入格式: {src_path}, 仅支持 .docx')

    conv = Converter(src_path,
                     code_heuristics=kwargs.get('code_heuristics', True),
                     include_toc=kwargs.get('include_toc', True),
                     heading_styles=kwargs.get('heading_styles'),
                     toc_styles=kwargs.get('toc_styles'),
                     code_lang=kwargs.get('code_lang', 'c'),
                     image_dir_name=kwargs.get('image_dir_name', 'images'),
                     strip_emphasis=kwargs.get('strip_emphasis', False),
                     style_config=kwargs.get('style_config'),
                     subheading_patterns=kwargs.get('subheading_patterns'))
    if out_dir is None:
        out_dir = os.path.join(os.path.dirname(src_path),
                               os.path.splitext(os.path.basename(src_path))[0] + '_md')
    return conv.convert(str(out_dir), extract_images=kwargs.get('extract_images', True))


def main(argv=None):
    ap = argparse.ArgumentParser(description='通用 Word(.docx) -> Markdown 转换器')
    ap.add_argument('--version', action='version',
                    version=f'word2md {__version__}')
    ap.add_argument('input', help='输入 .docx 文件')
    ap.add_argument('-o', '--output-dir', default=None, help='输出目录(缺省: 输入旁 <名>_md)')
    ap.add_argument('--no-images', action='store_true', help='不提取图片')
    ap.add_argument('--no-code-heuristics', action='store_true', help='关闭代码块启发式')
    ap.add_argument('--no-toc', action='store_true', help='关闭目录域处理')
    ap.add_argument('--no-bold', action='store_true', help='剥离加粗/粗斜体强调语法(为 AI 阅读降噪, 文本内容与 sub/sup/br 功能标签保留)')
    ap.add_argument('--code-lang', default='c', help='代码块语言标记(默认 c)')
    ap.add_argument('--styles', default=None, metavar='JSON',
                    help='外部样式映射配置(定制模板用)；缺省依次查找环境变量 '
                         'WORD2MD_STYLES 与 ~/.word2md/styles.json')
    ap.add_argument('--verify', action='store_true',
                    help='转换后立即做内容级一致性校验；不一致时退出码为 1')
    ap.add_argument('--fail-under', type=float, default=None, metavar='SCORE',
                    help='配合 --verify：AI 可读性分数低于该阈值时退出码为 1')
    ap.add_argument('--probe', action='store_true',
                    help='配合 --verify：额外输出 AI 可读性诊断摘要')
    args = ap.parse_args(argv)

    stats = convert_word_to_markdown(
        args.input, out_dir=args.output_dir,
        extract_images=not args.no_images,
        code_heuristics=not args.no_code_heuristics,
        include_toc=not args.no_toc,
        code_lang=args.code_lang,
        strip_emphasis=args.no_bold,
        style_config=args.styles)

    print(f"转换完成: {stats['md_path']}")
    print(f"  标题 {stats['headings']} | 表格 {stats['tables']} (GFM {stats['tables_gfm']}/HTML {stats['tables_html']})"
          f" | 图片 {stats['images']} | 代码块 {stats['code_blocks']} | 公式 {stats['formulas']}"
          f" | 目录条目 {stats['toc_entries']}")
    for w in stats['warnings']:
        print(f"  ⚠ {w}")

    if not (args.verify or args.fail_under is not None or args.probe):
        return 0

    # ---- 一步式校验: 转换完立刻证明没丢东西 ----
    # 延迟导入: 让"只想转换"的用户不承担 verify/probe 的模块加载开销
    from verify_consistency import verify

    report = verify(args.input, stats['md_path'])
    par = report['paragraphs']
    tab = report['tables']
    fml = report['formulas']
    print()
    print(f"【校验】段落 {par['pct']}% ({par['matched']}/{par['word_n']})"
          f" | 表格 {tab['pct']}% (单元格 {tab['cell_matched']}/{tab['cell_total']})"
          f" | 公式 {fml['word']}/{fml['md']}"
          f" | 图片缺失 {len(report['images'].get('missing') or [])}")
    print(f"【判定】{'通过 ✓' if report['pass'] else '存在差异 ✗'}")

    problems = []
    if not report['pass']:
        problems.append('内容级校验未通过')

    score = None
    if args.probe or args.fail_under is not None:
        from review_probe import probe
        pr = probe(args.input, stats['md_path'], sample=0)
        score = pr['readability']['score']
        print(f"【可读性】{score} 分 / {pr['readability']['grade']} 级"
              f"，缺陷 {len(pr['defects'])} 项")

    if args.fail_under is not None and score is not None and score < args.fail_under:
        problems.append(f'可读性 {score} < 阈值 {args.fail_under}')

    if problems:
        print('【门禁】' + '；'.join(problems), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
