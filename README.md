# word2md — Word 转 Markdown 工具包

为**样式重度的技术文档**做的 Word→Markdown 转换器：合并单元格忠实还原、转换后逐段逐格校验、
只依赖两个 Python 包（不装 Pandoc、不下模型、不联网）。

**版本**：1.0.1　·　[PyPI](https://pypi.org/project/word-vane/)　·　变更记录见 [CHANGELOG](https://github.com/A-electric-sheep/word-vane/blob/main/CHANGELOG.md)

## 它解决什么、不解决什么

先明确边界，省得你白试。

**适合**：中文技术文档（规格书、手册、白皮书、方案）迁往 Confluence / 语雀 / Docusaurus /
知识库；文档含**大量合并单元格**、OMML 公式、自动目录、拼音注音、EMF 矢量图；
环境**不能装 Pandoc、不能下模型**；需要**向他人证明**转换没丢内容。

**不适合**（请直接用别的工具）：

| 你的需求 | 建议 |
|---|---|
| 转 PDF / PPT / Excel / EPUB | [anydoc](https://github.com/firecrawl/anydoc)（14 种格式，毫秒级） |
| 扫描件 OCR | [MinerU](https://github.com/opendatalab/MinerU) / [Marker](https://github.com/datalab-to/marker) |
| 脚注/尾注、批注、修订痕迹 | Pandoc、[docx2md](https://github.com/ucli-tools/ms2md) |
| 表格重、要 ML 级结构识别 | [Docling](https://github.com/docling-project/docling) |
| 一次性转个简单文档 | 任选网页工具，不必装环境 |

## 为什么不用通用转换器

GFM 管道表语法上**没有** `colspan`/`rowspan`。所以"docx → GFM"路线的工具遇到合并单元格，
只能三选一：重复值、留空、丢表——**三种做法都让读者无法判断原表有几处合并、跨几行几列**。

本工具对含合并的表输出 HTML `<table>`，保留原生 `colspan`/`rowspan`。

实测对照（2026-11-04，`firecrawl-anydoc` 2.x，夹具 5 张表 / 真值 4 处横向 + 3 处纵向）：

| 工具 | colspan 属性 | rowspan 属性 | 输出形态 |
|---|---|---|---|
| **word2md** | **4/4** | **3/3** | HTML `<table>` + colspan/rowspan |
| anydoc | 0/4 | 0/3 | GFM 管道表，合并语义不可恢复 |

复现：`python bench/merged_cells_bench.py --with-anydoc`（口径与局限见 [bench/README](https://github.com/A-electric-sheep/word-vane/blob/main/bench/README.md)）

> 说明：这是**构造夹具**上的单点对照，不代表两工具的整体优劣。anydoc 在格式覆盖、
> 速度（中位 4.4ms vs 本项目 20–40s）上远优于本项目。如果你要处理的文档没有合并单元格，
> 用 anydoc 更合适。

## 关于"转换后校验"

转换后自动校验不是本项目首创——[`ucli-tools/docx2md`](https://github.com/ucli-tools/ms2md)
的 `docx2md verify` 早已提供同类能力，思路值得致谢。

本项目的差异在于把它做成了**工作流而非单点功能**：`--verify` 一步完成转换+校验、
`--fail-under` 把可读性分数接成 CI 门禁、`REVIEW_SPEC.md` 定义了人工复核的判定口径
（PASS / WARN / FAIL）与证据要求。

## 安装

```bash
pip install word-vane
```

> 安装名是 **`word-vane`**，导入名与命令行仍是 `word2md`。
> PyPI 上的 `word2md` 是**另一个项目**（hnrobert/word2md），与本项目无关，请勿混装。

需要 Python 3.9 及以上。仅两个依赖：`python-docx`（解析 .docx）、`Pillow`（图片光栅化）。
不依赖 Office / LibreOffice，无需联网，可离线部署。

从源码安装（开发用）：

```bash
git clone https://github.com/A-electric-sheep/word-vane.git
cd word-vane
pip install -e ".[dev]"
```

## 最短路径：一条命令

```bash
word2md 你的文档.docx -o 输出目录 --verify
```

输出 `输出目录/你的文档.md` + `输出目录/images/*.png`，并打印校验结论：

```
【校验】段落 100.0% (1240/1240) | 表格 100.0% (单元格 3871/3871) | 公式 14/14 | 图片缺失 0
【判定】通过 ✓
```

内容有差异时退出码为 `1`，可直接用于 CI。

命令行工具一览（`pip install` 后即可用）：

| 命令 | 作用 |
|---|---|
| `word2md` | 转换单个文档；`--verify` 附带校验 |
| `word2md-verify` | 内容级一致性校验（退出码 0/1） |
| `word2md-batch` | 批量转换整个目录 |
| `word2md-batch-verify` | 批量校验；支持 `--fail-under` |
| `word2md-probe` | AI 可读性诊断（0–100 分 + 缺陷清单） |
| `word2md-mcp` | 以 MCP server 方式提供能力（需 `pip install "word-vane[mcp]"`） |

也可直接运行脚本，无需安装：`python word2md.py 你的文档.docx -o 输出目录`

校验输出最末行的 `【总判定】 通过 ✓` 即表示内容完整。

---

## 各脚本分工

| 场景 | 模块 | 命令行 | 脚本直调 |
|---|---|---|---|
| 转单个文档 | `word2md.py` | `word2md 输入.docx -o 输出目录` | `python word2md.py ...` |
| 校验单个文档 | `verify_consistency.py` | `word2md-verify 输入.docx 输出.md` | `python verify_consistency.py ...` |
| 转整个目录 | `batch_convert.py` | `word2md-batch 文档根目录 --out 输出根目录` | 同左 |
| 批量校验 | `batch_verify.py` | `word2md-batch-verify 文档根目录 --converted 输出根目录` | 同左 |
| AI 可读性诊断 | `review_probe.py` | `word2md-probe 输入.docx 输出.md` | 同左 |

`math_latex.py` 是 `word2md.py` 的依赖（公式转 LaTeX），不需单独调用。

### 批量处理示例

```bash
# 递归转换整个目录，镜像原目录结构输出
python batch_convert.py ./文档库 --out ./转换结果 --workers 4

# 批量校验（--converted 指向与上一步相同的输出根目录）
python batch_verify.py ./文档库 --converted ./转换结果 --workers 4
```

批量脚本的结果 JSON **写入当前工作目录**（不污染源文档目录），文件名分别为
`batch_convert_result.json`、`batch_verify_result.json`。可用 `--json 路径` 自定义。

---

## 转换器参数

```
python word2md.py 输入.docx [-o 输出目录] [--no-images] [--no-code-heuristics]
                            [--no-toc] [--no-bold] [--code-lang c]
                            [--verify] [--probe] [--fail-under SCORE] [--version]
```

| 参数 | 作用 | 什么时候用 |
|---|---|---|
| `-o, --output-dir` | 输出目录 | 缺省写在输入文件旁的 `<名>_md/` |
| `--no-images` | 不提取图片 | 只要文字，不需要图片时 |
| `--no-code-heuristics` | 关闭代码块识别 | 寄存器描述被误判成代码块时 |
| `--no-toc` | 跳过目录域处理 | 原文目录混乱、不想要目录时 |
| `--no-bold` | 剥离 `**` 加粗语法 | **给 AI 阅读时推荐**，消除加粗碎片噪声 |
| `--code-lang` | 代码块语言标记 | 默认 `c`，可改 `python`/`verilog` 等 |
| `--styles` | 外部样式映射配置 JSON | 模板用了自定义章节样式名时，见「自定义样式映射」 |
| `--verify` | 转换后立即做一致性校验 | 想一条命令拿到"没丢东西"的结论 |
| `--probe` | 附带可读性诊断摘要 | 需要知道生成的 MD 适不适合喂 LLM |
| `--fail-under` | 可读性阈值，低于则退出码 1 | 接进 CI / 质量门禁 |

### 何时该开 `--no-bold`

若转换目的是把 Markdown 直接喂给大模型阅读，建议开启。Word 常把一个词拆成多个格式相同的片段，会产生 `**注意：该位仅可在****BUSY=0****时...**` 这类碎片。开启后文本内容与 `<sub>/<sup>/<br>` 等功能标签全部保留，仅去掉强调符号。

---

## 支持的文档元素

| 元素 | 处理方式 |
|---|---|
| 标题 | Heading 1-6 / Title / 中文「标题 1-4」/ 自定义章节样式 → `#`~`######` |
| 段落 | 保序输出，空段保留 |
| 富文本 | 粗体/斜体 → `**`/`*`；上下标 → `<sub>`/`<sup>`；删除线 → `<s>`；拼音注音只留正文 |
| 表格 | 无合并 → GFM 管道表；含合并单元格/嵌套表 → HTML `<table>` + colspan/rowspan |
| 图片 | 自动提取到 `images/`；EMF/WMF 矢量图 3× LANCZOS 光栅化为 PNG |
| 列表 | 按 Word 编号定义还原有序/无序与层级缩进（含**样式级编号**） |
| 代码块 | 启发式识别（C 前缀/赋值/控制语句/函数调用/自增/8051 汇编），自动排除数学公式与图注 |
| 链接 | 内部锚点 → `[文字](#slug)`；外部链接保留 |
| 公式 | OMML → LaTeX（行内 `$...$` / 块级 `$$...$$`） |
| 目录 | 解析 Word 目录域为可点击锚点，**在原文位置**输出 |

### 代码块判定规则（v1.1）

`_looks_like_code` 按以下规则判定，任一命中即为代码行：

| 规则 | 形态 | 示例 |
|---|---|---|
| C 语言前缀 | 以声明关键字开头 | `#include <reg51.h>`、`unsigned char i;` |
| 赋值语句 | `左值 运算符 右值` | `EA = 1;`、`P1 [3] = 0;`、`p->x = 1;` |
| 控制语句 | `for/if/while/switch/return` 或单独括号 | `for(i=0;i<12;i++)`、`{`、`}` |
| 函数调用 | `标识符(参数);` **必须带分号** | `_nop_();`、`I2C_START();` |
| 自增自减 | `左值++;` 或 `++左值;` | `counter++;`、`++i;` |
| 8051 汇编 | 助记符(全大写) + 合法操作数；含分号即为代码 | `MOV R1,#83H; 注释`、`NOP`、`JC $+50H` |

**排除规则**（防止正文被误包进代码块）：

- 无分号且等号右侧为纯标识符算式 → 判为数学关系式，如 `duty1 = duty0`
- 无分号且等号右侧含 2 个以上英文单词 → 判为说明文字/图注

**关键区分线是分号**。正文中形如 `REG_A1 (CBH)`、`SOP16L(150mil)`、
`PWMn周期寄存器高8位(读/写)` 的寄存器与封装说明**没有分号**，因此不会被误判；
而真实代码行几乎总以分号收尾。改动此逻辑时请守住这条线。

**汇编的额外坑**：汇编代码与「指令说明」形状几乎一样，只看首词必然误伤。
`_looks_like_asm` 按以下优先级判定：

| 顺序 | 判据 | 结果 |
|---|---|---|
| 1 | 含 `;` | 代码（汇编注释一律以 `;` 起头） |
| 2 | 含全角 `（`，或操作数后大间距接 `(` | 说明文字，不入块 |
| 3 | 逐个操作数校验 8051 合法形态 | 合法才算代码 |
| 4 | 助记符非全大写 | 不入块（排除 `Push the PC value…`） |

对照：

| 真汇编代码 | 指令说明（正文） |
|---|---|
| `MOV R1,#83H<TAB>; 注释` | `MOV  A, #50H    （这条指令是将立即数50H送到累加器A中）` |
| `NOP` | `INC  R0      是指（R0）+1→R0` |

无操作数指令目前只认 `NOP`——`RET`/`RETI` 若放宽会误伤「RET 指令说明」类正文。

**语言标记**：`flush_code()` 按块内内容判定，过半行为汇编则标 `asm`，
否则用 `--code-lang` 指定值（默认 `c`）。C 块混一两行汇编注释仍标 `c`。

---

## 适用范围与已知限制

**适用**（预期高质量输出）：以「文字 + 表格 + 图片」为主的技术类 `.docx`；中英文混排；
含合并单元格、嵌套表、数学公式的专业文档。

**不适用 / 效果受限**：

| 场景 | 说明与建议 |
|---|---|
| 旧版 `.doc` 格式 | 先另存为 `.docx`，或 `soffice --headless --convert-to docx 文件.doc` |
| 扫描件 / 纯图片 PDF | 无文本层，需先 OCR |
| **脚注 / 尾注** | **当前不导出**（已知缺口）；有此需求请用 Pandoc 或 docx2md |
| 文本框 / 艺术字内容 | 图形内的文字不在正文流，按设计不抽取；需从原图查看 |
| 嵌入对象（Excel 图表、Visio、OLE） | 不解析 OLE 二进制，会跳过 |
| 批注 / 修订痕迹 | 不导出批注；修订按当前显示版本处理 |
| 宏（VBA） | `.docx` 不含宏；`.docm` 转出后宏丢失 |
| 自定义标题样式 | 若文档未用标准 Heading 样式，标题可能识别不到；可用 `heading_styles` 参数补充映射 |
| 超大文档性能 | 单份含大量 EMF/WMF 的文档约 20–40 秒，耗时主要在矢量图光栅化；可用 `--no-images` 跳过 |

### 失败边界

1. **样式缺失**：文档未用标准 Heading 样式 → 改用 `heading_styles` 传自定义映射。
2. **图片格式异常**：Pillow 不支持的格式会以原格式落盘，并在 `warnings` 中提示。
3. **代码块误判**：用 `--no-code-heuristics` 关闭。
4. **HTML 表格无边框**：部分纯文本 Markdown 预览器不渲染边框样式，但内容完整。

---

## 校验与诊断

### 一致性校验（verify_consistency.py）

三层校验，`report['pass'] == True` 表示四项全过：

1. **结构层**：段落流块数、表格数两侧相等
2. **内容层**：段落逐块比对、表格逐格比对、公式占位数量对齐
3. **资源层**：图片引用全部存在

若为 `False`，`paragraphs.differences` 与 `tables.diff_indices` 会给出具体差异位置。

`word2md-verify` 的退出码为 `0`（全部通过）或 `1`（存在差异），可直接作为 CI 门禁：

```bash
word2md doc.docx -o out/
word2md-verify doc.docx out/doc.md || exit 1
```

### AI 可读性诊断（review_probe.py）

分析生成的 Markdown 是否适合交给 AI 阅读，输出结构指标、缺陷清单与 0-100 评分：

```bash
word2md-probe 输入.docx 输出.md --sample 8
```

检查项：标题层级是否可导航、表格列数是否一致、图片引用是否存在、公式是否闭合、有无残留占位符、目录锚点是否有效。

## 自定义样式映射

Word 模板常把章节标题定义成**自定义样式**（不是内置的 `Heading 1/2/3`），这时需要告诉
转换器「哪个样式对应几级标题」。本工具有两条路径：先看通用规则能否自动命中，不行再配。

### 第一层：通用结构规则（无需配置）

样式名符合下列命名习惯时会被**自动识别**，与具体前缀无关：

| 样式名形态 | 示例 | 判定 |
|---|---|---|
| `<前缀><数字>-<层级>` | `chapter1-2`、`sec2-3`、`part1-4` | 后缀数字即层级（h1~h4） |
| `<前缀><数字>-<名字>` | `chapter1-content`、`part2-cover` | 视为一级章节（h1） |
| `<前缀><数字>` | `section1`、`chapter2` | 视为一级章节（h1），并排除 `list`/`bullet`/`table` 等非章节词根 |
| `*多栏` / `*multicol` | `正文多栏` | 正文像小标题时提升为 h3 |

其中 `<前缀><数字>-<层级>` 形态明确，直接采信；后两种歧义较大，会额外校验正文特征
（长度、无句末标点、非图注、非「数值+单位」条目）才提升。

### 第二层：外部配置（用于特例）

命名不符合上述规则、或需要覆盖默认判定的，用 JSON 配置注入。**源码里不硬编码任何
特定厂商或模板的专有样式名**——这类名称属于内部资产，不应该出现在公开仓库里。

配置文件查找顺序（前者优先）：

1. `--styles 路径`
2. 环境变量 `WORD2MD_STYLES` 指向的 JSON
3. `~/.word2md/styles.json`

格式（各键均可选）：

```json
{
  "heading_styles": { "mycover1": 1, "mybodytex": 3 },
  "toc_styles": { "mytoc 1": 1 },
  "subheading_patterns": ["^mymulticol\\d*$"]
}
```

```bash
word2md doc.docx -o out/ --styles ./my-styles.json

# 或一次性设置，之后所有转换都生效
export WORD2MD_STYLES=/path/to/my-styles.json
```

库调用同样支持：

```python
from word2md import convert_word_to_markdown

convert_word_to_markdown('doc.docx', out_dir='out',
                         style_config='./my-styles.json',
                         # 也可直接传字典，优先级高于配置文件
                         heading_styles={'mycover1': 1})
```

### 排查：标题没被识别

不确定某个样式名为何没生效时，先用 `word2md-probe` 看实际结构，或检查该样式是否真的
应用到了段落上（Word 里 `Ctrl+Shift+S` 可查看当前样式名）。常见原因有两个：

- 文档未使用样式、而是手工放大字号加粗——这类「伪标题」无法通过样式识别；
- 样式名属于特例且未配置，按上面第二层补一条即可。

## 质量门禁（接进 CI）

三个层面都可作为构建门禁，失败时退出码为 `1`：

```bash
# 1) 内容级：转换 + 校验一条命令，有差异即失败
word2md doc.docx -o out/ --verify || exit 1

# 2) 可读性：低于 90 分即失败（适合"喂给 LLM 之前"做把关）
word2md doc.docx -o out/ --verify --probe --fail-under 90 || exit 1
word2md-probe doc.docx out/doc.md --fail-under 90 || exit 1

# 3) 批量：任一份段落或表格一致率低于 99% 即失败
word2md-batch docs/ --out out/ || exit 1
word2md-batch-verify docs/ --converted out/ --fail-under 99 || exit 1
```

典型 CI 片段：

```yaml
- name: Convert and gate
  run: |
    pip install word-vane
    word2md manual.docx -o build/ --verify --fail-under 90
```

## MCP Server（给 Agent 用）

```bash
pip install "word-vane[mcp]"
```

只暴露**一个**工具 `word2md_convert_and_verify`，一次调用完成转换 → 校验 → 评分。
不拆成三个工具，是因为拆开会给每次调用增加一轮模型往返，还要在对话里传中间产物路径。

工具参数与返回：

| 参数 | 说明 |
|---|---|
| `path` | 本地 `.docx` 绝对路径（**不接受 URL / data URI**） |
| `out_dir` | 输出目录，缺省落系统临时目录，不写源目录 |
| `fail_under` | 可读性阈值，低于则 `status` 为 `below_threshold` |
| `keep_images` | 是否提取图片，默认 `true` |

返回 `status`（`ok` / `mismatch` / `below_threshold` / `error`）、`converted`、
`readability{score,grade,verdict}`、`consistency{paragraph_pct,table_pct,images_missing}`、
`structure{...}`、`artifacts{md_path,images_dir}`、`warnings`。

在 MCP 客户端里配置：

```json
{
  "mcpServers": {
    "word2md": {
      "command": "word2md-mcp",
      "args": []
    }
  }
}
```

安全边界：只接受本地路径（拒绝 URL，避免 SSRF）；不加载第三方 MCP 插件；
不执行文档内嵌宏或 OLE 对象。

兼容性：mcp SDK 1.x（`FastMCP`）与 2.x（`MCPServer`）均支持，启动时自动选择。

---

## 库调用

```python
import sys
sys.path.insert(0, '此目录路径')

from word2md import convert_word_to_markdown
stats = convert_word_to_markdown('input.docx', out_dir='out')
print(stats['md_path'])    # 输出 .md 路径
print(stats['warnings'])   # 转换警告

from verify_consistency import verify
report = verify('input.docx', 'out/input.md')
print(report['pass'])
```

常用关键字参数：`extract_images`、`code_heuristics`、`include_toc`、`strip_emphasis`、`heading_styles`、`code_lang`。

---

## 文件清单

| 文件 | 职责 | 必需 |
|---|---|---|
| `word2md.py` | 转换器主程序 + CLI | 是 |
| `math_latex.py` | OMML→LaTeX（word2md 依赖） | 是 |
| `verify_consistency.py` | 内容级一致性校验器 | 是 |
| `batch_convert.py` | 批量转换驱动 | 批量时 |
| `batch_verify.py` | 批量校验驱动 | 批量时 |
| `review_probe.py` | AI 可读性诊断 | 可选 |
| `mcp_server.py` | MCP server（单一工具 `convert_and_verify`） | 可选 |
| `tools/convert_emf.ps1` | EMF/WMF 高分光栅化脚本（Windows PowerShell，可选） | 可选 |
| `pyproject.toml` | 打包与工具链配置（ruff / pytest） | 是 |
| `requirements.txt` | 依赖清单（等价于 pyproject 的 dependencies） | 可选 |
| `REVIEW_SPEC.md` | 复核规程（供 AI 子代理遵循） | 可选 |
| `CHANGELOG.md` | 变更日志（含已知限制与版本号来源说明） | 是 |
| `tests/` | pytest 测试（夹具运行时生成，仓库不含 .docx） | 开发 |
| `bench/` | 合并单元格保真度基准（可复现对照） | 可选 |

---

## 开发

```bash
pip install -e ".[dev]"

pytest                                # 84 条用例（70 个测试函数，参数化展开）
ruff check .                          # 风格检查
python bench/merged_cells_bench.py    # 合并单元格基准
```

CI（GitHub Actions）在 Ubuntu 与 Windows 上跑 Python 3.9 / 3.11 / 3.13 三档，
并验证 wheel 构建与安装后 CLI 可用。

测试夹具全部在运行时用 `python-docx` 生成，仓库**不保存任何 `.docx` 二进制**——
既便于审阅"测的是什么"，也保证公开仓库不携带真实文档。

### 基准数据说明

本仓库不宣称未经复现的准确率数字。所有性能/保真度结论都附可执行脚本，
你可以用 `bench/` 在自己的文档上重跑，以你的语料结果为准。

---

## 常见问题

**Q：转换后标题全部丢失？**
文档没用标准 Heading 样式。检查 Word 里标题用的样式名，然后用 `heading_styles={'你的样式名': 1}` 传入映射。

**Q：转换后表格变成了一堆 `<table>` 标签？**
原表格含合并单元格或嵌套表，这是有意为之——GFM 管道表无法表达 `colspan`/`rowspan`，只能用 HTML 保留结构。浏览器渲染正常。

**Q：列表编号错乱或丢失？**
先确认 Word 里编号是否正常显示。本工具已支持样式级编号继承，若仍异常请附上文档样式名反馈。

**Q：批量转换的结果 JSON 找不到？**
它在**你执行命令时所在的目录**，不是输入目录。或用 `--json 指定路径`。

**Q：转换很慢？**
单份含大量 EMF/WMF 图片的文档约需 20~40 秒，主要耗时在矢量图光栅化。批量处理时用 `--workers` 提高并发（建议不超过 CPU 核数）。