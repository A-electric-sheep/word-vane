# word2md — Word 转 Markdown 工具包

将 Word 文档（`.docx`）转换为结构保真的 Markdown，并校验转换结果的完整性。

经 70 份真实技术文档全量验证：表格 100% 逐格一致、公式 100% 对齐、图片零缺失、段落一致率 99.998%。

**版本**：1.0.0

---

## 一、安装

需要 Python 3.9 及以上。

```bash
pip install -r requirements.txt
```

只需两个依赖：`python-docx`（解析 .docx）、`Pillow`（图片光栅化）。不依赖 Office 或 LibreOffice，无需联网。

---

## 二、最短路径：三步跑通

以下命令在**本文件所在目录**执行。

```bash
# 1. 转换
python word2md.py 你的文档.docx -o 输出目录

# 2. 校验（强烈建议，确认无内容丢失）
python verify_consistency.py 你的文档.docx 输出目录/你的文档.md

# 3. 查看结果
# 输出：输出目录/你的文档.md  +  输出目录/images/*.png
```

校验输出最末行的 `【总判定】 通过 ✓` 即表示内容完整。

---

## 三、四个脚本的分工

| 场景 | 脚本 | 用法 |
|---|---|---|
| 转单个文档 | `word2md.py` | `python word2md.py 输入.docx -o 输出目录` |
| 校验单个文档 | `verify_consistency.py` | `python verify_consistency.py 输入.docx 输出.md` |
| 转整个目录 | `batch_convert.py` | `python batch_convert.py 文档根目录 --out 输出根目录` |
| 批量校验 | `batch_verify.py` | `python batch_verify.py 文档根目录 --converted 输出根目录` |

`math_latex.py` 是 `word2md.py` 的依赖（公式转 LaTeX），不需单独调用。
`review_probe.py` 是进阶诊断工具，用于分析 Markdown 的 AI 可读性（见第七节）。

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

## 四、转换器参数

```
python word2md.py 输入.docx [-o 输出目录] [--no-images] [--no-code-heuristics]
                            [--no-toc] [--no-bold] [--code-lang c] [--version]
```

| 参数 | 作用 | 什么时候用 |
|---|---|---|
| `-o, --output-dir` | 输出目录 | 缺省写在输入文件旁的 `<名>_md/` |
| `--no-images` | 不提取图片 | 只要文字，不需要图片时 |
| `--no-code-heuristics` | 关闭代码块识别 | 寄存器描述被误判成代码块时 |
| `--no-toc` | 跳过目录域处理 | 原文目录混乱、不想要目录时 |
| `--no-bold` | 剥离 `**` 加粗语法 | **给 AI 阅读时推荐**，消除加粗碎片噪声 |
| `--code-lang` | 代码块语言标记 | 默认 `c`，可改 `python`/`verilog` 等 |

### 何时该开 `--no-bold`

若转换目的是把 Markdown 直接喂给大模型阅读，建议开启。Word 常把一个词拆成多个格式相同的片段，会产生 `**注意：该位仅可在****BUSY=0****时...**` 这类碎片。开启后文本内容与 `<sub>/<sup>/<br>` 等功能标签全部保留，仅去掉强调符号。

---

## 五、支持的文档元素

| 元素 | 处理方式 |
|---|---|
| 标题 | Heading 1-6 / Title / 中文「标题 1-4」/ 自定义章节样式 → `#`~`######` |
| 段落 | 保序输出，空段保留 |
| 富文本 | 粗体/斜体 → `**`/`*`；上下标 → `<sub>`/`<sup>`；删除线 → `<s>`；拼音注音只留正文 |
| 表格 | 无合并 → GFM 管道表；含合并单元格/嵌套表 → HTML `<table>` + colspan/rowspan |
| 图片 | 自动提取到 `images/`；EMF/WMF 矢量图 3× LANCZOS 光栅化为 PNG |
| 列表 | 按 Word 编号定义还原有序/无序与层级缩进（含**样式级编号**） |
| 代码块 | 启发式识别（C 前缀/赋值/控制语句），自动排除数学公式与图注 |
| 链接 | 内部锚点 → `[文字](#slug)`；外部链接保留 |
| 公式 | OMML → LaTeX（行内 `$...$` / 块级 `$$...$$`） |
| 目录 | 解析 Word 目录域为可点击锚点，**在原文位置**输出 |

---

## 六、适用范围与已知限制

**适合**：报告、技术规格书、手册、方案、论文等「文字 + 表格 + 图片」为主的 `.docx`；中英文混排；含合并单元格、嵌套表、数学公式的专业文档。

**不适合 / 效果受限**：

| 场景 | 说明与建议 |
|---|---|
| 旧版 `.doc` 格式 | 先另存为 `.docx`，或 `soffice --headless --convert-to docx 文件.doc` |
| 扫描件 / 纯图片 PDF | 无文本层，需先 OCR |
| 文本框 / 艺术字内容 | 图形内的文字不在正文流，按设计不抽取；需从原图查看 |
| 嵌入对象（Excel 图表、Visio、OLE） | 不解析 OLE 二进制，会跳过 |
| 批注 / 修订痕迹 | 不导出批注；修订按当前显示版本处理 |
| 宏（VBA） | `.docx` 不含宏；`.docm` 转出后宏丢失 |
| 自定义标题样式 | 若文档未用标准 Heading 样式，标题可能识别不到；可用 `heading_styles` 参数补充映射 |

### 失败边界

1. **样式缺失**：文档未用标准 Heading 样式 → 改用 `heading_styles` 传自定义映射。
2. **图片格式异常**：Pillow 不支持的格式会以原格式落盘，并在 `warnings` 中提示。
3. **代码块误判**：用 `--no-code-heuristics` 关闭。
4. **HTML 表格无边框**：部分纯文本 Markdown 预览器不渲染边框样式，但内容完整。

---

## 七、校验与诊断

### 一致性校验（verify_consistency.py）

三层校验，`report['pass'] == True` 表示四项全过：

1. **结构层**：段落流块数、表格数两侧相等
2. **内容层**：段落逐块比对、表格逐格比对、公式占位数量对齐
3. **资源层**：图片引用全部存在

若为 `False`，`paragraphs.differences` 与 `tables.diff_indices` 会给出具体差异位置。

### AI 可读性诊断（review_probe.py）

分析生成的 Markdown 是否适合交给 AI 阅读，输出结构指标、缺陷清单与 0-100 评分：

```bash
python review_probe.py 输入.docx 输出.md --sample 8
```

检查项：标题层级是否可导航、表格列数是否一致、图片引用是否存在、公式是否闭合、有无残留占位符、目录锚点是否有效。

---

## 八、库调用

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

## 九、文件清单

| 文件 | 职责 | 必需 |
|---|---|---|
| `word2md.py` | 转换器主程序 + CLI | 是 |
| `math_latex.py` | OMML→LaTeX（word2md 依赖） | 是 |
| `verify_consistency.py` | 内容级一致性校验器 | 是 |
| `batch_convert.py` | 批量转换驱动 | 批量时 |
| `batch_verify.py` | 批量校验驱动 | 批量时 |
| `review_probe.py` | AI 可读性诊断 | 可选 |
| `tools/convert_emf.ps1` | EMF/WMF 高分光栅化脚本（Windows PowerShell，可选） | 可选 |
| `requirements.txt` | 依赖清单 | 是 |
| `REVIEW_SPEC.md` | 复核规程（供 AI 子代理遵循） | 可选 |

---

## 十、常见问题

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