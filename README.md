# Word → Markdown 可复用转换流程（word2md）

将任意 Word 文档（.docx）转换为结构保真的 Markdown。本工具源自某芯片规格书转换项目的实战沉淀，经内容级全量校验（正文段落流 100%、241 个表格逐格 100% 一致）验证。

---

## 一、快速开始

```bash
# CLI：转换单个文档
python word2md.py 输入.docx -o 输出目录

# 库调用
from word2md import convert_word_to_markdown
stats = convert_word_to_markdown('input.docx', out_dir='out')
print(stats['md_path'])    # 输出的 .md 路径
print(stats['warnings'])   # 转换警告
```

输出约定：Markdown 与同名 `images/` 目录（图片提取后自动光栅化 EMF/WMF 为 PNG）。

转换后校验：

```bash
# CLI：内容级一致性校验（Word 原文 ↔ 生成 Markdown）
python verify_consistency.py 输入.docx 输出.md

# 库调用
from verify_consistency import verify
report = verify('输入.docx', '输出.md')
print(report['pass'])      # 全部通过为 True
```

## 二、支持的元素

| 元素 | 处理方式 |
| --- | --- |
| 标题 | Heading 1-6 / Title / 中文「标题 1-4」样式 → `#`~`######` |
| 段落 | 普通段落保序输出，空段保留 |
| 富文本 | 粗体/斜体 → `**`/`*`；上下标 → `<sub>`/`<sup>`；删除线 → `<s>`；拼音注音（ruby）→ 只输出主体文本 |
| 表格 | 无合并 → GFM 管道；含合并单元格/嵌套表 → HTML `<table border="1">` + colspan/rowspan 忠实还原 |
| 图片 | 自动提取到 `images/`；EMF/WMF 矢量图用 Pillow 3× LANCZOS 光栅化为 PNG |
| 列表 | 编号列表按 abstractNum 还原有序/无序与层级缩进 |
| 代码块 | 启发式识别（C 前缀/赋值/控制语句），围栏包裹 |
| 链接 | 内部锚点 → `[文字](#slug)`；外部链接保留 |
| 公式 | OMML（Word 公式）→ LaTeX（`$...$` 行内 / `$$...$$` 块级，含自动编号） |
| 目录 | Word 目录域解析为可点击锚点目录，并在原文档位置输出（不移到文首） |

## 三、CLI 参数

```
python word2md.py 输入.docx [-o 输出目录] [--no-images] [--no-code-heuristics] [--no-toc] [--no-bold] [--code-lang c]
```

与库参数对应：`extract_images`、`code_heuristics`、`include_toc`、`strip_emphasis`、`heading_styles`（额外标题样式映射）、`code_lang`。

**`--no-bold`（AI 阅读降噪模式）**：剥离加粗/粗斜体强调语法（GFM 的 `**`/`***` 与 HTML 表格的 `<strong>`/`<em>`），消除「**注意：该位仅可在****BUSY=0****时...**」式碎片噪声；全部文本内容与 `<sub>/<sup>/<s>/<br>` 功能标签保留，供 API 直传原始 Markdown 给 LLM 阅读的场景使用。

## 四、适用范围

**适用**（预期高质量输出）：

- 常规 Office 文档：报告、规格书、手册、方案、论文等「文字 + 表格 + 图片」为主的 .docx
- 中文/英文/多语言混排（Unicode 全支持）
- 含合并单元格、嵌套表、数学公式的专业文档

**不适用 / 效果受限**：

| 场景 | 原因与建议 |
| --- | --- |
| 旧版 .doc 二进制格式 | zip 容器无法直读。先 Word/WPS 另存为 .docx，或 `soffice --headless --convert-to docx` |
| 扫描件/纯图片 PDF 转来的文档 | 无文本层，需 OCR 工具（如 PaddleOCR）前置处理 |
| 复杂文本框/艺术字（w:txbxContent） | 文本框内容不在正文流，需单独抽取 |
| 嵌入对象（Excel 图表、Visio、内嵌 OLE） | 转换器不解析 OLE 二进制，会跳过 |
| 宏（VBA） | docx 不含宏（宏在 .docm），转出的 .docx 中宏已经丢失 |
| 批注/修订 | 目前不导出批注与修订痕迹；修订内容按当前显示版本处理 |
| 本身就是讲 Markdown 语法的文档 | 正文示例（`# 一级标题`、`\|列\|`）会被渲染器当作真实语法解析，可视差异不代表内容丢失（已验证字符级 0 缺失） |

## 五、失败边界与降级行为

1. **样式缺失**：文档未用标准 Heading 样式 → 标题识别不到。可通过 `heading_styles` 参数补充自定义样式映射。
2. **图片格式**：Pillow 不支持的嵌入格式（如 .tif 特殊编码）会以原格式落盘并在 `warnings` 中提示。
3. **代码块误判**：启发式识别可能把非代码段落（如 `EA=1；` 样式的寄存器描述）围进代码块。出现此类误报时用 `--no-code-heuristics` 关闭。
4. **合并表格**：HTML 表格在无 CSS 的环境（如某些纯文本 Markdown 预览器）边框样式不显示，但结构与内容完整。
5. **超长表格/嵌套极深**：多层嵌套表依赖浏览器 HTML 解析，极深嵌套可能渲染异常（本工具支持一级嵌套，实测通过）。

## 六、如何校验转换结果

三层校验，对应 `verify_consistency.py` 的返回字段：

1. **结构层**：段落流块数两侧相等、表格数两侧相等（`paragraphs.word_n == md_n`、`tables.n_total`）。
2. **内容层**：
   - 正文段落流逐块比对（`paragraphs.pct` 应为 100%）
   - 表格逐格比对（`tables.pct` 应为 100%）
   - 公式占位数量对齐（`formulas.word == formulas.md`）
3. **资源层**：图片引用全部存在（`images.missing` 应为空）。

判定口径（已内置）：空白归一、装饰星号剥离、链接/HTML 标签剥离、代码块逐行对齐、ruby 注音两侧丢弃、vMerge 占位格与空单元格两侧跳过、嵌套表文本并入所在格、公式统一占位。

`report['pass'] == True` 表示四项全过。若为 False，`paragraphs.differences` 与 `tables.diff_indices` 会给出具体差异位置供人工核对。

## 七、文件清单

| 文件 | 职责 |
| --- | --- |
| `word2md.py` | 通用转换库 + CLI（主入口） |
| `verify_consistency.py` | 内容级一致性校验器 + CLI |
| `math_latex.py` | OMML→LaTeX 转换器（单元级移植自 eqword2llm，MIT） |
| `tools/convert_emf.ps1` | 独立 EMF/WMF 高分光栅化脚本（Windows PowerShell，可选，转换器内置了 Pillow 路径） |

历史验证基准（内部项目，样例名已脱敏）：`<样例规格书>.docx` —— 标题 185、表格 241（GFM 148 + HTML 93）、图片 58、公式 14、代码块 4、目录 164 条；段落流与表格流 100% 一致。

## 八、隐私与脱敏说明

本仓库为内部工具的开源化版本，已做以下脱敏处理：

- **移除全部内部测试文档本体与转换产物**（源 .docx、生成的 .md、提取图片等均不入库）；
- **移除硬编码本地绝对路径**（原代码曾内置内部项目文件路径，现已移除并改为必填参数，缺失时抛出明确错误）；
- **泛化内部芯片型号与寄存器名**（样式映射注释、校验器示例注释中的真实型号已替换为通用示例，并标注 `[DESENSITIZED]`）;
- **保留全部公开可展示的技术指标**：元素覆盖清单、命令行参数、三层校验口径、历史实测准确率（段落流/表格流 100% 一致）等。

如发现遗漏的敏感信息，欢迎通过 Issue 反馈。