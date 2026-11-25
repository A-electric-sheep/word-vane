# 子代理复核规程（Word→Markdown 转换质量）

## 你的任务
对分配给你的每一份文档，独立完成两项检查，并输出结构化结论：

1. **可读性检查**：假设你是一名只拿到这份 Markdown 的 AI 读者，能否正确理解文档结构、定位章节、读懂表格与图片引用。
2. **一致性检查**：把 Markdown 内容与原始 Word 逐段比对，确认文本、表格、公式、图片无遗漏、无篡改、无错位。

## 输入
以下路径均为示例，实际以派发时给出的根目录为准：

- 待复核清单：`<工作根目录>/review_index.json`（按 `root` 分组；你的任务是其中指定的若干组）
- 原始 Word：`<工作根目录>/source_docs/<root>/<子目录>/<name>.docx`
- 转换结果：`<工作根目录>/converted/<root>/<子目录>/<name>/<name>.md`
- 图片目录：与 md 同级的 `images/`

## 工具（必须使用，勿凭印象）
每份文档先跑探针：

```bash
python review_probe.py "<docx绝对路径>" "<md绝对路径>" --sample 8
```

探针输出 JSON：`structure`（体量/标题/表格/图片/公式）、`defects`（结构缺陷）、
`readability`（机器评分与扣分项）、`paragraph_check`（段落块对位）、
`table_check`（表格对位）、`samples`（等间距对位抽样）。这是你的**事实基线**，
不要与之矛盾；若你认为探针判断有误，必须在结论中说明依据。

再跑权威一致性校验：

```bash
python verify_consistency.py "<docx绝对路径>" "<md绝对路径>"
```

## 检查要点（逐项落实）

### A. 可读性（AI 视角）
- 标题层级是否可导航：`#`~`######` 是否呈树状、有无整篇无标题、有无跳级。
- 表格是否可解析：GFM 管道表列数是否一致；HTML 表是否有 `colspan/rowspan`；
  有无表格被压成纯文本。
- 图片是否可用：`![...](images/xxx.png)` 引用是否真实存在。
- 公式是否可读：`$...$`/`$$...$$` 是否闭合、有无 LaTeX 断裂。
- 代码块围栏是否成对、语言标记是否合理。
- 有无残留占位符（`<<TOC_PLACEHOLDER>>`、`[[IMG:`、`\x00`）。
- 目录锚点链接是否指向真实存在的标题。

### B. 一致性（Word 对照）
必须实际打开 Word 原文对照，不能只依赖 md。建议：

```bash
python - <<'PY'
from docx import Document
d = Document(r"<docx绝对路径>")
for i, p in enumerate(d.paragraphs[:60]):
    if p.text.strip():
        print(i, "|", p.style.name, "|", p.text[:100])
PY
```

- 抽查开头、中段、结尾各若干段落，确认 md 中能找到对应文本且语义未变形。
- 抽查若干表格：随机 3~5 个，逐格核对关键单元格（表头、字段名、数值、单位）。
- 确认图片数量与 Word 内嵌图片数量相符（探针给 md 侧数量，Word 侧可用 python-docx 统计 drawing/pict）。
- 若探针 `paragraph_check` 不等，必须定位差异位置并判定属于哪一类：
  - `真实缺失`：内容和语义在 md 中确实找不到 → 必须记录位置与原文。
  - `校验口径`：文本其实存在，只是 Word 侧来自文本框/艺术字等非正文流，
    或归一化差异导致对位失败 → 需说明并举例。
  - `源文档特性`：已人工核定的已知项，归为可忽略，不再重复调查（见下方清单）。

#### 已核定的可忽略项（不必重复排查）

| 项 | 表现 | 结论 |
| --- | --- | --- |
| 图片在翻译流程中被误转为文本框 | Word 用绘图工具（`wps:txbx` / `v:textbox`）画的图，标注文字装在独立文本框里。校验器在 Word 侧读到、MD 侧无对应段落。典型形态：同一张波形图在某一语言版本里被转成上百个文本框，另一语言版本同位置为标准 EMF 图片。 | **根因在文档制作环节（图片被转成文本框），不是转换工具问题。** 应回源文档核对该位置在各语言版本中的图形形态是否一致，确认后再反馈给文档作者修复原文件。 |

### C. 判定分级（严格按此口径输出）
- `PASS`：verify 通过；或未通过但差异经你核实为「校验口径/已知边界」，内容实际完整。
- `WARN`：内容完整但存在结构性瑕疵（标题跳级、表格列不齐、公式断裂、
  少量图片缺失、加粗碎片等），会影响 AI 阅读效率但不影响信息获取。
- `FAIL`：存在**真实内容缺失或篡改**（段落/表格/公式丢失、文本错误、图片大片缺失）。

> 注意：判定为 PASS/WARN 但探针报了 `defects`，必须把这些 defects 逐条写进报告的
> `issues` 字段，不得因为「整体通过」而隐去局部缺陷。

## 输出格式（必须严格遵守）
把结果写成 JSON 文件，路径由派发时指定（如 `review_parts/<组名>.json`）：

```json
{
  "root": "<组名>",
  "docs": [
    {
      "name": "<文档名>.docx",
      "docx": "<绝对路径>",
      "md": "<绝对路径>",
      "status": "PASS|WARN|FAIL",
      "readability_score": 92.0,
      "readability_grade": "A",
      "readability_verdict": "一句话结论",
      "structure": {"headings": 148, "tables_gfm": 120, "tables_html": 69,
                    "images": 39, "formulas": 3, "chars": 171518},
      "consistency": {"paragraph_pct": 100.0, "table_pct": 100.0,
                      "paragraph_evidence": "抽样位置与比对结果的具体描述",
                      "table_evidence": "抽查了哪几张表、核对了哪些格、结论",
                      "image_evidence": "Word 内嵌 N 张 / md 引用 N 张 / 缺失 0"},
      "issues": [
        {"severity": "warn",
         "type": "heading_level_jump|table_column_mismatch|formula_broken|missing_image|emphasis_fragment|textbox_excluded|other",
         "location": "章节名或行号或表序号或段落索引",
         "detail": "具体现象",
         "cause": "成因判断（转换器/原文档/校验口径）",
         "impact": "对 AI 阅读的实际影响",
         "suggestion": "修复建议"}
      ],
      "conclusion": "2~3 句总结：该文档是否可放心交给 AI 阅读，残余风险是什么"
    }
  ],
  "summary": {
    "total": 2, "pass": 2, "warn": 0, "fail": 0,
    "avg_readability": 92.0,
    "notes": "本组整体结论"
  }
}
```

## 硬性要求
1. **必须实际读取每份 docx 与 md**，不得凭文件名或探针分数直接下结论。
2. 每份文档至少提供 3 条具体证据（段落抽样位置、表格核对、图片核对）。
3. `issues` 里每条都要写清 `location`（能定位到章节/表号/行号），不接受「某处」。
4. 不要修改任何源文件或转换结果，只写你自己的 JSON 报告。
5. 最后用 5 行以内文字总结本组结论，作为你的返回内容（JSON 已落盘，无需在正文重复全量）。