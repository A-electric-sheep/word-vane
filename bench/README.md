# 基准（bench）

本目录存放**可复现**的能力基准。原则：不靠形容词，靠一条命令跑出数字。

## merged_cells_bench.py — 合并单元格保真度

### 要回答的问题

GFM 管道表在语法上**没有** colspan/rowspan。因此所有"docx → GFM"的工具遇到合并
单元格，只能在三条退路里选：

1. **重复值** —— 把 `Q1 2026` 复制进被合并的每一格（Pandoc 的做法）
2. **留空** —— 只第一格留值，其余全空（多数工具的做法）
3. **丢表** —— 直接跳过这张表

三种做法都让读者**无法判断原表存在几处合并、跨几行几列**。本仓库选择第四条路：
输出 HTML `<table>`，用原生 `colspan`/`rowspan` 保留结构。

### 运行

```bash
# 仅测 word2md
python bench/merged_cells_bench.py

# 与 anydoc 对照（需先 pip install firecrawl-anydoc）
python bench/merged_cells_bench.py --with-anydoc

# 输出 JSON / 保留临时目录查看生成物
python bench/merged_cells_bench.py --json result.json --keep
```

退出码：`0` = 通过，`1` = 合并语义有丢失（可直接用于 CI 门禁）。

### 判定口径

**主判定**：`verify_consistency` 的单元格级一致率 ≥ 99%，且输出使用 HTML 表格。
之所以用校验器而非自己数属性，是因为校验器直接读 Word XML 与 Markdown 表格流
对位比较，不受"我怎么构造夹具"影响。

**辅助指标**：输出中 `colspan`/`rowspan` 属性个数 vs 夹具真实合并数。
真值由 python-docx 读回 docx 后统计（`gridSpan>1` 计横向，`vMerge=restart` 计纵向），
不是硬编码，夹具改动后指标自动跟随。

### 实测结果（2026-11-04）

夹具：5 张表，真实合并 4 处横向 + 3 处纵向。

| 工具 | colspan 属性 | rowspan 属性 | 输出形态 |
|---|---|---|---|
| **word2md** | **4/4** | **3/3** | HTML `<table>` + colspan/rowspan |
| anydoc | 0/4 | 0/3 | GFM 管道表（合并语义不可恢复） |

word2md 单元格级一致率 **100%**（比对 44 个单元格）。

### 两个已知的实现坑（写在这里省得后来人再踩）

1. **不能用 `row.cells` 统计合并数**。python-docx 对合并区域内每一个网格位都返回
   同一个 `_tc`，会把同一处合并重复计数——实测把 4 处横向合并数成 10 处，保留率
   被误判成 53%。必须直接遍历 `w:tc` 元素。
2. **夹具小节标题不要含 `colspan`/`rowspan` 字样**。标题会作为普通段落写进
   Markdown，全文正则会把标题误计为属性——实测使指标从 100% 虚高到 125%。

### 局限

- 夹具为**构造样本**，未覆盖真实文档中的跨页表、极深嵌套表、单元格内嵌
  多段落等形态；
- 未做交叉合并（同一格既横跨又纵跨）：Word 会把这类格拆成多个 `gridSpan`
  起始格，属性计数随之上浮，使辅助指标失真。这类形态建议看校验器的单元格级
  一致率，而非属性计数。
- 对照项仅 anydoc；如需对比 Pandoc / MarkItDown，可按同样方式接入
  `run_anydoc` 的写法扩展。