# 更新日志

本文件记录本项目所有值得注意的变更。

格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## 关于版本号

`1.0.0` 已于 2026-11-25 发布到 [PyPI](https://pypi.org/project/word-vane/1.0.0/)。

开发期间版本号未随功能递增（7 个功能性提交都停留在 1.0.0），
因此下方条目按**日期**而非版本号分组，首次发布后的变更将按语义化版本递增。

当前版本可通过任一方式核对，三者应一致：

```bash
word2md --version          # word2md 1.0.0
python -c "import word2md; print(word2md.__version__)"   # 1.0.0
pip show word-vane | grep Version                        # Version: 1.0.0
```

版本号的**唯一来源**是 `word2md.py` 中的 `__version__`：
`pyproject.toml` 通过 `[tool.setuptools.dynamic]` 读取它，不单独维护。

---

## [1.0.0] — 2026-11-25

首次发布版本，已上架 [PyPI](https://pypi.org/project/word-vane/1.0.0/)。
安装名 `word-vane`，导入名与命令行为 `word2md`。

### 新增

**转换能力**
- 标题识别：`Heading 1-6`、`Title`、中文「标题 1-5」，以及自定义章节样式
- 表格双通道输出：无合并走 GFM 管道表；含合并单元格或嵌套表走 HTML
  `<table>` + `colspan`/`rowspan`，保留合并语义（GFM 语法不支持合并）
- 图片提取：自动落入 `images/`；EMF/WMF 矢量图按 3× LANCZOS 光栅化为 PNG
- 公式转换：OMML → LaTeX（行内 `$...$` / 块级 `$$...$$`），移植自 eqword2llm(MIT)
- 目录域解析：Word 目录域转为可点击锚点，并**在原文位置**输出
- 列表还原：按 Word 编号定义重建有序/无序与层级缩进，含样式级编号
- 代码块启发式识别：C 前缀、赋值、控制语句、函数调用、自增自减、8051 汇编
- 拼音注音处理：只保留正文，丢弃注音部分

**自定义样式映射**
- 通用结构规则，与前缀无关：`<前缀><数字>-<层级>`（如 `chapter1-2`）、
  `<前缀><数字>-<名字>`（如 `chapter1-content`）、`<前缀><数字>`（如 `section1`）
- 外部配置注入：`--styles` 路径 / 环境变量 `WORD2MD_STYLES` /
  `~/.word2md/styles.json`，支持 `heading_styles`、`toc_styles`、
  `subheading_patterns`

**校验与诊断**
- 内容级一致性校验：段落逐块、表格逐格、公式占位、图片引用四层比对
- AI 可读性诊断：输出结构指标、缺陷清单与 0-100 评分
- 转义感知切分：GFM 行内转义竖线 `\|` 不再被误计为列分隔

**质量门禁**
- `word2md --verify`：转换后立即校验，一条命令拿到结论
- `--fail-under`：可读性分数低于阈值时退出码为 1
- `word2md-batch-verify --fail-under`：批量场景，任一份低于阈值即失败
- 所有门禁失败时退出码统一为 `1`，可直接接入 CI

**集成与工程**
- MCP Server：单一工具 `word2md_convert_and_verify`，一次调用完成
  转换 → 校验 → 评分；兼容 mcp SDK 1.x(`FastMCP`) 与 2.x(`MCPServer`)
- 六个命令行入口：`word2md`、`word2md-verify`、`word2md-batch`、
  `word2md-batch-verify`、`word2md-probe`、`word2md-mcp`
- 测试套件：64 个测试函数 / 78 条用例（参数化展开），夹具运行时生成，
  仓库不含任何 `.docx` 二进制
- CI：GitHub Actions，lint + 测试矩阵（Ubuntu/Windows × Python 3.9/3.11/3.13）
  + wheel 构建与安装验证
- 可复现基准：`bench/merged_cells_bench.py` 量化合并单元格保真度，
  退出码可作门禁

### 修复

以下为开发期发现并修复的缺陷，按发现顺序记录：

- **缩进代码行漏判**：`_looks_like_code` 未剥离前导空白，`    int n = 0;`
  的 `startswith` 与 `^` 锚点全部失配，导致单个代码块被切成多段
- **代码块缩进丢失**：入队使用已 `strip()` 的文本，4 空格缩进被压平
- **目录样式漏判**：默认目录样式只收录 `toc 1`/`TOC1`，漏掉 Word 实际使用的
  `TOC 1`（大写带空格），致使整块目录退化为正文
- **`.doc` 提示为死代码**：`endswith(('.docx','.doc'))` 的外层条件对 `.doc`
  恒为 False，内层「请先另存为 .docx」的提示永不可达，用户实际收到底层
  `PackageNotFoundError`
- **`run` 内制表符与软回车丢失**：只收集 `w:t` 导致 `w:tab`/`w:cr`/`w:br`
  被丢弃，`int\ta = 1;` 变成 `inta = 1;`（标识符被篡改）
- **C 指针被误伤**：代码块曾事后正则剔除 `*`，会破坏 `char *p`；改为走
  plain 纯文本通道，不再事后处理
- **段落标记格式污染**：段落标记的 `b`/`i`/`strike` 曾被当作文字格式继承，
  凭空产出 `<s>`/`<strong>`
- **表格伪差异**：单元格内转义竖线 `\|` 被当成列分隔，导致整表比对失败

### 变更

- 分发名使用 `word-vane`。PyPI 上的 `word2md` 属于另一个项目
  （[hnrobert/word2md](https://github.com/hnrobert/word2md)），
  两者无关，请勿混装
- 源码中不再硬编码任何特定模板的专有样式名，全部改为「通用规则 + 外部配置」
- 依赖收敛为两个：`python-docx`（解析 .docx）、`Pillow`（图片光栅化）。
  `mcp` 为可选依赖，仅 `pip install "word-vane[mcp]"` 时安装

### 已知限制

如实记录，避免误用：

| 场景 | 现状 |
|---|---|
| 脚注 / 尾注 | **不导出**。有此需求请用 Pandoc 或 docx2md |
| 批注 / 修订痕迹 | 不导出批注；修订按当前显示版本处理 |
| 嵌入对象（Excel 图表、Visio、OLE） | 不解析 OLE 二进制，跳过 |
| 宏（VBA） | `.docx` 不含宏；`.docm` 转出后宏丢失 |
| 旧版 `.doc` | 不直接支持，会给出明确转换指引 |
| 文本框 / 艺术字 | 图形内文字不在正文流，按设计不抽取 |
| 扫描件 / 纯图片 | 无文本层，需先 OCR |
| 大文档性能 | 大量 EMF/WMF 时约 20–40 秒（矢量图光栅化开销），可用 `--no-images` 跳过 |

### 备注

- 本项目在 `1.0.0` 阶段**尚无外部用户**，API 仍可能调整。若你的场景依赖
  具体函数签名，建议锁定提交哈希而非版本号
- 「转换后校验」并非本项目首创：[`ucli-tools/docx2md`](https://github.com/ucli-tools/ms2md)
  的 `docx2md verify` 早已提供同类能力
- 合并单元格对照数据为**构造夹具上的单点对照**，不代表整体优劣；
  竞品在格式覆盖与速度上可能显著优于本项目

---

## 后续计划（未实现）

以下为讨论中提出但尚未落地的能力，不构成承诺：

- 脚注 / 尾注导出
- 批注导出开关
- 文本框内容抽取
- `--fast` 性能档（跳过矢量图光栅化）