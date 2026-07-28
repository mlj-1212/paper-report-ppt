---
name: paper-report-ppt
description: "Turn research PDFs into editable group-meeting PPTX following the paper's structure; embed figures verbatim; support user templates. Invoke when converting a paper into a group meeting presentation."
---

# Paper Report PPT Skill

把研究生文献（PDF）按文章脉络生成用于组会汇报的**可编辑** PPTX。文献配图**原样嵌入不改**，模板可由使用者提供。

本 skill 是**场景编排层**，提供**双轨生成路径**：
- **路径 A（SVG 管线）**：通过 ppt-master 的 SVG → DrawingML 管线生成，质量最高，适合 TRAE 等深度集成环境
- **路径 B（直接生成）**：通过 `scripts/gen_pptx.py` 用 python-pptx 直接构建 DrawingML，**所有 AI 环境通用**，质量与路径 A 接近

> **重要**：路径 B 是为 WorkBuddy / Claude Code / Cursor / Qwen 等非 TRAE 环境设计的。这些环境的 AI 通常无法可靠地逐页手写符合 ppt-master 严格规范的 SVG 文件，但完全有能力生成 JSON 结构化数据交给脚本处理。**所有 AI 环境优先使用路径 B**，TRAE 环境可选择路径 A。

---

## 协作引擎与路径解析

本 skill 提供双轨生成路径。`ppt-master` 已**内置**在本 skill 的 `vendor/ppt-master/` 目录中，用户无需单独安装。

### `ppt-master` 定位

| 运行环境 | `ppt-master` 状态 | 获取方式 |
|---|---|---|
| 任何有文件系统的 AI 环境 | ✅ 已内置 | 克隆本 skill 即包含 `vendor/ppt-master/`，无需额外操作 |
| TRAE（SOLO CN / Cloud） | 内置 + vendor 双保险 | 优先用 TRAE 内置版本，回退到 vendor 版本 |
| 无文件系统的 Chatbot | 不适用 | 本 skill 不支持纯对话环境 |

### `${PPT_MASTER_DIR}` 路径解析（S0 自动执行）

工作流进入 S0 时按以下优先级解析 `ppt-master` 路径：

1. 环境变量 `PPT_MASTER_DIR`（如果已设置）
2. **`vendor/ppt-master`（内置版本，最高优先级）** — `${PAPER_REPORT_PPT_DIR}/vendor/ppt-master`
3. TRAE 内置：`~/.trae-cn/skills/ppt-master`、`~/.trae/skills/ppt-master`
4. 其他环境：`~/.claude/skills/ppt-master`、`~/.cursor/skills/ppt-master`、`~/.codex/skills/ppt-master`
5. Windows：`%APPDATA%\TRAE SOLO CN\skills\ppt-master`、`%LOCALAPPDATA%\TRAE\skills\ppt-master`

> **关键设计**：`vendor/ppt-master/` 是 `ppt-master` 的完整副本（含 scripts / references / templates / workflows），与本 skill 一起克隆。用户只需克隆一个仓库，无需单独获取 `ppt-master`。

### 其他约定

- **本 skill 目录**：`${PAPER_REPORT_PPT_DIR}` 为当前 SKILL.md 所在目录（即 `paper-report-ppt` skill 根目录）
- **可选协作**：`pdf` skill（用于超长文献的逐页检索证据，通常为内置）
- **运行环境**：Windows 下若 `python3` 不可用，改用 `python`（ppt-master 已声明此规则）。
- **工作目录约定**：
  - 中间素材（MD / pages / outline）放临时工作目录 `<work_dir>`
  - 最终 PPTX 交付到用户工作区

---

## 触发条件

当用户出现以下意图时调用本 skill：
- "把这篇论文做成组会汇报 PPT"
- "这篇文献帮我整理成汇报 slides"
- "研究生组会汇报，按文章脉络来"
- 上传 PDF 并要求生成汇报演示文稿
- "paper report" / "literature presentation" / "journal club slides"

---

## 输入契约

| 输入 | 必需 | 说明 |
|---|---|---|
| 文献 PDF | 是 | 研究生文献（论文 / 预印本 / 学位论文），本地路径 |
| 汇报场景 | 否 | 组会 / 开题 / 中期 / 答辩；默认组会 |
| 目标页数 | 否 | 默认 12–18 页（含封面目录） |
| 语言 | 否 | 默认跟随文献语言；中文文献默认中文汇报 |
| 模板 | 否 | raw PPTX 文件路径；不提供则 free design |
| 侧重 | 否 | IMRaD 均衡 / 问题驱动 / 创新点驱动 / 综述对比；默认均衡 |
| 演讲稿 | 否 | 是否同步生成完整演讲稿（独立 DOCX 文稿）；默认生成 |

---

## 工作流（S0–S5）

### S0 — 意图确认与需求收集

- **输入**：用户对话 + PDF 路径（+ 可选模板 PPTX）
- **输出**：需求摘要（场景 / 页数 / 语言 / 模板 / 侧重 + 矢量图处理决策）+ 已解析的 `${PPT_MASTER_DIR}` + **生成路径决策**
- **脚本**：`scripts/install_check.py`（可选，用于环境自检）
- **门禁**：⛔ ① 环境自检通过 ② 用户确认需求摘要后进入 S1

#### S0.1 环境自检与生成路径决策（首要步骤）

进入 S0 后**第一件事**是运行环境自检：

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/install_check.py --json
```

根据自检结果，**自动选择生成路径**：

| 检测条件 | 生成路径 | 说明 |
|---|---|---|
| TRAE 环境（通过工作目录检测）+ 用户未指定 | **路径 A（SVG 管线）** | TRAE 深度集成 ppt-master，质量最高 |
| TRAE 环境 + 用户明确指定 | 路径 A 或路径 B | 由用户选择 |
| 非 TRAE 环境（WorkBuddy / Claude / Cursor / Qwen 等） | **路径 B（直接生成）** | 这些环境无法可靠执行 SVG 管线 |

> **核心原则**：`install_check.py` 通过**当前工作目录**判断正在运行的 AI 环境（而非机器上装了什么），确保同一台机器上 TRAE 和 WorkBuddy 共存时能正确识别。非 TRAE 环境一律走路径 B。路径 B 通过 `scripts/gen_pptx.py` 用 python-pptx 直接生成 PPTX，不依赖 SVG 手写能力，所有有文件系统的 AI 环境均可使用。

`install_check.py` 行为：
- **任何环境**：检测 Python + python-pptx + PyMuPDF（路径 B 最小依赖）
- **TRAE 环境**：额外检测 ppt-master（路径 A 需要的 SVG 管线脚本）
- **纯对话环境**：提示本 skill 需要文件系统支持，礼貌退出
- 脚本退出码：0=就绪，2=需用户操作，3=环境不支持

#### S0.2 需求收集

**收集项**：
1. PDF 路径
2. 汇报场景与目标页数
3. 是否提供模板 PPTX（⚠️ 路径 B 暂不支持用户模板，自动回退 free design）
4. 汇报语言
5. 脉络侧重（IMRaD 均衡 / 问题驱动 / 创新点驱动 / 综述对比）
6. PDF 中矢量图处理方式（默认不提取；如需保留接受栅格化或尝试 EMF）
7. 是否同步生成演讲稿（DOCX 格式，默认生成）

> **矢量图决策**：PDF 矢量图（图表 / 流程图）默认不提取。若用户要求保留：方案一栅格化（`source_to_md.py --render-vector-figures`，180 DPI）；方案二尝试导出为 EMF/WMF 嵌入（仅当 PDF 内嵌 EMF 时）。在 S0 与用户确认。

---

### S1 — 文献解析与素材提取

- **输入**：PDF
- **输出**：
  - `<work_dir>/<stem>.md`（结构化 Markdown，含 `<!-- Page N -->` 标记）
  - `<work_dir>/<stem>_files/`（原样提取的配图 PNG/JPG + `image_manifest.json`）
  - `<work_dir>/pages/page_XXXX.txt`（逐页文本，长文献检索用）
  - `<work_dir>/formula_list.json`（公式检测清单，见 `references/formula-rendering.md`）
- **门禁**：✅ `<stem>.md` 与 `image_manifest.json` 存在

**主脚本**（ppt-master 管线，原样提取图片不重采样）：

```bash
python ${PPT_MASTER_DIR}/scripts/source_to_md.py <pdf_path> -o <work_dir>/<stem>.md
```

> `source_to_md.py` 调度 `pdf_to_md.py`，使用 PyMuPDF 提取，**图片原样写 bytes 不重采样**（`f.write(image_data)`），并生成 `<stem>_files/image_manifest.json`（含 bbox / pixel_w / pixel_h / sha256 / source_kind / caption / figure_number）。

**公式检测**（新增）：在 `<stem>.md` 中搜索 LaTeX 数学标记（`$...$` / `$$...$$` / `\(...\)` / `\[...\]`），生成 `formula_list.json`（每条记录：id / latex / source_page / complexity / render_as）。详细规则见 `references/formula-rendering.md`。

**长文献补充检索**（可选，>30 页时启用，用 pdf skill）：

```bash
python ${PDF_SKILL_DIR}/scripts/extract_pages.py <pdf_path> <work_dir>
python ${PDF_SKILL_DIR}/scripts/search_extracted.py <work_dir> --query-file <work_dir>/queries.txt
```

`queries.txt` 按组会脉络预设查询：研究背景 / 问题定义 / 方法 / 实验 / 结果 / 讨论 / 结论 / 局限性 / 创新点。

---

### S2 — 组会汇报脉络大纲生成

- **输入**：`<stem>.md` + `image_manifest.json` + `formula_list.json`
- **输出**：`<work_dir>/outline.md`（组会脉络页序列）
- **脚本**：无（主 agent 基于脉络模板预设生成）
- **门禁**：⛔ 用户确认大纲后进入 S3

#### 脉络模板预设

根据 S0 收集的"侧重"方向，从 `references/outline-templates.md` 选取对应模板，自动套用页序列：

| 侧重方向 | 模板 | 页数范围 | 适用场景 |
|----------|------|:---:|------|
| IMRaD 均衡 | 模板 1 | 14–17 | 常规组会汇报（默认） |
| 问题驱动 | 模板 2 | 10–12 | 紧凑汇报，聚焦"问题→解决" |
| 创新点驱动 | 模板 3 | 13–16 | 开题/中期/答辩，强调创新 |
| 综述对比 | 模板 4 | 12–15 | 文献综述，横向对比 |

> 用户未指定时默认使用模板 1。模板可微调：确认大纲时增减页面或调整配图分配。

#### 配图智能筛选与排序

严格按照 `references/image-selection.md` 的规则对配图进行筛选、排序和去重：

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/filter_images.py <work_dir>/<stem>_files/image_manifest.json --max-per-page 2
```

输出 `image_manifest_filtered.json`（筛选后列表），包含：
- 分类：按 caption 语义判断图片类型，自动过滤装饰性小图/logo/二维码
- 去重：按 sha256 去除重复图片
- 排序：按章节顺序排列（引言→方法→结果→讨论）
- 标记：标记图片所属章节，便于分配到对应页

**执行规则**：
1. **筛选**：带 `Figure N` caption 的核心图优先保留，作者照片/logo/二维码过滤掉
2. **去重**：sha256 完全相同的图片只保留一次
3. **排序**：按 caption 出现的页码顺序，严格对应论文章节
4. **分配**：每页最多 1–2 张配图，方法图→方法页，结果图→结果页，模型图→工作模型页

#### 每页大纲需标注

- 该页对应文献的章节 / 页码区间
- 该页应嵌入的配图文件名（从 `image_manifest.json` 经筛选排序后选取，标注 figure 编号与源页码）
- 该页应嵌入的公式（从 `formula_list.json` 选取，标注公式 ID）
- 关键论点（≤3 条）

---

### S3 — 模板决策与准备

- **输入**：S0 收集的模板信息
- **输出**：模板决策（free design / workspace root 路径）
- **门禁**：✅ 明确 free design 或拿到 workspace root 路径

**分支 A — 无模板（free design，默认）**：直接进入 S4。

**分支 B — 用户提供 raw PPTX**：引导走 ppt-master 的 Create Template 路由生成 workspace：

```bash
python ${PPT_MASTER_DIR}/scripts/pptx_template_import.py "<user_template.pptx>"
```

然后按 `c:\Users\Administrator\.trae-cn\skills\ppt-master\workflows\create-template.md` 完成 Create Brand / Layout / Deck，产出 workspace root（含 `templates/design_spec.md` + SVGs）。把该 workspace root 路径记为 S4 的 Step 3 输入。

**分支 C — 用户已有 workspace root**（含 `templates/design_spec.md`）：直接作为 S4 的 Step 3 输入。

> **硬规则**：raw PPTX **不能**直接作为 Generate PPTX Step 3 的 workspace（ppt-master 硬规则）。必须先走 Create Template 生成 workspace，再回 Generate PPTX Step 3。bare 模板名 / 风格描述不触发模板路径，按 free design 处理。模板只影响视觉风格与版式，不影响组会脉络结构。

---

### S4 — PPTX 生成

S0 的路径决策决定 S4 的执行路线：

- **路径 A（SVG 管线）** → S4-A：ppt-master 完整管线（7 步）
- **路径 B（直接生成）** → S4-B：gen_pptx.py 单脚本生成

---

#### S4-A — 路径 A：ppt-master SVG 管线（TRAE 环境默认）

仅当 S0 决策为路径 A 时执行。

- **输入**：`<stem>.md` + `<stem>_files/` + `outline.md` + 模板决策
- **输出**：`exports/<project>_<timestamp>.pptx` + `validation/<project>_<timestamp>.report.json`
- **门禁**：✅ PPTX 生成 + postflight report `passed`

按 ppt-master `workflows/generate-pptx.md` 依次执行 7 步，本 skill 在关键步骤注入场景适配：

#### Step 1（source_to_md）

S1 已完成，跳过。

#### Step 2（project init + import-sources）

```bash
python ${PPT_MASTER_DIR}/scripts/project_manager.py init <project_name> --format ppt169
python ${PPT_MASTER_DIR}/scripts/project_manager.py import-sources <project_path> <pdf_path> <work_dir>/<stem>.md
```

> `import-sources` 会自动运行 PPTX intake（若源是 PPTX）；对 PDF 源，把 MD 和原 PDF 纳入 `sources/`。

#### Step 3（模板选项）

- free design → 直接进 Step 4
- workspace root → 触发 `apply-template-workspace` runbook（见 `workflows/stages/apply-template-workspace.md`）

#### Step 4（Strategist 阶段） — 场景适配关键

**图片资源契约强制**：所有文献配图行必须标记：
- `Acquire Via: user`
- `Status: Existing`
- `Crop Policy: no-crop`
- `Dimensions` / `Ratio` 从 `analysis/image_analysis.csv` 的 Width / Height / AspectRatio 推导（Prepared-user fast path，见 `references/strategist-image.md` §4）

**大纲来源**：把 S2 的 `outline.md` 作为 §IX 页序列的输入依据。

**确认 UI 三阶段**（stage1 通信契约 / stage2 完整方案 / stage3 生产机制）照常走，写 `design_spec.md` + `spec_lock.md`。

**图片分析**：

```bash
python ${PPT_MASTER_DIR}/scripts/analyze_images.py <project_path>/images
```

#### Step 5（图片获取）

文献配图全为 `user / Existing`，**整步跳过**（无 ai / web / slice 行）。

#### Step 5.5（公式渲染） — 新增

若 `formula_list.json` 中存在复杂公式（`complexity=complex`），在 Executor 生成 SVG 前先渲染公式为 PNG：

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/render_formula.py <work_dir>/formula_list.json <project_path>/images
```

> 公式图片以 `formula_XX.png` 命名输出到 `images/` 目录，策略标记为 `Acquire Via: formula | Status: Existing | Crop Policy: no-crop`。渲染参数：300 DPI，透明背景，自动裁剪白边。详细规则见 `references/formula-rendering.md`。

**渲染后处理**：
- 脚本自动回写 `formula_list.json`，补充 `rendered_path`、`rendered_width`、`rendered_height`
- PNG 文件纳入图片资源清单，与文献配图统一管理
- 渲染完成后，在 SVG 中嵌入公式 PNG（display 公式单独占一行，inline 公式与正文同行）

#### Step 6（Executor 阶段）

- 按 `outline.md` 逐页手写 SVG → `svg_output/`
- 配图用 `<image href="../images/<filename>" preserveAspectRatio="xMidYMid meet"/>`（no-crop 完整显示）
- 首页门禁：

```bash
python ${PPT_MASTER_DIR}/scripts/svg_quality_checker.py <project_path> --stage first-page --json
```

- 终检：

```bash
python ${PPT_MASTER_DIR}/scripts/svg_quality_checker.py <project_path> --stage final --json
```

- speaker notes 生成（按 `references/executor-notes.md`）

#### Step 6.5（演讲稿生成） — 场景适配关键

在 speaker notes 生成完成后，基于 notes + outline + 文献解析 MD 生成一份**完整的口头演讲文字稿**。

**演讲稿与 speaker notes 的区别**：

| 维度 | Speaker Notes | 演讲稿 |
|---|---|---|
| 定位 | PPT 备注栏提示 | 独立完整文稿 |
| 粒度 | 每页几句要点 | 连贯叙事，含开场白/过渡语/结尾 |
| 长度 | 每页 50-100 字 | 全文 3000-6000 字（20 分钟演讲） |
| 用途 | 演讲时瞄一眼 | 逐字练习/留存参考 |

**输入**：
- `notes/total.md`（speaker notes，每页要点）
- `<work_dir>/outline.md`（组会脉络大纲）
- `<work_dir>/<stem>.md`（文献解析全文，用于补充细节）
- `design_spec.md` §IX（页序列与每页核心论点）

**输出**：`<project_path>/exports/<project>_speech.docx`

**演讲稿结构**：

```markdown
# 文献汇报演讲稿：<论文标题>

> 文献：作者, 年份, 期刊, 卷: 页码
> 汇报场景：研究生组会
> 预计时长：约 XX 分钟
> 生成日期：YYYY-MM-DD

---

## 开场白（约 1 分钟）

[完整的开场口语文字，介绍文献来源、选择理由、核心悬念，
引出汇报框架。以"各位老师同学好"类口语开场，
以"下面开始汇报"类过渡语收束]

---

## 第一部分：研究背景（对应 P03-P04）

### P03 — <页面标题>

[该页的完整口头讲解文字，200-400 字。
比 speaker notes 更详细，包含：
- 该页要讲什么（引导听众看 PPT）
- 关键数据的口头表述（"大家注意看这个数字..."）
- 与上一页的过渡逻辑
- 与文献原文的对应关系（"原文在 Results 第三段提到..."）]

### P04 — <页面标题>

[同上]

---

## 第二部分：系统筛选（对应 P05-P06）

[过渡语]
[各页讲解]

---

## 第三部分：UBP15 深入研究（对应 P07-P13）

[过渡语]
[各页讲解，配图页适当加长]

---

## 第四部分：遗传交互（对应 P14）

[过渡语]
[该页讲解，重点讲"反直觉发现"]

---

## 第五部分：结论与展望（对应 P15-P16）

[过渡语]
[结论页讲解]
[展望页讲解]

---

## 结束语（约 30 秒）

[总结核心发现的口头表述，
以"以上就是本次汇报的全部内容，感谢大家的聆听，
欢迎提问和讨论"收束]

---

## 演讲提示

- [时长分配建议]
- [重点强调提示]
- [可能被提问的预判]
```

> 上述结构以 Markdown 示意内容层次，实际输出为 DOCX 文件（通过 docx-js 渲染为标题层级、正文段落、时长表格、页码页脚等原生 Word 元素）。

**生成规则**：

1. **口语化**：用"大家看""请注意""这里关键的一点是"等口语引导词，避免书面语
2. **连贯性**：每页讲解之间用过渡语衔接（"接下来看""基于上面的结果""这就引出了下一个问题"）
3. **数据口头化**：把 PPT 上的数字转化为口头表述（"细胞数减少了大约 40%，也就是从将近 300 降到了 180 左右"）
4. **配图引导**：配图页必须包含"请大家看这张图""图中左侧是..."类引导语
5. **时长标注**：每部分标注预计时长，总时长与汇报场景匹配（组会默认 20 分钟）
6. **文献对应**：关键论点后标注文献位置（"原文 Results 第三段""Figure 5 的数据"）
7. **问题预判**：在演讲提示部分预判 2-3 个可能被提问的问题并给出简要回答思路

**DOCX 生成方法**：使用 `docx` (npm) 库生成 DOCX 文稿，确保 CJK 字体正确渲染。

```bash
# 工作目录需安装 docx 依赖
npm install docx
# 运行生成脚本（脚本内含演讲稿全文内容 + docx-js 排版逻辑）
node <work_dir>/gen_speech_docx.js <project_path>/exports/<project>_speech.docx
```

**DOCX 排版规范**：
- CJK 字体：`font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }`
- 标题层级：Heading 1（文档标题）、Heading 2（部分标题）、Heading 3（页面标题）
- 正文段落：size 24（12pt），行距 360（1.5 倍）
- 时长分配表：三列表格（部分 / 页面 / 预计时长），表头浅绿底色 `EDF5E8`
- 页脚：居中页码（`第 N 页`）
- 页面尺寸：A4（11906 × 16838 DXA），1 英寸边距

**生成时机**：Step 6.4（notes 生成）完成后立即生成，与 Step 7（后处理导出）并行无依赖，可先生成演讲稿再导出 PPTX，也可并行。

**门禁**：✅ `exports/<project>_speech.docx` 存在且包含开场白 + 全部页面讲解 + 结束语

#### Step 7（后处理导出）

```bash
python ${PPT_MASTER_DIR}/scripts/total_md_split.py <project_path>
python ${PPT_MASTER_DIR}/scripts/finalize_svg.py <project_path>
python ${PPT_MASTER_DIR}/scripts/svg_to_pptx.py <project_path>
```

---

#### S4-B — 路径 B：gen_pptx.py 直接生成（所有非 TRAE 环境默认）

仅当 S0 决策为路径 B 时执行。**这是 WorkBuddy / Claude Code / Cursor / Qwen 等环境的推荐路径。**

- **输入**：S2 的 `outline.md` + `image_manifest_filtered.json` + 配图文件
- **输出**：`<work_dir>/output.pptx`（可编辑 PPTX）
- **门禁**：✅ PPTX 文件存在且可打开
- **前置依赖**：`pip install python-pptx`（若未安装）

##### Step B1 — 生成 slides.json

**AI 的工作**：根据 S2 的 `outline.md`，逐页生成结构化 JSON 数据，写入 `<work_dir>/slides.json`。

这是路径 B 中 **AI 唯一需要做的创造性工作**——把 outline 中每页的文字内容、配图分配、speaker notes 组织为 JSON 格式。不需要手写 SVG，不需要理解 DrawingML 规范。

**slides.json 格式**：

```json
[
  {
    "page_num": 1,
    "page_type": "cover",
    "title": "文献精读汇报：RSV感染诱导去泛素化酶UBP16上调稳定SHMT1促进病毒感染",
    "subtitle": "Wang et al., Stress Biology, 2025",
    "bullets": [],
    "image_path": null,
    "image_caption": null,
    "notes": "今天汇报的文献是2025年发表在Stress Biology上的研究..."
  },
  {
    "page_num": 2,
    "page_type": "toc",
    "title": "汇报提纲",
    "sections": ["研究背景", "科学问题", "方法总览", "主要结果", "讨论与创新", "结论与展望"],
    "bullets": [],
    "image_path": null,
    "image_caption": null,
    "notes": "本次汇报分为六个部分..."
  },
  {
    "page_num": 3,
    "page_type": "content",
    "title": "研究背景：RSV与植物防御",
    "bullets": [
      "水稻条纹病毒（RSV）是最具破坏性的水稻病毒之一",
      "植物泛素化-去泛素化通路在抗病毒防御中发挥关键作用",
      "去泛素化酶（DUBs）移除泛素链，稳定靶蛋白",
      "已有研究表明病毒可劫持宿主泛素化通路促进感染"
    ],
    "image_path": null,
    "image_caption": null,
    "highlights": [
      {"title": "知识缺口", "content": "植物去泛素化酶是否以及如何调控病毒感染尚不清楚"}
    ],
    "notes": "首先介绍研究背景..."
  },
  {
    "page_num": 7,
    "page_type": "figure",
    "title": "系统筛选：NbUBP16响应RSV感染",
    "bullets": [],
    "image_path": "../s44154-025-00265-2_files/88bb1143-ef92-4902-9283-5c3da02d7305_b4e6069d-486c-4603-95e2-b0db7e55bffd_s44154-025-00265-2_p3_0.jpeg",
    "image_caption": "Figure 1: NbUBP16 responds to RSV infection.",
    "notes": "请看这张图，Figure 1展示了通过转录组分析筛选出的去泛素化酶..."
  },
  {
    "page_num": 17,
    "page_type": "qa",
    "title": "感谢聆听",
    "key_message": "Q&A / 欢迎提问",
    "bullets": [],
    "image_path": null,
    "image_caption": null,
    "notes": "以上就是本次汇报的全部内容..."
  }
]
```

**page_type 取值**：

| page_type | 用途 | 必需字段 | 可选字段 |
|---|---|---|---|
| `cover` | 封面页 | `title`, `subtitle` | `notes` |
| `toc` | 目录页 | `title`, `sections`(数组) | `notes` |
| `section` | 章节分隔页 | `title` | `notes` |
| `content` | 内容页 | `title`, `bullets`(数组) | `highlights`, `notes` |
| `figure` | 配图页 | `title`, `image_path` | `image_caption`, `bullets`, `notes` |
| `model` | 工作模型页 | `title`, `image_path` | `image_caption`, `notes` |
| `conclusion` | 结论页 | `title`, `key_message` | `bullets`, `notes` |
| `qa` | 致谢页 | `title` | `key_message`, `notes` |

**AI 生成 slides.json 的要点**：
1. 从 `outline.md` 提取每页的标题和关键论点，转化为 `bullets`
2. 从 `image_manifest_filtered.json` 获取配图路径，填入 `image_path`（相对路径或绝对路径均可）
3. 为每页生成 50-100 字的 `notes`（speaker notes），供演讲时参考
4. `image_path` 必须是实际存在的文件路径（从 `<stem>_files/` 中提取的原始图片）
5. JSON 必须合法（无注释、无尾逗号、字符串正确转义）

##### Step B2 — 执行 gen_pptx.py 生成 PPTX

```bash
pip install python-pptx   # 若未安装
python ${PAPER_REPORT_PPT_DIR}/scripts/gen_pptx.py \
  --input <work_dir>/slides.json \
  --images-dir <work_dir>/<stem>_files/ \
  --output <work_dir>/output.pptx \
  --theme academic
```

**参数说明**：
- `--input`：Step B1 生成的 slides.json 路径
- `--images-dir`：配图文件所在目录（`<stem>_files/`）
- `--output`：输出 PPTX 路径
- `--theme`：视觉主题，`academic`（默认，专业学术风）或 `minimal`（简洁风）

**脚本行为**：
- 读取 slides.json，逐页生成 python-pptx 原生 DrawingML 对象
- 每页包含：顶部装饰条 + 标题栏 + 内容区 + 底部页码 + 角落装饰形状（15-35 个形状/页）
- 配图以 `preserveAspectRatio` 方式完整嵌入，不裁剪
- speaker notes 写入每页备注栏
- 总输出约 300-500 个原生可编辑对象（接近路径 A 的 466 个）

##### Step B3 — 演讲稿生成（与 S4-A Step 6.5 相同）

路径 B 的演讲稿生成与路径 A 完全一致，使用 docx-js 生成独立 DOCX 文稿。参考 S4-A Step 6.5 的详细规范。

---

### S5 — 交付与质检

- **输入**：PPTX + report.json
- **输出**：交付清单 + 三项核验报告
- **脚本**：无（主 agent 核验，可用 python-pptx 抽检）

**核验项**：

1. **可编辑性核验**：用 python-pptx 抽检文本框 / 形状 / 图片是否为原生对象（非整页图片）
2. **图片原样核验**：比对 PPTX 内 media 文件 sha256 与 `image_manifest.json` 的 source_sha256
3. **组会脉络核验**：幻灯片标题序列与 `outline.md` 一致

---

## 关键约束

### 1. 文献配图原样不改（四层保证）

| 层 | 保证机制 | 实现位置 |
|---|---|---|
| 提取层 | `pdf_to_md.py` 直接写 image bytes，`f.write(image_data)`，不重采样 | `source_to_md/pdf_to_md.py` |
| 契约层 | 所有文献配图强制 `Acquire Via: user \| Status: Existing \| Crop Policy: no-crop` | S4 Step 4 Strategist 注入 |
| 嵌入层 | Executor 用 `preserveAspectRatio="xMidYMid meet"`（完整显示，不裁剪） | S4 Step 6 + `references/svg-image-embedding.md` |
| 导出层 | `svg_to_pptx.py` 映射为原生 pic，media 文件保持原始字节 | S4 Step 7.3 |

**矢量图例外**（须告知用户）：
- PDF 中的矢量图（图表 / 流程图）默认不提取（`pdf_to_md.py` 默认不开 `--render-vector-figures`）
- 若用户要求保留：方案一接受栅格化（`--render-vector-figures`，180 DPI）；方案二尝试导出为 EMF/WMF 嵌入（仅当 PDF 内嵌 EMF 时）
- 决策在 S0 与用户确认

### 2. 可编辑 PPTX 保证

**路径 A**：`svg_to_pptx.py` 把 SVG 映射为 DrawingML 原生对象：文本→text body、形状→autoshape、图片→pic，**非整页图片插入**。

**路径 B**：`gen_pptx.py` 用 python-pptx 直接构建 DrawingML 原生对象：文本→text frame、形状→autoshape、图片→picture，**非整页图片插入**。每页 15-35 个原生形状，总计 300-500 个。

两条路径均保证：所有文本框可改文字、形状可改属性、图片可替换。导出后用 python-pptx 抽检验证。

### 3. 模板支持边界

- raw PPTX **不能**直接作为 Step 3 workspace（ppt-master 硬规则）
- 必须先走 Create Template 生成 workspace，再回 Generate PPTX Step 3
- 用户给 bare 模板名 / 风格描述 → 不触发模板路径，按 free design 处理
- 模板只影响视觉风格与版式，不影响组会脉络结构（脉络由 S2 的 outline 决定）

### 4. 文献配图选取规则

- 优先选取带 `Figure N` caption 的配图（image_manifest.json 的 source_kind=pdf_image）
- 每页最多 1–2 张配图，避免信息过载
- 配图在 speaker notes 中标注源页码与 figure 编号
- 装饰性小图 / logo（被 `should_keep_image` 过滤的）不纳入

---

## 输出契约

| 产物 | 路径 | 说明 |
|---|---|---|
| 可编辑 PPTX | `<project_path>/exports/<project>_<timestamp>.pptx` | 主交付物 |
| 演讲稿 | `<project_path>/exports/<project>_speech.docx` | 完整口头演讲文字稿（DOCX），可独立阅读 |
| 质检报告 | `<project_path>/validation/<project>_<timestamp>.report.json` | postflight |
| 组会大纲 | `<work_dir>/outline.md` | 脉络文档 |
| 文献解析 | `<work_dir>/<stem>.md` + `<stem>_files/` | 素材留档 |
| 设计规约 | `<project_path>/design_spec.md` + `spec_lock.md` | ppt-master 规约 |

---

## 关键设计决策

1. **为何用 Generate PPTX 路由而非 Fill Native PPTX**：Fill 路由 v1 不支持替换图片（"Replace images: Not in v1"），而文献汇报的核心需求之一是嵌入配图；Generate 路由支持 `user/Existing/no-crop` 原样图片嵌入，且产出可编辑 PPTX，完全满足需求。

2. **为何不重写 PDF 解析**：ppt-master 自带的 `source_to_md.py → pdf_to_md.py` 已实现"原样提取图片 bytes + 生成 image_manifest.json + 结构化 MD"，这正是组会汇报所需的素材形态；重写会破坏与 ppt-master 管线的契约。pdf skill 的 `extract_pages.py` 仅作为长文献检索的补充手段。

3. **图片契约为何强制 `no-crop`**：文献配图（实验结果图、方法流程图、数据图表）裁剪会丢失标签 / 坐标轴 / 证据信息，`strategist-image.md` §4 明确指出 "screenshots, charts, dense diagrams" 应标记 `no-crop`；`no-crop` + `preserveAspectRatio="xMidYMid meet"` 保证完整显示不裁剪。

4. **模板链路为何要绕一圈**：ppt-master 硬规则禁止 raw PPTX 直接作为 Step 3 workspace，必须先 Create Template 生成 workspace 再回 Generate PPTX。这是为了保证 Master/Layout 结构的合法性与可复用性，不能绕过。

5. **S2 大纲为何要用户确认**：组会汇报的脉络选择（方法优先 vs 创新点优先、是否含相关工作页）强依赖汇报人意图，且直接影响 §IX 页序列，属于 ppt-master Strategist 阶段之前的场景级决策，应在进入重量级确认流程前锁定。

---

## 优化方向

1. **多文献综述模式**：支持输入 2–5 篇文献，生成对比矩阵页（方法对比表、结果对比表），脉络切换为综述式（背景→分类→对比→趋势→展望）。

2. **配图智能筛选与排序**：✅ 已实现。基于 figure caption 语义和正文引用位置，自动筛选核心配图、按章节排序、sha256 去重。详细规则见 `references/image-selection.md`。

3. **演讲稿与 speaker notes 双层生成**：✅ 已实现。speaker notes 作为 PPT 备注栏简短提示，演讲稿作为独立 DOCX 完整口头文字稿，两者同步生成。

4. **公式保真**：✅ 已实现。检测文献中 LaTeX 数学标记，复杂公式渲染为高清 PNG（300 DPI，透明背景）嵌入 PPTX，简单公式保留为可编辑文本。详细规则见 `references/formula-rendering.md`。

5. **双栏论文阅读顺序适配**：检测双栏排版（PDF bbox 横向分布），分栏提取并按正确阅读顺序合并文本，避免左右栏文字交错。

6. **组会脉络模板预设**：✅ 已实现。预置 4 种脉络模板（IMRaD 均衡 / 问题驱动 / 创新点驱动 / 综述对比），用户在 S0 一键选择，S2 自动套用对应页序列。详细规则见 `references/outline-templates.md`。

7. **增量更新**：文献更新（新版本 PDF）后，按 image_manifest.json 的 source_sha256 比对，只重新生成配图变化的页与受影响章节页，而非全量重做。

8. **双轨生成路径**：✅ 已实现。路径 A（SVG 管线）适合 TRAE 等深度集成环境，路径 B（gen_pptx.py 直接生成）适合所有 AI 环境。非 TRAE 环境自动走路径 B，保证跨平台一致的高质量输出。
