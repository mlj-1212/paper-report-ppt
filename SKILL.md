---
name: paper-report-ppt
description: "Turn research PDFs into editable group-meeting PPTX following the paper's structure; embed figures verbatim; support user templates. Invoke when converting a paper into a group meeting presentation."
---

# Paper Report PPT Skill

把研究生文献（PDF）按文章脉络生成用于组会汇报的**可编辑** PPTX。文献配图**原样嵌入不改**，模板可由使用者提供。

本 skill 是**场景编排层**，复用 `ppt-master` 的 Generate PPTX 路由完成 PPTX 生成，自身只负责文献解析、组会脉络结构化、配图契约适配与模板引导。

---

## 依赖与前置条件

- **必须依赖**：`ppt-master` skill 已安装。脚本根目录：
  `${PPT_MASTER_DIR}` = `c:\Users\Administrator\.trae-cn\skills\ppt-master`
- **可选依赖**：`pdf` skill（用于超长文献的逐页检索证据）。脚本根目录：
  `${PDF_SKILL_DIR}` = `c:\Users\Administrator\.trae-cn\builtin\work\hebe\skills\pdf`
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
- **输出**：需求摘要（场景 / 页数 / 语言 / 模板 / 侧重 + 矢量图处理决策）
- **脚本**：无（对话收集）
- **门禁**：⛔ 用户确认需求摘要后进入 S1

**收集项**：
1. PDF 路径
2. 汇报场景与目标页数
3. 是否提供模板 PPTX
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
- **门禁**：✅ `<stem>.md` 与 `image_manifest.json` 存在

**主脚本**（ppt-master 管线，原样提取图片不重采样）：

```bash
python ${PPT_MASTER_DIR}/scripts/source_to_md.py <pdf_path> -o <work_dir>/<stem>.md
```

> `source_to_md.py` 调度 `pdf_to_md.py`，使用 PyMuPDF 提取，**图片原样写 bytes 不重采样**（`f.write(image_data)`），并生成 `<stem>_files/image_manifest.json`（含 bbox / pixel_w / pixel_h / sha256 / source_kind）。

**长文献补充检索**（可选，>30 页时启用，用 pdf skill）：

```bash
python ${PDF_SKILL_DIR}/scripts/extract_pages.py <pdf_path> <work_dir>
python ${PDF_SKILL_DIR}/scripts/search_extracted.py <work_dir> --query-file <work_dir>/queries.txt
```

`queries.txt` 按组会脉络预设查询：研究背景 / 问题定义 / 方法 / 实验 / 结果 / 讨论 / 结论 / 局限性 / 创新点。

---

### S2 — 组会汇报脉络大纲生成

- **输入**：`<stem>.md` + `image_manifest.json`
- **输出**：`<work_dir>/outline.md`（组会脉络页序列）
- **脚本**：无（主 agent 基于 IMRaD 结构生成）
- **门禁**：⛔ 用户确认大纲后进入 S3

**默认脉络**（IMRaD 适配，可按侧重调整）：

| 页 | 内容 | 必备 |
|---|---|---|
| P01 | 封面（标题 / 作者 / 汇报人 / 日期） | 是 |
| P02 | 目录 | 是 |
| P03 | 文献选择依据（为什么读这篇） | 是 |
| P04 | 研究背景与问题 | 是 |
| P05 | 相关工作（简） | 否 |
| P06 | 方法 / 模型 | 是 |
| P07 | 方法关键图（配图原样） | 是 |
| P08 | 实验设置 | 是 |
| P09 | 主要结果（配图原样） | 是 |
| P10 | 结果分析 / 讨论 | 是 |
| P11 | 创新点总结 | 是 |
| P12 | 局限性与未来工作 | 是 |
| P13 | 结论 | 是 |
| P14 | Q&A / 致谢 | 是 |

**每页大纲需标注**：
- 该页对应文献的章节 / 页码区间
- 该页应嵌入的配图文件名（从 `image_manifest.json` 选取，标注 figure 编号与源页码）
- 关键论点（≤3 条）

**配图选取规则**：
- 优先选取带 `Figure N` caption 的配图
- 每页最多 1–2 张配图，避免信息过载
- 装饰性小图 / logo 不纳入
- 用 `image_manifest.json` 的 sha256 去重

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

### S4 — ppt-master Generate PPTX 编排

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

- `svg_to_pptx.py` 把 SVG 映射为 DrawingML 原生对象：文本→text body、形状→autoshape、图片→pic，**非整页图片插入**
- `design_spec.md` 的 `pptx_structure.mode` 默认 `flat`，保证 Slide-local 可编辑
- 导出后用 python-pptx 抽检：文本框可改文字、形状可改属性、图片可替换

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

2. **配图智能筛选与排序**：基于 figure caption 语义和正文引用位置（"as shown in Figure 3"），自动筛选核心配图并按章节排序；过滤纯装饰图与重复图（image_manifest.json 的 source_sha256 去重）。

3. **演讲稿与 speaker notes 双层生成**：✅ 已实现。speaker notes 作为 PPT 备注栏简短提示（每页 50-100 字），演讲稿作为独立 DOCX 完整口头文字稿（全文 3000-6000 字，含开场白/过渡语/结束语/问题预判，docx-js 排版含标题层级/时长表格/页码），两者同步生成、互为补充。

4. **公式保真**：检测文献中的关键公式（`$...$` / `$$...$$`），通过 ppt-master 的 `latex_render.py` 渲染为 PNG 原样嵌入（mixed 策略：复杂公式渲染、简单 inline 保留可编辑文本），公式行标记 `Acquire Via: formula \| Crop Policy: no-crop`。

5. **双栏论文阅读顺序适配**：检测双栏排版（PDF bbox 横向分布），分栏提取并按正确阅读顺序合并文本，避免左右栏文字交错。

6. **组会脉络模板预设**：预置 4 种脉络模板（IMRaD 均衡 / 问题驱动 / 创新点驱动 / 综述对比），用户在 S0 一键选择，S2 自动套用对应页序列。

7. **增量更新**：文献更新（新版本 PDF）后，按 image_manifest.json 的 source_sha256 比对，只重新生成配图变化的页与受影响章节页，而非全量重做。
