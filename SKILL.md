---
name: paper-report-ppt
description: "Turn research PDFs into editable group-meeting PPTX following the paper's structure; embed figures verbatim. Invoke when converting a paper into a group meeting presentation."
---

# Paper Report PPT Skill

把研究生文献（PDF）按文章脉络生成用于组会汇报的**可编辑** PPTX。文献配图**原样嵌入不改**，使用自由设计路径生成学术专业风格演示文稿。

本 skill 是**自包含场景编排层**，所有核心能力通过内置 Python 脚本实现，不依赖任何外部 skill。只需 `pip install python-pptx PyMuPDF python-docx Pillow matplotlib` 即可在任何 AI 环境中使用。

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
| 侧重 | 否 | IMRaD 均衡 / 问题驱动 / 创新点驱动 / 综述对比；默认均衡 |
| 演讲稿 | 否 | 是否同步生成完整演讲稿（独立 DOCX 文稿）；默认生成 |

---

## 工作流（S0–S5）

### S0 — 意图确认与需求收集

- **输入**：用户对话 + PDF 路径
- **输出**：需求摘要（场景 / 页数 / 语言 / 侧重 + 矢量图处理决策）
- **脚本**：`scripts/install_check.py`
- **门禁**：⛔ ① 环境自检通过 ② 用户确认需求摘要后进入 S1

#### S0.1 环境自检（首要步骤）

进入 S0 后**第一件事**是运行环境自检：

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/install_check.py --json
```

自检内容：
- Python >= 3.8
- 3 个必需 pip 包：`python-pptx`、`PyMuPDF`、`python-docx`
- 1 个可选 pip 包：`matplotlib`（公式渲染）

退出码：0=就绪，2=需安装依赖。若退出码为 2，提示用户运行：
```bash
pip install python-pptx PyMuPDF python-docx matplotlib
```

#### S0.2 需求收集

**收集项**：
1. PDF 路径
2. 汇报场景与目标页数
3. 汇报语言
4. 脉络侧重（IMRaD 均衡 / 问题驱动 / 创新点驱动 / 综述对比）
5. PDF 中矢量图处理方式（默认不提取；如需保留接受栅格化）
6. 是否同步生成演讲稿（DOCX 格式，默认生成）

> **矢量图决策**：PDF 矢量图（图表 / 流程图）默认不提取。若用户要求保留：栅格化（`parse_pdf.py --render-vector-figures`，180 DPI）。在 S0 与用户确认。

---

### S1 — 文献解析与素材提取

- **输入**：PDF
- **输出**：
  - `<work_dir>/<stem>.md`（结构化 Markdown，含 `<!-- Page N -->` 标记）
  - `<work_dir>/<stem>_files/`（原样提取的配图 PNG/JPG + `image_manifest.json`）
  - `<work_dir>/pages/page_XXXX.txt`（逐页文本，长文献检索用）
  - `<work_dir>/formula_list.json`（公式检测清单）
- **门禁**：✅ `<stem>.md` 与 `image_manifest.json` 存在

**PDF 解析**（自包含，仅依赖 PyMuPDF）：

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/parse_pdf.py <pdf_path> -o <work_dir>/<stem>.md
```

> `parse_pdf.py` 是本 skill 自带的 PDF 解析器。功能包括：标题级别检测、加粗/斜体格式化、列表检测、页眉页脚过滤、图片原样提取（不重采样）、image_manifest.json 生成（字段：filename / sha256 / page_number / pixel_width / pixel_height / caption / figure_number）、基础表格检测、矢量图渲染（`--render-vector-figures`）。

**公式检测**：在 `<stem>.md` 中搜索 LaTeX 数学标记（`$...$` / `$$...$$` / `\(...\)` / `\[...\]`），生成 `formula_list.json`（每条记录：id / latex / source_page / complexity / render_as）。详细规则见 `references/formula-rendering.md`。

---

### S2 — 组会汇报脉络大纲生成

- **输入**：`<stem>.md` + `image_manifest.json` + `formula_list.json`
- **输出**：`<work_dir>/outline.md`（组会脉络页序列）
- **脚本**：`scripts/filter_images.py`（配图筛选）
- **门禁**：⛔ 用户确认大纲后进入 S3

#### 脉络模板预设

根据 S0 收集的"侧重"方向，从 `references/outline-templates.md` 选取对应模板：

| 侧重方向 | 模板 | 页数范围 | 适用场景 |
|----------|------|:---:|------|
| IMRaD 均衡 | 模板 1 | 14–17 | 常规组会汇报（默认） |
| 问题驱动 | 模板 2 | 10–12 | 紧凑汇报，聚焦"问题→解决" |
| 创新点驱动 | 模板 3 | 13–16 | 开题/中期/答辩，强调创新 |
| 综述对比 | 模板 4 | 12–15 | 文献综述，横向对比 |

#### 配图智能筛选与排序

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/filter_images.py <work_dir>/<stem>_files/image_manifest.json --max-per-page 2
```

输出 `image_manifest_filtered.json`，包含：
- 分类：按 caption 语义判断图片类型，自动过滤装饰性小图/logo/二维码
- 去重：按 sha256 去除重复图片
- 排序：按章节顺序排列（引言→方法→结果→讨论）
- 标记：标记图片所属章节

详细规则见 `references/image-selection.md`。

#### 每页大纲需标注

- 该页对应文献的章节 / 页码区间
- 该页应嵌入的配图文件名（从 `image_manifest_filtered.json` 选取）
- 该页应嵌入的公式（从 `formula_list.json` 选取）
- 关键论点（≤3 条）

---

### S3 — PPTX 生成（自由设计）

- **输入**：S2 的 `outline.md` + `image_manifest_filtered.json` + 配图文件
- **输出**：`<work_dir>/output.pptx`（可编辑 PPTX）
- **门禁**：✅ PPTX 文件存在且可打开

#### Step 1 — 公式渲染（如有）

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/render_formula.py <work_dir>/formula_list.json <work_dir>/<stem>_files/
```

> 公式渲染为 300 DPI PNG（透明背景，自动裁剪白边）。详细规则见 `references/formula-rendering.md`。

##### Step 2 — 生成 slides.json

**AI 的工作**：根据 `outline.md`，逐页生成结构化 JSON 数据，写入 `<work_dir>/slides.json`。

这是 AI 唯一需要做的创造性工作——把 outline 中每页的文字内容、配图分配、speaker notes 组织为 JSON 格式。不需要手写 SVG，不需要理解 DrawingML 规范。

**slides.json 格式**：

```json
[
  {
    "page_num": 1,
    "page_type": "cover",
    "cn_title": "文献精读汇报：<中文论文标题>",
    "en_title": "<English Paper Title>",
    "subtitle": "作者 et al., 期刊, 年份",
    "presenter": "汇报人：研究生组会汇报",
    "date": "2026年7月29日",
    "notes": "今天汇报的文献是..."
  },
  {
    "page_num": 2,
    "page_type": "toc",
    "title": "汇报提纲",
    "sections": ["研究背景", "科学问题", "方法总览", "主要结果", "讨论与创新", "结论与展望"],
    "notes": "本次汇报分为六个部分..."
  },
  {
    "page_num": 3,
    "page_type": "content",
    "title": "研究背景",
    "bullets": ["要点1", "要点2", "要点3"],
    "highlights": [{"title": "知识缺口", "content": "..."}],
    "notes": "首先介绍研究背景..."
  },
  {
    "page_num": 7,
    "page_type": "figure",
    "title": "系统筛选结果",
    "image_path": "<stem>_files/figure_1.png",
    "image_caption": "Figure 1: ...",
    "bullets": [],
    "notes": "请看这张图..."
  },
  {
    "page_num": 17,
    "page_type": "qa",
    "title": "感谢聆听",
    "key_message": "Q&A / 欢迎提问",
    "notes": "以上就是本次汇报的全部内容..."
  }
]
```

**page_type 取值**：

| page_type | 用途 | 必需字段 | 可选字段 |
|---|---|---|---|
| `cover` | 封面页 | `cn_title`或`title`, `subtitle` | `en_title`, `presenter`, `date`, `notes` |
| `toc` | 目录页 | `title`, `sections`(数组) | `notes` |
| `section` | 章节分隔页 | `title` | `notes` |
| `content` | 内容页 | `title`, `bullets`(数组) | `highlights`, `notes` |
| `figure` | 配图页 | `title`, `image_path` | `image_caption`, `bullets`, `notes` |
| `model` | 工作模型页 | `title`, `image_path` | `image_caption`, `notes` |
| `conclusion` | 结论页 | `title`, `key_message` | `bullets`, `notes` |
| `qa` | 致谢页 | `title` | `key_message`, `notes` |

##### Step 3 — 执行 gen_pptx.py 生成 PPTX

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/gen_pptx.py \
  --input <work_dir>/slides.json \
  --images-dir <work_dir>/<stem>_files/ \
  --output <work_dir>/output.pptx \
  --theme academic
```

**参数说明**：
- `--input`：slides.json 路径
- `--images-dir`：配图文件所在目录
- `--output`：输出 PPTX 路径
- `--theme`：视觉主题（`academic` 默认 / `minimal` / `trae` / `ref`）
  - `ref`：参考模板风格（深蓝顶部条 + 白色直角卡片 + 海军蓝标题 + 中英文对照封面），推荐用于正式汇报

**脚本行为**：
- 读取 slides.json，逐页生成 python-pptx 原生 DrawingML 对象
- 每页包含：顶部装饰条 + 标题栏 + 内容区 + 底部页码
- 配图以 `preserveAspectRatio` 方式完整嵌入，不裁剪
- speaker notes 写入每页备注栏
- 总输出约 300-500 个原生可编辑对象

---

### S4 — 演讲稿生成

- **输入**：slides.json + outline.md + `<stem>.md`
- **输出**：`<work_dir>/output_speech.docx`（完整口头演讲文字稿）
- **脚本**：`scripts/gen_speech_docx.py`
- **门禁**：✅ DOCX 文件存在且包含开场白 + 全部页面讲解 + 结束语

**演讲稿与 speaker notes 的区别**：

| 维度 | Speaker Notes | 演讲稿 |
|---|---|---|
| 定位 | PPT 备注栏提示 | 独立完整文稿 |
| 粒度 | 每页几句要点 | 连贯叙事，含开场白/过渡语/结尾 |
| 长度 | 每页 50-100 字 | 全文 3000–6000 字（20 分钟演讲） |
| 用途 | 演讲时瞄一眼 | 逐字练习/留存参考 |

**Step 1 — AI 生成 speech_data.json**

AI 基于 slides.json 的 notes + outline.md + 文献解析 MD，生成结构化演讲稿内容：

```json
{
  "title": "文献汇报演讲稿：<论文标题>",
  "meta": {
    "literature": "作者, 年份, 期刊",
    "scenario": "研究生组会",
    "duration_minutes": 20,
    "date": "2026-07-29"
  },
  "opening": "各位老师同学好，今天汇报的文献是...",
  "sections": [
    {
      "part_title": "第一部分：研究背景",
      "pages": [
        {
          "page_num": "P03",
          "page_title": "研究背景",
          "duration_minutes": 1.5,
          "content": "该页的完整口头讲解文字，200-400字..."
        }
      ]
    }
  ],
  "closing": "以上就是本次汇报的全部内容，感谢大家的聆听...",
  "duration_table": [
    {"part": "开场白", "pages": "-", "duration": "1分钟"},
    {"part": "第一部分", "pages": "P03-P04", "duration": "3分钟"}
  ],
  "tips": ["时长分配建议", "重点强调提示", "可能被提问的预判"]
}
```

**生成规则**：
1. **口语化**：用"大家看""请注意""这里关键的一点是"等口语引导词
2. **连贯性**：每页讲解之间用过渡语衔接
3. **数据口头化**：把 PPT 上的数字转化为口头表述
4. **配图引导**：配图页必须包含"请大家看这张图"类引导语
5. **时长标注**：每部分标注预计时长，总时长与汇报场景匹配
6. **文献对应**：关键论点后标注文献位置
7. **问题预判**：预判 2-3 个可能被提问的问题

**Step 2 — 执行 gen_speech_docx.py 生成 DOCX**

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/gen_speech_docx.py \
  --input <work_dir>/speech_data.json \
  --output <work_dir>/output_speech.docx \
  --verbose
```

**DOCX 排版规范**：
- CJK 字体：Microsoft YaHei（通过 XML eastAsia 属性设置）
- 标题层级：H1（文档标题）、H2（部分标题）、H3（页面标题）
- 正文段落：12pt，1.5 倍行距
- 时长分配表：三列表格（部分 / 页面 / 预计时长），表头浅绿底色 `#EDF5E8`
- 页脚：居中页码（`第 N 页`）
- 页面尺寸：A4，1 英寸边距

---

### S5 — 交付与质检

- **输入**：PPTX + image_manifest.json + outline.md
- **输出**：交付清单 + 质检报告
- **脚本**：`scripts/validate_pptx.py`
- **门禁**：✅ 三项质检通过

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/validate_pptx.py <work_dir>/output.pptx \
  --manifest <work_dir>/<stem>_files/image_manifest.json \
  --outline <work_dir>/outline.md
```

**三项质检**：

1. **可编辑性核验**：用 python-pptx 检查每页 shape 类型分布（textbox / autoshape / picture 计数），确认非整页图片
2. **图片原样核验**：提取 `ppt/media/` 文件计算 sha256，与 `image_manifest.json` 的 sha256 比对
3. **脉络一致性核验**：提取每页标题文本，与 `outline.md` 页面标题序列比对

输出 JSON 报告，退出码 0=全部通过，1=有未通过项。

---

## 关键约束

### 1. 文献配图原样不改（三层保证）

| 层 | 保证机制 | 实现位置 |
|---|---|---|
| 提取层 | `parse_pdf.py` 直接写 image bytes，不重采样 | `scripts/parse_pdf.py` |
| 嵌入层 | `gen_pptx.py` 用 `preserveAspectRatio` 完整显示，不裁剪 | `scripts/gen_pptx.py` |
| 验证层 | `validate_pptx.py` 比对 sha256 确保图片未被修改 | `scripts/validate_pptx.py` |

### 2. 可编辑 PPTX 保证

`gen_pptx.py` 用 python-pptx 直接构建 DrawingML 原生对象：文本→text frame、形状→autoshape、图片→picture，**非整页图片插入**。每页 15-35 个原生形状，总计 300-500 个。所有文本框可改文字、形状可改属性、图片可替换。

### 3. 文献配图选取规则

- 优先选取带 `Figure N` caption 的配图
- 每页最多 1–2 张配图，避免信息过载
- 装饰性小图 / logo 不纳入
- 详细规则见 `references/image-selection.md`

---

## 输出契约

| 产物 | 路径 | 说明 |
|---|---|---|
| 可编辑 PPTX | `<work_dir>/output.pptx` | 主交付物 |
| 演讲稿 | `<work_dir>/output_speech.docx` | 完整口头演讲文字稿（DOCX） |
| 质检报告 | 终端输出 / `--json` 可导出 | 三项核验结果 |
| 组会大纲 | `<work_dir>/outline.md` | 脉络文档 |
| 文献解析 | `<work_dir>/<stem>.md` + `<stem>_files/` | 素材留档 |

---

## 脚本清单

| 脚本 | 功能 | 依赖 |
|---|---|---|
| `install_check.py` | 环境自检 | 无 |
| `parse_pdf.py` | PDF → 结构化 MD + 配图提取 | PyMuPDF |
| `filter_images.py` | 配图筛选/去重/排序 | 无 |
| `render_formula.py` | LaTeX 公式 → PNG | matplotlib |
| `gen_pptx.py` | slides.json → 可编辑 PPTX | python-pptx |
| `gen_speech_docx.py` | speech_data.json → 演讲稿 DOCX | python-docx |
| `validate_pptx.py` | PPTX 三项质检 | python-pptx |

---

## 跨环境兼容性

本 skill v4.0 为**自包含架构**，在所有有文件系统的 AI 环境中均可使用：

| 环境 | 兼容性 | 说明 |
|---|---|---|
| TRAE (SOLO CN / Cloud) | ✅ 完全兼容 | 所有脚本均可在 TRAE 中运行 |
| Claude Code | ✅ 完全兼容 | 仅需 pip install 5 个包 |
| Cursor | ✅ 完全兼容 | 同上 |
| WorkBuddy | ✅ 完全兼容 | 同上 |
| Qwen | ✅ 完全兼容 | 同上 |
| 纯对话 Chatbot | ❌ 不支持 | 需要文件系统 |

**安装命令**：
```bash
pip install python-pptx PyMuPDF python-docx matplotlib Pillow
```

无 ppt-master 依赖、无 Node.js 依赖、无任何外部 skill 依赖。