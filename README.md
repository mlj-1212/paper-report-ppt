# paper-report-ppt

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![TRAE Skill](https://img.shields.io/badge/TRAE-Skill-green.svg)]()
[![Python 3.8+](https://img.shields.io/badge/Python-3.8+-blue.svg)]()

> 上传一篇文献 PDF，自动生成**可编辑**的组会汇报 PPT。论文中的实验图表原样保留，同步生成配套的演讲稿 Word 文档。**自包含架构，无外部 skill 依赖，所有 AI 桌面应用通用。**

---

## 效果预览

以下是用真实文献生成的组会汇报 PPT 截图：

| 封面页 | 结果页（实验图原样嵌入） | 工作模型页 |
|:------:|:------:|:------:|
| [![封面页](assets/screenshots/slide_01.png)](assets/screenshots/slide_01.png) | [![结果页](assets/screenshots/slide_07.png)](assets/screenshots/slide_07.png) | [![工作模型](assets/screenshots/slide_13.png)](assets/screenshots/slide_13.png) |

**17 页完整 PPT，7 张实验图原样嵌入，三重质检全部通过。** [查看完整案例](#真实案例)

---

## 30 秒开始

```bash
git clone https://github.com/mlj-1212/paper-report-ppt.git
cd paper-report-ppt
python install.py
```

`install.py` 会自动安装全部依赖并完成环境自检，看到 `🎉 安装完成！` 即可使用。

> **AI 环境用户**：也可以直接告诉 AI「安装 https://github.com/mlj-1212/paper-report-ppt 这个 skill」，AI 会自动完成克隆和安装。
>
> **仓库体积仅 ~220 KB**（不含截图） — 完全自包含，无外部 skill 依赖，无 Node.js 依赖。

---

## 用法：两步生成 PPT

### 第 1 步：上传 PDF

把你的文献 PDF 发给 AI，说一句"生成组会汇报 PPT"就行。AI 会自动处理语言识别、页数、结构、图表嵌入、演讲稿生成等所有细节。

### 第 2 步：确认大纲，等待生成

AI 会先给你一个汇报大纲（类似目录），你回复"确认"后，AI 开始生成 PPT。处理完成后，你会收到：

- **一份 PPTX 文件** — 可编辑的组会汇报 PPT，论文图表原样保留
- **一份 DOCX 文件** — 完整的演讲稿，约 3000–6000 字

---

## 进阶用法

如果你想自定义汇报风格，可以在上传 PDF 时附带说明：

| 你想做什么 | 告诉 AI |
|---|---|
| 侧重创新点 | "侧重展示创新点和核心发现，控制在 10 页以内" |
| 英文汇报 | "用英文生成汇报" |
| 开题/答辩场景 | "用于开题汇报，约 20 页，重点展示研究背景和方法对比" |
| 切换视觉主题 | "使用 ref 主题风格（深蓝顶部条 + 中英文对照封面）" |

---

## 功能特性

| 特性 | 说明 |
|------|------|
| **脉络模板预设** | 4 种汇报模板：IMRaD 均衡 / 问题驱动 / 创新点驱动 / 综述对比，一键切换 |
| **配图智能筛选** | 基于 caption 语义自动筛选和排序配图，过滤装饰图，sha256 去重 |
| **公式保真渲染** | 检测文献中 LaTeX 数学公式，复杂公式渲染为高清 PNG 嵌入 PPT，简单公式保留可编辑文本 |
| **文献脉络还原** | 按引言-方法-结果-讨论的结构自动组织汇报内容 |
| **配图原样保留** | 实验图、数据图从 PDF 原样提取，SHA256 校验，不做裁剪 |
| **多视觉主题** | 内置 academic / minimal / trae / ref 四种主题，自由设计路径生成专业学术风格 |
| **同步演讲稿** | 生成完整口头演讲文字稿（DOCX），含开场白、逐页讲解、时长分配、问题预判 |
| **完全可编辑** | 输出为原生 PPTX，文本、形状、图片均可修改 |
| **自动质检** | 可编辑性 + 图片完整性 + 脉络一致性三重核验 |
| **跨平台通用** | 纯 Python 实现，任何有文件系统的 AI 环境均可使用 |

---

## 适用场景

- 研究生组会文献汇报（Journal Club）
- 开题 / 中期 / 答辩演示
- 学术会议快速准备

---

## 真实案例

以下是用本工具处理一篇真实植物病理学文献的完整测试结果。

### 文献信息

| 项目 | 内容 |
|------|------|
| 论文 | *RSV infection induces upregulation of deubiquitinase UBP16 to stabilize SHMT1 for promoting viral infection* |
| 期刊 | **Stress Biology**, 2025, 5:62 |
| 作者 | Wang et al. |
| PDF | 10 页，含 7 张配图（Figure 1–7） |
| 汇报语言 | 中文 |

### 生成结果

| 指标 | 数值 |
|------|------|
| 总页数 | **17 页**（IMRaD 均衡模式） |
| 配图嵌入 | 7/7 原样嵌入（SHA256 全匹配） |
| PPTX 大小 | 1.15 MB |
| 演讲稿字数 | ~3,500 字（约 20 分钟） |

### 幻灯片结构

```
P01  封面          — 标题、作者、期刊、汇报人
P02  目录          — 六大板块导航
P03  文献选择依据   — 为什么选这篇
P04  研究背景      — RSV 与植物泛素化防御
P05  科学问题      — 核心假说与研究策略
P06  方法总览      — 多层次验证体系
P07  系统筛选      — NbUBP16 响应 RSV 感染 [Figure 1]
P08  蛋白互作      — NbUBP16.1 与 NbSHMT1 互作 [Figure 2]
P09  酶活验证      — 去泛素化稳定 NbSHMT1 [Figure 3]
P10  功能分析      — OE-NbSHMT1 抑制防御促进感染 [Figure 4]
P11  遗传验证      — NbUBP16.1 正向调控 RSV 感染 [Figure 5]
P12  ROS 检测      — NbUBP16.1 抑制 ROS 积累 [Figure 6]
P13  工作模型      — 病毒劫持宿主去泛素化通路 [Figure 7]
P14  讨论与创新点
P15  局限性与展望
P16  结论
P17  Q&A / 致谢
```

### 质检报告

| 检查项 | 结果 | 详情 |
|--------|------|------|
| **可编辑性** | ✅ PASS | 原生对象（文本框 + 自选图形 + 图片），零整页图片 |
| **图片完整性** | ✅ PASS | 7/7 媒体文件 SHA256 与源文件完全匹配，图片原样未改 |
| **脉络一致性** | ✅ PASS | 17 页幻灯片与大纲一一对应，IMRaD 4/4 板块全部覆盖 |

### 输出文件

| 文件 | 格式 | 大小 |
|------|------|------|
| RSV_UBP16_组会汇报.pptx | PPTX | 1.15 MB |
| RSV_UBP16_组会汇报演讲稿.docx | DOCX | 14.6 KB |

### 演讲稿内容摘录

> 各位老师、同学们好！今天我汇报的文献是 2025 年发表在 Stress Biology 上的研究，题目是"RSV 感染诱导去泛素化酶 UBP16 上调，稳定 SHMT1 抑制 ROS 积累促进病毒感染"。这项研究首次揭示了植物去泛素化酶调控病毒感染的分子机制，提出了"病毒劫持宿主去泛素化通路"的新范式……

演讲稿包含：开场白 → 五个部分逐页讲解 → 结束语 → 时长分配表 → 问题预判（3 个 Q&A）。

---

## 平台支持

| 平台 | 支持状态 | 说明 |
|------|---------|------|
| **TRAE** | ✅ 完全兼容 | 所有脚本均可在 TRAE 中运行 |
| **WorkBuddy** | ✅ 完全兼容 | 仅需 pip install 5 个包 |
| **Claude Code** | ✅ 完全兼容 | 克隆到 skills 目录即可 |
| **Cursor** | ✅ 完全兼容 | 同上 |
| **Qwen** | ✅ 完全兼容 | 同上 |
| **普通 Chatbot** | ❌ 不支持 | 需要文件系统支持 |

> **所有平台使用同一套脚本，同一套流程**，无路径分支，无环境差异。

---

## 工作流程

```
S0  确认需求
     ↓ 你告诉 AI：页数、语言、是否需要演讲稿
S1  解析文献
     ↓ parse_pdf.py 读取 PDF，提取文字内容和实验图表
S2  生成大纲
     ↓ AI 给出汇报目录，你回复"确认"或调整
     ↓ filter_images.py 筛选配图
S3  生成 PPT
     ↓ AI 生成 slides.json → gen_pptx.py 自动构建 PPTX
     ↓ 配图原样嵌入，speaker notes 写入备注栏
S4  演讲稿生成
     ↓ AI 生成 speech_data.json → gen_speech_docx.py 生成 DOCX
S5  交付质检
     ↓ validate_pptx.py 自动检查：可编辑性 / 图片完整性 / 脉络一致性
```

---

## 依赖

### Python 包

```bash
pip install python-pptx PyMuPDF python-docx Pillow matplotlib
```

| 包 | 用途 | 必需 |
|---|---|---|
| `python-pptx` | PPTX 生成 | ✅ 必需 |
| `PyMuPDF` | PDF 解析 | ✅ 必需 |
| `python-docx` | 演讲稿 DOCX 生成 | ✅ 必需 |
| `Pillow` | 图片处理（尺寸读取/格式转换） | ✅ 必需 |
| `matplotlib` | 公式渲染 | ⚠️ 可选 |

### 系统要求

- Python 3.8+
- 无 Node.js 依赖
- 无外部 skill 依赖

### 脚本清单

| 脚本 | 功能 |
|---|---|
| `install.py` | 一键安装（自动安装依赖 + 环境自检） |
| `scripts/install_check.py` | 环境自检 |
| `scripts/parse_pdf.py` | PDF → 结构化 MD + 配图提取 |
| `scripts/filter_images.py` | 配图筛选/去重/排序 |
| `scripts/render_formula.py` | LaTeX 公式 → PNG |
| `scripts/gen_pptx.py` | slides.json → 可编辑 PPTX |
| `scripts/gen_speech_docx.py` | speech_data.json → 演讲稿 DOCX |
| `scripts/validate_pptx.py` | PPTX 三项质检 |

---

## 输出产物

| 产物 | 格式 | 说明 |
|------|------|------|
| 可编辑 PPTX | `.pptx` | 主交付物，12–18 页原生可编辑幻灯片 |
| 演讲稿 | `.docx` | 完整口头演讲文字稿，3000–6000 字 |
| 质检报告 | 终端输出 / JSON | 可编辑性 / 图片完整性 / 脉络一致性 |
| 组会大纲 | `.md` | 脉络文档，含每页对应文献章节 |
| 文献解析 | `.md` + 配图 | 素材留档 |

---

## FAQ

**Q：如何安装？**
A：克隆仓库后运行 `python install.py` 即可，脚本会自动安装全部依赖并完成自检。

**Q：需要安装 ppt-master 或 Node.js 吗？**
A：都不需要。v4.0 是完全自包含的，所有核心能力通过内置 Python 脚本实现，演讲稿生成使用 python-docx（Python 库）。

**Q：PDF 中的矢量图（流程图/图表）能保留吗？**
A：位图原样嵌入；矢量图默认不提取，可在第一步选择栅格化（180 DPI）保留。

**Q：支持自定义 PPT 模板吗？**
A：v4.0 使用自由设计路径生成 PPT，内置 academic / minimal / trae / ref 四种视觉主题。ref 主题提供深蓝顶部条 + 中英文对照封面的正式学术风格，推荐用于组会汇报。

**Q：演讲稿和 PPT 备注有什么区别？**
A：PPT 备注是每页 50–100 字的简短提示；演讲稿是独立 Word 文档（3000–6000 字），含开场白、过渡语、结束语和问题预判。

**Q：支持英文文献吗？**
A：支持。语言默认跟随文献语言，英文文献生成英文汇报。

**Q：可以一次处理多篇文献吗？**
A：当前版本支持单篇文献。多文献综述模式在优化方向中规划。

---

## 优化方向

- [ ] 多文献综述模式（2–5 篇文献对比）
- [x] 配图智能筛选与排序（基于 caption 语义，含 filter_images.py 脚本）
- [x] 公式保真（LaTeX 渲染为 PNG 嵌入，含 render_formula.py 脚本）
- [x] 组会脉络模板预设（IMRaD / 问题驱动 / 创新点驱动 / 综述对比）
- [x] 演讲稿同步生成（python-docx，含 gen_speech_docx.py 脚本）
- [x] PPTX 自动质检（可编辑性 / 图片完整性 / 脉络一致性，含 validate_pptx.py 脚本）
- [ ] 增量更新（文献更新后只重新生成变化页面）

---

## License

[MIT](LICENSE) © 2026 mlj-1212
