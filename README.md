# paper-report-ppt

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![TRAE Skill](https://img.shields.io/badge/TRAE-Skill-green.svg)]()

> 把研究生文献（PDF）一键转成**可编辑**的组会汇报 PPTX。文献配图**原样嵌入**，支持用户模板，同步生成完整演讲稿。

---

## 效果预览

以下是用真实文献生成的组会汇报 PPT 截图：

| 封面页 | 结果页（实验图原样嵌入） | 工作模型页 |
|:------:|:------:|:------:|
| [![封面页](assets/screenshots/slide_01.png)](assets/screenshots/slide_01.png) | [![结果页](assets/screenshots/slide_07.png)](assets/screenshots/slide_07.png) | [![工作模型](assets/screenshots/slide_13.png)](assets/screenshots/slide_13.png) |

**17 页完整 PPT，7 张实验图原样嵌入，三重质检全部通过。** [查看完整案例](#真实案例)

---

## 快速开始

```
用户：把这篇论文做成组会汇报 PPT
AI  ：收到。请上传 PDF 文献，并告诉我汇报场景（组会/开题/答辩）。

用户：[上传 PDF] 组会汇报，15 页左右，用我提供的模板
AI  ：已读取文献。以下是汇报大纲（S2），请确认：
      P01 封面 → P02 目录 → P03 文献选择依据 → ... → P17 Q&A
      配图：Figure 1-7 原样嵌入
      是否同步生成演讲稿？

用户：确认，生成演讲稿
AI  ：正在生成 PPTX + 演讲稿 DOCX ... 已完成！
```

**3 分钟后你得到**：
- `report.pptx` — 17 页可编辑 PPT，文献配图原样保留
- `report_speech.docx` — 完整演讲稿（含开场白/逐页讲解/时长分配/问题预判）

---

## 功能特性

| 特性 | 说明 |
|------|------|
| **文献脉络还原** | 按 IMRaD 结构（引言-方法-结果-讨论）自动组织汇报内容 |
| **配图原样保留** | 实验图、数据图从 PDF 原样提取，SHA256 校验，不做裁剪 |
| **PPT 模板支持** | 可基于用户提供的 PPTX 模板生成统一风格演示文稿 |
| **同步演讲稿** | 生成完整口头演讲文字稿（DOCX），含开场白、逐页讲解、时长分配、问题预判 |
| **完全可编辑** | 输出为原生 PPTX（DrawingML 对象），文本/形状/图片均可修改 |
| **质检报告** | 自动生成可编辑性 + 图片完整性 + 脉络一致性三重核验报告 |

---

## 适用场景

- 研究生组会文献汇报（Journal Club）
- 开题/中期/答辩演示
- 学术会议快速准备

---

## 真实案例

以下是用本 skill 处理一篇真实植物病理学文献的完整测试结果。

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
| **可编辑性** | ✅ PASS | 523 个原生对象（516 文本框 + 176 自选图形 + 7 图片），零整页图片 |
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
```

---

## 工作流程（S0–S5）

```
S0 意图确认    → 收集 PDF、场景、页数、语言、模板、侧重、演讲稿需求
     ↓
S1 文献解析    → 提取 Markdown + 原样配图 + image_manifest.json
     ↓
S2 脉络大纲    → 生成 IMRaD 结构 outline.md，用户确认
     ↓
S3 模板准备    → Free design 或用户模板 → workspace
     ↓
S4 PPTX 生成   → SVG 逐页手写 → 质量检查 → 导出 PPTX + 演讲稿 DOCX
     ↓
S5 交付质检    → 可编辑性 + 图片 SHA256 + 脉络一致性核验
```

---

## 依赖

- [ppt-master](https://github.com/trae-ai/ppt-master) skill（PPTX 生成引擎）
- Python 3.x（Windows 下自动回退到 `python`）
- Node.js + `docx` npm 包（演讲稿 DOCX 生成）

---

## 安装

将本仓库克隆到 TRAE skills 目录：

```bash
cd ~/.trae/skills
git clone https://github.com/mlj-1212/paper-report-ppt.git
```

重启 TRAE 后即可使用。

---

## 输出产物

| 产物 | 格式 | 说明 |
|------|------|------|
| 可编辑 PPTX | `.pptx` | 主交付物，12–18 页原生可编辑幻灯片 |
| 演讲稿 | `.docx` | 完整口头演讲文字稿，3000–6000 字 |
| 质检报告 | `.json` | 可编辑性 / 图片完整性 / 脉络一致性 |
| 组会大纲 | `.md` | 脉络文档，含每页对应文献章节 |
| 文献解析 | `.md` + 配图 | 素材留档，含 image_manifest.json |

---

## FAQ

**Q：PDF 中的矢量图（流程图/图表）能保留吗？**
A：位图原样嵌入；矢量图默认不提取，可在 S0 选择栅格化（180 DPI）保留。

**Q：PPT 模板有什么限制？**
A：raw PPTX 需先走 Create Template 流程生成 workspace，不能直接作为生成源。模板只影响视觉风格，不影响汇报结构。

**Q：演讲稿和 Speaker Notes 有什么区别？**
A：Speaker Notes 是 PPT 备注栏的简短提示（每页 50–100 字）；演讲稿是独立 DOCX 完整文稿（3000–6000 字），含开场白、过渡语、结束语和问题预判。

**Q：支持英文文献吗？**
A：支持。语言默认跟随文献语言，英文文献生成英文汇报。

**Q：可以一次处理多篇文献吗？**
A：当前版本支持单篇文献。多文献综述模式在优化方向中规划。

---

## 优化方向

- [ ] 多文献综述模式（2–5 篇文献对比）
- [ ] 配图智能筛选与排序（基于 caption 语义）
- [ ] 公式保真（LaTeX 渲染为 PNG 嵌入）
- [ ] 组会脉络模板预设（IMRaD / 问题驱动 / 创新点驱动 / 综述对比）
- [ ] 增量更新（文献更新后只重新生成变化页面）

---

## License

[MIT](LICENSE) © 2026 mlj-1212
