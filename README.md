# paper-report-ppt

TRAE Skill：将研究生文献（PDF）按文章脉络生成**可编辑** PPTX 组会汇报演示文稿。文献配图**原样嵌入不改**，支持用户提供 PPT 模板，同步生成完整演讲稿（DOCX）。

## 功能特性

- **文献脉络还原**：按 IMRaD 结构（引言-方法-结果-讨论）组织汇报内容
- **配图原样保留**：文献中的实验图、数据图原样提取，不做裁剪或修改
- **PPT 模板支持**：可基于用户提供的 PPTX 模板生成统一风格的演示文稿
- **同步演讲稿**：生成完整口头演讲文字稿（DOCX），含开场白、逐页讲解、时长分配、问题预判
- **完全可编辑**：输出为原生 PPTX（DrawingML 对象），非整页图片

## 适用场景

- 研究生组会文献汇报
- 开题/中期/答辩演示
- 期刊俱乐部（Journal Club）

## 触发方式

在 TRAE 中说出以下任意指令即可调用：

- "把这篇论文做成组会汇报 PPT"
- "这篇文献帮我整理成汇报 slides"
- "研究生组会汇报，按文章脉络来"

## 依赖

- [ppt-master](https://github.com/trae-ai/ppt-master) skill（PPTX 生成引擎）
- Python 3.x（Windows 下自动回退到 `python`）
- Node.js + `docx` 包（演讲稿 DOCX 生成）

## 工作流程（S0–S5）

| 阶段 | 内容 | 产物 |
|------|------|------|
| S0 | 意图确认与需求收集 | 需求摘要 |
| S1 | 文献解析与素材提取 | Markdown + 配图 + image_manifest.json |
| S2 | 组会汇报脉络大纲生成 | outline.md（IMRaD 结构） |
| S3 | 模板决策与准备 | 模板 workspace（可选） |
| S4 | PPTX 生成与演讲稿生成 | PPTX + DOCX 演讲稿 |
| S5 | 交付质检 | 可编辑性 + 图片原样 + 脉络核验报告 |

## 输出产物

| 产物 | 格式 | 说明 |
|------|------|------|
| 可编辑 PPTX | `.pptx` | 主交付物，含 12–18 页幻灯片 |
| 演讲稿 | `.docx` | 完整口头演讲文字稿，3000–6000 字 |
| 质检报告 | `.json` | 可编辑性 / 图片完整性 / 脉络一致性 |
| 组会大纲 | `.md` | 脉络文档 |
| 文献解析 | `.md` + 配图 | 素材留档 |

## 安装

将本仓库克隆到 TRAE skills 目录：

```bash
cd ~/.trae/skills
git clone https://github.com/mlj-1212/paper-report-ppt.git
```

## License

MIT License — 详见 [LICENSE](LICENSE)
