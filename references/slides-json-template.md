# slides.json 黄金模板 —— AI 生成指南

## 概述

本文件是 `gen_pptx.py` 的输入规范。AI 在生成 `slides.json` 时必须严格遵循此模板，确保 PPT 输出效果稳定一致。

`slides.json` 是一个 JSON 数组，数组中每个对象描述一页幻灯片。`gen_pptx.py` 会读取该数组，根据每页的 `page_type` 调用对应的构建函数渲染 PPT。请确保：

- 整个文件是合法的 JSON（无双引号缺失、无尾逗号、无注释）。
- 数组元素按页码顺序排列。
- 每页的 `page_type` 必须是下文规定的取值之一。

> **完整示例**：真实的 17 页 slides.json 示例见同目录下的 `example-slides.json`。AI 生成时应严格模仿其结构和字段填写方式，确保跨环境一致性。

---

## 必需字段说明

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| page_num | int | 是 | 页码，从1开始 |
| page_type | string | 是 | 页面类型：cover/toc/section/content/figure/conclusion/qa |
| title | string | 是* | 页面标题（*cover页可用cn_title替代） |
| cn_title | string | cover可选 | 中文标题（封面页主标题，启用中英文对照） |
| en_title | string | cover可选 | 英文标题（封面页副标题，与cn_title配合使用） |
| subtitle | string | cover可选 | 引文信息（作者, 期刊, 年份），显示在深蓝圆角引文条中 |
| sections | array | toc必需 | 目录章节列表，字符串数组 |
| bullets | array | content/figure可选 | 要点列表，字符串数组 |
| sub_title | string | content可选 | 小标题（显示在标题下方，无竖线装饰） |
| analysis_title | string | figure可选 | 图表分析小标题（等同sub_title） |
| conclusion | string | content/figure可选 | 结论文字（显示在底部结论框中） |
| key_message | string | conclusion/qa可选 | 核心信息（结论页大字展示） |
| highlights | array | content可选 | 高亮信息，字典数组[{title, content}] |
| image_path | string | figure可选 | 图片文件路径 |
| image_caption | string | figure可选 | 图片说明文字 |
| notes | string | 全部可选 | 演讲者备注 |
| kicker | string | 全部可选 | 标题栏右上角标签（如"遗传学验证"） |
| section_no | int | section可选 | 章节编号（不填则自动分配） |
| presenter | string | cover可选 | 汇报人信息（默认"汇报人：研究生组会汇报"） |
| date | string | cover可选 | 汇报日期（默认当天日期） |

---

## 各页面类型完整示例

### cover 页示例

封面页**必须**使用中英文对照格式（cn_title + en_title），**禁止**仅使用 title：

```json
{
  "page_num": 1,
  "page_type": "cover",
  "cn_title": "泛素特异性蛋白酶16通过调控Na+/H+逆向转运活性调节拟南芥的耐盐性",
  "en_title": "UBIQUITIN-SPECIFIC PROTEASE16 Modulates Salt Tolerance in Arabidopsis",
  "subtitle": "Zhou et al., The Plant Cell, 2012, Vol. 24: 5106-5122",
  "presenter": "汇报人：研究生组会汇报",
  "date": "2026年7月29日",
  "notes": "演讲备注"
}
```

### toc 页示例

```json
{
  "page_num": 2,
  "page_type": "toc",
  "title": "汇报提纲",
  "sections": ["研究背景", "科学问题", "方法", "结果", "讨论", "结论"],
  "notes": "演讲备注"
}
```

### section 页示例

```json
{
  "page_num": 3,
  "page_type": "section",
  "title": "研究背景与科学问题",
  "notes": "演讲备注"
}
```

### content 页示例（完整版）

```json
{
  "page_num": 4,
  "page_type": "content",
  "title": "研究背景：RSV与植物防御",
  "sub_title": "前期工作与知识空白",
  "kicker": "背景介绍",
  "bullets": [
    "水稻条纹病毒（RSV）是水稻生产最具破坏性的病毒之一",
    "植物抗病毒免疫主要为两层：PTI与ETI，伴随ROS爆发",
    "前期工作：E3连接酶MEL泛素化并降解SHMT1，激活防御",
    "知识空白：去泛素化酶是否参与调控植物病毒侵染尚不清楚"
  ],
  "conclusion": "去泛素化酶是否参与调控病毒侵染是本文要回答的核心问题",
  "notes": "演讲备注"
}
```

### figure 页示例（完整版）

```json
{
  "page_num": 7,
  "page_type": "figure",
  "title": "图1 NbUBP16响应并调控RSV侵染",
  "sub_title": "RNA-seq筛选与功能验证",
  "kicker": "实验结果",
  "bullets": [
    "RNA-seq：RSV侵染后15个DUBs中仅NbUBP16明显上调（4-6倍）",
    "VIGS敲低NbUBP16后，RSV症状与外壳蛋白积累显著下降"
  ],
  "conclusion": "NbUBP16是RSV侵染的正调控因子",
  "image_path": "图片文件路径",
  "image_caption": "Fig. 1 NbUBP16 responds to and regulates RSV infection.",
  "notes": "演讲备注"
}
```

### conclusion 页示例

```json
{
  "page_num": 16,
  "page_type": "conclusion",
  "title": "结论",
  "key_message": "RSV诱导NbUBP16上调，去泛素化稳定NbSHMT1、抑制ROS，从而促进病毒侵染",
  "bullets": [
    "揭示病毒劫持宿主去泛素化通路的新策略",
    "为抗病靶点设计提供线索"
  ],
  "notes": "演讲备注"
}
```

### qa 页示例

```json
{
  "page_num": 17,
  "page_type": "qa",
  "title": "感谢与讨论",
  "key_message": "Q & A · 欢迎提问",
  "bullets": ["感谢各位老师同学"],
  "notes": "演讲备注"
}
```

---

## 生成规则（必须严格遵守！）

> ⚠️ 以下规则为**硬性约束**，不是建议。违反任一规则将导致不同 AI 环境生成的 PPT 不一致。

1. 每页必须有 `page_num`、`page_type` 字段。封面页**必须**用 `cn_title`（**禁止**仅用 `title`）。
2. 封面页**必须**同时包含 `cn_title` + `en_title` + `subtitle`，实现中英文对照。
3. content 页**必须**填写 `sub_title` 和 `conclusion`；figure 页**必须**填写 `sub_title` 和 `image_caption`。
4. bullets 中可包含学术关键词（如"增加""减少""促进""抑制"），脚本会自动标红加粗。
5. figure 页的 `image_path` **必须**使用 `image_manifest_filtered.json` 中的 `filename` 字段值（仅文件名，不含路径前缀）。
6. toc 页的 `sections` **必须**固定为四段：`["研究背景与科学问题", "材料与方法", "主要结果", "讨论与结论"]`。
7. 如果不填 `conclusion`，脚本会自动从 highlights 或 key_message 中提取（但**建议**显式填写以保证一致性）。
8. 总页数**必须**为 17 页，遵循固定 page_type 序列（见下方"完整17页示例大纲"）。
9. 每个 figure 页**只放一张**配图，**禁止**合并多张图到同一页。
10. bullets 数量：content 页 3-4 条，figure 页 2-3 条，每条不超过 40 个中文字符。

---

## 跨环境一致性校验清单

生成 slides.json 后，AI **必须**逐项自检以下内容，全部通过方可提交：

| # | 检查项 | 通过条件 |
|---|---|---|
| C1 | 总页数 = 17 | `len(slides) == 17` |
| C2 | 封面含 cn_title + en_title | `slides[0]` 同时有这两个字段 |
| C3 | 第 2 页 sections 为四段式 | `slides[1]["sections"]` 长度 = 4 |
| C4 | page_type 序列正确 | cover→toc→section→content→content→section→figure×N→section→content→conclusion→qa |
| C5 | 所有 figure 页有 image_path | 每个 figure 页的 image_path 非空 |
| C6 | image_path 与 manifest 一致 | 每个 image_path 值存在于 image_manifest_filtered.json 的 filename 列表中 |
| C7 | 所有 content/figure 页有 sub_title | 非封面/目录/章节/结论/QA 页均有 sub_title |
| C8 | 无合并图 | 每个 figure 页仅 1 个 image_path |

---

## 完整17页示例大纲

标准17页的 page_type 序列如下：

```
cover(1) → toc(2) → section(3) → content(4) → content(5) →
section(6) → figure(7-13) → section(14) → content(15) →
conclusion(16) → qa(17)
```

对应页面类型分布：

| 页码 | page_type | 说明 |
|------|-----------|------|
| 1 | cover | 封面：论文标题、作者、期刊、年份 |
| 2 | toc | 目录：汇报提纲导航 |
| 3 | section | 章节页：研究背景与科学问题 |
| 4 | content | 研究背景介绍 |
| 5 | content | 科学问题与研究策略 |
| 6 | section | 章节页：实验结果 |
| 7 | figure | 主要结果图1 |
| 8 | figure | 主要结果图2 |
| 9 | figure | 主要结果图3 |
| 10 | figure | 主要结果图4 |
| 11 | figure | 主要结果图5 |
| 12 | figure | 主要结果图6 |
| 13 | figure | 主要结果图7 |
| 14 | section | 章节页：讨论与结论 |
| 15 | content | 讨论与创新点总结 |
| 16 | conclusion | 结论：核心发现一句话总结 |
| 17 | qa | 致谢与讨论 |
