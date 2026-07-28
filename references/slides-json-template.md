# slides.json 黄金模板 —— AI 生成指南

## 概述

本文件是 `gen_pptx.py` 的输入规范。AI 在生成 `slides.json` 时必须严格遵循此模板，确保 PPT 输出效果稳定一致。

`slides.json` 是一个 JSON 数组，数组中每个对象描述一页幻灯片。`gen_pptx.py` 会读取该数组，根据每页的 `page_type` 调用对应的构建函数渲染 PPT。请确保：

- 整个文件是合法的 JSON（无双引号缺失、无尾逗号、无注释）。
- 数组元素按页码顺序排列。
- 每页的 `page_type` 必须是下文规定的取值之一。

---

## 必需字段说明

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| page_num | int | 是 | 页码，从1开始 |
| page_type | string | 是 | 页面类型：cover/toc/section/content/figure/conclusion/qa |
| title | string | 是 | 页面标题 |
| subtitle | string | cover可选 | 副标题（封面页的作者期刊信息） |
| sections | array | toc必需 | 目录章节列表，字符串数组 |
| bullets | array | content/figure可选 | 要点列表，字符串数组 |
| sub_title | string | content可选 | 小标题（显示在标题下方，带蓝色竖线装饰） |
| analysis_title | string | figure可选 | 图表分析小标题（等同sub_title） |
| conclusion | string | content/figure可选 | 结论文字（显示在底部浅蓝色结论框中） |
| key_message | string | conclusion/qa可选 | 核心信息（结论页大字展示） |
| highlights | array | content可选 | 高亮信息，字典数组[{title, content}] |
| image_path | string | figure可选 | 图片文件路径 |
| image_caption | string | figure可选 | 图片说明文字 |
| notes | string | 全部可选 | 演讲者备注 |
| kicker | string | 全部可选 | 标题栏右上角标签（如"遗传学验证"） |
| section_no | int | section可选 | 章节编号（不填则自动分配） |

---

## 各页面类型完整示例

### cover 页示例

```json
{
  "page_num": 1,
  "page_type": "cover",
  "title": "论文标题",
  "subtitle": "作者, 期刊, 年份",
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

## 生成规则（重要！）

1. 每页必须有 `page_num`、`page_type`、`title` 三个字段。
2. content 和 figure 页强烈建议填写 `sub_title`、`conclusion` 字段（直接影响PPT美观度）。
3. bullets 中可包含学术关键词（如"增加""减少""促进""抑制"），脚本会自动标红加粗。
4. figure 页的 `image_path` 必须是图片文件的绝对路径。
5. sections 数组可以是字符串数组，也可以是字典数组 `[{title, en, desc}]`。
6. 如果不填 `conclusion`，脚本会自动从 highlights 或 key_message 中提取。
7. 总页数建议14-17页，遵循IMRaD结构。

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
