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
| image_path | string | figure/model必需 | 配图相对路径（相对于images-dir） |
| image_caption | string | figure/model可选 | 图片说明文字 |
| presenter | string | cover可选 | 汇报人信息 |
| date | string | cover可选 | 汇报日期 |
| notes | string | 所有类型可选 | Speaker notes（演讲提示） |

---

## page_type 详细规范

### cover — 封面页

```json
{
  "page_num": 1,
  "page_type": "cover",
  "cn_title": "文献精读汇报：水稻条纹病毒劫持宿主去泛素化通路的分子机制",
  "en_title": "RSV Infection Induces Upregulation of Deubiquitinase UBP16 to Stabilize SHMT1 for Promoting Viral Infection",
  "subtitle": "Wang et al., Stress Biology, 2025",
  "presenter": "汇报人：张三",
  "date": "2026年7月30日",
  "notes": "今天汇报的文献是..."
}
```

### toc — 目录页

```json
{
  "page_num": 2,
  "page_type": "toc",
  "title": "汇报提纲",
  "sections": ["研究背景", "科学问题", "方法总览", "主要结果", "讨论与创新", "结论与展望"],
  "notes": "本次汇报分为六个部分..."
}
```

### section — 章节分隔页

```json
{
  "page_num": 5,
  "page_type": "section",
  "title": "主要结果",
  "notes": "接下来进入结果部分..."
}
```

### content — 内容页

```json
{
  "page_num": 3,
  "page_type": "content",
  "title": "研究背景",
  "sub_title": "RSV与植物泛素化防御",
  "bullets": [
    "水稻条纹病毒（RSV）是水稻最重要的病毒病原之一",
    "植物通过泛素-蛋白酶体系统（UPS）调控抗病毒免疫",
    "去泛素化酶（DUB）在病毒感染中的调控机制尚不清楚"
  ],
  "highlights": [
    {"title": "知识缺口", "content": "UBP16在RSV感染中的功能未知"}
  ],
  "conclusion": "因此，解析UBP16的抗病毒调控机制具有重要意义",
  "notes": "首先介绍研究背景..."
}
```

### figure — 配图页

```json
{
  "page_num": 7,
  "page_type": "figure",
  "title": "系统筛选结果",
  "analysis_title": "NbUBP16响应RSV感染",
  "image_path": "paper_files/figure_1.png",
  "image_caption": "Figure 1: 病毒诱导的UBP16上调验证",
  "bullets": [
    "RSV感染显著诱导NbUBP16转录本上调",
    "Western blot证实蛋白水平同步上升",
    "沉默NbUBP16显著抑制病毒积累"
  ],
  "conclusion": "NbUBP16正向调控RSV感染",
  "notes": "请看这张图..."
}
```

### conclusion — 结论页

```json
{
  "page_num": 16,
  "page_type": "conclusion",
  "title": "结论",
  "key_message": "RSV劫持宿主去泛素化通路稳定SHMT1抑制ROS积累",
  "bullets": [
    "首次揭示植物DUB调控病毒感染的分子机制",
    "发现病毒-宿主互作新范式"
  ],
  "notes": "总结本次汇报的核心结论..."
}
```

### qa — 致谢页

```json
{
  "page_num": 17,
  "page_type": "qa",
  "title": "感谢聆听",
  "key_message": "Q&A / 欢迎提问",
  "notes": "以上就是本次汇报的全部内容..."
}
```

---

## 注意事项

1. **图片路径**：`image_path` 使用相对于 `--images-dir` 的相对路径
2. **JSON 合法性**：确保没有 trailing commas、所有字符串用双引号
3. **页码连续**：page_num 应从 1 开始连续递增
4. **notes 长度**：每页 notes 建议 50-150 字，作为演讲提示
