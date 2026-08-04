# 配图智能筛选与排序规则

## 概述

基于 `image_manifest.json` 和文献 caption 语义，自动筛选、排序、去重配图。

---

## 筛选规则

### 1. 保留优先级

| 图片类型 | 优先级 | 说明 |
|----------|:------:|------|
| 带 `Figure N` caption 的独立图片 | ⭐⭐⭐⭐⭐ | 优先保留，每个 figure 对应一页或半页 |
| 带 `Fig. N` caption 的图片 | ⭐⭐⭐⭐⭐ | 同上 |
| 带 `Figure S1` 等补充材料图片 | ⭐⭐⭐ | 仅当页数宽松时保留 |
| 正文内嵌的小示意图 | ⭐⭐⭐ | 用于说明机制，保留 1-2 张 |
| 化学结构式 | ⭐⭐⭐ | 保留（如果是关键结构） |
| 作者照片 / 机构 logo / 二维码 | ⭐ | 过滤掉，不纳入 PPT |
| 装饰性插图 | ⭐ | 过滤掉，不纳入 PPT |

### 2. 标题 caption 检测

识别以下形式的 figure caption：

```
Figure 1: ...
Fig. 2. ...
Figure 1A, 1B: ...
Supplementary Figure S1: ...
Fig S1: ...
```

检测到 caption 的图片**必然留下**，因为它们是论文作者明确标注的核心证据。

### 3. 去重规则

- 比较 `image_manifest.json` 中的 `sha256`，完全相同的图片只保留一次
- 同一 figure 的多面板子图，保持原样保留（它们是同一 figure 的不同部分）
- 同一实验重复数据图，只保留结论性的一张

### 4. 容量控制

- 每页最多保留 **1–2 张**配图
- 总配图数量 ≈ 总页数 × 0.7–1.0（15 页 ≈ 10–15 张）
- 如果文献配图超过容量上限：优先保留正文多次引用的 figure，过滤掉低引用次数的补充图

---

## 排序规则

### 1. 按章节顺序排序

- 按照 PDF 中 caption 出现的章节顺序排列
- 引言 → 方法 → 结果 → 讨论，顺序严格对应
- 同一章节内多个 figure 按出现顺序排列

### 2. 按引用位置排序

- 解析正文文本中的 "as shown in Figure 3"
- 正文引用次数越多的 figure 优先级越高
- 如果一张图在正文多次引用，优先放在核心结果页

### 3. 最终排序步骤

```
1. 所有通过筛选的图按 caption 出现顺序分组
2. 每组标记对应章节
3. 章节之间按引言 → 方法 → 结果 → 讨论排序
4. 同一章节内按 caption 出现顺序排序
5. 检查是否偏离论文原有逻辑，微调
```

---

## 分配到页规则

- **结论页**：不需要配图（纯文字总结）
- **方法页**：保留 1 张方法示意图 / 流程图
- **结果页**：每页保留 1 张核心结果图（对应 1-2 个实验）
- **讨论页**：不需要配图（或保留 1 张工作模型图）
- 多图可分配到相邻页，避免一页堆砌多张图

---

## 生成 image_manifest.json 字段说明

`image_manifest.json` 每条记录：

```json
{
  "filename": "figure_1.jpg",
  "sha256": "2b04a59a4c54779ad390edca6706a4bb7c03c0a5924d544672a7130e4ac74179",
  "source_kind": "pdf_image",
  "caption": "NbUBP16 responds to RSV infection.",
  "caption_text": "NbUBP16 responds to RSV infection.",
  "figure_number": 1,
  "page_number": 7,
  "pixel_width": 600,
  "pixel_height": 450,
  "bbox": [x0, y0, x1, y1]
}
```

**语义分析使用 caption_text**：
- 如果 caption 包含 "model"、"mechanism"、"work model" → 放在工作模型页
- 如果 caption 包含 "method"、"protocol"、"procedure" → 放在方法页
- 如果 caption 包含 "result"、"data"、"expression" → 放在结果页

---

## 策略决策树

```
开始
  ↓
读取 image_manifest.json
  ↓
按优先级筛选：
  有 caption → 保留
  是 logo/二维码 → 过滤
  是 decorative → 过滤
  ↓
去重（按 sha256）
  ↓
按 caption 所在页码排序
  ↓
按章节分组（引言/方法/结果/讨论）
  ↓
按每页最多 2 张分配到对应页
  ↓
输出：每页配图列表（filename + caption + 占位）
结束
```