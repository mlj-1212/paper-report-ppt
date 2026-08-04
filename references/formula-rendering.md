# 公式保真渲染规则

## 概述

检测文献中的关键数学公式（LaTeX 风格），通过 LaTeX 渲染为高清 PNG 嵌入 PPTX，保证公式可读性。

---

## 公式检测

### 1. 检测范围

在 S1 文献解析时，从 `<stem>.md` 中检测以下 LaTeX 公式标记：

| 标记形式 | 示例 | 处理方式 |
|----------|------|----------|
| `$...$` (inline) | `$E = mc^2$` | 渲染为 PNG，inline 嵌入 |
| `$$...$$` (display) | `$$\int_0^1 f(x)dx$$` | 渲染为 PNG，单独占一行 |
| `\(...\)` (inline) | `\(x^2 + y^2 = r^2\)` | 渲染为 PNG，inline 嵌入 |
| `\[...\]` (display) | `\[\sum_{i=1}^n x_i\]` | 渲染为 PNG，单独占一行 |

### 2. 检测策略

- 在 `<stem>.md` 中搜索所有 LaTeX 数学标记
- 生成公式清单，每公式记录：页码、上下文、复杂度评估
- 判断规则：
  - **简单公式**（单行、无矩阵、无嵌套分式）：尝试保留为可编辑文本（Unicode 数学符号 + 上下标）
  - **复杂公式**（多行、矩阵、分式、积分、求和）：渲染为高清 PNG

### 3. 公式清单格式

```json
{
  "formulas": [
    {
      "id": "formula_01",
      "latex": "\\frac{d[P]}{dt} = k_{cat}[E]_0\\frac{[S]}{K_m + [S]}",
      "source_page": 5,
      "context": "Michaelis-Menten equation",
      "complexity": "complex",
      "render_as": "png",
      "width": 250,
      "height": 60
    },
    {
      "id": "formula_02",
      "latex": "p < 0.05",
      "source_page": 8,
      "context": "Statistical significance",
      "complexity": "simple",
      "render_as": "text"
    }
  ]
}
```

---

## LaTeX 渲染为 PNG

### 1. 渲染脚本

使用 `matplotlib` + `mathtext` 渲染（不需要完整 LaTeX 环境）：

```python
# render_formula.py
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

def render_latex_to_png(latex_str, output_path, dpi=300):
    """
    将 LaTeX 公式渲染为高清 PNG。
    dpi=300 确保高分辨率，适合 PPT 嵌入。
    """
    fig, ax = plt.subplots(figsize=(4, 0.8))
    ax.axis('off')
    ax.text(0.5, 0.5, f'${latex_str}$',
            fontsize=14,
            horizontalalignment='center',
            verticalalignment='center',
            transform=ax.transAxes)
    fig.savefig(output_path, dpi=dpi, bbox_inches='tight',
                pad_inches=0.1, transparent=True)
    plt.close(fig)
```

### 2. 调用方式

```bash
python ${PAPER_REPORT_PPT_DIR}/scripts/render_formula.py <formula_list.json> <output_dir>
```

输出：`<output_dir>/formula_01.png`, `formula_02.png`, ...

### 3. 渲染参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| dpi | 300 | 输出分辨率，300 DPI 保证 PPT 中清晰显示 |
| fontsize | 14 | 公式字号 |
| bbox_inches | tight | 自动裁剪白边 |
| transparent | true | 透明背景，适合嵌入 PPT |
| pad_inches | 0.1 | 最小内边距 |

---

## 公式嵌入 PPTX

### 1. 图片契约

公式图片在 Strategist 阶段标记为：

```
Acquire Via: formula
Status: Existing
Crop Policy: no-crop
Dimensions: 见 render_formula.py 输出
```

### 2. 嵌入方式

- **Display 公式**（`$$...$$`）：在 SVG 中单独占一行，居中显示
- **Inline 公式**（`$...$`）：在 SVG 中与正文同行，保持基线对齐
- **简单公式**（保留为文本）：在 SVG 中用 `<text>` 元素 + Unicode 数学符号

### 3. SVG 嵌入示例

```xml
<!-- Display 公式（复杂，渲染为 PNG） -->
<image href="../images/formula_01.png"
       x="100" y="400" width="300" height="60"
       preserveAspectRatio="xMidYMid meet"/>

<!-- Inline 公式（简单，保留为文本） -->
<text x="100" y="500" font-size="14" fill="#333">
  Statistical significance: p &lt; 0.05
</text>
```

---

## 公式质量检查

### 1. 渲染检查

- 检查 PNG 是否成功生成，文件大小 > 0
- 检查公式是否清晰可读（无模糊、无截断）
- 检查公式背景是否透明

### 2. 内容检查

- 对比渲染结果与原文公式，确保符号、上下标、括号匹配
- 化学方程式（如 `H₂O → H⁺ + OH⁻`）优先用 Unicode 而非 LaTeX，避免过度渲染

---

## 完整流程

```
S1 文献解析
  ↓
检测 <stem>.md 中的 LaTeX 公式
  ↓
生成 formula_list.json（公式清单 + 复杂度评估）
  ↓
S2 大纲生成
  ↓
标记哪些页需要嵌入公式
  ↓
S4 生成前
  ↓
render_formula.py 渲染公式为 PNG
  ↓
公式图片作为 user/Existing 图片纳入 PPTX
  ↓
S4 Executor 阶段
  ↓
在 SVG 中按 display/inline 方式嵌入公式
  ↓
S5 质检
  ↓
检查公式清晰度 + 内容匹配
```