# Formula Rendering Guidelines

## Overview

This document describes how `render_formula.py` handles LaTeX math formulas found in research papers.

## Detection Rules

`parse_pdf.py` scans the extracted Markdown for LaTeX math markers:
- Inline math: `$...$`, `\(...\)`
- Display math: `$$...$$`, `\[...\]`

## Rendering Strategy

| Complexity | Action | Output |
|---|---|---|
| Simple (≤15 chars, no fractions/summation/integrals) | Keep as editable text in PPT | Plain text in text frame |
| Complex (>15 chars or contains \frac, \sum, \int, etc.) | Render to PNG via matplotlib | 300 DPI PNG with transparent background |

## PNG Specifications

- **Resolution**: 300 DPI
- **Background**: Transparent (alpha channel)
- **Auto-crop**: White margins removed
- **Format**: PNG

## Formula List JSON Schema

```json
[
  {
    "id": "f001",
    "latex": "E = mc^2",
    "source_page": 3,
    "complexity": "simple",
    "render_as": "text"
  },
  {
    "id": "f002",
    "latex": "\\int_{0}^{\\infty} e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}",
    "source_page": 5,
    "complexity": "complex",
    "render_as": "png"
  }
]
```