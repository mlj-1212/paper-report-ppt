#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
parse_pdf.py — paper-report-ppt 自包含 PDF 解析器
====================================================
从 ppt-master 的 pdf_to_md.py 中提取核心能力，使路径 B 完全自包含。

功能：
  1. PDF → Markdown（标题级别检测、加粗/斜体、列表、页眉页脚过滤）
  2. 图片原样提取（不重采样）+ image_manifest.json 生成
  3. 基础表格检测（PyMuPDF 原生表格）
  4. 页码标记（<!-- Page N -->）

依赖：仅需 PyMuPDF (fitz) + Python 标准库

用法:
    python parse_pdf.py <pdf_path> -o <output.md>
    python parse_pdf.py <pdf_path> -o <output.md> --images all
    python parse_pdf.py <pdf_path> -o <output.md> --render-vector-figures

输出:
    <stem>.md              — 结构化 Markdown
    <stem>_files/          — 原样提取的图片
    <stem>_files/image_manifest.json — 图片清单（兼容 filter_images.py）
"""

import argparse
import hashlib
import json
import os
import re
import sys
from pathlib import Path
from collections import Counter

try:
    import fitz  # PyMuPDF
except ImportError:
    print("[ERROR] PyMuPDF not installed. Run: pip install PyMuPDF", file=sys.stderr)
    sys.exit(1)

try:
    from PIL import Image as PILImage, ImageOps
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


# ═══════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════

FONT_BODY_SIZE = 12
FONT_H1_SIZE = 24
FONT_H2_SIZE = 18
FONT_H3_SIZE = 14
HEADER_FOOTER_SAMPLE_LIMIT = 40
HEADER_FOOTER_EDGE_SAMPLE_SIZE = 20
CONTROL_CHARS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

# Image filtering thresholds
MIN_IMAGE_PIXELS = 100
MIN_IMAGE_AREA = 30000
MIN_IMAGE_BYTES = 2048
MIN_PAGE_RATIO = 0.05
MIN_VISIBLE_IMAGE_WIDTH = 40
MIN_VISIBLE_IMAGE_HEIGHT = 40
MIN_VISIBLE_IMAGE_AREA_RATIO = 0.01
MAX_ASPECT_RATIO = 12
MAX_LOW_INFO_BPP = 0.08
MAX_LOW_INFO_AREA = 500000

# Vector figure thresholds
MIN_VECTOR_FIGURE_WIDTH = 100
MIN_VECTOR_FIGURE_HEIGHT = 80
MIN_VECTOR_FIGURE_AREA = 30000
MAX_VECTOR_FIGURE_ASPECT_RATIO = 8
VECTOR_FIGURE_PADDING = 4
VECTOR_FIGURE_DPI = 180

# Figure caption regex
FIGURE_CAPTION_RE = re.compile(
    r'^(?:Figure|Fig\.?)\s*\d+\s*[:.|｜]', re.IGNORECASE
)


# ═══════════════════════════════════════════════════════════
#  字号分析与标题检测
# ═══════════════════════════════════════════════════════════

def analyze_font_sizes(doc: fitz.Document) -> dict:
    """分析字号分布，推断 body/h1/h2/h3。"""
    size_counter = Counter()

    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block["type"] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        size = round(span["size"], 1)
                        text = span["text"].strip()
                        if text:
                            size_counter[size] += len(text)

    if not size_counter:
        return {
            "body": FONT_BODY_SIZE,
            "h1": FONT_H1_SIZE,
            "h2": FONT_H2_SIZE,
            "h3": FONT_H3_SIZE,
        }

    sorted_sizes = sorted(size_counter.items(), key=lambda x: x[1], reverse=True)
    body_size = sorted_sizes[0][0]

    all_sizes = sorted(size_counter.keys(), reverse=True)
    larger_sizes = [s for s in all_sizes if s > body_size + 1]

    size_map = {"body": body_size}
    if len(larger_sizes) >= 1:
        size_map["h1"] = larger_sizes[0]
    if len(larger_sizes) >= 2:
        size_map["h2"] = larger_sizes[1]
    if len(larger_sizes) >= 3:
        size_map["h3"] = larger_sizes[2]

    return size_map


def get_heading_level(size: float, size_map: dict, text: str = "",
                      flags: int = 0, strict: bool = True) -> int:
    """多启发式判断标题级别（字号+加粗+正则）。"""
    level = 0
    if "h1" in size_map and size >= size_map["h1"] - 0.5:
        level = 1
    elif "h2" in size_map and size >= size_map["h2"] - 0.5:
        level = 2
    elif "h3" in size_map and size >= size_map["h3"] - 0.5:
        level = 3

    if level == 0:
        return 0

    if not strict or not text:
        return level

    text = text.strip()

    if len(text) > 80:
        return 0

    sentence_endings = '.。!！?？'
    if text and text[-1] in sentence_endings:
        if not re.match(r'^[\d第]+[.、章节]', text):
            return 0

    is_bold = flags & 16
    if not is_bold and level >= 2:
        body_size = size_map.get("body", 12)
        if size < body_size + 2:
            return 0

    return level


# ═══════════════════════════════════════════════════════════
#  文本格式化
# ═══════════════════════════════════════════════════════════

def is_monospace_font(font_name: str) -> bool:
    """识别等宽字体（代码用）。"""
    if not font_name:
        return False
    font_lower = font_name.lower()
    mono_fonts = [
        'courier', 'consolas', 'monaco', 'menlo', 'monospace',
        'source code', 'fira code', 'jetbrains', 'inconsolata',
        'dejavu sans mono', 'liberation mono', 'ubuntu mono',
        'roboto mono', 'robotomono', 'sf mono', 'cascadia', 'hack'
    ]
    return any(f in font_lower for f in mono_fonts)


def format_span_text(text: str, flags: int) -> str:
    """根据 font flags 输出 **bold** / *italic*。"""
    text = CONTROL_CHARS_RE.sub('', text)
    text = text.strip()
    if not text:
        return ""

    is_bold = flags & 16
    is_italic = flags & 2

    if is_bold and is_italic:
        return f"***{text}***"
    elif is_bold:
        return f"**{text}**"
    elif is_italic:
        return f"*{text}*"
    return text


def detect_list_item(text: str) -> tuple:
    """检测有序/无序列表。"""
    text = text.strip()

    ul_patterns = [
        (r'^[•●○◦▪▸►]\s*', '-'),
        (r'^[-–—]\s+', '-'),
        (r'^\*\s+', '-'),
    ]
    for pattern, marker in ul_patterns:
        match = re.match(pattern, text)
        if match:
            return (True, 'ul', marker + ' ' + text[match.end():])

    ol_pattern = r'^(\d+)[.、)]\s*'
    match = re.match(ol_pattern, text)
    if match:
        num = match.group(1)
        return (True, 'ol', f"{num}. " + text[match.end():])

    return (False, None, text)


def remove_page_footer(text: str) -> str:
    """移除页脚页码模式。"""
    months_en = r'(?:January|February|March|April|May|June|July|August|September|October|November|December)'
    pattern_en = rf'\s*{months_en}\s+\d{{4}}\s+\d{{1,3}}\s*$'
    text = re.sub(pattern_en, '', text, flags=re.IGNORECASE)

    pattern_cn = r'\s*\d{4}年\d{1,2}月\s+\d{1,3}\s*$'
    text = re.sub(pattern_cn, '', text)

    return text.rstrip()


def merge_adjacent_formatting(text: str) -> str:
    """合并相邻同风格格式化碎片。"""
    text = re.sub(r'\*{6}', ' ', text)
    text = re.sub(r'\*{4}', ' ', text)
    return text


def is_sentence_end(text: str) -> bool:
    """检查文本是否以句末标点结尾。"""
    text = text.rstrip()
    if not text:
        return True
    end_puncts = '.。!！?？:：;；'
    return text[-1] in end_puncts


def should_merge_lines(current: dict, next_line: dict) -> bool:
    """判断两行是否应合并为同一段落。"""
    if current.get("is_heading") or next_line.get("is_heading"):
        return False
    if current.get("is_list") or next_line.get("is_list"):
        return False
    if is_sentence_end(current.get("content", "")):
        return False
    return True


# ═══════════════════════════════════════════════════════════
#  页眉页脚检测
# ═══════════════════════════════════════════════════════════

def detect_headers_footers(doc: fitz.Document, threshold_ratio: float = 0.6) -> set:
    """统计检测跨页重复的页眉页脚文本。"""
    if len(doc) < 3:
        return set()

    headers = []
    footers = []

    pages_to_scan = list(range(len(doc)))
    if len(doc) > HEADER_FOOTER_SAMPLE_LIMIT:
        pages_to_scan = (
            pages_to_scan[:HEADER_FOOTER_EDGE_SAMPLE_SIZE]
            + pages_to_scan[-HEADER_FOOTER_EDGE_SAMPLE_SIZE:]
        )

    for i in pages_to_scan:
        page = doc[i]
        rect = page.rect
        h = rect.height

        top_rect = fitz.Rect(0, 0, rect.width, h * 0.15)
        bottom_rect = fitz.Rect(0, h * 0.85, rect.width, h)

        blocks = page.get_text("blocks")
        for b in blocks:
            b_rect = fitz.Rect(b[:4])
            text = b[4].strip()
            if not text:
                continue

            if b_rect.intersects(top_rect):
                headers.append(text)
            elif b_rect.intersects(bottom_rect):
                footers.append(text)

    noise_texts = set()
    total_scanned = len(pages_to_scan)

    for collection in [headers, footers]:
        counter = Counter(collection)
        for text, count in counter.items():
            if count / total_scanned > threshold_ratio:
                noise_texts.add(text)

    return noise_texts


# ═══════════════════════════════════════════════════════════
#  标题合并
# ═══════════════════════════════════════════════════════════

def merge_adjacent_headings(elements: list) -> list:
    """合并相邻同级别短标题。"""
    if not elements:
        return elements

    merged = []
    i = 0

    while i < len(elements):
        el = elements[i]

        if el.get("type") != 0 or not el.get("is_heading"):
            merged.append(el)
            i += 1
            continue

        content = el["content"]
        match = re.match(r'^(#{1,6})\s+(.+)$', content)
        if not match:
            merged.append(el)
            i += 1
            continue

        level = match.group(1)
        title_text = match.group(2)

        j = i + 1
        while j < len(elements) and len(title_text) < 60:
            next_el = elements[j]
            if next_el.get("type") != 0 or not next_el.get("is_heading"):
                break

            next_match = re.match(r'^(#{1,6})\s+(.+)$', next_el["content"])
            if not next_match or next_match.group(1) != level:
                break

            next_text = next_match.group(2)
            if len(next_text) > 40:
                break

            title_text += " " + next_text
            j += 1

        el["content"] = f"{level} {title_text}"
        merged.append(el)
        i = j

    return merged


# ═══════════════════════════════════════════════════════════
#  图片过滤
# ═══════════════════════════════════════════════════════════

def should_keep_image(block: dict, page_rect: fitz.Rect,
                      seen_hashes: set = None) -> bool:
    """过滤小图/装饰图/重复图。"""
    w, h = block.get("width", 0), block.get("height", 0)
    bbox = block.get("bbox", (0, 0, 0, 0))
    render_w = bbox[2] - bbox[0]
    render_h = bbox[3] - bbox[1]
    page_area = page_rect.width * page_rect.height
    render_area_ratio = (render_w * render_h) / page_area if page_area > 0 else 0
    visibly_placed = (
        render_w >= MIN_VISIBLE_IMAGE_WIDTH
        and render_h >= MIN_VISIBLE_IMAGE_HEIGHT
        and render_area_ratio >= MIN_VISIBLE_IMAGE_AREA_RATIO
    )

    if not visibly_placed and (w < MIN_IMAGE_PIXELS or h < MIN_IMAGE_PIXELS):
        return False

    area = w * h
    if not visibly_placed and area < MIN_IMAGE_AREA:
        return False

    image_data = block.get("image", b"")
    if not visibly_placed and len(image_data) < MIN_IMAGE_BYTES:
        return False

    if seen_hashes is not None:
        img_hash = hashlib.md5(image_data).hexdigest()
        if img_hash in seen_hashes and not visibly_placed:
            return False
        seen_hashes.add(img_hash)

    page_w = page_rect.width
    page_h = page_rect.height
    if page_w > 0 and page_h > 0:
        if render_w / page_w < MIN_PAGE_RATIO and render_h / page_h < MIN_PAGE_RATIO:
            return False

    aspect = max(w, h) / max(min(w, h), 1)
    if aspect > MAX_ASPECT_RATIO:
        return False

    bpp = len(image_data) / area if area > 0 else 0
    if bpp < MAX_LOW_INFO_BPP and area < MAX_LOW_INFO_AREA and not visibly_placed:
        return False

    return True


# ═══════════════════════════════════════════════════════════
#  图片反色检测与修复
# ═══════════════════════════════════════════════════════════

def fix_inverted_image(image_data: bytes, ext: str, block: dict = None) -> tuple:
    """检测并修复 PDF 提取图片的反色（黑底）问题。

    根因：PDF 中的 /ImageMask true 模板蒙版图（1-bit、colorspace=CSNone、bpc=1）
    被 PyMuPDF 导出为灰度 PNG 时，墨迹=1(白)、背景=0(黑)，导致插入 PPT 后
    出现"黑底白线"的异常图。

    检测采用分层策略，避免误伤真实暗背景图（如荧光显微图，通常是 RGB 模式）：
    - 第一层（确定性）：block["colorspace"]==0 且 block["bpc"]==1 判定为蒙版，直接反色。
    - 第二层（兜底，仅限 1-bit / 灰度 L 模式）：亮度阈值判定反色。
    - RGB / P / CMYK 模式绝不做亮度反色（防止荧光显微图等被错误反色）。

    返回: (修复后的 image_data, 是否修复)
    """
    if not PIL_AVAILABLE:
        return image_data, False

    try:
        import io
        img = PILImage.open(io.BytesIO(image_data))

        # ── 第一层：基于 PDF 元数据的确定性蒙版检测 ──
        is_mask = False
        if block is not None:
            cs = block.get("colorspace")
            bpc = block.get("bpc")
            # /ImageMask true 图像：colorspace=0(CSNone) 且 bpc=1
            if cs == 0 and bpc == 1:
                is_mask = True

        if is_mask:
            # 蒙版图：无论 PIL 解码为 1-bit 还是灰度 L，统一反色为白底。
            # （真实暗背景图如荧光显微图为 RGB 模式且 colorspace≠0，不会进入此分支）
            gray = img.convert("L")
            gray = ImageOps.invert(gray)
            gray = gray.convert("RGB")
            buf = io.BytesIO()
            gray.save(buf, format="PNG")
            return buf.getvalue(), True

        # ── 第二层：兜底，仅限 1-bit / 灰度 L 模式的亮度启发式 ──
        # 真实暗背景图（荧光显微图等）多为 RGB，不在此分支，不会被误伤。
        if img.mode == "1":
            hist = img.histogram()
            black_ratio = hist[0] / sum(hist) if sum(hist) > 0 else 0
            if black_ratio > 0.6:
                img = img.convert("L")
                img = ImageOps.invert(img)
                img = img.convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                return buf.getvalue(), True
        elif img.mode == "L":
            small = img.resize((50, 50))
            pixels = list(small.getdata())
            avg_brightness = sum(pixels) / len(pixels)
            if avg_brightness < 20:
                img = ImageOps.invert(img)
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                return buf.getvalue(), True

        # RGB / P / CMYK：不做亮度反色，避免误伤真实暗背景图
        return image_data, False

    except Exception:
        return image_data, False
# ═══════════════════════════════════════════════════════════

def detect_vector_figure_rects(page: fitz.Page, tab_rects: list) -> list:
    """检测 PDF 中的矢量图区域（用于渲染为 PNG）。"""
    if not page:
        return []

    drawings = page.get_drawings()
    if not drawings:
        return []

    page_rect = page.rect
    page_area = page_rect.width * page_rect.height
    if page_area <= 0:
        return []

    drawing_rects = []
    for d in drawings:
        rect = d.get("rect")
        if not rect:
            continue
        w = rect.width
        h = rect.height
        area = w * h
        if (w < MIN_VECTOR_FIGURE_WIDTH or h < MIN_VECTOR_FIGURE_HEIGHT
                or area < MIN_VECTOR_FIGURE_AREA):
            continue
        aspect = max(w, h) / max(min(w, h), 1)
        if aspect > MAX_VECTOR_FIGURE_ASPECT_RATIO:
            continue
        if area / page_area > MAX_VECTOR_BACKGROUND_AREA_RATIO:
            continue

        padded = fitz.Rect(
            max(0, rect.x0 - VECTOR_FIGURE_PADDING),
            max(0, rect.y0 - VECTOR_FIGURE_PADDING),
            min(page_rect.x1, rect.x1 + VECTOR_FIGURE_PADDING),
            min(page_rect.y1, rect.y1 + VECTOR_FIGURE_PADDING),
        )

        overlaps_table = False
        for tab_rect in tab_rects:
            if padded.intersects(tab_rect):
                intersect_area = (padded & tab_rect).get_area()
                if intersect_area > 0.5 * padded.get_area():
                    overlaps_table = True
                    break
        if overlaps_table:
            continue

        drawing_rects.append(padded)

    if not drawing_rects:
        return []

    merged = []
    for rect in drawing_rects:
        absorbed = False
        for i, existing in enumerate(merged):
            if rect.intersects(existing):
                union = rect | existing
                if union.get_area() / page_area < 0.8:
                    merged[i] = union
                    absorbed = True
                    break
        if not absorbed:
            merged.append(rect)

    return merged


MAX_VECTOR_BACKGROUND_AREA_RATIO = 1.9


# ═══════════════════════════════════════════════════════════
#  图注检测
# ═══════════════════════════════════════════════════════════

def find_figure_caption(page: fitz.Page, figure_rect: fitz.Rect) -> str:
    """在矢量图附近查找图注文本。"""
    blocks = page.get_text("blocks")
    caption_y_start = figure_rect.y1
    caption_y_end = figure_rect.y1 + 50
    caption_x_range = (figure_rect.x0 - 20, figure_rect.x1 + 20)

    for b in blocks:
        b_rect = fitz.Rect(b[:4])
        text = b[4].strip()
        if not text:
            continue
        if (caption_y_start - 10 <= b_rect.y0 <= caption_y_end
                and b_rect.x0 >= caption_x_range[0]
                and b_rect.x1 <= caption_x_range[1] + 100):
            if FIGURE_CAPTION_RE.match(text):
                return text
    return ""


# ═══════════════════════════════════════════════════════════
#  表格提取（简化版：仅 PyMuPDF 原生表格）
# ═══════════════════════════════════════════════════════════

def find_page_tables(page: fitz.Page) -> list:
    """使用 PyMuPDF 原生表格检测，返回 [{bbox, content, method}]。"""
    candidates = []
    try:
        for tab in page.find_tables():
            content = _table_to_markdown(tab)
            if content:
                candidates.append({
                    "bbox": fitz.Rect(tab.bbox) if hasattr(tab, 'bbox') else None,
                    "content": content,
                    "method": "native",
                })
    except Exception:
        pass
    return candidates


def _table_to_markdown(tab) -> str:
    """PyMuPDF 原生表格 → Markdown 表格。"""
    try:
        rows = tab.extract()
        if not rows or len(rows) < 2:
            return ""

        lines = []
        for i, row in enumerate(rows):
            cells = [str(c).strip() if c else "" for c in row]
            lines.append("| " + " | ".join(cells) + " |")
            if i == 0:
                sep = ["---"] * len(cells)
                lines.append("| " + " | ".join(sep) + " |")

        return "\n".join(lines)
    except Exception:
        return ""


# ═══════════════════════════════════════════════════════════
#  核心提取函数
# ═══════════════════════════════════════════════════════════

def extract_pdf_to_markdown(
    pdf_path: str,
    output_path: str = None,
    images: str = "filtered",
    render_vector_figures: bool = False,
    vector_figure_dpi: int = VECTOR_FIGURE_DPI,
) -> str:
    """从 PDF 提取文本、图片和表格，转换为 Markdown。

    Args:
        pdf_path: PDF 文件路径
        output_path: Markdown 输出路径
        images: 图片提取模式 (filtered/all/none)
        render_vector_figures: 是否渲染矢量图为 PNG
        vector_figure_dpi: 矢量图渲染 DPI
    """
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"[ERROR] Failed to open PDF file: {e}")
        return ""

    if len(doc) >= 200:
        print(f"[HINT] {len(doc)} pages — for very large PDFs, consider splitting "
              f"the source by chapter beforehand.")

    filename = Path(pdf_path).stem
    title = re.sub(r'^\d+-', '', filename).strip()

    print(f"[INFO] Analyzing document structure...")
    size_map = analyze_font_sizes(doc)
    print(f"   Font size mapping: body={size_map.get('body', 'N/A')}, "
          f"H1={size_map.get('h1', 'N/A')}, H2={size_map.get('h2', 'N/A')}, "
          f"H3={size_map.get('h3', 'N/A')}")

    print(f"[INFO] Detecting repeated headers/footers...")
    noise_texts = detect_headers_footers(doc)
    if noise_texts:
        print(f"   Found {len(noise_texts)} repeated noise texts (will be removed)")

    markdown_content = f"# {title}\n\n"
    seen_image_hashes = set()

    img_dir = None
    rel_img_dir = None
    if output_path:
        output_path = Path(output_path)
        rel_img_dir = f"{output_path.stem}_files"
        img_dir = output_path.parent / rel_img_dir

    img_count = 0
    image_manifest = []

    for page_num, page in enumerate(doc, 1):
        if page_num > 1:
            markdown_content += f"\n\n<!-- Page {page_num} -->\n\n"

        table_candidates = find_page_tables(page)
        tab_rects = [
            c["bbox"] for c in table_candidates
            if isinstance(c["bbox"], fitz.Rect)
        ]

        page_elements = []

        for table in table_candidates:
            bbox = table["bbox"]
            if not isinstance(bbox, fitz.Rect):
                continue
            page_elements.append({
                "y0": bbox.y0,
                "type": 2,
                "content": table["content"]
            })
            print(f"  [OK] Found table: P{page_num} ({table['method']})")

        if render_vector_figures:
            for figure_rect in detect_vector_figure_rects(page, tab_rects):
                page_elements.append({
                    "y0": figure_rect.y0,
                    "type": 3,
                    "content": figure_rect,
                })
                print(f"  [OK] Found vector figure region: P{page_num}")

        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            block_rect = fitz.Rect(block["bbox"])

            is_in_table = False
            for tab_rect in tab_rects:
                intersect = block_rect & tab_rect
                if intersect.get_area() > 0.6 * block_rect.get_area():
                    is_in_table = True
                    break
            if is_in_table:
                continue

            if block["type"] == 0:
                block_text_full = "".join(
                    span["text"] for line in block["lines"]
                    for span in line["spans"]
                ).strip()
                if block_text_full in noise_texts:
                    continue

                for line in block["lines"]:
                    line_text = ""
                    line_size = 0
                    line_flags = 0
                    is_code_line = False
                    formatted_spans = []

                    for span in line["spans"]:
                        span_text = CONTROL_CHARS_RE.sub('', span["text"])
                        if not span_text.strip():
                            if span_text:
                                formatted_spans.append(span_text)
                            continue

                        span_size = span["size"]
                        span_flags = span["flags"]
                        line_size = max(line_size, span_size)
                        line_flags |= span_flags

                        heading_level = get_heading_level(
                            span_size, size_map, span_text, span_flags
                        )

                        font_name = span.get("font", "")
                        if is_monospace_font(font_name):
                            is_code_line = True
                            formatted_spans.append(span_text)
                        elif heading_level > 0:
                            formatted_spans.append(span_text.strip())
                        else:
                            formatted_spans.append(
                                format_span_text(span_text, span_flags)
                            )

                    line_text = ''.join(formatted_spans).strip()
                    if not line_text:
                        continue

                    if line_text in noise_texts:
                        continue

                    line_text = merge_adjacent_formatting(line_text)
                    heading_level = get_heading_level(
                        line_size, size_map, line_text, line_flags
                    )
                    is_list, list_type, list_content = detect_list_item(line_text)

                    if heading_level > 0:
                        prefix = '#' * heading_level + ' '
                        clean_line = re.sub(r'\*+([^*]+)\*+', r'\1', line_text)
                        final_text = prefix + clean_line
                    elif is_list:
                        final_text = list_content
                    else:
                        final_text = line_text

                    page_elements.append({
                        "y0": line["bbox"][1],
                        "type": 0,
                        "content": final_text,
                        "is_heading": heading_level > 0,
                        "is_list": is_list,
                        "is_code": is_code_line
                    })

            elif block["type"] == 1:
                if images == "none":
                    pass
                elif images == "all" or should_keep_image(
                    block, page.rect, seen_image_hashes
                ):
                    page_elements.append({
                        "y0": block["bbox"][1],
                        "type": 1,
                        "content": block
                    })
                else:
                    w, h = block.get("width", 0), block.get("height", 0)
                    print(f"  [SKIP] Filtered image: {w}x{h}px")

        page_elements.sort(key=lambda x: x["y0"])
        page_elements = merge_adjacent_headings(page_elements)

        # 合并相邻文本行
        merged_elements = []
        i = 0
        while i < len(page_elements):
            el = page_elements[i]
            if el["type"] == 0 and not el.get("is_heading") and not el.get("is_list"):
                merged_content = el["content"]
                j = i + 1
                while j < len(page_elements):
                    next_el = page_elements[j]
                    if next_el["type"] != 0:
                        break
                    if not should_merge_lines(
                        {"content": merged_content, "is_heading": False, "is_list": False},
                        next_el
                    ):
                        break
                    merged_content += " " + next_el["content"]
                    j += 1
                merged_elements.append({
                    "type": 0,
                    "content": remove_page_footer(merged_content),
                    "is_heading": False,
                    "is_list": False
                })
                i = j
            else:
                merged_elements.append(el)
                i += 1

        prev_was_list = False
        prev_was_code = False
        code_block_lines = []

        def flush_code_block():
            nonlocal code_block_lines, markdown_content
            if code_block_lines:
                markdown_content += "```\n"
                markdown_content += "\n".join(code_block_lines) + "\n"
                markdown_content += "```\n\n"
                code_block_lines = []

        for el in merged_elements:
            if el["type"] == 0:
                is_list = el.get("is_list", False)
                is_heading = el.get("is_heading", False)
                is_code = el.get("is_code", False)

                if is_code:
                    if prev_was_list:
                        markdown_content += "\n"
                        prev_was_list = False
                    code_block_lines.append(el["content"])
                    prev_was_code = True
                else:
                    if prev_was_code:
                        flush_code_block()
                        prev_was_code = False

                    if is_heading:
                        if prev_was_list:
                            markdown_content += "\n"
                        markdown_content += el["content"] + "\n\n"
                        prev_was_list = False
                    elif is_list:
                        markdown_content += el["content"] + "\n"
                        prev_was_list = True
                    else:
                        if prev_was_list:
                            markdown_content += "\n"
                        markdown_content += el["content"] + "\n\n"
                        prev_was_list = False

            elif el["type"] == 2:
                if prev_was_code:
                    flush_code_block()
                    prev_was_code = False
                if prev_was_list:
                    markdown_content += "\n"
                markdown_content += el["content"] + "\n\n"
                prev_was_list = False

            elif el["type"] == 1:
                if prev_was_code:
                    flush_code_block()
                    prev_was_code = False
                if img_dir:
                    block = el["content"]
                    ext = block["ext"]
                    image_data = block["image"]

                    # 反色检测与修复（传入 block 元数据做分层判定，避免误伤暗背景图）
                    image_data, was_fixed = fix_inverted_image(image_data, ext, block)
                    if was_fixed:
                        ext = "png"  # 修复后统一存为 PNG
                        print(f"  [FIX] Image color inversion detected and fixed (page {page_num})")

                    safe_filename = filename.replace(" ", "_")
                    image_name = f"{safe_filename}_p{page_num}_{img_count}.{ext}"
                    image_path = img_dir / image_name

                    try:
                        img_dir.mkdir(parents=True, exist_ok=True)
                        with open(image_path, "wb") as f:
                            f.write(image_data)

                        if prev_was_list:
                            markdown_content += "\n"
                        markdown_content += f"![{image_name}]({rel_img_dir}/{image_name})\n\n"

                        width = int(block.get("width", 0) or 0)
                        height = int(block.get("height", 0) or 0)
                        ratio = width / height if width > 0 and height > 0 else None
                        sha = hashlib.sha256(image_data).hexdigest()

                        # 查找图注
                        caption = find_figure_caption(
                            page, fitz.Rect(block.get("bbox", [0, 0, 0, 0]))
                        )
                        figure_num_match = re.search(
                            r'(?:Figure|Fig\.?)\s*(\d+)', caption, re.IGNORECASE
                        )
                        figure_number = int(figure_num_match.group(1)) if figure_num_match else None

                        # 兼容 filter_images.py 的字段名
                        image_manifest.append({
                            "index": len(image_manifest) + 1,
                            "filename": image_name,
                            "original_filename": image_name,
                            "asset_kind": "bitmap",
                            "svg_renderable": True,
                            "pptx_native_supported": True,
                            "source_kind": "pdf_image",
                            "source_ext": f".{ext}",
                            "page_number": page_num,
                            "page_index": page_num,
                            "occurrence_index": img_count + 1,
                            "pixel_width": width or None,
                            "pixel_height": height or None,
                            "pixel_ratio": round(ratio, 6) if ratio else None,
                            "display_ratio": round(ratio, 6) if ratio else None,
                            "sha256": sha,
                            "source_sha256": sha,
                            "bbox": list(block.get("bbox", [])),
                            "caption": caption,
                            "caption_text": caption,
                            "figure_number": figure_number,
                        })
                        img_count += 1
                        prev_was_list = False
                        print(f"  [OK] Extracted image: {image_name}")
                    except Exception as e:
                        print(f"  [WARN] Failed to save image: {e}")

            elif el["type"] == 3:
                if prev_was_code:
                    flush_code_block()
                    prev_was_code = False
                if img_dir:
                    figure_rect = el["content"]
                    safe_filename = filename.replace(" ", "_")
                    image_name = f"{safe_filename}_p{page_num}_figure_{img_count}.png"
                    image_path = img_dir / image_name

                    try:
                        img_dir.mkdir(parents=True, exist_ok=True)
                        scale = vector_figure_dpi / 72
                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(scale, scale),
                            clip=figure_rect,
                            alpha=False,
                        )
                        pix.save(str(image_path))

                        if prev_was_list:
                            markdown_content += "\n"
                        markdown_content += f"![{image_name}]({rel_img_dir}/{image_name})\n\n"

                        ratio = pix.width / pix.height if pix.width > 0 and pix.height > 0 else None
                        caption = find_figure_caption(page, figure_rect)
                        figure_num_match = re.search(
                            r'(?:Figure|Fig\.?)\s*(\d+)', caption, re.IGNORECASE
                        )
                        figure_number = int(figure_num_match.group(1)) if figure_num_match else None

                        # 读取渲染后的文件计算 sha256
                        with open(image_path, "rb") as f:
                            sha = hashlib.sha256(f.read()).hexdigest()

                        image_manifest.append({
                            "index": len(image_manifest) + 1,
                            "filename": image_name,
                            "original_filename": image_name,
                            "asset_kind": "bitmap",
                            "svg_renderable": True,
                            "pptx_native_supported": True,
                            "source_kind": "pdf_vector_figure",
                            "source_ext": ".png",
                            "page_number": page_num,
                            "page_index": page_num,
                            "occurrence_index": img_count + 1,
                            "pixel_width": pix.width,
                            "pixel_height": pix.height,
                            "pixel_ratio": round(ratio, 6) if ratio else None,
                            "display_ratio": round(ratio, 6) if ratio else None,
                            "sha256": sha,
                            "source_sha256": sha,
                            "bbox": [
                                figure_rect.x0, figure_rect.y0,
                                figure_rect.x1, figure_rect.y1,
                            ],
                            "caption": caption,
                            "caption_text": caption,
                            "figure_number": figure_number,
                        })
                        img_count += 1
                        prev_was_list = False
                        print(f"  [OK] Rendered vector figure: {image_name}")
                    except Exception as e:
                        print(f"  [WARN] Failed to render vector figure: {e}")

        if prev_was_code:
            flush_code_block()

    doc.close()

    markdown_content = CONTROL_CHARS_RE.sub('', markdown_content)
    markdown_content = re.sub(r'\n{3,}', '\n\n', markdown_content)
    markdown_content = markdown_content.strip() + "\n"

    if output_path:
        os.makedirs(os.path.dirname(str(output_path)) or '.', exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        if img_dir and image_manifest:
            (img_dir / "image_manifest.json").write_text(
                json.dumps(image_manifest, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        print(f"[OK] Saved Markdown to: {output_path}")
        if image_manifest:
            print(f"   Extracted {len(image_manifest)} images")
            print(f"   Manifest: {img_dir / 'image_manifest.json'}")

    return markdown_content


# ═══════════════════════════════════════════════════════════
#  CLI 入口
# ═══════════════════════════════════════════════════════════

def main() -> int:
    parser = argparse.ArgumentParser(
        description='PDF to Markdown converter (self-contained, for paper-report-ppt)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python parse_pdf.py paper.pdf -o output.md
  python parse_pdf.py paper.pdf -o output.md --images all
  python parse_pdf.py paper.pdf -o output.md --render-vector-figures

Output:
  <stem>.md                    — Structured Markdown
  <stem>_files/                — Extracted images (original bytes, no resampling)
  <stem>_files/image_manifest.json — Image manifest (compatible with filter_images.py)

Features:
  - Auto-detect heading levels (font size analysis)
  - Bold/italic formatting
  - Ordered/unordered list detection
  - PyMuPDF native table extraction
  - Repeated header/footer removal
  - Page markers (<!-- Page N -->)
  - Image filtering (size/quality/dedup)
  - Vector figure rendering (optional)
  - Figure caption detection
'''
    )

    parser.add_argument('input', help='PDF file path')
    parser.add_argument('-o', '--output', required=True,
                        help='Output Markdown file path')
    parser.add_argument('--images', choices=['all', 'filtered', 'none'],
                        default='filtered',
                        help='Image extraction mode (default: filtered)')
    parser.add_argument('--render-vector-figures', action='store_true',
                        help='Render large vector drawing regions as PNG')
    parser.add_argument('--vector-figure-dpi', type=int, default=VECTOR_FIGURE_DPI,
                        help=f'DPI for vector figure rendering (default: {VECTOR_FIGURE_DPI})')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"[ERROR] File not found: {args.input}", file=sys.stderr)
        return 1

    result = extract_pdf_to_markdown(
        args.input,
        args.output,
        images=args.images,
        render_vector_figures=args.render_vector_figures,
        vector_figure_dpi=args.vector_figure_dpi,
    )

    return 0 if result else 1


if __name__ == '__main__':
    sys.exit(main())
