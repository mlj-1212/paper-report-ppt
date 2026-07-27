#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_pptx.py —— paper-report-ppt 技能的"直接生成"路径
==========================================
直接使用 python-pptx 生成高质量、完全可编辑的 PPTX 文件，
绕过 SVG 管线，输出质量对标 SVG 管线（每页 15-35 个形状、丰富装饰、精细布局）。

用法:
    python gen_pptx.py --input slides.json --images-dir ./images --output output.pptx --theme academic
"""

import argparse
import json
import os
import sys
from pathlib import Path

# ── 依赖检查 ──────────────────────────────────────────────
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    print("错误: 缺少 python-pptx 库。请执行: pip install python-pptx")
    sys.exit(1)

# ═══════════════════════════════════════════════════════════
#  主题配置
# ═══════════════════════════════════════════════════════════

THEMES = {
    "academic": {
        # 颜色
        "background": RGBColor(0xF8, 0xF9, 0xFA),       # #F8F9FA 浅灰背景
        "header_bar": RGBColor(0x2C, 0x3E, 0x50),        # #2C3E50 深蓝灰顶部装饰条
        "accent": RGBColor(0x34, 0x98, 0xDB),             # #3498DB 蓝色强调
        "highlight": RGBColor(0xE8, 0xF4, 0xFD),         # #E8F4FD 浅蓝高亮
        "body_text": RGBColor(0x2C, 0x3E, 0x50),          # #2C3E50 正文
        "secondary_text": RGBColor(0x7F, 0x8C, 0x8D),     # #7F8C8D 次要文字
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "border_gray": RGBColor(0xBD, 0xC3, 0xC7),        # #BDC3C7 图片边框
        "decor_corner": RGBColor(0xE8, 0xF4, 0xFD),       # 角落装饰（50%透明度通过代码控制）
        "section_bg": RGBColor(0x34, 0x98, 0xDB),          # 章节分隔页背景带
        # 字体
        "font_title": "Microsoft YaHei",
        "font_body": "Microsoft YaHei",
        "font_caption": "Microsoft YaHei",
        "title_size_pt": 28,
        "body_size_pt": 14,
        "caption_size_pt": 12,
        # 布局
        "slide_width_inch": 13.33,
        "slide_height_inch": 7.5,
        "top_bar_height_inch": 0.3,
        "header_area_height_inch": 1.1,
        "margin_inch": 0.8,
        # 装饰开关
        "decorations": True,
    },
    "minimal": {
        "background": RGBColor(0xFF, 0xFF, 0xFF),
        "header_bar": RGBColor(0x2C, 0x3E, 0x50),
        "accent": RGBColor(0x34, 0x98, 0xDB),
        "highlight": RGBColor(0xF5, 0xF5, 0xF5),
        "body_text": RGBColor(0x33, 0x33, 0x33),
        "secondary_text": RGBColor(0x99, 0x99, 0x99),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "border_gray": RGBColor(0xDD, 0xDD, 0xDD),
        "decor_corner": RGBColor(0xF0, 0xF0, 0xF0),
        "section_bg": RGBColor(0x34, 0x98, 0xDB),
        "font_title": "Microsoft YaHei",
        "font_body": "Microsoft YaHei",
        "font_caption": "Microsoft YaHei",
        "title_size_pt": 26,
        "body_size_pt": 14,
        "caption_size_pt": 12,
        "slide_width_inch": 13.33,
        "slide_height_inch": 7.5,
        "top_bar_height_inch": 0.15,
        "header_area_height_inch": 1.0,
        "margin_inch": 1.0,
        "decorations": False,
    },
}

# ═══════════════════════════════════════════════════════════
#  辅助函数
# ═══════════════════════════════════════════════════════════

def _set_font(run, font_name, size_pt, bold=False, color=None, italic=False):
    """设置文本 run 的字体属性，同时处理 CJK 字体兼容。"""
    run.font.name = font_name
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if size_pt:
        run.font.size = Pt(size_pt)
    if color:
        run.font.color.rgb = color
    # CJK 字体: 通过 XML 设置 eastAsia 字体名，确保中文显示正确
    rPr = run._r.get_or_add_rPr()
    nsmap = rPr.nsmap if hasattr(rPr, "nsmap") else None
    # 构造 eastAsia 属性
    ea_elem = rPr.makeelement(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}eastAsia",
        {"typeface": font_name},
    )
    rPr.append(ea_elem)


def _add_textbox(slide, left, top, width, height, text, font_name, size_pt,
                color, bold=False, italic=False, alignment=PP_ALIGN.LEFT,
                word_wrap=True):
    """在幻灯片上添加一个文本框并返回 shape 对象。"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    _set_font(run, font_name, size_pt, bold=bold, color=color, italic=italic)
    return txBox


def _add_rect(slide, left, top, width, height, fill_color, line_color=None,
              line_width=None):
    """添加矩形装饰形状。"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, height
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def _add_rounded_rect(slide, left, top, width, height, fill_color,
                      line_color=None, line_width=None):
    """添加圆角矩形装饰形状。"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def _add_oval(slide, left, top, width, height, fill_color, line_color=None):
    """添加圆形/椭圆装饰形状。"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, left, top, width, height
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def _set_shape_opacity(shape, opacity_pct):
    """通过 XML 设置形状填充的透明度 (0-100, 0=完全透明)。"""
    try:
        spPr = shape._element.spPr
        solidFill = spPr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
        if solidFill is not None:
            srgbClr = solidFill.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
            if srgbClr is not None:
                ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
                alpha_elem = srgbClr.makeelement(
                    f"{{{ns}}}alpha",
                    {"val": str(int((1 - opacity_pct / 100) * 100000))},
                )
                # 移除已有的 alpha 元素
                for old in srgbClr.findall(f"{{{ns}}}alpha"):
                    srgbClr.remove(old)
                srgbClr.append(alpha_elem)
    except Exception:
        pass  # 透明度设置失败不影响主流程


def _add_notes(slide, notes_text):
    """为幻灯片添加演讲者备注。"""
    if notes_text:
        notes_slide = slide.notes_slide
        notes_tf = notes_slide.notes_text_frame
        notes_tf.text = notes_text


# ═══════════════════════════════════════════════════════════
#  SlideBuilder 主类
# ═══════════════════════════════════════════════════════════

class SlideBuilder:
    """PPTX 幻灯片构建器，支持多种页面类型和装饰。"""

    def __init__(self, theme_name="academic"):
        self.theme_name = theme_name
        self.theme = THEMES.get(theme_name, THEMES["academic"])
        self.sw = Inches(self.theme["slide_width_inch"])
        self.sh = Inches(self.theme["slide_height_inch"])

    def create_presentation(self):
        """创建空白演示文稿，设置 16:9 尺寸。"""
        prs = Presentation()
        prs.slide_width = self.sw
        prs.slide_height = self.sh
        return prs

    # ── 公共装饰方法 ──────────────────────────────────────

    def _add_top_bar(self, slide):
        """每页顶部的装饰横条 (深蓝灰)。"""
        _add_rect(
            slide,
            left=Emu(0), top=Emu(0),
            width=self.sw, height=Inches(self.theme["top_bar_height_inch"]),
            fill_color=self.theme["header_bar"],
        )

    def _add_bottom_bar(self, slide, page_num):
        """底部页码条。"""
        # 底部细线
        _add_rect(
            slide,
            left=Emu(0),
            top=self.sh - Inches(0.05),
            width=self.sw,
            height=Inches(0.05),
            fill_color=self.theme["accent"],
        )
        # 页码
        _add_textbox(
            slide,
            left=self.sw - Inches(1.0),
            top=self.sh - Inches(0.45),
            width=Inches(0.8),
            height=Inches(0.35),
            text=str(page_num),
            font_name=self.theme["font_caption"],
            size_pt=10,
            color=self.theme["secondary_text"],
            alignment=PP_ALIGN.RIGHT,
        )

    def _add_corner_decorations(self, slide):
        """在角落添加半透明装饰形状（仅 academic 主题）。"""
        if not self.theme["decorations"]:
            return
        t = self.theme
        # 右上角小圆形
        c1 = _add_oval(
            slide,
            left=self.sw - Inches(0.8),
            top=Inches(0.5),
            width=Inches(0.5),
            height=Inches(0.5),
            fill_color=t["decor_corner"],
        )
        _set_shape_opacity(c1, 50)
        # 左下角小方形
        s1 = _add_rect(
            slide,
            left=Inches(0.3),
            top=self.sh - Inches(1.0),
            width=Inches(0.35),
            height=Inches(0.35),
            fill_color=t["decor_corner"],
        )
        _set_shape_opacity(s1, 50)
        # 右下角圆形
        c2 = _add_oval(
            slide,
            left=self.sw - Inches(1.5),
            top=self.sh - Inches(1.3),
            width=Inches(0.3),
            height=Inches(0.3),
            fill_color=t["accent"],
        )
        _set_shape_opacity(c2, 70)

    # ── 各页面类型构建方法 ────────────────────────────────

    def build_cover(self, prs, slide_data, page_num):
        """构建封面页。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
        t = self.theme

        # 背景色
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        # 顶部装饰条
        self._add_top_bar(slide)

        # 装饰: 左上角圆角矩形
        dr1 = _add_rounded_rect(
            slide,
            left=Inches(0.5),
            top=Inches(0.6),
            width=Inches(1.2),
            height=Inches(0.4),
            fill_color=t["accent"],
        )
        _set_shape_opacity(dr1, 30)

        # 装饰: 右下角圆角矩形
        dr2 = _add_rounded_rect(
            slide,
            left=self.sw - Inches(2.0),
            top=self.sh - Inches(1.5),
            width=Inches(1.5),
            height=Inches(0.35),
            fill_color=t["accent"],
        )
        _set_shape_opacity(dr2, 20)

        # 标题 — 居中大号加粗
        title_text = slide_data.get("title", "论文标题")
        _add_textbox(
            slide,
            left=Inches(1.5),
            top=Inches(1.8),
            width=self.sw - Inches(3.0),
            height=Inches(1.5),
            text=title_text,
            font_name=t["font_title"],
            size_pt=40,
            color=t["header_bar"],
            bold=True,
            alignment=PP_ALIGN.CENTER,
        )

        # 装饰分隔线
        _add_rect(
            slide,
            left=Inches(4.5),
            top=Inches(3.5),
            width=Inches(4.33),
            height=Inches(0.04),
            fill_color=t["accent"],
        )

        # 副标题
        subtitle = slide_data.get("subtitle", "")
        if subtitle:
            _add_textbox(
                slide,
                left=Inches(2.0),
                top=Inches(3.8),
                width=self.sw - Inches(4.0),
                height=Inches(0.8),
                text=subtitle,
                font_name=t["font_body"],
                size_pt=20,
                color=t["secondary_text"],
                alignment=PP_ALIGN.CENTER,
            )

        # 作者 / 期刊信息
        info_lines = slide_data.get("bullets", [])
        y_pos = Inches(4.8)
        for line in info_lines:
            _add_textbox(
                slide,
                left=Inches(2.0),
                top=y_pos,
                width=self.sw - Inches(4.0),
                height=Inches(0.4),
                text=line,
                font_name=t["font_body"],
                size_pt=14,
                color=t["body_text"],
                alignment=PP_ALIGN.CENTER,
            )
            y_pos += Inches(0.45)

        # 底部日期
        _add_textbox(
            slide,
            left=Inches(2.0),
            top=self.sh - Inches(1.0),
            width=self.sw - Inches(4.0),
            height=Inches(0.4),
            text="2026",
            font_name=t["font_caption"],
            size_pt=12,
            color=t["secondary_text"],
            alignment=PP_ALIGN.CENTER,
        )

        self._add_bottom_bar(slide, page_num)
        _add_notes(slide, slide_data.get("notes", ""))

    def build_toc(self, prs, slide_data, page_num):
        """构建目录页。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        t = self.theme

        # 背景色
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        self._add_top_bar(slide)

        # 标题
        _add_textbox(
            slide,
            left=t.get("margin_inch", 0.8) and Inches(0.8),
            top=Inches(0.55),
            width=Inches(6.0),
            height=Inches(0.7),
            text=slide_data.get("title", "目录"),
            font_name=t["font_title"],
            size_pt=t["title_size_pt"],
            color=t["header_bar"],
            bold=True,
        )

        # 装饰左侧竖条
        _add_rect(
            slide,
            left=Inches(0.5),
            top=Inches(1.5),
            width=Inches(0.06),
            height=Inches(5.2),
            fill_color=t["accent"],
        )

        sections = slide_data.get("sections", [])
        left_col_x = Inches(1.2)
        right_col_x = Inches(7.0)
        y_start = Inches(1.8)
        line_height = Inches(0.7)

        for i, section in enumerate(sections):
            y = y_start + i * line_height
            # 编号圆圈
            circle = _add_oval(
                slide,
                left=left_col_x,
                top=y + Inches(0.05),
                width=Inches(0.45),
                height=Inches(0.45),
                fill_color=t["accent"],
            )
            # 编号文字
            circle_tf = circle.text_frame
            circle_tf.word_wrap = False
            p = circle_tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(i + 1)
            _set_font(run, t["font_body"], 14, bold=True, color=t["white"])

            # 章节名
            _add_textbox(
                slide,
                left=left_col_x + Inches(0.65),
                top=y,
                width=Inches(4.5),
                height=Inches(0.45),
                text=section,
                font_name=t["font_body"],
                size_pt=16,
                color=t["body_text"],
                bold=True,
            )

            # 右侧简要描述
            highlights = slide_data.get("highlights", [])
            if i < len(highlights):
                desc = highlights[i].get("content", "")
                _add_textbox(
                    slide,
                    left=right_col_x,
                    top=y + Inches(0.05),
                    width=Inches(5.0),
                    height=Inches(0.45),
                    text=desc,
                    font_name=t["font_body"],
                    size_pt=13,
                    color=t["secondary_text"],
                )

            # 分隔线（非最后一项）
            if i < len(sections) - 1:
                _add_rect(
                    slide,
                    left=left_col_x,
                    top=y + line_height - Inches(0.1),
                    width=self.sw - Inches(2.5),
                    height=Inches(0.01),
                    fill_color=t["border_gray"],
                )

        self._add_corner_decorations(slide)
        self._add_bottom_bar(slide, page_num)
        _add_notes(slide, slide_data.get("notes", ""))

    def build_section(self, prs, slide_data, page_num):
        """构建章节分隔页。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        t = self.theme

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        self._add_top_bar(slide)

        # 上方装饰线
        _add_rect(
            slide,
            left=Inches(3.0),
            top=Inches(2.2),
            width=Inches(7.33),
            height=Inches(0.04),
            fill_color=t["accent"],
        )

        # 彩色强调带
        _add_rect(
            slide,
            left=Inches(3.0),
            top=Inches(2.5),
            width=Inches(7.33),
            height=Inches(0.12),
            fill_color=t["section_bg"],
        )

        # 章节编号
        page_num_text = slide_data.get("page_num", page_num)
        _add_textbox(
            slide,
            left=Inches(3.0),
            top=Inches(2.9),
            width=Inches(7.33),
            height=Inches(0.7),
            text=f"Section {page_num_text}",
            font_name=t["font_body"],
            size_pt=18,
            color=t["accent"],
            bold=False,
            alignment=PP_ALIGN.CENTER,
        )

        # 章节名 — 居中大号
        section_title = slide_data.get("title", "")
        _add_textbox(
            slide,
            left=Inches(2.0),
            top=Inches(3.5),
            width=self.sw - Inches(4.0),
            height=Inches(1.2),
            text=section_title,
            font_name=t["font_title"],
            size_pt=36,
            color=t["header_bar"],
            bold=True,
            alignment=PP_ALIGN.CENTER,
        )

        # 下方装饰线
        _add_rect(
            slide,
            left=Inches(3.0),
            top=Inches(5.2),
            width=Inches(7.33),
            height=Inches(0.04),
            fill_color=t["accent"],
        )

        # 装饰: 两侧小圆
        if t["decorations"]:
            c1 = _add_oval(
                slide,
                left=Inches(1.5),
                top=Inches(3.5),
                width=Inches(0.6),
                height=Inches(0.6),
                fill_color=t["decor_corner"],
            )
            _set_shape_opacity(c1, 40)
            c2 = _add_oval(
                slide,
                left=self.sw - Inches(2.3),
                top=Inches(3.5),
                width=Inches(0.6),
                height=Inches(0.6),
                fill_color=t["decor_corner"],
            )
            _set_shape_opacity(c2, 40)

        self._add_bottom_bar(slide, page_num)
        _add_notes(slide, slide_data.get("notes", ""))

    def build_content(self, prs, slide_data, page_num):
        """构建内容页（带要点列表）。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        t = self.theme
        margin = Inches(t["margin_inch"])

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        # 顶部装饰条
        self._add_top_bar(slide)

        # 标题区域: 彩色条纹 + 标题文字
        header_y = Inches(t["top_bar_height_inch"])
        _add_rect(
            slide,
            left=Emu(0),
            top=header_y,
            width=self.sw,
            height=Inches(t["header_area_height_inch"]) - header_y,
            fill_color=t["white"],
        )
        # 标题区左侧蓝色强调线
        _add_rect(
            slide,
            left=Emu(0),
            top=header_y,
            width=Inches(0.08),
            height=Inches(t["header_area_height_inch"]) - header_y,
            fill_color=t["accent"],
        )

        # 标题文字
        _add_textbox(
            slide,
            left=margin,
            top=header_y + Inches(0.15),
            width=self.sw - margin * 2,
            height=Inches(0.6),
            text=slide_data.get("title", ""),
            font_name=t["font_title"],
            size_pt=t["title_size_pt"],
            color=t["header_bar"],
            bold=True,
        )

        # 标题区底部细线分隔
        header_bottom = Inches(t["header_area_height_inch"])
        _add_rect(
            slide,
            left=margin,
            top=header_bottom,
            width=self.sw - margin * 2,
            height=Inches(0.015),
            fill_color=t["border_gray"],
        )

        # 要点列表
        bullets = slide_data.get("bullets", [])
        body_top = header_bottom + Inches(0.3)
        bullet_left = margin + Inches(0.3)
        bullet_width = self.sw - margin * 2 - Inches(0.6)
        bullet_height = Inches(0.55)

        for i, bullet in enumerate(bullets):
            y = body_top + i * Inches(0.7)

            # 子弹符号 — 小圆点
            dot = _add_oval(
                slide,
                left=margin + Inches(0.05),
                top=y + Inches(0.12),
                width=Inches(0.15),
                height=Inches(0.15),
                fill_color=t["accent"],
            )

            # 要点文本框（独立文本框，可编辑）
            _add_textbox(
                slide,
                left=bullet_left,
                top=y,
                width=bullet_width,
                height=bullet_height,
                text=bullet,
                font_name=t["font_body"],
                size_pt=t["body_size_pt"],
                color=t["body_text"],
            )

        # 高亮框
        highlights = slide_data.get("highlights", [])
        if highlights:
            hl_y = body_top + len(bullets) * Inches(0.7) + Inches(0.3)
            for j, hl in enumerate(highlights):
                hl_box = _add_rounded_rect(
                    slide,
                    left=margin + Inches(0.3),
                    top=hl_y + j * Inches(1.0),
                    width=Inches(5.0),
                    height=Inches(0.8),
                    fill_color=t["highlight"],
                    line_color=t["accent"],
                    line_width=1,
                )
                # 高亮标题
                _add_textbox(
                    slide,
                    left=margin + Inches(0.6),
                    top=hl_y + j * Inches(1.0) + Inches(0.05),
                    width=Inches(4.5),
                    height=Inches(0.3),
                    text=hl.get("title", ""),
                    font_name=t["font_body"],
                    size_pt=13,
                    color=t["accent"],
                    bold=True,
                )
                # 高亮内容
                _add_textbox(
                    slide,
                    left=margin + Inches(0.6),
                    top=hl_y + j * Inches(1.0) + Inches(0.35),
                    width=Inches(4.5),
                    height=Inches(0.4),
                    text=hl.get("content", ""),
                    font_name=t["font_body"],
                    size_pt=12,
                    color=t["body_text"],
                )

        self._add_corner_decorations(slide)
        self._add_bottom_bar(slide, page_num)
        _add_notes(slide, slide_data.get("notes", ""))

    def build_figure(self, prs, slide_data, page_num, images_dir=None):
        """构建图表页，居中展示图片。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        t = self.theme

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        self._add_top_bar(slide)

        # 标题
        _add_textbox(
            slide,
            left=Inches(0.8),
            top=Inches(0.5),
            width=self.sw - Inches(1.6),
            height=Inches(0.6),
            text=slide_data.get("title", "图表"),
            font_name=t["font_title"],
            size_pt=t["title_size_pt"],
            color=t["header_bar"],
            bold=True,
        )

        # 标题下分隔线
        _add_rect(
            slide,
            left=Inches(0.8),
            top=Inches(1.15),
            width=self.sw - Inches(1.6),
            height=Inches(0.015),
            fill_color=t["border_gray"],
        )

        # 图片区域 — 预留边距 1 英寸
        img_margin = Inches(1.0)
        img_area_left = img_margin
        img_area_top = Inches(1.5)
        img_area_width = self.sw - img_margin * 2
        img_area_height = self.sh - Inches(2.8)

        # 图片背景框（浅色底）
        _add_rounded_rect(
            slide,
            left=img_area_left - Inches(0.1),
            top=img_area_top - Inches(0.1),
            width=img_area_width + Inches(0.2),
            height=img_area_height + Inches(0.2),
            fill_color=t["white"],
            line_color=t["border_gray"],
            line_width=1,
        )

        # 加载图片
        image_path = slide_data.get("image_path", "")
        actual_image_file = None
        if image_path:
            # 支持绝对路径和相对于 images_dir 的路径
            if os.path.isabs(image_path):
                actual_image_file = image_path
            elif images_dir:
                actual_image_file = os.path.join(images_dir, image_path)

        if actual_image_file and os.path.isfile(actual_image_file):
            try:
                # 保持宽高比计算
                from PIL import Image as PILImage
                pil_img = PILImage.open(actual_image_file)
                img_w, img_h = pil_img.size
                aspect = img_w / img_h

                # 在可用区域内按比例缩放
                avail_w = img_area_width - Inches(0.4)
                avail_h = img_area_height - Inches(0.4)
                if avail_w / avail_h > aspect:
                    # 高度受限
                    final_h = avail_h
                    final_w = int(final_h * aspect)
                else:
                    # 宽度受限
                    final_w = avail_w
                    final_h = int(final_w / aspect)

                # 居中放置
                img_left = img_area_left + (img_area_width - final_w) // 2
                img_top = img_area_top + (img_area_height - final_h) // 2

                slide.shapes.add_picture(
                    actual_image_file,
                    img_left, img_top,
                    final_w, final_h,
                )
            except ImportError:
                # 没有 Pillow, 使用 python-pptx 直接放置，不自动缩放
                slide.shapes.add_picture(
                    actual_image_file,
                    img_area_left + Inches(0.2),
                    img_area_top + Inches(0.2),
                    img_area_width - Inches(0.4),
                    img_area_height - Inches(0.4),
                )
            except Exception as e:
                print(f"警告: 无法加载图片 '{actual_image_file}': {e}")
                # 显示占位文本
                _add_textbox(
                    slide,
                    left=img_area_left,
                    top=img_area_top + Inches(2.0),
                    width=img_area_width,
                    height=Inches(0.5),
                    text="[图片加载失败]",
                    font_name=t["font_body"],
                    size_pt=16,
                    color=t["secondary_text"],
                    alignment=PP_ALIGN.CENTER,
                )
        else:
            # 无图片时显示占位
            placeholder = _add_rounded_rect(
                slide,
                left=img_area_left + Inches(1.5),
                top=img_area_top + Inches(1.5),
                width=img_area_width - Inches(3.0),
                height=img_area_height - Inches(3.0),
                fill_color=t["highlight"],
                line_color=t["border_gray"],
                line_width=1,
            )
            _add_textbox(
                slide,
                left=img_area_left + Inches(2.0),
                top=img_area_top + Inches(2.5),
                width=img_area_width - Inches(4.0),
                height=Inches(0.5),
                text="Figure Placeholder",
                font_name=t["font_body"],
                size_pt=16,
                color=t["secondary_text"],
                alignment=PP_ALIGN.CENTER,
            )

        # 图片说明
        caption = slide_data.get("image_caption", "")
        if caption:
            _add_textbox(
                slide,
                left=Inches(1.0),
                top=self.sh - Inches(1.3),
                width=self.sw - Inches(2.0),
                height=Inches(0.5),
                text=caption,
                font_name=t["font_caption"],
                size_pt=t["caption_size_pt"],
                color=t["secondary_text"],
                italic=True,
                alignment=PP_ALIGN.CENTER,
            )

        self._add_corner_decorations(slide)
        self._add_bottom_bar(slide, page_num)
        _add_notes(slide, slide_data.get("notes", ""))

    def build_conclusion(self, prs, slide_data, page_num):
        """构建结论页。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        t = self.theme

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        self._add_top_bar(slide)

        # 标题
        _add_textbox(
            slide,
            left=Inches(0.8),
            top=Inches(0.55),
            width=self.sw - Inches(1.6),
            height=Inches(0.7),
            text=slide_data.get("title", "总结"),
            font_name=t["font_title"],
            size_pt=t["title_size_pt"],
            color=t["header_bar"],
            bold=True,
        )

        # 分隔线
        _add_rect(
            slide,
            left=Inches(0.8),
            top=Inches(1.3),
            width=self.sw - Inches(1.6),
            height=Inches(0.015),
            fill_color=t["border_gray"],
        )

        # 核心信息 — 居中彩色框内
        key_message = slide_data.get("key_message", slide_data.get("title", ""))
        # 背景强调框
        box_w = Inches(9.0)
        box_h = Inches(1.6)
        box_left = (self.sw - box_w) // 2
        box_top = Inches(2.0)
        _add_rounded_rect(
            slide,
            left=box_left,
            top=box_top,
            width=box_w,
            height=box_h,
            fill_color=t["highlight"],
            line_color=t["accent"],
            line_width=2,
        )

        # 核心信息文字
        _add_textbox(
            slide,
            left=box_left + Inches(0.5),
            top=box_top + Inches(0.2),
            width=box_w - Inches(1.0),
            height=box_h - Inches(0.4),
            text=key_message,
            font_name=t["font_title"],
            size_pt=24,
            color=t["header_bar"],
            bold=True,
            alignment=PP_ALIGN.CENTER,
        )

        # 支撑要点
        bullets = slide_data.get("bullets", [])
        y = Inches(4.2)
        for i, bullet in enumerate(bullets):
            # 小装饰条
            _add_rect(
                slide,
                left=Inches(3.0),
                top=y + Inches(0.08),
                width=Inches(0.06),
                height=Inches(0.25),
                fill_color=t["accent"],
            )
            _add_textbox(
                slide,
                left=Inches(3.3),
                top=y,
                width=Inches(7.0),
                height=Inches(0.45),
                text=bullet,
                font_name=t["font_body"],
                size_pt=t["body_size_pt"],
                color=t["body_text"],
            )
            y += Inches(0.6)

        self._add_corner_decorations(slide)
        self._add_bottom_bar(slide, page_num)
        _add_notes(slide, slide_data.get("notes", ""))

    def build_qa(self, prs, slide_data, page_num):
        """构建问答/致谢页。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        t = self.theme

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        self._add_top_bar(slide)

        # 中央大文字 "Thank You / Q&A"
        _add_textbox(
            slide,
            left=Inches(2.0),
            top=Inches(1.8),
            width=self.sw - Inches(4.0),
            height=Inches(1.5),
            text=slide_data.get("title", "Thank You"),
            font_name=t["font_title"],
            size_pt=48,
            color=t["header_bar"],
            bold=True,
            alignment=PP_ALIGN.CENTER,
        )

        # 装饰分隔线
        _add_rect(
            slide,
            left=Inches(4.5),
            top=Inches(3.5),
            width=Inches(4.33),
            height=Inches(0.04),
            fill_color=t["accent"],
        )

        # Q&A 文字
        _add_textbox(
            slide,
            left=Inches(2.0),
            top=Inches(3.8),
            width=self.sw - Inches(4.0),
            height=Inches(0.8),
            text="Q & A",
            font_name=t["font_title"],
            size_pt=32,
            color=t["accent"],
            alignment=PP_ALIGN.CENTER,
        )

        # 底部装饰圆
        if t["decorations"]:
            c1 = _add_oval(
                slide,
                left=Inches(2.5),
                top=Inches(5.0),
                width=Inches(0.8),
                height=Inches(0.8),
                fill_color=t["decor_corner"],
            )
            _set_shape_opacity(c1, 40)
            c2 = _add_oval(
                slide,
                left=self.sw - Inches(3.5),
                top=Inches(5.0),
                width=Inches(0.8),
                height=Inches(0.8),
                fill_color=t["decor_corner"],
            )
            _set_shape_opacity(c2, 40)
            c3 = _add_oval(
                slide,
                left=Inches(5.8),
                top=Inches(5.3),
                width=Inches(1.2),
                height=Inches(1.2),
                fill_color=t["accent"],
            )
            _set_shape_opacity(c3, 20)

        # 副信息
        bullets = slide_data.get("bullets", [])
        if bullets:
            y = Inches(4.8)
            for bullet in bullets:
                _add_textbox(
                    slide,
                    left=Inches(3.0),
                    top=y,
                    width=Inches(7.0),
                    height=Inches(0.4),
                    text=bullet,
                    font_name=t["font_body"],
                    size_pt=14,
                    color=t["secondary_text"],
                    alignment=PP_ALIGN.CENTER,
                )
                y += Inches(0.45)

        self._add_bottom_bar(slide, page_num)
        _add_notes(slide, slide_data.get("notes", ""))


# ═══════════════════════════════════════════════════════════
#  页面类型路由
# ═══════════════════════════════════════════════════════════

PAGE_TYPE_BUILDERS = {
    "cover":     "build_cover",
    "toc":       "build_toc",
    "section":   "build_section",
    "content":   "build_content",
    "figure":    "build_figure",
    "model":     "build_content",       # model 页复用 content 布局
    "conclusion":"build_conclusion",
    "qa":        "build_qa",
}


def build_slide(prs, builder, slide_data, page_idx, images_dir):
    """根据页面类型路由到对应的构建方法。"""
    page_type = slide_data.get("page_type", "content")
    method_name = PAGE_TYPE_BUILDERS.get(page_type, "build_content")
    method = getattr(builder, method_name)

    # figure 类型额外传入 images_dir
    if method_name == "build_figure":
        method(prs, slide_data, page_idx, images_dir=images_dir)
    else:
        method(prs, slide_data, page_idx)


# ═══════════════════════════════════════════════════════════
#  输入验证
# ═══════════════════════════════════════════════════════════

REQUIRED_FIELDS = {"page_num", "page_type", "title"}
OPTIONAL_FIELDS = {
    "subtitle", "sections", "bullets", "image_path", "image_caption",
    "notes", "highlights", "key_message",
}


def validate_slides(slides):
    """验证 slides.json 数据结构，返回错误列表。"""
    errors = []
    for i, s in enumerate(slides):
        # 检查必需字段
        missing = REQUIRED_FIELDS - set(s.keys())
        if missing:
            errors.append(f"  幻灯片 #{i+1}: 缺少必需字段 {missing}")
        # 检查 page_type
        if "page_type" in s and s["page_type"] not in PAGE_TYPE_BUILDERS:
            errors.append(
                f"  幻灯片 #{i+1}: 未知 page_type '{s['page_type']}', "
                f"支持: {list(PAGE_TYPE_BUILDERS.keys())}"
            )
    return errors


# ═══════════════════════════════════════════════════════════
#  主函数
# ═══════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="gen_pptx —— paper-report-ppt 直接生成路径, "
                    "使用 python-pptx 生成高质量可编辑 PPTX。",
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="slides.json 路径 (AI 生成的幻灯片定义文件)",
    )
    parser.add_argument(
        "--images-dir",
        default="./images",
        help="图片目录路径 (默认: ./images)",
    )
    parser.add_argument(
        "--output", "-o",
        default="./output.pptx",
        help="输出 PPTX 文件路径 (默认: ./output.pptx)",
    )
    parser.add_argument(
        "--theme",
        default="academic",
        choices=["academic", "minimal"],
        help="主题风格 (默认: academic)",
    )
    args = parser.parse_args()

    # ── 检查输入文件 ──
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误: 输入文件不存在: {args.input}")
        print("请确保 slides.json 文件存在，格式参见文档。")
        sys.exit(1)

    # ── 读取 JSON ──
    try:
        with open(input_path, "r", encoding="utf-8") as f:
            slides_data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"错误: slides.json 格式不合法: {e}")
        sys.exit(1)

    if not isinstance(slides_data, list) or len(slides_data) == 0:
        print("错误: slides.json 应为非空数组 (list of slide objects)。")
        sys.exit(1)

    # ── 验证数据 ──
    errors = validate_slides(slides_data)
    if errors:
        print("数据验证错误:")
        for err in errors:
            print(err)
        sys.exit(1)

    # ── 构建演示文稿 ──
    builder = SlideBuilder(theme_name=args.theme)
    prs = builder.create_presentation()

    print(f"主题: {args.theme}")
    print(f"幻灯片数量: {len(slides_data)}")
    print(f"图片目录: {args.images_dir}")
    print("-" * 50)

    for i, slide_data in enumerate(slides_data):
        page_num = slide_data.get("page_num", i + 1)
        page_type = slide_data.get("page_type", "content")
        title = slide_data.get("title", "(无标题)")

        build_slide(prs, builder, slide_data, page_num, args.images_dir)

        print(f"  [{page_num}/{len(slides_data)}] {page_type}: {title}")

    # ── 保存 ──
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))

    print("-" * 50)
    print(f"已生成: {output_path.resolve()}")
    print("完成!")


if __name__ == "__main__":
    main()
