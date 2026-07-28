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
        "title_band": RGBColor(0x1A, 0x52, 0x76),          # #1A5276 深蓝色标题带
        "conclusion_bg": RGBColor(0xEB, 0xF5, 0xFB),       # #EBF5FB 浅蓝色结论框背景
        "keyword_red": RGBColor(0xC0, 0x39, 0x2B),         # #C0392B 关键词红色
        "sub_title_bg": RGBColor(0xF0, 0xF7, 0xFC),        # #F0F7FC 小标题浅蓝背景
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
        "decorations": False,
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

    # ── TRAE 风格辅助方法 ────────────────────────────────

    # 学术关键词列表（用于 _add_rich_textbox 智能高亮）
    ACADEMIC_KEYWORDS = [
        "增加", "减少", "升高", "降低", "促进", "抑制", "上调", "下调",
        "增强", "减弱", "正向调控", "负向调控", "加重", "减轻",
        "稳定", "降解", "激活", "破坏",
    ]

    def _add_title_header(self, slide, title_text, theme, kicker=None):
        """在页面顶部画一个全宽深蓝色矩形标题带（替代原 _add_top_bar + 单独标题）。

        - 高 0.8 英寸，颜色用 title_band
        - 标题文字白色、加粗、左对齐（左边距 0.5 英寸，垂直居中）
        - 若提供 kicker，在右上角画一个浅色小标签（圆角矩形，白字）
        """
        band_height = Inches(0.8)
        # 全宽深蓝色矩形条
        _add_rect(
            slide,
            left=Emu(0), top=Emu(0),
            width=self.sw, height=band_height,
            fill_color=theme["title_band"],
        )
        # 标题文字（白色、加粗、左对齐、垂直居中）
        title_box = _add_textbox(
            slide,
            left=Inches(0.5),
            top=Emu(0),
            width=self.sw - Inches(1.0),
            height=band_height,
            text=title_text,
            font_name=theme["font_title"],
            size_pt=theme["title_size_pt"],
            color=theme["white"],
            bold=True,
            alignment=PP_ALIGN.LEFT,
        )
        # 垂直居中
        try:
            title_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception:
            pass

        # 右上角 kicker 小标签
        if kicker:
            kicker_w = Inches(1.8)
            kicker_h = Inches(0.35)
            kicker_left = self.sw - kicker_w - Inches(0.4)
            kicker_top = (band_height - kicker_h) // 2
            k_shape = _add_rounded_rect(
                slide,
                left=kicker_left,
                top=kicker_top,
                width=kicker_w,
                height=kicker_h,
                fill_color=theme["accent"],
            )
            k_tf = k_shape.text_frame
            k_tf.word_wrap = False
            try:
                k_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            except Exception:
                pass
            k_p = k_tf.paragraphs[0]
            k_p.alignment = PP_ALIGN.CENTER
            k_run = k_p.add_run()
            k_run.text = kicker
            _set_font(k_run, theme["font_body"], 11, bold=True, color=theme["white"])

    def _add_conclusion_box(self, slide, left, top, width, text, theme):
        """画一个浅蓝色圆角矩形结论框，左侧带蓝色竖线装饰。

        - conclusion_bg 填充，无边框
        - 左侧 accent 色竖线（宽 0.06 英寸）
        - 文字为 header_bar 色（深蓝），加粗，size 13pt，前加"结论："前缀（加粗）
        """
        box_height = Inches(0.7)
        # 浅蓝色圆角矩形背景
        _add_rounded_rect(
            slide,
            left=left,
            top=top,
            width=width,
            height=box_height,
            fill_color=theme["conclusion_bg"],
            line_color=None,
        )
        # 左侧蓝色竖线装饰
        _add_rect(
            slide,
            left=left,
            top=top,
            width=Inches(0.06),
            height=box_height,
            fill_color=theme["accent"],
        )
        # 文字框（"结论："前缀加粗 + 正文）
        txBox = slide.shapes.add_textbox(
            left + Inches(0.25),
            top,
            width - Inches(0.4),
            box_height,
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        try:
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception:
            pass
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        # "结论：" 前缀
        prefix_run = p.add_run()
        prefix_run.text = "结论："
        _set_font(prefix_run, theme["font_body"], 13, bold=True,
                  color=theme["header_bar"])
        # 正文（支持关键词高亮）
        self._add_keyword_runs(p, text, theme["font_body"], 13,
                               theme["header_bar"], theme)
        return box_height

    def _add_keyword_runs(self, paragraph, text, font_name, size_pt,
                          default_color, theme, bold=False):
        """将文本按学术关键词切分，添加多个 run（关键词标红加粗）。

        关键词用 keyword_red + bold，其余用 default_color。
        """
        keywords = self.ACADEMIC_KEYWORDS
        # 按出现位置切分文本
        remaining = text
        while remaining:
            # 找最早出现的关键词
            earliest_pos = -1
            earliest_kw = None
            for kw in keywords:
                pos = remaining.find(kw)
                if pos != -1 and (earliest_pos == -1 or pos < earliest_pos):
                    earliest_pos = pos
                    earliest_kw = kw
            if earliest_pos == -1:
                # 没有关键词了，输出剩余文本
                run = paragraph.add_run()
                run.text = remaining
                _set_font(run, font_name, size_pt, bold=bold, color=default_color)
                break
            # 输出关键词前的普通文本
            if earliest_pos > 0:
                normal_text = remaining[:earliest_pos]
                run = paragraph.add_run()
                run.text = normal_text
                _set_font(run, font_name, size_pt, bold=bold, color=default_color)
            # 输出关键词（标红加粗）
            kw_run = paragraph.add_run()
            kw_run.text = earliest_kw
            _set_font(kw_run, font_name, size_pt, bold=True,
                      color=theme["keyword_red"])
            # 继续处理剩余文本
            remaining = remaining[earliest_pos + len(earliest_kw):]

    def _add_rich_textbox(self, slide, left, top, width, height, text,
                          font_name, size_pt, color, theme, bold=False):
        """智能关键词高亮文本框：检测学术关键词并标红加粗。

        关键词部分用 keyword_red + bold，其余用正常 color。
        用 add_run 机制，处理一段文本中多个关键词的情况。
        """
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        self._add_keyword_runs(p, text, font_name, size_pt, color, theme, bold=bold)
        return txBox

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

        # 顶部装饰条（使用 title_band 色）
        _add_rect(
            slide,
            left=Emu(0), top=Emu(0),
            width=self.sw, height=Inches(t["top_bar_height_inch"]),
            fill_color=t["title_band"],
        )

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

            # 章节名 —— 兼容字符串和字典两种 sections 格式
            section_title = _extract_section_title(section)

            _add_textbox(
                slide,
                left=left_col_x + Inches(0.65),
                top=y,
                width=Inches(4.5),
                height=Inches(0.45),
                text=section_title,
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

        # 顶部 title_band 色带（替代原 header_bar 细条）
        _add_rect(
            slide,
            left=Emu(0), top=Emu(0),
            width=self.sw, height=Inches(0.4),
            fill_color=t["title_band"],
        )

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
            fill_color=t["title_band"],
        )

        # 章节编号（优先用 section_no，回退到 page_num）
        section_no = slide_data.get("section_no", page_num)
        _add_textbox(
            slide,
            left=Inches(3.0),
            top=Inches(2.9),
            width=Inches(7.33),
            height=Inches(0.7),
            text=f"Section {section_no}",
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

        # 标题带（替代原 _add_top_bar + 单独标题）
        kicker = slide_data.get("kicker")
        self._add_title_header(slide, slide_data.get("title", ""), t, kicker=kicker)

        # 小标题（如果有 sub_title）
        sub_y = Inches(0.92)
        sub_title = slide_data.get("sub_title")
        if sub_title:
            # 蓝色竖线装饰
            _add_rect(
                slide,
                left=margin,
                top=sub_y + Inches(0.02),
                width=Inches(0.06),
                height=Inches(0.32),
                fill_color=t["accent"],
            )
            # 小标题浅蓝背景
            _add_rounded_rect(
                slide,
                left=margin + Inches(0.15),
                top=sub_y,
                width=self.sw - margin * 2 - Inches(0.15),
                height=Inches(0.36),
                fill_color=t["sub_title_bg"],
                line_color=None,
            )
            _add_textbox(
                slide,
                left=margin + Inches(0.3),
                top=sub_y + Inches(0.02),
                width=self.sw - margin * 2 - Inches(0.45),
                height=Inches(0.32),
                text=sub_title,
                font_name=t["font_body"],
                size_pt=14,
                color=t["title_band"],
                bold=True,
            )
            body_top = sub_y + Inches(0.52)
        else:
            body_top = Inches(1.15)

        # 要点列表
        bullets = slide_data.get("bullets", [])
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

            # 要点文本框（支持关键词高亮）
            self._add_rich_textbox(
                slide,
                left=bullet_left,
                top=y,
                width=bullet_width,
                height=bullet_height,
                text=bullet,
                font_name=t["font_body"],
                size_pt=t["body_size_pt"],
                color=t["body_text"],
                theme=t,
            )

        # 高亮框
        highlights = slide_data.get("highlights", [])
        hl_y = body_top + len(bullets) * Inches(0.7) + Inches(0.2)
        if highlights:
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

        # 结论框（如果有 conclusion 字段）
        conclusion = slide_data.get("conclusion")
        if conclusion:
            concl_top = self.sh - Inches(1.5)
            self._add_conclusion_box(
                slide,
                left=margin + Inches(0.3),
                top=concl_top,
                width=self.sw - margin * 2 - Inches(0.6),
                text=conclusion,
                theme=t,
            )

        self._add_corner_decorations(slide)
        self._add_bottom_bar(slide, page_num)
        _add_notes(slide, slide_data.get("notes", ""))

    def build_figure(self, prs, slide_data, page_num, images_dir=None):
        """构建图表页，支持左右分栏（文字+图片）或居中图片两种布局。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        t = self.theme

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        self._add_title_header(slide, slide_data.get("title", "图表"), t,
                               kicker=slide_data.get("kicker"))

        bullets = slide_data.get("bullets", [])
        has_bullets = bool(bullets)

        if has_bullets:
            # ── 左右分栏布局：左侧文字分析，右侧图片 ──
            text_col_width = Inches(5.0)
            img_col_width = self.sw - text_col_width - Inches(1.6)
            col_gap = Inches(0.4)
            text_left = Inches(0.8)
            img_left = text_left + text_col_width + col_gap
            content_top = Inches(1.15)
            content_height = self.sh - Inches(2.5)

            # 左侧文字区：极淡背景 + 无边框，靠留白自然区分
            _add_rounded_rect(
                slide,
                left=text_left - Inches(0.05),
                top=content_top - Inches(0.05),
                width=text_col_width + Inches(0.1),
                height=content_height + Inches(0.1),
                fill_color=RGBColor(0xFC, 0xFC, 0xFC),  # 极浅灰，几乎看不见
                line_color=None,
            )

            # 小标题（如果有 sub_title 或 analysis_title）
            sub_title = slide_data.get("sub_title") or slide_data.get("analysis_title")
            if sub_title:
                # 蓝色竖线装饰
                _add_rect(
                    slide,
                    left=text_left + Inches(0.05),
                    top=content_top + Inches(0.1),
                    width=Inches(0.06),
                    height=Inches(0.32),
                    fill_color=t["accent"],
                )
                # 小标题浅蓝背景
                _add_rounded_rect(
                    slide,
                    left=text_left + Inches(0.18),
                    top=content_top + Inches(0.08),
                    width=text_col_width - Inches(0.25),
                    height=Inches(0.36),
                    fill_color=t["sub_title_bg"],
                    line_color=None,
                )
                _add_textbox(
                    slide,
                    left=text_left + Inches(0.32),
                    top=content_top + Inches(0.1),
                    width=text_col_width - Inches(0.4),
                    height=Inches(0.32),
                    text=sub_title,
                    font_name=t["font_body"],
                    size_pt=14,
                    color=t["title_band"],
                    bold=True,
                )
                body_top = content_top + Inches(0.55)
            else:
                # 左侧蓝色竖线装饰（仿TRAE风格，强调文字区）
                _add_rect(
                    slide,
                    left=text_left + Inches(0.05),
                    top=content_top + Inches(0.2),
                    width=Inches(0.04),
                    height=Inches(0.6),
                    fill_color=t["accent"],
                )
                body_top = content_top + Inches(0.2)

            # 要点列表
            bullet_left = text_left + Inches(0.3)
            bullet_width = text_col_width - Inches(0.6)
            bullet_height = Inches(0.55)
            max_bullets = min(len(bullets), 6)  # 防止溢出

            for i, bullet in enumerate(bullets[:max_bullets]):
                y = body_top + i * Inches(0.7)
                # 子弹符号
                _add_oval(
                    slide,
                    left=text_left + Inches(0.08),
                    top=y + Inches(0.12),
                    width=Inches(0.15),
                    height=Inches(0.15),
                    fill_color=t["accent"],
                )
                # 要点文本（支持关键词高亮）
                self._add_rich_textbox(
                    slide,
                    left=bullet_left,
                    top=y,
                    width=bullet_width,
                    height=bullet_height,
                    text=bullet,
                    font_name=t["font_body"],
                    size_pt=t["body_size_pt"],
                    color=t["body_text"],
                    theme=t,
                )

            # 结论框（如果有 conclusion 或 key_message）
            conclusion = slide_data.get("conclusion")
            if not conclusion:
                conclusion = slide_data.get("key_message")
            if conclusion:
                concl_top = content_top + content_height - Inches(0.8)
                self._add_conclusion_box(
                    slide,
                    left=text_left + Inches(0.1),
                    top=concl_top,
                    width=text_col_width - Inches(0.2),
                    text=conclusion,
                    theme=t,
                )

            # 右侧图片区：极淡背景 + 无边框
            _add_rounded_rect(
                slide,
                left=img_left - Inches(0.05),
                top=content_top - Inches(0.05),
                width=img_col_width + Inches(0.1),
                height=content_height + Inches(0.1),
                fill_color=RGBColor(0xFA, 0xFA, 0xFA),  # 比左侧再浅一点
                line_color=None,
            )

            # 在右侧区域加载图片（为图注预留底部空间）
            img_area_height_caption = content_height - Inches(0.5)
            self._place_image_in_area(
                slide, slide_data, images_dir,
                img_left, content_top, img_col_width, img_area_height_caption, t
            )

            # 图注放在右侧图片正下方
            caption = slide_data.get("image_caption", "")
            if caption:
                _add_textbox(
                    slide,
                    left=img_left,
                    top=content_top + img_area_height_caption + Inches(0.05),
                    width=img_col_width,
                    height=Inches(0.4),
                    text=caption,
                    font_name=t["font_caption"],
                    size_pt=t["caption_size_pt"],
                    color=t["secondary_text"],
                    italic=True,
                    alignment=PP_ALIGN.CENTER,
                )

        else:
            # ── 无文字时：居中图片布局（原版） ──
            img_margin = Inches(1.0)
            img_area_left = img_margin
            img_area_top = Inches(1.15)
            img_area_width = self.sw - img_margin * 2
            img_area_height = self.sh - Inches(2.5)

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

            self._place_image_in_area(
                slide, slide_data, images_dir,
                img_area_left, img_area_top, img_area_width, img_area_height, t
            )

            # 图注放在图片正下方（居中）
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

    def _place_image_in_area(self, slide, slide_data, images_dir,
                             area_left, area_top, area_width, area_height, t):
        """在指定区域内加载并居中放置图片。"""
        image_path = slide_data.get("image_path", "")
        actual_image_file = None
        if image_path:
            if os.path.isabs(image_path):
                actual_image_file = image_path
            elif images_dir:
                actual_image_file = os.path.join(images_dir, image_path)

        if actual_image_file and os.path.isfile(actual_image_file):
            try:
                from PIL import Image as PILImage
                pil_img = PILImage.open(actual_image_file)
                img_w, img_h = pil_img.size
                aspect = img_w / img_h

                avail_w = area_width - Inches(0.4)
                avail_h = area_height - Inches(0.4)
                if avail_w / avail_h > aspect:
                    final_h = avail_h
                    final_w = int(final_h * aspect)
                else:
                    final_w = avail_w
                    final_h = int(final_w / aspect)

                img_left = area_left + (area_width - final_w) // 2
                img_top = area_top + (area_height - final_h) // 2

                slide.shapes.add_picture(
                    actual_image_file,
                    img_left, img_top,
                    final_w, final_h,
                )
            except ImportError:
                slide.shapes.add_picture(
                    actual_image_file,
                    area_left + Inches(0.2),
                    area_top + Inches(0.2),
                    area_width - Inches(0.4),
                    area_height - Inches(0.4),
                )
            except Exception as e:
                print(f"警告: 无法加载图片 '{actual_image_file}': {e}")
                _add_textbox(
                    slide,
                    left=area_left,
                    top=area_top + area_height // 2 - Inches(0.25),
                    width=area_width,
                    height=Inches(0.5),
                    text="[图片加载失败]",
                    font_name=t["font_body"],
                    size_pt=16,
                    color=t["secondary_text"],
                    alignment=PP_ALIGN.CENTER,
                )
        else:
            _add_rounded_rect(
                slide,
                left=area_left + Inches(0.5),
                top=area_top + Inches(0.5),
                width=area_width - Inches(1.0),
                height=area_height - Inches(1.0),
                fill_color=t["highlight"],
                line_color=t["border_gray"],
                line_width=1,
            )
            _add_textbox(
                slide,
                left=area_left + Inches(1.0),
                top=area_top + area_height // 2 - Inches(0.25),
                width=area_width - Inches(2.0),
                height=Inches(0.5),
                text="Figure Placeholder",
                font_name=t["font_body"],
                size_pt=16,
                color=t["secondary_text"],
                alignment=PP_ALIGN.CENTER,
            )

    def build_conclusion(self, prs, slide_data, page_num):
        """构建结论页。"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        t = self.theme

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = t["background"]

        # 标题带（替代原 _add_top_bar + 单独标题）
        self._add_title_header(slide, slide_data.get("title", "总结"), t,
                               kicker=slide_data.get("kicker"))

        # 核心信息 — 更大的结论框样式
        key_message = slide_data.get("key_message", slide_data.get("title", ""))
        box_w = Inches(10.0)
        box_h = Inches(1.8)
        box_left = (self.sw - box_w) // 2
        box_top = Inches(1.4)

        # 浅蓝色圆角矩形背景
        _add_rounded_rect(
            slide,
            left=box_left,
            top=box_top,
            width=box_w,
            height=box_h,
            fill_color=t["conclusion_bg"],
            line_color=None,
        )
        # 左侧蓝色竖线装饰（更粗）
        _add_rect(
            slide,
            left=box_left,
            top=box_top,
            width=Inches(0.1),
            height=box_h,
            fill_color=t["accent"],
        )

        # 核心信息文字（"结论：" 前缀 + 正文，更大字号，支持关键词高亮）
        kbox = slide.shapes.add_textbox(
            box_left + Inches(0.4),
            box_top + Inches(0.15),
            box_w - Inches(0.7),
            box_h - Inches(0.3),
        )
        ktf = kbox.text_frame
        ktf.word_wrap = True
        try:
            ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception:
            pass
        kp = ktf.paragraphs[0]
        kp.alignment = PP_ALIGN.LEFT
        # "结论：" 前缀（加粗）
        prefix_run = kp.add_run()
        prefix_run.text = "结论："
        _set_font(prefix_run, t["font_title"], 22, bold=True,
                  color=t["header_bar"])
        # 正文（更大字号 + 关键词高亮）
        self._add_keyword_runs(kp, key_message, t["font_title"], 22,
                               t["header_bar"], t, bold=True)

        # 支撑要点
        bullets = slide_data.get("bullets", [])
        y = Inches(3.7)
        for i, bullet in enumerate(bullets):
            # 小装饰条
            _add_rect(
                slide,
                left=Inches(2.5),
                top=y + Inches(0.08),
                width=Inches(0.06),
                height=Inches(0.25),
                fill_color=t["accent"],
            )
            # 要点文本（支持关键词高亮）
            self._add_rich_textbox(
                slide,
                left=Inches(2.8),
                top=y,
                width=Inches(8.0),
                height=Inches(0.45),
                text=bullet,
                font_name=t["font_body"],
                size_pt=t["body_size_pt"],
                color=t["body_text"],
                theme=t,
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

        # 标题带（替代原 _add_top_bar + 单独标题）
        self._add_title_header(slide, slide_data.get("title", "Thank You"), t,
                               kicker=slide_data.get("kicker"))

        # 中央大文字 "Q & A"（使用 key_message 或默认值）
        key_message = slide_data.get("key_message", "Q & A")
        _add_textbox(
            slide,
            left=Inches(2.0),
            top=Inches(1.8),
            width=self.sw - Inches(4.0),
            height=Inches(1.5),
            text=key_message,
            font_name=t["font_title"],
            size_pt=48,
            color=t["title_band"],
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

        # "Thank You" 文字
        _add_textbox(
            slide,
            left=Inches(2.0),
            top=Inches(3.8),
            width=self.sw - Inches(4.0),
            height=Inches(0.8),
            text="Thank You",
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
    "sub_title", "analysis_title", "conclusion", "kicker", "section_no",
}


def _extract_section_title(section):
    """从 sections 数组项中提取标题文字，兼容字符串和字典两种格式。

    - 字符串: "研究背景" → "研究背景"
    - 字典:   {"title": "研究背景", "en": "Background", "desc": "..."} → "研究背景"
    """
    if isinstance(section, str):
        return section
    if isinstance(section, dict):
        return section.get("title", section.get("name", str(section)))
    return str(section)


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
        # 检查 sections 格式（toc 页）
        sections = s.get("sections")
        if sections is not None:
            if not isinstance(sections, list):
                errors.append(f"  幻灯片 #{i+1}: sections 必须是数组")
            else:
                for j, sec in enumerate(sections):
                    # 允许字符串或字典（含 title 键）
                    if not isinstance(sec, (str, dict)):
                        errors.append(
                            f"  幻灯片 #{i+1}: sections[{j}] 必须是字符串或字典, "
                            f"实际类型: {type(sec).__name__}"
                        )
        # 检查 highlights 格式
        highlights = s.get("highlights")
        if highlights is not None:
            if not isinstance(highlights, list):
                errors.append(f"  幻灯片 #{i+1}: highlights 必须是数组")
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
        choices=["academic", "minimal", "trae"],
        help="主题风格 (默认: academic, 可选: minimal, trae)",
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

    # ── 预处理：为 section 页分配章节序号 ──
    section_counter = 0
    for slide_data in slides_data:
        if slide_data.get("page_type") == "section":
            section_counter += 1
            # 如果 slides.json 未提供 section_no，自动分配
            if "section_no" not in slide_data:
                slide_data["section_no"] = section_counter

    # ── 自动补全：确保每页都有足够的字段，效果不依赖AI生成质量 ──
    for slide_data in slides_data:
        page_type = slide_data.get("page_type", "content")
        # content/figure 页：如果没有 conclusion 但有 highlights，提取最后一条作为结论
        if page_type in ("content", "figure"):
            if "conclusion" not in slide_data:
                highlights = slide_data.get("highlights", [])
                if highlights:
                    if isinstance(highlights[-1], dict):
                        slide_data["conclusion"] = highlights[-1].get("content", "")
                    elif isinstance(highlights[-1], str):
                        slide_data["conclusion"] = highlights[-1]
                elif "key_message" in slide_data:
                    slide_data["conclusion"] = slide_data["key_message"]
        # conclusion 页：确保有 key_message
        if page_type == "conclusion" and "key_message" not in slide_data:
            bullets = slide_data.get("bullets", [])
            if bullets:
                slide_data["key_message"] = bullets[0] if isinstance(bullets[0], str) else str(bullets[0])

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
