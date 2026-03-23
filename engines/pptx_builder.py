"""
PPT Forge v2 — PPTX Builder Engine
Assembles professional PPTX files from config + generated art.
Supports 12 slide types, 20+ art styles, 8 themes, 19 palettes.
"""

import io
import random
from typing import List, Dict, Optional, Tuple
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
from PIL import Image, ImageFilter, ImageDraw, ImageEnhance
import numpy as np

from .art_engine import generate_art, get_palette, palette_color, adjust_brightness, with_alpha
from .layout_engine import SmartLayout, TYPE_SCALE
from .effects_pipeline import apply_effects
from .illustration_engine import auto_illustrate_slide, generate_illustration
from .svg_engine import svg_to_png as _svg_to_png


# ─── Theme Definitions ─────────────────────────────────────────────────────

THEMES: Dict[str, Dict] = {
    "minimal": {
        "bg_color": (255, 255, 255),
        "title_color": (25, 25, 28),
        "body_color": (65, 65, 75),
        "accent_color": (0, 90, 200),
        "accent2_color": (80, 140, 240),
        "font_title": "Arial",
        "font_body": "Arial",
        "title_size": 44,
        "body_size": 16,
        "default_art": "gradient_mesh",
        "palette": "mono",
        "overlay_opacity": 0.05,
    },
    "dark": {
        "bg_color": (14, 14, 20),
        "title_color": (245, 245, 250),
        "body_color": (175, 175, 195),
        "accent_color": (99, 179, 237),
        "accent2_color": (56, 161, 240),
        "font_title": "Arial",
        "font_body": "Arial",
        "title_size": 44,
        "body_size": 16,
        "default_art": "flow_field",
        "palette": "midnight",
        "overlay_opacity": 0.12,
    },
    "gradient": {
        "bg_color": (25, 0, 45),
        "title_color": (255, 255, 255),
        "body_color": (220, 200, 240),
        "accent_color": (255, 100, 150),
        "accent2_color": (200, 80, 255),
        "font_title": "Arial",
        "font_body": "Arial",
        "title_size": 44,
        "body_size": 16,
        "default_art": "gradient_mesh",
        "palette": "cosmic",
        "overlay_opacity": 0.10,
    },
    "corporate": {
        "bg_color": (245, 247, 250),
        "title_color": (15, 45, 90),
        "body_color": (55, 75, 105),
        "accent_color": (0, 85, 175),
        "accent2_color": (0, 130, 220),
        "font_title": "Arial",
        "font_body": "Arial",
        "title_size": 40,
        "body_size": 16,
        "default_art": "wave",
        "palette": "cool",
        "overlay_opacity": 0.06,
    },
    "creative": {
        "bg_color": (255, 248, 238),
        "title_color": (35, 15, 5),
        "body_color": (85, 55, 35),
        "accent_color": (255, 90, 70),
        "accent2_color": (255, 140, 50),
        "font_title": "Arial",
        "font_body": "Arial",
        "title_size": 44,
        "body_size": 16,
        "default_art": "geometric",
        "palette": "warm",
        "overlay_opacity": 0.08,
    },
    "tech": {
        "bg_color": (6, 10, 18),
        "title_color": (0, 255, 200),
        "body_color": (140, 195, 215),
        "accent_color": (0, 200, 255),
        "accent2_color": (120, 0, 255),
        "font_title": "Arial",
        "font_body": "Arial",
        "title_size": 44,
        "body_size": 16,
        "default_art": "constellation",
        "palette": "neon",
        "overlay_opacity": 0.10,
    },
    "nature": {
        "bg_color": (242, 238, 228),
        "title_color": (25, 55, 25),
        "body_color": (60, 80, 45),
        "accent_color": (40, 130, 80),
        "accent2_color": (80, 170, 90),
        "font_title": "Arial",
        "font_body": "Arial",
        "title_size": 40,
        "body_size": 16,
        "default_art": "fractal_tree",
        "palette": "forest",
        "overlay_opacity": 0.08,
    },
    "neon": {
        "bg_color": (4, 4, 8),
        "title_color": (255, 0, 128),
        "body_color": (195, 195, 205),
        "accent_color": (0, 255, 255),
        "accent2_color": (255, 0, 200),
        "font_title": "Arial",
        "font_body": "Arial",
        "title_size": 48,
        "body_size": 18,
        "default_art": "bokeh",
        "palette": "neon",
        "overlay_opacity": 0.15,
    },
}

# Slide standard dimensions
SLIDE_W_INCH = 13.333
SLIDE_H_INCH = 7.5
SLIDE_W = Inches(SLIDE_W_INCH)
SLIDE_H = Inches(SLIDE_H_INCH)

# Pixel resolution for art generation (at 96 DPI)
ART_W = int(SLIDE_W_INCH * 96)
ART_H = int(SLIDE_H_INCH * 96)

layout = SmartLayout()


# ─── Helpers ───────────────────────────────────────────────────────────────

def _rgb(r, g, b):
    return RGBColor(r, g, b)


def _rgb_tuple(t):
    return RGBColor(*t[:3])


def _save_img(img, fmt="PNG") -> io.BytesIO:
    buf = io.BytesIO()
    img.save(buf, format=fmt)
    buf.seek(0)
    return buf


def _set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = _rgb_tuple(color)


def _ensure_pt(val):
    """Ensure value is a Pt object for font size."""
    if isinstance(val, Pt):
        return val
    if isinstance(val, (int, float)):
        # EMU values from layout engine are > 100000 (e.g. Pt(55).emu = 5029200)
        # Plain point values are small (e.g. 18, 44, 69)
        if val > 100000:
            # Convert EMU to points: 1 pt = 12700 EMU
            val = int(val / 12700)
        return Pt(int(val))
    return Pt(18)


def _add_text(slide, left, top, width, height, text, font_size=18,
               color=(255, 255, 255), bold=False, alignment=PP_ALIGN.LEFT,
               font_name="Arial", line_spacing=1.25, italic=False):
    """Add a rich text box."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    fs = _ensure_pt(font_size)
    fs_val = fs.pt if isinstance(fs, Pt) else 18

    if isinstance(text, list):
        for i, line in enumerate(text):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = fs
            p.font.color.rgb = _rgb_tuple(color)
            p.font.bold = bold
            p.font.name = font_name
            p.font.italic = italic
            p.alignment = alignment
            p.space_after = Pt(fs_val * 0.6)
            if line_spacing and line_spacing != 1.0:
                p.line_spacing = Pt(fs_val * line_spacing)
    else:
        p = tf.paragraphs[0]
        p.text = str(text)
        p.font.size = fs
        p.font.color.rgb = _rgb_tuple(color)
        p.font.bold = bold
        p.font.name = font_name
        p.font.italic = italic
        p.alignment = alignment

    return txBox


def _add_line(slide, left, top, width, color, height_pt=2):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(height_pt))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb_tuple(color)
    shape.line.fill.background()
    return shape


def _add_accent_bar(slide, color, width_emu=Inches(0.07)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, width_emu, SLIDE_H)
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb_tuple(color)
    shape.line.fill.background()


def _add_circle(slide, cx, cy, r, color, opacity=1.0):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL,
        cx - r, cy - r, r * 2, r * 2)
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb_tuple(color)
    shape.line.fill.background()
    return shape


def _generate_bg_art(slide_config, theme_cfg, art_w=ART_W, art_h=ART_H):
    """Generate background art image."""
    art_style = slide_config.get("bg_style", theme_cfg.get("default_art", "gradient_mesh"))
    palette = slide_config.get("palette", theme_cfg.get("palette", "sunset"))
    seed = slide_config.get("seed", random.randint(1, 999999))

    img = generate_art(art_style, art_w, art_h, palette_name=palette, seed=seed)
    return img, palette


def _apply_bg_art(slide, img, theme_cfg, blur_radius=2, darken=100):
    """Apply art as slide background with optional effects."""
    # Blur for readability
    if blur_radius > 0:
        img = img.filter(ImageFilter.GaussianBlur(radius=blur_radius))

    # Darken overlay
    overlay = Image.new("RGBA", img.size, (0, 0, 0, darken))
    img = img.convert("RGBA")
    img = Image.alpha_composite(img, overlay)
    img = img.convert("RGB")

    buf = _save_img(img)
    slide.shapes.add_picture(buf, 0, 0, SLIDE_W, SLIDE_H)


def _apply_subtle_bg_art(slide, img, theme_cfg, opacity=0.12):
    """Apply art as subtle background texture."""
    img = img.convert("RGBA")
    alpha = img.split()[3] if img.mode == "RGBA" else img.convert("L")
    alpha = alpha.point(lambda p: int(p * opacity))
    img.putalpha(alpha)

    bg = Image.new("RGBA", img.size, theme_cfg["bg_color"] + (255,))
    bg = Image.alpha_composite(bg, img).convert("RGB")
    buf = _save_img(bg)
    slide.shapes.add_picture(buf, 0, 0, SLIDE_W, SLIDE_H)


def _add_illustration(slide, slide_cfg, theme_cfg, pos=None):
    """Auto-generate and add contextual illustration to a slide."""
    if pos is None:
        pos = (Inches(9.0), Inches(1.5), Inches(3.5), Inches(4.5))

    palette_name = slide_cfg.get("palette", theme_cfg.get("palette", "forest"))
    seed = slide_cfg.get("seed", random.randint(1, 999999))

    try:
        svg_str = auto_illustrate_slide(slide_cfg, palette_name=palette_name, seed=seed)
        import tempfile, os
        fd, png_path = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        _svg_to_png(svg_str, 400, 400, png_path)

        if os.path.exists(png_path) and os.path.getsize(png_path) > 0:
            buf = io.BytesIO(open(png_path, "rb").read())
            slide.shapes.add_picture(buf, *pos)
            os.unlink(png_path)
    except Exception:
        pass  # Silently skip if illustration fails


# ─── Slide Builders ───────────────────────────────────────────────────────

def _build_title(prs, cfg, theme):
    """Full-bleed art title slide with overlay text."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])

    art_img, _ = _generate_bg_art(cfg, theme_cfg)
    _apply_bg_art(slide, art_img, theme_cfg, blur_radius=3, darken=120)

    pos = layout.title_position(vertical_center=True)
    heading = cfg.get("heading", "Presentation Title")
    subheading = cfg.get("subheading", "")

    _add_text(slide, *pos["title"], heading,
              font_size=pos["size_title"], color=(255, 255, 255),
              bold=True, alignment=PP_ALIGN.CENTER, font_name=theme_cfg["font_title"])

    _add_line(slide, *pos["divider"], (255, 255, 255), height_pt=2)

    if subheading:
        _add_text(slide, *pos["subtitle"], subheading,
                  font_size=pos["size_subtitle"], color=(210, 210, 230),
                  alignment=PP_ALIGN.CENTER, font_name=theme_cfg["font_body"])


def _build_content(prs, cfg, theme):
    """Content slide with heading, accent bar, and bullets."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])

    # Subtle background art
    art_style = cfg.get("bg_style")
    if art_style:
        art_img, _ = _generate_bg_art(cfg, theme_cfg)
        _apply_subtle_bg_art(slide, art_img, theme_cfg)
    else:
        _set_bg(slide, theme_cfg["bg_color"])

    # Left accent bar
    _add_accent_bar(slide, theme_cfg["accent_color"])

    pos = layout.content_position()
    heading = cfg.get("heading", "Section Title")
    body = cfg.get("body", [])

    _add_text(slide, *pos["heading"], heading,
              font_size=pos["size_heading"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, *pos["divider"], theme_cfg["accent_color"], height_pt=3)

    if isinstance(body, str):
        body = [body]
    _add_text(slide, *pos["body"], body,
              font_size=pos["size_body"], color=theme_cfg["body_color"],
              font_name=theme_cfg["font_body"])

    # Auto-generate contextual illustration
    if not cfg.get("no_illustration", False):
        ill_pos = pos.get("illustration", (Inches(9.0), Inches(1.5), Inches(3.5), Inches(4.5)))
        _add_illustration(slide, cfg, theme_cfg, ill_pos)


def _build_two_column(prs, cfg, theme):
    """Two-column comparison slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])
    _set_bg(slide, theme_cfg["bg_color"])

    pos = layout.two_column_position()
    heading = cfg.get("heading", "Comparison")

    _add_text(slide, *pos["heading"], heading,
              font_size=pos["size_heading"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, Inches(0.8), Inches(1.25), Inches(2), theme_cfg["accent_color"], 3)

    left = cfg.get("left", {})
    right = cfg.get("right", {})

    # Left column
    _add_text(slide, *pos["left_title"], left.get("title", "Column A"),
              font_size=pos["size_col_title"], color=theme_cfg["accent_color"],
              bold=True, font_name=theme_cfg["font_title"])

    items = left.get("items", [])
    if isinstance(items, str):
        items = [items]
    _add_text(slide, *pos["left_body"], items,
              font_size=pos["size_body"], color=theme_cfg["body_color"],
              font_name=theme_cfg["font_body"])

    # Vertical divider (4-tuple: left, top, width, height)
    div_key = "divider_v" if "divider_v" in pos else "divider"
    div = pos[div_key]
    if len(div) == 4:
        div_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *div)
        div_shape.fill.solid()
        div_shape.fill.fore_color.rgb = _rgb_tuple(theme_cfg["accent_color"])
        div_shape.line.fill.background()
    else:
        _add_line(slide, *div, theme_cfg["accent_color"], height_pt=1)

    # Right column
    _add_text(slide, *pos["right_title"], right.get("title", "Column B"),
              font_size=pos["size_col_title"], color=theme_cfg["accent_color"],
              bold=True, font_name=theme_cfg["font_title"])

    items = right.get("items", [])
    if isinstance(items, str):
        items = [items]
    _add_text(slide, *pos["right_body"], items,
              font_size=pos["size_body"], color=theme_cfg["body_color"],
              font_name=theme_cfg["font_body"])

    # Auto-generate contextual illustrations for columns
    if not cfg.get("no_illustration", False):
        # Small illustration in bottom-right corner
        _add_illustration(slide, cfg, theme_cfg,
                         (Inches(10.5), Inches(5.5), Inches(2.2), Inches(1.5)))


def _build_three_column(prs, cfg, theme):
    """Three-column layout slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])
    _set_bg(slide, theme_cfg["bg_color"])

    pos = layout.three_column_position()
    heading = cfg.get("heading", "Overview")

    _add_text(slide, *pos["heading"], heading,
              font_size=pos["size_heading"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, Inches(0.8), Inches(1.25), Inches(2), theme_cfg["accent_color"], 3)

    columns = cfg.get("columns", [{}, {}, {}])
    for i, col_cfg in enumerate(columns[:3]):
        col_pos = pos["columns"][i]
        _add_text(slide, *col_pos["title"], col_cfg.get("title", f"Column {i+1}"),
                  font_size=pos["size_col_title"], color=theme_cfg["accent_color"],
                  bold=True, font_name=theme_cfg["font_title"])

        items = col_cfg.get("items", [])
        if isinstance(items, str):
            items = [items]
        _add_text(slide, *col_pos["body"], items,
                  font_size=pos["size_body"], color=theme_cfg["body_color"],
                  font_name=theme_cfg["font_body"])

    # Auto illustration in top-right
    if not cfg.get("no_illustration", False):
        _add_illustration(slide, cfg, theme_cfg,
                         (Inches(10.5), Inches(0.2), Inches(2.2), Inches(1.5)))


def _build_image_full(prs, cfg, theme):
    """Full-bleed art image slide with overlay text."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])

    art_img, _ = _generate_bg_art(cfg, theme_cfg, ART_W, ART_H)

    # Apply optional effects
    effects = cfg.get("effects", [])
    if effects:
        art_img = apply_effects(art_img, effects)

    # Gradient overlay at bottom for text readability
    w, h = art_img.size
    overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    for y in range(int(h * 0.55), h):
        alpha = int(180 * ((y - h * 0.55) / (h * 0.45)))
        draw.line([(0, y), (w, y)], fill=(0, 0, 0, min(180, alpha)))

    art_img = art_img.convert("RGBA")
    art_img = Image.alpha_composite(art_img, overlay).convert("RGB")

    buf = _save_img(art_img)
    slide.shapes.add_picture(buf, 0, 0, SLIDE_W, SLIDE_H)

    pos = layout.image_full_position()
    heading = cfg.get("heading", "")
    caption = cfg.get("caption", "")

    if heading:
        _add_text(slide, *pos["heading"], heading,
                  font_size=pos["size_heading"], color=(255, 255, 255),
                  bold=True, alignment=PP_ALIGN.CENTER, font_name=theme_cfg["font_title"])
    if caption:
        _add_text(slide, *pos["caption"], caption,
                  font_size=pos["size_caption"], color=(200, 200, 210),
                  alignment=PP_ALIGN.CENTER, italic=True)


def _build_chart(prs, cfg, theme):
    """Data chart slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])
    _set_bg(slide, theme_cfg["bg_color"])

    pos = layout.chart_position()
    heading = cfg.get("heading", "Data")

    _add_text(slide, *pos["heading"], heading,
              font_size=pos["size_heading"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, Inches(0.8), Inches(1.25), Inches(2), theme_cfg["accent_color"], 3)

    chart_type_str = cfg.get("chart_type", "bar")
    chart_type_map = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE,
        "area": XL_CHART_TYPE.AREA,
        "pie": XL_CHART_TYPE.PIE,
        "donut": XL_CHART_TYPE.DOUGHNUT,
        "scatter": XL_CHART_TYPE.XY_SCATTER,
    }

    data = cfg.get("data", {"labels": ["A", "B", "C"], "values": [30, 50, 20]})
    labels = data.get("labels", [])
    values = data.get("values", [])

    # Build chart data
    from pptx.chart.data import CategoryChartData
    chart_data = CategoryChartData()
    chart_data.categories = labels
    chart_data.add_series(cfg.get("series_name", "Series 1"), values)

    # Additional series support
    series_list = data.get("series", [])
    for s in series_list:
        if isinstance(s, dict):
            chart_data.add_series(s.get("name", ""), s.get("values", []))

    chart_shape = slide.shapes.add_chart(
        chart_type_map.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED),
        *pos["chart"], chart_data
    )

    chart = chart_shape.chart
    chart.has_legend = cfg.get("show_legend", True)

    # Style the chart
    plot = chart.plots[0]
    plot.gap_width = 100

    # Color the series
    palette_name = cfg.get("palette", theme_cfg.get("palette", "sunset"))
    palette = get_palette(palette_name)
    for i, series in enumerate(plot.series):
        color = palette[i % len(palette)]
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = _rgb_tuple(color)


def _build_end(prs, cfg, theme):
    """Thank you / closing slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])

    art_img, _ = _generate_bg_art(cfg, theme_cfg)
    _apply_bg_art(slide, art_img, theme_cfg, blur_radius=3, darken=120)

    pos = layout.end_position()
    heading = cfg.get("heading", "Thank You")
    subheading = cfg.get("subheading", "")

    # Support both key names (title or heading)
    title_key = "title" if "title" in pos else "heading"
    sub_key = "subtitle" if "subtitle" in pos else "subheading"

    _add_text(slide, *pos[title_key], heading,
              font_size=pos["size_title"] if "size_title" in pos else pos["size_heading"],
              color=(255, 255, 255),
              bold=True, alignment=PP_ALIGN.CENTER, font_name=theme_cfg["font_title"])

    _add_line(slide, *pos["divider"], (255, 255, 255), height_pt=2)

    if subheading:
        _add_text(slide, *pos[sub_key], subheading,
                  font_size=pos["size_subtitle"] if "size_subtitle" in pos else pos.get("size_subheading", 22),
                  color=(200, 200, 220),
                  alignment=PP_ALIGN.CENTER)


def _build_agenda(prs, cfg, theme):
    """Agenda / outline slide with numbered items."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])
    _set_bg(slide, theme_cfg["bg_color"])

    pos = layout.agenda_position()
    heading = cfg.get("heading", "Agenda")

    _add_text(slide, *pos["heading"], heading,
              font_size=pos["size_heading"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, Inches(0.8), Inches(1.25), Inches(2), theme_cfg["accent_color"], 3)

    items = cfg.get("items", [])
    palette = get_palette(cfg.get("palette", theme_cfg.get("palette", "sunset")))

    for i, item in enumerate(items):
        y = Inches(1.8) + Inches(i * 0.75)
        number_color = palette[(i + 1) % len(palette)]

        # Number circle
        _add_circle(slide, Inches(1.5), y + Inches(0.18), Inches(0.22), number_color)

        # Number text
        _add_text(slide, Inches(1.2), y, Inches(0.6), Inches(0.5),
                  str(i + 1), font_size=pos["size_number"],
                  color=(255, 255, 255), bold=True, alignment=PP_ALIGN.CENTER)

        # Item text
        _add_text(slide, Inches(2.1), y, Inches(9), Inches(0.5),
                  item, font_size=pos["size_item"],
                  color=theme_cfg["body_color"], font_name=theme_cfg["font_body"])

    # Auto illustration
    if not cfg.get("no_illustration", False):
        _add_illustration(slide, cfg, theme_cfg,
                         (Inches(10.2), Inches(5.5), Inches(2.5), Inches(1.5)))


def _build_timeline(prs, cfg, theme):
    """Timeline / process flow slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])
    _set_bg(slide, theme_cfg["bg_color"])

    pos = layout.timeline_position()
    heading = cfg.get("heading", "Timeline")

    _add_text(slide, *pos["heading"], heading,
              font_size=pos["size_heading"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, Inches(0.8), Inches(1.15), Inches(2), theme_cfg["accent_color"], 3)

    items = cfg.get("items", [])
    palette = get_palette(cfg.get("palette", theme_cfg.get("palette", "sunset")))
    n = len(items)

    if n > 0:
        line_y = Inches(4.0)
        start_x = Inches(1.5)
        end_x = Inches(11.8)

        # Horizontal line
        _add_line(slide, start_x, line_y, end_x - start_x, theme_cfg["body_color"], 2)

        # Items along timeline
        for i, item in enumerate(items):
            t = i / max(1, n - 1)
            x = start_x + (end_x - start_x) * t

            # Dot
            dot_color = palette[(i + 1) % len(palette)]
            _add_circle(slide, x, line_y - Inches(0.1), Inches(0.15), dot_color)

            # Date/label above
            label = item.get("date", item.get("label", f"Step {i+1}"))
            _add_text(slide, x - Inches(0.8), line_y - Inches(0.8), Inches(1.6), Inches(0.5),
                      label, font_size=pos["size_label"],
                      color=theme_cfg["accent_color"], bold=True, alignment=PP_ALIGN.CENTER)

            # Title below
            title = item.get("title", "")
            if title:
                _add_text(slide, x - Inches(0.9), line_y + Inches(0.3), Inches(1.8), Inches(0.8),
                          title, font_size=pos["size_title"],
                          color=theme_cfg["body_color"], alignment=PP_ALIGN.CENTER)

    # Auto illustration
    if not cfg.get("no_illustration", False):
        _add_illustration(slide, cfg, theme_cfg,
                         (Inches(0.3), Inches(5.5), Inches(2.5), Inches(1.5)))


def _build_quote(prs, cfg, theme):
    """Quote / testimonial slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])

    art_style = cfg.get("bg_style")
    if art_style:
        art_img, _ = _generate_bg_art(cfg, theme_cfg)
        _apply_subtle_bg_art(slide, art_img, theme_cfg, opacity=0.08)
    else:
        _set_bg(slide, theme_cfg["bg_color"])

    pos = layout.quote_position()

    # Large quote marks
    _add_text(slide, Inches(0.5), Inches(1.0), Inches(1.5), Inches(1.5),
              "\u201C", font_size=120, color=theme_cfg["accent_color"],
              alignment=PP_ALIGN.LEFT)

    quote_text = cfg.get("text", "Your quote here")
    author = cfg.get("author", "")

    _add_text(slide, *pos["quote"], quote_text,
              font_size=pos["size_quote"], color=theme_cfg["title_color"],
              bold=False, alignment=PP_ALIGN.LEFT, italic=True,
              font_name=theme_cfg["font_title"])

    _add_line(slide, *pos["divider"], theme_cfg["accent_color"], 2)

    if author:
        _add_text(slide, *pos["author"], f"— {author}",
                  font_size=pos["size_author"], color=theme_cfg["body_color"],
                  alignment=PP_ALIGN.CENTER)

    # Auto illustration
    if not cfg.get("no_illustration", False):
        _add_illustration(slide, cfg, theme_cfg,
                         (Inches(10.5), Inches(5.8), Inches(2.0), Inches(1.3)))


def _build_data_table(prs, cfg, theme):
    """Data table slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])
    _set_bg(slide, theme_cfg["bg_color"])

    pos = layout.data_table_position()
    heading = cfg.get("heading", "Data Table")

    _add_text(slide, *pos["heading"], heading,
              font_size=pos["size_heading"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, Inches(0.8), Inches(1.25), Inches(2), theme_cfg["accent_color"], 3)

    headers = cfg.get("headers", [])
    rows = cfg.get("rows", [])

    if headers and rows:
        n_rows = len(rows) + 1
        n_cols = len(headers)
        tbl_shape = slide.shapes.add_table(n_rows, n_cols, *pos["table"])
        table = tbl_shape.table

        # Header row
        palette = get_palette(cfg.get("palette", theme_cfg.get("palette", "sunset")))
        for j, h in enumerate(headers):
            cell = table.cell(0, j)
            cell.text = h
            p = cell.text_frame.paragraphs[0]
            p.font.size = _ensure_pt(pos["size_header"])
            p.font.bold = True
            p.font.color.rgb = _rgb(255, 255, 255)
            p.font.name = theme_cfg["font_title"]
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb_tuple(palette[1] if len(palette) > 1 else theme_cfg["accent_color"])

        # Data rows
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                cell = table.cell(i + 1, j)
                cell.text = str(val)
                p = cell.text_frame.paragraphs[0]
                p.font.size = _ensure_pt(pos["size_cell"])
                p.font.color.rgb = _rgb_tuple(theme_cfg["body_color"])
                p.font.name = theme_cfg["font_body"]

                # Alternating row colors
                if i % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = _rgb_tuple(
                        adjust_brightness(theme_cfg["bg_color"], 0.95)
                    )

    # Auto illustration
    if not cfg.get("no_illustration", False):
        _add_illustration(slide, cfg, theme_cfg,
                         (Inches(10.8), Inches(6.0), Inches(2.0), Inches(1.2)))


def _build_metrics(prs, cfg, theme):
    """Big number / KPI metrics slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])
    _set_bg(slide, theme_cfg["bg_color"])

    pos = layout.metrics_position()
    heading = cfg.get("heading", "Key Metrics")

    _add_text(slide, *pos["heading"], heading,
              font_size=pos["size_heading"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, Inches(0.8), Inches(1.1), Inches(2), theme_cfg["accent_color"], 3)

    metrics = cfg.get("metrics", [])
    n = len(metrics)
    palette = get_palette(cfg.get("palette", theme_cfg.get("palette", "sunset")))

    col_width = Inches(11.3 / max(1, n))
    start_x = Inches(1.0)

    for i, m in enumerate(metrics):
        cx = start_x + col_width * i + col_width / 2

        # Large number
        number_color = palette[(i + 1) % len(palette)]
        _add_text(slide, cx - Inches(1.5), Inches(2.5), Inches(3), Inches(2),
                  m.get("value", "0"),
                  font_size=pos["size_number"], color=number_color,
                  bold=True, alignment=PP_ALIGN.CENTER,
                  font_name=theme_cfg["font_title"])

        # Label
        _add_text(slide, cx - Inches(1.5), Inches(4.8), Inches(3), Inches(0.8),
                  m.get("label", ""),
                  font_size=pos["size_label"], color=theme_cfg["body_color"],
                  alignment=PP_ALIGN.CENTER, font_name=theme_cfg["font_body"])

        # Sub-label
        sub = m.get("sublabel", "")
        if sub:
            _add_text(slide, cx - Inches(1.5), Inches(5.5), Inches(3), Inches(0.6),
                      sub, font_size=TYPE_SCALE["sm"], color=theme_cfg["body_color"],
                      alignment=PP_ALIGN.CENTER, italic=True)

        # Decorative card background
        if "card_bg" in pos:
            card = pos["cards"][i]["card_bg"] if i < len(pos.get("cards", [])) else None
            if card:
                bg_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *card)
                bg_shape.fill.solid()
                bg_shape.fill.fore_color.rgb = _rgb_tuple(
                    adjust_brightness(theme_cfg["bg_color"], 1.15)
                )
                bg_shape.line.fill.background()

    # Auto illustration
    if not cfg.get("no_illustration", False):
        _add_illustration(slide, cfg, theme_cfg,
                         (Inches(10.8), Inches(6.0), Inches(2.0), Inches(1.2)))


def _build_image_grid(prs, cfg, theme):
    """Image grid layout (2x2 or 3x2)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme_cfg = THEMES.get(theme, THEMES["dark"])
    _set_bg(slide, theme_cfg["bg_color"])

    heading = cfg.get("heading", "Gallery")
    _add_text(slide, Inches(0.8), Inches(0.3), Inches(11), Inches(0.7),
              heading, font_size=TYPE_SCALE["2xl"], color=theme_cfg["title_color"],
              bold=True, font_name=theme_cfg["font_title"])

    _add_line(slide, Inches(0.8), Inches(1.05), Inches(2), theme_cfg["accent_color"], 3)

    art_styles = cfg.get("art_styles", ["gradient_mesh", "flow_field", "voronoi", "bokeh"])
    palette = cfg.get("palette", theme_cfg.get("palette", "sunset"))
    seed = cfg.get("seed", random.randint(1, 999999))

    cols = min(3, len(art_styles))
    rows = (len(art_styles) + cols - 1) // cols
    img_w = int(10.5 / cols * 96)
    img_h = int(5.0 / rows * 96)

    for i, style in enumerate(art_styles[:6]):
        col = i % cols
        row = i // cols
        x = Inches(1.0 + col * (10.5 / cols) + 0.15)
        y = Inches(1.4 + row * (5.0 / rows) + 0.15)
        w = Inches(10.5 / cols - 0.3)
        h = Inches(5.0 / rows - 0.3)

        art = generate_art(style, img_w, img_h, palette_name=palette, seed=seed + i)
        # Rounded corner effect via slight blur on edges
        art = art.filter(ImageFilter.GaussianBlur(radius=1))
        buf = _save_img(art)
        slide.shapes.add_picture(buf, x, y, w, h)


# ─── Slide Type Registry ──────────────────────────────────────────────────

SLIDE_BUILDERS = {
    "title": _build_title,
    "content": _build_content,
    "two_column": _build_two_column,
    "three_column": _build_three_column,
    "image_full": _build_image_full,
    "chart": _build_chart,
    "end": _build_end,
    "agenda": _build_agenda,
    "timeline": _build_timeline,
    "quote": _build_quote,
    "data_table": _build_data_table,
    "metrics": _build_metrics,
    "image_grid": _build_image_grid,
}


# ─── Main Builder ─────────────────────────────────────────────────────────

def build_ppt(config: Dict, output_path: str) -> str:
    """Build a complete PPTX from config dictionary."""
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    theme = config.get("theme", "dark")
    slides_config = config.get("slides", [])
    title = config.get("title", "Presentation")

    # Set presentation title
    prs.core_properties.title = title

    for slide_cfg in slides_config:
        slide_type = slide_cfg.get("type", "content")
        builder = SLIDE_BUILDERS.get(slide_type, _build_content)

        # Auto-seed if not provided
        if "seed" not in slide_cfg:
            slide_cfg["seed"] = random.randint(1, 999999)

        builder(prs, slide_cfg, theme)

    prs.save(output_path)
    return output_path
