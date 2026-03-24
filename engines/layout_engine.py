"""
PPT Forge v2 — Enhanced Smart Layout Engine
Golden ratio grid system, modular type scale, adaptive compositions.
Supports Chinese typography, educational layouts, and rich visual hierarchies.
"""

from typing import List, Tuple, Dict, Optional
from pptx.util import Inches, Pt, Emu


# ─── Design System Constants ──────────────────────────────────────────────

# 8px base grid with golden ratio spacing
SPACING = {
    "xxs": Inches(0.04),    # ~4px  micro spacing
    "xs":  Inches(0.08),    # ~8px  tight
    "sm":  Inches(0.16),    # ~16px small
    "md":  Inches(0.32),    # ~32px medium
    "lg":  Inches(0.56),    # ~56px large
    "xl":  Inches(0.8),     # ~80px x-large
    "2xl": Inches(1.2),     # ~120px xx-large
    "3xl": Inches(1.6),     # ~160px huge
    "4xl": Inches(2.4),     # ~240px massive
}

# Modular type scale (major third 1.25 ratio) — ENHANCED for Chinese
TYPE_SCALE = {
    "xs":    10,
    "sm":    12,
    "base":  14,
    "md":    18,
    "lg":    22,
    "xl":    28,
    "2xl":   35,
    "3xl":   44,
    "4xl":   55,
    "5xl":   69,
    "hero":  86,
    # Chinese education scale (much larger)
    "cn_title":    48,     # Section titles
    "cn_heading":  36,     # Content headings
    "cn_body":     24,     # Body text (larger for classroom PPTs)
    "cn_subtitle": 28,     # Subtitles
    "cn_huge":     72,     # Task separators
    "cn_massive":  120,    # Hero Chinese titles
    "cn_giant":    166,    # Maximum impact titles
    "cn_label":    20,     # Labels and annotations
    "cn_caption":  16,     # Small captions
}

# Line height ratios per size category
LINE_HEIGHT = {
    "display": 1.1,   # Large headings
    "heading": 1.2,   # Section headings
    "body":    1.5,   # Body text
    "caption": 1.4,   # Captions/labels
    "mono":    1.3,   # Code/data
}

# Letter spacing (tracking) multipliers
TRACKING = {
    "tight":  -0.02,
    "normal": 0.0,
    "wide":   0.05,
    "wider":  0.1,
}


class SmartLayout:
    """Professional slide layout engine with golden ratio grid."""

    # 16:9 standard
    SLIDE_W = 13.333
    SLIDE_H = 7.5

    def __init__(self, safe_margin: float = 0.5):
        self.margin = safe_margin
        self.content_w = self.SLIDE_W - safe_margin * 2
        self.content_h = self.SLIDE_H - safe_margin * 2

    @property
    def slide_width(self):
        return Inches(self.SLIDE_W)

    @property
    def slide_height(self):
        return Inches(self.SLIDE_H)

    def inches(self, v: float):
        return Inches(v)

    def pt(self, v: float):
        return Pt(v)

    # ─── Grid System ──────────────────────────────────────────────────

    def col(self, n: int, total: int = 12, gap: float = 0.15) -> float:
        """Calculate column width (12-column grid)."""
        gutter = gap * (total - 1)
        return (self.content_w - gutter) * n / total

    def col_offset(self, start: int, total: int = 12, gap: float = 0.15) -> float:
        """Calculate column offset."""
        col_w = self.col(1, total, gap)
        return self.margin + (start - 1) * (col_w + gap)

    def golden_split(self, total: float, ratio: float = 0.618) -> Tuple[float, float]:
        """Split a dimension by golden ratio."""
        return total * ratio, total * (1 - ratio)

    # ─── Title Slide Layouts ──────────────────────────────────────────

    def title_centered(self) -> Dict:
        """Centered title layout (most common)."""
        return {
            "title":    (Inches(1.2), Inches(2.2), Inches(10.9), Inches(1.8)),
            "divider":  (Inches(5.2), Inches(4.1), Inches(2.9)),
            "subtitle": (Inches(2.5), Inches(4.4), Inches(8.3), Inches(1.2)),
            "size_title":    Pt(TYPE_SCALE["4xl"]),
            "size_subtitle": Pt(TYPE_SCALE["lg"]),
            "title_bold":    True,
            "subtitle_bold": False,
        }

    def title_left_aligned(self) -> Dict:
        """Left-aligned title with accent bar."""
        return {
            "accent_bar": (Inches(0.6), Inches(1.5), Inches(0.07), Inches(4.5)),
            "title":      (Inches(1.0), Inches(1.8), Inches(10), Inches(1.5)),
            "divider":    (Inches(1.0), Inches(3.4), Inches(2.5)),
            "subtitle":   (Inches(1.0), Inches(3.7), Inches(9), Inches(1)),
            "size_title":    Pt(TYPE_SCALE["3xl"]),
            "size_subtitle": Pt(TYPE_SCALE["md"]),
        }

    def title_hero(self) -> Dict:
        """Large hero title — maximum impact."""
        return {
            "title":    (Inches(0.8), Inches(1.5), Inches(11.7), Inches(2.5)),
            "subtitle": (Inches(0.8), Inches(4.2), Inches(11.7), Inches(1.2)),
            "divider":  (Inches(0.8), Inches(4.0), Inches(4)),
            "size_title":    Pt(TYPE_SCALE["hero"]),
            "size_subtitle": Pt(TYPE_SCALE["xl"]),
        }

    # ─── Chinese Education Layouts ────────────────────────────────────

    def cn_title_slide(self) -> Dict:
        """Chinese educational title — large characters with decorative elements."""
        return {
            "title_area":   (Inches(2), Inches(1.5), Inches(9.3), Inches(3)),
            "subtitle":     (Inches(3), Inches(4.8), Inches(7.3), Inches(1)),
            "size_title":   Pt(TYPE_SCALE["cn_massive"]),
            "size_subtitle": Pt(TYPE_SCALE["cn_subtitle"]),
            "decoration_tl": (Inches(0.3), Inches(0.3), Inches(2.5), Inches(1.5)),
            "decoration_br": (Inches(10), Inches(5.5), Inches(3), Inches(1.8)),
        }

    def cn_task_separator(self) -> Dict:
        """Full-page task/section separator — Chinese education style."""
        return {
            "text_area":    (Inches(2), Inches(2), Inches(9.3), Inches(3.5)),
            "size_text":    Pt(TYPE_SCALE["cn_huge"]),
            "decoration_top": (Inches(1.5), Inches(0.5), Inches(10.3), Inches(1.5)),
            "decoration_bot": (Inches(1.5), Inches(5.5), Inches(10.3), Inches(1.5)),
        }

    def cn_content_with_side(self) -> Dict:
        """Chinese educational content with vertical side label."""
        return {
            "side_label":   (Inches(0.3), Inches(1), Inches(1.2), Inches(5.5)),
            "content_area": (Inches(1.8), Inches(0.8), Inches(7.5), Inches(6)),
            "decoration":   (Inches(9.5), Inches(4.5), Inches(3.5), Inches(2.8)),
            "heading":      (Inches(2), Inches(1), Inches(7), Inches(0.8)),
            "body":         (Inches(2), Inches(2), Inches(7), Inches(4.5)),
            "size_side":    Pt(TYPE_SCALE["cn_title"]),
            "size_heading": Pt(TYPE_SCALE["cn_heading"]),
            "size_body":    Pt(TYPE_SCALE["cn_body"]),
        }

    def cn_scenario(self) -> Dict:
        """Chinese education scenario layout — situation + thought bubble."""
        return {
            "scenario_label": (Inches(0.2), Inches(0), Inches(2.5), Inches(1.2)),
            "scenario_title": (Inches(0.4), Inches(0.2), Inches(2), Inches(0.7)),
            "thought_bubble": (Inches(0.2), Inches(4.5), Inches(4.3), Inches(2.5)),
            "content_box":    (Inches(4.8), Inches(0.8), Inches(8.3), Inches(6.2)),
            "content_text":   (Inches(5.2), Inches(1.8), Inches(7.5), Inches(2.5)),
            "content_image":  (Inches(6), Inches(3.5), Inches(5.5), Inches(3.5)),
            "size_label":   Pt(TYPE_SCALE["cn_heading"]),
            "size_content": Pt(TYPE_SCALE["cn_body"]),
            "size_bubble":  Pt(TYPE_SCALE["cn_label"]),
        }

    def cn_mind_map(self) -> Dict:
        """Mind map layout for knowledge organization."""
        return {
            "heading":     (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "map_area":    (Inches(1), Inches(1.2), Inches(11.3), Inches(6)),
            "center_node": (Inches(5.5), Inches(3.8)),
            "decoration":  (Inches(10), Inches(5.5), Inches(3), Inches(1.8)),
            "size_heading": Pt(TYPE_SCALE["cn_heading"]),
        }

    def cn_evaluation_table(self) -> Dict:
        """Evaluation/rubric table layout."""
        return {
            "heading":  (Inches(1.5), Inches(0.5), Inches(10), Inches(0.8)),
            "table":    (Inches(0.8), Inches(1.5), Inches(11.5), Inches(5)),
            "decoration": (Inches(0.2), Inches(4.5), Inches(1), Inches(2.5)),
            "size_heading": Pt(TYPE_SCALE["cn_heading"]),
            "size_header":  Pt(TYPE_SCALE["cn_label"]),
            "size_cell":    Pt(TYPE_SCALE["cn_caption"]),
        }

    def cn_flowchart(self) -> Dict:
        """Process flow / writing guide layout."""
        return {
            "heading":    (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "flow_area":  (Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.8)),
            "decoration": (Inches(9.8), Inches(5), Inches(3.2), Inches(2.2)),
            "size_heading": Pt(TYPE_SCALE["cn_heading"]),
        }

    def cn_split_content(self) -> Dict:
        """Split layout — left illustration, right content with annotations."""
        return {
            "left_image":   (Inches(0.3), Inches(0.5), Inches(5.5), Inches(6.5)),
            "right_content": (Inches(6.2), Inches(0.5), Inches(6.8), Inches(6.5)),
            "heading":      (Inches(6.5), Inches(0.8), Inches(6), Inches(0.8)),
            "body":         (Inches(6.5), Inches(1.8), Inches(6), Inches(4)),
            "annotation":   (Inches(6.5), Inches(5.5), Inches(6), Inches(1.2)),
            "size_heading": Pt(TYPE_SCALE["cn_heading"]),
            "size_body":    Pt(TYPE_SCALE["cn_body"]),
            "size_annot":   Pt(TYPE_SCALE["cn_label"]),
        }

    def cn_reading_passage(self) -> Dict:
        """Reading passage with analysis sidebar."""
        return {
            "title_bar":   (Inches(1.5), Inches(0.5), Inches(6), Inches(0.7)),
            "passage":     (Inches(1.5), Inches(1.5), Inches(7), Inches(5.5)),
            "sidebar":     (Inches(9), Inches(1), Inches(3.8), Inches(5.5)),
            "decoration":  (Inches(9.2), Inches(5), Inches(3.5), Inches(2.2)),
            "size_title":   Pt(TYPE_SCALE["cn_heading"]),
            "size_body":    Pt(TYPE_SCALE["cn_body"]),
            "size_sidebar": Pt(TYPE_SCALE["cn_label"]),
        }

    # ─── Content Slide Layouts ────────────────────────────────────────

    def content_standard(self) -> Dict:
        """Standard content with heading + body."""
        return {
            "accent":   (Inches(0.5), Inches(0.3), Inches(0.06), Inches(6.9)),
            "heading":  (Inches(0.8), Inches(0.4), Inches(11), Inches(0.9)),
            "divider":  (Inches(0.8), Inches(1.35), Inches(2.2)),
            "body":     (Inches(0.8), Inches(1.7), Inches(11), Inches(5.2)),
            "illustration": (Inches(9.0), Inches(1.5), Inches(3.5), Inches(4.5)),
            "size_heading": Pt(TYPE_SCALE["2xl"]),
            "size_body":    Pt(TYPE_SCALE["base"]),
        }

    def content_with_image(self, side: str = "right") -> Dict:
        """Content with illustration on one side."""
        text_w = self.content_w * 0.55
        img_w = self.content_w * 0.38
        if side == "left":
            return {
                "accent":        (Inches(img_w + 1.0), Inches(0.3), Inches(0.06), Inches(6.9)),
                "heading":       (Inches(img_w + 1.3), Inches(0.4), Inches(text_w), Inches(0.8)),
                "divider":       (Inches(img_w + 1.3), Inches(1.25), Inches(2)),
                "body":          (Inches(img_w + 1.3), Inches(1.6), Inches(text_w), Inches(5)),
                "illustration":  (Inches(0.5), Inches(1.2), Inches(img_w), Inches(5)),
                "size_heading": Pt(TYPE_SCALE["xl"]),
                "size_body":    Pt(TYPE_SCALE["base"]),
            }
        return {
            "accent":        (Inches(0.5), Inches(0.3), Inches(0.06), Inches(6.9)),
            "heading":       (Inches(0.8), Inches(0.4), Inches(text_w), Inches(0.8)),
            "divider":       (Inches(0.8), Inches(1.25), Inches(2)),
            "body":          (Inches(0.8), Inches(1.6), Inches(text_w), Inches(5)),
            "illustration":  (Inches(text_w + 1.0), Inches(1.2), Inches(img_w), Inches(5)),
            "size_heading": Pt(TYPE_SCALE["xl"]),
            "size_body":    Pt(TYPE_SCALE["base"]),
        }

    # ─── Multi-column Layouts ─────────────────────────────────────────

    def two_column(self, gap: float = 0.5) -> Dict:
        """Two-column comparison."""
        col_w = (self.content_w - gap) / 2
        left_x = self.margin
        right_x = self.margin + col_w + gap
        return {
            "heading":      (Inches(0.8), Inches(0.35), Inches(11), Inches(0.8)),
            "divider_h":    (Inches(0.8), Inches(1.2), Inches(2.2)),
            "left_title":   (Inches(left_x), Inches(1.5), Inches(col_w), Inches(0.6)),
            "left_body":    (Inches(left_x), Inches(2.2), Inches(col_w), Inches(4.5)),
            "left_icon":    (Inches(left_x), Inches(1.3), Inches(0.5), Inches(0.5)),
            "divider_v":    (Inches(self.margin + col_w + gap/2), Inches(1.6), Inches(0.04), Inches(4.5)),
            "right_title":  (Inches(right_x), Inches(1.5), Inches(col_w), Inches(0.6)),
            "right_body":   (Inches(right_x), Inches(2.2), Inches(col_w), Inches(4.5)),
            "right_icon":   (Inches(right_x), Inches(1.3), Inches(0.5), Inches(0.5)),
            "size_heading":     Pt(TYPE_SCALE["2xl"]),
            "size_col_title":   Pt(TYPE_SCALE["lg"]),
            "size_body":        Pt(TYPE_SCALE["base"]),
        }

    def three_column(self, gap: float = 0.35) -> Dict:
        """Three-column layout."""
        col_w = (self.content_w - gap * 2) / 3
        cols = []
        for i in range(3):
            left = self.margin + i * (col_w + gap)
            cols.append({
                "icon":  (Inches(left), Inches(1.3), Inches(0.5), Inches(0.5)),
                "title": (Inches(left), Inches(1.8), Inches(col_w), Inches(0.5)),
                "body":  (Inches(left), Inches(2.4), Inches(col_w), Inches(4.2)),
            })
        return {
            "heading":  (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "divider":  (Inches(0.8), Inches(1.05), Inches(2.2)),
            "columns":  cols,
            "size_heading":     Pt(TYPE_SCALE["xl"]),
            "size_col_title":   Pt(TYPE_SCALE["md"]),
            "size_body":        Pt(TYPE_SCALE["sm"]),
        }

    def four_column(self, gap: float = 0.25) -> Dict:
        """Four-column layout."""
        col_w = (self.content_w - gap * 3) / 4
        cols = []
        for i in range(4):
            left = self.margin + i * (col_w + gap)
            cols.append({
                "number": (Inches(left), Inches(1.3), Inches(col_w), Inches(0.8)),
                "title":  (Inches(left), Inches(2.3), Inches(col_w), Inches(0.5)),
                "body":   (Inches(left), Inches(2.9), Inches(col_w), Inches(3.8)),
            })
        return {
            "heading":  (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "divider":  (Inches(0.8), Inches(1.05), Inches(2.2)),
            "columns":  cols,
            "size_heading":    Pt(TYPE_SCALE["xl"]),
            "size_number":     Pt(TYPE_SCALE["3xl"]),
            "size_col_title":  Pt(TYPE_SCALE["base"]),
            "size_body":       Pt(TYPE_SCALE["sm"]),
        }

    # ─── Special Slide Layouts ────────────────────────────────────────

    def quote_layout(self) -> Dict:
        """Quote / testimonial."""
        return {
            "mark_open":  (Inches(0.8), Inches(1.0), Inches(1.2), Inches(1.5)),
            "quote":      (Inches(1.5), Inches(2.0), Inches(10.3), Inches(2.5)),
            "divider":    (Inches(5.0), Inches(4.7), Inches(3.3)),
            "author":     (Inches(2), Inches(5.0), Inches(9.3), Inches(0.6)),
            "size_quote":  Pt(TYPE_SCALE["2xl"]),
            "size_author": Pt(TYPE_SCALE["md"]),
            "quote_italic": True,
        }

    def image_full(self) -> Dict:
        """Full-bleed image with overlay text at bottom."""
        return {
            "image":    (0, 0, Inches(self.SLIDE_W), Inches(self.SLIDE_H)),
            "overlay":  (0, Inches(4.8), Inches(self.SLIDE_W), Inches(2.7)),
            "heading":  (Inches(1.2), Inches(5.0), Inches(10.9), Inches(1.2)),
            "caption":  (Inches(1.2), Inches(6.3), Inches(10.9), Inches(0.6)),
            "size_heading": Pt(TYPE_SCALE["3xl"]),
            "size_caption": Pt(TYPE_SCALE["sm"]),
        }

    def chart_layout(self) -> Dict:
        """Chart / data visualization."""
        return {
            "heading":    (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "divider":    (Inches(0.8), Inches(1.05), Inches(2.2)),
            "chart":      (Inches(1.0), Inches(1.4), Inches(7.5), Inches(5.5)),
            "sidebar":    (Inches(9.0), Inches(1.4), Inches(3.5), Inches(5.5)),
            "size_heading": Pt(TYPE_SCALE["xl"]),
        }

    def metrics_layout(self, n_metrics: int = 3) -> Dict:
        """Big number KPI cards."""
        gap = 0.4
        col_w = (self.content_w - gap * (n_metrics - 1)) / n_metrics
        cards = []
        for i in range(n_metrics):
            left = self.margin + i * (col_w + gap)
            cards.append({
                "number":   (Inches(left), Inches(2.2), Inches(col_w), Inches(1.8)),
                "label":    (Inches(left), Inches(4.2), Inches(col_w), Inches(0.6)),
                "sublabel": (Inches(left), Inches(4.9), Inches(col_w), Inches(0.5)),
                "card_bg":  (Inches(left + 0.1), Inches(1.8), Inches(col_w - 0.2), Inches(4.0)),
            })
        return {
            "heading":    (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "divider":    (Inches(0.8), Inches(1.05), Inches(2.2)),
            "cards":      cards,
            "size_heading":   Pt(TYPE_SCALE["xl"]),
            "size_number":    Pt(TYPE_SCALE["5xl"]),
            "size_label":     Pt(TYPE_SCALE["base"]),
            "size_sublabel":  Pt(TYPE_SCALE["sm"]),
        }

    def agenda_layout(self) -> Dict:
        """Numbered agenda / outline."""
        return {
            "heading":     (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "divider":     (Inches(0.8), Inches(1.05), Inches(2.2)),
            "items_area":  (Inches(1.0), Inches(1.5), Inches(11), Inches(5.5)),
            "size_heading": Pt(TYPE_SCALE["2xl"]),
            "size_item":   Pt(TYPE_SCALE["md"]),
            "size_number": Pt(TYPE_SCALE["xl"]),
        }

    def timeline_layout(self) -> Dict:
        """Horizontal timeline."""
        return {
            "heading":  (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "divider":  (Inches(0.8), Inches(1.05), Inches(2.2)),
            "line":     (Inches(1.2), Inches(4.0), Inches(10.9), Inches(0.04)),
            "size_heading": Pt(TYPE_SCALE["xl"]),
            "size_label":   Pt(TYPE_SCALE["sm"]),
            "size_title":   Pt(TYPE_SCALE["base"]),
        }

    def data_table_layout(self) -> Dict:
        """Data table."""
        return {
            "heading":  (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "divider":  (Inches(0.8), Inches(1.05), Inches(2.2)),
            "table":    (Inches(0.6), Inches(1.4), Inches(12.1), Inches(5.2)),
            "size_heading": Pt(TYPE_SCALE["xl"]),
            "size_header":  Pt(TYPE_SCALE["sm"]),
            "size_cell":    Pt(TYPE_SCALE["sm"]),
        }

    def image_grid_layout(self, n_images: int = 6) -> Dict:
        """Gallery grid."""
        cols = min(3, n_images)
        rows = (n_images + cols - 1) // cols
        gap = 0.2
        cell_w = (self.content_w - gap * (cols - 1)) / cols
        cell_h = (self.content_h - 1.5 - gap * (rows - 1)) / rows

        cells = []
        for i in range(n_images):
            col_i = i % cols
            row_i = i // cols
            left = self.margin + col_i * (cell_w + gap)
            top = Inches(1.4) + row_i * (cell_h + gap)
            cells.append((Inches(left), top, Inches(cell_w), Inches(cell_h)))

        return {
            "heading":  (Inches(0.8), Inches(0.3), Inches(11), Inches(0.7)),
            "divider":  (Inches(0.8), Inches(1.05), Inches(2.2)),
            "cells":    cells,
            "size_heading": Pt(TYPE_SCALE["xl"]),
        }

    def end_centered(self) -> Dict:
        """Centered closing slide."""
        return {
            "title":    (Inches(1.2), Inches(2.2), Inches(10.9), Inches(1.8)),
            "divider":  (Inches(5.2), Inches(4.1), Inches(2.9)),
            "subtitle": (Inches(2.5), Inches(4.4), Inches(8.3), Inches(1.2)),
            "size_title":    Pt(TYPE_SCALE["4xl"]),
            "size_subtitle": Pt(TYPE_SCALE["lg"]),
        }

    # ─── Backward Compatibility Aliases ───────────────────────────────

    def title_position(self, vertical_center: bool = True) -> Dict:
        return self.title_centered() if vertical_center else self.title_left_aligned()

    def content_position(self, has_image: bool = False, image_side: str = "right") -> Dict:
        if has_image:
            return self.content_with_image(image_side)
        return self.content_standard()

    def two_column_position(self) -> Dict:
        return self.two_column()

    def three_column_position(self) -> Dict:
        return self.three_column()

    def quote_position(self) -> Dict:
        return self.quote_layout()

    def image_full_position(self) -> Dict:
        return self.image_full()

    def chart_position(self, chart_type: str = "bar") -> Dict:
        return self.chart_layout()

    def end_position(self) -> Dict:
        return self.end_centered()

    def agenda_position(self) -> Dict:
        return self.agenda_layout()

    def timeline_position(self) -> Dict:
        return self.timeline_layout()

    def data_table_position(self) -> Dict:
        return self.data_table_layout()

    def metrics_position(self) -> Dict:
        return self.metrics_layout()
