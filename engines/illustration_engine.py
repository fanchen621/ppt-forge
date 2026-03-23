"""
PPT Forge v2 — Intelligent Illustration Engine
Generates contextually relevant SVG illustrations based on slide content.
Parses keywords from headings/body text and produces matching art.
Zero paid APIs — pure SVG + Pillow rendering.
"""

import re
import math
import random
import subprocess
import tempfile
import os
from typing import List, Tuple, Dict, Optional
from .art_engine import PALETTES, get_palette, palette_color, adjust_brightness, with_alpha, lerp_color


def _rgb(color: Tuple) -> str:
    return f"rgb({color[0]},{color[1]},{color[2]})"

def _rgba(color: Tuple, a: float) -> str:
    return f"rgba({color[0]},{color[1]},{color[2]},{a})"

def _svg_wrap(width, height, body):
    return f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">{body}</svg>'


# ─── Keyword Detection ─────────────────────────────────────────────────────

CATEGORY_KEYWORDS = {
    "plant": ["花", "草", "树", "植物", "种子", "叶", "根", "茎", "果", "苗", "芽",
              "flower", "plant", "seed", "leaf", "tree", "root", "stem", "grow",
              "凤仙花", "光合", "生长", "开花", "发芽", "结果", "标本", "种植", "浇水"],
    "science": ["科学", "实验", "观察", "测量", "数据", "分析", "研究", "探究", "对照",
                "science", "experiment", "observe", "measure", "data", "research",
                "记录", "统计", "图表", "折线", "曲线", "对比"],
    "education": ["学习", "教学", "学生", "课堂", "课程", "教育", "知识", "能力",
                  "learn", "teach", "student", "class", "education", "school",
                  "素养", "目标", "评价", "汇报", "展示", "项目"],
    "time": ["时间", "周", "阶段", "计划", "安排", "日程", "路线", "步骤",
             "time", "week", "phase", "plan", "schedule", "timeline",
             "第1", "第2", "第3", "第4", "第5", "第6", "第7", "第8"],
    "data": ["数据", "图表", "统计", "分析", "指标", "对比", "百分", "增长",
             "data", "chart", "graph", "metric", "analysis", "growth"],
    "team": ["合作", "小组", "团队", "交流", "讨论", "分享",
             "team", "group", "collaborate", "share"],
    "safety": ["安全", "提示", "注意", "防护", "洗手", "保护",
               "safety", "protect", "warning"],
    "resource": ["资源", "材料", "工具", "物资", "清单",
                 "resource", "material", "tool", "supply"],
    "award": ["成果", "展示", "评价", "优秀", "荣誉", "汇报",
              "result", "show", "evaluate", "award", "excellent"],
    "tech": ["技术", "数字", "APP", "视频", "智能", "编程",
             "tech", "digital", "app", "video", "smart"],
    "nature": ["自然", "生态", "环境", "天气", "季节", "阳光", "水", "土壤",
               "nature", "ecology", "environment", "weather", "sun", "water"],
}

def detect_categories(text: str) -> List[str]:
    """Detect content categories from text."""
    text_lower = text.lower()
    scores = {}
    for cat, keywords in CATEGORY_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw.lower() in text_lower)
        if score > 0:
            scores[cat] = score
    # Return top categories sorted by relevance
    sorted_cats = sorted(scores.keys(), key=lambda k: scores[k], reverse=True)
    return sorted_cats[:3] if sorted_cats else ["education"]


# ─── SVG Illustration Generators ───────────────────────────────────────────

def _gen_plant_illustration(w, h, palette, seed=None):
    """Botanical illustration: stem, leaves, flower, roots."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    # Background circle/oval
    cx, cy = w // 2, h // 2
    parts.append(f'<circle cx="{cx}" cy="{cy}" r="{min(w,h)*0.42}" fill="{_rgba(c[0], 0.08)}"/>')

    # Soil line
    soil_y = int(h * 0.72)
    parts.append(f'<ellipse cx="{cx}" cy="{soil_y+10}" rx="{w*0.35}" ry="{h*0.08}" fill="{_rgba(c[0], 0.15)}"/>')

    # Roots
    root_color = _rgb(adjust_brightness(c[0], 0.6))
    for i in range(5):
        rx = cx + random.randint(-40, 40)
        ry = soil_y + random.randint(5, 30)
        rx2 = rx + random.randint(-30, 30)
        ry2 = ry + random.randint(20, 50)
        parts.append(f'<path d="M{cx} {soil_y} Q{rx} {ry} {rx2} {ry2}" fill="none" stroke="{root_color}" stroke-width="2" opacity="0.5"/>')

    # Stem
    stem_color = _rgb(adjust_brightness(c[1], 0.8))
    stem_top = int(h * 0.18)
    parts.append(f'<path d="M{cx} {soil_y} C{cx-5} {int(soil_y*0.7)} {cx+8} {int(soil_y*0.4)} {cx} {stem_top}" fill="none" stroke="{stem_color}" stroke-width="4"/>')

    # Leaves along stem
    leaf_color1 = _rgb(c[1])
    leaf_color2 = _rgb(adjust_brightness(c[1], 1.2))
    for i in range(4):
        ly = int(soil_y - (soil_y - stem_top) * (i + 1) / 5)
        side = 1 if i % 2 == 0 else -1
        lx = cx + side * 5
        # Leaf shape (bezier curve)
        tip_x = lx + side * random.randint(30, 55)
        tip_y = ly + random.randint(-10, 10)
        ctrl1_x = lx + side * random.randint(15, 30)
        ctrl1_y = ly - random.randint(10, 25)
        ctrl2_x = tip_x - side * random.randint(5, 15)
        ctrl2_y = tip_y + random.randint(5, 15)
        parts.append(f'<path d="M{lx} {ly} C{ctrl1_x} {ctrl1_y} {ctrl2_x} {ctrl2_y} {tip_x} {tip_y} Q{(lx+tip_x)//2} {ly+10} {lx} {ly}" fill="{leaf_color1}" opacity="0.8"/>')
        # Leaf vein
        parts.append(f'<line x1="{lx}" y1="{ly}" x2="{tip_x}" y2="{tip_y}" stroke="{leaf_color2}" stroke-width="1" opacity="0.4"/>')

    # Flower at top
    petal_count = random.randint(5, 7)
    flower_cx, flower_cy = cx, stem_top - 5
    petal_colors = [c[2], c[3], c[4] if len(c) > 4 else c[3]]
    for i in range(petal_count):
        angle = 2 * math.pi * i / petal_count
        px = flower_cx + 18 * math.cos(angle)
        py = flower_cy + 18 * math.sin(angle)
        pc = petal_colors[i % len(petal_colors)]
        parts.append(f'<ellipse cx="{px}" cy="{py}" rx="12" ry="18" fill="{_rgb(pc)}" opacity="0.85" transform="rotate({math.degrees(angle)} {px} {py})"/>')

    # Flower center
    parts.append(f'<circle cx="{flower_cx}" cy="{flower_cy}" r="8" fill="{_rgb(petal_colors[0])}"/>')
    parts.append(f'<circle cx="{flower_cx}" cy="{flower_cy}" r="5" fill="{_rgb(adjust_brightness(petal_colors[0], 1.3))}"/>')

    # Seed pods (small circles along stem)
    for i in range(3):
        sy = int(stem_top + (soil_y - stem_top) * random.uniform(0.3, 0.8))
        sx = cx + random.randint(-15, 15)
        parts.append(f'<circle cx="{sx}" cy="{sy}" r="4" fill="{_rgb(c[3])}" opacity="0.6"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_science_illustration(w, h, palette, seed=None):
    """Science illustration: magnifier, flask, microscope, data."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Magnifying glass
    mx, my = int(w * 0.35), int(h * 0.4)
    mr = 35
    parts.append(f'<circle cx="{mx}" cy="{my}" r="{mr}" fill="{_rgba(c[0], 0.1)}" stroke="{_rgb(c[2])}" stroke-width="3"/>')
    parts.append(f'<line x1="{mx+mr*0.7}" y1="{my+mr*0.7}" x2="{mx+mr+20}" y2="{my+mr+20}" stroke="{_rgb(c[2])}" stroke-width="5" stroke-linecap="round"/>')
    # Lens reflection
    parts.append(f'<circle cx="{mx-8}" cy="{my-8}" r="8" fill="{_rgba(c[4], 0.2)}"/>')
    # Cells inside magnifier
    for i in range(5):
        cell_x = mx + random.randint(-15, 15)
        cell_y = my + random.randint(-15, 15)
        cell_r = random.randint(4, 9)
        parts.append(f'<circle cx="{cell_x}" cy="{cell_y}" r="{cell_r}" fill="none" stroke="{_rgb(c[1])}" stroke-width="1.5" opacity="0.6"/>')
        parts.append(f'<circle cx="{cell_x}" cy="{cell_y}" r="2" fill="{_rgb(c[3])}" opacity="0.8"/>')

    # Test tube
    tx, ty = int(w * 0.65), int(h * 0.2)
    tube_w, tube_h = 22, 80
    parts.append(f'<rect x="{tx}" y="{ty}" width="{tube_w}" height="{tube_h}" rx="11" fill="{_rgba(c[0], 0.08)}" stroke="{_rgb(c[2])}" stroke-width="2"/>')
    # Liquid level
    liquid_h = random.randint(30, 60)
    parts.append(f'<rect x="{tx+2}" y="{ty+tube_h-liquid_h}" width="{tube_w-4}" height="{liquid_h}" rx="9" fill="{_rgba(c[3], 0.5)}"/>')
    # Bubbles
    for _ in range(4):
        bx = tx + random.randint(5, tube_w - 5)
        by = ty + random.randint(tube_h - liquid_h + 5, tube_h - 10)
        br = random.randint(2, 4)
        parts.append(f'<circle cx="{bx}" cy="{by}" r="{br}" fill="{_rgba(c[4], 0.3)}"/>')

    # Data chart (mini bar chart)
    chart_x, chart_y = int(w * 0.2), int(h * 0.72)
    bar_w = 12
    for i in range(5):
        bh = random.randint(15, 50)
        bx = chart_x + i * (bar_w + 5)
        parts.append(f'<rect x="{bx}" y="{chart_y - bh}" width="{bar_w}" height="{bh}" rx="2" fill="{_rgb(c[i % len(c)])}" opacity="0.7"/>')
    # Axis
    parts.append(f'<line x1="{chart_x-3}" y1="{chart_y}" x2="{chart_x + 5*(bar_w+5)}" y2="{chart_y}" stroke="{_rgb(c[2])}" stroke-width="1.5"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_education_illustration(w, h, palette, seed=None):
    """Education illustration: book, lightbulb, pencil, graduation cap."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Open book
    bx, by = int(w * 0.3), int(h * 0.55)
    book_w, book_h = 100, 70
    # Left page
    parts.append(f'<path d="M{bx+book_w//2} {by} Q{bx+book_w//4} {by-8} {bx} {by} L{bx} {by+book_h} Q{bx+book_w//4} {by+book_h-8} {bx+book_w//2} {by+book_h} Z" fill="{_rgba(c[0], 0.1)}" stroke="{_rgb(c[2])}" stroke-width="1.5"/>')
    # Right page
    parts.append(f'<path d="M{bx+book_w//2} {by} Q{bx+3*book_w//4} {by-8} {bx+book_w} {by} L{bx+book_w} {by+book_h} Q{bx+3*book_w//4} {by+book_h-8} {bx+book_w//2} {by+book_h} Z" fill="{_rgba(c[0], 0.1)}" stroke="{_rgb(c[2])}" stroke-width="1.5"/>')
    # Text lines on pages
    for i in range(4):
        ly = by + 15 + i * 14
        lw = random.randint(20, 40)
        parts.append(f'<line x1="{bx+10}" y1="{ly}" x2="{bx+10+lw}" y2="{ly}" stroke="{_rgb(c[3])}" stroke-width="1" opacity="0.4"/>')
        parts.append(f'<line x1="{bx+book_w//2+10}" y1="{ly}" x2="{bx+book_w//2+10+lw}" y2="{ly}" stroke="{_rgb(c[3])}" stroke-width="1" opacity="0.4"/>')

    # Lightbulb
    lbx, lby = int(w * 0.7), int(h * 0.3)
    # Glow
    parts.append(f'<circle cx="{lbx}" cy="{lby}" r="35" fill="{_rgba(c[4], 0.1)}"/>')
    parts.append(f'<circle cx="{lbx}" cy="{lby}" r="25" fill="{_rgba(c[4], 0.15)}"/>')
    # Bulb
    parts.append(f'<path d="M{lbx-15} {lby+5} A18 18 0 1 1 {lbx+15} {lby+5} L{lbx+10} {lby+20} L{lbx-10} {lby+20} Z" fill="{_rgba(c[4], 0.3)}" stroke="{_rgb(c[3])}" stroke-width="2"/>')
    # Base
    parts.append(f'<rect x="{lbx-8}" y="{lby+20}" width="16" height="8" rx="2" fill="{_rgb(c[2])}"/>')
    # Rays
    for i in range(8):
        angle = 2 * math.pi * i / 8
        r1, r2 = 30, 42
        parts.append(f'<line x1="{lbx+r1*math.cos(angle)}" y1="{lby+r1*math.sin(angle)}" x2="{lbx+r2*math.cos(angle)}" y2="{lby+r2*math.sin(angle)}" stroke="{_rgb(c[4])}" stroke-width="2" opacity="0.5"/>')

    # Pencil
    px, py = int(w * 0.75), int(h * 0.7)
    parts.append(f'<rect x="{px}" y="{py}" width="8" height="50" rx="1" fill="{_rgb(c[1])}" transform="rotate(-20 {px+4} {py+25})"/>')
    parts.append(f'<polygon points="{px} {py} {px+8} {py} {px+4} {py-12}" fill="{_rgb(c[3])}" transform="rotate(-20 {px+4} {py})"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_timeline_illustration(w, h, palette, seed=None):
    """Timeline/path illustration with milestones."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    # Winding path
    points = []
    y = int(h * 0.5)
    for x in range(40, w - 40, 8):
        y_offset = math.sin(x * 0.02) * (h * 0.15)
        points.append((x, int(y + y_offset)))

    path_d = f"M{points[0][0]} {points[0][1]}"
    for i in range(1, len(points)):
        path_d += f" L{points[i][0]} {points[i][1]}"
    parts.append(f'<path d="{path_d}" fill="none" stroke="{_rgba(c[2], 0.4)}" stroke-width="4" stroke-linecap="round"/>')

    # Milestone dots
    n_milestones = 6
    for i in range(n_milestones):
        idx = int(len(points) * (i + 0.5) / n_milestones)
        mx, my = points[idx]
        mc = c[(i + 1) % len(c)]
        parts.append(f'<circle cx="{mx}" cy="{my}" r="12" fill="{_rgb(mc)}" opacity="0.8"/>')
        parts.append(f'<circle cx="{mx}" cy="{my}" r="6" fill="white" opacity="0.9"/>')
        # Flag/stem above
        parts.append(f'<line x1="{mx}" y1="{my-12}" x2="{mx}" y2="{my-35}" stroke="{_rgb(mc)}" stroke-width="2"/>')
        parts.append(f'<rect x="{mx}" y="{my-35}" width="20" height="12" rx="2" fill="{_rgba(mc, 0.6)}"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_data_illustration(w, h, palette, seed=None):
    """Data visualization illustration: charts, pie, lines."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Pie chart
    pie_cx, pie_cy = int(w * 0.3), int(h * 0.4)
    pie_r = 40
    start_angle = 0
    slices = [0.3, 0.25, 0.2, 0.15, 0.1]
    for i, size in enumerate(slices):
        end_angle = start_angle + size * 2 * math.pi
        x1 = pie_cx + pie_r * math.cos(start_angle)
        y1 = pie_cy + pie_r * math.sin(start_angle)
        x2 = pie_cx + pie_r * math.cos(end_angle)
        y2 = pie_cy + pie_r * math.sin(end_angle)
        large = 1 if size > 0.5 else 0
        parts.append(f'<path d="M{pie_cx} {pie_cy} L{x1} {y1} A{pie_r} {pie_r} 0 {large} 1 {x2} {y2} Z" fill="{_rgb(c[i % len(c)])}" opacity="0.7"/>')
        start_angle = end_angle

    # Line chart
    line_x, line_y = int(w * 0.55), int(h * 0.25)
    line_w, line_h = int(w * 0.35), int(h * 0.4)
    # Axes
    parts.append(f'<line x1="{line_x}" y1="{line_y+line_h}" x2="{line_x+line_w}" y2="{line_y+line_h}" stroke="{_rgb(c[2])}" stroke-width="1.5"/>')
    parts.append(f'<line x1="{line_x}" y1="{line_y}" x2="{line_x}" y2="{line_y+line_h}" stroke="{_rgb(c[2])}" stroke-width="1.5"/>')
    # Data line
    pts = []
    for i in range(6):
        px = line_x + int(line_w * i / 5)
        py = line_y + line_h - int(line_h * random.uniform(0.2, 0.9))
        pts.append(f"{px},{py}")
    parts.append(f'<polyline points="{" ".join(pts)}" fill="none" stroke="{_rgb(c[3])}" stroke-width="2.5"/>')
    # Dots
    for pt in pts:
        x, y = pt.split(",")
        parts.append(f'<circle cx="{x}" cy="{y}" r="4" fill="{_rgb(c[3])}"/>')

    # Bar chart
    bar_x, bar_y = int(w * 0.15), int(h * 0.78)
    for i in range(4):
        bh = random.randint(15, 45)
        bx = bar_x + i * 22
        parts.append(f'<rect x="{bx}" y="{bar_y - bh}" width="16" height="{bh}" rx="3" fill="{_rgb(c[i % len(c)])}" opacity="0.6"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_team_illustration(w, h, palette, seed=None):
    """Team/collaboration illustration: people icons, connections."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Connection lines between people
    n_people = 5
    positions = []
    for i in range(n_people):
        angle = 2 * math.pi * i / n_people - math.pi / 2
        r = min(w, h) * 0.28
        px = cx + r * math.cos(angle)
        py = cy + r * math.sin(angle)
        positions.append((int(px), int(py)))

    # Lines between all
    for i in range(len(positions)):
        for j in range(i + 1, len(positions)):
            parts.append(f'<line x1="{positions[i][0]}" y1="{positions[i][1]}" x2="{positions[j][0]}" y2="{positions[j][1]}" stroke="{_rgba(c[2], 0.2)}" stroke-width="1.5"/>')

    # People (head + body)
    for i, (px, py) in enumerate(positions):
        pc = c[(i + 1) % len(c)]
        # Head
        parts.append(f'<circle cx="{px}" cy="{py-8}" r="10" fill="{_rgb(pc)}" opacity="0.8"/>')
        # Body
        parts.append(f'<path d="M{px-12} {py+15} A12 12 0 0 1 {px+12} {py+15}" fill="{_rgb(pc)}" opacity="0.6"/>')
        # Center connection dot
        parts.append(f'<circle cx="{px}" cy="{py}" r="3" fill="{_rgb(c[0])}"/>')

    # Center symbol
    parts.append(f'<circle cx="{cx}" cy="{cy}" r="15" fill="{_rgba(c[3], 0.3)}" stroke="{_rgb(c[3])}" stroke-width="2"/>')
    parts.append(f'<text x="{cx}" y="{cy+5}" text-anchor="middle" font-size="14" fill="{_rgb(c[3])}" font-weight="bold">★</text>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_safety_illustration(w, h, palette, seed=None):
    """Safety/health illustration: shield, hands, heart."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Shield
    sx, sy = cx, int(h * 0.35)
    shield_w, shield_h = 60, 75
    parts.append(f'<path d="M{sx} {sy-shield_h//2} L{sx+shield_w//2} {sy-shield_h//4} L{sx+shield_w//2} {sy+shield_h//6} Q{sx+shield_w//4} {sy+shield_h//2} {sx} {sy+shield_h//2+10} Q{sx-shield_w//4} {sy+shield_h//2} {sx-shield_w//2} {sy+shield_h//6} L{sx-shield_w//2} {sy-shield_h//4} Z" fill="{_rgba(c[1], 0.2)}" stroke="{_rgb(c[1])}" stroke-width="2"/>')
    # Checkmark inside shield
    parts.append(f'<path d="M{sx-12} {sy+5} L{sx-3} {sy+15} L{sx+15} {sy-10}" fill="none" stroke="{_rgb(c[1])}" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>')

    # Hands washing
    hx, hy = int(w * 0.65), int(h * 0.5)
    # Left hand
    parts.append(f'<ellipse cx="{hx-15}" cy="{hy}" rx="18" ry="12" fill="{_rgba(c[2], 0.4)}" stroke="{_rgb(c[2])}" stroke-width="1.5"/>')
    # Right hand
    parts.append(f'<ellipse cx="{hx+15}" cy="{hy}" rx="18" ry="12" fill="{_rgba(c[2], 0.4)}" stroke="{_rgb(c[2])}" stroke-width="1.5"/>')
    # Water drops
    for i in range(5):
        dx = hx + random.randint(-20, 20)
        dy = hy + random.randint(15, 40)
        parts.append(f'<path d="M{dx} {dy-5} Q{dx-3} {dy} {dx} {dy+4} Q{dx+3} {dy} {dx} {dy-5}" fill="{_rgba(c[3], 0.5)}"/>')

    # Heart
    hcx, hcy = int(w * 0.3), int(h * 0.72)
    parts.append(f'<path d="M{hcx} {hcy+12} C{hcx-20} {hcy-5} {hcx-20} {hcy-20} {hcx} {hcy-10} C{hcx+20} {hcy-20} {hcx+20} {hcy-5} {hcx} {hcy+12}" fill="{_rgba(c[4], 0.4)}" stroke="{_rgb(c[4])}" stroke-width="1.5"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_resource_illustration(w, h, palette, seed=None):
    """Resource/tool illustration: toolbox, materials."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Toolbox
    tx, ty = int(w * 0.25), int(h * 0.35)
    tw, th = 80, 50
    parts.append(f'<rect x="{tx}" y="{ty}" width="{tw}" height="{th}" rx="5" fill="{_rgba(c[1], 0.2)}" stroke="{_rgb(c[1])}" stroke-width="2"/>')
    parts.append(f'<rect x="{tx+tw//2-15}" y="{ty-5}" width="30" height="10" rx="3" fill="{_rgba(c[1], 0.3)}" stroke="{_rgb(c[1])}" stroke-width="1.5"/>')
    # Handle
    parts.append(f'<path d="M{tx+tw//2-20} {ty-5} Q{tx+tw//2} {ty-20} {tx+tw//2+20} {ty-5}" fill="none" stroke="{_rgb(c[2])}" stroke-width="3"/>')

    # Scattered tools
    # Wrench
    wx, wy = int(w * 0.7), int(h * 0.3)
    parts.append(f'<rect x="{wx}" y="{wy}" width="6" height="40" rx="2" fill="{_rgb(c[2])}" transform="rotate(30 {wx+3} {wy+20})"/>')
    parts.append(f'<circle cx="{wx}" cy="{wy}" r="8" fill="none" stroke="{_rgb(c[2])}" stroke-width="3" transform="rotate(30 {wx} {wy})"/>')

    # Seeds scattered
    for i in range(8):
        sx = random.randint(int(w*0.3), int(w*0.8))
        sy = random.randint(int(h*0.6), int(h*0.85))
        sr = random.randint(3, 7)
        parts.append(f'<ellipse cx="{sx}" cy="{sy}" rx="{sr}" ry="{sr*0.6}" fill="{_rgb(c[(i+2)%len(c)])}" opacity="0.6"/>')

    # Ruler
    rx, ry = int(w * 0.6), int(h * 0.6)
    parts.append(f'<rect x="{rx}" y="{ry}" width="70" height="10" rx="1" fill="{_rgba(c[3], 0.4)}" stroke="{_rgb(c[3])}" stroke-width="1"/>')
    for i in range(8):
        mx = rx + i * 10
        mh = 4 if i % 2 == 0 else 2
        parts.append(f'<line x1="{mx}" y1="{ry}" x2="{mx}" y2="{ry+mh}" stroke="{_rgb(c[3])}" stroke-width="0.5"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_award_illustration(w, h, palette, seed=None):
    """Award/achievement illustration: trophy, star, ribbon."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Star
    sx, sy = int(w * 0.35), int(h * 0.35)
    star_r = 35
    star_pts = []
    for i in range(10):
        angle = math.pi * i / 5 - math.pi / 2
        r = star_r if i % 2 == 0 else star_r * 0.45
        star_pts.append(f"{sx + r * math.cos(angle)},{sy + r * math.sin(angle)}")
    parts.append(f'<polygon points="{" ".join(star_pts)}" fill="{_rgba(c[4], 0.5)}" stroke="{_rgb(c[3])}" stroke-width="2"/>')

    # Trophy
    tx, ty = int(w * 0.6), int(h * 0.3)
    parts.append(f'<path d="M{tx-20} {ty} L{tx-15} {ty+35} L{tx-8} {ty+35} L{tx-5} {ty+45} L{tx+5} {ty+45} L{tx+8} {ty+35} L{tx+15} {ty+35} L{tx+20} {ty} Z" fill="{_rgba(c[4], 0.4)}" stroke="{_rgb(c[3])}" stroke-width="2"/>')
    # Handles
    parts.append(f'<path d="M{tx-20} {ty+5} Q{tx-30} {ty+5} {tx-30} {ty+15} Q{tx-30} {ty+25} {tx-20} {ty+20}" fill="none" stroke="{_rgb(c[3])}" stroke-width="2"/>')
    parts.append(f'<path d="M{tx+20} {ty+5} Q{tx+30} {ty+5} {tx+30} {ty+15} Q{tx+30} {ty+25} {tx+20} {ty+20}" fill="none" stroke="{_rgb(c[3])}" stroke-width="2"/>')

    # Ribbon
    rx, ry = int(w * 0.35), int(h * 0.7)
    parts.append(f'<path d="M{rx-20} {ry} L{rx} {ry-25} L{rx+20} {ry} L{rx+15} {ry+30} L{rx} {ry+20} L{rx-15} {ry+30} Z" fill="{_rgba(c[1], 0.4)}" stroke="{_rgb(c[1])}" stroke-width="1.5"/>')

    # Confetti
    for i in range(15):
        fx = random.randint(0, w)
        fy = random.randint(0, h)
        fw = random.randint(4, 10)
        fh = random.randint(2, 5)
        angle = random.randint(0, 180)
        fc = random.choice(c)
        parts.append(f'<rect x="{fx}" y="{fy}" width="{fw}" height="{fh}" rx="1" fill="{_rgb(fc)}" opacity="0.4" transform="rotate({angle} {fx+fw//2} {fy+fh//2})"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_nature_illustration(w, h, palette, seed=None):
    """Nature illustration: sun, cloud, rain, mountains."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Sun
    sun_x, sun_y = int(w * 0.75), int(h * 0.2)
    sun_r = 25
    parts.append(f'<circle cx="{sun_x}" cy="{sun_y}" r="{sun_r+15}" fill="{_rgba(c[4], 0.1)}"/>')
    parts.append(f'<circle cx="{sun_x}" cy="{sun_y}" r="{sun_r}" fill="{_rgba(c[4], 0.4)}" stroke="{_rgb(c[3])}" stroke-width="2"/>')
    for i in range(12):
        angle = 2 * math.pi * i / 12
        r1 = sun_r + 5
        r2 = sun_r + 15
        parts.append(f'<line x1="{sun_x+r1*math.cos(angle)}" y1="{sun_y+r1*math.sin(angle)}" x2="{sun_x+r2*math.cos(angle)}" y2="{sun_y+r2*math.sin(angle)}" stroke="{_rgb(c[3])}" stroke-width="2"/>')

    # Cloud
    cloud_x, cloud_y = int(w * 0.3), int(h * 0.2)
    parts.append(f'<ellipse cx="{cloud_x}" cy="{cloud_y}" rx="30" ry="15" fill="{_rgba(c[0], 0.15)}" stroke="{_rgba(c[2], 0.3)}" stroke-width="1"/>')
    parts.append(f'<ellipse cx="{cloud_x-15}" cy="{cloud_y+5}" rx="20" ry="12" fill="{_rgba(c[0], 0.12)}"/>')
    parts.append(f'<ellipse cx="{cloud_x+15}" cy="{cloud_y+5}" rx="22" ry="13" fill="{_rgba(c[0], 0.12)}"/>')

    # Mountains
    mt_y = int(h * 0.65)
    parts.append(f'<polygon points="0,{mt_y+40} {int(w*0.2)},{mt_y-30} {int(w*0.4)},{mt_y+40}" fill="{_rgba(c[1], 0.3)}"/>')
    parts.append(f'<polygon points="{int(w*0.15)},{mt_y+40} {int(w*0.45)},{mt_y-15} {int(w*0.7)},{mt_y+40}" fill="{_rgba(c[0], 0.2)}"/>')
    parts.append(f'<polygon points="{int(w*0.35)},{mt_y+40} {int(w*0.6)},{mt_y-40} {int(w*0.85)},{mt_y+40}" fill="{_rgba(c[1], 0.25)}"/>')

    # Trees
    for i in range(5):
        tx = int(w * (0.1 + i * 0.2))
        ty = mt_y + random.randint(-10, 10)
        tree_h = random.randint(20, 40)
        parts.append(f'<rect x="{tx-2}" y="{ty}" width="4" height="{tree_h//3}" fill="{_rgb(c[0])}"/>')
        parts.append(f'<polygon points="{tx} {ty-tree_h//2} {tx-tree_h//4} {ty} {tx+tree_h//4} {ty}" fill="{_rgb(c[1])}" opacity="0.7"/>')

    # Ground
    parts.append(f'<rect x="0" y="{mt_y+40}" width="{w}" height="{h-mt_y-40}" fill="{_rgba(c[0], 0.1)}"/>')

    # Water drops (rain)
    for i in range(8):
        dx = random.randint(0, w)
        dy = random.randint(int(h*0.3), int(h*0.55))
        parts.append(f'<path d="M{dx} {dy} Q{dx-2} {dy+5} {dx} {dy+8} Q{dx+2} {dy+5} {dx} {dy}" fill="{_rgba(c[3], 0.3)}"/>')

    return _svg_wrap(w, h, "\n".join(parts))


def _gen_abstract_illustration(w, h, palette, seed=None):
    """Abstract decorative illustration as fallback."""
    if seed: random.seed(seed)
    c = palette
    parts = []

    cx, cy = w // 2, h // 2

    # Abstract circles
    for i in range(8):
        ax = random.randint(0, w)
        ay = random.randint(0, h)
        ar = random.randint(15, 60)
        ac = random.choice(c)
        parts.append(f'<circle cx="{ax}" cy="{ay}" r="{ar}" fill="{_rgba(ac, 0.15)}"/>')

    # Connecting curves
    for i in range(5):
        x1, y1 = random.randint(0, w), random.randint(0, h)
        x2, y2 = random.randint(0, w), random.randint(0, h)
        ctrl_x = (x1 + x2) // 2 + random.randint(-50, 50)
        ctrl_y = (y1 + y2) // 2 + random.randint(-50, 50)
        ac = random.choice(c)
        parts.append(f'<path d="M{x1} {y1} Q{ctrl_x} {ctrl_y} {x2} {y2}" fill="none" stroke="{_rgb(ac)}" stroke-width="1.5" opacity="0.4"/>')

    # Small dots
    for i in range(20):
        dx = random.randint(0, w)
        dy = random.randint(0, h)
        dr = random.randint(2, 6)
        dc = random.choice(c)
        parts.append(f'<circle cx="{dx}" cy="{dy}" r="{dr}" fill="{_rgb(dc)}" opacity="0.5"/>')

    return _svg_wrap(w, h, "\n".join(parts))


# ─── Registry ──────────────────────────────────────────────────────────────

ILLUSTRATION_GENERATORS = {
    "plant": _gen_plant_illustration,
    "science": _gen_science_illustration,
    "education": _gen_education_illustration,
    "time": _gen_timeline_illustration,
    "data": _gen_data_illustration,
    "team": _gen_team_illustration,
    "safety": _gen_safety_illustration,
    "resource": _gen_resource_illustration,
    "award": _gen_award_illustration,
    "nature": _gen_nature_illustration,
}


# ─── Public API ────────────────────────────────────────────────────────────

def generate_illustration(category: str, width: int = 300, height: int = 300,
                           palette_name: str = "forest", seed: int = None) -> str:
    """Generate SVG illustration for a given category."""
    gen = ILLUSTRATION_GENERATORS.get(category, _gen_abstract_illustration)
    palette = get_palette(palette_name)
    return gen(width, height, palette, seed=seed)


def generate_contextual_illustration(text: str, width: int = 300, height: int = 300,
                                      palette_name: str = "forest", seed: int = None) -> str:
    """Auto-detect content category and generate matching illustration."""
    cats = detect_categories(text)
    cat = cats[0] if cats else "education"
    return generate_illustration(cat, width, height, palette_name, seed)


def auto_illustrate_slide(slide_config: dict, palette_name: str = "forest",
                           seed: int = None) -> str:
    """Generate the most relevant illustration for a slide based on its content."""
    # Combine all text fields
    text_parts = []
    text_parts.append(slide_config.get("heading", ""))
    text_parts.append(slide_config.get("subheading", ""))
    text_parts.append(slide_config.get("text", ""))
    text_parts.append(slide_config.get("caption", ""))
    body = slide_config.get("body", [])
    if isinstance(body, list):
        text_parts.extend(body)
    elif isinstance(body, str):
        text_parts.append(body)

    # Also check columns and other nested content
    for key in ("left", "right"):
        col = slide_config.get(key, {})
        if isinstance(col, dict):
            text_parts.append(col.get("title", ""))
            items = col.get("items", [])
            if isinstance(items, list):
                text_parts.extend(items)

    for col in slide_config.get("columns", []):
        if isinstance(col, dict):
            text_parts.append(col.get("title", ""))
            items = col.get("items", [])
            if isinstance(items, list):
                text_parts.extend(items)

    combined_text = " ".join(str(t) for t in text_parts)
    return generate_contextual_illustration(combined_text, width=300, height=300,
                                             palette_name=palette_name, seed=seed)
