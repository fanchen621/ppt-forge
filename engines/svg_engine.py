"""
PPT Forge v2 — SVG Art Engine
Generates vector-quality SVG art, converts to PNG via rsvg-convert.
Produces crisp geometric patterns, icons, and decorative elements.
"""

import os
import math
import random
import subprocess
import tempfile
from typing import List, Tuple, Optional


# ─── SVG Primitives ────────────────────────────────────────────────────────

def svg_header(width: int, height: int) -> str:
    return f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">'


SVG_FOOTER = '</svg>'


def svg_defs_gradient_linear(id_name: str, colors: List[Tuple], x1="0%", y1="0%", x2="100%", y2="100%") -> str:
    """Create a linear gradient SVG def."""
    stops = ""
    for i, (r, g, b) in enumerate(colors):
        offset = i / max(1, len(colors) - 1) * 100
        stops += f'<stop offset="{offset:.0f}%" stop-color="rgb({r},{g},{b})"/>'
    return f'''<defs><linearGradient id="{id_name}" x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}">{stops}</linearGradient></defs>'''


def svg_defs_radial_gradient(id_name: str, colors: List[Tuple]) -> str:
    stops = ""
    for i, (r, g, b) in enumerate(colors):
        offset = i / max(1, len(colors) - 1) * 100
        stops += f'<stop offset="{offset:.0f}%" stop-color="rgb({r},{g},{b})"/>'
    return f'''<defs><radialGradient id="{id_name}">{stops}</radialGradient></defs>'''


def svg_circle(cx, cy, r, fill=None, stroke=None, stroke_width=1, opacity=1.0) -> str:
    s = f'<circle cx="{cx}" cy="{cy}" r="{r}"'
    if fill: s += f' fill="{fill}"'
    if stroke: s += f' stroke="{stroke}" stroke-width="{stroke_width}"'
    s += f' opacity="{opacity}"'
    return s + '/>'


def svg_rect(x, y, w, h, fill=None, rx=0, opacity=1.0) -> str:
    s = f'<rect x="{x}" y="{y}" width="{w}" height="{h}"'
    if fill: s += f' fill="{fill}"'
    if rx > 0: s += f' rx="{rx}"'
    s += f' opacity="{opacity}"'
    return s + '/>'


def svg_polygon(points: List[Tuple], fill=None, opacity=1.0) -> str:
    pts = " ".join(f"{x},{y}" for x, y in points)
    return f'<polygon points="{pts}" fill="{fill}" opacity="{opacity}"/>'


def svg_path(d: str, fill=None, stroke=None, stroke_width=1, opacity=1.0) -> str:
    s = f'<path d="{d}"'
    if fill: s += f' fill="{fill}"'
    else: s += ' fill="none"'
    if stroke: s += f' stroke="{stroke}" stroke-width="{stroke_width}"'
    s += f' opacity="{opacity}"'
    return s + '/>'


def svg_filter_blur(id_name: str, std_dev: float = 5) -> str:
    return f'''<defs><filter id="{id_name}"><feGaussianBlur stdDeviation="{std_dev}"/></filter></defs>'''


def svg_filter_glow(id_name: str, color: Tuple, std_dev: float = 4) -> str:
    r, g, b = color
    return f'''<defs><filter id="{id_name}"><feGaussianBlur stdDeviation="{std_dev}" result="blur"/>
    <feFlood flood-color="rgb({r},{g},{b})" flood-opacity="0.6" result="color"/>
    <feComposite in="color" in2="blur" operator="in" result="glow"/>
    <feMerge><feMergeNode in="glow"/><feMergeNode in="SourceGraphic"/></feMerge></filter></defs>'''


# ─── Generators ────────────────────────────────────────────────────────────

def _rgb(color: Tuple) -> str:
    return f"rgb({color[0]},{color[1]},{color[2]})"


def _rgba(color: Tuple, a: float) -> str:
    return f"rgba({color[0]},{color[1]},{color[2]},{a})"


def generate_svg_geometric_pattern(width, height, palette, seed=None):
    """Abstract geometric pattern with overlapping shapes."""
    if seed is not None:
        random.seed(seed)
    parts = [svg_header(width, height)]

    # Background
    parts.append(svg_rect(0, 0, width, height, fill=_rgb(palette[0])))

    # Large overlapping circles
    for _ in range(random.randint(5, 10)):
        cx = random.randint(-width//4, width + width//4)
        cy = random.randint(-height//4, height + height//4)
        r = random.randint(100, min(width, height))
        color = random.choice(palette[1:])
        parts.append(svg_circle(cx, cy, r, fill=_rgba(color, 0.15)))

    # Geometric lines
    for _ in range(random.randint(10, 20)):
        x1, y1 = random.randint(0, width), random.randint(0, height)
        angle = random.uniform(0, 2 * math.pi)
        length = random.randint(100, 400)
        x2 = x1 + length * math.cos(angle)
        y2 = y1 + length * math.sin(angle)
        color = random.choice(palette[2:])
        parts.append(f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{_rgb(color)}" stroke-width="{random.randint(1,3)}" opacity="0.4"/>')

    # Small decorative dots
    for _ in range(random.randint(30, 60)):
        x, y = random.randint(0, width), random.randint(0, height)
        r = random.randint(2, 8)
        color = random.choice(palette[1:])
        parts.append(svg_circle(x, y, r, fill=_rgb(color), opacity=0.5))

    parts.append(SVG_FOOTER)
    return "\n".join(parts)


def generate_svg_hex_grid(width, height, palette, seed=None):
    """Honeycomb hexagonal grid pattern."""
    if seed is not None:
        random.seed(seed)
    parts = [svg_header(width, height)]
    parts.append(svg_rect(0, 0, width, height, fill=_rgb(palette[0])))

    hex_size = random.randint(30, 60)
    hex_h = hex_size * math.sqrt(3)

    row = 0
    y = -hex_size
    while y < height + hex_size:
        x_offset = (hex_size * 1.5) * (row % 2)
        x = -hex_size + x_offset
        while x < width + hex_size:
            pts = []
            for i in range(6):
                angle = math.pi / 3 * i - math.pi / 6
                hx = x + hex_size * math.cos(angle)
                hy = y + hex_size * math.sin(angle)
                pts.append((hx, hy))

            color = random.choice(palette[1:])
            fill_opa = random.uniform(0.2, 0.7)
            stroke_color = palette[-1]

            parts.append(svg_polygon(pts, fill=_rgba(color, fill_opa)))
            parts.append(svg_polygon(pts, fill="none"))

            # Outline with slightly transparent stroke
            pts_str = " ".join(f"{px:.1f},{py:.1f}" for px, py in pts)
            parts.append(f'<polygon points="{pts_str}" fill="none" stroke="{_rgb(stroke_color)}" stroke-width="1" opacity="0.3"/>')

            x += hex_size * 3
        y += hex_h
        row += 1

    parts.append(SVG_FOOTER)
    return "\n".join(parts)


def generate_svg_circuit(width, height, palette, seed=None):
    """Tech circuit board pattern."""
    if seed is not None:
        random.seed(seed)
    parts = [svg_header(width, height)]
    parts.append(svg_rect(0, 0, width, height, fill=_rgb(palette[0])))

    # Grid-based circuit paths
    grid = 40
    nodes = set()
    for x in range(0, width + grid, grid):
        for y in range(0, height + grid, grid):
            if random.random() < 0.4:
                nodes.add((x, y))

    # Draw connections
    for (x1, y1) in nodes:
        for (x2, y2) in nodes:
            if (x1, y1) == (x2, y2):
                continue
            dist = math.sqrt((x2-x1)**2 + (y2-y1)**2)
            if dist < grid * 3 and random.random() < 0.3:
                color = random.choice(palette[1:])
                # L-shaped paths (circuit style)
                mid_x, mid_y = x2, y1
                path = f"M {x1} {y1} L {mid_x} {mid_y} L {x2} {y2}"
                parts.append(svg_path(path, stroke=_rgb(color), stroke_width=random.choice([1, 2]), opacity=0.5))

    # Nodes (small circles)
    for (x, y) in nodes:
        color = random.choice(palette[1:])
        parts.append(svg_circle(x, y, 3, fill=_rgb(color), opacity=0.8))
        if random.random() < 0.3:
            parts.append(svg_circle(x, y, 8, fill=_rgba(color, 0.3)))

    # IC chip rectangles
    for _ in range(random.randint(5, 12)):
        x, y = random.choice(list(nodes)) if nodes else (random.randint(0, width), random.randint(0, height))
        w, h = random.randint(20, 60), random.randint(15, 40)
        color = random.choice(palette[2:])
        parts.append(svg_rect(x - w//2, y - h//2, w, h, fill=_rgba(color, 0.4), rx=3))
        # Pins
        for px in range(x - w//2 + 5, x + w//2, 8):
            parts.append(svg_rect(px, y - h//2 - 5, 2, 5, fill=_rgb(palette[-1]), opacity=0.3))
            parts.append(svg_rect(px, y + h//2, 2, 5, fill=_rgb(palette[-1]), opacity=0.3))

    parts.append(SVG_FOOTER)
    return "\n".join(parts)


def generate_svg_wave_pattern(width, height, palette, seed=None):
    """Flowing wave pattern with gradient fills."""
    if seed is not None:
        random.seed(seed)
    parts = [svg_header(width, height)]

    # Define gradients
    grad_defs = "<defs>"
    for i in range(4):
        c1 = palette[i % len(palette)]
        c2 = palette[(i + 1) % len(palette)]
        stops = f'<stop offset="0%" stop-color="{_rgba(c1, 0.6)}"/><stop offset="100%" stop-color="{_rgba(c2, 0.6)}"/>'
        grad_defs += f'<linearGradient id="wg{i}" x1="0%" y1="0%" x2="100%" y2="0%">{stops}</linearGradient>'
    grad_defs += "</defs>"
    parts.append(grad_defs)

    parts.append(svg_rect(0, 0, width, height, fill=_rgb(palette[0])))

    for wave_i in range(5):
        y_base = height * (0.3 + 0.15 * wave_i)
        amp = random.uniform(30, 80)
        freq = random.uniform(0.005, 0.015)
        phase = random.uniform(0, math.pi * 2)

        d = f"M 0 {y_base}"
        for x in range(0, width + 10, 5):
            y = y_base + amp * math.sin(x * freq + phase)
            y += amp * 0.3 * math.sin(x * freq * 2.5 + phase * 1.3)
            d += f" L {x} {y}"
        d += f" L {width} {height} L 0 {height} Z"

        parts.append(svg_path(d, fill=f"url(#wg{wave_i % 4})", opacity=0.6))

    parts.append(SVG_FOOTER)
    return "\n".join(parts)


def generate_svg_abstract_art(width, height, palette, seed=None):
    """Bauhaus/abstract art inspired composition."""
    if seed is not None:
        random.seed(seed)
    parts = [svg_header(width, height)]
    parts.append(svg_rect(0, 0, width, height, fill=_rgb(palette[0])))

    # Large background shapes
    for _ in range(random.randint(3, 6)):
        shape_type = random.choice(["circle", "rect", "arc"])
        color = random.choice(palette[1:])
        opa = random.uniform(0.15, 0.4)

        if shape_type == "circle":
            cx = random.randint(0, width)
            cy = random.randint(0, height)
            r = random.randint(100, min(width, height))
            parts.append(svg_circle(cx, cy, r, fill=_rgba(color, opa)))

        elif shape_type == "rect":
            x = random.randint(-100, width)
            y = random.randint(-100, height)
            w = random.randint(150, width)
            h = random.randint(100, height)
            rot = random.randint(0, 45)
            parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="{_rgba(color, opa)}" transform="rotate({rot} {x+w//2} {y+h//2})"/>')

        elif shape_type == "arc":
            cx, cy = random.randint(0, width), random.randint(0, height)
            r = random.randint(100, 300)
            angle1 = random.uniform(0, math.pi)
            angle2 = angle1 + random.uniform(math.pi * 0.5, math.pi * 1.5)
            x1, y1 = cx + r * math.cos(angle1), cy + r * math.sin(angle1)
            x2, y2 = cx + r * math.cos(angle2), cy + r * math.sin(angle2)
            large = 1 if (angle2 - angle1) > math.pi else 0
            d = f"M {cx} {cy} L {x1} {y1} A {r} {r} 0 {large} 1 {x2} {y2} Z"
            parts.append(svg_path(d, fill=_rgba(color, opa)))

    # Accent lines
    for _ in range(random.randint(5, 15)):
        x1, y1 = random.randint(0, width), random.randint(0, height)
        x2 = x1 + random.randint(-300, 300)
        y2 = y1 + random.randint(-300, 300)
        color = random.choice(palette[2:])
        parts.append(f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{_rgb(color)}" stroke-width="{random.uniform(1,4):.1f}" opacity="0.5"/>')

    # Small dots
    for _ in range(20):
        x, y = random.randint(0, width), random.randint(0, height)
        r = random.randint(3, 12)
        color = random.choice(palette[1:])
        parts.append(svg_circle(x, y, r, fill=_rgb(color), opacity=0.6))

    parts.append(SVG_FOOTER)
    return "\n".join(parts)


def generate_svg_diagonal_stripes(width, height, palette, seed=None):
    """Diagonal stripe pattern with gradients."""
    if seed is not None:
        random.seed(seed)
    parts = [svg_header(width, height)]
    parts.append(svg_rect(0, 0, width, height, fill=_rgb(palette[0])))

    stripe_width = random.randint(20, 60)
    angle = random.uniform(30, 60)

    for i in range(-height, width + height, stripe_width):
        color = palette[(i // stripe_width) % len(palette)]
        x1, y1 = i, 0
        x2, y2 = i + stripe_width, 0
        x3, y3 = i + stripe_width + height, height
        x4, y4 = i + height, height

        # Apply rotation
        rad = math.radians(angle)
        cx, cy = width / 2, height / 2

        def rotate(x, y):
            dx, dy = x - cx, y - cy
            rx = dx * math.cos(rad) - dy * math.sin(rad) + cx
            ry = dx * math.sin(rad) + dy * math.cos(rad) + cy
            return rx, ry

        pts = [rotate(x1, y1), rotate(x2, y2), rotate(x3, y3), rotate(x4, y4)]
        parts.append(svg_polygon(pts, fill=_rgba(color, 0.4)))

    parts.append(SVG_FOOTER)
    return "\n".join(parts)


def generate_svg_radial_sunburst(width, height, palette, seed=None):
    """Radial sunburst pattern."""
    if seed is not None:
        random.seed(seed)
    parts = [svg_header(width, height)]

    # Gradient bg
    parts.append("<defs>")
    c1, c2 = palette[0], palette[1]
    parts.append(f'<radialGradient id="bg"><stop offset="0%" stop-color="{_rgb(c2)}"/><stop offset="100%" stop-color="{_rgb(c1)}"/></radialGradient>')
    parts.append("</defs>")
    parts.append(svg_rect(0, 0, width, height, fill="url(#bg)"))

    cx, cy = width // 2, height // 2
    n_rays = random.randint(12, 36)

    for i in range(n_rays):
        angle1 = 2 * math.pi * i / n_rays
        angle2 = 2 * math.pi * (i + 0.5) / n_rays
        r = max(width, height)

        x1 = cx + r * math.cos(angle1)
        y1 = cy + r * math.sin(angle1)
        x2 = cx + r * math.cos(angle2)
        y2 = cy + r * math.sin(angle2)

        color = palette[(i % (len(palette) - 1)) + 1]
        d = f"M {cx} {cy} L {x1} {y1} A {r} {r} 0 0 1 {x2} {y2} Z"
        parts.append(svg_path(d, fill=_rgba(color, 0.3)))

    # Center circle
    parts.append(svg_circle(cx, cy, 60, fill=_rgba(palette[-1], 0.5)))
    parts.append(svg_circle(cx, cy, 40, fill=_rgb(palette[0])))

    parts.append(SVG_FOOTER)
    return "\n".join(parts)


# ─── Registry & Converter ─────────────────────────────────────────────────

SVG_GENERATORS = {
    "geometric_pattern": generate_svg_geometric_pattern,
    "hex_grid": generate_svg_hex_grid,
    "circuit": generate_svg_circuit,
    "wave_pattern": generate_svg_wave_pattern,
    "abstract_art": generate_svg_abstract_art,
    "diagonal_stripes": generate_svg_diagonal_stripes,
    "radial_sunburst": generate_svg_radial_sunburst,
}


def generate_svg_art(style: str, width: int, height: int, palette: List, seed: int = None) -> str:
    """Generate SVG string by style name."""
    gen = SVG_GENERATORS.get(style)
    if gen is None:
        gen = generate_svg_geometric_pattern
    return gen(width, height, palette, seed=seed)


def svg_to_png(svg_string: str, width: int, height: int, output_path: str = None) -> str:
    """Convert SVG string to PNG using rsvg-convert."""
    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix=".png")
        os.close(fd)

    svg_path = output_path.replace(".png", ".svg")
    with open(svg_path, "w") as f:
        f.write(svg_string)

    try:
        subprocess.run(
            ["rsvg-convert", "-w", str(width), "-h", str(height),
             "-o", output_path, svg_path],
            check=True, capture_output=True, timeout=30
        )
    finally:
        if os.path.exists(svg_path):
            os.unlink(svg_path)

    return output_path
