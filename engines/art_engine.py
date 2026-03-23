"""
PPT Forge v2 — Advanced Generative Art Engine
Produces professional-grade visuals using pure math & algorithms.
Zero paid APIs. Zero external services. Pure creative computation.
"""

import math
import random
import colorsys
from typing import List, Tuple, Optional, Dict
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageChops
import numpy as np


# ─── Color Science ─────────────────────────────────────────────────────────

PALETTES: Dict[str, List[Tuple[int, int, int]]] = {
    "sunset":   [(255, 94, 77), (255, 154, 0), (255, 206, 84), (255, 56, 96), (128, 0, 128)],
    "ocean":    [(0, 31, 63), (0, 116, 217), (127, 219, 255), (170, 240, 209), (46, 204, 64)],
    "forest":   [(11, 83, 41), (46, 139, 87), (144, 195, 119), (199, 220, 164), (255, 250, 205)],
    "neon":     [(255, 0, 128), (0, 255, 255), (255, 255, 0), (128, 0, 255), (0, 255, 128)],
    "pastel":   [(255, 209, 220), (181, 234, 215), (199, 206, 234), (255, 241, 191), (226, 240, 203)],
    "mono":     [(20, 20, 20), (60, 60, 60), (100, 100, 100), (160, 160, 160), (220, 220, 220)],
    "warm":     [(193, 66, 66), (203, 100, 53), (211, 141, 53), (218, 180, 70), (227, 217, 93)],
    "cool":     [(11, 36, 71), (29, 66, 113), (48, 104, 141), (98, 162, 177), (177, 215, 208)],
    "fire":     [(255, 0, 0), (255, 80, 0), (255, 160, 0), (255, 220, 50), (255, 255, 150)],
    "aurora":   [(0, 255, 128), (0, 200, 255), (128, 0, 255), (255, 0, 200), (0, 255, 200)],
    "midnight": [(10, 10, 35), (25, 25, 80), (50, 50, 140), (100, 80, 180), (180, 140, 255)],
    "sakura":   [(255, 183, 197), (255, 150, 170), (240, 120, 150), (220, 90, 130), (200, 60, 110)],
    "lavender": [(75, 0, 130), (138, 43, 226), (186, 85, 211), (218, 112, 214), (255, 182, 193)],
    "emerald":  [(0, 48, 32), (0, 100, 60), (0, 163, 108), (80, 200, 120), (144, 238, 144)],
    "crimson":  [(60, 0, 0), (120, 0, 0), (180, 20, 20), (220, 60, 60), (255, 120, 100)],
    "arctic":   [(200, 230, 255), (150, 200, 255), (100, 170, 255), (50, 130, 240), (0, 90, 200)],
    "golden":   [(139, 90, 0), (184, 134, 11), (218, 165, 32), (238, 200, 80), (255, 230, 130)],
    "cosmic":   [(10, 0, 30), (40, 0, 80), (100, 20, 150), (180, 50, 200), (255, 100, 255)],
    "mint":     [(0, 60, 50), (0, 110, 90), (0, 160, 130), (100, 200, 170), (180, 240, 220)],
    "coral":    [(100, 20, 20), (180, 60, 50), (230, 100, 70), (255, 150, 100), (255, 200, 160)],
}


def get_palette(name: str) -> List[Tuple[int, int, int]]:
    return PALETTES.get(name, PALETTES["sunset"])


def hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    return f"#{r:02x}{g:02x}{b:02x}"


def lerp_color(c1: Tuple, c2: Tuple, t: float) -> Tuple[int, int, int]:
    t = max(0.0, min(1.0, t))
    return tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


def palette_color(palette: List, t: float) -> Tuple[int, int, int]:
    """Sample a continuous color from palette at position t (0-1)."""
    t = max(0, min(1, t))
    idx = t * (len(palette) - 1)
    i = int(idx)
    if i >= len(palette) - 1:
        return palette[-1]
    return lerp_color(palette[i], palette[i + 1], idx - i)


def adjust_brightness(color: Tuple, factor: float) -> Tuple[int, int, int]:
    """Adjust brightness: factor > 1 = lighter, < 1 = darker."""
    return tuple(max(0, min(255, int(c * factor))) for c in color)


def complementary(color: Tuple) -> Tuple[int, int, int]:
    """Get complementary color."""
    r, g, b = color
    h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
    h = (h + 0.5) % 1.0
    r2, g2, b2 = colorsys.hsv_to_rgb(h, s, v)
    return (int(r2 * 255), int(g2 * 255), int(b2 * 255))


def analogous_colors(color: Tuple, n: int = 3, spread: float = 0.08) -> List[Tuple[int, int, int]]:
    """Generate analogous colors."""
    r, g, b = color
    h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
    colors = []
    for i in range(n):
        offset = (i - n // 2) * spread
        new_h = (h + offset) % 1.0
        r2, g2, b2 = colorsys.hsv_to_rgb(new_h, s, v)
        colors.append((int(r2 * 255), int(g2 * 255), int(b2 * 255)))
    return colors


def with_alpha(color: Tuple, alpha: int) -> Tuple:
    """Add alpha channel to RGB."""
    return color + (alpha,) if len(color) == 3 else color[:3] + (alpha,)


# ─── Noise Generators ──────────────────────────────────────────────────────

class PerlinNoise:
    """Optimized 2D Perlin noise with octaves."""
    def __init__(self, seed=None):
        if seed is not None:
            random.seed(seed)
        self.perm = list(range(256))
        random.shuffle(self.perm)
        self.perm *= 2
        self._grad2 = [(1,1),(-1,1),(1,-1),(-1,-1),(1,0),(-1,0),(0,1),(0,-1)]

    def _fade(self, t):
        return t * t * t * (t * (t * 6 - 15) + 10)

    def _dot2(self, g, x, y):
        return g[0] * x + g[1] * y

    def noise(self, x, y):
        xi = int(math.floor(x)) & 255
        yi = int(math.floor(y)) & 255
        xf = x - math.floor(x)
        yf = y - math.floor(y)
        u = self._fade(xf)
        v = self._fade(yf)
        aa = self.perm[self.perm[xi] + yi] & 7
        ab = self.perm[self.perm[xi] + yi + 1] & 7
        ba = self.perm[self.perm[xi + 1] + yi] & 7
        bb = self.perm[self.perm[xi + 1] + yi + 1] & 7
        x1 = (self._dot2(self._grad2[aa], xf, yf) * (1 - u) +
              self._dot2(self._grad2[ba], xf - 1, yf) * u)
        x2 = (self._dot2(self._grad2[ab], xf, yf - 1) * (1 - u) +
              self._dot2(self._grad2[bb], xf - 1, yf - 1) * u)
        return x1 * (1 - v) + x2 * v

    def octave_noise(self, x, y, octaves=6, persistence=0.5, lacunarity=2.0):
        total = 0
        amplitude = 1
        frequency = 1
        max_val = 0
        for _ in range(octaves):
            total += self.noise(x * frequency, y * frequency) * amplitude
            max_val += amplitude
            amplitude *= persistence
            frequency *= lacunarity
        return total / max_val


class SimplexNoise:
    """2D Simplex noise — smoother than Perlin for organic textures."""
    F2 = 0.5 * (math.sqrt(3) - 1)
    G2 = (3 - math.sqrt(3)) / 6

    def __init__(self, seed=None):
        if seed is not None:
            random.seed(seed)
        self.perm = list(range(256))
        random.shuffle(self.perm)
        self.perm *= 2
        self._grad3 = [(1,1,0),(-1,1,0),(1,-1,0),(-1,-1,0),
                       (1,0,1),(-1,0,1),(1,0,-1),(-1,0,-1),
                       (0,1,1),(0,-1,1),(0,1,-1),(0,-1,-1)]

    def noise(self, x, y):
        s = (x + y) * self.F2
        i = int(math.floor(x + s))
        j = int(math.floor(y + s))
        t = (i + j) * self.G2
        x0, y0 = x - (i - t), y - (j - t)
        i1, j1 = (1, 0) if x0 > y0 else (0, 1)
        x1, y1 = x0 - i1 + self.G2, y0 - j1 + self.G2
        x2, y2 = x0 - 1 + 2 * self.G2, y0 - 1 + 2 * self.G2
        ii, jj = i & 255, j & 255

        def contrib(gi, xc, yc):
            t = 0.5 - xc * xc - yc * yc
            if t < 0:
                return 0
            t *= t
            g = self._grad3[self.perm[ii + self.perm[jj + gi]] % 12]
            return t * t * (g[0] * xc + g[1] * yc)

        n0 = contrib(0, x0, y0)
        n1 = contrib(i1 + j1 * 2, x1, y1)
        n2 = contrib(1, x2, y2)
        return 70 * (n0 + n1 + n2)

    def octave_noise(self, x, y, octaves=6, persistence=0.5, lacunarity=2.0):
        total = 0
        amplitude = 1
        frequency = 1
        max_val = 0
        for _ in range(octaves):
            total += self.noise(x * frequency, y * frequency) * amplitude
            max_val += amplitude
            amplitude *= persistence
            frequency *= lacunarity
        return total / max_val


class WorleyNoise:
    """Worley (cellular) noise — great for organic/cellular patterns."""
    def __init__(self, seed=None):
        if seed is not None:
            random.seed(seed)
        self.points = []

    def generate_points(self, width, height, n_points=50):
        self.points = [(random.uniform(0, width), random.uniform(0, height))
                        for _ in range(n_points)]

    def noise(self, x, y):
        if not self.points:
            return 0
        dists = sorted([math.sqrt((x - px)**2 + (y - py)**2)
                        for px, py in self.points])
        return dists[0]

    def noise_f2(self, x, y):
        """Return F2 - F1 (edge detection pattern)."""
        if len(self.points) < 2:
            return 0
        dists = sorted([math.sqrt((x - px)**2 + (y - py)**2)
                        for px, py in self.points])
        return dists[1] - dists[0]


# ─── Art Generators ────────────────────────────────────────────────────────

def generate_gradient_mesh(width, height, palette_name="sunset", seed=None, **kw):
    """Smooth multi-point mesh gradient (Apple iOS wallpaper style)."""
    if seed is not None:
        np.random.seed(seed)
        random.seed(seed)
    palette = get_palette(palette_name)
    img = np.zeros((height, width, 3), dtype=np.float64)

    n_points = random.randint(5, 10)
    points = [(random.randint(0, width - 1), random.randint(0, height - 1)) for _ in range(n_points)]
    colors = [palette[i % len(palette)] for i in range(n_points)]

    ys, xs = np.mgrid[0:height, 0:width]
    total_weight = np.zeros((height, width))

    for (px, py), color in zip(points, colors):
        dist = np.sqrt((xs - px) ** 2 + (ys - py) ** 2) + 1.0
        weight = 1.0 / (dist ** 2.8)
        for c in range(3):
            img[:, :, c] += color[c] * weight
        total_weight += weight

    for c in range(3):
        img[:, :, c] /= total_weight

    return Image.fromarray(np.clip(img, 0, 255).astype(np.uint8))


def generate_perlin_noise(width, height, palette_name="ocean", seed=None, **kw):
    """Organic Perlin noise texture — vectorized for speed."""
    scale = kw.get('scale', 0.012)
    octaves = kw.get('octaves', 7)
    pn = PerlinNoise(seed)
    palette = get_palette(palette_name)

    step = 3
    small_w = width // step + 2
    small_h = height // step + 2
    small = np.zeros((small_h, small_w))

    for y in range(small_h):
        for x in range(small_w):
            small[y, x] = pn.octave_noise(x * step * scale, y * step * scale, octaves)

    small_img = Image.fromarray(((small + 1) / 2 * 255).astype(np.uint8), mode='L')
    noise_img = small_img.resize((width, height), Image.LANCZOS)
    noise_arr = np.array(noise_img, dtype=np.float64) / 255.0

    result = np.zeros((height, width, 3), dtype=np.uint8)
    n_colors = len(palette)

    for i in range(n_colors - 1):
        lo = i / (n_colors - 1)
        hi = (i + 1) / (n_colors - 1)
        mask = (noise_arr >= lo) & (noise_arr < hi)
        if not np.any(mask):
            continue
        t_local = (noise_arr[mask] - lo) / (hi - lo)
        for c in range(3):
            result[:, :, c][mask] = (palette[i][c] + (palette[i + 1][c] - palette[i][c]) * t_local).astype(np.uint8)

    mask_top = noise_arr >= 1.0
    for c in range(3):
        result[:, :, c][mask_top] = palette[-1][c]

    return Image.fromarray(result)


def generate_simplex_noise(width, height, palette_name="aurora", seed=None, **kw):
    """Simplex noise — smoother organic textures than Perlin."""
    scale = kw.get('scale', 0.008)
    octaves = kw.get('octaves', 7)
    sn = SimplexNoise(seed)
    palette = get_palette(palette_name)

    step = 3
    small_w = width // step + 2
    small_h = height // step + 2
    small = np.zeros((small_h, small_w))

    for y in range(small_h):
        for x in range(small_w):
            small[y, x] = sn.octave_noise(x * step * scale, y * step * scale, octaves)

    small_img = Image.fromarray(((small + 1) / 2 * 255).astype(np.uint8), mode='L')
    noise_img = small_img.resize((width, height), Image.LANCZOS)
    noise_arr = np.array(noise_img, dtype=np.float64) / 255.0

    result = np.zeros((height, width, 3), dtype=np.uint8)
    n_colors = len(palette)

    for i in range(n_colors - 1):
        lo = i / (n_colors - 1)
        hi = (i + 1) / (n_colors - 1)
        mask = (noise_arr >= lo) & (noise_arr < hi)
        if not np.any(mask):
            continue
        t_local = (noise_arr[mask] - lo) / (hi - lo)
        for c in range(3):
            result[:, :, c][mask] = (palette[i][c] + (palette[i + 1][c] - palette[i][c]) * t_local).astype(np.uint8)

    return Image.fromarray(result)


def generate_voronoi(width, height, palette_name="neon", seed=None, **kw):
    """Voronoi cell tessellation with smooth edges and glow."""
    if seed is not None:
        np.random.seed(seed)
        random.seed(seed)
    n_cells = kw.get('n_cells', 35)
    palette = get_palette(palette_name)

    pts = np.array([(random.randint(0, width), random.randint(0, height)) for _ in range(n_cells)])
    colors = [palette[i % len(palette)] for i in range(n_cells)]

    ys, xs = np.mgrid[0:height, 0:width]
    img = np.zeros((height, width, 3), dtype=np.uint8)

    min_dist = np.full((height, width), np.inf)
    nearest = np.zeros((height, width), dtype=int)

    for i, (px, py) in enumerate(pts):
        dist = (xs - px) ** 2 + (ys - py) ** 2
        closer = dist < min_dist
        min_dist[closer] = dist[closer]
        nearest[closer] = i

    for i, color in enumerate(colors):
        mask = nearest == i
        for c in range(3):
            img[:, :, c][mask] = color[c]

    # Soft edge glow
    img_pil = Image.fromarray(img)
    edge_mask = np.zeros((height, width), dtype=bool)
    shifted_funcs = [
        lambda n: np.roll(n, 1, axis=0),
        lambda n: np.roll(n, -1, axis=0),
        lambda n: np.roll(n, 1, axis=1),
        lambda n: np.roll(n, -1, axis=1),
    ]
    for fn in shifted_funcs:
        shifted = fn(nearest)
        edge_mask |= (nearest != shifted)

    edge_coords = np.argwhere(edge_mask)
    for y, x in edge_coords:
        if 0 <= y < height and 0 <= x < width:
            img_pil.putpixel((x, y), (255, 255, 255))

    return img_pil.filter(ImageFilter.GaussianBlur(radius=0.5))


def generate_flow_field(width, height, palette_name="aurora", seed=None, **kw):
    """Particle flow field — organic, dynamic, professional."""
    if seed is not None:
        random.seed(seed)
        np.random.seed(seed)
    n_particles = kw.get('n_particles', 4000)
    steps = kw.get('steps', 100)
    palette = get_palette(palette_name)
    pn = PerlinNoise(seed)

    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    scale = 0.004

    for _ in range(n_particles):
        x = random.uniform(0, width)
        y = random.uniform(0, height)
        color = random.choice(palette)
        alpha = random.randint(8, 45)
        points = []

        for _ in range(steps):
            if 0 <= x < width and 0 <= y < height:
                points.append((x, y))
                angle = pn.octave_noise(x * scale, y * scale, 4) * math.pi * 4
                x += math.cos(angle) * 2.5
                y += math.sin(angle) * 2.5
            else:
                break

        if len(points) > 1:
            draw.line(points, fill=with_alpha(color, alpha), width=1)

    bg = Image.new("RGB", (width, height), palette[0])
    bg.paste(img, (0, 0), img)
    return bg


def generate_geometric(width, height, palette_name="creative", seed=None, **kw):
    """Geometric tessellation with layered depth and transparency."""
    if seed is not None:
        random.seed(seed)
    palette = get_palette(palette_name)

    bg_color = palette[0]
    img = Image.new("RGB", (width, height), bg_color)
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    cx, cy = width // 2, height // 2
    max_r = int(math.sqrt(cx**2 + cy**2))

    # Concentric polygons
    for r in range(max_r, 20, -12):
        color = palette_color(palette, r / max_r)
        sides = random.choice([3, 4, 5, 6, 8])
        angle_offset = random.uniform(0, math.pi)
        pts = []
        for i in range(sides):
            angle = 2 * math.pi * i / sides + angle_offset
            pts.append((cx + r * math.cos(angle), cy + r * math.sin(angle)))
        alpha = random.randint(20, 60)
        draw.polygon(pts, fill=with_alpha(color, alpha))

    # Scattered circles
    for _ in range(30):
        x, y = random.randint(0, width), random.randint(0, height)
        r = random.randint(5, 40)
        color = random.choice(palette[1:])
        alpha = random.randint(30, 80)
        draw.ellipse([x - r, y - r, x + r, y + r], fill=with_alpha(color, alpha))

    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    return img


def generate_fractal_tree(width, height, palette_name="forest", seed=None, **kw):
    """Recursive fractal branching pattern."""
    if seed is not None:
        random.seed(seed)
    palette = get_palette(palette_name)

    img = Image.new("RGB", (width, height), palette[0])
    draw = ImageDraw.Draw(img)

    def branch(x, y, angle, length, depth, width_line=3):
        if depth <= 0 or length < 3:
            return
        x2 = x + length * math.cos(angle)
        y2 = y + length * math.sin(angle)
        color = palette_color(palette, 1 - depth / 10)
        draw.line([(x, y), (x2, y2)], fill=color, width=max(1, width_line))

        spread = random.uniform(0.3, 0.7)
        shrink = random.uniform(0.6, 0.8)
        branch(x2, y2, angle - spread, length * shrink, depth - 1, max(1, width_line - 1))
        branch(x2, y2, angle + spread, length * shrink, depth - 1, max(1, width_line - 1))

    start_x = width // 2
    start_y = int(height * 0.85)
    branch(start_x, start_y, -math.pi / 2, height * 0.22, 10, 4)

    # Glow effect
    glow = img.filter(ImageFilter.GaussianBlur(radius=8))
    img = Image.blend(img, glow, alpha=0.3)
    return img


def generate_bokeh(width, height, palette_name="neon", seed=None, **kw):
    """Soft bokeh circles with depth-of-field effect."""
    if seed is not None:
        random.seed(seed)
    palette = get_palette(palette_name)

    bg_color = adjust_brightness(palette[0], 0.3)
    img = Image.new("RGB", (width, height), bg_color)

    n_circles = random.randint(20, 50)
    circles_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))

    for _ in range(n_circles):
        x = random.randint(-50, width + 50)
        y = random.randint(-50, height + 50)
        r = random.randint(30, 200)
        color = random.choice(palette)
        alpha = random.randint(15, 50)

        circle = Image.new("RGBA", (r * 2, r * 2), (0, 0, 0, 0))
        draw_c = ImageDraw.Draw(circle)
        draw_c.ellipse([0, 0, r * 2 - 1, r * 2 - 1], fill=with_alpha(color, alpha))

        # Inner glow
        inner_r = int(r * 0.7)
        offset = r - inner_r
        draw_c.ellipse([offset, offset, offset + inner_r * 2 - 1, offset + inner_r * 2 - 1],
                       fill=with_alpha(adjust_brightness(color, 1.5), alpha // 2))

        circle = circle.filter(ImageFilter.GaussianBlur(radius=r // 3))
        circles_layer.paste(circle, (x - r, y - r), circle)

    # Composite
    img = Image.alpha_composite(img.convert("RGBA"), circles_layer).convert("RGB")

    # Extra soft blur for dreamy effect
    blur = img.filter(ImageFilter.GaussianBlur(radius=6))
    img = Image.blend(img, blur, alpha=0.2)
    return img


def generate_wave(width, height, palette_name="ocean", seed=None, **kw):
    """Layered sine waves with gradient colors."""
    if seed is not None:
        random.seed(seed)
    palette = get_palette(palette_name)

    bg_color = palette[0]
    img = Image.new("RGB", (width, height), adjust_brightness(bg_color, 0.5))

    n_waves = random.randint(6, 12)
    for w in range(n_waves):
        t = w / n_waves
        y_base = int(height * (0.2 + 0.7 * t))
        amplitude = random.uniform(20, 60)
        frequency = random.uniform(0.003, 0.01)
        phase = random.uniform(0, math.pi * 2)
        color = palette_color(palette, t)
        alpha = random.randint(40, 100)

        wave_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw_w = ImageDraw.Draw(wave_layer)

        points = []
        for x in range(0, width, 2):
            y = y_base + amplitude * math.sin(x * frequency + phase)
            y += amplitude * 0.3 * math.sin(x * frequency * 2.3 + phase * 1.5)
            points.append((x, int(y)))

        # Fill below wave
        if points:
            fill_points = points + [(width, height), (0, height)]
            draw_w.polygon(fill_points, fill=with_alpha(color, alpha))

        img = Image.alpha_composite(img.convert("RGBA"), wave_layer).convert("RGB")

    return img


def generate_constellation(width, height, palette_name="midnight", seed=None, **kw):
    """Connected star/dot pattern with glow."""
    if seed is not None:
        random.seed(seed)
    palette = get_palette(palette_name)

    img = Image.new("RGB", (width, height), palette[0])
    draw = ImageDraw.Draw(img)

    n_stars = random.randint(60, 120)
    stars = [(random.randint(0, width), random.randint(0, height)) for _ in range(n_stars)]

    # Connection lines
    max_dist = min(width, height) * 0.15
    for i in range(len(stars)):
        for j in range(i + 1, len(stars)):
            dist = math.sqrt((stars[i][0] - stars[j][0])**2 + (stars[i][1] - stars[j][1])**2)
            if dist < max_dist:
                alpha = max(5, int(30 * (1 - dist / max_dist)))
                color = palette_color(palette, dist / max_dist)
                draw.line([stars[i], stars[j]], fill=adjust_brightness(color, 0.6), width=1)

    # Stars with glow
    for x, y in stars:
        r = random.randint(1, 4)
        color = random.choice(palette[2:])
        draw.ellipse([x - r, y - r, x + r, y + r], fill=color)

    # Glow on bright stars
    glow = img.filter(ImageFilter.GaussianBlur(radius=3))
    img = Image.blend(img, glow, alpha=0.25)
    draw = ImageDraw.Draw(img)
    for x, y in random.sample(stars, min(25, len(stars))):
        r = random.randint(2, 5)
        draw.ellipse([x - r, y - r, x + r, y + r], fill=(255, 255, 255))

    return img


def generate_glass_morphism(width, height, palette_name="cool", seed=None, **kw):
    """Frosted glass layered shapes on gradient base."""
    if seed is not None:
        random.seed(seed)
    palette = get_palette(palette_name)

    img = generate_gradient_mesh(width, height, palette_name, seed)
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    for _ in range(random.randint(4, 8)):
        x = random.randint(0, width)
        y = random.randint(0, height)
        w = random.randint(150, 500)
        h = random.randint(100, 400)
        color = (255, 255, 255, random.randint(15, 50))
        r = random.randint(20, 50)
        draw.rounded_rectangle([x, y, x + w, y + h], radius=r, fill=color)

    overlay = overlay.filter(ImageFilter.GaussianBlur(radius=20))
    img.paste(overlay, (0, 0), overlay)
    return img


def generate_marble(width, height, palette_name="mono", seed=None, **kw):
    """Marble/stone texture using layered noise."""
    if seed is not None:
        random.seed(seed)
        np.random.seed(seed)
    palette = get_palette(palette_name)
    pn = PerlinNoise(seed)
    sn = SimplexNoise(seed) if seed else SimplexNoise(random.randint(1, 999999))

    step = 2
    small_w = width // step + 2
    small_h = height // step + 2
    noise1 = np.zeros((small_h, small_w))
    noise2 = np.zeros((small_h, small_w))

    for y in range(small_h):
        for x in range(small_w):
            n1 = pn.octave_noise(x * step * 0.008, y * step * 0.008, 8)
            noise1[y, x] = n1
            n2 = sn.octave_noise(x * step * 0.02, y * step * 0.02, 4)
            noise2[y, x] = n2

    # Marble: abs(sin(y * freq + noise))
    noise1_img = Image.fromarray(((noise1 + 1) / 2 * 255).astype(np.uint8), mode='L')
    noise1_full = np.array(noise1_img.resize((width, height), Image.LANCZOS), dtype=np.float64)

    ys_arr = np.arange(height).reshape(-1, 1) * np.ones((1, width))
    marble = np.sin(ys_arr * 0.05 + noise1_full * 8)
    marble = (marble + 1) / 2

    # Map to palette
    result = np.zeros((height, width, 3), dtype=np.uint8)
    for i in range(len(palette) - 1):
        lo = i / (len(palette) - 1)
        hi = (i + 1) / (len(palette) - 1)
        mask = (marble >= lo) & (marble < hi)
        t_local = (marble[mask] - lo) / (hi - lo)
        for c in range(3):
            result[:, :, c][mask] = (palette[i][c] + (palette[i+1][c] - palette[i][c]) * t_local).astype(np.uint8)

    return Image.fromarray(result)


def generate_isometric_grid(width, height, palette_name="tech", seed=None, **kw):
    """Isometric 3D grid pattern with depth."""
    if seed is not None:
        random.seed(seed)
    palette = get_palette(palette_name)

    img = Image.new("RGB", (width, height), palette[0])
    draw = ImageDraw.Draw(img)

    tile_w = 60
    tile_h = 35
    depth = 20

    for row in range(-2, height // tile_h + 3):
        for col in range(-2, width // (tile_w) + 3):
            x = col * tile_w + (row % 2) * (tile_w // 2)
            y = row * tile_h

            if random.random() < 0.3:
                continue

            color = random.choice(palette[1:])
            alpha_idx = random.randint(0, len(palette) - 1)
            c = palette[alpha_idx]

            # Top face
            pts_top = [(x, y - depth), (x + tile_w // 2, y - tile_h // 2 - depth),
                       (x + tile_w, y - depth), (x + tile_w // 2, y + tile_h // 2 - depth)]
            draw.polygon(pts_top, fill=adjust_brightness(c, 1.2), outline=palette[-1])

            # Left face
            pts_left = [(x, y - depth), (x + tile_w // 2, y + tile_h // 2 - depth),
                        (x + tile_w // 2, y + tile_h // 2), (x, y)]
            draw.polygon(pts_left, fill=adjust_brightness(c, 0.7))

            # Right face
            pts_right = [(x + tile_w // 2, y + tile_h // 2 - depth), (x + tile_w, y - depth),
                         (x + tile_w, y), (x + tile_w // 2, y + tile_h // 2)]
            draw.polygon(pts_right, fill=adjust_brightness(c, 0.5))

    return img


def generate_topographic(width, height, palette_name="earth", seed=None, **kw):
    """Topographic contour map pattern."""
    if seed is not None:
        random.seed(seed)
    pn = PerlinNoise(seed)
    palette = get_palette(palette_name)

    img = Image.new("RGB", (width, height), palette[0])
    draw = ImageDraw.Draw(img)

    scale = 0.006
    step = 2
    levels = 15

    for y in range(0, height, step):
        for x in range(0, width, step):
            n = pn.octave_noise(x * scale, y * scale, 6)
            n = (n + 1) / 2  # 0-1

            # Determine contour level
            level = int(n * levels)
            t = level / levels
            color = palette_color(palette, t)

            # Check if this is a contour edge
            n_right = pn.octave_noise((x + step) * scale, y * scale, 6)
            n_right = (n_right + 1) / 2
            level_right = int(n_right * levels)

            if level != level_right:
                draw.rectangle([x, y, x + step - 1, y + step - 1], fill=adjust_brightness(color, 0.6))
            else:
                draw.rectangle([x, y, x + step - 1, y + step - 1], fill=color)

    return img


def generate_double_exposure(width, height, palette_name="sunset", seed=None, **kw):
    """Double-exposure effect with layered silhouettes."""
    if seed is not None:
        random.seed(seed)
    palette = get_palette(palette_name)

    # Base gradient
    base = generate_gradient_mesh(width, height, palette_name, seed)

    # Mountain silhouette
    mountain = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(mountain)

    for layer in range(3):
        points = [(0, height)]
        n_peaks = random.randint(3, 7)
        for i in range(n_peaks + 1):
            x = int(width * i / n_peaks)
            peak_h = random.uniform(0.3, 0.7) + layer * 0.1
            y = int(height * (1 - peak_h))
            points.append((x, y))
        points.append((width, height))

        alpha = 60 - layer * 15
        color = palette_color(palette, layer / 3)
        draw.polygon(points, fill=with_alpha(color, max(20, alpha)))

    # Tree silhouettes
    for _ in range(15):
        tx = random.randint(0, width)
        ty = int(height * random.uniform(0.5, 0.85))
        tree_h = random.randint(40, 120)
        trunk_w = random.randint(3, 8)
        draw.rectangle([tx - trunk_w // 2, ty, tx + trunk_w // 2, ty + tree_h // 3],
                        fill=(0, 0, 0, 100))
        # Crown
        for crown in range(3):
            cy = ty - crown * tree_h // 6
            cr = int(tree_h * 0.3 * (1 - crown * 0.2))
            draw.ellipse([tx - cr, cy - cr, tx + cr, cy + cr],
                          fill=with_alpha(palette[0], 80))

    base.paste(mountain, (0, 0), mountain)
    return base


# ─── Registry ──────────────────────────────────────────────────────────────

GENERATORS = {
    "gradient_mesh":    generate_gradient_mesh,
    "perlin_noise":     generate_perlin_noise,
    "simplex_noise":    generate_simplex_noise,
    "voronoi":          generate_voronoi,
    "flow_field":       generate_flow_field,
    "geometric":        generate_geometric,
    "fractal_tree":     generate_fractal_tree,
    "bokeh":            generate_bokeh,
    "wave":             generate_wave,
    "constellation":    generate_constellation,
    "glass_morphism":   generate_glass_morphism,
    "marble":           generate_marble,
    "isometric_grid":   generate_isometric_grid,
    "topographic":      generate_topographic,
    "double_exposure":  generate_double_exposure,
}


def generate_art(style: str, width: int, height: int, palette_name: str = "sunset",
                  seed: int = None, **kwargs) -> Image.Image:
    """Generate art by style name. Falls back to gradient_mesh if unknown."""
    gen = GENERATORS.get(style)
    if gen is None:
        gen = generate_gradient_mesh
    return gen(width, height, palette_name=palette_name, seed=seed, **kwargs)
