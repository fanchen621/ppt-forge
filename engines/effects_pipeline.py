"""
PPT Forge v2 — Visual Effects Pipeline
Professional-grade image effects using Pillow + numpy.
Mimics high-end design tools: blur, glow, grain, chromatic aberration,
glass morphism, vignette, duotone, halftone, and more.
"""

import math
import random
from typing import Tuple, Optional
from PIL import Image, ImageDraw, ImageFilter, ImageChops, ImageEnhance
import numpy as np


def apply_effects(img: Image.Image, effects: list) -> Image.Image:
    """Apply a chain of effects to an image. Returns processed image."""
    result = img.copy()
    for effect in effects:
        if isinstance(effect, str):
            effect = {"type": effect}
        effect_type = effect.get("type", "")
        handler = EFFECTS_REGISTRY.get(effect_type)
        if handler:
            result = handler(result, **effect)
    return result


# ─── Effect Implementations ───────────────────────────────────────────────

def effect_blur(img: Image.Image, radius: float = 5, **kw) -> Image.Image:
    """Gaussian blur."""
    return img.filter(ImageFilter.GaussianBlur(radius=radius))


def effect_box_blur(img: Image.Image, radius: int = 5, **kw) -> Image.Image:
    """Box blur — faster, more uniform."""
    return img.filter(ImageFilter.BoxBlur(radius))


def effect_glow(img: Image.Image, radius: float = 8, intensity: float = 0.4,
                 color: Tuple = None, **kw) -> Image.Image:
    """Outer glow effect."""
    glow = img.filter(ImageFilter.GaussianBlur(radius=radius))
    if color:
        overlay = Image.new("RGB", img.size, color)
        glow = Image.blend(glow, overlay, alpha=0.3)
    return Image.blend(img, glow, alpha=intensity)


def effect_vignette(img: Image.Image, intensity: float = 0.6, **kw) -> Image.Image:
    """Radial vignette darkening effect."""
    w, h = img.size
    cx, cy = w / 2, h / 2
    max_dist = math.sqrt(cx**2 + cy**2)

    # Create radial gradient mask
    Y, X = np.ogrid[:h, :w]
    dist = np.sqrt((X - cx)**2 + (Y - cy)**2)
    vignette = np.clip(1 - (dist / max_dist) * intensity * 1.5, 0, 1)
    vignette = np.dstack([vignette] * 3)

    arr = np.array(img, dtype=np.float64)
    arr *= vignette
    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8))


def effect_grain(img: Image.Image, amount: float = 0.15, **kw) -> Image.Image:
    """Film grain effect."""
    arr = np.array(img, dtype=np.float64)
    noise = np.random.normal(0, amount * 255, arr.shape)
    arr += noise
    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8))


def effect_chromatic_aberration(img: Image.Image, shift: int = 3, **kw) -> Image.Image:
    """RGB channel shift for chromatic aberration effect."""
    if img.mode != "RGB":
        img = img.convert("RGB")
    r, g, b = img.split()

    # Shift red channel right, blue channel left
    r_arr = np.array(r)
    b_arr = np.array(b)

    r_shifted = np.zeros_like(r_arr)
    b_shifted = np.zeros_like(b_arr)

    if shift > 0:
        r_shifted[:, shift:] = r_arr[:, :-shift]
        b_shifted[:, :-shift] = b_arr[:, shift:]
    else:
        shift = abs(shift)
        r_shifted[:, :-shift] = r_arr[:, shift:]
        b_shifted[:, shift:] = b_arr[:, :-shift]

    r = Image.fromarray(r_shifted)
    b = Image.fromarray(b_shifted)

    return Image.merge("RGB", (r, g, b))


def effect_duotone(img: Image.Image, color1: Tuple = (30, 0, 60),
                    color2: Tuple = (255, 100, 150), **kw) -> Image.Image:
    """Duotone color mapping (two-color gradient)."""
    gray = img.convert("L")
    arr = np.array(gray, dtype=np.float64) / 255.0

    result = np.zeros((*arr.shape, 3), dtype=np.uint8)
    for c in range(3):
        result[:, :, c] = (color1[c] * (1 - arr) + color2[c] * arr).astype(np.uint8)

    return Image.fromarray(result)


def effect_halftone(img: Image.Image, dot_spacing: int = 8, **kw) -> Image.Image:
    """Halftone dot pattern effect."""
    gray = img.convert("L")
    w, h = img.size

    result = Image.new("RGB", (w, h), (255, 255, 255))
    draw = ImageDraw.Draw(result)
    pixels = gray.load()

    for y in range(0, h, dot_spacing):
        for x in range(0, w, dot_spacing):
            brightness = pixels[x, y] / 255.0
            radius = (1 - brightness) * dot_spacing * 0.5
            if radius > 0.5:
                draw.ellipse(
                    [x - radius, y - radius, x + radius, y + radius],
                    fill=(0, 0, 0)
                )

    return result


def effect_pixelate(img: Image.Image, pixel_size: int = 10, **kw) -> Image.Image:
    """Pixelation effect."""
    w, h = img.size
    small = img.resize((w // pixel_size, h // pixel_size), Image.NEAREST)
    return small.resize((w, h), Image.NEAREST)


def effect_sharpen(img: Image.Image, **kw) -> Image.Image:
    """Sharpen filter."""
    return img.filter(ImageFilter.SHARPEN)


def effect_unsharp_mask(img: Image.Image, radius: float = 2, percent: int = 150,
                         threshold: int = 3, **kw) -> Image.Image:
    """Unsharp mask for controlled sharpening."""
    return img.filter(ImageFilter.UnsharpMask(radius=radius, percent=percent, threshold=threshold))


def effect_emboss(img: Image.Image, **kw) -> Image.Image:
    """Emboss/relief effect."""
    return img.filter(ImageFilter.EMBOSS)


def effect_edge_enhance(img: Image.Image, **kw) -> Image.Image:
    """Edge enhancement."""
    return img.filter(ImageFilter.EDGE_ENHANCE_MORE)


def effect_contrast(img: Image.Image, factor: float = 1.5, **kw) -> Image.Image:
    """Contrast adjustment."""
    return ImageEnhance.Contrast(img).enhance(factor)


def effect_brightness(img: Image.Image, factor: float = 1.2, **kw) -> Image.Image:
    """Brightness adjustment."""
    return ImageEnhance.Brightness(img).enhance(factor)


def effect_saturation(img: Image.Image, factor: float = 1.5, **kw) -> Image.Image:
    """Color saturation adjustment."""
    return ImageEnhance.Color(img).enhance(factor)


def effect_sepia(img: Image.Image, intensity: float = 0.6, **kw) -> Image.Image:
    """Sepia tone effect."""
    if img.mode != "RGB":
        img = img.convert("RGB")
    arr = np.array(img, dtype=np.float64)

    # Sepia matrix
    sepia_r = arr[:,:,0] * 0.393 + arr[:,:,1] * 0.769 + arr[:,:,2] * 0.189
    sepia_g = arr[:,:,0] * 0.349 + arr[:,:,1] * 0.686 + arr[:,:,2] * 0.168
    sepia_b = arr[:,:,0] * 0.272 + arr[:,:,1] * 0.534 + arr[:,:,2] * 0.131

    sepia = np.stack([sepia_r, sepia_g, sepia_b], axis=2)

    # Blend with original
    result = arr * (1 - intensity) + sepia * intensity
    return Image.fromarray(np.clip(result, 0, 255).astype(np.uint8))


def effect_color_overlay(img: Image.Image, color: Tuple = (0, 0, 0),
                          opacity: float = 0.3, **kw) -> Image.Image:
    """Solid color overlay."""
    overlay = Image.new("RGB", img.size, color)
    return Image.blend(img, overlay, alpha=opacity)


def effect_mirror(img: Image.Image, axis: str = "horizontal", **kw) -> Image.Image:
    """Mirror/flip effect."""
    if axis == "horizontal":
        return img.transpose(Image.FLIP_LEFT_RIGHT)
    return img.transpose(Image.FLIP_TOP_BOTTOM)


def effect_rotate(img: Image.Image, angle: float = 90, **kw) -> Image.Image:
    """Rotation with expansion."""
    return img.rotate(angle, expand=True, resample=Image.BICUBIC)


def effect_noise(img: Image.Image, amount: float = 0.05, **kw) -> Image.Image:
    """Uniform noise overlay."""
    arr = np.array(img, dtype=np.float64)
    noise = np.random.uniform(-amount * 255, amount * 255, arr.shape)
    arr += noise
    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8))


def effect_polaroid(img: Image.Image, **kw) -> Image.Image:
    """Polaroid instant photo look."""
    img = ImageEnhance.Contrast(img).enhance(1.1)
    img = ImageEnhance.Color(img).enhance(0.9)
    img = effect_sepia(img, intensity=0.15)

    # Add white border (thicker at bottom)
    w, h = img.size
    border = min(w, h) // 15
    bottom_extra = border * 2

    result = Image.new("RGB", (w + border * 2, h + border + bottom_extra), (255, 255, 255))
    result.paste(img, (border, border))
    return result


def effect_invert(img: Image.Image, **kw) -> Image.Image:
    """Invert colors."""
    return ImageChops.invert(img)


# ─── Registry ──────────────────────────────────────────────────────────────

EFFECTS_REGISTRY = {
    "blur": effect_blur,
    "box_blur": effect_box_blur,
    "glow": effect_glow,
    "vignette": effect_vignette,
    "grain": effect_grain,
    "chromatic_aberration": effect_chromatic_aberration,
    "duotone": effect_duotone,
    "halftone": effect_halftone,
    "pixelate": effect_pixelate,
    "sharpen": effect_sharpen,
    "unsharp_mask": effect_unsharp_mask,
    "emboss": effect_emboss,
    "edge_enhance": effect_edge_enhance,
    "contrast": effect_contrast,
    "brightness": effect_brightness,
    "saturation": effect_saturation,
    "sepia": effect_sepia,
    "color_overlay": effect_color_overlay,
    "mirror": effect_mirror,
    "rotate": effect_rotate,
    "noise": effect_noise,
    "polaroid": effect_polaroid,
    "invert": effect_invert,
}
