"""
PPT Forge v2 — Intelligent Video Engine
Converts PPTX presentations into narrated MP4 videos using ffmpeg.
Supports slide transitions, text animations, and background music.
Zero paid APIs — pure ffmpeg + Pillow rendering.
"""

import os
import sys
import math
import random
import subprocess
import tempfile
import shutil
from typing import List, Dict, Optional, Tuple
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from art_engine import generate_art, get_palette, PALETTES
from effects_pipeline import apply_effects


# ─── Configuration ─────────────────────────────────────────────────────────

VIDEO_WIDTH = 1920
VIDEO_HEIGHT = 1080
FPS = 30
DEFAULT_SLIDE_DURATION = 5.0  # seconds per slide
TRANSITION_DURATION = 1.0     # seconds for transitions


# ─── Text Rendering ────────────────────────────────────────────────────────

def _try_load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    """Try to load a system font."""
    font_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc" if bold else
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc" if bold else
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()


def _render_text_overlay(width: int, height: int, heading: str, body_lines: list,
                          accent_color: Tuple, bg_color: Tuple, font_scale: float = 1.0) -> Image.Image:
    """Render a text overlay frame."""
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Semi-transparent background panel
    panel_h = int(height * 0.35)
    panel_y = height - panel_h
    for y in range(panel_y, height):
        alpha = int(180 * ((y - panel_y) / panel_h))
        draw.line([(0, y), (width, y)], fill=(0, 0, 0, min(180, alpha)))

    # Accent line
    draw.rectangle([60, panel_y + 20, 66, panel_y + 20 + 50], fill=accent_color + (200,))

    # Heading
    heading_size = int(42 * font_scale)
    font_heading = _try_load_font(heading_size, bold=True)
    draw.text((85, panel_y + 15), heading, fill=(255, 255, 255, 240), font=font_heading)

    # Body lines
    body_size = int(24 * font_scale)
    font_body = _try_load_font(body_size)
    for i, line in enumerate(body_lines[:6]):
        y_pos = panel_y + 75 + i * 32
        draw.text((85, y_pos), line, fill=(200, 200, 210, 200), font=font_body)

    return img


# ─── Transition Effects ───────────────────────────────────────────────────

def _transition_fade(img1: Image.Image, img2: Image.Image, progress: float) -> Image.Image:
    """Cross-fade transition between two frames."""
    return Image.blend(img1.convert("RGB"), img2.convert("RGB"), progress)


def _transition_slide_left(img1: Image.Image, img2: Image.Image, progress: float) -> Image.Image:
    """Slide left transition."""
    w, h = img1.size
    offset = int(w * progress)
    result = Image.new("RGB", (w, h), (0, 0, 0))
    # Crop and paste img1 moving left
    if offset < w:
        crop1 = img1.crop((offset, 0, w, h))
        result.paste(crop1, (0, 0))
    # Crop and paste img2 entering from right
    if offset > 0:
        crop2 = img2.crop((0, 0, offset, h))
        result.paste(crop2, (w - offset, 0))
    return result


def _transition_zoom(img1: Image.Image, img2: Image.Image, progress: float) -> Image.Image:
    """Zoom transition — img1 zooms out, img2 fades in."""
    w, h = img1.size
    # Zoom out img1
    scale = 1.0 + progress * 0.3
    new_w, new_h = int(w * scale), int(h * scale)
    zoomed = img1.resize((new_w, new_h), Image.LANCZOS)
    # Center crop
    left = (new_w - w) // 2
    top = (new_h - h) // 2
    zoomed = zoomed.crop((left, top, left + w, top + h))
    return Image.blend(zoomed, img2.convert("RGB"), progress)


TRANSITIONS = {
    "fade": _transition_fade,
    "slide": _transition_slide_left,
    "zoom": _transition_zoom,
}


# ─── Frame Generators ─────────────────────────────────────────────────────

def _generate_title_frame(slide_config: dict, theme_cfg: dict, art_img: Image.Image,
                           width: int, height: int) -> Image.Image:
    """Generate a video frame for title slide."""
    # Art background with dark overlay
    frame = art_img.resize((width, height), Image.LANCZOS).convert("RGB")
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 120))
    frame = Image.alpha_composite(frame.convert("RGBA"), overlay).convert("RGB")

    draw = ImageDraw.Draw(frame)

    # Heading
    heading = slide_config.get("heading", "")
    font_size = 56
    font = _try_load_font(font_size, bold=True)
    bbox = draw.textbbox((0, 0), heading, font=font)
    tw = bbox[2] - bbox[0]
    x = (width - tw) // 2
    y = int(height * 0.35)
    draw.text((x, y), heading, fill=(255, 255, 255), font=font)

    # Divider
    div_w = 200
    div_x = (width - div_w) // 2
    draw.rectangle([div_x, y + 80, div_x + div_w, y + 83], fill=(255, 255, 255))

    # Subheading
    sub = slide_config.get("subheading", "")
    if sub:
        font_sub = _try_load_font(28)
        bbox = draw.textbbox((0, 0), sub, font=font_sub)
        tw = bbox[2] - bbox[0]
        x = (width - tw) // 2
        draw.text((x, y + 100), sub, fill=(200, 200, 220), font=font_sub)

    return frame


def _generate_content_frame(slide_config: dict, theme_cfg: dict, art_img: Image.Image,
                             width: int, height: int, accent_color: Tuple) -> Image.Image:
    """Generate a video frame for content slide."""
    # Subtle art background
    bg = art_img.resize((width, height), Image.LANCZOS).convert("RGBA")
    alpha = bg.split()[3]
    alpha = alpha.point(lambda p: int(p * 0.12))
    bg.putalpha(alpha)
    base = Image.new("RGBA", (width, height), theme_cfg.get("bg_color", (14, 14, 20)) + (255,))
    frame = Image.alpha_composite(base, bg).convert("RGB")
    draw = ImageDraw.Draw(frame)

    # Left accent bar
    draw.rectangle([0, 0, 6, height], fill=accent_color)

    # Heading
    heading = slide_config.get("heading", "")
    font_h = _try_load_font(38, bold=True)
    draw.text((50, 30), heading, fill=theme_cfg.get("title_color", (245, 245, 250)), font=font_h)

    # Divider line
    draw.rectangle([50, 85, 200, 88], fill=accent_color)

    # Body text
    body = slide_config.get("body", [])
    if isinstance(body, str):
        body = [body]
    font_b = _try_load_font(22)
    for i, line in enumerate(body[:10]):
        y = 110 + i * 36
        draw.text((50, y), line, fill=theme_cfg.get("body_color", (175, 175, 195)), font=font_b)

    return frame


def _generate_end_frame(slide_config: dict, theme_cfg: dict, art_img: Image.Image,
                         width: int, height: int) -> Image.Image:
    """Generate closing frame."""
    frame = art_img.resize((width, height), Image.LANCZOS).convert("RGB")
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 120))
    frame = Image.alpha_composite(frame.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(frame)

    heading = slide_config.get("heading", "Thank You")
    font = _try_load_font(56, bold=True)
    bbox = draw.textbbox((0, 0), heading, font=font)
    tw = bbox[2] - bbox[0]
    draw.text(((width - tw) // 2, int(height * 0.35)), heading, fill=(255, 255, 255), font=font)

    sub = slide_config.get("subheading", "")
    if sub:
        font_s = _try_load_font(26)
        bbox = draw.textbbox((0, 0), sub, font=font_s)
        tw = bbox[2] - bbox[0]
        draw.text(((width - tw) // 2, int(height * 0.35) + 80), sub, fill=(200, 200, 220), font=font_s)

    return frame


def _generate_generic_frame(slide_config: dict, theme_cfg: dict, art_img: Image.Image,
                             width: int, height: int, accent_color: Tuple) -> Image.Image:
    """Generic frame for other slide types."""
    return _generate_content_frame(slide_config, theme_cfg, art_img, width, height, accent_color)


# ─── Main Video Builder ───────────────────────────────────────────────────

def build_video(config: Dict, output_path: str, fps: int = FPS,
                slide_duration: float = DEFAULT_SLIDE_DURATION,
                transition: str = "fade",
                transition_duration: float = TRANSITION_DURATION,
                title_duration: float = 6.0,
                end_duration: float = 5.0) -> str:
    """
    Generate an MP4 video from a slide config.
    Creates frames, then uses ffmpeg to encode.
    """
    theme = config.get("theme", "dark")
    slides = config.get("slides", [])
    theme_cfg = {
        "bg_color": (14, 14, 20),
        "title_color": (245, 245, 250),
        "body_color": (175, 175, 195),
        "accent_color": (99, 179, 237),
    }

    # Try to import theme data
    try:
        from pptx_builder import THEMES
        if theme in THEMES:
            theme_cfg.update(THEMES[theme])
    except ImportError:
        pass

    accent_color = theme_cfg.get("accent_color", (99, 179, 237))
    palette_name = theme_cfg.get("palette", "midnight")
    transition_fn = TRANSITIONS.get(transition, _transition_fade)

    # Create temp directory for frames
    tmp_dir = tempfile.mkdtemp(prefix="ppt_forge_video_")

    try:
        all_frames = []
        frame_idx = 0

        for slide_i, slide_cfg in enumerate(slides):
            slide_type = slide_cfg.get("type", "content")
            seed = slide_cfg.get("seed", random.randint(1, 999999))

            # Generate art background
            art_style = slide_cfg.get("bg_style", slide_cfg.get("art_style",
                       theme_cfg.get("default_art", "gradient_mesh")))
            art_palette = slide_cfg.get("palette", palette_name)
            art_img = generate_art(art_style, VIDEO_WIDTH, VIDEO_HEIGHT,
                                    palette_name=art_palette, seed=seed)

            # Duration based on slide type
            if slide_type == "title":
                dur = title_duration
            elif slide_type == "end":
                dur = end_duration
            else:
                dur = slide_duration

            n_frames = int(dur * fps)

            # Generate base frame
            if slide_type == "title":
                frame = _generate_title_frame(slide_cfg, theme_cfg, art_img,
                                               VIDEO_WIDTH, VIDEO_HEIGHT)
            elif slide_type == "end":
                frame = _generate_end_frame(slide_cfg, theme_cfg, art_img,
                                             VIDEO_WIDTH, VIDEO_HEIGHT)
            else:
                frame = _generate_generic_frame(slide_cfg, theme_cfg, art_img,
                                                 VIDEO_WIDTH, VIDEO_HEIGHT, accent_color)

            # Ken Burns effect: slow zoom on the art
            for fi in range(n_frames):
                progress = fi / max(1, n_frames - 1)
                # Subtle zoom (1.0 -> 1.05)
                scale = 1.0 + progress * 0.05
                sw, sh = int(VIDEO_WIDTH * scale), int(VIDEO_HEIGHT * scale)
                zoomed_art = art_img.resize((sw, sh), Image.LANCZOS)
                left = (sw - VIDEO_WIDTH) // 2
                top = (sh - VIDEO_HEIGHT) // 2
                zoomed_art = zoomed_art.crop((left, top, left + VIDEO_WIDTH, top + VIDEO_HEIGHT))

                # Blend frame with zoomed background
                blended = Image.blend(frame.convert("RGBA"),
                                       zoomed_art.convert("RGBA"), 0.15).convert("RGB")

                frame_path = os.path.join(tmp_dir, f"frame_{frame_idx:06d}.png")
                blended.save(frame_path)
                all_frames.append(frame_path)
                frame_idx += 1

            # Transition frames to next slide
            if slide_i < len(slides) - 1:
                next_cfg = slides[slide_i + 1]
                next_seed = next_cfg.get("seed", random.randint(1, 999999))
                next_art_style = next_cfg.get("bg_style", next_cfg.get("art_style",
                                theme_cfg.get("default_art", "gradient_mesh")))
                next_palette = next_cfg.get("palette", palette_name)
                next_art = generate_art(next_art_style, VIDEO_WIDTH, VIDEO_HEIGHT,
                                         palette_name=next_palette, seed=next_seed)

                if next_cfg.get("type") == "title":
                    next_frame = _generate_title_frame(next_cfg, theme_cfg, next_art,
                                                        VIDEO_WIDTH, VIDEO_HEIGHT)
                elif next_cfg.get("type") == "end":
                    next_frame = _generate_end_frame(next_cfg, theme_cfg, next_art,
                                                      VIDEO_WIDTH, VIDEO_HEIGHT)
                else:
                    next_frame = _generate_generic_frame(next_cfg, theme_cfg, next_art,
                                                          VIDEO_WIDTH, VIDEO_HEIGHT, accent_color)

                n_trans = int(transition_duration * fps)
                for ti in range(n_trans):
                    progress = ti / max(1, n_trans - 1)
                    trans_frame = transition_fn(frame, next_frame, progress)
                    frame_path = os.path.join(tmp_dir, f"frame_{frame_idx:06d}.png")
                    trans_frame.save(frame_path)
                    all_frames.append(frame_path)
                    frame_idx += 1

        # Encode with ffmpeg
        cmd = [
            "ffmpeg", "-y",
            "-framerate", str(fps),
            "-i", os.path.join(tmp_dir, "frame_%06d.png"),
            "-c:v", "libx264",
            "-preset", "medium",
            "-crf", "23",
            "-pix_fmt", "yuv420p",
            "-vf", f"scale={VIDEO_WIDTH}:{VIDEO_HEIGHT}",
            output_path
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if result.returncode != 0:
            raise RuntimeError(f"ffmpeg failed: {result.stderr[:500]}")

        return output_path

    finally:
        # Cleanup temp frames
        shutil.rmtree(tmp_dir, ignore_errors=True)


def video_from_pptx(pptx_path: str, config: Dict, output_path: str,
                     fps: int = FPS, slide_duration: float = DEFAULT_SLIDE_DURATION,
                     transition: str = "fade") -> str:
    """Generate video from existing PPTX + its config."""
    return build_video(config, output_path, fps=fps, slide_duration=slide_duration,
                       transition=transition)
