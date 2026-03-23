"""PPT Forge v2 — Professional-grade presentation generator, zero paid APIs."""
from .art_engine import generate_art, GENERATORS, PALETTES
from .svg_engine import generate_svg_art, svg_to_png
from .effects_pipeline import apply_effects
from .layout_engine import SmartLayout
from .illustration_engine import generate_illustration, auto_illustrate_slide, detect_categories
from .video_engine import build_video
from .pptx_builder import build_ppt, THEMES
