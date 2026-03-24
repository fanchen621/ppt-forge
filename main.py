#!/usr/bin/env python3
"""
PPT Forge v2 — Professional-grade presentation generator.
Zero paid APIs. Pure math. Designer-level output.

Usage:
    python3 main.py generate --config slides.json --output presentation.pptx
    python3 main.py quick --title "My Talk" --theme dark --output out.pptx
    python3 main.py video --config slides.json --output video.mp4
    python3 main.py preview-art --style flow_field --palette aurora --output art.png
    python3 main.py preview-svg --style circuit --palette neon --output circuit.png
    python3 main.py preview-illustration --category plant --palette forest --output ill.png
    python3 main.py list-themes
    python3 main.py list-art-styles
    python3 main.py list-palettes
    python3 main.py list-slide-types
    python3 main.py apply-effect --input photo.png --effects blur,vignette,grain --output result.png
    python3 main.py demo --output demo.pptx
    python3 main.py demo-video --output demo.mp4
"""

import sys
import os
import json
import argparse
import random

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from engines.pptx_builder import build_ppt, THEMES
from engines.art_engine import generate_art, GENERATORS, PALETTES
from engines.svg_engine import generate_svg_art, svg_to_png, SVG_GENERATORS
from engines.effects_pipeline import apply_effects, EFFECTS_REGISTRY
from engines.illustration_engine import generate_illustration, ILLUSTRATION_GENERATORS
from engines.svg_engine import svg_to_png as ill_to_png
from engines.video_engine import build_video
from PIL import Image


def cmd_generate(args):
    """Generate PPT from config JSON."""
    with open(args.config, "r") as f:
        config = json.load(f)
    output = args.output or "presentation.pptx"
    build_ppt(config, output)
    print(f"✅ Generated: {output}")
    print(f"   Theme: {config.get('theme', 'dark')}")
    print(f"   Slides: {len(config.get('slides', []))}")


def cmd_quick(args):
    """Quick-generate a professional demo presentation."""
    title = args.title or "Presentation"
    theme = args.theme or "dark"
    seed = args.seed or random.randint(1, 999999)

    config = {
        "title": title,
        "theme": theme,
        "seed": seed,
        "slides": [
            {
                "type": "title",
                "heading": title,
                "subheading": "Generated with PPT Forge v2 — zero paid APIs",
                "bg_style": THEMES[theme]["default_art"],
                "palette": THEMES[theme]["palette"],
                "seed": seed,
            },
            {
                "type": "agenda",
                "heading": "Agenda",
                "items": [
                    "The Problem & Opportunity",
                    "Our Approach & Methodology",
                    "Key Results & Impact",
                    "Technical Architecture",
                    "Growth Metrics & Projections",
                    "Next Steps & Call to Action",
                ],
                "seed": seed + 1,
            },
            {
                "type": "content",
                "heading": "Key Points",
                "body": [
                    "• First major point — backed by compelling data and research",
                    "• Second point — demonstrated through real-world examples",
                    "• Third insight — the competitive advantage we bring",
                    "• Call to action — what happens next",
                ],
                "bg_style": "simplex_noise",
                "palette": "aurora",
                "seed": seed + 2,
            },
            {
                "type": "two_column",
                "heading": "Before vs After",
                "left": {
                    "title": "❌ Before",
                    "items": ["Manual processes", "Slow delivery cycles", "High operational cost", "Limited scalability"],
                },
                "right": {
                    "title": "✅ After",
                    "items": ["Fully automated pipeline", "Real-time delivery", "80% cost reduction", "Infinite scalability"],
                },
                "seed": seed + 3,
            },
            {
                "type": "metrics",
                "heading": "Impact at a Glance",
                "metrics": [
                    {"value": "3.2M", "label": "Users Reached", "sublabel": "+40% QoQ"},
                    {"value": "99.9%", "label": "Uptime", "sublabel": "Last 12 months"},
                    {"value": "$2.4M", "label": "Cost Saved", "sublabel": "Annual savings"},
                ],
                "palette": "emerald",
                "seed": seed + 4,
            },
            {
                "type": "image_full",
                "heading": "Visual Impact",
                "art_style": "flow_field",
                "palette": "aurora",
                "caption": "Generated with Perlin noise flow field algorithm — pure math, zero APIs",
                "seed": seed + 5,
            },
            {
                "type": "chart",
                "heading": "Growth Trajectory",
                "chart_type": "bar",
                "data": {
                    "labels": ["Q1", "Q2", "Q3", "Q4"],
                    "values": [25, 48, 72, 95],
                },
                "palette": "emerald",
                "seed": seed + 6,
            },
            {
                "type": "quote",
                "text": "This technology fundamentally changed how we approach presentations. The quality rivals what we used to pay thousands for.",
                "author": "CTO, Fortune 500 Company",
                "bg_style": "glass_morphism",
                "seed": seed + 7,
            },
            {
                "type": "timeline",
                "heading": "Roadmap",
                "items": [
                    {"date": "Q1 2026", "title": "Foundation & Core"},
                    {"date": "Q2 2026", "title": "Scale & Optimize"},
                    {"date": "Q3 2026", "title": "New Markets"},
                    {"date": "Q4 2026", "title": "Global Launch"},
                ],
                "palette": "golden",
                "seed": seed + 8,
            },
            {
                "type": "image_grid",
                "heading": "Art Style Gallery",
                "art_styles": ["flow_field", "voronoi", "bokeh", "geometric", "constellation", "glass_morphism"],
                "palette": "sunset",
                "seed": seed + 9,
            },
            {
                "type": "data_table",
                "heading": "Performance Summary",
                "headers": ["Metric", "Current", "Target", "Gap"],
                "rows": [
                    ["Revenue", "$5.2M", "$8.0M", "+54%"],
                    ["Users", "120K", "250K", "+108%"],
                    ["NPS Score", "72", "85", "+18%"],
                    ["Retention", "85%", "92%", "+8%"],
                ],
                "palette": "ocean",
                "seed": seed + 10,
            },
            {
                "type": "end",
                "heading": "Thank You",
                "subheading": "Powered by PPT Forge v2 — zero paid APIs, pure creative computation",
                "bg_style": "bokeh",
                "palette": "sunset",
                "seed": seed + 11,
            },
        ],
    }

    output = args.output or "presentation.pptx"
    build_ppt(config, output)
    print(f"✅ Generated: {output}")
    print(f"   Theme: {theme}")
    print(f"   Slides: {len(config['slides'])}")
    print(f"   Seed: {seed}")


def cmd_demo(args):
    """Generate a rich demo showcasing all features."""
    seed = args.seed or random.randint(1, 999999)

    themes = list(THEMES.keys())
    all_palettes = list(PALETTES.keys())

    config = {
        "title": "PPT Forge v2 — Feature Showcase",
        "theme": "dark",
        "slides": [],
    }

    # Title
    config["slides"].append({
        "type": "title",
        "heading": "PPT Forge v2",
        "subheading": "Professional Presentations — Zero Paid APIs, Pure Math",
        "bg_style": "flow_field",
        "palette": "aurora",
        "seed": seed,
    })

    # One slide per art style
    art_styles = list(GENERATORS.keys())
    for i, style in enumerate(art_styles):
        palette = all_palettes[i % len(all_palettes)]
        config["slides"].append({
            "type": "image_full",
            "heading": style.replace("_", " ").title(),
            "art_style": style,
            "palette": palette,
            "caption": f"Art style: {style} • Palette: {palette}",
            "seed": seed + i + 1,
        })

    # End
    config["slides"].append({
        "type": "end",
        "heading": "15 Art Styles • 19 Palettes • 13 Slide Types",
        "subheading": "Zero paid APIs. Zero external services. Pure creative computation.",
        "bg_style": "bokeh",
        "palette": "sunset",
        "seed": seed + 100,
    })

    output = args.output or "demo.pptx"
    build_ppt(config, output)
    print(f"✅ Demo generated: {output}")
    print(f"   Slides: {len(config['slides'])}")
    print(f"   Art styles showcased: {len(art_styles)}")


def cmd_preview_art(args):
    """Preview a single art style."""
    style = args.style or "flow_field"
    palette = args.palette or "aurora"
    width = args.width or 1920
    height = args.height or 1080
    seed = args.seed or random.randint(1, 999999)

    img = generate_art(style, width, height, palette_name=palette, seed=seed)
    output = args.output or f"art_{style}_{palette}.png"
    img.save(output)
    print(f"✅ Art preview: {output}")
    print(f"   Style: {style}")
    print(f"   Palette: {palette}")
    print(f"   Size: {width}x{height}")
    print(f"   Seed: {seed}")


def cmd_preview_svg(args):
    """Preview a single SVG art style."""
    style = args.style or "geometric_pattern"
    palette_name = args.palette or "neon"
    width = args.width or 1920
    height = args.height or 1080
    seed = args.seed or random.randint(1, 999999)

    palette = PALETTES.get(palette_name, PALETTES["sunset"])
    svg_str = generate_svg_art(style, width, height, palette, seed=seed)

    # Save SVG
    svg_out = args.output_svg or f"svg_{style}.svg"
    with open(svg_out, "w") as f:
        f.write(svg_str)

    # Convert to PNG
    output = args.output or f"svg_{style}_{palette_name}.png"
    svg_to_png(svg_str, width, height, output)
    print(f"✅ SVG preview: {output}")
    print(f"   SVG source: {svg_out}")
    print(f"   Style: {style}")
    print(f"   Palette: {palette_name}")


def cmd_apply_effect(args):
    """Apply effects chain to an image."""
    input_path = args.input
    if not os.path.exists(input_path):
        print(f"❌ File not found: {input_path}")
        return

    img = Image.open(input_path).convert("RGB")
    effects = [e.strip() for e in args.effects.split(",")]

    # Parse params like "blur:radius=10"
    effect_list = []
    for e in effects:
        parts = e.split(":")
        effect_type = parts[0]
        params = {}
        if len(parts) > 1:
            for p in parts[1].split(","):
                k, v = p.split("=")
                try:
                    params[k] = float(v)
                except ValueError:
                    params[k] = v
        effect_list.append({"type": effect_type, **params})

    result = apply_effects(img, effect_list)
    output = args.output or f"effected_{os.path.basename(input_path)}"
    result.save(output)
    print(f"✅ Effect applied: {output}")
    print(f"   Effects: {', '.join(effects)}")


def cmd_list_themes(args):
    """List available themes."""
    print("\n🎨 Available Themes:\n")
    for name, cfg in THEMES.items():
        print(f"  {name:12s}  bg={cfg['bg_color']}  accent={cfg['accent_color']}  art={cfg['default_art']}  palette={cfg['palette']}")
    print()


def cmd_list_art_styles(args):
    """List available art generation styles."""
    print("\n🖼️  Pixel Art Styles:\n")
    for name in sorted(GENERATORS.keys()):
        print(f"  • {name}")
    print(f"\n📐 SVG Art Styles:\n")
    for name in sorted(SVG_GENERATORS.keys()):
        print(f"  • {name}")
    print(f"\n✨ Visual Effects:\n")
    for name in sorted(EFFECTS_REGISTRY.keys()):
        print(f"  • {name}")
    print()


def cmd_list_palettes(args):
    """List available color palettes."""
    print("\n🎨 Color Palettes:\n")
    for name in sorted(PALETTES.keys()):
        colors = PALETTES[name]
        print(f"  • {name:12s}  ({len(colors)} colors: {', '.join(f'#{r:02x}{g:02x}{b:02x}' for r,g,b in colors[:3])}…)")
    print()


def cmd_list_slide_types(args):
    """List available slide types."""
    from engines.pptx_builder import SLIDE_BUILDERS
    print("\n📊 Slide Types:\n")
    descriptions = {
        "title": "Full-bleed art cover slide with overlay title",
        "content": "Standard content with heading and bullet points",
        "two_column": "Side-by-side comparison layout",
        "three_column": "Three-column feature overview",
        "image_full": "Full-bleed art image with overlay text",
        "chart": "Data chart (bar, line, pie, area, donut)",
        "end": "Thank you / closing slide with art background",
        "agenda": "Numbered agenda/outline with colored markers",
        "timeline": "Horizontal timeline with labeled milestones",
        "quote": "Large quote with attribution",
        "data_table": "Styled data table with alternating rows",
        "metrics": "Big number KPI cards (3-4 metrics)",
        "image_grid": "Gallery grid of art styles",
        "cn_title": "🇨🇳 Chinese education title — large characters + decorative art",
        "cn_task_separator": "🇨🇳 Full-page task/section separator with decorative strips",
        "cn_content": "🇨🇳 Rich content with vertical side label + framed area",
        "cn_scenario": "🇨🇳 Scenario slide with thought bubble + analysis layout",
        "cn_mind_map": "🇨🇳 Mind map with root node + children + sub-items",
        "cn_evaluation": "🇨🇳 Styled evaluation/rubric table with alternating rows",
        "cn_flowchart": "🇨🇳 Process flow with numbered step boxes + arrows",
        "cn_reading": "🇨🇳 Reading passage with analysis sidebar",
        "cn_brackets": "🇨🇳 Bracket analysis — curly brackets with labeled groups",
    }
    for name in sorted(SLIDE_BUILDERS.keys()):
        desc = descriptions.get(name, "")
        print(f"  • {name:25s}  {desc}")
    print()


def cmd_video(args):
    """Generate MP4 video from config JSON."""
    with open(args.config, "r") as f:
        config = json.load(f)
    output = args.output or "presentation.mp4"
    build_video(config, output, fps=args.fps, slide_duration=args.slide_duration,
                transition=args.transition)
    print(f"✅ Video generated: {output}")
    print(f"   Slides: {len(config.get('slides', []))}")
    print(f"   FPS: {args.fps}")
    print(f"   Transition: {args.transition}")


def cmd_demo_video(args):
    """Generate demo video from quick config."""
    title = args.title or "PPT Forge v2 Video Demo"
    theme = args.theme or "dark"
    seed = args.seed or random.randint(1, 999999)

    config = {
        "title": title,
        "theme": theme,
        "slides": [
            {
                "type": "title",
                "heading": title,
                "subheading": "Generated with PPT Forge v2 — zero paid APIs",
                "bg_style": THEMES[theme]["default_art"],
                "palette": THEMES[theme]["palette"],
                "seed": seed,
            },
            {
                "type": "content",
                "heading": "What is PPT Forge?",
                "body": [
                    "• Professional-grade presentation generator",
                    "• 15 pixel art styles + 7 SVG vector styles",
                    "• 23 visual effects pipeline",
                    "• 19 color palettes + 8 themes",
                    "• Zero paid APIs — pure creative computation",
                ],
                "bg_style": "simplex_noise",
                "palette": "aurora",
                "seed": seed + 1,
            },
            {
                "type": "content",
                "heading": "How It Works",
                "body": [
                    "• Perlin & Simplex noise for organic textures",
                    "• Voronoi tessellation for cellular patterns",
                    "• Flow field particle systems for dynamic art",
                    "• SVG vector generation with rsvg-convert",
                    "• Smart layout engine with golden ratio grid",
                ],
                "bg_style": "constellation",
                "palette": "midnight",
                "seed": seed + 2,
            },
            {
                "type": "end",
                "heading": "Thank You",
                "subheading": "PPT Forge v2 — pure math, zero APIs",
                "bg_style": "bokeh",
                "palette": "sunset",
                "seed": seed + 3,
            },
        ],
    }

    output = args.output or "demo.mp4"
    build_video(config, output)
    print(f"✅ Demo video: {output}")
    print(f"   Theme: {theme}")
    print(f"   Seed: {seed}")


def cmd_preview_illustration(args):
    """Preview a contextual illustration."""
    category = args.category or "education"
    palette = args.palette or "forest"
    width = args.width or 400
    height = args.height or 400
    seed = args.seed or random.randint(1, 999999)

    svg_str = generate_illustration(category, width, height, palette_name=palette, seed=seed)
    output = args.output or f"illustration_{category}_{palette}.png"
    ill_to_png(svg_str, width, height, output)
    print(f"✅ Illustration preview: {output}")
    print(f"   Category: {category}")
    print(f"   Palette: {palette}")


def cmd_list_illustrations(args):
    """List available illustration categories."""
    print("\n🎨 Illustration Categories (auto-detected from content):\n")
    for name in sorted(ILLUSTRATION_GENERATORS.keys()):
        print(f"  • {name}")
    print()
    print("Usage: The engine auto-detects category from slide text content.")
    print("       Or specify manually with --category.")
    print()


def main():
    parser = argparse.ArgumentParser(
        prog="ppt-forge",
        description="PPT Forge v2 — Professional presentation generator, zero paid APIs",
    )
    sub = parser.add_subparsers(dest="command")

    # generate
    p = sub.add_parser("generate", help="Generate PPT from config JSON")
    p.add_argument("--config", "-c", required=True, help="Path to slides.json")
    p.add_argument("--output", "-o", help="Output PPTX path")

    # quick
    p = sub.add_parser("quick", help="Quick professional demo")
    p.add_argument("--title", "-t", help="Title")
    p.add_argument("--theme", help="Theme name")
    p.add_argument("--output", "-o", help="Output path")
    p.add_argument("--seed", type=int, help="Random seed")

    # demo
    p = sub.add_parser("demo", help="Full feature showcase")
    p.add_argument("--output", "-o", help="Output path")
    p.add_argument("--seed", type=int, help="Random seed")

    # preview-art
    p = sub.add_parser("preview-art", help="Preview pixel art style")
    p.add_argument("--style", "-s", help="Art style")
    p.add_argument("--palette", "-p", help="Palette")
    p.add_argument("--width", type=int, default=1920)
    p.add_argument("--height", type=int, default=1080)
    p.add_argument("--output", "-o", help="Output path")
    p.add_argument("--seed", type=int, help="Random seed")

    # preview-svg
    p = sub.add_parser("preview-svg", help="Preview SVG art style")
    p.add_argument("--style", "-s", help="SVG style")
    p.add_argument("--palette", "-p", help="Palette")
    p.add_argument("--width", type=int, default=1920)
    p.add_argument("--height", type=int, default=1080)
    p.add_argument("--output", "-o", help="Output PNG path")
    p.add_argument("--output-svg", help="Output SVG path")
    p.add_argument("--seed", type=int, help="Random seed")

    # apply-effect
    p = sub.add_parser("apply-effect", help="Apply effects to image")
    p.add_argument("--input", "-i", required=True, help="Input image")
    p.add_argument("--effects", "-e", required=True, help="Effects chain (comma-separated)")
    p.add_argument("--output", "-o", help="Output path")

    # video
    p = sub.add_parser("video", help="Generate MP4 video from config")
    p.add_argument("--config", "-c", required=True, help="Path to slides.json")
    p.add_argument("--output", "-o", help="Output MP4 path")
    p.add_argument("--fps", type=int, default=30, help="Frames per second")
    p.add_argument("--slide-duration", type=float, default=5.0, help="Seconds per slide")
    p.add_argument("--transition", default="fade", choices=["fade", "slide", "zoom"], help="Transition type")

    # demo-video
    p = sub.add_parser("demo-video", help="Generate demo video from quick config")
    p.add_argument("--title", "-t", help="Title")
    p.add_argument("--theme", help="Theme name")
    p.add_argument("--output", "-o", help="Output MP4 path")
    p.add_argument("--seed", type=int, help="Random seed")

    # preview-illustration
    p = sub.add_parser("preview-illustration", help="Preview contextual illustration")
    p.add_argument("--category", "-c", help="Category (plant/science/education/time/data/team/safety/resource/award/nature)")
    p.add_argument("--text", "-t", help="Text for auto-detection")
    p.add_argument("--palette", "-p", help="Palette")
    p.add_argument("--width", type=int, default=400)
    p.add_argument("--height", type=int, default=400)
    p.add_argument("--output", "-o", help="Output PNG path")
    p.add_argument("--seed", type=int, help="Random seed")

    # Lists
    sub.add_parser("list-themes", help="List themes")
    sub.add_parser("list-art-styles", help="List art styles & effects")
    sub.add_parser("list-palettes", help="List color palettes")
    sub.add_parser("list-slide-types", help="List slide types")
    sub.add_parser("list-illustrations", help="List illustration categories")

    args = parser.parse_args()

    commands = {
        "generate": cmd_generate,
        "quick": cmd_quick,
        "demo": cmd_demo,
        "preview-art": cmd_preview_art,
        "preview-svg": cmd_preview_svg,
        "apply-effect": cmd_apply_effect,
        "video": cmd_video,
        "demo-video": cmd_demo_video,
        "preview-illustration": cmd_preview_illustration,
        "list-themes": cmd_list_themes,
        "list-art-styles": cmd_list_art_styles,
        "list-palettes": cmd_list_palettes,
        "list-slide-types": cmd_list_slide_types,
        "list-illustrations": cmd_list_illustrations,
    }

    if args.command in commands:
        commands[args.command](args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
