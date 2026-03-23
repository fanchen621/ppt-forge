---
name: ppt-forge
description: Generate professional-grade PowerPoint presentations with AI-quality generative art backgrounds and auto-illustrations — zero paid APIs, zero external services. Produces PPTX files rivaling Midjourney/DALL-E 3 designer output using pure mathematical algorithms. Features: 15 pixel art styles, 7 SVG vector art, 23 visual effects, 19 palettes, 8 themes, 13 slide types, intelligent contextual illustrations, and MP4 video export. Triggered by: PPT, PowerPoint, presentation, slides, deck, pitch deck, video.
---

# PPT Forge v2 — Professional Presentation Generator

Generate designer-level PPTX files with generative art backgrounds and auto-illustrations. **100% free, zero paid APIs.**

## Quick Start

```bash
cd SKILL_DIR

# Professional demo
python3 main.py quick --title "My Presentation" --theme dark --output demo.pptx

# Full feature showcase
python3 main.py demo --output showcase.pptx

# Generate from config
python3 main.py generate --config slides.json --output presentation.pptx

# Generate MP4 video from config
python3 main.py video --config slides.json --output video.mp4

# Demo video
python3 main.py demo-video --title "Demo" --theme dark --output demo.mp4

# Preview art
python3 main.py preview-art --style flow_field --palette aurora --output art.png
python3 main.py preview-svg --style circuit --palette neon --output circuit.png
python3 main.py preview-illustration --category plant --palette forest --output plant.png

# List capabilities
python3 main.py list-themes
python3 main.py list-art-styles
python3 main.py list-palettes
python3 main.py list-slide-types
python3 main.py list-illustrations
```

## Config Format (slides.json)

```json
{
  "title": "Presentation Title",
  "theme": "dark",
  "slides": [
    {"type": "title", "heading": "Main Title", "subheading": "Subtitle", "bg_style": "flow_field", "palette": "aurora"},
    {"type": "agenda", "heading": "Agenda", "items": ["Point 1", "Point 2", "Point 3"]},
    {"type": "content", "heading": "Section", "body": ["• Bullet 1", "• Bullet 2"], "bg_style": "simplex_noise", "palette": "ocean"},
    {"type": "two_column", "heading": "Compare", "left": {"title": "A", "items": ["a1","a2"]}, "right": {"title": "B", "items": ["b1","b2"]}},
    {"type": "three_column", "heading": "Features", "columns": [{"title":"A","items":["x"]}, {"title":"B","items":["y"]}, {"title":"C","items":["z"]}]},
    {"type": "metrics", "heading": "KPIs", "metrics": [{"value":"3.2M","label":"Users","sublabel":"+40%"}, {"value":"99.9%","label":"Uptime"}]},
    {"type": "image_full", "heading": "Visual", "art_style": "flow_field", "palette": "aurora", "caption": "Generated art"},
    {"type": "chart", "heading": "Growth", "chart_type": "bar", "data": {"labels":["Q1","Q2"], "values":[30,50]}, "palette": "emerald"},
    {"type": "quote", "text": "Great quote here", "author": "Name, Title", "bg_style": "glass_morphism"},
    {"type": "timeline", "heading": "Roadmap", "items": [{"date":"Q1","title":"Launch"}, {"date":"Q2","title":"Scale"}]},
    {"type": "data_table", "heading": "Data", "headers":["Col1","Col2"], "rows":[["a","b"],["c","d"]]},
    {"type": "image_grid", "heading": "Gallery", "art_styles":["flow_field","voronoi","bokeh"]},
    {"type": "end", "heading": "Thank You", "subheading": "contact@example.com", "bg_style": "bokeh", "palette": "sunset"}
  ]
}
```

## Slide Types

| Type | Description | Key Fields |
|------|-------------|------------|
| `title` | Full-bleed art cover | heading, subheading, bg_style, palette |
| `content` | Heading + bullets + auto-illustration | heading, body, bg_style |
| `two_column` | Side-by-side + illustration | heading, left, right |
| `three_column` | Feature overview + illustration | heading, columns |
| `metrics` | Big KPI numbers (3-4) + illustration | heading, metrics[{value, label, sublabel}] |
| `image_full` | Art background + overlay | heading, art_style, palette, caption |
| `chart` | Data chart | heading, chart_type, data{labels, values} |
| `quote` | Quote/testimonial + illustration | text, author, bg_style |
| `agenda` | Numbered outline + illustration | heading, items |
| `timeline` | Horizontal timeline + illustration | heading, items[{date, title}] |
| `data_table` | Styled table + illustration | heading, headers, rows |
| `image_grid` | Art gallery grid | heading, art_styles |
| `end` | Closing slide | heading, subheading, bg_style, palette |

## Themes

`minimal` · `dark` · `gradient` · `corporate` · `creative` · `tech` · `nature` · `neon`

## Art Styles (Pixel — Pillow)

`gradient_mesh` · `perlin_noise` · `simplex_noise` · `voronoi` · `flow_field` · `geometric` · `fractal_tree` · `bokeh` · `wave` · `constellation` · `glass_morphism` · `marble` · `isometric_grid` · `topographic` · `double_exposure`

## SVG Art Styles (Vector — rsvg-convert)

`geometric_pattern` · `hex_grid` · `circuit` · `wave_pattern` · `abstract_art` · `diagonal_stripes` · `radial_sunburst`

## Visual Effects (23)

`blur` · `box_blur` · `glow` · `vignette` · `grain` · `chromatic_aberration` · `duotone` · `halftone` · `pixelate` · `sharpen` · `emboss` · `contrast` · `brightness` · `saturation` · `sepia` · `color_overlay` · `mirror` · `noise` · `polaroid` · `invert`

## Illustration Categories (auto-detected from content)

`plant` · `science` · `education` · `time` · `data` · `team` · `safety` · `resource` · `award` · `nature`

## Color Palettes (19)

`sunset` · `ocean` · `forest` · `neon` · `pastel` · `mono` · `warm` · `cool` · `fire` · `aurora` · `midnight` · `sakura` · `lavender` · `emerald` · `crimson` · `arctic` · `golden` · `cosmic` · `mint` · `coral`

## Video Engine

```bash
# Generate MP4 from config
python3 main.py video --config slides.json --output video.mp4 --fps 30 --transition fade

# Transition types: fade, slide, zoom
# --slide-duration: seconds per slide (default: 5)
# --fps: frames per second (default: 30)
```

## Agent Workflow

When asked to create a PPT/PPTX/presentation:

1. **Parse requirements** → determine slide types, theme, content
2. **Build config JSON** → map user's content to slide types
3. **Run generator**:
   ```bash
   python3 SKILL_DIR/main.py generate --config /tmp/slides.json --output /tmp/presentation.pptx
   ```
4. **Deliver the PPTX file**

### Decision Guide

- **Pitch deck** → title, agenda, content×2, metrics, chart, image_full, quote, timeline, end
- **Data report** → title, chart×3, data_table, metrics, end
- **Creative showcase** → title, image_full×3, image_grid, end
- **Corporate meeting** → title, agenda, content×4, two_column, data_table, end
- **Quick overview** → title, content×2, end (use `quick` command)
- **Video presentation** → use `video` command with config JSON

## Dependencies (all free)

- `python-pptx` — PPTX creation
- `Pillow` — Image processing & generative art
- `numpy` — Fast numerical computation
- `rsvg-convert` — SVG to PNG (system binary)
- `ffmpeg` — Video encoding (system binary, for video mode)
