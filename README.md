# 🔥 PPT Forge v2

**Professional-grade presentation generator — zero paid APIs, pure math.**

Generates stunning PPTX files with generative art backgrounds and **auto-contextual illustrations** rivaling Midjourney/DALL-E 3 designer output. Everything runs locally using open-source algorithms.

## ✨ What's New in v2

| Feature | v1 | v2 |
|---------|-----|-----|
| **Pixel Art Styles** | 10 | **15** (+simplex_noise, marble, isometric_grid, topographic, double_exposure) |
| **SVG Art Engine** | ❌ | **7 vector styles** |
| **Visual Effects** | basic | **23 professional effects** |
| **Color Palettes** | 12 | **19** |
| **Slide Types** | 6 | **13** |
| **Smart Illustrations** | ❌ | **10 auto-detected categories** (plant, science, education, data, team...) |
| **Video Export** | ❌ | **MP4 with Ken Burns + transitions** (fade, slide, zoom) |
| **Layout Engine** | manual | **Golden ratio grid + modular type scale** |

## 🎨 Art Styles

### Pixel Art (Pillow)
`gradient_mesh` · `perlin_noise` · `simplex_noise` · `voronoi` · `flow_field` · `geometric` · `fractal_tree` · `bokeh` · `wave` · `constellation` · `glass_morphism` · `marble` · `isometric_grid` · `topographic` · `double_exposure`

### SVG Vector Art (rsvg-convert)
`geometric_pattern` · `hex_grid` · `circuit` · `wave_pattern` · `abstract_art` · `diagonal_stripes` · `radial_sunburst`

### Visual Effects Pipeline
`blur` · `glow` · `vignette` · `grain` · `chromatic_aberration` · `duotone` · `halftone` · `sepia` · `polaroid` · `invert` (and 13 more)

### 🤖 Intelligent Illustrations
Auto-detects content from slide text and generates matching SVG illustrations:
- **plant** — botanical stem, leaves, flower, roots
- **science** — magnifier, test tube, data charts
- **education** — open book, lightbulb, pencil
- **data** — pie chart, line chart, bar chart
- **team** — people icons with connections
- **safety** — shield, handwashing, heart
- **resource** — toolbox, seeds, ruler
- **award** — trophy, star, confetti
- **nature** — sun, clouds, mountains, trees
- **time** — winding path with milestones

### 🎬 Video Engine
Converts presentations into MP4 videos:
- Ken Burns slow-zoom effect on backgrounds
- Cross-fade / slide-left / zoom transitions
- Configurable FPS and duration
- Pure ffmpeg encoding

## 🚀 Quick Start

```bash
# Professional demo
python3 main.py quick --title "My Presentation" --theme dark --output demo.pptx

# Full feature showcase
python3 main.py demo --output showcase.pptx

# Generate from config
python3 main.py generate --config slides.json --output presentation.pptx

# Generate video
python3 main.py video --config slides.json --output video.mp4

# Preview illustrations
python3 main.py preview-illustration --category plant --palette forest --output plant.png
```

## 📋 Config Format

```json
{
  "title": "My Presentation",
  "theme": "dark",
  "slides": [
    {"type": "title", "heading": "Title", "subheading": "Sub", "bg_style": "flow_field", "palette": "aurora"},
    {"type": "content", "heading": "Key Points", "body": ["• Item 1", "• Item 2"]},
    {"type": "end", "heading": "Thank You", "subheading": "contact@me.com"}
  ]
}
```

## 🎨 19 Color Palettes

`sunset` · `ocean` · `forest` · `neon` · `pastel` · `mono` · `warm` · `cool` · `fire` · `aurora` · `midnight` · `sakura` · `lavender` · `emerald` · `crimson` · `arctic` · `golden` · `cosmic` · `mint` · `coral`

## 🔧 Dependencies

- `python-pptx` — PPTX creation
- `Pillow` — Image processing & generative art
- `numpy` — Fast numerical computation
- `rsvg-convert` — SVG to PNG (system binary)
- `ffmpeg` — Video encoding (system binary, for video mode)

All free. No paid APIs. No external services. No data leaves your machine.

## 📄 License

MIT
