# Mck PPT Design Skill

A comprehensive McKinsey-style PowerPoint design system for creating professional, consultant-grade presentations using `python-pptx`.

## Features

- **Complete Design System** - Color palette, typography hierarchy, line treatments, layout patterns
- **Production-Ready Code** - Copy-paste Python helper functions for instant use
- **Shadow-Free Rendering** - Dual-defense approach to eliminate OOXML shadow/3D artifacts
- **East Asian Font Support** - Proper Chinese character rendering with KaiTi font
- **Multiple Layout Patterns** - Cover, Action Title, Table, Three-Column Overview, TOC
- **Post-Save Theme Cleanup** - Automated XML processing to ensure clean output

## Design Principles

| Principle | Description |
|-----------|-------------|
| **Extreme Minimalism** | No unnecessary visual elements, no gradients |
| **Flat Design** | No shadows, no 3D, no reflections |
| **Strict Hierarchy** | Title (22pt) > Sub-header (18pt) > Body (14pt) > Footnote (9pt) |
| **Consistency** | Unified colors, fonts, spacing across all slides |

## Color Palette

| Color | Hex | Usage |
|-------|-----|-------|
| NAVY | `#051C2C` | Primary titles, circles, highlights |
| BLACK | `#000000` | Lines, text separators |
| DARK_GRAY | `#333333` | Body text |
| MED_GRAY | `#666666` | Secondary text, labels |
| LINE_GRAY | `#CCCCCC` | Table row separators |
| BG_GRAY | `#F2F2F2` | Background panels |

## Quick Start

```bash
# Install dependencies
pip install python-pptx lxml

# Run minimal example
cd examples
python minimal_example.py
```

## Usage as a Skill

Place the `SKILL.md` file in your skills directory:

```
~/.workbuddy/skills/workbuddy-ppt-design/SKILL.md
```

The skill will be automatically available when creating PowerPoint presentations.

## File Structure

```
.
├── SKILL.md              # Core design specification (685 lines)
├── README.md             # This file
├── LICENSE               # Apache 2.0
├── CHANGELOG.md          # Version history
├── .gitignore            # Git exclusions
└── examples/
    ├── minimal_example.py   # Working 2-slide demo
    └── requirements.txt     # Python dependencies
```

## Key Technical Details

### Shadow Removal (Critical)

python-pptx auto-attaches `<p:style>` to connectors referencing theme effects containing `outerShdw`. This skill provides:

1. **Inline removal** - Strip `<p:style>` immediately after creating each connector
2. **Theme cleanup** - Post-save XML processing to remove all shadow/3D effects from theme

### Font Handling

All Chinese text requires explicit East Asian font setting via `set_ea_font()` to render as KaiTi instead of the default English font.

## Dependencies

- Python 3.8+
- python-pptx >= 0.6.21
- lxml >= 4.9.0

## License

Apache License 2.0 - See [LICENSE](LICENSE) for details.

## Contributing

Issues and Pull Requests are welcome. Please ensure any contributions follow the design principles outlined in `SKILL.md`.
