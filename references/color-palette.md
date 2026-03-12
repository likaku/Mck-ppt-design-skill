# Color Palette Quick Reference

| Name | Hex | RGB | Usage |
|------|-----|-----|-------|
| NAVY | #051C2C | (5, 28, 44) | Primary — titles, circles, TOC highlights |
| BLACK | #000000 | (0, 0, 0) | Separator lines, title underlines |
| WHITE | #FFFFFF | (255, 255, 255) | Backgrounds, text on navy |
| DARK_GRAY | #333333 | (51, 51, 51) | Body text |
| MED_GRAY | #666666 | (102, 102, 102) | Secondary text, labels, source notes |
| LINE_GRAY | #CCCCCC | (204, 204, 204) | Table row separators |
| BG_GRAY | #F2F2F2 | (242, 242, 242) | Background panels, takeaway areas |

## Python Constants

```python
NAVY      = RGBColor(0x05, 0x1C, 0x2C)
BLACK     = RGBColor(0x00, 0x00, 0x00)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
MED_GRAY  = RGBColor(0x66, 0x66, 0x66)
LINE_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
BG_GRAY   = RGBColor(0xF2, 0xF2, 0xF2)
```

## Accent Colors (v1.7.0 — Dark Palette)

For multi-item visual differentiation (3+ parallel items):

| Name | Hex | RGB | Light BG | Usage |
|------|-----|-----|----------|-------|
| ACCENT_BLUE | #0A2E4D | (10, 46, 77) | #E8EDF2 | First item accent |
| ACCENT_GREEN | #0C3626 | (12, 54, 38) | #E6EDE9 | Second item accent |
| ACCENT_ORANGE | #4D2E0A | (77, 46, 10) | #F0EBE4 | Third item accent |
| ACCENT_RED | #4A1015 | (74, 16, 21) | #F0E6E7 | Fourth item / warning |

```python
ACCENT_BLUE   = RGBColor(0x0A, 0x2E, 0x4D)
ACCENT_GREEN  = RGBColor(0x0C, 0x36, 0x26)
ACCENT_ORANGE = RGBColor(0x4D, 0x2E, 0x0A)
ACCENT_RED    = RGBColor(0x4A, 0x10, 0x15)
LIGHT_BLUE    = RGBColor(0xE8, 0xED, 0xF2)
LIGHT_GREEN   = RGBColor(0xE6, 0xED, 0xE9)
LIGHT_ORANGE  = RGBColor(0xF0, 0xEB, 0xE4)
LIGHT_RED     = RGBColor(0xF0, 0xE6, 0xE7)
```

## Font Size Hierarchy

| Size | Usage |
|------|-------|
| 44pt | Cover title only |
| 28pt | Section header |
| 22pt | Action title (bold, Georgia) |
| 18pt | Sub-header |
| 14pt | Body text (primary) |
| 9pt  | Footnote / source |
