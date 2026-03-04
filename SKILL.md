---
name: mck-ppt-design
description: "Create professional, consultant-grade PowerPoint presentations from scratch using python-pptx with McKinsey-style design. Use when user asks to create slides, pitch decks, business presentations, strategy decks, quarterly reviews, board meeting slides, or any professional PPTX. Generates clean, flat-design presentations with 36 layout patterns, consistent typography, and zero file-corruption issues."
license: Apache-2.0
version: "1.3.0"
homepage: https://github.com/likaku/Mck-ppt-design-skill
user-invocable: true
allowed-tools:
  - Read
  - Write
  - Bash
metadata: {"openclaw":{"emoji":"📊","requires":{"bins":["python3","pip"]}}}
---

# McKinsey PPT Design Framework

## Overview

This skill encodes the complete design specification for **professional business presentations** — a consultant-grade PowerPoint framework based on McKinsey design principles. It includes:

- **36 layout patterns** across 7 categories (structure, data, framework, comparison, narrative, timeline, team)
- **Color system** and strict typography hierarchy
- **Python-pptx code patterns** ready to copy and customize
- **Three-layer defense** against file corruption (zero `p:style` leaks)
- **Chinese + English font handling** (KaiTi / Georgia / Arial)

All specifications have been refined through iterative user feedback to ensure visual consistency and professional polish.

---

## When to Use This Skill

Use this skill when users ask to:

1. **Create presentations** — pitch decks, strategy presentations, quarterly reviews, board meeting slides, consulting deliverables, project proposals, business plans
2. **Generate slides programmatically** — using python-pptx to produce .pptx files from scratch
3. **Apply professional design** — McKinsey / BCG / Bain consulting style, clean flat design, no shadows or gradients
4. **Build specific slide types** — cover pages, data dashboards, 2x2 matrices, timelines, funnels, team introductions, executive summaries, comparison layouts
5. **Fix PPT issues** — file corruption ("needs repair"), shadow/3D artifacts, inconsistent fonts, Chinese text rendering problems
6. **Maintain design consistency** — unified color palette, font hierarchy, spacing, and line treatments across all slides

---

## Core Design Philosophy

### McKinsey Design Principles

1. **Extreme Minimalism** - Remove all non-essential visual elements
   - No color blocks unless absolutely necessary
   - Lines: thin, flat, no shadows or 3D effects
   - Shapes: simple, clean, no gradients
   - Text: clear hierarchy, maximum readability

2. **Consistency** - Repeat visual language across all slides
   - Unified color palette (navy + cyan + grays)
   - Consistent font sizes and weights for same content types
   - Aligned spacing and margins
   - Matching line widths and styles

3. **Hierarchy** - Guide viewer through information
   - Title bar (22pt) → Sub-headers (18pt) → Body (14pt) → Details (9pt)
   - Navy for primary elements, gray for secondary, black for divisions
   - Visual weight through bold, color, size (not through effects)

4. **Flat Design** - No 3D, shadows, or gradients
   - Pure solid colors only
   - All lines are simple strokes with no effects
   - Shapes have no shadow or reflection effects
   - Circles are solid fills, not 3D spheres

---

## Design Specifications

### Color Palette

All colors in RGB format for python-pptx:

| Color Name | Hex | RGB | Usage | Notes |
|-----------|-----|-----|-------|-------|
| **NAVY** | #051C2C | (5, 28, 44) | Primary, titles, circles | Corporate, formal tone |
| **CYAN** | #00A9F4 | (0, 169, 244) | Originally used in v1 | **DEPRECATED** - Use NAVY for consistency |
| **WHITE** | #FFFFFF | (255, 255, 255) | Backgrounds, text | On navy backgrounds only |
| **BLACK** | #000000 | (0, 0, 0) | Lines, text separators | For clarity and contrast |
| **DARK_GRAY** | #333333 | (51, 51, 51) | Body text, descriptions | Main content text |
| **MED_GRAY** | #666666 | (102, 102, 102) | Secondary text, labels | Softer tone than DARK_GRAY |
| **LINE_GRAY** | #CCCCCC | (204, 204, 204) | Light separators, table rows | Table separators only |
| **BG_GRAY** | #F2F2F2 | (242, 242, 242) | Background panels | Takeaway/highlight areas |

**Key Rule**: Use navy (#051C2C) everywhere, especially for:
- All circle indicators (A, B, C, 1, 2, 3)
- All action titles
- All primary section headers
- All TOC highlight colors

---

### Typography System

#### Font Families

```
English Headers:  Georgia (serif, elegant)
English Body:     Arial (sans-serif, clean)
Chinese (ALL):    KaiTi (楷体, traditional brush style)
                  (fallback: SimSun 宋体)
```

**Critical Implementation**:
```python
def set_ea_font(run, typeface='KaiTi'):
    """Set East Asian font for Chinese text"""
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', typeface)
```

Every paragraph with Chinese text MUST apply `set_ea_font()` to all runs.

#### Font Size Hierarchy

| Size | Type | Examples | Notes |
|------|------|----------|-------|
| **44pt** | Cover Title | "项目名称" | Cover slide only, Georgia |
| **28pt** | Section Header | "目录" (TOC title) | Largest body content, Georgia |
| **24pt** | Subtitle | Tagline on cover | Cover slide only |
| **22pt** | Action Title | Slide title bars | Main content titles, **bold**, Georgia |
| **18pt** | Sub-Header | Column headers, section names | Supporting titles |
| **16pt** | Emphasis Text | Bottom takeaway on slide 8 | Callout text, bold |
| **14pt** | Body Text | Tables, lists, descriptions | **PRIMARY BODY SIZE**, all main content |
| **9pt** | Footnote | Source attribution | Smallest, gray color only |

**No other sizes should be used** - stick to this hierarchy exclusively.

---

### Line Treatment (CRITICAL)

#### Line Rendering Rules

1. **All lines are FLAT** - no shadows, no effects, no 3D
2. **Remove theme style references** - prevents automatic shadow application
3. **Solid color only** - no gradients or patterns
4. **Width varies by context** - see table below

#### Line Width Specifications

| Usage | Width | Color | Context |
|-------|-------|-------|---------|
| **Title separator** (under action titles) | 0.5pt | BLACK | Below 22pt title |
| **Column/section divider** (under headers) | 0.5pt | BLACK | Below 18pt headers |
| **Table header line** | 1.0pt | BLACK | Between header and first row |
| **Table row separator** | 0.5pt | LINE_GRAY (#CCCCCC) | Between table rows |
| **Timeline line** (roadmap) | 0.75pt | LINE_GRAY | Background for phase indicators |
| **Cover accent line** | 2.0pt | NAVY | Decorative bottom-left on cover |
| **Column internal divider** | 0.5pt | BLACK | Between "是什么" and "独到之处" |

#### Code Implementation (v1.1 — Rectangle-based Lines)

**CRITICAL**: Do NOT use `slide.shapes.add_connector()` for lines. Connectors carry `<p:style>` elements that reference theme effects and cause file corruption. Instead, draw lines as ultra-thin rectangles:

```python
def add_hline(slide, x, y, length, color=BLACK, thickness=Pt(0.5)):
    """Draw a horizontal line using a thin rectangle (no connector, no p:style)."""
    from pptx.util import Emu
    h = max(int(thickness), Emu(6350))  # minimum ~0.5pt
    return add_rect(slide, x, y, length, h, color)
```

**Why rectangle lines**: python-pptx `add_connector()` automatically attaches `<p:style>` with `effectRef idx="2"`, which references theme effects including `outerShdw` (outer shadow). Even after removing `<p:style>` at creation time, PowerPoint may re-apply effects on open. Using thin rectangles eliminates this class of bugs entirely.

**Old approach (DEPRECATED — do NOT use)**:
```python
# DEPRECATED: causes file corruption in some PowerPoint versions
# c = slide.shapes.add_connector(1, x1, y1, x2, y2)
```

#### Post-Save Full Cleanup (v1.1 — Nuclear Sanitization)

After `prs.save(outpath)`, ALWAYS run full cleanup that sanitizes **both** theme XML **and** all slide XML:

```python
import zipfile, os
from lxml import etree

def full_cleanup(outpath):
    """Remove ALL p:style from every slide + theme shadows/3D."""
    tmppath = outpath + '.tmp'
    with zipfile.ZipFile(outpath, 'r') as zin:
        with zipfile.ZipFile(tmppath, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith('.xml'):
                    root = etree.fromstring(data)
                    ns_p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
                    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    # Remove ALL p:style elements from all shapes/connectors
                    for style in root.findall(f'.//{{{ns_p}}}style'):
                        style.getparent().remove(style)
                    # Remove shadows and 3D from theme
                    if 'theme' in item.filename.lower():
                        for tag in ['outerShdw', 'innerShdw', 'scene3d', 'sp3d']:
                            for el in root.findall(f'.//{{{ns_a}}}{tag}'):
                                el.getparent().remove(el)
                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
    os.replace(tmppath, outpath)
```

**v1.1 improvement**: The old cleanup only handled theme XML. The new `full_cleanup()` also removes `<p:style>` from every shape in every slide, preventing `effectRef idx="2"` references that cause "File needs repair" errors.

---

### Text Box & Shape Treatment

#### Text Box Padding

All text boxes must have consistent internal padding to prevent text touching edges:

```python
bodyPr = tf._txBody.find(qn('a:bodyPr'))
# All margins: 45720 EMUs = ~0.05 inches
for attr in ['lIns','tIns','rIns','bIns']:
    bodyPr.set(attr, '45720')
```

#### Vertical Anchoring

Text must be anchored correctly based on usage:

| Content Type | Anchor | Code | Notes |
|--------------|--------|------|-------|
| Action titles | MIDDLE | `anchor='ctr'` | Centered vertically in bar |
| Body text | TOP | `anchor='t'` | Default, aligns to top |
| Labels | CENTER | `anchor='ctr'` | For circle labels |

```python
anchor_map = {
    MSO_ANCHOR.MIDDLE: 'ctr', 
    MSO_ANCHOR.BOTTOM: 'b', 
    MSO_ANCHOR.TOP: 't'
}
bodyPr.set('anchor', anchor_map.get(anchor, 't'))
```

#### Shape Styling

All shapes (rectangles, circles) must have:
- Solid fill color (no gradients)
- NO border/line (`shape.line.fill.background()`)
- **p:style removed** immediately after creation (`_clean_shape()`)
- No shadow effects (enforced by both inline cleanup and post-save full_cleanup)

```python
def _clean_shape(shape):
    """Remove p:style from any shape to prevent effect references."""
    sp = shape._element
    style = sp.find(qn('p:style'))
    if style is not None:
        sp.remove(style)

shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
shape.fill.solid()
shape.fill.fore_color.rgb = BG_GRAY
shape.line.fill.background()  # CRITICAL: removes border
_clean_shape(shape)            # CRITICAL: removes p:style
```

---

## Layout Patterns

### Slide Dimensions

```python
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
```

Widescreen format (16:9), standard for all presentations.

### Standard Margin/Padding

| Position | Size | Usage |
|----------|------|-------|
| **Left margin** | 0.8" | Default left edge |
| **Right margin** | 0.8" | Default right edge |
| **Top (below title)** | 1.4" | Content start position |
| **Bottom** | 7.05" | Source text baseline |
| **Title bar height** | 0.9" | Action title bar |
| **Title bar top** | 0.15" | From slide top |

### Slide Type Patterns

#### 1. Cover Slide (Slide 1)

Layout:
- Navy bar at very top (0.05" height)
- Main title centered (44pt, Georgia, navy) at y=2.2"
- Subtitle (24pt, dark gray) at y=3.5"
- Date/info (14pt, med gray) at y=4.5"
- Decorative navy line at x=1", y=6.8" (2" wide, 2pt)

Code template:
```python
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(s1, 0, 0, prs.slide_width, Inches(0.05), NAVY)
add_text(s1, Inches(1), Inches(2.2), Inches(11), Inches(1.0),
         '项目名称', font_size=Pt(44), font_name='Georgia',
         font_color=NAVY, bold=True, ea_font='KaiTi')
add_text(s1, Inches(1), Inches(3.5), Inches(11), Inches(0.6),
         '副标题描述', font_size=Pt(24),
         font_color=DARK_GRAY, ea_font='KaiTi')
add_text(s1, Inches(1), Inches(4.5), Inches(11), Inches(0.5),
         '演示文稿  |  2026年3月', font_size=BODY_SIZE,
         font_color=MED_GRAY, ea_font='KaiTi')
add_line(s1, Inches(1), Inches(6.8), Inches(4), Inches(6.8),
         color=NAVY, width=Pt(2))
```

#### 2. Action Title Slide (Most Content Slides)

Every main content slide has this structure:

```
┌─────────────────────────────────────────┐ 0.15"
│ ▌ Action Title (22pt, bold, black)      │ ← TITLE_BAR_H = 0.9"
├─────────────────────────────────────────┤ 1.05"
│                                         │
│  Content area (starts at 1.4")          │
│  [Tables, lists, text, etc.]            │
│                                         │
│                                         │
│  ──────────────────────────────────────  │ 7.05"
│  Source: ...                            │ 9pt, med gray
└─────────────────────────────────────────┘ 7.5"
```

Code pattern:
```python
s = prs.slides.add_slide(prs.slide_layouts[6])
add_action_title(s, 'Slide Title Here')
# Then add content below y=1.4"
add_source(s, 'Source attribution')
```

#### 3. Table Layout (Slide 4 - Five Capabilities)

Structure:
- Header row with column names (BODY_SIZE, gray, bold)
- 1.0pt black line under header
- Data rows (1.0" height each, 14pt text)
- 0.5pt gray line between rows
- 3 columns: Module (1.6" wide), Function (5.0"), Scene (5.1")

```python
# Headers
add_text(s4, left, top, width, height, text,
         font_size=BODY_SIZE, font_color=MED_GRAY, bold=True)

# Header line (thicker)
add_line(s4, left, top + Inches(0.5), left + full_width, top + Inches(0.5),
         color=BLACK, width=Pt(1.0))

# Rows
for i, (col1, col2, col3) in enumerate(rows):
    y = header_y + row_height * i
    add_text(s4, left, y, col1_w, row_h, col1, ...)
    add_text(s4, left + col1_w, y, col2_w, row_h, col2, ...)
    add_text(s4, left + col1_w + col2_w, y, col3_w, row_h, col3, ...)
    # Row separator
    add_line(s4, left, y + row_h, left + full_w, y + row_h,
             color=LINE_GRAY, width=Pt(0.5))
```

#### 4. Three-Column Overview (Slide 5)

Layout:
- Left column (4.1" wide): "是什么"
- Middle column (4.1" wide): "独到之处"
- Right 1/4 (2.5" wide) gray panel: "Key Takeaways"

```
0.8"  4.9"  5.3"  9.4"  10.0" 12.5"
|-----|-----|-----|-----|------|
│左 │ │ 中 │ │ 右 │
└─────────────────────────────┘
```

Code:
```python
left_x = Inches(0.8)
col_w5 = Inches(4.1)
mid_x = Inches(5.3)
takeaway_left = Inches(10.0)
takeaway_width = Inches(2.5)

# Left column
add_text(s5, left_x, content_top, col_w5, ...)
add_multiline(s5, left_x, content_top + Inches(0.6), col_w5, ..., 
              bullet=True, line_spacing_pt=8)

# Right gray takeaway area
add_rect(s5, takeaway_left, Inches(1.2), takeaway_width, Inches(5.6), BG_GRAY)
add_text(s5, takeaway_left + Inches(0.15), Inches(1.35), takeaway_width - Inches(0.3), ...,
         'Key Takeaways', font_size=BODY_SIZE, color=NAVY, bold=True)
add_multiline(s5, takeaway_left + Inches(0.15), Inches(1.9), takeaway_width - Inches(0.3), ...,
              [f'{i+1}. {t}' for i, t in enumerate(takeaways)], line_spacing_pt=10)
```

---

### 类别 A：结构导航

#### 5. Section Divider (章节分隔页)

**适用场景**: 多章节演示文稿的章节过渡页，用于视觉上分隔不同主题模块。

```
┌──┬──────────────────────────────────────┐
│N │                                      │
│A │  第一部分                             │
│V │  章节标题（28pt, NAVY, bold）          │
│Y │  副标题说明文字                        │
│  │                                      │
└──┴──────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_rect(s, 0, 0, Inches(0.6), SH, NAVY)
add_text(s, Inches(1.2), Inches(2.0), Inches(10), Inches(0.8),
         '第一部分', font_size=SUB_HEADER_SIZE, font_color=MED_GRAY, font_name='Georgia')
add_text(s, Inches(1.2), Inches(2.8), Inches(10), Inches(1.2),
         '章节标题', font_size=HEADER_SIZE, font_color=NAVY, bold=True, font_name='Georgia')
add_text(s, Inches(1.2), Inches(4.2), Inches(10), Inches(0.6),
         '副标题说明文字', font_size=BODY_SIZE, font_color=DARK_GRAY)
```

#### 6. Table of Contents / Agenda (目录/议程页)

**适用场景**: 演示文稿开头的目录或会议议程，列出各章节及说明。

```
┌─────────────────────────────────────────┐
│ ▌ 目录                                  │
├─────────────────────────────────────────┤
│                                         │
│  (1)  章节一标题     简要描述            │
│  ─────────────────────────────────────  │
│  (2)  章节二标题     简要描述            │
│  ─────────────────────────────────────  │
│  (3)  章节三标题     简要描述            │
│                                         │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '目录')
items = [('1', '引言与背景', '项目起源与核心问题'),
         ('2', '市场分析', '竞争格局与机会识别'),
         ('3', '战略建议', '三大核心行动方案')]
iy = Inches(1.6)
for num, title, desc in items:
    add_oval(s, LM, iy, num, size=Inches(0.5))
    add_text(s, LM + Inches(0.7), iy, Inches(4.0), Inches(0.4),
             title, font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True)
    add_text(s, Inches(5.5), iy + Inches(0.05), Inches(6.5), Inches(0.4),
             desc, font_size=BODY_SIZE, font_color=MED_GRAY)
    iy += Inches(0.7)
    add_hline(s, LM, iy, CONTENT_W, LINE_GRAY)
    iy += Inches(0.3)
```

#### 7. Appendix Title (附录标题页)

**适用场景**: 正文结束后的附录/备用材料分隔页。

```
┌─────────────────────────────────────────┐
│                                         │
│                                         │
│           附录                           │
│           Appendix                      │
│           ────────                      │
│           补充数据与参考资料              │
│                                         │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_rect(s, 0, 0, SW, Inches(0.05), NAVY)
add_text(s, Inches(1), Inches(2.5), Inches(11.3), Inches(1.0),
         '附录', font_size=Pt(36), font_color=NAVY, bold=True,
         font_name='Georgia', alignment=PP_ALIGN.CENTER)
add_hline(s, Inches(5.5), Inches(3.8), Inches(2.3), NAVY, Pt(1.5))
add_text(s, Inches(1), Inches(4.2), Inches(11.3), Inches(0.5),
         '补充数据与参考资料', font_size=BODY_SIZE, font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER)
```

---

### 类别 B：数据统计

#### 8. Big Number / Factoid (大数据展示页)

**适用场景**: 用一个醒目的大数字引出核心发现或关键数据点。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│                                         │
│  ┌─NAVY─────────┐                       │
│  │    95%        │   右侧上下文说明      │
│  │  子标题       │   详细解释数据含义     │
│  └──────────────┘                       │
│                                         │
│  ┌─BG_GRAY──────────────────────────┐   │
│  │  关键洞见：详细分析文字            │   │
│  └──────────────────────────────────┘   │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '关键发现标题')
add_rect(s, LM, Inches(1.4), Inches(3.5), Inches(1.8), NAVY)
add_text(s, LM + Inches(0.2), Inches(1.5), Inches(3.1), Inches(0.8),
         '95%', font_size=Pt(44), font_color=WHITE, bold=True,
         font_name='Georgia', alignment=PP_ALIGN.CENTER)
add_text(s, LM + Inches(0.2), Inches(2.3), Inches(3.1), Inches(0.7),
         '描述数据含义', font_size=Pt(12), font_color=WHITE, alignment=PP_ALIGN.CENTER)
add_text(s, Inches(5.0), Inches(1.5), Inches(7.5), Inches(2.0),
         '上下文说明与详细解释', font_size=BODY_SIZE)
add_rect(s, LM, Inches(4.5), CONTENT_W, Inches(2.2), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(4.6), Inches(1.5), Inches(0.4),
         '关键洞见', font_size=BODY_SIZE, font_color=NAVY, bold=True)
add_multiline(s, LM + Inches(0.3), Inches(5.1), CONTENT_W - Inches(0.6), Inches(1.4),
              ['洞见要点一', '洞见要点二', '洞见要点三'], line_spacing_pt=8)
add_source(s, 'Source: ...')
```

#### 9. Two-Stat Comparison (双数据对比页)

**适用场景**: 并排展示两个关键指标的对比（如同比、环比、A vs B）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│                                         │
│  ┌──NAVY───────┐    ┌──BG_GRAY────┐     │
│  │   $2.4B     │    │   $1.8B     │     │
│  │  2026年目标  │    │  2025年实际  │     │
│  └─────────────┘    └─────────────┘     │
│                                         │
│  分析说明文字                            │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '对比标题')
stats = [('$2.4B', '2026年目标', True), ('$1.8B', '2025年实际', False)]
sw = Inches(5.5)
sg = Inches(0.733)
for i, (big, small, is_navy) in enumerate(stats):
    sx = LM + (sw + sg) * i
    fill = NAVY if is_navy else BG_GRAY
    bc = WHITE if is_navy else NAVY
    sc = WHITE if is_navy else DARK_GRAY
    add_rect(s, sx, Inches(1.8), sw, Inches(2.5), fill)
    add_text(s, sx + Inches(0.3), Inches(2.0), sw - Inches(0.6), Inches(1.0),
             big, font_size=Pt(44), font_color=bc, bold=True,
             font_name='Georgia', alignment=PP_ALIGN.CENTER)
    add_text(s, sx + Inches(0.3), Inches(3.2), sw - Inches(0.6), Inches(0.5),
             small, font_size=BODY_SIZE, font_color=sc, alignment=PP_ALIGN.CENTER)
add_text(s, LM, Inches(5.0), CONTENT_W, Inches(1.5),
         '分析说明文字', font_size=BODY_SIZE)
add_source(s, 'Source: ...')
```

#### 10. Three-Stat Dashboard (三指标仪表盘)

**适用场景**: 同时展示三个关键业务指标（如 KPI、季度数据）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  ┌──NAVY──┐   ┌──BG_GRAY─┐  ┌──BG_GRAY─┐│
│  │  数字1  │   │  数字2   │  │  数字3   ││
│  │ 小标题  │   │  小标题  │  │  小标题  ││
│  └────────┘   └─────────┘  └─────────┘│
│                                         │
│  详细说明文字                            │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '三大关键指标')
stats = [('87%', '客户满意度', True),
         ('+23%', '同比增长', False),
         ('4.2x', '投资回报率', False)]
sw = Inches(3.5)
sg = (CONTENT_W - sw * 3) / 2
for i, (big, small, is_navy) in enumerate(stats):
    sx = LM + (sw + sg) * i
    fill = NAVY if is_navy else BG_GRAY
    bc = WHITE if is_navy else NAVY
    sc = WHITE if is_navy else DARK_GRAY
    add_rect(s, sx, Inches(1.4), sw, Inches(1.8), fill)
    add_text(s, sx + Inches(0.2), Inches(1.5), sw - Inches(0.4), Inches(0.7),
             big, font_size=Pt(28), font_color=bc, bold=True,
             font_name='Georgia', alignment=PP_ALIGN.CENTER)
    add_text(s, sx + Inches(0.2), Inches(2.25), sw - Inches(0.4), Inches(0.6),
             small, font_size=Pt(12), font_color=sc, alignment=PP_ALIGN.CENTER)
add_source(s, 'Source: ...')
```

#### 11. Data Table with Headers (数据表格页)

**适用场景**: 结构化数据展示，如财务数据、功能对比矩阵、项目清单。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  列1         列2         列3     列4    │
│  ═══════════════════════════════════    │
│  数据1-1     数据1-2     ...     ...    │
│  ───────────────────────────────────    │
│  数据2-1     数据2-2     ...     ...    │
│  ───────────────────────────────────    │
│  数据3-1     数据3-2     ...     ...    │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '数据概览')
headers = ['模块', '功能', '状态', '负责人']
col_widths = [Inches(2.5), Inches(4.0), Inches(2.5), Inches(2.7)]
hdr_y = Inches(1.5)
cx = LM
for hdr, cw in zip(headers, col_widths):
    add_text(s, cx, hdr_y, cw, Inches(0.4), hdr,
             font_size=BODY_SIZE, font_color=MED_GRAY, bold=True)
    cx += cw
add_hline(s, LM, Inches(2.0), CONTENT_W, BLACK, Pt(1.0))
# Data rows
rows = [['模块A', '核心功能描述', '已上线', '张三'], ...]
row_h = Inches(0.8)
for ri, row in enumerate(rows):
    ry = Inches(2.1) + row_h * ri
    cx = LM
    for val, cw in zip(row, col_widths):
        add_text(s, cx, ry, cw, row_h, val, font_size=BODY_SIZE)
        cx += cw
    add_hline(s, LM, ry + row_h, CONTENT_W, LINE_GRAY)
add_source(s, 'Source: ...')
```

#### 12. Metric Cards Row (指标卡片行)

**适用场景**: 3-4个并排卡片展示独立指标，每个卡片含标题+描述。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│ ┌─BG_GRAY─┐ ┌─BG_GRAY─┐ ┌─BG_GRAY─┐   │
│ │ (A)     │ │ (B)     │ │ (C)     │   │
│ │ 标题    │ │ 标题    │ │ 标题    │   │
│ │ ───     │ │ ───     │ │ ───     │   │
│ │ 描述    │ │ 描述    │ │ 描述    │   │
│ └─────────┘ └─────────┘ └─────────┘   │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '核心指标概览')
cards = [('A', '用户增长', '月活用户达到 120 万\n同比增长 35%'),
         ('B', '营收表现', '季度营收 ¥8,500 万\n超出预期 12%'),
         ('C', '运营效率', '客诉响应时间 < 2h\n满意度 94%')]
cw = Inches(3.5)
cg = (CONTENT_W - cw * 3) / 2
for i, (letter, title, desc) in enumerate(cards):
    cx = LM + (cw + cg) * i
    cy = Inches(1.5)
    add_rect(s, cx, cy, cw, Inches(4.5), BG_GRAY)
    add_oval(s, cx + Inches(1.5), cy + Inches(0.2), letter)
    add_text(s, cx + Inches(0.2), cy + Inches(0.8), cw - Inches(0.4), Inches(0.4),
             title, font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER)
    add_hline(s, cx + Inches(0.4), cy + Inches(1.3), cw - Inches(0.8), LINE_GRAY)
    add_multiline(s, cx + Inches(0.2), cy + Inches(1.5), cw - Inches(0.4), Inches(2.5),
                  desc.split('\n'), line_spacing_pt=8, alignment=PP_ALIGN.CENTER)
add_source(s, 'Source: ...')
```

---

### 类别 C：框架矩阵

#### 13. 2x2 Matrix (四象限矩阵)

**适用场景**: 战略分析（如 BCG 矩阵、优先级排序、风险评估）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│         高 Y轴                           │
│  ┌─NAVY──────┐  ┌─BG_GRAY───┐          │
│  │ 左上象限   │  │ 右上象限   │          │
│  └───────────┘  └───────────┘          │
│  ┌─BG_GRAY───┐  ┌─BG_GRAY───┐          │
│  │ 左下象限   │  │ 右下象限   │          │
│  └───────────┘  └───────────┘          │
│         低           高 X轴              │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '战略优先级矩阵')
mx = LM + Inches(1.5)
my = Inches(1.8)
cw = Inches(4.5)
ch = Inches(2.5)
# Axis labels
add_text(s, mx - Inches(1.3), my + Inches(0.8), Inches(1.1), Inches(0.4),
         '高影响', font_size=BODY_SIZE, font_color=NAVY, bold=True, alignment=PP_ALIGN.CENTER)
add_text(s, mx + Inches(0.8), my - Inches(0.5), Inches(3.0), Inches(0.4),
         '高可行性', font_size=BODY_SIZE, font_color=NAVY, bold=True, alignment=PP_ALIGN.CENTER)
# Quadrants
add_rect(s, mx, my, cw, ch, NAVY)  # best quadrant
add_rect(s, mx + cw + Inches(0.15), my, cw, ch, BG_GRAY)
add_rect(s, mx, my + ch + Inches(0.15), cw, ch, BG_GRAY)
add_rect(s, mx + cw + Inches(0.15), my + ch + Inches(0.15), cw, ch, BG_GRAY)
# Quadrant titles + descriptions
add_text(s, mx + Inches(0.3), my + Inches(0.3), cw - Inches(0.6), Inches(0.5),
         '立即执行', font_size=SUB_HEADER_SIZE, font_color=WHITE, bold=True)
# ... repeat for other 3 quadrants with DARK_GRAY text
add_source(s, 'Source: ...')
```

#### 14. Three-Pillar Framework (三支柱框架)

**适用场景**: 展示三个并列的核心策略、能力或主题模块。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│ ┌──NAVY──┐   ┌──NAVY──┐   ┌──NAVY──┐   │
│ │ 标题1  │   │ 标题2  │   │ 标题3  │   │
│ ├────────┤   ├────────┤   ├────────┤   │
│ │ 要点   │   │ 要点   │   │ 要点   │   │
│ │ 要点   │   │ 要点   │   │ 要点   │   │
│ │ 要点   │   │ 要点   │   │ 要点   │   │
│ └────────┘   └────────┘   └────────┘   │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '三大核心战略')
pillars = [('数字化转型', ['建设数据中台', '打通全渠道', '自动化运营']),
           ('组织升级', ['扁平化管理', '敏捷团队', '人才梯队']),
           ('客户深耕', ['精细化运营', '会员体系', 'LTV 提升'])]
pw = Inches(3.5)
pg = (CONTENT_W - pw * 3) / 2
for i, (title, points) in enumerate(pillars):
    px = LM + (pw + pg) * i
    add_rect(s, px, Inches(1.5), pw, Inches(0.6), NAVY)
    add_text(s, px + Inches(0.15), Inches(1.5), pw - Inches(0.3), Inches(0.6),
             title, font_size=SUB_HEADER_SIZE, font_color=WHITE, bold=True,
             anchor=MSO_ANCHOR.MIDDLE, alignment=PP_ALIGN.CENTER)
    add_rect(s, px, Inches(2.1), pw, Inches(4.0), BG_GRAY)
    add_multiline(s, px + Inches(0.2), Inches(2.3), pw - Inches(0.4), Inches(3.5),
                  [f'• {p}' for p in points], line_spacing_pt=10)
add_source(s, 'Source: ...')
```

#### 15. Pyramid / Hierarchy (金字塔/层级图)

**适用场景**: 展示层级关系（如 Maslow 需求层次、战略-战术-执行分层）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│           ┌──NAVY──┐                    │
│           │ 愿景   │    右侧说明        │
│         ┌─┴────────┴─┐                  │
│         │  战略目标   │  右侧说明        │
│       ┌─┴────────────┴─┐                │
│       │   执行措施      │  右侧说明      │
│       └────────────────┘                │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '战略层级框架')
levels = [('愿景', '成为行业第一', Inches(3.5)),
          ('战略目标', '三年收入翻倍', Inches(5.0)),
          ('执行措施', '渠道+产品+组织', Inches(6.5))]
for i, (label, desc, w) in enumerate(levels):
    lx = Inches(6.666) - w / 2  # centered
    ly = Inches(1.8) + Inches(1.5) * i
    h = Inches(1.2)
    fill = NAVY if i == 0 else BG_GRAY
    tc = WHITE if i == 0 else NAVY
    add_rect(s, lx, ly, w, h, fill)
    add_text(s, lx + Inches(0.2), ly + Inches(0.1), w - Inches(0.4), Inches(0.4),
             label, font_size=SUB_HEADER_SIZE, font_color=tc, bold=True,
             alignment=PP_ALIGN.CENTER)
    add_text(s, lx + Inches(0.2), ly + Inches(0.55), w - Inches(0.4), Inches(0.5),
             desc, font_size=BODY_SIZE, font_color=tc if i == 0 else DARK_GRAY,
             alignment=PP_ALIGN.CENTER)
add_source(s, 'Source: ...')
```

#### 16. Process Chevron (流程箭头页)

**适用场景**: 线性流程展示（3-5步），如实施路径、业务流程、方法论步骤。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│                                         │
│  ┌NAVY┐ -> ┌GRAY┐ -> ┌GRAY┐ -> ┌GRAY┐  │
│  │ S1 │    │ S2 │    │ S3 │    │ S4 │  │
│  └────┘    └────┘    └────┘    └────┘  │
│   描述      描述      描述      描述    │
│                                         │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '实施路径')
steps = [('诊断', '现状评估\n痛点识别'),
         ('设计', '方案制定\n资源规划'),
         ('实施', '分阶段落地\n快速迭代'),
         ('优化', '效果追踪\n持续改进')]
sw = Inches(2.5)
sg = (CONTENT_W - sw * len(steps)) / (len(steps) - 1)
for i, (title, desc) in enumerate(steps):
    sx = LM + (sw + sg) * i
    fill = NAVY if i == 0 else BG_GRAY
    tc = WHITE if i == 0 else NAVY
    add_rect(s, sx, Inches(2.0), sw, Inches(1.2), fill)
    add_oval(s, sx + Inches(0.1), Inches(2.1), str(i + 1),
             bg=WHITE if i == 0 else NAVY, fg=NAVY if i == 0 else WHITE)
    add_text(s, sx + Inches(0.6), Inches(2.1), sw - Inches(0.8), Inches(1.0),
             title, font_size=SUB_HEADER_SIZE, font_color=tc, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, sx + Inches(0.1), Inches(3.4), sw - Inches(0.2), Inches(1.5),
             desc, font_size=BODY_SIZE, alignment=PP_ALIGN.CENTER)
    if i < len(steps) - 1:
        add_text(s, sx + sw + Inches(0.05), Inches(2.3), Inches(0.4), Inches(0.5),
                 '->', font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True)
add_source(s, 'Source: ...')
```

#### 17. Venn Diagram Concept (维恩图概念页)

**适用场景**: 展示两三个概念的交集关系（如能力交叉、市场定位）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│          ┌──BG──┐                       │
│         ╱概念A  ╲                       │
│        ╱  ┌──┐   ╲      右侧说明       │
│       │   │交│    │                     │
│        ╲  └──┘   ╱                     │
│         ╲概念B  ╱                       │
│          └──────┘                       │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '核心能力交叉')
# Use overlapping rectangles to represent Venn concept
add_rect(s, Inches(1.5), Inches(1.8), Inches(4.5), Inches(4.0), BG_GRAY)
add_text(s, Inches(1.7), Inches(2.0), Inches(2.0), Inches(0.4),
         '技术能力', font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True)
add_rect(s, Inches(3.5), Inches(2.8), Inches(4.5), Inches(4.0), BG_GRAY)
add_text(s, Inches(5.5), Inches(5.5), Inches(2.0), Inches(0.4),
         '业务洞察', font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True)
# Intersection area
add_rect(s, Inches(3.5), Inches(2.8), Inches(2.5), Inches(3.0), NAVY)
add_text(s, Inches(3.7), Inches(3.5), Inches(2.1), Inches(0.8),
         '核心\n竞争力', font_size=SUB_HEADER_SIZE, font_color=WHITE, bold=True,
         alignment=PP_ALIGN.CENTER)
# Right explanation
add_text(s, Inches(9.0), Inches(2.0), Inches(3.5), Inches(4.0),
         '当技术能力与业务洞察交叉时...', font_size=BODY_SIZE)
add_source(s, 'Source: ...')
```

#### 18. Temple / House Framework (殿堂框架)

**适用场景**: 展示"屋顶-支柱-基座"的结构（如企业架构、能力体系）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  ┌═══════════NAVY（愿景/屋顶）══════════┐│
│  ├────┤  ├────┤  ├────┤  ├────┤        ││
│  │支柱│  │支柱│  │支柱│  │支柱│        ││
│  │ 1  │  │ 2  │  │ 3  │  │ 4  │        ││
│  ├════╧══╧════╧══╧════╧══╧════╧════════┤│
│  │        基座（基础能力/文化）          ││
│  └──────────────────────────────────────┘│
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '企业能力架构')
# Roof
add_rect(s, LM, Inches(1.5), CONTENT_W, Inches(0.8), NAVY)
add_text(s, LM + Inches(0.3), Inches(1.5), CONTENT_W - Inches(0.6), Inches(0.8),
         '企业愿景：成为行业领先的数字化平台',
         font_size=SUB_HEADER_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE, alignment=PP_ALIGN.CENTER)
# Pillars
pillars = ['产品力', '技术力', '运营力', '品牌力']
pw = Inches(2.5)
pg = (CONTENT_W - pw * 4) / 3
for i, name in enumerate(pillars):
    px = LM + (pw + pg) * i
    add_rect(s, px, Inches(2.5), pw, Inches(3.0), BG_GRAY)
    add_text(s, px + Inches(0.15), Inches(2.6), pw - Inches(0.3), Inches(0.5),
             name, font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER)
# Foundation
add_rect(s, LM, Inches(5.7), CONTENT_W, Inches(0.8), NAVY)
add_text(s, LM + Inches(0.3), Inches(5.7), CONTENT_W - Inches(0.6), Inches(0.8),
         '基座：数据驱动 + 人才体系 + 企业文化',
         font_size=BODY_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE, alignment=PP_ALIGN.CENTER)
add_source(s, 'Source: ...')
```

---

### 类别 D：对比评估

#### 19. Side-by-Side Comparison (左右对比页)

**适用场景**: 两个方案/选项/产品的并排对比分析。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  ┌──方案 A──────┐  ┌──方案 B──────┐     │
│  │ 标题（NAVY） │  │ 标题（NAVY） │     │
│  ├──────────────┤  ├──────────────┤     │
│  │ 优势         │  │ 优势         │     │
│  │ 劣势         │  │ 劣势         │     │
│  │ 成本         │  │ 成本         │     │
│  └──────────────┘  └──────────────┘     │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '方案对比分析')
cw = Inches(5.5)
cg = Inches(0.733)
options = [('方案 A：自建团队', ['投入可控', '周期较长', '成本 ¥500万/年']),
           ('方案 B：外部合作', ['快速启动', '依赖供应商', '成本 ¥300万/年'])]
for i, (title, points) in enumerate(options):
    cx = LM + (cw + cg) * i
    add_rect(s, cx, Inches(1.5), cw, Inches(0.6), NAVY)
    add_text(s, cx + Inches(0.15), Inches(1.5), cw - Inches(0.3), Inches(0.6),
             title, font_size=SUB_HEADER_SIZE, font_color=WHITE, bold=True,
             anchor=MSO_ANCHOR.MIDDLE, alignment=PP_ALIGN.CENTER)
    add_rect(s, cx, Inches(2.1), cw, Inches(4.0), BG_GRAY)
    add_multiline(s, cx + Inches(0.3), Inches(2.3), cw - Inches(0.6), Inches(3.5),
                  [f'• {p}' for p in points], line_spacing_pt=10)
add_source(s, 'Source: ...')
```

#### 20. Before / After (前后对比页)

**适用场景**: 展示变革前后的对比（如流程优化、组织变革）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  ┌──BG_GRAY────┐  ──>  ┌──NAVY────┐    │
│  │  现状       │       │  目标    │    │
│  │  (Before)   │       │  (After) │    │
│  │  痛点列表   │       │  改进点  │    │
│  └─────────────┘       └─────────┘    │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '从现状到目标')
hw = Inches(5.0)
# Before
add_rect(s, LM, Inches(1.5), hw, Inches(4.5), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(1.6), hw - Inches(0.6), Inches(0.5),
         'X  现状（Before）', font_size=SUB_HEADER_SIZE, font_color=DARK_GRAY, bold=True)
add_hline(s, LM + Inches(0.3), Inches(2.2), hw - Inches(0.6), LINE_GRAY)
add_multiline(s, LM + Inches(0.3), Inches(2.4), hw - Inches(0.6), Inches(3.0),
              ['痛点一', '痛点二', '痛点三'], line_spacing_pt=10)
# Arrow
add_text(s, LM + hw + Inches(0.1), Inches(3.2), Inches(1.5), Inches(0.5),
         '->', font_size=Pt(36), font_color=NAVY, bold=True, alignment=PP_ALIGN.CENTER)
# After
ax = LM + hw + Inches(1.733)
add_rect(s, ax, Inches(1.5), hw, Inches(4.5), NAVY)
add_text(s, ax + Inches(0.3), Inches(1.6), hw - Inches(0.6), Inches(0.5),
         'V  目标（After）', font_size=SUB_HEADER_SIZE, font_color=WHITE, bold=True)
add_hline(s, ax + Inches(0.3), Inches(2.2), hw - Inches(0.6), WHITE)
add_multiline(s, ax + Inches(0.3), Inches(2.4), hw - Inches(0.6), Inches(3.0),
              ['改进一', '改进二', '改进三'], font_color=WHITE, line_spacing_pt=10)
add_source(s, 'Source: ...')
```

#### 21. Pros and Cons (优劣分析页)

**适用场景**: 评估某个决策/方案的优势与风险。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  V 优势                  X 风险         │
│  ───────────             ──────────     │
│  • 要点1                 • 要点1        │
│  • 要点2                 • 要点2        │
│  • 要点3                 • 要点3        │
│                                         │
│  ┌──BG_GRAY 结论/建议───────────────┐   │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '方案评估：优势与风险')
hw = Inches(5.5)
# Pros column
add_text(s, LM, Inches(1.5), hw, Inches(0.4),
         'V  优势', font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True)
add_hline(s, LM, Inches(2.0), hw, NAVY)
add_multiline(s, LM, Inches(2.2), hw, Inches(2.5),
              ['• 优势要点一', '• 优势要点二', '• 优势要点三'], line_spacing_pt=10)
# Cons column
cx = LM + hw + Inches(0.733)
add_text(s, cx, Inches(1.5), hw, Inches(0.4),
         'X  风险', font_size=SUB_HEADER_SIZE, font_color=DARK_GRAY, bold=True)
add_hline(s, cx, Inches(2.0), hw, DARK_GRAY)
add_multiline(s, cx, Inches(2.2), hw, Inches(2.5),
              ['• 风险要点一', '• 风险要点二', '• 风险要点三'], line_spacing_pt=10)
# Bottom conclusion
add_rect(s, LM, Inches(5.2), CONTENT_W, Inches(1.5), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.3), Inches(1.5), Inches(0.4),
         '结论', font_size=BODY_SIZE, font_color=NAVY, bold=True)
add_text(s, LM + Inches(0.3), Inches(5.8), CONTENT_W - Inches(0.6), Inches(0.6),
         '综合评估建议文字', font_size=BODY_SIZE)
add_source(s, 'Source: ...')
```

#### 22. Traffic Light / RAG Status (红绿灯状态页)

**适用场景**: 多项目/多模块的状态总览（绿=正常、黄=关注、红=风险）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  项目        状态    进度     备注       │
│  ═══════════════════════════════════    │
│  项目A       (G)    85%     按计划推进  │
│  ───────────────────────────────────    │
│  项目B       (Y)    60%     需关注资源  │
│  ───────────────────────────────────    │
│  项目C       (R)    30%     存在阻塞    │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '项目状态总览')
# Header
headers = ['项目', '状态', '进度', '备注']
widths = [Inches(3.0), Inches(1.5), Inches(2.0), Inches(5.233)]
hx = LM
for hdr, w in zip(headers, widths):
    add_text(s, hx, Inches(1.5), w, Inches(0.4), hdr,
             font_size=BODY_SIZE, font_color=MED_GRAY, bold=True)
    hx += w
add_hline(s, LM, Inches(2.0), CONTENT_W, BLACK, Pt(1.0))
# Rows with status indicators
rows = [('产品研发', 'NAVY', '85%', '按计划推进'),
        ('市场推广', 'MED_GRAY', '60%', '需关注预算'),
        ('团队扩招', 'DARK_GRAY', '30%', '存在阻塞')]
color_map = {'NAVY': NAVY, 'MED_GRAY': MED_GRAY, 'DARK_GRAY': DARK_GRAY}
ry = Inches(2.2)
for name, status_color, pct, note in rows:
    add_text(s, LM, ry, Inches(3.0), Inches(0.6), name, font_size=BODY_SIZE)
    add_oval(s, LM + Inches(3.3), ry + Inches(0.05), '', size=Inches(0.35),
             bg=color_map[status_color])
    add_text(s, LM + Inches(4.5), ry, Inches(2.0), Inches(0.6), pct, font_size=BODY_SIZE)
    add_text(s, LM + Inches(6.5), ry, Inches(5.233), Inches(0.6), note, font_size=BODY_SIZE)
    ry += Inches(0.7)
    add_hline(s, LM, ry, CONTENT_W, LINE_GRAY)
    ry += Inches(0.15)
add_source(s, 'Source: ...')
```

#### 23. Scorecard (计分卡页)

**适用场景**: 展示多项评估维度的得分/评级，如供应商评估、团队绩效。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  评估维度          得分   评级           │
│  ═══════════════════════════════════    │
│  客户满意度         92    ████████░░    │
│  产品质量           85    ███████░░░    │
│  交付速度           78    ██████░░░░    │
│  创新能力           65    █████░░░░░    │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '综合评估计分卡')
headers = ['评估维度', '得分', '评级']
add_text(s, LM, Inches(1.5), Inches(4.0), Inches(0.4), headers[0],
         font_size=BODY_SIZE, font_color=MED_GRAY, bold=True)
add_text(s, Inches(5.0), Inches(1.5), Inches(1.5), Inches(0.4), headers[1],
         font_size=BODY_SIZE, font_color=MED_GRAY, bold=True)
add_text(s, Inches(7.0), Inches(1.5), Inches(5.5), Inches(0.4), headers[2],
         font_size=BODY_SIZE, font_color=MED_GRAY, bold=True)
add_hline(s, LM, Inches(2.0), CONTENT_W, BLACK, Pt(1.0))
items = [('客户满意度', '92', 0.92), ('产品质量', '85', 0.85),
         ('交付速度', '78', 0.78), ('创新能力', '65', 0.65)]
ry = Inches(2.2)
bar_max = Inches(5.0)
for name, score, pct in items:
    add_text(s, LM, ry, Inches(4.0), Inches(0.5), name, font_size=BODY_SIZE)
    add_text(s, Inches(5.0), ry, Inches(1.5), Inches(0.5), score,
             font_size=BODY_SIZE, font_color=NAVY, bold=True)
    add_rect(s, Inches(7.0), ry + Inches(0.1), bar_max, Inches(0.3), BG_GRAY)
    add_rect(s, Inches(7.0), ry + Inches(0.1), Inches(5.0 * pct), Inches(0.3), NAVY)
    ry += Inches(0.7)
    add_hline(s, LM, ry, CONTENT_W, LINE_GRAY)
    ry += Inches(0.15)
add_source(s, 'Source: ...')
```

---

### 类别 E：内容叙事

#### 24. Executive Summary (执行摘要页)

**适用场景**: 演示文稿的核心结论汇总，通常放在开头或结尾。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│ ┌──NAVY（核心结论）────────────────────┐ │
│ │  一句话核心结论                       │ │
│ └──────────────────────────────────────┘ │
│                                         │
│  (1) 支撑论点一      详细说明           │
│  (2) 支撑论点二      详细说明           │
│  (3) 支撑论点三      详细说明           │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '执行摘要')
add_rect(s, LM, Inches(1.4), CONTENT_W, Inches(1.0), NAVY)
add_text(s, LM + Inches(0.3), Inches(1.4), CONTENT_W - Inches(0.6), Inches(1.0),
         '核心结论：一句话概括最重要的发现或建议',
         font_size=SUB_HEADER_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
points = [('1', '论点一标题', '支撑论点的详细说明文字'),
          ('2', '论点二标题', '支撑论点的详细说明文字'),
          ('3', '论点三标题', '支撑论点的详细说明文字')]
iy = Inches(2.8)
for num, title, desc in points:
    add_oval(s, LM, iy, num)
    add_text(s, LM + Inches(0.6), iy, Inches(3.5), Inches(0.4),
             title, font_size=BODY_SIZE, font_color=NAVY, bold=True)
    add_text(s, Inches(5.0), iy, Inches(7.5), Inches(0.4),
             desc, font_size=BODY_SIZE)
    iy += Inches(0.6)
    add_hline(s, LM, iy, CONTENT_W, LINE_GRAY)
    iy += Inches(0.3)
add_source(s, 'Source: ...')
```

#### 25. Key Takeaway with Detail (核心洞见页)

**适用场景**: 左侧详细论述 + 右侧灰底要点提炼，用于核心发现页。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│                      ┌──BG_GRAY────────┐│
│  左侧正文内容        │ Key Takeaways   ││
│  详细分析论述        │ 1. 要点一        ││
│  数据与支撑          │ 2. 要点二        ││
│                      │ 3. 要点三        ││
│                      └─────────────────┘│
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '核心发现')
# Left content
add_text(s, LM, Inches(1.5), Inches(7.5), Inches(0.4),
         '分析标题', font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True)
add_hline(s, LM, Inches(2.0), Inches(7.5), LINE_GRAY)
add_multiline(s, LM, Inches(2.2), Inches(7.5), Inches(4.0),
              ['详细分析段落一', '', '详细分析段落二'], line_spacing_pt=8)
# Right takeaway
tk_x = Inches(9.0)
tk_w = Inches(3.5)
add_rect(s, tk_x, Inches(1.5), tk_w, Inches(5.0), BG_GRAY)
add_text(s, tk_x + Inches(0.2), Inches(1.7), tk_w - Inches(0.4), Inches(0.4),
         'Key Takeaways', font_size=BODY_SIZE, font_color=NAVY, bold=True)
add_hline(s, tk_x + Inches(0.2), Inches(2.2), tk_w - Inches(0.4), LINE_GRAY)
add_multiline(s, tk_x + Inches(0.2), Inches(2.4), tk_w - Inches(0.4), Inches(3.8),
              ['1. 要点一', '2. 要点二', '3. 要点三'], line_spacing_pt=10)
add_source(s, 'Source: ...')
```

#### 26. Quote / Insight Page (引言/洞见页)

**适用场景**: 突出一段重要引言、专家观点或核心洞察。

```
┌─────────────────────────────────────────┐
│                                         │
│            ──────────                   │
│                                         │
│      "引言内容，居中显示，               │
│       大字号强调核心观点"                │
│                                         │
│            ──────────                   │
│         — 来源/作者                      │
│                                         │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_rect(s, 0, 0, SW, Inches(0.05), NAVY)
add_hline(s, Inches(5.5), Inches(2.0), Inches(2.3), NAVY, Pt(1.5))
add_text(s, Inches(1.5), Inches(2.5), Inches(10.3), Inches(2.5),
         '"引言内容，用于强调某个核心观点或专家洞见"',
         font_size=Pt(24), font_color=DARK_GRAY, alignment=PP_ALIGN.CENTER)
add_hline(s, Inches(5.5), Inches(5.3), Inches(2.3), NAVY, Pt(1.5))
add_text(s, Inches(1.5), Inches(5.6), Inches(10.3), Inches(0.5),
         '— 作者姓名，来源',
         font_size=BODY_SIZE, font_color=MED_GRAY, alignment=PP_ALIGN.CENTER)
```

#### 27. Two-Column Text (双栏文本页)

**适用场景**: 平衡展示两个主题/方面，每列独立标题+正文。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  (A) 左栏标题         (B) 右栏标题      │
│  ─────────────        ─────────────     │
│  左栏正文内容         右栏正文内容       │
│  详细分析             详细分析           │
│                                         │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '双维度分析')
cw = Inches(5.5)
cg = Inches(0.733)
cols = [('A', '维度一标题', ['分析要点一', '分析要点二', '分析要点三']),
        ('B', '维度二标题', ['分析要点一', '分析要点二', '分析要点三'])]
for i, (letter, title, points) in enumerate(cols):
    cx = LM + (cw + cg) * i
    add_oval(s, cx, Inches(1.5), letter)
    add_text(s, cx + Inches(0.6), Inches(1.5), cw - Inches(0.6), Inches(0.4),
             title, font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True)
    add_hline(s, cx, Inches(2.0), cw, LINE_GRAY)
    add_multiline(s, cx, Inches(2.2), cw, Inches(4.0),
                  [f'• {p}' for p in points], line_spacing_pt=10)
add_source(s, 'Source: ...')
```

#### 28. Four-Column Overview (四栏概览页)

**适用场景**: 四个并列维度的概览（如四大业务线、四个能力域）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  (1)       (2)       (3)       (4)      │
│  标题1     标题2     标题3     标题4     │
│  ────      ────      ────      ────     │
│  描述      描述      描述      描述      │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '四大业务板块')
items = [('1', '板块一', '描述内容\n关键数据'),
         ('2', '板块二', '描述内容\n关键数据'),
         ('3', '板块三', '描述内容\n关键数据'),
         ('4', '板块四', '描述内容\n关键数据')]
cw = Inches(2.7)
cg = (CONTENT_W - cw * 4) / 3
for i, (num, title, desc) in enumerate(items):
    cx = LM + (cw + cg) * i
    add_rect(s, cx, Inches(1.5), cw, Inches(4.8), BG_GRAY)
    add_oval(s, cx + Inches(1.1), Inches(1.65), num)
    add_text(s, cx + Inches(0.15), Inches(2.3), cw - Inches(0.3), Inches(0.4),
             title, font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER)
    add_hline(s, cx + Inches(0.3), Inches(2.8), cw - Inches(0.6), LINE_GRAY)
    add_multiline(s, cx + Inches(0.15), Inches(3.0), cw - Inches(0.3), Inches(3.0),
                  desc.split('\n'), line_spacing_pt=8, alignment=PP_ALIGN.CENTER)
add_source(s, 'Source: ...')
```

---

### 类别 F：时间流程

#### 29. Timeline / Roadmap (时间轴/路线图)

**适用场景**: 展示时间维度的里程碑计划（季度/月度/年度路线图）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│                                         │
│  (1)──────(2)──────(3)──────(4)         │
│  Q1       Q2       Q3       Q4         │
│  里程碑1  里程碑2  里程碑3  里程碑4     │
│                                         │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '2026 年度路线图')
# Timeline bar
add_hline(s, LM + Inches(0.5), Inches(3.0), Inches(10.7), LINE_GRAY, Pt(2))
milestones = [('Q1', '产品 MVP\n发布'), ('Q2', '用户增长\n达到10万'),
              ('Q3', '盈利\n突破'), ('Q4', '国际化\n拓展')]
spacing = Inches(10.7) / (len(milestones) - 1)
for i, (label, desc) in enumerate(milestones):
    mx = LM + Inches(0.5) + spacing * i
    add_oval(s, mx - Inches(0.225), Inches(2.775), str(i + 1))
    add_text(s, mx - Inches(1.0), Inches(2.0), Inches(2.0), Inches(0.5),
             label, font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER)
    add_text(s, mx - Inches(1.0), Inches(3.5), Inches(2.0), Inches(1.5),
             desc, font_size=BODY_SIZE, alignment=PP_ALIGN.CENTER)
add_source(s, 'Source: ...')
```

#### 30. Vertical Steps (垂直步骤页)

**适用场景**: 从上到下的操作步骤或实施阶段。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  (1) 步骤一标题      详细说明           │
│  ─────────────────────────────────────  │
│  (2) 步骤二标题      详细说明           │
│  ─────────────────────────────────────  │
│  (3) 步骤三标题      详细说明           │
│  ─────────────────────────────────────  │
│  (4) 步骤四标题      详细说明           │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '实施步骤')
steps = [('1', '需求分析', '深入调研用户需求与业务痛点'),
         ('2', '方案设计', '制定技术架构与实施计划'),
         ('3', '开发实施', '分阶段迭代交付核心功能'),
         ('4', '上线运营', '监控效果并持续优化')]
iy = Inches(1.5)
for num, title, desc in steps:
    add_oval(s, LM, iy, num)
    add_text(s, LM + Inches(0.6), iy, Inches(3.5), Inches(0.4),
             title, font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True)
    add_text(s, Inches(5.0), iy, Inches(7.5), Inches(0.4),
             desc, font_size=BODY_SIZE)
    iy += Inches(0.6)
    add_hline(s, LM, iy, CONTENT_W, LINE_GRAY)
    iy += Inches(0.5)
add_source(s, 'Source: ...')
```

#### 31. Cycle / Loop (循环图页)

**适用场景**: 闭环流程或迭代循环（如 PDCA、敏捷迭代、反馈循环）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│         ┌──阶段1──┐                     │
│         │        │                      │
│  ┌阶段4┐│        │┌阶段2┐   右侧说明   │
│  │     │└────────┘│     │              │
│  └─────┘          └─────┘              │
│         ┌──阶段3──┐                     │
│         └────────┘                      │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '持续改进循环（PDCA）')
phases = [('Plan\n计划', Inches(2.8), Inches(1.5)),
          ('Do\n执行', Inches(5.0), Inches(3.0)),
          ('Check\n检查', Inches(2.8), Inches(4.5)),
          ('Act\n改进', Inches(0.6), Inches(3.0))]
for i, (label, px, py) in enumerate(phases):
    fill = NAVY if i == 0 else BG_GRAY
    tc = WHITE if i == 0 else NAVY
    add_rect(s, LM + px, py, Inches(2.2), Inches(1.2), fill)
    add_text(s, LM + px + Inches(0.1), py + Inches(0.1), Inches(2.0), Inches(1.0),
             label, font_size=SUB_HEADER_SIZE, font_color=tc, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# Arrows between phases (text arrows)
add_text(s, LM + Inches(4.5), Inches(2.0), Inches(1.0), Inches(0.5),
         '->', font_size=Pt(24), font_color=NAVY, alignment=PP_ALIGN.CENTER)
add_text(s, LM + Inches(5.0), Inches(4.0), Inches(1.0), Inches(0.5),
         'v', font_size=Pt(24), font_color=NAVY, alignment=PP_ALIGN.CENTER)
add_text(s, LM + Inches(2.0), Inches(5.0), Inches(1.0), Inches(0.5),
         '<-', font_size=Pt(24), font_color=NAVY, alignment=PP_ALIGN.CENTER)
add_text(s, LM + Inches(0.8), Inches(2.0), Inches(1.0), Inches(0.5),
         '^', font_size=Pt(24), font_color=NAVY, alignment=PP_ALIGN.CENTER)
# Right side explanation
add_rect(s, Inches(8.5), Inches(1.5), Inches(4.0), Inches(5.0), BG_GRAY)
add_text(s, Inches(8.8), Inches(1.7), Inches(3.4), Inches(0.4),
         '循环要点', font_size=BODY_SIZE, font_color=NAVY, bold=True)
add_multiline(s, Inches(8.8), Inches(2.3), Inches(3.4), Inches(3.5),
              ['每个阶段的说明...'], line_spacing_pt=10)
add_source(s, 'Source: ...')
```

#### 32. Funnel (漏斗图页)

**适用场景**: 转化漏斗（如销售漏斗、用户转化路径）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  ┌════════════════════════════┐  100%   │
│  │         认知               │         │
│  ├══════════════════════┤      60%      │
│  │       兴趣           │               │
│  ├════════════════┤           35%       │
│  │     购买       │                     │
│  ├══════════┤                 15%       │
│  │   留存   │                           │
│  └─────────┘                            │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '用户转化漏斗')
stages = [('认知', '100,000', 1.0), ('兴趣', '60,000', 0.6),
          ('购买', '35,000', 0.35), ('留存', '15,000', 0.15)]
max_w = Inches(8.0)
fy = Inches(1.6)
for i, (name, count, pct) in enumerate(stages):
    w = max_w * pct
    fx = Inches(6.666) - w / 2  # center
    fill = NAVY if i == 0 else BG_GRAY
    tc = WHITE if i == 0 else NAVY
    add_rect(s, fx, fy, w, Inches(1.0), fill)
    add_text(s, fx + Inches(0.2), fy, w - Inches(0.4), Inches(1.0),
             name, font_size=SUB_HEADER_SIZE, font_color=tc, bold=True,
             anchor=MSO_ANCHOR.MIDDLE, alignment=PP_ALIGN.CENTER)
    add_text(s, fx + w + Inches(0.3), fy + Inches(0.2), Inches(2.5), Inches(0.5),
             f'{count} ({int(pct*100)}%)', font_size=BODY_SIZE, font_color=NAVY, bold=True)
    fy += Inches(1.2)
add_source(s, 'Source: ...')
```

---

### 类别 G：团队专题

#### 33. Meet the Team (团队介绍页)

**适用场景**: 团队成员/核心高管/项目组简介。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  ┌─BG──┐    ┌─BG──┐    ┌─BG──┐        │
│  │(头像)│    │(头像)│    │(头像)│        │
│  │ 姓名 │    │ 姓名 │    │ 姓名 │        │
│  │ 职位 │    │ 职位 │    │ 职位 │        │
│  │ 简介 │    │ 简介 │    │ 简介 │        │
│  └──────┘    └──────┘    └──────┘        │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '核心团队')
members = [('张三', 'CEO', '15年行业经验\n前XX公司VP'),
           ('李四', 'CTO', '技术架构专家\n前XX公司总监'),
           ('王五', 'COO', '运营管理专家\n前XX公司负责人')]
cw = Inches(3.5)
cg = (CONTENT_W - cw * 3) / 2
for i, (name, role, bio) in enumerate(members):
    cx = LM + (cw + cg) * i
    add_rect(s, cx, Inches(1.5), cw, Inches(5.0), BG_GRAY)
    add_oval(s, cx + Inches(1.25), Inches(1.7), name[0], size=Inches(1.0))
    add_text(s, cx + Inches(0.15), Inches(2.9), cw - Inches(0.3), Inches(0.4),
             name, font_size=SUB_HEADER_SIZE, font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER)
    add_text(s, cx + Inches(0.15), Inches(3.4), cw - Inches(0.3), Inches(0.4),
             role, font_size=BODY_SIZE, font_color=MED_GRAY, alignment=PP_ALIGN.CENTER)
    add_hline(s, cx + Inches(0.3), Inches(3.9), cw - Inches(0.6), LINE_GRAY)
    add_multiline(s, cx + Inches(0.15), Inches(4.1), cw - Inches(0.3), Inches(2.0),
                  bio.split('\n'), line_spacing_pt=8, alignment=PP_ALIGN.CENTER)
add_source(s, 'Source: ...')
```

#### 34. Case Study (案例研究页)

**适用场景**: 展示成功案例，按"情境-行动-结果"结构组织。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  ┌─Situation──┐ ┌─Approach──┐ ┌Result─┐ │
│  │ 背景/挑战  │ │ 采取行动  │ │ 成果  │ │
│  │            │ │           │ │       │ │
│  └────────────┘ └───────────┘ └───────┘ │
│                                         │
│  ┌──BG_GRAY 客户评价/关键指标──────────┐ │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '案例研究：XX项目')
sections = [('S', 'Situation\n情境', '客户面临的\n挑战描述'),
            ('A', 'Approach\n方法', '我们采取的\n解决方案'),
            ('R', 'Result\n成果', '取得的量化\n成果数据')]
sw = Inches(3.5)
sg = (CONTENT_W - sw * 3) / 2
for i, (letter, title, desc) in enumerate(sections):
    sx = LM + (sw + sg) * i
    fill = NAVY if i == 2 else BG_GRAY
    tc = WHITE if i == 2 else NAVY
    dc = WHITE if i == 2 else DARK_GRAY
    add_rect(s, sx, Inches(1.5), sw, Inches(3.0), fill)
    add_oval(s, sx + Inches(0.15), Inches(1.65), letter,
             bg=WHITE if i == 2 else NAVY, fg=NAVY if i == 2 else WHITE)
    add_text(s, sx + Inches(0.15), Inches(2.2), sw - Inches(0.3), Inches(0.8),
             title, font_size=BODY_SIZE, font_color=tc, bold=True,
             alignment=PP_ALIGN.CENTER)
    add_text(s, sx + Inches(0.15), Inches(3.1), sw - Inches(0.3), Inches(1.0),
             desc, font_size=BODY_SIZE, font_color=dc, alignment=PP_ALIGN.CENTER)
# Bottom highlight
add_rect(s, LM, Inches(5.0), CONTENT_W, Inches(1.5), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.1), Inches(1.5), Inches(0.4),
         '关键成果', font_size=BODY_SIZE, font_color=NAVY, bold=True)
add_text(s, LM + Inches(0.3), Inches(5.6), CONTENT_W - Inches(0.6), Inches(0.6),
         '营收增长 45%  |  客户满意度 92%  |  运营效率提升 30%',
         font_size=BODY_SIZE, font_color=DARK_GRAY)
add_source(s, 'Source: ...')
```

#### 35. Action Items / Next Steps (行动计划页)

**适用场景**: 演示文稿结尾的下一步行动清单。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  ┌──NAVY──┐   ┌──NAVY──┐   ┌──NAVY──┐  │
│  │行动一  │   │行动二  │   │行动三  │  │
│  ├─BG─────┤   ├─BG─────┤   ├─BG─────┤  │
│  │ 时间   │   │ 时间   │   │ 时间   │  │
│  │ 描述   │   │ 描述   │   │ 描述   │  │
│  │ 负责人 │   │ 负责人 │   │ 负责人 │  │
│  └────────┘   └────────┘   └────────┘  │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '下一步行动')
actions = [('建立数据中台', '2026 Q2', '完成核心数据资产盘点\n搭建基础架构', '技术团队'),
           ('启动用户增长计划', '2026 Q3', '渠道拓展+内容营销\n目标新增50万用户', '市场团队'),
           ('优化运营流程', '2026 Q4', '自动化率提升至80%\n降本增效', '运营团队')]
cw = Inches(3.5)
cg = (CONTENT_W - cw * 3) / 2
for i, (title, timeline, desc, owner) in enumerate(actions):
    cx = LM + (cw + cg) * i
    add_rect(s, cx, Inches(1.5), cw, Inches(0.6), NAVY)
    add_text(s, cx + Inches(0.15), Inches(1.5), cw - Inches(0.3), Inches(0.6),
             title, font_size=BODY_SIZE, font_color=WHITE, bold=True,
             anchor=MSO_ANCHOR.MIDDLE, alignment=PP_ALIGN.CENTER)
    add_rect(s, cx, Inches(2.1), cw, Inches(0.4), BG_GRAY)
    add_text(s, cx + Inches(0.15), Inches(2.1), cw - Inches(0.3), Inches(0.4),
             timeline, font_size=BODY_SIZE, font_color=NAVY, bold=True,
             anchor=MSO_ANCHOR.MIDDLE, alignment=PP_ALIGN.CENTER)
    add_multiline(s, cx + Inches(0.15), Inches(2.7), cw - Inches(0.3), Inches(2.0),
                  desc.split('\n'), line_spacing_pt=8, alignment=PP_ALIGN.CENTER)
    add_hline(s, cx + Inches(0.3), Inches(4.9), cw - Inches(0.6), LINE_GRAY)
    add_text(s, cx + Inches(0.15), Inches(5.1), cw - Inches(0.3), Inches(0.4),
             f'负责人：{owner}', font_size=BODY_SIZE, font_color=MED_GRAY,
             alignment=PP_ALIGN.CENTER)
add_source(s, 'Source: ...')
```

#### 36. Closing / Thank You (结束页)

**适用场景**: 演示文稿结尾的致谢或总结收尾页。

```
┌─────────────────────────────────────────┐
│  ═══                                    │
│                                         │
│           核心总结语句                    │
│           ──────────                    │
│           结束寄语                       │
│                                         │
│  ─────                                  │
└─────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(BL)
add_rect(s, 0, 0, SW, Inches(0.05), NAVY)
add_text(s, Inches(1.5), Inches(2.0), Inches(10.3), Inches(1.0),
         '核心总结语句', font_size=Pt(28), font_color=NAVY, bold=True,
         font_name='Georgia', alignment=PP_ALIGN.CENTER)
add_hline(s, Inches(5.5), Inches(3.3), Inches(2.3), NAVY, Pt(1.5))
add_text(s, Inches(1.5), Inches(3.8), Inches(10.3), Inches(2.0),
         '结束寄语或核心思想的延伸表达',
         font_size=SUB_HEADER_SIZE, font_color=DARK_GRAY, alignment=PP_ALIGN.CENTER)
add_hline(s, Inches(1), Inches(6.8), Inches(3), NAVY, Pt(2))
```

---

## Python Code Patterns

### Helper Functions (Copy Directly)

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn


def _clean_shape(shape):
    """Remove p:style from any shape to prevent effect references."""
    sp = shape._element
    style = sp.find(qn('p:style'))
    if style is not None:
        sp.remove(style)


def set_ea_font(run, typeface='KaiTi'):
    """Set East Asian font for Chinese text"""
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', typeface)


def add_text(slide, left, top, width, height, text, font_size=Pt(14),
             font_name='Arial', font_color=RGBColor(0x33, 0x33, 0x33), bold=False,
             alignment=PP_ALIGN.LEFT, ea_font='KaiTi', anchor=MSO_ANCHOR.TOP):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    anchor_map = {MSO_ANCHOR.MIDDLE: 'ctr', MSO_ANCHOR.BOTTOM: 'b', MSO_ANCHOR.TOP: 't'}
    bodyPr.set('anchor', anchor_map.get(anchor, 't'))
    for attr in ['lIns','tIns','rIns','bIns']:
        bodyPr.set(attr, '45720')
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = font_size
    p.font.name = font_name
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.alignment = alignment
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    for run in p.runs:
        set_ea_font(run, ea_font)
    return txBox


def add_multiline(slide, left, top, width, height, lines, font_size=Pt(14),
                  font_name='Arial', font_color=RGBColor(0x33, 0x33, 0x33), bold=False,
                  alignment=PP_ALIGN.LEFT, ea_font='KaiTi', line_spacing_pt=6):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    for attr in ['lIns','tIns','rIns','bIns']:
        bodyPr.set(attr, '45720')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = font_size
        p.font.name = font_name
        p.font.color.rgb = font_color
        p.font.bold = bold
        p.alignment = alignment
        p.space_before = Pt(line_spacing_pt)
        p.space_after = Pt(0)
        for run in p.runs:
            set_ea_font(run, ea_font)
    return txBox


def add_rect(slide, left, top, width, height, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    _clean_shape(shape)  # CRITICAL: remove p:style
    return shape


def add_hline(slide, x, y, length, color=RGBColor(0, 0, 0), thickness=Pt(0.5)):
    """Draw a horizontal line using a thin rectangle (no connector)."""
    h = max(int(thickness), Emu(6350))  # minimum ~0.5pt
    return add_rect(slide, x, y, length, h, color)


def add_oval(slide, x, y, letter, size=Inches(0.45),
             bg=RGBColor(0x05, 0x1C, 0x2C), fg=RGBColor(0xFF, 0xFF, 0xFF)):
    """Add a circle label with a letter (e.g. 'A', '1').
    Uses Arial font to match body text consistency."""
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    c.fill.solid()
    c.fill.fore_color.rgb = bg
    c.line.fill.background()
    tf = c.text_frame
    tf.paragraphs[0].text = letter
    tf.paragraphs[0].font.size = Pt(14)
    tf.paragraphs[0].font.name = 'Arial'
    tf.paragraphs[0].font.color.rgb = fg
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    for run in tf.paragraphs[0].runs:
        set_ea_font(run, 'KaiTi')
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    bodyPr.set('anchor', 'ctr')
    _clean_shape(c)  # CRITICAL: remove p:style
    return c


def add_action_title(slide, text, title_size=Pt(22)):
    """White bg, black text, thin line below."""
    add_text(slide, Inches(0.8), Inches(0.15), Inches(11.7), Inches(0.9),
             text, font_size=title_size, font_color=RGBColor(0, 0, 0), bold=True,
             font_name='Georgia', ea_font='KaiTi', anchor=MSO_ANCHOR.MIDDLE)
    add_hline(slide, Inches(0.8), Inches(1.05), Inches(11.7),
             color=RGBColor(0, 0, 0), thickness=Pt(0.5))


def add_source(slide, text, y=Inches(7.05)):
    add_text(slide, Inches(0.8), y, Inches(11), Inches(0.3),
             text, font_size=Pt(9), font_color=RGBColor(0x66, 0x66, 0x66))
```

---

## Common Issues & Solutions

### Problem 1: PPT Won't Open / "File Needs Repair"

**Cause**: Shapes or connectors carry `<p:style>` with `effectRef idx="2"`, referencing theme effects (shadows/3D)

**Solution** (v1.1 — three-layer defense):
1. **Never use connectors** — use `add_hline()` (thin rectangle) instead of `add_connector()`
2. **Inline cleanup** — every `add_rect()` and `add_oval()` calls `_clean_shape()` to remove `p:style`
3. **Post-save nuclear cleanup** — `full_cleanup()` removes ALL `<p:style>` from every slide XML + theme effects
4. Verify with:
   ```python
   import zipfile
   from lxml import etree
   zf = zipfile.ZipFile('file.pptx', 'r')
   for name in zf.namelist():
       if name.endswith('.xml'):
           root = etree.fromstring(zf.read(name))
           ns_p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
           styles = root.findall(f'.//{{{ns_p}}}style')
           if styles:
               print(f"ERROR: {name} has {len(styles)} p:style elements!")
   ```

### Problem 2: Text Not Displaying Correctly in PowerPoint

**Cause**: Chinese characters rendered as English font instead of KaiTi

**Solution**:
- Use `set_ea_font(run, 'KaiTi')` in every paragraph with Chinese text
- Call it inside the loop that creates runs:
  ```python
  for run in p.runs:
      set_ea_font(run, 'KaiTi')
  ```

### Problem 3: Lines Appearing With Shadows

**Cause**: Connectors inherently carry theme effect references

**Solution** (v1.1):
1. **Do NOT use connectors at all** — use `add_hline()` which draws lines as thin rectangles
2. Run `full_cleanup()` after save to remove any residual `p:style`
3. Verify zero `p:style` and zero shadow elements in the final file

### Problem 4: Font Sizes Inconsistent Across Slides

**Cause**: Using custom sizes instead of defined hierarchy

**Solution**:
- Define constants:
  ```python
  TITLE_SIZE = Pt(22)
  BODY_SIZE = Pt(14)
  SUB_HEADER_SIZE = Pt(18)
  LABEL_SIZE = Pt(14)
  SMALL_SIZE = Pt(9)
  ```
- Use these constants everywhere
- Never use arbitrary sizes like `Pt(13)` or `Pt(15)`

### Problem 5: Columns/Lists Not Aligning Vertically

**Cause**: Mixing different line spacing or not accounting for text height

**Solution**:
- Use consistent `line_spacing_pt` in `add_multiline()` calls
- Calculate row heights in tables based on actual text size:
  - For 14pt text with spacing: use 1.0" height minimum
  - For lists with bullets: use 0.35" height per line + 8pt spacing
- Test by saving and opening in PowerPoint to verify alignment

---

## Refining Existing Presentations

### User Feedback Implementation Checklist

When users provide feedback, follow this checklist:

- [ ] **Font size changes** → Update `TITLE_SIZE`, `BODY_SIZE`, etc. constants and regenerate
- [ ] **Color changes** → Update RGB tuples in color palette section
- [ ] **Line width changes** → Update calls to `add_hline()` with new `thickness=Pt(X)`
- [ ] **Layout spacing changes** → Adjust `Inches(X)` values in coordinate calculations
- [ ] **Circle color changes** → Update `CIRCLE_COLOR` constant and regenerate all circles
- [ ] **Shadow/effect issues** → Run theme cleanup immediately after save
- [ ] **Text wrapping issues** → Increase textbox height or reduce font size incrementally

### Iterative Improvement Process

1. **Collect feedback** - Note specific issues (font size, colors, spacing)
2. **Make changes to code** - Update constants or helper functions
3. **Regenerate slides** - Re-run python-pptx generation
4. **Post-process** - Always run theme cleanup
5. **Test in PowerPoint** - Open and verify visually
6. **Document changes** - Update this section with refinements

---

## Edge Cases

### Handling Large Presentations (20+ Slides)

- Break generation into batches of 5-8 slides, saving and verifying after each batch
- Always call `full_cleanup()` once at the end, not per-batch
- Memory: python-pptx holds the entire presentation in memory; for 50+ slides, monitor usage

### Font Availability

- **KaiTi / SimSun** may not be installed on non-Chinese systems — the presentation will render but fall back to a default CJK font
- **Georgia** is available on Windows/macOS by default; on Linux, install `ttf-mscorefonts-installer`
- If target audience uses macOS only, consider using `PingFang SC` as `ea_font` fallback

### Slide Dimensions

- All layouts assume **13.333" × 7.5"** (widescreen 16:9). Using 4:3 or custom sizes will break coordinate calculations
- If custom dimensions are required, scale all `Inches()` values proportionally

### PowerPoint vs LibreOffice

- Generated files are optimized for **Microsoft PowerPoint** (Windows/macOS)
- LibreOffice Impress may render fonts and spacing slightly differently
- `full_cleanup()` is still recommended for LibreOffice compatibility

---

## Error Handling

### "File needs repair" on open

```
Cause: Residual <p:style> elements referencing theme effects
Fix:   Ensure full_cleanup() runs AFTER prs.save()
Verify: Zero <p:style> elements in all slide XML (see validation code in Common Issues)
```

### Chinese text shows as boxes/squares

```
Cause: Missing set_ea_font() call on text runs containing CJK characters
Fix:   Every paragraph with Chinese text must call set_ea_font(run, 'KaiTi') on all runs
```

### ModuleNotFoundError: python-pptx / lxml

```
Fix: pip install python-pptx lxml
Note: lxml requires a C compiler on some Linux distros — install libxml2-dev libxslt-dev first
```

### Shapes overlapping or misaligned

```
Cause: Incorrect Inches() coordinate calculations when adapting templates
Fix:   Use the provided constants (LM, CONTENT_W, etc.) and verify visually after every 2-3 slides
```

---

## Best Practices

1. **Always start from scratch** - Don't try to edit existing .pptx files with python-pptx; regenerate
2. **Test early** - Save and open in PowerPoint after every 2-3 slides to catch issues
3. **Use constants** - Define all colors, sizes, positions as named constants at the top
4. **Keep code DRY** - Use helper functions like `add_text()`, `add_hline()`, `add_oval()`, etc.
5. **Never use connectors** - Always draw lines as thin rectangles via `add_hline()`
6. **Validate XML** - After `full_cleanup()`, verify zero `p:style` and zero shadows remain
6. **Document decisions** - Comment code explaining why specific colors/sizes are chosen
7. **Version control** - Save Python generation script alongside .pptx output

---

## Dependencies

- **python-pptx** >= 0.6.21 - For PowerPoint generation
- **lxml** - For XML processing during theme cleanup
- **zipfile** (built-in) - For PPTX manipulation
- Python 3.8+

Install with:
```bash
pip install python-pptx lxml
```

---

## Example: Complete Minimal Presentation

See `scripts/minimal_example.py` for a complete, working example that generates:
- Cover slide
- Table of contents
- Content slide with title + body text
- Source attribution
- Proper theme cleanup

---

## File References

Generated presentations are typically saved to:
```
./output/presentation.pptx
```

All colors, fonts, and dimensions referenced in code should match this document exactly.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.3.0 | 2026-03-04 | ClawHub release: optimized description for discoverability, added metadata/homepage, added Edge Cases & Error Handling sections |
| 1.2.0 | 2026-03-04 | Fixed circle shape number font inconsistency; `add_oval()` now sets `font_name='Arial'` + `set_ea_font()` for consistent typography |
| | | - Circle numbers simplified: use `1, 2, 3` instead of `01, 02, 03` |
| | | - Removed product-specific references from skill description |
| 1.1.0 | 2026-03-03 | **Breaking**: Replaced connector-based lines with rectangle-based `add_hline()` |
| | | - `add_line()` deprecated, use `add_hline()` instead |
| | | - `add_circle_label()` renamed to `add_oval()` with bg/fg params |
| | | - `add_rect()` now auto-removes `p:style` via `_clean_shape()` |
| | | - `cleanup_theme()` upgraded to `full_cleanup()` (sanitizes all slide XML) |
| | | - Three-layer defense against file corruption |
| | | - `add_multiline()` bullet param removed; use `'\u2022 '` prefix in text |
| 1.0.0 | 2026-03-02 | Initial complete specification, all refinements documented |
| | | - Color palette finalized (NAVY primary) |
| | | - Typography hierarchy locked (22pt title, 14pt body) |
| | | - Line treatment standardized (no shadows) |
| | | - Theme cleanup process documented |
| | | - All helper functions optimized |

