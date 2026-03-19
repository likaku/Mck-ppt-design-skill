---
name: mck-ppt-design
description: >-
  Create professional, consultant-grade PowerPoint presentations from scratch
  using python-pptx with McKinsey-style design. Use when user asks to create
  slides, pitch decks, business presentations, strategy decks, quarterly
  reviews, board meeting slides, or any professional PPTX. Generates clean,
  flat-design presentations with 70 layout patterns across 12 categories
  (structure, data, framework, comparison, narrative, timeline, team, charts,
  images, advanced viz, dashboards, visual storytelling), consistent
  typography, zero file-corruption issues, BLOCK_ARC native shapes for
  circular charts (donut, pie, gauge), and production-hardened guard rails
  for spacing, overflow, legend consistency, title style uniformity,
  dynamic sizing for variable-count layouts, and chart rendering.
---

# McKinsey PPT Design Framework

> **Version**: 1.10.4 · **License**: Apache-2.0 · **Author**: [likaku](https://github.com/likaku/Mck-ppt-design-skill)
>
> **Required tools**: Read, Write, Bash · **Requires**: python3, pip

## Overview

This skill encodes the complete design specification for **professional business presentations** — a consultant-grade PowerPoint framework based on McKinsey design principles. It includes:

- **70 layout patterns** across 12 categories (structure, data, framework, comparison, narrative, timeline, team, charts, **images**, **advanced viz**, **dashboards**, **visual storytelling**)
- **Color system** and strict typography hierarchy
- **Python-pptx code patterns** ready to copy and customize
- **Three-layer defense** against file corruption (zero `p:style` leaks)
- **Chinese + English font handling** (KaiTi / Georgia / Arial)
- **Image placeholder system** for image-containing layouts (v1.8)
- **BLOCK_ARC native shapes for charts** — donut, pie, gauge rendered with 3-4 shapes instead of hundreds of blocks, 60-80% smaller files (v2.0)
- **Production Guard Rails** — 9 mandatory rules including spacing/overflow protection, legend color consistency, title style uniformity, axis label centering, dynamic sizing, BLOCK_ARC chart rendering (v1.9+v2.0)
- **Code Efficiency guidelines** — variable reuse patterns, constant extraction, loop optimization for faster generation (v1.9)

All specifications have been refined through iterative production feedback to ensure visual consistency, professional polish, and zero-defect output.

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

#### Accent Colors (for multi-item differentiation)

When a slide contains **3 or more parallel items** (e.g., comparison cards, pillar frameworks, multi-category overviews), use these accent colors to create visual distinction between items. Without accent colors, parallel items become visually indistinguishable.

| Accent Name | Hex | RGB | Paired Light BG | Usage |
|-------------|-----|-----|-----------------|-------|
| **ACCENT_BLUE** | #006BA6 | (0, 107, 166) | #E3F2FD | First item accent |
| **ACCENT_GREEN** | #007A53 | (0, 122, 83) | #E8F5E9 | Second item accent |
| **ACCENT_ORANGE** | #D46A00 | (212, 106, 0) | #FFF3E0 | Third item accent |
| **ACCENT_RED** | #C62828 | (198, 40, 40) | #FFEBEE | Fourth item / warning |

**Accent Color Rules**:
- Use accent colors for: **card top accent borders** (thin 0.06" rect), **circle labels** (`add_oval()` bg param), **section sub-headers** (font_color)
- Use paired light BG for: **card background fills** only
- Body text inside cards ALWAYS remains **DARK_GRAY (#333333)**
- NAVY remains the primary color for **single-focus** elements (one card, one stat, cover title)
- Use accent colors **ONLY** when the slide has 3+ parallel items that need visual distinction
- The fourth item (D) can use NAVY instead of ACCENT_RED if red feels inappropriate for the content

```python
# Accent color constants
ACCENT_BLUE   = RGBColor(0x00, 0x6B, 0xA6)
ACCENT_GREEN  = RGBColor(0x00, 0x7A, 0x53)
ACCENT_ORANGE = RGBColor(0xD4, 0x6A, 0x00)
ACCENT_RED    = RGBColor(0xC6, 0x28, 0x28)
LIGHT_BLUE    = RGBColor(0xE3, 0xF2, 0xFD)
LIGHT_GREEN   = RGBColor(0xE8, 0xF5, 0xE9)
LIGHT_ORANGE  = RGBColor(0xFF, 0xF3, 0xE0)
LIGHT_RED     = RGBColor(0xFF, 0xEB, 0xEE)
```

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

**IMPORTANT**: Never use `add_connector()` — it causes file corruption. Always use `add_hline()` (thin rectangle).

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

## Presentation Planning

This section provides **mandatory guidance** for planning presentation structure, selecting layouts, and ensuring adequate content density. These rules dramatically improve output quality across different LLM models.

### Recommended Slide Structures

When creating a presentation, follow these templates unless the user explicitly specifies a different structure:

#### Standard Presentation (10-12 slides)

```
 Slide 1:  Cover Slide (Pattern #1 or #4)
 Slide 2:  Table of Contents (Pattern #6) — list ALL content sections
 Slide 3:  Executive Summary / Core Thesis (Pattern #24 or #8+#10)
 Slides 4-7:  Supporting Arguments (one per slide, vary layouts)
 Slides 8-10: Case Studies / Evidence (Pattern #33 or #19)
 Slide 11: Synthesis / Roadmap (Pattern #29 or #16)
 Slide 12: Key Takeaways + Closing (Pattern #34 or #36)
```

#### Short Presentation (6-8 slides)

```
 Slide 1:  Cover Slide
 Slide 2:  Executive Summary (Pattern #24)
 Slides 3-5: Core Content (vary layouts: #8, #14, #19, #33)
 Slide 6:  Synthesis / Timeline (Pattern #29)
 Slide 7:  Key Takeaways (Pattern #34)
 Slide 8:  Closing (Pattern #36)
```

**CRITICAL RULES**:
- **Minimum slide count**: 8 slides for any substantive topic. If the user's content supports 10+, generate 10+.
- **Never stop early**: Generate ALL planned slides in a single script. Do not truncate.
- **TOC must list ALL sections**: The Table of Contents slide must enumerate every content slide by number and title.

### Layout Diversity Requirement

**Each content slide MUST use a DIFFERENT layout pattern from its neighbors.** Repeating the same layout on consecutive slides makes the presentation feel monotonous and unprofessional.

Match content type to the optimal layout pattern:

| Content Type | Recommended Layouts | Avoid |
|---|---|---|
| Single key statistic | Big Number (#8) | Plain text |
| 2 options comparison | Side-by-Side (#19), Before/After (#20), Metric Comparison Row (#62) | Two-column text |
| 3-4 parallel concepts | Three-Pillar (#14), Four-Column (#27), Metric Cards (#10), Icon Grid (#63) | Bullet list |
| Process / steps | Process Chevron (#16), Vertical Steps (#30), Value Chain (#67) | Numbered text |
| Timeline | Timeline/Roadmap (#29), Cycle (#31) | Bullet list |
| Data table | Data Table (#9), Scorecard (#22), Harvey Ball Table (#56) | Plain text |
| Case study | Case Study (#33), Case Study with Image (#45) | Two-column text |
| Summary / conclusion | Executive Summary (#24), Key Takeaway (#25) | Bullet list |
| Multiple KPIs | Three-Stat Dashboard (#12), Two-Stat Comparison (#11), KPI Tracker (#52), Dashboard (#57) | Plain text |
| **Time series + values/percentages** | **Grouped Bar (#37), Stacked Bar (#38), Line Chart (#50), Stacked Area (#70)** | **Data Table, Scorecard** |
| **Category ranking / comparison** | **Horizontal Bar (#39), Grouped Bar (#37), Pareto (#51)** | **Bullet list, Plain text** |
| **Part-of-whole / composition** | **Donut (#48), Pie (#64), Stacked Bar (#38)** | **Bullet list** |
| **Content with visual / photo** | **Content+Right Image (#40), Left Image+Content (#41), Three Images (#42)** | **Text-only layouts** |
| **Risk / evaluation matrix** | **Risk Matrix (#54), SWOT (#65), Harvey Ball (#56), 2x2 Matrix (#13)** | **Bullet list** |
| **Strategic recommendations** | **Numbered List+Panel (#69), Decision Tree (#60), Checklist (#61)** | **Two-column text** |
| **Multi-KPI executive dashboard** | **Dashboard KPI+Chart (#57), Dashboard Table+Chart (#58)** | **Simple table** |
| **Stakeholder / relationship** | **Stakeholder Map (#59)** | **Bullet list** |
| **Meeting agenda** | **Agenda (#66)** | **Plain text** |

**NEVER** use Two-Column Text (#26) for more than 1 slide per deck. It is the least visually engaging layout.

**CHART PRIORITY RULE**: When data contains dates/periods + numeric values or percentages (e.g., `3/4 正面 20% 中性 80%` or `Q1: ¥850万`), you **MUST** use a Chart pattern (#37-#39, #48-#56, #64, #70) instead of a text-based layout. Charts maximize data-ink ratio and are the most visually compelling way to present time-series data.

**IMAGE PRIORITY RULE** (v1.8): When the content involves case studies, product showcases, location overviews, or any scenario where a visual/photo would strengthen the narrative, prefer Image+Content layouts (#40-#47, #68) over text-only layouts. The `add_image_placeholder()` function creates gray placeholder boxes that users replace with real images after generation.

### Content Density Requirements

"Minimalism" in McKinsey design means **removing decorative noise** (shadows, gradients, clip-art), NOT removing content. A slide with 80% whitespace is not minimalist — it is EMPTY.

**Mandatory minimums per content slide**:

1. **At least 3 distinct visual blocks** — e.g., title bar + content area + takeaway box, or title + left panel + right panel
2. **Body text area utilization ≥ 50%** of the available content space (between title bar at 1.4" and source line at 7.05")
3. **Action Title must be a FULL SENTENCE** expressing the slide's key insight:
   - ✅ `"连接组约束的AI模型将自由参数减少90%，实现单细胞精度预测"`
   - ❌ `"连接组约束的AI模型"`
4. **Use specific data points** when the user provides them (numbers, percentages, names) — display them prominently with Big Number or Metric Card patterns
5. **Source attribution** (`add_source()`) on every content slide with specific references, not generic labels

### Production Guard Rails (v1.9 / v2.0)

These rules address **recurring production defects** observed across multiple presentation generations. Each rule is derived from real-world user feedback and must be followed without exception.

#### Rule 1: Spacing Between Content Blocks and Bottom Bars

**Problem observed**: Tables, charts, or content grids placed immediately above a bottom summary/action bar (e.g., "行动公式", "趋势判读", "风险提示") with zero vertical gap, making them visually merged.

**MANDATORY**: There MUST be **at least 0.15" vertical gap** between the last content block and any bottom bar/summary box. Calculate positions explicitly:

```python
# ❌ WRONG: content ends at Inches(6.15), bottom bar also at Inches(6.15)
last_content_bottom = content_top + num_rows * row_height
bar_y = last_content_bottom  # NO GAP!

# ✅ CORRECT: explicit gap
BOTTOM_BAR_GAP = Inches(0.2)
bar_y = last_content_bottom + BOTTOM_BAR_GAP
```

**Validation formula**: `bottom_bar_y >= last_content_bottom + Inches(0.15)`

#### Rule 2: Content Overflow Protection

**Problem observed**: Text or shapes extending beyond the right margin (left_margin + content_width) or bottom margin (source line at 7.05").

**MANDATORY** overflow checks:

1. **Right margin**: Every element's `left + width ≤ LM + CW` (i.e., `Inches(0.8) + Inches(11.733) = Inches(12.533)`)
2. **Bottom margin**: Every element's `top + height ≤ Inches(6.95)` (leaving room for source line at 7.05")
3. **Text in bounded boxes**: When placing text inside a colored `add_rect()` box, the text box MUST be **inset by at least 0.15"** on each side:

```python
# ✅ CORRECT: text inset within its container box
box_left = LM
box_width = CW
add_rect(s, box_left, box_y, box_width, box_h, BG_GRAY)
add_text(s, box_left + Inches(0.3), box_y, box_width - Inches(0.6), box_h,
         text, ...)  # 0.3" padding on each side
```

4. **Multi-column layouts**: When calculating column widths, account for inter-column gaps AND the right margin:
   ```python
   # total available = CW = Inches(11.733)
   num_cols = 3
   gap = Inches(0.2)
   col_w = (CW - gap * (num_cols - 1)) / num_cols  # NOT CW / num_cols
   ```

5. **Long text truncation**: If generated text may exceed box boundaries, reduce `font_size` by 1-2pt or abbreviate text. Never allow visible overflow.

#### Rule 3: Bottom Whitespace Elimination

**Problem observed**: Charts or content areas end at ~Inches(5.5) while the bottom bar sits at ~Inches(6.3), leaving ~0.8" of dead whitespace.

**MANDATORY**: The bottom summary bar should be positioned at **no higher than Inches(6.1)** and **no lower than Inches(6.4)**. Adjust chart/content heights to fill available space. Target: visible whitespace between content and bottom bar ≤ 0.3".

```python
# ✅ CORRECT: Compute bottom bar position dynamically
content_bottom = chart_top + chart_height
# Place bottom bar close to content (but with minimum gap)
bar_y = max(content_bottom + Inches(0.15), Inches(6.1))
bar_y = min(bar_y, Inches(6.4))  # don't push past safe zone
```

#### Rule 4: Legend Color Consistency

**Problem observed**: Chart legends using plain black text "■" symbols (`■ 基准值 ■ 增加 ■ 减少`) while actual chart bars use NAVY, ACCENT_RED, ACCENT_GREEN — colors don't match.

**MANDATORY**: Every legend indicator MUST use a **colored square** (`add_rect()`) matching the exact color used in the chart below it. Never use text-only legends with "■" character.

```python
# ❌ WRONG: Text-only legend with black squares
add_text(s, LM, legend_y, CW, Inches(0.25),
         '■ 基准值  ■ 增加  ■ 减少', ...)

# ✅ CORRECT: Color-matched legend squares
lgx = LM + Inches(5)
add_rect(s, lgx, legend_y, Inches(0.15), Inches(0.15), NAVY)
add_text(s, lgx + Inches(0.2), legend_y, Inches(0.9), Inches(0.25),
         '基准值', font_size=Pt(10), font_color=MED_GRAY)
add_rect(s, lgx + Inches(1.3), legend_y, Inches(0.15), Inches(0.15), ACCENT_RED)
add_text(s, lgx + Inches(1.5), legend_y, Inches(0.9), Inches(0.25),
         '增加', font_size=Pt(10), font_color=MED_GRAY)
# ... repeat for each series
```

**Legend placement**: Inline with or immediately below the chart subtitle line (typically at Inches(1.15)-Inches(1.20)). Legend squares are 0.15" × 0.15" with 0.05" gap to label text.

#### Rule 5: Title Style Consistency

**Problem observed**: Some slides using `add_navy_title_bar()` (full-width navy background + white text) while others use `add_action_title()` (white background + black text + underline), creating jarring visual inconsistency.

**MANDATORY**: Use **`add_action_title()`** (`aat()`) as the **ONLY** title style for ALL content slides. The navy title bar (`antb()`) is **DEPRECATED for content slides** and should only appear if explicitly requested by the user.

```python
# ❌ DEPRECATED: Navy background title bar
def add_navy_title_bar(slide, text):
    add_rect(s, 0, 0, SW, Inches(0.75), NAVY)
    add_text(s, LM, 0, CW, Inches(0.75), text, font_color=WHITE, ...)

# ✅ CORRECT: Consistent white-background action title (bottom-anchored)
def add_action_title(slide, text, title_size=Pt(22)):
    add_text(s, Inches(0.8), Inches(0.15), Inches(11.7), Inches(0.9), text,
             font_size=title_size, font_color=BLACK, bold=True, font_name='Georgia',
             anchor=MSO_ANCHOR.BOTTOM)  # BOTTOM: text sits flush against separator
    add_hline(s, Inches(0.8), Inches(1.05), Inches(11.7), BLACK, Pt(0.5))
```

**Note**: When `add_action_title()` is used, content starts at **Inches(1.25)** (not Inches(1.0)). Account for this when positioning grids, tables, or charts below the title.

#### Rule 6: Axis Label Centering in Matrix/Grid Charts

**Problem observed**: In 2×2 matrix layouts (#13, #59, #65), axis labels ("用户规模↑", "技术壁垒→") positioned at fixed offsets rather than centered on their respective axes, causing visual misalignment.

**MANDATORY**: Axis labels MUST be **centered on the full span** of their axis:

```python
# Grid dimensions
grid_left = LM + Inches(2.0)
grid_top = Inches(1.65)
cell_w = Inches(4.5)  # width of each quadrant
cell_h = Inches(2.0)  # height of each quadrant
grid_w = 2 * cell_w   # full grid width
grid_h = 2 * cell_h   # full grid height

# ✅ CORRECT: Y-axis label centered vertically on FULL grid height
add_text(s, LM, grid_top, Inches(1.8), grid_h,
         'Y轴标签↑', alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ✅ CORRECT: X-axis label centered horizontally on FULL grid width
add_text(s, grid_left, grid_top + grid_h + Inches(0.1), grid_w, Inches(0.3),
         'X轴标签 →', alignment=PP_ALIGN.CENTER)
```

#### Rule 7: Image Placeholder Slide Requirement

**Problem observed**: Presentations generated with zero image-containing slides, resulting in a wall of text/charts that feels monotonous and lacks visual relief.

**MANDATORY**: For presentations with **8+ slides**, at least **1 slide** must include image placeholders (using `add_image_placeholder()` or custom gray boxes with "请插入图片" labels). Preferred positions:

- After the first 2-3 content slides (as a visual break)
- For case studies, product showcases, or ecosystem overviews

**Standard placeholder style** (when not using `add_image_placeholder()` helper):

```python
# Large placeholder
img_l = LM; img_t = Inches(1.3); img_w = Inches(6.5); img_h = Inches(4.0)
add_rect(s, img_l, img_t, img_w, img_h, BG_GRAY)
add_rect(s, img_l + Inches(0.04), img_t + Inches(0.04),
         img_w - Inches(0.08), img_h - Inches(0.08), WHITE)
add_rect(s, img_l + Inches(0.08), img_t + Inches(0.08),
         img_w - Inches(0.16), img_h - Inches(0.16), RGBColor(0xF8, 0xF8, 0xF8))
add_text(s, img_l, img_t + img_h // 2 - Inches(0.3), img_w, Inches(0.5),
         '[ 请插入图片 ]', font_size=Pt(22), font_color=LINE_GRAY,
         bold=True, alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, img_l, img_t + img_h // 2 + Inches(0.2), img_w, Inches(0.3),
         '图片描述标签', font_size=Pt(13), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
```

This triple-border style (BG_GRAY → WHITE → #F8F8F8) creates a professional, clearly identifiable placeholder that prompts users to insert real images.

#### Rule 8: Dynamic Sizing for Variable-Count Layouts (v1.10.4)

**Problem observed**: Layouts with a variable number of items (checklist rows, value chain stages, cover multi-line titles) use **fixed dimensions** that only work for a specific count. When item count differs, content either overflows past page boundaries or leaves excessive whitespace.

**MANDATORY**: For any layout where the number of items is variable, compute dimensions dynamically:

```python
# ✅ Horizontal items (value chain, flow): fill content width
n = len(items)
gap = Inches(0.35)
item_w = (CW - gap * (n - 1)) / n   # NOT a fixed Inches(2.0)

# ✅ Vertical items (checklist, table rows): fit within available height
bottom_limit = BOTTOM_BAR_Y if bottom_bar else SOURCE_Y - Inches(0.05)
available_h = bottom_limit - content_start_y
item_h = min(MAX_ITEM_H, available_h / max(n, 1))  # cap at comfortable max

# ✅ Multi-line titles: height scales with line count
n_lines = text.count('\n') + 1
title_h = Inches(0.8 + 0.65 * max(n_lines - 1, 0))
# Position following elements relative to title bottom, NOT at fixed y
```

**Anti-patterns** (❌ NEVER DO):
- `stage_w = Inches(2.0)` for N stages → use `(CW - gap*(N-1)) / N`
- `row_h = Inches(0.55)` for N rows → use `min(0.85, available / N)`
- `subtitle_y = Inches(3.5)` on cover → use `title_y + title_h + Inches(0.3)`

#### Rule 9: BLOCK_ARC Native Shapes for Circular Charts (v2.0)

**Problem observed**: Donut charts (#48), pie charts (#64), and gauge dials (#55) rendered with hundreds to thousands of small `add_rect()` blocks. This creates 100-2800 shapes per chart, inflates file size by 60-80%, slows generation to 2+ minutes, and produces visual artifacts (gaps between blocks, jagged edges).

**MANDATORY**: Use **BLOCK_ARC** preset shapes via `python-pptx` + XML adjustment for all circular/arc charts. Each segment = 1 shape (total: 3-5 shapes per chart vs. hundreds).

**BLOCK_ARC angle convention** (PPT coordinate system):
- Angles measured **clockwise from 12 o'clock** (top), in **60000ths of a degree**
- Top = 0°, Right = 90°, Bottom = 180°, Left = 270°
- Example: a full-circle donut segment from 12 o'clock CW to 3 o'clock = adj1=0, adj2=5400000

**Three adj parameters**:
- `adj1`: start angle (60000ths of degree, CW from top)
- `adj2`: end angle (60000ths of degree, CW from top)
- `adj3`: inner radius ratio (0 = solid sector / pie, 50000 = max / invisible ring)

```python
from pptx.oxml.ns import qn

def add_block_arc(slide, left, top, width, height, start_deg, end_deg, inner_ratio, color):
    """Draw a BLOCK_ARC shape with precise angle and ring-width control.

    Args:
        slide: pptx slide object
        left, top, width, height: bounding box (width == height for circular arc)
        start_deg: start angle in degrees, CW from 12 o'clock (0=top, 90=right, 180=bottom, 270=left)
        end_deg: end angle in degrees, CW from 12 o'clock
        inner_ratio: 0 = solid pie sector, 50000 = max (paper-thin ring).
                     For ~10px ring width: int((outer_r - Pt(10)) / outer_r * 50000)
        color: RGBColor fill color
    """
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    _clean_shape(sh)  # remove p:style to prevent file corruption

    sp = sh._element.find(qn('p:spPr'))
    prstGeom = sp.find(qn('a:prstGeom'))
    if prstGeom is not None:
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            avLst = prstGeom.makeelement(qn('a:avLst'), {})
            prstGeom.append(avLst)
        for gd in avLst.findall(qn('a:gd')):
            avLst.remove(gd)
        gd1 = avLst.makeelement(qn('a:gd'), {'name': 'adj1', 'fmla': f'val {int(start_deg * 60000)}'})
        gd2 = avLst.makeelement(qn('a:gd'), {'name': 'adj2', 'fmla': f'val {int(end_deg * 60000)}'})
        gd3 = avLst.makeelement(qn('a:gd'), {'name': 'adj3', 'fmla': f'val {inner_ratio}'})
        avLst.append(gd1)
        avLst.append(gd2)
        avLst.append(gd3)
    return sh
```

**Usage patterns**:

```python
# ── Donut chart: 4 segments, ~10px ring width ──
outer_r = Inches(1.6)
inner_ratio = int((outer_r - Pt(10)) / outer_r * 50000)  # ~10px ring
cum_deg = 0  # start at top (0° = 12 o'clock)
for pct, color, label in segments:
    sweep = pct * 360
    add_block_arc(s, cx - outer_r, cy - outer_r, outer_r * 2, outer_r * 2,
                  cum_deg, cum_deg + sweep, inner_ratio, color)
    cum_deg += sweep

# ── Pie chart (solid sectors): inner_ratio = 0 ──
add_block_arc(s, cx - r, cy - r, r * 2, r * 2, 0, 151.2, 0, NAVY)  # 42%

# ── Horizontal rainbow gauge (semi-circle, left→top→right) ──
# PPT coords: left=270°, top=0°, right=90°
gauge_segs = [(0.40, ACCENT_RED), (0.30, ACCENT_ORANGE), (0.30, ACCENT_GREEN)]
inner_ratio = int((outer_r - Pt(10)) / outer_r * 50000)
ppt_cum = 270  # start at left
for pct, color in gauge_segs:
    sweep = pct * 180
    add_block_arc(s, cx - outer_r, cy - outer_r, outer_r * 2, outer_r * 2,
                  ppt_cum % 360, (ppt_cum + sweep) % 360, inner_ratio, color)
    ppt_cum += sweep
```

**Anti-patterns** (❌ NEVER DO for circular charts):
- Nested `for deg in range(...): for r in range(...): add_rect(...)` — generates hundreds/thousands of tiny squares
- Drawing a white circle on top of a filled circle to "fake" a donut — fragile, misaligns on resize
- Using `math.cos/sin` + `add_rect()` loops for arcs — always use `BLOCK_ARC` instead

### Mandatory Slide Elements

EVERY content slide (except Cover and Closing) MUST include ALL of these:

| Element | Function | Position |
|---------|----------|----------|
| Action Title | `add_action_title(slide, text)` | Top (0.15" from top) |
| Title separator line | Included in `add_action_title()` | 1.05" from top |
| Content area | Layout-specific content blocks | 1.4" to 6.5" |
| Source attribution | `add_source(slide, text)` | 7.05" from top |
| Page number | `add_page_number(slide, n, total)` | Bottom-right corner |

Page number helper function:
```python
def add_page_number(slide, num, total):
    add_text(slide, Inches(12.2), Inches(7.1), Inches(1), Inches(0.3),
             f"{num}/{total}", font_size=Pt(9), font_color=MED_GRAY,
             alignment=PP_ALIGN.RIGHT)
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
- Main title (44pt, Georgia, navy) at y=1.2" — **height computed dynamically from line count**
- Subtitle (24pt, dark gray) positioned **below title dynamically**
- Date/info (14pt, med gray) follows subtitle
- Decorative navy line at x=1", y=6.8" (4" wide, 2pt)

Code template:
```python
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(s1, 0, 0, prs.slide_width, Inches(0.05), NAVY)

# Dynamic title height based on line count
lines = title.split('\n') if isinstance(title, str) else title
n_lines = len(lines) if isinstance(lines, list) else title.count('\n') + 1
title_h = Inches(0.8 + 0.65 * max(n_lines - 1, 0))

add_text(s1, Inches(1), Inches(1.2), Inches(11), title_h,
         '项目名称', font_size=Pt(44), font_name='Georgia',
         font_color=NAVY, bold=True, ea_font='KaiTi')

# Position elements BELOW title dynamically — never use fixed y
sub_y = Inches(1.2) + title_h + Inches(0.3)
add_text(s1, Inches(1), sub_y, Inches(11), Inches(0.8),
         '副标题描述', font_size=Pt(24),
         font_color=DARK_GRAY, ea_font='KaiTi')
sub_y += Inches(1.0)

add_text(s1, Inches(1), sub_y + Inches(0.3), Inches(11), Inches(0.5),
         '演示文稿  |  2026年3月', font_size=BODY_SIZE,
         font_color=MED_GRAY, ea_font='KaiTi')
add_hline(s1, Inches(1), Inches(6.8), Inches(4), NAVY, Pt(2))
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
add_text(s5, left_x, content_top + Inches(0.6), col_w5, ..., 
              bullet=True, line_spacing=Pt(8))

# Right gray takeaway area
add_rect(s5, takeaway_left, Inches(1.2), takeaway_width, Inches(5.6), BG_GRAY)
add_text(s5, takeaway_left + Inches(0.15), Inches(1.35), takeaway_width - Inches(0.3), ...,
         'Key Takeaways', font_size=BODY_SIZE, color=NAVY, bold=True)
add_text(s5, takeaway_left + Inches(0.15), Inches(1.9), takeaway_width - Inches(0.3), ...,
              [f'{i+1}. {t}' for i, t in enumerate(takeaways)], line_spacing=Pt(10))
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
add_text(s, LM + Inches(0.3), Inches(5.1), CONTENT_W - Inches(0.6), Inches(1.4),
              ['洞见要点一', '洞见要点二', '洞见要点三'], line_spacing=Pt(8))
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
    add_text(s, cx + Inches(0.2), cy + Inches(1.5), cw - Inches(0.4), Inches(2.5),
                  desc.split('\n'), line_spacing=Pt(8), alignment=PP_ALIGN.CENTER)
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
    add_text(s, px + Inches(0.2), Inches(2.3), pw - Inches(0.4), Inches(3.5),
                  [f'• {p}' for p in points], line_spacing=Pt(10))
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
    add_text(s, cx + Inches(0.3), Inches(2.3), cw - Inches(0.6), Inches(3.5),
                  [f'• {p}' for p in points], line_spacing=Pt(10))
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
add_text(s, LM + Inches(0.3), Inches(2.4), hw - Inches(0.6), Inches(3.0),
              ['痛点一', '痛点二', '痛点三'], line_spacing=Pt(10))
# Arrow
add_text(s, LM + hw + Inches(0.1), Inches(3.2), Inches(1.5), Inches(0.5),
         '->', font_size=Pt(36), font_color=NAVY, bold=True, alignment=PP_ALIGN.CENTER)
# After
ax = LM + hw + Inches(1.733)
add_rect(s, ax, Inches(1.5), hw, Inches(4.5), NAVY)
add_text(s, ax + Inches(0.3), Inches(1.6), hw - Inches(0.6), Inches(0.5),
         'V  目标（After）', font_size=SUB_HEADER_SIZE, font_color=WHITE, bold=True)
add_hline(s, ax + Inches(0.3), Inches(2.2), hw - Inches(0.6), WHITE)
add_text(s, ax + Inches(0.3), Inches(2.4), hw - Inches(0.6), Inches(3.0),
              ['改进一', '改进二', '改进三'], font_color=WHITE, line_spacing=Pt(10))
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
add_text(s, LM, Inches(2.2), hw, Inches(2.5),
              ['• 优势要点一', '• 优势要点二', '• 优势要点三'], line_spacing=Pt(10))
# Cons column
cx = LM + hw + Inches(0.733)
add_text(s, cx, Inches(1.5), hw, Inches(0.4),
         'X  风险', font_size=SUB_HEADER_SIZE, font_color=DARK_GRAY, bold=True)
add_hline(s, cx, Inches(2.0), hw, DARK_GRAY)
add_text(s, cx, Inches(2.2), hw, Inches(2.5),
              ['• 风险要点一', '• 风险要点二', '• 风险要点三'], line_spacing=Pt(10))
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
add_text(s, LM, Inches(2.2), Inches(7.5), Inches(4.0),
              ['详细分析段落一', '', '详细分析段落二'], line_spacing=Pt(8))
# Right takeaway
tk_x = Inches(9.0)
tk_w = Inches(3.5)
add_rect(s, tk_x, Inches(1.5), tk_w, Inches(5.0), BG_GRAY)
add_text(s, tk_x + Inches(0.2), Inches(1.7), tk_w - Inches(0.4), Inches(0.4),
         'Key Takeaways', font_size=BODY_SIZE, font_color=NAVY, bold=True)
add_hline(s, tk_x + Inches(0.2), Inches(2.2), tk_w - Inches(0.4), LINE_GRAY)
add_text(s, tk_x + Inches(0.2), Inches(2.4), tk_w - Inches(0.4), Inches(3.8),
              ['1. 要点一', '2. 要点二', '3. 要点三'], line_spacing=Pt(10))
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
    add_text(s, cx, Inches(2.2), cw, Inches(4.0),
                  [f'• {p}' for p in points], line_spacing=Pt(10))
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
    add_text(s, cx + Inches(0.15), Inches(3.0), cw - Inches(0.3), Inches(3.0),
                  desc.split('\n'), line_spacing=Pt(8), alignment=PP_ALIGN.CENTER)
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
add_text(s, Inches(8.8), Inches(2.3), Inches(3.4), Inches(3.5),
              ['每个阶段的说明...'], line_spacing=Pt(10))
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
    add_text(s, cx + Inches(0.15), Inches(4.1), cw - Inches(0.3), Inches(2.0),
                  bio.split('\n'), line_spacing=Pt(8), alignment=PP_ALIGN.CENTER)
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
    add_text(s, cx + Inches(0.15), Inches(2.7), cw - Inches(0.3), Inches(2.0),
                  desc.split('\n'), line_spacing=Pt(8), alignment=PP_ALIGN.CENTER)
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
add_hline(s, LM, Inches(6.8), CW, NAVY, Pt(2))  # Full content width — not Inches(3)
```

---

### 类别 H：数据图表

> **触发规则**：当用户提供的内容包含 **日期/时间 + 数值/百分比** 的结构化数据（如舆情变化、销售趋势、KPI 周报、转化率变化等），**必须优先使用本类别的图表模式**，而不是 Data Table (#11) 或 Scorecard (#23)。
>
> **识别信号**（满足任一即触发）：
> - 数据中出现 `日期 + 百分比` 或 `日期 + 数值` 的组合
> - 提示词含 `████` 进度条字符 + 百分比
> - 内容涉及"趋势"、"演变"、"变化"、"走势"、"周报"、"日报"等时序关键词
> - 数据行数 ≥ 3 且每行包含至少一个类别和一个数值

#### 37. Grouped Bar Chart（分组柱状图 / 情绪热力图）

**适用场景**: 多个类别在不同时间点的数值对比（如舆情情绪分布、多产品销售对比、多指标周变化）。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  100% ─                                 │
│   80% ─  ██                             │
│   60% ─  ██ ██                          │
│   40% ─  ██ ██      ██      ██          │
│   20% ─  ██ ██ ██   ██ ██   ██ ██       │
│    0% ────────────────────────────────  │
│         3/4   3/6   3/8   3/10          │
│                                         │
│  ■ 正面  ■ 中性  ■ 负面                 │
│                                         │
│  ┌─BG_GRAY 趋势总结──────────────────┐  │
│  │ 总结文字                           │  │
│  └───────────────────────────────────┘  │
└─────────────────────────────────────────┘
```

**设计规范**:
- 柱状图使用 `add_rect()` 手工绘制，不依赖 matplotlib
- Y 轴标签（百分比）用 `add_text()` 左对齐
- X 轴标签（日期）用 `add_text()` 居中
- 每组柱子间留 0.3" 间距，组内柱子间留 0.05" 间距
- 图例用小矩形色块 + 文字标签，放在图表下方
- 底部可选趋势总结区（BG_GRAY）

**颜色分配**:
- 第一类别：NAVY (#051C2C) — 主要/正面
- 第二类别：LINE_GRAY (#CCCCCC) — 中性/基准
- 第三类别：MED_GRAY (#666666) — 次要/负面
- 第四类别：ACCENT_BLUE (#006BA6) — 扩展
- 若类别有语义色（如正面=NAVY, 负面=MED_GRAY），优先使用语义色

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '一周舆情演变：情绪分布从中性主导转向正面主导')

# ── 数据定义 ──
dates = ['3/4', '3/6', '3/8', '3/10', '3/11']
categories = ['正面', '中性', '负面']
cat_colors = [NAVY, LINE_GRAY, MED_GRAY]
# 每行 = 一个日期，每列 = 一个类别的百分比值
data = [
    [20, 80, 0],    # 3/4
    [75, 15, 10],   # 3/6
    [75, 20, 5],    # 3/8
    [75, 20, 5],    # 3/10
    [75, 20, 5],    # 3/11
]

# ── 图表区域参数 ──
chart_left = LM + Inches(0.8)         # 柱子起始 X（留 Y 轴标签空间）
chart_top = Inches(1.6)               # 图表顶部
chart_bottom = Inches(5.0)            # 图表底部（X 轴位置）
chart_height = chart_bottom - chart_top
chart_right = Inches(11.5)            # 图表右侧边界
chart_width = chart_right - chart_left

n_dates = len(dates)
n_cats = len(categories)
group_width = chart_width / n_dates   # 每组占据的总宽度
bar_width = Inches(0.35)              # 单根柱子宽度
bar_gap = Inches(0.05)               # 组内柱子间距
group_bar_width = bar_width * n_cats + bar_gap * (n_cats - 1)  # 一组柱子总宽

max_val = 100  # Y 轴最大值

# ── Y 轴刻度标签 + 水平参考线 ──
y_ticks = [0, 20, 40, 60, 80, 100]
for tick in y_ticks:
    tick_y = chart_bottom - chart_height * (tick / max_val)
    # Y 轴标签
    add_text(s, LM, tick_y - Inches(0.15), Inches(0.7), Inches(0.3),
             f'{tick}%', font_size=Pt(9), font_color=MED_GRAY,
             alignment=PP_ALIGN.RIGHT)
    # 水平参考线（极细浅灰色）
    if tick > 0:
        add_hline(s, chart_left, tick_y, chart_width, LINE_GRAY, Pt(0.25))

# ── X 轴基线 ──
add_hline(s, chart_left, chart_bottom, chart_width, BLACK, Pt(0.5))

# ── 绘制柱子 ──
for di, date in enumerate(dates):
    group_x = chart_left + group_width * di + (group_width - group_bar_width) / 2
    for ci, cat in enumerate(categories):
        val = data[di][ci]
        bar_h = chart_height * (val / max_val)
        bar_x = group_x + (bar_width + bar_gap) * ci
        bar_y = chart_bottom - bar_h
        if val > 0:
            add_rect(s, bar_x, bar_y, bar_width, bar_h, cat_colors[ci])
            # 柱顶数值标签（仅当值 >= 10% 时显示）
            if val >= 10:
                add_text(s, bar_x - Inches(0.05), bar_y - Inches(0.25),
                         bar_width + Inches(0.1), Inches(0.25),
                         f'{val}%', font_size=Pt(9), font_color=DARK_GRAY,
                         alignment=PP_ALIGN.CENTER)
    # X 轴日期标签
    add_text(s, chart_left + group_width * di, chart_bottom + Inches(0.05),
             group_width, Inches(0.3), date,
             font_size=BODY_SIZE, font_color=DARK_GRAY, alignment=PP_ALIGN.CENTER)

# ── 图例（图表下方居中）──
legend_y = Inches(5.5)
legend_start_x = Inches(4.5)
for ci, cat in enumerate(categories):
    lx = legend_start_x + Inches(1.8) * ci
    add_rect(s, lx, legend_y + Inches(0.05), Inches(0.2), Inches(0.2), cat_colors[ci])
    add_text(s, lx + Inches(0.3), legend_y, Inches(1.2), Inches(0.3),
             cat, font_size=Pt(12), font_color=DARK_GRAY)

# ── 底部趋势总结区域（可选）──
add_rect(s, LM, Inches(6.0), CONTENT_W, Inches(0.8), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(6.0), Inches(1.5), Inches(0.8),
         '趋势总结', font_size=BODY_SIZE, font_color=NAVY, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, LM + Inches(2.0), Inches(6.0), CONTENT_W - Inches(2.3), Inches(0.8),
         '舆情情绪从 3/4 的中性主导（80%）迅速转向正面主导（75%），负面情绪始终控制在 10% 以内',
         font_size=BODY_SIZE, font_color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)
add_source(s, 'Source: 舆情监测平台数据')
add_page_number(s, 5, 12)
```

#### 38. Stacked Bar Chart（堆叠柱状图 / 百分比占比图）

**适用场景**: 展示各类别在总体中的占比随时间变化（如市场份额演变、预算分配变化、渠道贡献占比）。适合强调"构成比例"而非"绝对值"。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  100% ─ ┌──┐  ┌──┐  ┌──┐  ┌──┐        │
│         │C │  │  │  │  │  │  │        │
│   50% ─ │B │  │B │  │  │  │  │        │
│         │  │  │  │  │B │  │B │        │
│         │A │  │A │  │A │  │A │        │
│    0% ──└──┘──└──┘──└──┘──└──┘────────  │
│         Q1    Q2    Q3    Q4            │
│                                         │
│  ■ A类  ■ B类  ■ C类                    │
│                                         │
│  ┌─BG_GRAY 关键发现──────────────────┐  │
│  │ 分析文字                           │  │
│  └───────────────────────────────────┘  │
└─────────────────────────────────────────┘
```

**设计规范**:
- 每根柱子内部从底部到顶部依次堆叠各类别
- 柱子宽度统一为 0.8"~1.2"（比分组柱状图更宽）
- 各段之间无间距，直接堆叠
- 百分比标签写在对应色块内部（当色块高度足够时），或省略
- 右侧可选放置"直接标签"指向最后一根柱子的各段

**颜色分配**（从底到顶）:
- 第一层（最大/最重要）：NAVY (#051C2C)
- 第二层：ACCENT_BLUE (#006BA6)
- 第三层：LINE_GRAY (#CCCCCC)
- 第四层：BG_GRAY (#F2F2F2) + 细边框
- 更多层级：使用 ACCENT_GREEN, ACCENT_ORANGE

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '渠道贡献占比：线上渠道在四个季度内从 30% 增长到 55%')

# ── 数据定义 ──
periods = ['Q1', 'Q2', 'Q3', 'Q4']
categories = ['线上直营', '线下门店', '经销商']
cat_colors = [NAVY, ACCENT_BLUE, LINE_GRAY]
# 每行 = 一个时间段，值为百分比（总和应为 100）
data = [
    [30, 45, 25],   # Q1
    [38, 40, 22],   # Q2
    [48, 32, 20],   # Q3
    [55, 28, 17],   # Q4
]

# ── 图表区域参数 ──
chart_left = LM + Inches(0.8)
chart_top = Inches(1.6)
chart_bottom = Inches(5.0)
chart_height = chart_bottom - chart_top
chart_right = Inches(9.5)
chart_width = chart_right - chart_left

n_periods = len(periods)
bar_width = Inches(1.0)              # 堆叠柱宽度
bar_spacing = chart_width / n_periods

max_val = 100

# ── Y 轴刻度标签 ──
y_ticks = [0, 25, 50, 75, 100]
for tick in y_ticks:
    tick_y = chart_bottom - chart_height * (tick / max_val)
    add_text(s, LM, tick_y - Inches(0.15), Inches(0.7), Inches(0.3),
             f'{tick}%', font_size=Pt(9), font_color=MED_GRAY,
             alignment=PP_ALIGN.RIGHT)
    if tick > 0:
        add_hline(s, chart_left, tick_y, chart_width, LINE_GRAY, Pt(0.25))

# ── X 轴基线 ──
add_hline(s, chart_left, chart_bottom, chart_width, BLACK, Pt(0.5))

# ── 绘制堆叠柱子 ──
for pi, period in enumerate(periods):
    bar_x = chart_left + bar_spacing * pi + (bar_spacing - bar_width) / 2
    cumulative = 0  # 从底部累积
    for ci in range(len(categories)):
        val = data[pi][ci]
        seg_h = chart_height * (val / max_val)
        seg_y = chart_bottom - chart_height * ((cumulative + val) / max_val)
        if val > 0:
            add_rect(s, bar_x, seg_y, bar_width, seg_h, cat_colors[ci])
            # 段内百分比标签（当段高 >= 0.4" 时显示）
            if seg_h >= Inches(0.4):
                label_color = WHITE if ci == 0 else (WHITE if ci == 1 else DARK_GRAY)
                add_text(s, bar_x, seg_y, bar_width, seg_h,
                         f'{val}%', font_size=Pt(11), font_color=label_color,
                         bold=True, alignment=PP_ALIGN.CENTER,
                         anchor=MSO_ANCHOR.MIDDLE)
        cumulative += val
    # X 轴标签
    add_text(s, chart_left + bar_spacing * pi, chart_bottom + Inches(0.05),
             bar_spacing, Inches(0.3), period,
             font_size=BODY_SIZE, font_color=DARK_GRAY, alignment=PP_ALIGN.CENTER)

# ── 右侧直接标签（指向最后一根柱子）──
last_bar_right = chart_left + bar_spacing * (n_periods - 1) + (bar_spacing + bar_width) / 2
label_x = last_bar_right + Inches(0.2)
cumulative = 0
for ci in range(len(categories)):
    val = data[-1][ci]
    mid_y = chart_bottom - chart_height * ((cumulative + val / 2) / max_val)
    add_text(s, label_x, mid_y - Inches(0.15), Inches(2.5), Inches(0.3),
             f'{categories[ci]} {val}%', font_size=Pt(11),
             font_color=cat_colors[ci] if ci < 2 else DARK_GRAY, bold=True)
    cumulative += val

# ── 图例（图表下方）──
legend_y = Inches(5.5)
legend_start_x = Inches(4.0)
for ci, cat in enumerate(categories):
    lx = legend_start_x + Inches(2.2) * ci
    add_rect(s, lx, legend_y + Inches(0.05), Inches(0.2), Inches(0.2), cat_colors[ci])
    add_text(s, lx + Inches(0.3), legend_y, Inches(1.6), Inches(0.3),
             cat, font_size=Pt(12), font_color=DARK_GRAY)

# ── 底部关键发现 ──
add_rect(s, LM, Inches(6.0), CONTENT_W, Inches(0.8), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(6.0), Inches(1.5), Inches(0.8),
         '关键发现', font_size=BODY_SIZE, font_color=NAVY, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, LM + Inches(2.0), Inches(6.0), CONTENT_W - Inches(2.3), Inches(0.8),
         '线上直营渠道占比从 Q1 的 30% 稳步增长至 Q4 的 55%，成为第一大收入来源',
         font_size=BODY_SIZE, font_color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)
add_source(s, 'Source: 内部销售数据')
add_page_number(s, 6, 12)
```

#### 39. Horizontal Bar Chart（水平柱状图 / 排名图）

**适用场景**: 类别名称较长的排名对比（如部门绩效排名、品牌认知度、功能使用率排行）。横向柱状图在类别较多时可读性更好。

```
┌─────────────────────────────────────────┐
│ ▌ Action Title                          │
├─────────────────────────────────────────┤
│  类别 A    ████████████████████  92%    │
│  类别 B    ████████████████     85%     │
│  类别 C    ██████████████       78%     │
│  类别 D    ████████████         65%     │
│  类别 E    ████████             52%     │
│                                         │
│  ┌─BG_GRAY 说明──────────────────────┐  │
│  │ 分析文字                           │  │
│  └───────────────────────────────────┘  │
└─────────────────────────────────────────┘
```

**设计规范**:
- 类别标签左对齐，柱子起始位置统一
- 最长柱子 = 100% 参考宽度
- 每根柱子右侧标注数值
- 第一名用 NAVY，其余用 BG_GRAY（或渐变灰色）
- 行间距均匀

```python
s = prs.slides.add_slide(BL)
add_action_title(s, '功能使用率排名：智能推荐以 92% 使用率位居第一')

# ── 数据定义（已排序）──
items = [
    ('智能推荐', 92),
    ('搜索功能', 85),
    ('个人中心', 78),
    ('消息通知', 65),
    ('社区互动', 52),
    ('数据报表', 38),
]

# ── 图表区域参数 ──
label_x = LM
label_w = Inches(2.0)
bar_x = LM + Inches(2.2)
bar_max_w = Inches(7.5)
value_x = bar_x + bar_max_w + Inches(0.2)
row_h = Inches(0.65)
start_y = Inches(1.6)
max_val = 100

# ── 绘制水平柱子 ──
for i, (name, val) in enumerate(items):
    ry = start_y + row_h * i
    bar_w = bar_max_w * (val / max_val)
    fill = NAVY if i == 0 else BG_GRAY
    tc = NAVY if i == 0 else DARK_GRAY
    # 类别标签
    add_text(s, label_x, ry, label_w, row_h, name,
             font_size=BODY_SIZE, font_color=tc, bold=(i == 0),
             anchor=MSO_ANCHOR.MIDDLE)
    # 背景轨道（浅灰底）
    add_rect(s, bar_x, ry + Inches(0.12), bar_max_w, Inches(0.4), RGBColor(0xF2, 0xF2, 0xF2))
    # 数据柱
    add_rect(s, bar_x, ry + Inches(0.12), bar_w, Inches(0.4), fill)
    # 数值标签
    add_text(s, value_x, ry, Inches(1.0), row_h, f'{val}%',
             font_size=BODY_SIZE, font_color=tc, bold=(i == 0),
             anchor=MSO_ANCHOR.MIDDLE)
    # 行分隔线
    if i < len(items) - 1:
        add_hline(s, label_x, ry + row_h, bar_max_w + Inches(2.5), LINE_GRAY, Pt(0.25))

# ── 底部说明 ──
add_rect(s, LM, Inches(5.8), CONTENT_W, Inches(0.9), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.8), Inches(1.5), Inches(0.9),
         '分析', font_size=BODY_SIZE, font_color=NAVY, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, LM + Inches(2.0), Inches(5.8), CONTENT_W - Inches(2.3), Inches(0.9),
         '智能推荐和搜索功能是用户最高频使用的两大核心功能，社区互动和数据报表仍有较大提升空间',
         font_size=BODY_SIZE, font_color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)
add_source(s, 'Source: 产品埋点数据，2026年2月')
add_page_number(s, 7, 12)
```

---

### Category I: Image + Content Layouts

> **Image Placeholder Convention**: Since python-pptx cannot embed web images at generation time, all image positions use a **gray placeholder rectangle** with crosshair lines and a label. The user replaces these with real images after generation.

#### Helper: `add_image_placeholder()`

```python
def add_image_placeholder(slide, left, top, width, height, label='Image'):
    """Draw a gray placeholder box with crosshair + label for image positions."""
    PLACEHOLDER_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
    # Background rect
    rect = add_rect(slide, left, top, width, height, PLACEHOLDER_GRAY)
    # Crosshair lines (diagonal from corners)
    add_hline(slide, left, top + height // 2, width, RGBColor(0xBB, 0xBB, 0xBB), Pt(0.5))
    # Vertical center line as thin rect
    vw = Pt(0.5)
    add_rect(slide, left + width // 2 - vw // 2, top, vw, height, RGBColor(0xBB, 0xBB, 0xBB))
    # Label
    add_text(slide, left, top + height // 2 - Inches(0.2), width, Inches(0.4),
             f'[ {label} ]', font_size=Pt(12), font_color=RGBColor(0x99, 0x99, 0x99),
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return rect
```

---

#### #40 — Content + Right Image

**Use case**: Text explanation on the left, supporting visual on the right — product screenshot, photo, diagram.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├───────────────────────┬──────────────────────┤
│  Heading              │                      │
│  • Bullet point 1     │   ┌──────────────┐   │
│  • Bullet point 2     │   │  IMAGE        │   │
│  • Bullet point 3     │   │  PLACEHOLDER  │   │
│                       │   └──────────────┘   │
│  Takeaway box (gray)  │                      │
├───────────────────────┴──────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '数字化转型的三大核心能力',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Left: text content (55%) ──
left_w = Inches(6.5)
ty = Inches(1.1)

add_text(s, LM, ty, left_w, Inches(0.4),
         '组织需要构建三项关键能力以驱动转型',
         font_size=Pt(18), font_color=NAVY, bold=True)

bullets = [
    '• 数据驱动决策：建立端到端数据采集、清洗与分析体系',
    '• 敏捷运营模式：从瀑布式开发转向双周迭代交付',
    '• 人才梯队建设：培养兼具业务理解与技术能力的复合型团队'
]
add_text(s, LM, ty + Inches(0.5), left_w, Inches(2.4),
         bullets, font_size=BODY_SIZE, font_color=DARK_GRAY, line_spacing=Pt(8))

# Takeaway box
add_rect(s, LM, Inches(4.5), left_w, Inches(0.8), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(4.5), left_w - Inches(0.6), Inches(0.8),
         '关键洞见：数据能力是三项能力中投资回报率最高的切入点',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

# ── Right: image placeholder (45%) ──
img_x = LM + left_w + Inches(0.3)
img_w = CONTENT_W - left_w - Inches(0.3)
add_image_placeholder(s, img_x, Inches(1.1), img_w, Inches(4.2), '产品截图 / 架构图')

add_source(s, 'Source: McKinsey Digital, 2026')
add_page_number(s, 3, 12)
```

---

#### #41 — Left Image + Content

**Use case**: Visual-first layout — image on left draws attention, text on right provides context.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────┬───────────────────────┤
│                      │  Heading              │
│  ┌──────────────┐    │  • Bullet point 1     │
│  │  IMAGE        │   │  • Bullet point 2     │
│  │  PLACEHOLDER  │   │  • Bullet point 3     │
│  └──────────────┘    │                       │
│                      │  Takeaway box (gray)  │
├──────────────────────┴───────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '客户旅程优化的关键触点分析',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Left: image placeholder (45%) ──
img_w = Inches(5.4)
add_image_placeholder(s, LM, Inches(1.1), img_w, Inches(4.2), '客户旅程地图')

# ── Right: text content (55%) ──
rx = LM + img_w + Inches(0.3)
rw = CONTENT_W - img_w - Inches(0.3)
ty = Inches(1.1)

add_text(s, rx, ty, rw, Inches(0.4),
         '五个关键触点决定80%的客户满意度',
         font_size=Pt(18), font_color=NAVY, bold=True)

bullets = [
    '• 首次接触：品牌认知与第一印象建立',
    '• 产品体验：核心功能的易用性与稳定性',
    '• 售后服务：响应速度与问题解决率',
    '• 续约决策：价值感知与竞品比较',
]
add_text(s, rx, ty + Inches(0.5), rw, Inches(2.4),
         bullets, font_size=BODY_SIZE, font_color=DARK_GRAY, line_spacing=Pt(8))

# Takeaway box
add_rect(s, rx, Inches(4.5), rw, Inches(0.8), BG_GRAY)
add_text(s, rx + Inches(0.2), Inches(4.5), rw - Inches(0.4), Inches(0.8),
         '建议优先投资"首次接触"和"产品体验"两个高杠杆触点',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 客户满意度调研数据，2026 Q1')
add_page_number(s, 4, 12)
```

---

#### #42 — Three Images + Descriptions

**Use case**: Visual comparison of three products, locations, or concepts side by side.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────┬──────────────┬────────────────┤
│ ┌──────────┐ │ ┌──────────┐ │ ┌──────────┐  │
│ │  IMAGE 1 │ │ │  IMAGE 2 │ │ │  IMAGE 3 │  │
│ └──────────┘ │ └──────────┘ │ └──────────┘  │
│  Title 1     │  Title 2     │  Title 3      │
│  Description │  Description │  Description  │
├──────────────┴──────────────┴────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '三大标杆项目的实施效果对比',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

items = [
    ('项目 A：智慧零售', '通过数字化门店改造，客流转化率提升35%，单店日均营收增长28%'),
    ('项目 B：供应链优化', '端到端库存周转天数从45天缩短至28天，缺货率降低至2.1%'),
    ('项目 C：会员体系', '会员复购率从22%提升至41%，ARPU值增长56%'),
]
col_w = Inches(3.7)
gap = Inches(0.35)
img_h = Inches(2.5)
ty = Inches(1.0)

for i, (title, desc) in enumerate(items):
    cx = LM + i * (col_w + gap)
    # Image placeholder
    add_image_placeholder(s, cx, ty, col_w, img_h, f'项目{chr(65+i)}实景照片')
    # Title
    add_text(s, cx, ty + img_h + Inches(0.15), col_w, Inches(0.35),
             title, font_size=Pt(16), font_color=NAVY, bold=True)
    # Description
    add_text(s, cx, ty + img_h + Inches(0.55), col_w, Inches(1.0),
             desc, font_size=BODY_SIZE, font_color=DARK_GRAY)

add_source(s, 'Source: 项目实施报告汇总，2025-2026')
add_page_number(s, 5, 12)
```

---

#### #43 — Image + Four Key Points

**Use case**: Central image/diagram with four callout points arranged around or beside it.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  ┌─────┬──────────┐  ┌─────┬──────────┐     │
│  │ 01  │ Point A   │  │ 02  │ Point B   │   │
│  └─────┴──────────┘  └─────┴──────────┘     │
│         ┌──────────────────────┐              │
│         │    IMAGE PLACEHOLDER │              │
│         └──────────────────────┘              │
│  ┌─────┬──────────┐  ┌─────┬──────────┐     │
│  │ 03  │ Point C   │  │ 04  │ Point D   │   │
│  └─────┴──────────┘  └─────┴──────────┘     │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '产品生态系统的四大核心模块',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Center image ──
img_w = Inches(5.0)
img_h = Inches(2.4)
img_x = LM + (CONTENT_W - img_w) / 2
add_image_placeholder(s, img_x, Inches(2.5), img_w, img_h, '产品生态架构图')

# ── Four points: 2 above, 2 below ──
points = [
    ('用户端', '移动App + 小程序，覆盖2亿月活用户'),
    ('商户端', 'SaaS管理平台，赋能50万商户'),
    ('数据中台', '实时数据处理能力达10亿条/天'),
    ('开放平台', 'API市场已接入300+合作伙伴'),
]
accents = [ACCENT_BLUE, ACCENT_GREEN, ACCENT_ORANGE, ACCENT_RED]
card_w = Inches(5.2)
card_h = Inches(0.7)
positions = [
    (LM + Inches(0.5), Inches(1.1)),
    (LM + CONTENT_W - card_w - Inches(0.5), Inches(1.1)),
    (LM + Inches(0.5), Inches(5.2)),
    (LM + CONTENT_W - card_w - Inches(0.5), Inches(5.2)),
]

for i, (title, desc) in enumerate(points):
    px, py = positions[i]
    add_oval(s, px, py + Inches(0.08), str(i + 1), bg=accents[i])
    add_text(s, px + Inches(0.55), py, Inches(1.5), Inches(0.35),
             title, font_size=Pt(16), font_color=accents[i], bold=True)
    add_text(s, px + Inches(0.55), py + Inches(0.35), card_w - Inches(0.55), Inches(0.35),
             desc, font_size=BODY_SIZE, font_color=DARK_GRAY)

add_source(s, 'Source: 产品架构文档，2026年3月')
add_page_number(s, 6, 12)
```

---

#### #44 — Full-Width Image with Overlay Text

**Use case**: Hero image covering the slide with semi-transparent overlay text — for visual storytelling, case study intros.

```
┌──────────────────────────────────────────────┐
│                                              │
│           FULL-WIDTH IMAGE                   │
│           PLACEHOLDER                        │
│                                              │
│    ┌─────────────────────────────────────┐   │
│    │ Semi-transparent dark overlay        │   │
│    │ "Quote or headline text"            │   │
│    │  — Attribution                       │   │
│    └─────────────────────────────────────┘   │
│                                              │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Full-width image placeholder ──
add_image_placeholder(s, Inches(0), Inches(0), Inches(13.333), Inches(6.5), '全幅背景图片')

# ── Dark semi-transparent overlay bar at bottom ──
overlay = add_rect(s, Inches(0), Inches(4.0), Inches(13.333), Inches(2.0),
                   RGBColor(0x05, 0x1C, 0x2C))
# Set transparency via alpha (70% opaque)
fill_elem = overlay._element.find(qn('p:spPr')).find(qn('a:solidFill'))
if fill_elem is not None:
    srgb = fill_elem.find(qn('a:srgbClr'))
    if srgb is not None:
        alpha = srgb.makeelement(qn('a:alpha'), {'val': '70000'})
        srgb.append(alpha)

add_text(s, LM, Inches(4.1), CONTENT_W, Inches(0.8),
         '"数字化不是选择题，而是生存题"',
         font_size=Pt(28), font_color=WHITE, bold=True,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, LM, Inches(4.9), CONTENT_W, Inches(0.4),
         '— 某全球500强企业CEO，2026年战略峰会',
         font_size=Pt(14), font_color=RGBColor(0xCC, 0xCC, 0xCC),
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 战略峰会实录，2026年1月')
add_page_number(s, 7, 12)
```

---

#### #45 — Case Study with Image

**Use case**: Extended case study with a visual — Situation, Approach, Result + supporting image.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├───────────────────────┬──────────────────────┤
│  SITUATION            │                      │
│  Background text...   │  ┌──────────────┐    │
│                       │  │  IMAGE        │   │
│  APPROACH             │  │  PLACEHOLDER  │   │
│  Method text...       │  └──────────────┘    │
│                       │                      │
│  RESULT               │  ┌─────┬─────┐      │
│  Outcome metrics...   │  │ KPI1│ KPI2│      │
│                       │  └─────┴─────┘      │
├───────────────────────┴──────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '案例：某零售集团全渠道数字化转型',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Left: SAR text (55%) ──
left_w = Inches(6.5)
sections = [
    ('背景', '该集团拥有3000+门店，线上渠道占比仅12%，客户数据分散在15个独立系统中'),
    ('方案', '分三阶段推进：1) 统一数据中台建设 2) 全渠道会员体系打通 3) 智能选品与补货系统上线'),
    ('成果', '线上渠道占比提升至38%，会员贡献收入占比从45%增至72%，库存周转提升40%'),
]
accents_sar = [ACCENT_BLUE, ACCENT_GREEN, ACCENT_ORANGE]
ty = Inches(1.0)
for i, (label, text) in enumerate(sections):
    # Section accent bar
    add_rect(s, LM, ty, Inches(0.06), Inches(1.0), accents_sar[i])
    add_text(s, LM + Inches(0.2), ty, Inches(1.2), Inches(0.3),
             label, font_size=Pt(16), font_color=accents_sar[i], bold=True)
    add_text(s, LM + Inches(0.2), ty + Inches(0.35), left_w - Inches(0.3), Inches(0.65),
             text, font_size=BODY_SIZE, font_color=DARK_GRAY)
    ty += Inches(1.3)

# ── Right: image + KPIs ──
rx = LM + left_w + Inches(0.3)
rw = CONTENT_W - left_w - Inches(0.3)
add_image_placeholder(s, rx, Inches(1.0), rw, Inches(2.5), '项目实施现场照片')

# KPI boxes
kpis = [('38%', '线上占比'), ('72%', '会员贡献'), ('+40%', '库存周转')]
kpi_w = rw / len(kpis)
for i, (val, label) in enumerate(kpis):
    kx = rx + i * kpi_w
    add_rect(s, kx, Inches(3.8), kpi_w - Inches(0.1), Inches(1.2), BG_GRAY)
    add_text(s, kx, Inches(3.85), kpi_w - Inches(0.1), Inches(0.6),
             val, font_size=Pt(28), font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER)
    add_text(s, kx, Inches(4.45), kpi_w - Inches(0.1), Inches(0.4),
             label, font_size=Pt(12), font_color=MED_GRAY,
             alignment=PP_ALIGN.CENTER)

add_source(s, 'Source: 项目交付报告，2025年12月')
add_page_number(s, 8, 12)
```

---

#### #46 — Quote with Background Image

**Use case**: Inspirational quote or key insight with a subtle background visual — for keynote-style emphasis slides.

```
┌──────────────────────────────────────────────┐
│                                              │
│       ┌──────────────────────────┐           │
│       │  IMAGE PLACEHOLDER       │           │
│       │  (subtle / blurred)      │           │
│       └──────────────────────────┘           │
│                                              │
│  ──────────────────────────────────          │
│  "Quote text in large font"                  │
│  — Speaker Name, Title                       │
│  ──────────────────────────────────          │
│                                              │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Background image (top half) ──
add_image_placeholder(s, Inches(0), Inches(0), Inches(13.333), Inches(3.2),
                      '主题相关背景图片（建议使用浅色/模糊效果）')

# ── White overlay for text area ──
add_rect(s, Inches(0), Inches(3.2), Inches(13.333), Inches(4.3), WHITE)

# ── Decorative lines ──
line_x = LM + Inches(1.0)
line_w = CONTENT_W - Inches(2.0)
add_hline(s, line_x, Inches(3.6), line_w, NAVY, Pt(1.0))

# ── Quote text ──
add_text(s, LM + Inches(1.5), Inches(3.8), CONTENT_W - Inches(3.0), Inches(1.4),
         '"最危险的不是变化本身，而是用昨天的逻辑做明天的决策"',
         font_size=Pt(24), font_color=NAVY, bold=True,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ── Attribution ──
add_text(s, LM + Inches(1.5), Inches(5.2), CONTENT_W - Inches(3.0), Inches(0.4),
         '— Peter Drucker，管理学大师',
         font_size=Pt(14), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER)

# ── Bottom decorative line ──
add_hline(s, line_x, Inches(5.7), line_w, NAVY, Pt(1.0))

add_source(s, 'Source: 《管理的实践》')
add_page_number(s, 9, 12)
```

---

#### #47 — Goals / Targets with Illustration

**Use case**: Strategic goals or OKRs with a supporting illustration — for goal-setting and planning slides.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├───────────────────────┬──────────────────────┤
│                       │                      │
│  ○ Goal 1 — desc      │  ┌──────────────┐   │
│  ○ Goal 2 — desc      │  │  IMAGE        │   │
│  ○ Goal 3 — desc      │  │  PLACEHOLDER  │   │
│  ○ Goal 4 — desc      │  └──────────────┘   │
│                       │                      │
│  Summary metric       │                      │
├───────────────────────┴──────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '2026年下半年四大战略目标',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Left: goals list (55%) ──
left_w = Inches(6.5)
goals = [
    ('营收增长', '实现全年营收目标120亿元，同比增长25%', ACCENT_BLUE),
    ('市场扩张', '新进入3个海外市场，海外收入占比提升至15%', ACCENT_GREEN),
    ('产品升级', 'AI功能覆盖率从40%提升至80%，用户NPS达到65+', ACCENT_ORANGE),
    ('组织发展', '关键岗位内部晋升率达60%，员工满意度≥4.2/5', ACCENT_RED),
]
ty = Inches(1.1)
for i, (title, desc, color) in enumerate(goals):
    # Accent bar
    add_rect(s, LM, ty, Inches(0.06), Inches(0.8), color)
    add_oval(s, LM + Inches(0.25), ty + Inches(0.15), str(i + 1), bg=color)
    add_text(s, LM + Inches(0.8), ty, left_w - Inches(1.0), Inches(0.35),
             title, font_size=Pt(16), font_color=color, bold=True)
    add_text(s, LM + Inches(0.8), ty + Inches(0.35), left_w - Inches(1.0), Inches(0.45),
             desc, font_size=BODY_SIZE, font_color=DARK_GRAY)
    ty += Inches(1.05)

# ── Right: illustration ──
rx = LM + left_w + Inches(0.3)
rw = CONTENT_W - left_w - Inches(0.3)
add_image_placeholder(s, rx, Inches(1.1), rw, Inches(4.2), '战略目标示意图 / 增长路线图')

add_source(s, 'Source: 2026年战略规划文件')
add_page_number(s, 10, 12)
```

---

### Category J: Advanced Data Visualization

> **Drawing Convention**: All charts are drawn with `add_rect()` and `add_oval()` — no matplotlib, no chart objects, no connectors. This ensures zero file corruption and full style control.

---

#### #48 — Donut Chart

**Use case**: Part-of-whole composition — market share, budget allocation, sentiment distribution. Up to 5 segments.

> **v2.0**: Uses BLOCK_ARC native shapes — only 4 shapes per chart (was hundreds of rect blocks). See Guard Rails Rule 9.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├───────────────────────┬──────────────────────┤
│                       │                      │
│    ┌───────────┐      │  ■ Segment A  45%    │
│    │  DONUT    │      │  ■ Segment B  28%    │
│    │ (BLOCK_   │      │  ■ Segment C  15%    │
│    │  ARC ×4)  │      │  ■ Segment D  12%    │
│    │  CENTER%  │      │                      │
│    └───────────┘      │  Insight text...     │
├───────────────────────┴──────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
from pptx.oxml.ns import qn

s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '2026年上半年营收渠道构成',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Donut chart using BLOCK_ARC (4 shapes total) ──
cx, cy = LM + Inches(3.0), Inches(3.2)  # center
outer_r = Inches(1.6)
inner_ratio = int((outer_r - Pt(10)) / outer_r * 50000)  # ~10px ring width
segments = [
    (0.45, NAVY, '线上直营'),
    (0.28, ACCENT_BLUE, '经销商'),
    (0.15, ACCENT_GREEN, '企业客户'),
    (0.12, ACCENT_ORANGE, '其他'),
]

cum_deg = 0  # start at top (0° = 12 o'clock, CW)
for pct, color, label in segments:
    sweep = pct * 360
    add_block_arc(s, cx - outer_r, cy - outer_r, outer_r * 2, outer_r * 2,
                  cum_deg, cum_deg + sweep, inner_ratio, color)
    cum_deg += sweep

# Center label (use WHITE for readability against colored ring)
add_text(s, cx - Inches(0.7), cy - Inches(0.3), Inches(1.4), Inches(0.6),
         '¥8.5亿', font_size=Pt(24), font_color=WHITE, bold=True,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         font_name='Georgia')
add_text(s, cx - Inches(0.7), cy + Inches(0.2), Inches(1.4), Inches(0.3),
         '总营收', font_size=Pt(12), font_color=WHITE,
         alignment=PP_ALIGN.CENTER)

# ── Legend (right side) ──
legend_x = LM + Inches(7.0)
legend_y = Inches(1.5)
for i, (pct, color, label) in enumerate(segments):
    ly = legend_y + i * Inches(0.8)
    add_rect(s, legend_x, ly + Inches(0.05), Inches(0.3), Inches(0.3), color)
    add_text(s, legend_x + Inches(0.45), ly, Inches(3.0), Inches(0.4),
             f'{label}  {int(pct*100)}%',
             font_size=Pt(16), font_color=DARK_GRAY, bold=True)

# ── Insight box ──
add_rect(s, legend_x, Inches(5.0), Inches(4.5), Inches(0.8), BG_GRAY)
add_text(s, legend_x + Inches(0.2), Inches(5.0), Inches(4.1), Inches(0.8),
         '线上直营渠道占比同比提升12个百分点，预计下半年将突破50%',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 财务报告，2026年H1')
add_page_number(s, 5, 12)
```

---

#### #49 — Waterfall Chart

**Use case**: Bridge from starting value to ending value showing incremental changes — revenue bridge, profit walk, budget variance.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│        ┌──┐                                  │
│  Start │  │ +A  -B  +C  -D  +E  ┌──┐ End   │
│        │  │ ┌┐  ┌┐  ┌┐  ┌┐  ┌┐  │  │       │
│        │  │ ││  ││  ││  ││  ││  │  │       │
│        │  │ └┘──└┘──└┘──└┘──└┘  │  │       │
│        └──┘                      └──┘       │
│  Takeaway text...                            │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '2026年H1利润增长桥接分析（百万元）',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Waterfall data ──
items = [
    ('2025 H2\n利润', 850, 'base'),
    ('营收\n增长', 220, 'up'),
    ('成本\n优化', 85, 'up'),
    ('人力\n支出', -120, 'down'),
    ('营销\n投入', -65, 'down'),
    ('新业务\n投资', -40, 'down'),
    ('2026 H1\n利润', 930, 'base'),
]

chart_left = LM + Inches(0.3)
chart_bottom = Inches(5.0)
chart_top = Inches(1.2)
chart_h = chart_bottom - chart_top
max_val = 1000  # Y-axis max
bar_w = Inches(1.2)
gap = Inches(0.4)

running = 0
for i, (label, val, typ) in enumerate(items):
    bx = chart_left + i * (bar_w + gap)

    if typ == 'base':
        # Full bar from bottom
        bar_h = int(chart_h * val / max_val)
        bar_top = chart_bottom - bar_h
        color = NAVY
        add_rect(s, bx, bar_top, bar_w, bar_h, color)
        running = val
    elif typ == 'up':
        bar_h = int(chart_h * val / max_val)
        bar_top = chart_bottom - int(chart_h * running / max_val) - bar_h
        color = ACCENT_GREEN
        add_rect(s, bx, bar_top, bar_w, bar_h, color)
        running += val
    else:  # down
        bar_h = int(chart_h * abs(val) / max_val)
        bar_top = chart_bottom - int(chart_h * running / max_val)
        color = ACCENT_RED
        add_rect(s, bx, bar_top, bar_w, bar_h, color)
        running += val

    # Value label above bar
    val_str = f'+{val}' if val > 0 and typ != 'base' else str(val)
    add_text(s, bx, bar_top - Inches(0.35), bar_w, Inches(0.3),
             val_str, font_size=Pt(14), font_color=DARK_GRAY, bold=True,
             alignment=PP_ALIGN.CENTER)
    # Category label below
    add_text(s, bx, chart_bottom + Inches(0.05), bar_w, Inches(0.5),
             label, font_size=Pt(11), font_color=MED_GRAY,
             alignment=PP_ALIGN.CENTER)

# ── Baseline axis ──
add_hline(s, chart_left, chart_bottom, Inches(11.5), LINE_GRAY, Pt(0.5))

# ── Takeaway ──
add_rect(s, LM, Inches(6.0), CONTENT_W, Inches(0.7), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(6.0), CONTENT_W - Inches(0.6), Inches(0.7),
         '关键发现：营收增长和成本优化共贡献305M增量，人力和营销支出消耗185M，净利润增长9.4%',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 财务管理部，2026年6月')
add_page_number(s, 6, 12)
```

---

#### #50 — Line / Trend Chart

**Use case**: Time-series trends — revenue growth, user count, market share over time. Supports 1-4 series.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  Y ─                                         │
│     ══════ Series A (black, bold) ═══ LabelA │
│     ══════ Series B (blue) ══════════ LabelB │
│     ══════ Series C (green) ═════════ LabelC │
│  0 ──────────────────────────────────        │
│     Q1'24  Q2'24  Q3'24  Q4'24  Q1'25       │
│  Takeaway text...                            │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '核心产品月活用户趋势（2024Q1 - 2026Q1）',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Chart area setup ──
chart_l = LM + Inches(0.8)
chart_r = LM + CONTENT_W - Inches(1.5)
chart_w = chart_r - chart_l
chart_top = Inches(1.3)
chart_bot = Inches(5.0)
chart_h = chart_bot - chart_top

# Y-axis labels
y_labels = ['0', '500万', '1000万', '1500万', '2000万']
for i, yl in enumerate(y_labels):
    yy = chart_bot - int(chart_h * i / (len(y_labels) - 1))
    add_text(s, LM, yy - Inches(0.12), Inches(0.7), Inches(0.24),
             yl, font_size=Pt(9), font_color=MED_GRAY, alignment=PP_ALIGN.RIGHT)
    if i > 0:
        add_hline(s, chart_l, yy, chart_w, RGBColor(0xE8, 0xE8, 0xE8), Pt(0.25))

# X-axis labels
x_labels = ['Q1\'24', 'Q2\'24', 'Q3\'24', 'Q4\'24', 'Q1\'25', 'Q2\'25',
            'Q3\'25', 'Q4\'25', 'Q1\'26']
n_pts = len(x_labels)
for i, xl in enumerate(x_labels):
    xx = chart_l + int(chart_w * i / (n_pts - 1))
    add_text(s, xx - Inches(0.3), chart_bot + Inches(0.05), Inches(0.6), Inches(0.25),
             xl, font_size=Pt(9), font_color=MED_GRAY, alignment=PP_ALIGN.CENTER)

# ── Data series ──
# Series values as fraction of max (2000万)
series = [
    ('产品A', [0.35,0.40,0.48,0.55,0.62,0.70,0.78,0.85,0.92], BLACK, Pt(3)),
    ('产品B', [0.20,0.22,0.25,0.30,0.35,0.38,0.42,0.45,0.50], ACCENT_BLUE, Pt(2)),
    ('产品C', [0.10,0.12,0.13,0.15,0.16,0.18,0.20,0.22,0.25], ACCENT_GREEN, Pt(2)),
]

for name, values, color, thickness in series:
    # Draw line segments as thin rects connecting data points
    for j in range(len(values) - 1):
        x1 = chart_l + int(chart_w * j / (n_pts - 1))
        y1 = chart_bot - int(chart_h * values[j])
        x2 = chart_l + int(chart_w * (j + 1) / (n_pts - 1))
        y2 = chart_bot - int(chart_h * values[j + 1])
        # Approximate line with thin rect (horizontal segment)
        seg_w = x2 - x1
        seg_y = min(y1, y2)
        seg_h = max(abs(y2 - y1), int(thickness))
        add_rect(s, x1, seg_y, seg_w, seg_h, color)
    # Data points as small squares
    for j, v in enumerate(values):
        px = chart_l + int(chart_w * j / (n_pts - 1))
        py = chart_bot - int(chart_h * v)
        dot_sz = Inches(0.08)
        add_rect(s, px - dot_sz // 2, py - dot_sz // 2, dot_sz, dot_sz, color)
    # End label
    last_x = chart_r + Inches(0.1)
    last_y = chart_bot - int(chart_h * values[-1])
    add_text(s, last_x, last_y - Inches(0.12), Inches(1.2), Inches(0.24),
             name, font_size=Pt(11), font_color=color, bold=True)

# ── Baseline axis ──
add_hline(s, chart_l, chart_bot, chart_w, BLACK, Pt(0.5))

# ── Takeaway ──
add_rect(s, LM, Inches(5.5), CONTENT_W, Inches(0.7), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.5), CONTENT_W - Inches(0.6), Inches(0.7),
         '关键趋势：产品A保持强劲增长势头，MAU有望在Q2\'26突破2000万大关',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 产品数据平台，2026年3月')
add_page_number(s, 4, 12)
```

---

#### #51 — Pareto Chart (Bar + Cumulative Line)

**Use case**: 80/20 analysis — identifying the vital few causes/items that account for most of the impact.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  Y₁ ─                                  ─ Y₂ │
│     ┌──┐                          ----100%   │
│     │  │┌──┐               ------             │
│     │  ││  │┌──┐     ------                   │
│     │  ││  ││  │┌──┐-                         │
│     │  ││  ││  ││  │┌──┐┌──┐                 │
│     └──┘└──┘└──┘└──┘└──┘└──┘    80% line     │
│  Takeaway: Top 3 items account for 78%       │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '客户投诉根因帕累托分析',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Data ──
items = [
    ('系统响应慢', 35), ('功能缺失', 25), ('操作复杂', 18),
    ('数据不准', 10), ('界面难看', 7), ('其他', 5),
]
total = sum(v for _, v in items)

chart_l = LM + Inches(0.5)
chart_bot = Inches(5.2)
chart_top = Inches(1.3)
chart_h = chart_bot - chart_top
bar_w = Inches(1.5)
gap = Inches(0.3)

cumulative = 0
cum_points = []
for i, (label, val) in enumerate(items):
    bx = chart_l + i * (bar_w + gap)
    pct = val / total
    bar_h = int(chart_h * pct)
    bar_top = chart_bot - bar_h

    # Bar
    add_rect(s, bx, bar_top, bar_w, bar_h, NAVY if i < 3 else LINE_GRAY)
    # Value label
    add_text(s, bx, bar_top - Inches(0.3), bar_w, Inches(0.25),
             f'{val}件 ({int(pct*100)}%)', font_size=Pt(11),
             font_color=DARK_GRAY, bold=(i < 3), alignment=PP_ALIGN.CENTER)
    # X-axis label
    add_text(s, bx, chart_bot + Inches(0.05), bar_w, Inches(0.4),
             label, font_size=Pt(11), font_color=MED_GRAY, alignment=PP_ALIGN.CENTER)

    # Cumulative point
    cumulative += pct
    cx_pt = bx + bar_w // 2
    cy_pt = chart_bot - int(chart_h * cumulative)
    cum_points.append((cx_pt, cy_pt))
    # Cumulative dot
    dot = Inches(0.1)
    add_rect(s, cx_pt - dot // 2, cy_pt - dot // 2, dot, dot, ACCENT_ORANGE)

# Connect cumulative dots with horizontal segments
for j in range(len(cum_points) - 1):
    x1, y1 = cum_points[j]
    x2, y2 = cum_points[j + 1]
    seg_w = x2 - x1
    seg_y = min(y1, y2)
    seg_h = max(abs(y2 - y1), Pt(2))
    add_rect(s, x1, seg_y, seg_w, seg_h, ACCENT_ORANGE)

# 80% threshold line (dashed approximation with small rects)
threshold_y = chart_bot - int(chart_h * 0.80)
dash_len = Inches(0.2)
total_w = len(items) * (bar_w + gap)
for d in range(0, int(total_w), int(dash_len * 2)):
    add_rect(s, chart_l + d, threshold_y, dash_len, Pt(1), ACCENT_RED)
add_text(s, chart_l + total_w + Inches(0.1), threshold_y - Inches(0.12),
         Inches(0.6), Inches(0.24), '80%',
         font_size=Pt(10), font_color=ACCENT_RED, bold=True)

# ── Baseline axis ──
add_hline(s, chart_l, chart_bot, total_w, BLACK, Pt(0.5))

# ── Takeaway ──
add_rect(s, LM, Inches(5.7), CONTENT_W, Inches(0.7), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.7), CONTENT_W - Inches(0.6), Inches(0.7),
         '分析：前3项根因占全部投诉的78%，优先解决"系统响应慢"可消除35%的投诉量',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 客服工单系统，2026年1-6月')
add_page_number(s, 7, 12)
```

---

#### #52 — Progress Bars / KPI Tracker

**Use case**: Multiple KPIs with target vs actual progress — project health, OKR tracking, sales pipeline.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  KPI Name          Actual / Target    Status │
│  ════════████████████░░░░░░░░   78%   ● On   │
│  ════════████████░░░░░░░░░░░░   52%   ● Risk │
│  ════════██████████████████░░   92%   ● On   │
│  ════════████████████████░░░   85%   ● On   │
│  ════════█████░░░░░░░░░░░░░░   38%   ● Off  │
│  Summary / insight text                      │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '2026年Q2 OKR达成进度追踪',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Header row ──
hy = Inches(1.0)
add_text(s, LM, hy, Inches(3.5), Inches(0.35),
         'KPI 指标', font_size=Pt(12), font_color=MED_GRAY, bold=True)
add_text(s, LM + Inches(3.5), hy, Inches(6.0), Inches(0.35),
         '进度', font_size=Pt(12), font_color=MED_GRAY, bold=True)
add_text(s, LM + Inches(9.5), hy, Inches(1.2), Inches(0.35),
         '达成率', font_size=Pt(12), font_color=MED_GRAY, bold=True,
         alignment=PP_ALIGN.CENTER)
add_text(s, LM + Inches(10.7), hy, Inches(1.0), Inches(0.35),
         '状态', font_size=Pt(12), font_color=MED_GRAY, bold=True,
         alignment=PP_ALIGN.CENTER)
add_hline(s, LM, hy + Inches(0.35), CONTENT_W, BLACK, Pt(0.75))

# ── KPI rows ──
kpis = [
    ('营收目标', 0.78, '¥9.4亿/¥12亿', 'on'),
    ('新客获取', 0.52, '2.6万/5万', 'risk'),
    ('客户留存率', 0.92, '92%/95%', 'on'),
    ('产品NPS', 0.85, '59/70', 'on'),
    ('成本控制', 0.38, '¥3.8亿/¥3.2亿', 'off'),
]

bar_x = LM + Inches(3.5)
bar_max_w = Inches(5.8)
bar_h = Inches(0.25)
row_h = Inches(0.7)

status_colors = {'on': ACCENT_GREEN, 'risk': ACCENT_ORANGE, 'off': ACCENT_RED}
status_labels = {'on': '达标', 'risk': '风险', 'off': '滞后'}

for i, (name, pct, detail, status) in enumerate(kpis):
    ry = Inches(1.6) + i * row_h
    # KPI name
    add_text(s, LM, ry, Inches(3.3), row_h,
             name, font_size=BODY_SIZE, font_color=DARK_GRAY, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
    # Progress bar background
    add_rect(s, bar_x, ry + (row_h - bar_h) / 2, bar_max_w, bar_h, BG_GRAY)
    # Progress bar fill
    fill_w = int(bar_max_w * min(pct, 1.0))
    fill_color = status_colors[status]
    add_rect(s, bar_x, ry + (row_h - bar_h) / 2, fill_w, bar_h, fill_color)
    # Percentage
    add_text(s, LM + Inches(9.5), ry, Inches(1.2), row_h,
             f'{int(pct*100)}%', font_size=Pt(16), font_color=DARK_GRAY, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Status indicator
    sc = status_colors[status]
    dot_sz = Inches(0.15)
    add_rect(s, LM + Inches(10.8), ry + (row_h - dot_sz) / 2, dot_sz, dot_sz, sc)
    add_text(s, LM + Inches(11.0), ry, Inches(0.7), row_h,
             status_labels[status], font_size=Pt(11), font_color=sc,
             anchor=MSO_ANCHOR.MIDDLE)
    # Row separator
    if i < len(kpis) - 1:
        add_hline(s, LM, ry + row_h, CONTENT_W, LINE_GRAY, Pt(0.25))

# ── Summary ──
add_rect(s, LM, Inches(5.5), CONTENT_W, Inches(0.8), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.5), Inches(1.5), Inches(0.8),
         '总结', font_size=BODY_SIZE, font_color=NAVY, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, LM + Inches(2.0), Inches(5.5), CONTENT_W - Inches(2.3), Inches(0.8),
         '5项KPI中3项达标，"新客获取"和"成本控制"需重点关注，建议Q3调整预算分配',
         font_size=BODY_SIZE, font_color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: OKR管理平台，2026年6月')
add_page_number(s, 8, 12)
```

---

#### #53 — Bubble / Scatter Plot

**Use case**: Two-variable comparison with size encoding — market attractiveness vs competitive position, impact vs effort.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  Y ─   High                                  │
│     ●(large)    ○(med)                       │
│              ●(small)    ○(large)             │
│     ○(med)         ●(med)                    │
│  0 ─   Low ──────────────────── High ─ X     │
│  Legend: ● Category A  ○ Category B          │
│  Takeaway text...                            │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '产品组合分析：市场吸引力 vs 竞争地位',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Axes ──
chart_l = LM + Inches(1.2)
chart_b = Inches(5.0)
chart_t = Inches(1.3)
chart_w = Inches(9.0)
chart_h = chart_b - chart_t

# X-axis
add_hline(s, chart_l, chart_b, chart_w, BLACK, Pt(0.5))
add_text(s, chart_l + chart_w // 2 - Inches(1.0), chart_b + Inches(0.15),
         Inches(2.0), Inches(0.3), '竞争地位 →',
         font_size=Pt(11), font_color=MED_GRAY, alignment=PP_ALIGN.CENTER)

# Y-axis (vertical rect)
add_rect(s, chart_l, chart_t, Pt(0.5), chart_h, BLACK)
add_text(s, LM, chart_t + chart_h // 2 - Inches(0.5), Inches(1.0), Inches(1.0),
         '市\n场\n吸\n引\n力\n↑', font_size=Pt(11), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ── Bubbles (oval shapes with size encoding) ──
# (x_pct, y_pct, size_inches, label, color)
bubbles = [
    (0.75, 0.80, 0.9, '产品A\n¥3.2亿', NAVY),
    (0.55, 0.60, 0.65, '产品B\n¥1.8亿', ACCENT_BLUE),
    (0.30, 0.75, 0.5, '产品C\n¥0.9亿', ACCENT_GREEN),
    (0.80, 0.35, 0.55, '产品D\n¥1.2亿', ACCENT_ORANGE),
    (0.20, 0.25, 0.35, '产品E\n¥0.4亿', LINE_GRAY),
    (0.45, 0.45, 0.7, '产品F\n¥2.1亿', ACCENT_BLUE),
]

for xp, yp, sz, label, color in bubbles:
    bx = chart_l + int(chart_w * xp) - Inches(sz / 2)
    by = chart_b - int(chart_h * yp) - Inches(sz / 2)
    oval = s.shapes.add_shape(MSO_SHAPE.OVAL, bx, by, Inches(sz), Inches(sz))
    oval.fill.solid()
    oval.fill.fore_color.rgb = color
    oval.line.fill.background()
    _clean_shape(oval)
    # Set 40% transparency
    fill_elem = oval._element.find(qn('p:spPr')).find(qn('a:solidFill'))
    if fill_elem is not None:
        srgb = fill_elem.find(qn('a:srgbClr'))
        if srgb is not None:
            alpha = srgb.makeelement(qn('a:alpha'), {'val': '60000'})
            srgb.append(alpha)
    # Label inside bubble
    add_text(s, bx, by + Inches(sz * 0.2), Inches(sz), Inches(sz * 0.6),
             label, font_size=Pt(9), font_color=WHITE, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ── Quadrant lines (dashed) ──
mid_x = chart_l + chart_w // 2
mid_y = chart_t + chart_h // 2
dash = Inches(0.15)
for d in range(0, int(chart_w), int(dash * 2)):
    add_rect(s, chart_l + d, mid_y, dash, Pt(0.5), RGBColor(0xDD, 0xDD, 0xDD))
for d in range(0, int(chart_h), int(dash * 2)):
    add_rect(s, mid_x, chart_t + d, Pt(0.5), dash, RGBColor(0xDD, 0xDD, 0xDD))

# ── Takeaway ──
add_rect(s, LM, Inches(5.5), CONTENT_W, Inches(0.7), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.5), CONTENT_W - Inches(0.6), Inches(0.7),
         '建议：优先投资产品A（高吸引力+强竞争力），观察产品F的增长潜力，逐步退出产品E',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 战略规划部，2026年Q1')
add_page_number(s, 9, 12)
```

---

#### #54 — Risk / Heat Matrix

**Use case**: Risk assessment — impact vs likelihood grid, with color-coded cells. Classic consulting risk register visualization.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│           Low Impact   Med Impact  High Impact│
│  High     ■ Yellow     ■ Orange   ■ Red      │
│  Prob     "Risk C"     "Risk A"   "Risk D"   │
│  Med      ■ Green      ■ Yellow   ■ Orange   │
│  Prob     "Risk F"     "Risk B"   "Risk E"   │
│  Low      ■ Green      ■ Green    ■ Yellow   │
│  Prob                              "Risk G"  │
│  Action items / mitigation plan              │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '项目风险评估矩阵',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Matrix setup ──
grid_l = LM + Inches(1.8)
grid_t = Inches(1.3)
cell_w = Inches(3.0)
cell_h = Inches(1.1)
rows = 3  # High, Medium, Low probability
cols = 3  # Low, Medium, High impact

# Color coding: [row][col] where row 0 = High prob, col 0 = Low impact
heat_colors = [
    [ACCENT_ORANGE, ACCENT_RED, ACCENT_RED],       # High prob
    [ACCENT_GREEN, ACCENT_ORANGE, ACCENT_RED],      # Med prob
    [ACCENT_GREEN, ACCENT_GREEN, ACCENT_ORANGE],    # Low prob
]
# Semi-transparent lighter versions for background
light_colors = [
    [LIGHT_ORANGE, LIGHT_RED, LIGHT_RED],
    [LIGHT_GREEN, LIGHT_ORANGE, LIGHT_RED],
    [LIGHT_GREEN, LIGHT_GREEN, LIGHT_ORANGE],
]

# Risks placed in cells: (row, col, name)
risks = [
    (0, 1, '数据泄露'), (0, 2, '核心人员流失'),
    (1, 0, '供应商延迟'), (1, 1, '预算超支'), (1, 2, '法规变更'),
    (2, 2, '汇率波动'),
]

# Y-axis labels
y_labels = ['高概率', '中概率', '低概率']
for r in range(rows):
    add_text(s, LM, grid_t + r * cell_h, Inches(1.6), cell_h,
             y_labels[r], font_size=Pt(13), font_color=DARK_GRAY, bold=True,
             alignment=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# X-axis labels
x_labels = ['低影响', '中影响', '高影响']
for c in range(cols):
    add_text(s, grid_l + c * cell_w, grid_t - Inches(0.35), cell_w, Inches(0.3),
             x_labels[c], font_size=Pt(13), font_color=DARK_GRAY, bold=True,
             alignment=PP_ALIGN.CENTER)

# Draw grid cells
for r in range(rows):
    for c in range(cols):
        cx = grid_l + c * cell_w
        cy = grid_t + r * cell_h
        add_rect(s, cx, cy, cell_w - Inches(0.05), cell_h - Inches(0.05),
                 light_colors[r][c])
        # Color indicator dot
        add_rect(s, cx + Inches(0.1), cy + Inches(0.1), Inches(0.2), Inches(0.2),
                 heat_colors[r][c])

# Place risk labels
for r, c, name in risks:
    cx = grid_l + c * cell_w
    cy = grid_t + r * cell_h
    add_text(s, cx + Inches(0.4), cy + Inches(0.25), cell_w - Inches(0.6), Inches(0.6),
             name, font_size=Pt(13), font_color=DARK_GRAY, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)

# ── Mitigation summary ──
add_rect(s, LM, Inches(4.8), CONTENT_W, Inches(1.2), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(4.85), Inches(1.5), Inches(0.3),
         '应对措施', font_size=Pt(14), font_color=NAVY, bold=True)
mitigations = [
    '• 红色区域（2项）：已制定应急预案，每周评审进展',
    '• 橙色区域（3项）：指定风险负责人，双周监控',
    '• 绿色区域（1项）：季度回顾，暂不采取主动措施',
]
add_text(s, LM + Inches(0.3), Inches(5.2), CONTENT_W - Inches(0.6), Inches(0.8),
         mitigations, font_size=BODY_SIZE, font_color=DARK_GRAY, line_spacing=Pt(4))

add_source(s, 'Source: 项目管理办公室，2026年Q2')
add_page_number(s, 10, 12)
```

**Variant: Matrix + Side Panel** — When the matrix needs an accompanying insight panel (e.g. "Key Changes", "Action Items"), use a compact grid (~60% width) with a side panel (~38% width). This prevents the panel from being crushed by a full-width grid.

```
┌──────────────────────────────────────────────────┐
│ [Action Title]                                   │
├──────────────────────────────────────────────────┤
│ Axis │ Col1   Col2   Col3 │ ┌─────────────────┐ │
│  ──  │ ■■■    ■■■    ■■■  │ │ Insight Panel   │ │
│  ↑   │ ■■■    ■■■    ■■■  │ │ • Bullet 1      │ │
│      │ ■■■    ■■■    ■■■  │ │ • Bullet 2      │ │
│      │   → Axis label →   │ │ ┌─────────────┐ │ │
│      │                     │ │ │ Summary box │ │ │
│      │                     │ │ └─────────────┘ │ │
│      │                     │ └─────────────────┘ │
├──────────────────────────────────────────────────┤
│ Source | Page N/Total                             │
└──────────────────────────────────────────────────┘
```

Layout math for the side-panel variant:

```python
# ── Compact grid + side panel layout ──
axis_label_w = Inches(0.65)            # Y-axis label column (tight)
grid_l       = LM + axis_label_w       # Grid left edge
cell_w       = Inches(2.15)            # Narrower cells (vs 3.0" default)
cell_h       = Inches(1.65)            # Taller cells to fill vertical space
grid_gap     = Inches(0.04)            # Minimal gap between cells

grid_right   = grid_l + 3 * (cell_w + grid_gap)
panel_gap    = Inches(0.25)
rx           = grid_right + panel_gap  # Panel left edge
rw           = LM + CONTENT_W - rx     # Panel width (~4.2")

# Panel height matches grid height
panel_h = 3 * (cell_h + grid_gap) - grid_gap

# Draw panel background
add_rect(s, rx, grid_t, rw, panel_h, BG_GRAY)
add_rect(s, rx, grid_t, rw, Inches(0.05), NAVY)  # Top accent line

# Optional: dark summary box at panel bottom
summary_h = Inches(0.65)
summary_y = grid_t + panel_h - summary_h
add_rect(s, rx, summary_y, rw, summary_h, NAVY)
add_text(s, rx + Inches(0.15), summary_y, rw - Inches(0.3), summary_h,
         'Key takeaway text here',
         font_size=Pt(11), font_color=WHITE, bold=True,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
```

> **Rule**: When a matrix needs a companion panel, shrink `cell_w` to ~2.15" (from 3.0") and `axis_label_w` to ~0.65" (from 1.8"). This yields a panel width of ~4.2" — enough for 6+ bullet items with comfortable reading. Never let the panel shrink below Inches(2.5).

---

#### #55 — Gauge / Dial Chart

**Use case**: Single KPI health indicator — customer satisfaction, system uptime, quality score. Visual "speedometer" metaphor.

> **v2.0**: Uses BLOCK_ARC native shapes — only 3 shapes for the arc (was 180+ rect blocks + white overlay). Horizontal rainbow arc (left→top→right). See Guard Rails Rule 9.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│                                              │
│         ╭──── ── ── ── ── ────╮              │
│      Red│   Orange    Green   │              │
│         ╰─────────────────────╯              │
│               78 / 100                       │
│                                              │
│  ┃ 当前NPS  ┃ 行业平均  ┃ 去年同期  ┃ 目标  │
│  ┃ 78       ┃ 52        ┃ 65        ┃ 80    │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
from pptx.oxml.ns import qn

s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '客户净推荐值（NPS）健康度仪表盘',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Gauge: horizontal rainbow arc using BLOCK_ARC (3 shapes) ──
# Arc goes from left (270° PPT) → top (0°) → right (90° PPT)
# Total sweep = 180° (a horizontal semi-circle, opening upward like ⌢)
cx = LM + CONTENT_W // 2
cy = Inches(3.8)
outer_r = Inches(2.2)
inner_ratio = int((outer_r - Pt(10)) / outer_r * 50000)  # ~10px ring width
score = 78  # out of 100

gauge_segs = [
    (0.40, ACCENT_RED),    # 0-40%: red zone (PPT 270° → 342°)
    (0.30, ACCENT_ORANGE), # 40-70%: orange zone (PPT 342° → 396°→36°)
    (0.30, ACCENT_GREEN),  # 70-100%: green zone (PPT 36° → 90°)
]

ppt_cum = 270  # start at left (270° in PPT CW from 12 o'clock)
for pct, color in gauge_segs:
    sweep = pct * 180  # half-circle, so 180° total
    add_block_arc(s, cx - outer_r, cy - outer_r, outer_r * 2, outer_r * 2,
                  ppt_cum % 360, (ppt_cum + sweep) % 360, inner_ratio, color)
    ppt_cum += sweep

# Center score
add_text(s, cx - Inches(0.8), cy - Inches(0.5), Inches(1.6), Inches(0.6),
         str(score), font_size=Pt(44), font_color=NAVY, bold=True,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         font_name='Georgia')
add_text(s, cx - Inches(0.5), cy + Inches(0.1), Inches(1.0), Inches(0.3),
         '/ 100', font_size=Pt(14), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER)

# Zone labels
add_text(s, cx - outer_r - Inches(0.3), cy + Inches(0.1), Inches(0.8), Inches(0.25),
         '差', font_size=Pt(11), font_color=ACCENT_RED, alignment=PP_ALIGN.CENTER)
add_text(s, cx - Inches(0.3), cy - outer_r - Inches(0.3), Inches(0.6), Inches(0.25),
         '良', font_size=Pt(11), font_color=ACCENT_ORANGE, alignment=PP_ALIGN.CENTER)
add_text(s, cx + outer_r - Inches(0.3), cy + Inches(0.1), Inches(0.8), Inches(0.25),
         '优', font_size=Pt(11), font_color=ACCENT_GREEN, alignment=PP_ALIGN.CENTER)

# ── Benchmark context ──
benchmarks = [
    ('当前 NPS', '78', ACCENT_GREEN),
    ('行业平均', '52', MED_GRAY),
    ('去年同期', '65', ACCENT_BLUE),
    ('目标值', '80', NAVY),
]
bx_start = LM + Inches(0.5)
by_row = Inches(5.0)
bw = Inches(2.5)
for i, (label, val, color) in enumerate(benchmarks):
    bx = bx_start + i * bw
    add_rect(s, bx, by_row, Inches(0.08), Inches(0.6), color)
    add_text(s, bx + Inches(0.2), by_row, bw - Inches(0.3), Inches(0.3),
             label, font_size=Pt(12), font_color=MED_GRAY)
    add_text(s, bx + Inches(0.2), by_row + Inches(0.3), bw - Inches(0.3), Inches(0.3),
             val, font_size=Pt(22), font_color=color, bold=True,
             font_name='Georgia')

add_source(s, 'Source: 客户体验部 NPS 调研，2026年Q2')
add_page_number(s, 11, 12)
```

---

#### #56 — Harvey Ball Status Table

**Use case**: Multi-criteria evaluation matrix — feature comparison, vendor assessment, capability maturity with visual fill indicators.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  Criteria     Option A   Option B   Option C │
│  ─────────────────────────────────────────── │
│  功能完整度     ●          ◕          ◑       │
│  用户体验       ◕          ●          ◔       │
│  技术可扩展     ◑          ◕          ●       │
│  实施成本       ◕          ◑          ●       │
│  供应商实力     ●          ◕          ◕       │
│  ─────────────────────────────────────────── │
│  Legend: ● Full  ◕ 75%  ◑ 50%  ◔ 25%  ○ 0% │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         'CRM系统供应商评估对比',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Table setup ──
criteria = ['功能完整度', '用户体验', '技术可扩展', '实施成本', '供应商实力', '数据安全']
options = ['供应商 A', '供应商 B', '供应商 C']
# Scores: 4=full, 3=75%, 2=50%, 1=25%, 0=empty
scores = [
    [4, 3, 2],
    [3, 4, 1],
    [2, 3, 4],
    [3, 2, 4],
    [4, 3, 3],
    [4, 4, 2],
]

col1_w = Inches(2.8)
col_w = Inches(2.5)
row_h = Inches(0.6)
table_l = LM
ty = Inches(1.1)

# Header row
add_text(s, table_l, ty, col1_w, row_h,
         '评估维度', font_size=Pt(13), font_color=NAVY, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
for j, opt in enumerate(options):
    add_text(s, table_l + col1_w + j * col_w, ty, col_w, row_h,
             opt, font_size=Pt(13), font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, table_l, ty + row_h, col1_w + len(options) * col_w, BLACK, Pt(0.75))

# Data rows
def draw_harvey_ball(slide, x, y, score, size=Inches(0.35)):
    """Draw Harvey ball: 0=empty circle, 1-3=partial fill, 4=full fill."""
    # Outer circle (always drawn)
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    oval.fill.solid()
    oval.fill.fore_color.rgb = BG_GRAY
    oval.line.color.rgb = NAVY
    oval.line.width = Pt(1.0)
    _clean_shape(oval)
    if score == 0:
        return
    # Fill proportion: draw a filled rect clipped visually by the oval context
    fill_w = int(size * score / 4)
    filled = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    filled.fill.solid()
    filled.fill.fore_color.rgb = NAVY
    filled.line.fill.background()
    _clean_shape(filled)
    if score < 4:
        # Overlay white rect to "erase" unfilled portion
        mask_x = x + fill_w
        mask_w = size - fill_w
        mask = add_rect(slide, mask_x, y, mask_w, size, WHITE)

for i, criterion in enumerate(criteria):
    ry = ty + row_h + Inches(0.05) + i * row_h
    add_text(s, table_l, ry, col1_w, row_h,
             criterion, font_size=BODY_SIZE, font_color=DARK_GRAY,
             anchor=MSO_ANCHOR.MIDDLE)
    for j in range(len(options)):
        ball_x = table_l + col1_w + j * col_w + (col_w - Inches(0.35)) // 2
        ball_y = ry + (row_h - Inches(0.35)) // 2
        draw_harvey_ball(s, ball_x, ball_y, scores[i][j])
    if i < len(criteria) - 1:
        add_hline(s, table_l, ry + row_h,
                  col1_w + len(options) * col_w, LINE_GRAY, Pt(0.25))

# ── Legend ──
legend_y = ty + row_h + len(criteria) * row_h + Inches(0.3)
add_hline(s, table_l, legend_y - Inches(0.1),
          col1_w + len(options) * col_w, BLACK, Pt(0.5))
legend_items = ['● 完全满足', '◕ 大部分满足', '◑ 部分满足', '◔ 少量满足']
lx = table_l
for item in legend_items:
    add_text(s, lx, legend_y, Inches(2.0), Inches(0.3),
             item, font_size=Pt(11), font_color=MED_GRAY)
    lx += Inches(2.2)

# ── Recommendation ──
add_rect(s, LM, Inches(5.5), CONTENT_W, Inches(0.8), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.5), Inches(1.5), Inches(0.8),
         '推荐', font_size=BODY_SIZE, font_color=NAVY, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, LM + Inches(2.0), Inches(5.5), CONTENT_W - Inches(2.3), Inches(0.8),
         '综合评估，供应商A在功能完整度和供应商实力方面领先，建议作为首选，供应商C可作为备选方案',
         font_size=BODY_SIZE, font_color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: IT采购评估小组，2026年Q2')
add_page_number(s, 7, 12)
```

---

### Category K: Dashboard Layouts

> **Dashboard Convention**: Dashboards pack multiple visual elements (KPIs, charts, tables) into a single dense slide. Use 3-4 distinct visual blocks minimum. Background panels (BG_GRAY) create clear section boundaries.

---

#### #57 — Dashboard: KPIs + Chart + Takeaways

**Use case**: Executive summary dashboard — top KPI cards, a chart in the middle, and key takeaways at the bottom.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├────────┬────────┬────────┬───────────────────┤
│  KPI 1 │  KPI 2 │  KPI 3 │  KPI 4           │
│  ¥8.5B │  +25%  │  78    │  92%             │
│  营收   │  增长率 │  NPS   │  留存率          │
├────────┴────────┴────────┴───────────────────┤
│                                              │
│  ┌──── Bar/Line Chart Area ─────────┐        │
│  │    (any chart pattern here)       │        │
│  └───────────────────────────────────┘        │
│                                              │
│  ┌──── Takeaway Panel ──────────────┐        │
│  │ • Key insight 1   • Key insight 2 │        │
│  └───────────────────────────────────┘        │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '业务健康度仪表盘 — 2026年Q2',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Top: KPI cards row ──
kpis = [
    ('¥8.5亿', '总营收', '+18% YoY', ACCENT_BLUE),
    ('+25%', '营收增速', '超目标5%', ACCENT_GREEN),
    ('78', 'NPS评分', '+13 vs去年', ACCENT_ORANGE),
    ('92%', '客户留存', '行业TOP10%', NAVY),
]
card_w = CONTENT_W / len(kpis) - Inches(0.15)
card_h = Inches(1.1)
ky = Inches(0.95)

for i, (val, label, detail, color) in enumerate(kpis):
    cx = LM + i * (card_w + Inches(0.15))
    add_rect(s, cx, ky, card_w, card_h, WHITE)
    add_rect(s, cx, ky, card_w, Inches(0.06), color)  # Top accent bar
    add_text(s, cx + Inches(0.2), ky + Inches(0.15), card_w - Inches(0.4), Inches(0.45),
             val, font_size=Pt(24), font_color=color, bold=True)
    add_text(s, cx + Inches(0.2), ky + Inches(0.6), Inches(1.5), Inches(0.2),
             label, font_size=Pt(11), font_color=MED_GRAY)
    add_text(s, cx + Inches(1.8), ky + Inches(0.6), card_w - Inches(2.0), Inches(0.2),
             detail, font_size=Pt(10), font_color=ACCENT_GREEN,
             alignment=PP_ALIGN.RIGHT)

# ── Middle: mini grouped bar chart ──
chart_y = Inches(2.3)
chart_h_area = Inches(2.5)
chart_l = LM + Inches(0.5)
chart_bot = chart_y + chart_h_area
months = ['1月', '2月', '3月', '4月', '5月', '6月']
# Two series: actual vs target
actual =  [1.2, 1.3, 1.4, 1.5, 1.4, 1.7]
target =  [1.3, 1.3, 1.4, 1.4, 1.5, 1.5]
max_val = 2.0
bar_w = Inches(0.6)
pair_gap = Inches(0.15)
group_w = bar_w * 2 + pair_gap
month_gap = Inches(0.5)

for i in range(len(months)):
    gx = chart_l + i * (group_w + month_gap)
    for j, (vals, color) in enumerate([(actual, NAVY), (target, BG_GRAY)]):
        bx = gx + j * (bar_w + pair_gap)
        val = vals[i]
        bh = int(chart_h_area * val / max_val)
        bt = chart_bot - bh
        add_rect(s, bx, bt, bar_w, bh, color)
        add_text(s, bx, bt - Inches(0.2), bar_w, Inches(0.18),
                 f'{val}亿', font_size=Pt(8), font_color=DARK_GRAY,
                 alignment=PP_ALIGN.CENTER)
    # Month label
    add_text(s, gx, chart_bot + Inches(0.03), group_w, Inches(0.2),
             months[i], font_size=Pt(9), font_color=MED_GRAY,
             alignment=PP_ALIGN.CENTER)

add_hline(s, chart_l, chart_bot, Inches(10.5), LINE_GRAY, Pt(0.5))

# Mini legend
add_rect(s, LM + Inches(9.0), chart_y, Inches(0.3), Inches(0.15), NAVY)
add_text(s, LM + Inches(9.4), chart_y - Inches(0.02), Inches(0.8), Inches(0.2),
         '实际', font_size=Pt(9), font_color=DARK_GRAY)
add_rect(s, LM + Inches(10.3), chart_y, Inches(0.3), Inches(0.15), BG_GRAY)
add_text(s, LM + Inches(10.7), chart_y - Inches(0.02), Inches(0.8), Inches(0.2),
         '目标', font_size=Pt(9), font_color=DARK_GRAY)

# ── Bottom: takeaway panel ──
add_rect(s, LM, Inches(5.3), CONTENT_W, Inches(0.9), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.3), Inches(1.5), Inches(0.9),
         '关键发现', font_size=BODY_SIZE, font_color=NAVY, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
takeaways = [
    '• 6月营收创单月新高（1.7亿），主要来自企业客户增长',
    '• NPS连续3个季度提升，但距行业标杆（85）仍有7分差距',
]
add_text(s, LM + Inches(2.0), Inches(5.3), CONTENT_W - Inches(2.3), Inches(0.9),
         takeaways, font_size=BODY_SIZE, font_color=DARK_GRAY,
         anchor=MSO_ANCHOR.MIDDLE, line_spacing=Pt(4))

add_source(s, 'Source: 业务数据平台，2026年6月')
add_page_number(s, 3, 12)
```

---

#### #58 — Dashboard: Table + Chart + Factoids

**Use case**: Data-dense overview — left table, right chart, bottom factoid cards. For board-level reporting.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├───────────────────────┬──────────────────────┤
│                       │                      │
│  ┌── Data Table ───┐  │  ┌── Chart ───────┐  │
│  │ Rows of data    │  │  │ Bars or lines  │  │
│  │ with values     │  │  │                │  │
│  └─────────────────┘  │  └────────────────┘  │
│                       │                      │
├────────┬──────────┬───┴──────────┬───────────┤
│ Fact 1 │ Fact 2   │  Fact 3      │ Fact 4    │
│ "120+" │ "¥2.3B"  │  "Top 5%"   │ "99.9%"   │
├────────┴──────────┴──────────────┴───────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '区域业绩总览 — 2026年H1',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Left: data table (55%) ──
left_w = Inches(6.2)
ty = Inches(1.0)
cols = ['区域', 'H1营收', '同比', '达成率']
col_ws = [Inches(1.5), Inches(1.5), Inches(1.2), Inches(1.5)]
rows_data = [
    ('华东', '¥3.2亿', '+22%', '105%'),
    ('华南', '¥2.1亿', '+18%', '98%'),
    ('华北', '¥1.8亿', '+15%', '92%'),
    ('西部', '¥0.9亿', '+28%', '110%'),
    ('海外', '¥0.5亿', '+45%', '85%'),
]

# Table header
hx = LM
for ci, (col_name, cw) in enumerate(zip(cols, col_ws)):
    add_text(s, hx, ty, cw, Inches(0.3),
             col_name, font_size=Pt(12), font_color=NAVY, bold=True)
    hx += cw
add_hline(s, LM, ty + Inches(0.3), left_w, BLACK, Pt(0.75))

# Table rows
for ri, row in enumerate(rows_data):
    ry = ty + Inches(0.4) + ri * Inches(0.5)
    rx = LM
    for ci, (val, cw) in enumerate(zip(row, col_ws)):
        fc = DARK_GRAY
        bld = False
        if ci == 2:  # Growth column
            fc = ACCENT_GREEN if '+' in val else ACCENT_RED
        if ci == 3:  # Achievement
            pct_val = int(val.replace('%', ''))
            fc = ACCENT_GREEN if pct_val >= 100 else (ACCENT_ORANGE if pct_val >= 90 else ACCENT_RED)
            bld = True
        add_text(s, rx, ry, cw, Inches(0.4),
                 val, font_size=BODY_SIZE, font_color=fc, bold=bld,
                 anchor=MSO_ANCHOR.MIDDLE)
        rx += cw
    if ri < len(rows_data) - 1:
        add_hline(s, LM, ry + Inches(0.45), left_w, LINE_GRAY, Pt(0.25))

# ── Right: mini horizontal bar chart (45%) ──
chart_x = LM + left_w + Inches(0.5)
chart_w = CONTENT_W - left_w - Inches(0.5)
chart_ty = Inches(1.0)
bar_max = Inches(3.5)
max_rev = 3.2

add_text(s, chart_x, chart_ty, chart_w, Inches(0.3),
         '各区域营收对比', font_size=Pt(12), font_color=NAVY, bold=True)
add_hline(s, chart_x, chart_ty + Inches(0.3), chart_w, BLACK, Pt(0.5))

regions = [('华东', 3.2), ('华南', 2.1), ('华北', 1.8), ('西部', 0.9), ('海外', 0.5)]
bar_h = Inches(0.3)
bar_gap = Inches(0.15)
for i, (region, rev) in enumerate(regions):
    by = chart_ty + Inches(0.45) + i * (bar_h + bar_gap)
    bw = int(bar_max * rev / max_rev)
    add_text(s, chart_x, by, Inches(0.8), bar_h,
             region, font_size=Pt(11), font_color=MED_GRAY,
             anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, chart_x + Inches(0.9), by, bw, bar_h, NAVY if i == 0 else ACCENT_BLUE)
    add_text(s, chart_x + Inches(0.9) + bw + Inches(0.1), by, Inches(0.8), bar_h,
             f'¥{rev}亿', font_size=Pt(11), font_color=DARK_GRAY, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)

# ── Bottom: factoid cards ──
facts = [
    ('120+', '服务客户数', ACCENT_BLUE),
    ('¥8.5亿', 'H1总营收', NAVY),
    ('Top 5%', '行业排名', ACCENT_GREEN),
    ('99.2%', '服务可用性', ACCENT_ORANGE),
]
card_y = Inches(4.8)
card_w = CONTENT_W / len(facts) - Inches(0.15)
card_h = Inches(1.0)
for i, (val, label, color) in enumerate(facts):
    fx = LM + i * (card_w + Inches(0.15))
    add_rect(s, fx, card_y, card_w, card_h, BG_GRAY)
    add_rect(s, fx, card_y, Inches(0.06), card_h, color)  # Left accent
    add_text(s, fx + Inches(0.2), card_y + Inches(0.1), card_w - Inches(0.3), Inches(0.5),
             val, font_size=Pt(24), font_color=color, bold=True)
    add_text(s, fx + Inches(0.2), card_y + Inches(0.6), card_w - Inches(0.3), Inches(0.3),
             label, font_size=Pt(11), font_color=MED_GRAY)

add_source(s, 'Source: 区域业务部 & 财务部，2026年H1')
add_page_number(s, 4, 12)
```

---

### Category L: Visual Storytelling & Special

> **Storytelling Convention**: These layouts emphasize visual narrative patterns commonly found in McKinsey decks — stakeholder maps, decision trees, checklists, and icon-driven grids. They add variety beyond standard charts and text layouts.

---

#### #59 — Stakeholder Map

**Use case**: Influence vs interest matrix for stakeholders — change management, project governance, communication planning.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  Interest ↑                                  │
│  High  ┌─────────────┬──────────────┐        │
│        │ Keep Informed│ Manage Closely│       │
│        │  (name)      │  (name)      │       │
│        ├─────────────┼──────────────┤        │
│  Low   │ Monitor     │ Keep Satisfied│       │
│        │  (name)      │  (name)      │       │
│        └─────────────┴──────────────┘        │
│             Low        High → Influence       │
│  Action plan text...                         │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '项目利益相关者影响力-关注度矩阵',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── 2x2 grid ──
grid_l = LM + Inches(2.0)
grid_t = Inches(1.2)
cell_w = Inches(4.5)
cell_h = Inches(1.8)

quadrant_labels = [
    ('保持沟通', 'Keep Informed', LIGHT_BLUE),
    ('重点管理', 'Manage Closely', LIGHT_RED),
    ('定期监测', 'Monitor', LIGHT_GREEN),
    ('维护满意', 'Keep Satisfied', LIGHT_ORANGE),
]
quadrant_stakeholders = [
    ['产品经理', '设计团队'],
    ['CEO', 'CTO', '投资方'],
    ['法务部', '行政部'],
    ['运维团队', '外部供应商'],
]

for qi, (label_cn, label_en, bg_color) in enumerate(quadrant_labels):
    row = qi // 2
    col = qi % 2
    qx = grid_l + col * cell_w
    qy = grid_t + row * cell_h
    # Cell background
    add_rect(s, qx, qy, cell_w - Inches(0.05), cell_h - Inches(0.05), bg_color)
    # Quadrant title
    add_text(s, qx + Inches(0.15), qy + Inches(0.1), cell_w - Inches(0.3), Inches(0.35),
             f'{label_cn} ({label_en})',
             font_size=Pt(13), font_color=NAVY, bold=True)
    # Stakeholder names
    names = quadrant_stakeholders[qi]
    for ni, name in enumerate(names):
        add_oval(s, qx + Inches(0.2), qy + Inches(0.55) + ni * Inches(0.4),
                 name[0], size=Inches(0.3), bg=NAVY)
        add_text(s, qx + Inches(0.6), qy + Inches(0.5) + ni * Inches(0.4),
                 Inches(2.5), Inches(0.35),
                 name, font_size=BODY_SIZE, font_color=DARK_GRAY)

# ── Axis labels ──
add_text(s, LM, grid_t + cell_h - Inches(0.3), Inches(1.8), Inches(0.6),
         '关\n注\n度\n↑', font_size=Pt(12), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, grid_l + cell_w - Inches(0.5), grid_t + 2 * cell_h + Inches(0.1),
         Inches(2.5), Inches(0.3),
         '影响力 →', font_size=Pt(12), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER)

# ── Action plan ──
add_rect(s, LM, Inches(5.2), CONTENT_W, Inches(0.8), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.2), CONTENT_W - Inches(0.6), Inches(0.8),
         '行动计划：本周安排CEO一对一沟通，每双周向投资方发送项目简报，产品团队纳入每日站会',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 项目管理办公室，2026年Q2')
add_page_number(s, 5, 12)
```

---

#### #60 — Issue / Decision Tree

**Use case**: Breaking down a complex decision into sub-decisions — problem decomposition, MECE logic tree, diagnostic framework.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│                                              │
│  ┌────────┐                                  │
│  │ Root   │──┬── ┌────────┐──┬── ┌────────┐ │
│  │ Issue  │  │   │ Branch │  │   │ Leaf 1 │ │
│  └────────┘  │   │   A    │  │   └────────┘ │
│              │   └────────┘  └── ┌────────┐ │
│              │                    │ Leaf 2 │ │
│              │                    └────────┘ │
│              └── ┌────────┐──┬── ┌────────┐ │
│                  │ Branch │  │   │ Leaf 3 │ │
│                  │   B    │  └── └────────┘ │
│                  └────────┘                  │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '利润下滑根因诊断树',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Tree structure (3 levels) ──
# Level 0: Root
L0_x, L0_y = LM + Inches(0.3), Inches(2.5)
L0_w, L0_h = Inches(2.2), Inches(1.2)
add_rect(s, L0_x, L0_y, L0_w, L0_h, NAVY)
add_text(s, L0_x, L0_y, L0_w, L0_h,
         '利润下滑\n15%',
         font_size=Pt(16), font_color=WHITE, bold=True,
         alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Level 1: branches
L1_items = [
    ('营收下降', '–8%'),
    ('成本上升', '+12%'),
]
L1_x = L0_x + L0_w + Inches(0.6)
L1_w, L1_h = Inches(2.0), Inches(0.9)

# Connecting lines from L0 to L1
conn_x = L0_x + L0_w
for i, (title, metric) in enumerate(L1_items):
    L1_y = Inches(1.5) + i * Inches(2.2)
    # Horizontal connector
    add_hline(s, conn_x, L0_y + L0_h // 2, L1_x - conn_x, LINE_GRAY, Pt(1))
    # Vertical segment to branch
    if i == 0:
        add_rect(s, L1_x - Inches(0.02), L1_y + L1_h // 2, Pt(1),
                 L0_y + L0_h // 2 - L1_y - L1_h // 2, LINE_GRAY)
    else:
        add_rect(s, L1_x - Inches(0.02), L0_y + L0_h // 2, Pt(1),
                 L1_y + L1_h // 2 - L0_y - L0_h // 2, LINE_GRAY)

    add_rect(s, L1_x, L1_y, L1_w, L1_h, ACCENT_BLUE if i == 0 else ACCENT_ORANGE)
    add_text(s, L1_x, L1_y, L1_w, Inches(0.5),
             title, font_size=Pt(14), font_color=WHITE, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, L1_x, L1_y + Inches(0.45), L1_w, Inches(0.4),
             metric, font_size=Pt(18), font_color=WHITE, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Level 2: leaves
L2_groups = [
    [('客户流失', '–5%'), ('价格下调', '–3%')],
    [('原材料涨价', '+7%'), ('人力成本', '+3%'), ('运营费用', '+2%')],
]
L2_x = L1_x + L1_w + Inches(0.6)
L2_w, L2_h = Inches(2.0), Inches(0.65)

for gi, group in enumerate(L2_groups):
    parent_y = Inches(1.5) + gi * Inches(2.2)
    parent_cx = L1_x + L1_w
    for li, (title, metric) in enumerate(group):
        L2_y = parent_y - Inches(0.3) + li * Inches(0.8)
        # Connector
        add_hline(s, parent_cx, parent_y + L1_h // 2,
                  L2_x - parent_cx, LINE_GRAY, Pt(0.5))
        color = BG_GRAY
        add_rect(s, L2_x, L2_y, L2_w, L2_h, color)
        add_text(s, L2_x + Inches(0.1), L2_y, L2_w * 0.6, L2_h,
                 title, font_size=Pt(12), font_color=DARK_GRAY,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, L2_x + L2_w * 0.6, L2_y, L2_w * 0.4, L2_h,
                 metric, font_size=Pt(14), font_color=NAVY, bold=True,
                 anchor=MSO_ANCHOR.MIDDLE, alignment=PP_ALIGN.CENTER)

# Level 3: action items (rightmost)
L3_x = L2_x + L2_w + Inches(0.4)
L3_w = CONTENT_W - (L3_x - LM)
add_rect(s, L3_x, Inches(1.2), L3_w, Inches(4.5), BG_GRAY)
add_text(s, L3_x + Inches(0.15), Inches(1.3), L3_w - Inches(0.3), Inches(0.3),
         '建议行动', font_size=Pt(14), font_color=NAVY, bold=True)
actions = [
    '1. 客户挽回计划：Top 20客户专项拜访',
    '2. 价格策略：阶梯定价替代统一折扣',
    '3. 供应链：锁定6个月期货合约',
    '4. 人效提升：AI工具导入减少20%人力',
    '5. 费用管控：暂停非核心项目支出',
]
add_text(s, L3_x + Inches(0.15), Inches(1.7), L3_w - Inches(0.3), Inches(3.5),
         actions, font_size=Pt(12), font_color=DARK_GRAY, line_spacing=Pt(8))

add_source(s, 'Source: 财务分析部，2026年Q2')
add_page_number(s, 6, 12)
```

---

#### #61 — Five-Row Checklist / Status

**Use case**: Task completion status, implementation checklist, audit findings — each row with status indicator.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  # │ Task / Item         │ Owner │ Status    │
│  ──┼─────────────────────┼───────┼────────── │
│  1 │ Data migration       │ TechOps│ ✓ Done   │
│  2 │ UAT testing          │ QA    │ ✓ Done    │
│  3 │ Security audit       │ InfoSec│ → Active │
│  4 │ Training rollout     │ HR    │ ○ Pending │
│  5 │ Go-live sign-off     │ PMO   │ ○ Pending │
│  ──┼─────────────────────┼───────┼────────── │
│  Progress: 2/5 complete (40%)               │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_action_title(s, '系统上线前检查清单 — 进度追踪')

# ── Table header ──
hy = CONTENT_TOP + Inches(0.1)
headers = [('#', Inches(0.5)), ('任务项', Inches(4.5)), ('负责人', Inches(1.8)),
           ('截止日期', Inches(1.8)), ('状态', Inches(2.0))]
hx = LM
for label, w in headers:
    add_text(s, hx, hy, w, Inches(0.35),
             label, font_size=Pt(12), font_color=NAVY, bold=True)
    hx += w
add_hline(s, LM, hy + Inches(0.35), CW, BLACK, Pt(0.75))

# ── Checklist rows ──
tasks = [
    ('1', '数据迁移与验证', '技术运维部', '3月15日', 'done'),
    ('2', 'UAT用户验收测试', 'QA团队', '3月20日', 'done'),
    ('3', '信息安全审计', '信息安全部', '3月25日', 'active'),
    ('4', '全员培训与上手', 'HR & 培训部', '4月1日', 'pending'),
    ('5', '上线审批签字', 'PMO', '4月5日', 'pending'),
    ('6', '灰度发布与监控', '技术运维部', '4月8日', 'pending'),
    ('7', '全量上线', 'PMO', '4月10日', 'pending'),
]

status_config = {
    'done':    ('✓ 完成', ACCENT_GREEN, LIGHT_GREEN),
    'active':  ('→ 进行中', ACCENT_ORANGE, LIGHT_ORANGE),
    'pending': ('○ 待启动', MED_GRAY, BG_GRAY),
}

# ── Dynamic row height: fit all rows without overflowing page ──
data_start_y = hy + Inches(0.5)
bottom_limit = SOURCE_Y - Inches(0.1)  # or BOTTOM_BAR_Y if using bottom bar
available_h = bottom_limit - data_start_y
row_h = min(Inches(0.85), available_h / max(len(tasks), 1))  # cap at 0.85"

# Use smaller font when rows are tight
row_font = SMALL_SIZE if row_h < Inches(0.65) else BODY_SIZE

for i, (num, task, owner, deadline, status) in enumerate(tasks):
    ry = data_start_y + i * row_h
    st_label, st_color, st_bg = status_config[status]

    # Row background for alternating
    if i % 2 == 0:
        add_rect(s, LM, ry, CW, row_h, RGBColor(0xFA, 0xFA, 0xFA))

    rx = LM
    vals = [(num, Inches(0.5)), (task, Inches(4.5)), (owner, Inches(1.8)),
            (deadline, Inches(1.8))]
    for val, w in vals:
        add_text(s, rx, ry, w, row_h,
                 val, font_size=row_font, font_color=DARK_GRAY,
                 anchor=MSO_ANCHOR.MIDDLE)
        rx += w

    # Status badge
    badge_h = min(row_h - Inches(0.2), Inches(0.35))
    add_rect(s, rx + Inches(0.1), ry + (row_h - badge_h) / 2,
             Inches(1.5), badge_h, st_bg)
    add_text(s, rx + Inches(0.1), ry, Inches(1.5), row_h,
             st_label, font_size=Pt(12), font_color=st_color, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    if i < len(tasks) - 1:
        add_hline(s, LM, ry + row_h, CW, LINE_GRAY, Pt(0.25))

add_source(s, 'Source: 项目管理办公室，2026年3月')
add_page_number(s, 8, 12)
```

---

#### #62 — Metric Comparison Row

**Use case**: Before/after or multi-period comparison with large numbers — performance review, transformation impact, A/B test results.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│                                              │
│  ┌────────────┐   →   ┌────────────┐        │
│  │  BEFORE     │       │  AFTER      │       │
│  │  ¥5.2亿     │       │  ¥8.5亿     │       │
│  │  营收        │       │  营收        │       │
│  └────────────┘       └────────────┘        │
│  ┌────────────┐   →   ┌────────────┐        │
│  │  45天       │       │  28天       │       │
│  │  库存周转    │       │  库存周转    │       │
│  └────────────┘       └────────────┘        │
│  Summary text...                             │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '数字化转型前后关键指标对比',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Metric pairs ──
metrics = [
    ('营收规模', '¥5.2亿', '¥8.5亿', '+63%'),
    ('库存周转', '45天', '28天', '–38%'),
    ('客户NPS', '52', '78', '+50%'),
    ('线上占比', '12%', '38%', '+217%'),
]

row_h = Inches(0.95)
before_x = LM + Inches(0.5)
after_x = LM + Inches(6.5)
card_w = Inches(4.0)
arrow_x = before_x + card_w + Inches(0.3)
delta_x = after_x + card_w + Inches(0.3)

# Column headers
add_text(s, before_x, Inches(1.0), card_w, Inches(0.3),
         '转型前 (2024)', font_size=Pt(13), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER)
add_text(s, after_x, Inches(1.0), card_w, Inches(0.3),
         '转型后 (2026)', font_size=Pt(13), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER)
add_text(s, delta_x, Inches(1.0), Inches(1.5), Inches(0.3),
         '变化', font_size=Pt(13), font_color=MED_GRAY,
         alignment=PP_ALIGN.CENTER)

for i, (label, before, after, delta) in enumerate(metrics):
    ry = Inches(1.5) + i * row_h

    # Before card
    add_rect(s, before_x, ry, card_w, row_h - Inches(0.1), BG_GRAY)
    add_text(s, before_x + Inches(0.2), ry, Inches(1.5), row_h - Inches(0.1),
             label, font_size=Pt(12), font_color=MED_GRAY,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, before_x + Inches(1.8), ry, Inches(2.0), row_h - Inches(0.1),
             before, font_size=Pt(22), font_color=DARK_GRAY, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Arrow
    add_text(s, arrow_x, ry, Inches(1.5), row_h - Inches(0.1),
             '→', font_size=Pt(24), font_color=LINE_GRAY,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # After card
    add_rect(s, after_x, ry, card_w, row_h - Inches(0.1), LIGHT_BLUE)
    add_text(s, after_x + Inches(0.2), ry, Inches(1.5), row_h - Inches(0.1),
             label, font_size=Pt(12), font_color=ACCENT_BLUE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, after_x + Inches(1.8), ry, Inches(2.0), row_h - Inches(0.1),
             after, font_size=Pt(22), font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Delta badge
    is_positive = delta.startswith('+')
    dc = ACCENT_GREEN if is_positive else ACCENT_RED
    add_rect(s, delta_x + Inches(0.1), ry + Inches(0.15),
             Inches(1.2), row_h - Inches(0.35), dc)
    add_text(s, delta_x + Inches(0.1), ry + Inches(0.15),
             Inches(1.2), row_h - Inches(0.35),
             delta, font_size=Pt(16), font_color=WHITE, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ── Summary ──
add_rect(s, LM, Inches(5.5), CONTENT_W, Inches(0.7), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.5), CONTENT_W - Inches(0.6), Inches(0.7),
         '总结：两年数字化转型使营收增长63%，运营效率（库存周转）改善38%，客户满意度（NPS）提升50%',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 战略转型办公室，2026年Q1')
add_page_number(s, 9, 12)
```

---

#### #63 — Icon Grid (4×2 or 3×3)

**Use case**: Capability overview, service catalog, feature grid — each cell with icon placeholder + title + description.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────┬──────────────┬────────────────┤
│  [icon]      │  [icon]      │  [icon]        │
│  Title A     │  Title B     │  Title C       │
│  Description │  Description │  Description   │
├──────────────┼──────────────┼────────────────┤
│  [icon]      │  [icon]      │  [icon]        │
│  Title D     │  Title E     │  Title F       │
│  Description │  Description │  Description   │
├──────────────┴──────────────┴────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '平台六大核心能力',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

items = [
    ('数据智能', '实时数据采集与AI驱动的智能分析，日处理10亿+数据点', ACCENT_BLUE),
    ('用户增长', '全渠道获客引擎，月活用户增长率保持25%+', ACCENT_GREEN),
    ('安全合规', 'ISO 27001认证，满足GDPR与等保三级要求', ACCENT_ORANGE),
    ('开放生态', 'API开放平台，已接入300+合作伙伴与ISV', ACCENT_RED),
    ('智能运维', 'AIOps平台实现99.99%系统可用性', ACCENT_BLUE),
    ('全球部署', '5大洲12个数据中心，端到端延迟<100ms', ACCENT_GREEN),
]

cols = 3
rows = 2
cell_w = CONTENT_W / cols - Inches(0.15)
cell_h = Inches(2.2)
ty = Inches(1.0)

for i, (title, desc, color) in enumerate(items):
    col = i % cols
    row = i // cols
    ix = LM + col * (cell_w + Inches(0.15))
    iy = ty + row * (cell_h + Inches(0.1))

    # Card background
    add_rect(s, ix, iy, cell_w, cell_h, WHITE)
    # Top accent line
    add_rect(s, ix, iy, cell_w, Inches(0.06), color)
    # Icon circle placeholder
    icon_sz = Inches(0.6)
    oval = s.shapes.add_shape(MSO_SHAPE.OVAL, ix + Inches(0.3), iy + Inches(0.25),
                               icon_sz, icon_sz)
    oval.fill.solid()
    oval.fill.fore_color.rgb = color
    oval.line.fill.background()
    _clean_shape(oval)
    add_text(s, ix + Inches(0.3), iy + Inches(0.25), icon_sz, icon_sz,
             title[0], font_size=Pt(18), font_color=WHITE, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Title
    add_text(s, ix + Inches(1.1), iy + Inches(0.25), cell_w - Inches(1.3), Inches(0.4),
             title, font_size=Pt(16), font_color=color, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
    # Description
    add_text(s, ix + Inches(0.3), iy + Inches(1.0), cell_w - Inches(0.6), Inches(1.0),
             desc, font_size=BODY_SIZE, font_color=DARK_GRAY)

add_source(s, 'Source: 产品白皮书，2026年')
add_page_number(s, 5, 12)
```

---

#### #64 — Pie Chart (Simple)

**Use case**: Simple part-of-whole with ≤5 segments — budget allocation, market share, time allocation.

> **v2.0**: Uses BLOCK_ARC native shapes with `inner_ratio=0` for solid pie sectors — only 4 shapes per chart (was 2000+ rect blocks). See Guard Rails Rule 9.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├───────────────────────┬──────────────────────┤
│                       │                      │
│    ┌───────────┐      │  ■ Segment A  42%    │
│    │   PIE     │      │  ■ Segment B  28%    │
│    │ (BLOCK_   │      │  ■ Segment C  18%    │
│    │  ARC ×4)  │      │  ■ Segment D  12%    │
│    └───────────┘      │                      │
│  Insight text box                            │
├───────────────────────┴──────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
from pptx.oxml.ns import qn

s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '2026年IT预算分配',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Pie chart using BLOCK_ARC with inner_ratio=0 (solid sectors) ──
cx, cy = LM + Inches(3.0), Inches(3.2)
radius = Inches(1.8)

segments = [
    (0.42, NAVY, '基础设施'),
    (0.28, ACCENT_BLUE, '应用开发'),
    (0.18, ACCENT_GREEN, '安全合规'),
    (0.12, ACCENT_ORANGE, '创新研发'),
]

cum_deg = 0  # start at top (0° = 12 o'clock, CW)
for pct, color, label in segments:
    sweep = pct * 360
    add_block_arc(s, cx - radius, cy - radius, radius * 2, radius * 2,
                  cum_deg, cum_deg + sweep, 0, color)  # inner_ratio=0 → solid sector
    cum_deg += sweep

# ── Legend (right side) ──
legend_x = LM + Inches(7.0)
for i, (pct, color, label) in enumerate(segments):
    ly = Inches(1.5) + i * Inches(0.9)
    add_rect(s, legend_x, ly + Inches(0.05), Inches(0.3), Inches(0.3), color)
    add_text(s, legend_x + Inches(0.45), ly, Inches(3.5), Inches(0.3),
             f'{label}', font_size=Pt(16), font_color=DARK_GRAY, bold=True)
    add_text(s, legend_x + Inches(0.45), ly + Inches(0.3), Inches(3.5), Inches(0.3),
             f'{int(pct*100)}% — ¥{pct*2.4:.1f}亿',
             font_size=Pt(13), font_color=MED_GRAY)

# ── Insight ──
add_rect(s, LM, Inches(5.3), CONTENT_W, Inches(0.8), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.3), CONTENT_W - Inches(0.6), Inches(0.8),
         '分析：基础设施占比42%符合行业水平，建议将创新研发占比从12%提升至18%以加速数字化转型',
         font_size=BODY_SIZE, font_color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: IT预算委员会，2026年度')
add_page_number(s, 6, 12)
```

---

#### #65 — SWOT Analysis

**Use case**: Classic strategic analysis — Strengths, Weaknesses, Opportunities, Threats in a 2×2 color-coded grid.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────┬───────────────────────┤
│  STRENGTHS (Blue)    │  WEAKNESSES (Orange)  │
│  • Point 1           │  • Point 1            │
│  • Point 2           │  • Point 2            │
├──────────────────────┼───────────────────────┤
│  OPPORTUNITIES (Green)│  THREATS (Red)        │
│  • Point 1           │  • Point 1            │
│  • Point 2           │  • Point 2            │
├──────────────────────┴───────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '市场竞争力SWOT分析',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── 2x2 SWOT grid ──
quadrants = [
    ('S — 优势', ACCENT_BLUE, LIGHT_BLUE,
     ['• 行业领先的技术架构', '• 强大的数据资产（10亿+用户画像）', '• 完善的合作伙伴生态']),
    ('W — 劣势', ACCENT_ORANGE, LIGHT_ORANGE,
     ['• 海外市场品牌知名度不足', '• 中小客户服务能力薄弱', '• 部分产品线老化需升级']),
    ('O — 机会', ACCENT_GREEN, LIGHT_GREEN,
     ['• AI技术赋能新业务场景', '• 东南亚市场快速增长', '• 政策鼓励企业数字化转型']),
    ('T — 威胁', ACCENT_RED, LIGHT_RED,
     ['• 头部竞品加大投入', '• 数据合规要求日益严格', '• 经济下行影响企业IT预算']),
]

cell_w = CONTENT_W / 2 - Inches(0.1)
cell_h = Inches(2.3)
grid_t = Inches(1.0)

for qi, (title, accent, bg, points) in enumerate(quadrants):
    row = qi // 2
    col = qi % 2
    qx = LM + col * (cell_w + Inches(0.15))
    qy = grid_t + row * (cell_h + Inches(0.1))

    # Cell background
    add_rect(s, qx, qy, cell_w, cell_h, bg)
    # Accent top bar
    add_rect(s, qx, qy, cell_w, Inches(0.06), accent)
    # Title
    add_text(s, qx + Inches(0.2), qy + Inches(0.15), cell_w - Inches(0.4), Inches(0.35),
             title, font_size=Pt(16), font_color=accent, bold=True)
    # Points
    add_text(s, qx + Inches(0.2), qy + Inches(0.55), cell_w - Inches(0.4), cell_h - Inches(0.7),
             points, font_size=BODY_SIZE, font_color=DARK_GRAY, line_spacing=Pt(6))

add_source(s, 'Source: 战略规划研讨会，2026年Q1')
add_page_number(s, 4, 12)
```

---

#### #66 — Agenda / Meeting Outline

**Use case**: Meeting agenda with time allocations, speaker assignments — for workshop facilitation, board meetings.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  Time    │ Topic             │ Speaker │ Min  │
│  ────────┼───────────────────┼─────────┼──── │
│  09:00   │ Opening & Context │ CEO     │ 15   │
│  09:15   │ Market Analysis   │ VP Mkt  │ 30   │
│  09:45   │ Product Roadmap   │ CPO     │ 30   │
│  10:15   │ Break             │         │ 15   │
│  10:30   │ Financial Review  │ CFO     │ 30   │
│  11:00   │ Q&A & Next Steps  │ All     │ 30   │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '季度战略评审会议议程',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Agenda header ──
hy = Inches(1.1)
col_defs = [('时间', Inches(1.5)), ('议题', Inches(5.5)), ('主讲人', Inches(2.0)),
            ('时长', Inches(1.2)), ('状态', Inches(1.5))]
hx = LM
for label, w in col_defs:
    add_text(s, hx, hy, w, Inches(0.35),
             label, font_size=Pt(13), font_color=NAVY, bold=True)
    hx += w
add_hline(s, LM, hy + Inches(0.35), CONTENT_W, BLACK, Pt(0.75))

# ── Agenda items ──
items = [
    ('09:00', '开场致辞与战略背景', 'CEO 张总', '15分钟', 'key'),
    ('09:15', '市场竞争格局分析', '市场VP 李总', '30分钟', 'key'),
    ('09:45', '产品路线图与技术规划', 'CPO 王总', '30分钟', 'key'),
    ('10:15', '茶歇', '', '15分钟', 'break'),
    ('10:30', '财务回顾与预算规划', 'CFO 赵总', '30分钟', 'normal'),
    ('11:00', 'Q&A 与下一步行动', '全体参会者', '30分钟', 'normal'),
    ('11:30', '总结与会议闭幕', 'CEO 张总', '15分钟', 'normal'),
]

row_h = Inches(0.6)
for i, (time, topic, speaker, duration, item_type) in enumerate(items):
    ry = Inches(1.6) + i * row_h

    # Row background
    if item_type == 'break':
        add_rect(s, LM, ry, CONTENT_W, row_h, BG_GRAY)
    elif item_type == 'key':
        add_rect(s, LM, ry, Inches(0.06), row_h, ACCENT_BLUE)

    rx = LM
    vals = [(time, Inches(1.5)), (topic, Inches(5.5)), (speaker, Inches(2.0)),
            (duration, Inches(1.2))]
    for val, w in vals:
        fc = MED_GRAY if item_type == 'break' else DARK_GRAY
        bld = item_type == 'key'
        add_text(s, rx, ry, w, row_h,
                 val, font_size=BODY_SIZE, font_color=fc, bold=bld,
                 anchor=MSO_ANCHOR.MIDDLE)
        rx += w

    # Status
    if item_type == 'key':
        add_rect(s, rx + Inches(0.1), ry + Inches(0.12), Inches(1.0), row_h - Inches(0.24),
                 LIGHT_BLUE)
        add_text(s, rx + Inches(0.1), ry, Inches(1.0), row_h,
                 '★ 重点', font_size=Pt(11), font_color=ACCENT_BLUE, bold=True,
                 alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    if i < len(items) - 1:
        add_hline(s, LM, ry + row_h, CONTENT_W, LINE_GRAY, Pt(0.25))

# ── Footer note ──
add_text(s, LM, Inches(5.9), CONTENT_W, Inches(0.3),
         '会议地点：总部大楼28层会议室A | 参会人数：12人 | 会议材料已于3月10日分发',
         font_size=Pt(10), font_color=MED_GRAY)

add_source(s, 'Source: 战略管理部，2026年Q1')
add_page_number(s, 2, 12)
```

---

#### #67 — Value Chain / Horizontal Flow

**Use case**: End-to-end value chain visualization — supply chain, service delivery pipeline, customer journey stages.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│                                              │
│ ┌───────┐  →  ┌───────┐  →  ┌───────┐  →  ┌───────┐  →  ┌───────┐ │
│ │Stage 1│     │Stage 2│     │Stage 3│     │Stage 4│     │Stage 5│ │
│ │ desc  │     │ desc  │     │ desc  │     │ desc  │     │ desc  │ │
│ │ KPI   │     │ KPI   │     │ KPI   │     │ KPI   │     │ KPI   │ │
│ └───────┘     └───────┘     └───────┘     └───────┘     └───────┘ │
│                                              │
│  Insight / bottleneck analysis               │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_action_title(s, '端到端价值链分析 — 从获客到复购')

# ── Value chain stages ──
stages = [
    ('获客', '多渠道投放\n精准触达', 'CAC ¥45', ACCENT_BLUE),
    ('转化', '产品试用\n销售跟进', '转化率 8%', ACCENT_GREEN),
    ('交付', '实施部署\n数据迁移', '周期 21天', ACCENT_ORANGE),
    ('服务', '客户成功\n技术支持', 'NPS 78', NAVY),
    ('复购', '续约管理\n增购推荐', '续约率 92%', ACCENT_BLUE),
]

# ── Dynamic sizing: fill full content width and height ──
n = len(stages)
arrow_w = Inches(0.4)
stage_w = (CW - arrow_w * (n - 1)) / n   # fills entire content width
stage_y = CONTENT_TOP + Inches(0.1)
# Fill down to bottom bar or source area
stage_h = SOURCE_Y - Inches(0.15) - stage_y

for i, (title, desc, kpi, color) in enumerate(stages):
    sx = LM + i * (stage_w + arrow_w)

    # Stage card
    add_rect(s, sx, stage_y, stage_w, stage_h, WHITE)
    add_rect(s, sx, stage_y, stage_w, Inches(0.06), color)  # Top accent
    # Stage number
    add_oval(s, sx + Inches(0.15), stage_y + Inches(0.2), str(i + 1),
             size=Inches(0.4), bg=color)
    # Title
    add_text(s, sx + Inches(0.65), stage_y + Inches(0.2), stage_w - Inches(0.8), Inches(0.4),
             title, font_size=Pt(16), font_color=color, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
    # Description
    desc_h = stage_h - Inches(1.6)  # space for title row + KPI box + padding
    add_text(s, sx + Inches(0.15), stage_y + Inches(0.8), stage_w - Inches(0.3), desc_h,
             desc, font_size=BODY_SIZE, font_color=DARK_GRAY,
             alignment=PP_ALIGN.CENTER)
    # KPI box at bottom
    add_rect(s, sx + Inches(0.1), stage_y + stage_h - Inches(0.7),
             stage_w - Inches(0.2), Inches(0.55), BG_GRAY)
    add_text(s, sx + Inches(0.1), stage_y + stage_h - Inches(0.7),
             stage_w - Inches(0.2), Inches(0.55),
             kpi, font_size=Pt(13), font_color=NAVY, bold=True,
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Arrow between stages
    if i < len(stages) - 1:
        ax = sx + stage_w + Inches(0.05)
        ay = stage_y + stage_h // 2
        add_text(s, ax, ay - Inches(0.15), arrow_w - Inches(0.1), Inches(0.3),
                 '→', font_size=Pt(22), font_color=LINE_GRAY,
                 alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 运营效率分析，2026年Q2')
add_page_number(s, 7, 12)
```

---

#### #68 — Two-Column Image + Text Grid

**Use case**: Visual catalog — 2 rows × 2 columns, each cell with image + title + description. Product showcase, location overview.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────┬───────────────────────┤
│ ┌──────┐ Title A     │ ┌──────┐ Title B      │
│ │IMAGE │ Description │ │IMAGE │ Description  │
│ └──────┘             │ └──────┘              │
├──────────────────────┼───────────────────────┤
│ ┌──────┐ Title C     │ ┌──────┐ Title D      │
│ │IMAGE │ Description │ │IMAGE │ Description  │
│ └──────┘             │ └──────┘              │
├──────────────────────┴───────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '四大区域办公室概览',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

items = [
    ('上海总部', '员工800+人，研发与运营中心，覆盖华东市场', ACCENT_BLUE),
    ('深圳分部', '员工350+人，技术创新中心，负责华南及出海业务', ACCENT_GREEN),
    ('北京分部', '员工200+人，政府关系与企业客户中心', ACCENT_ORANGE),
    ('新加坡办事处', '员工50+人，东南亚市场拓展前哨站', ACCENT_RED),
]

cell_w = CONTENT_W / 2 - Inches(0.15)
cell_h = Inches(2.2)
img_w = Inches(2.8)
img_h = Inches(1.8)
ty = Inches(1.0)

for i, (title, desc, color) in enumerate(items):
    col = i % 2
    row = i // 2
    cx = LM + col * (cell_w + Inches(0.15))
    cy = ty + row * (cell_h + Inches(0.1))

    # Card background
    add_rect(s, cx, cy, cell_w, cell_h, WHITE)
    # Left: image placeholder
    add_image_placeholder(s, cx + Inches(0.15), cy + Inches(0.15), img_w, img_h,
                          f'{title}办公室照片')
    # Right: text
    tx = cx + img_w + Inches(0.3)
    tw = cell_w - img_w - Inches(0.45)
    add_rect(s, tx - Inches(0.05), cy, Inches(0.06), cell_h, color)
    add_text(s, tx + Inches(0.1), cy + Inches(0.2), tw, Inches(0.35),
             title, font_size=Pt(16), font_color=color, bold=True)
    add_text(s, tx + Inches(0.1), cy + Inches(0.6), tw, Inches(1.2),
             desc, font_size=BODY_SIZE, font_color=DARK_GRAY)

add_source(s, 'Source: 人力资源部，2026年3月')
add_page_number(s, 8, 12)
```

---

#### #69 — Numbered List with Side Panel

**Use case**: Key recommendations or findings with a highlighted side panel — consulting recommendations, audit findings.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├────────────────────────┬─────────────────────┤
│                        │                     │
│  1  Recommendation A   │  ┌───────────────┐  │
│     Detail text...     │  │ HIGHLIGHT     │  │
│                        │  │ PANEL         │  │
│  2  Recommendation B   │  │               │  │
│     Detail text...     │  │ Key metric    │  │
│                        │  │ or quote      │  │
│  3  Recommendation C   │  │               │  │
│     Detail text...     │  └───────────────┘  │
│                        │                     │
├────────────────────────┴─────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '五项核心建议',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Left: numbered recommendations (65%) ──
left_w = Inches(7.5)
recs = [
    ('构建数据中台', '统一数据标准与治理体系，打通15个业务系统数据孤岛'),
    ('升级客户体验', '引入AI客服+智能推荐，目标NPS提升至85+'),
    ('拓展海外市场', '以东南亚为突破口，3年内海外营收占比达20%'),
    ('强化人才体系', '实施"数字化人才倍增计划"，年培养200+复合型人才'),
    ('优化成本结构', '通过自动化+云原生架构，IT运营成本降低30%'),
]

ty = Inches(1.0)
for i, (title, desc) in enumerate(recs):
    ry = ty + i * Inches(0.85)
    add_oval(s, LM, ry + Inches(0.05), str(i + 1), bg=NAVY)
    add_text(s, LM + Inches(0.6), ry, left_w - Inches(0.6), Inches(0.3),
             title, font_size=Pt(15), font_color=NAVY, bold=True)
    add_text(s, LM + Inches(0.6), ry + Inches(0.35), left_w - Inches(0.6), Inches(0.4),
             desc, font_size=BODY_SIZE, font_color=DARK_GRAY)
    if i < len(recs) - 1:
        add_hline(s, LM + Inches(0.6), ry + Inches(0.8), left_w - Inches(0.8),
                  LINE_GRAY, Pt(0.25))

# ── Right: highlight panel (35%) ──
rx = LM + left_w + Inches(0.3)
rw = CONTENT_W - left_w - Inches(0.3)
panel_y = Inches(1.0)
panel_h = Inches(4.8)

add_rect(s, rx, panel_y, rw, panel_h, NAVY)
add_text(s, rx + Inches(0.3), panel_y + Inches(0.3), rw - Inches(0.6), Inches(0.3),
         '预期回报', font_size=Pt(14), font_color=RGBColor(0xCC, 0xCC, 0xCC))
add_text(s, rx + Inches(0.3), panel_y + Inches(0.8), rw - Inches(0.6), Inches(0.6),
         '+¥3.2亿', font_size=Pt(36), font_color=WHITE, bold=True,
         alignment=PP_ALIGN.CENTER)
add_text(s, rx + Inches(0.3), panel_y + Inches(1.5), rw - Inches(0.6), Inches(0.3),
         '年化增量营收', font_size=Pt(13), font_color=RGBColor(0xCC, 0xCC, 0xCC),
         alignment=PP_ALIGN.CENTER)

# Divider
add_hline(s, rx + Inches(0.3), panel_y + Inches(2.1), rw - Inches(0.6),
          RGBColor(0x33, 0x44, 0x55), Pt(0.5))

# Additional metrics
panel_metrics = [
    ('投资回收期', '18个月'),
    ('三年ROI', '320%'),
    ('风险等级', '中低'),
]
for mi, (label, val) in enumerate(panel_metrics):
    my = panel_y + Inches(2.4) + mi * Inches(0.7)
    add_text(s, rx + Inches(0.3), my, rw - Inches(0.6), Inches(0.3),
             label, font_size=Pt(11), font_color=RGBColor(0xAA, 0xAA, 0xAA))
    add_text(s, rx + Inches(0.3), my + Inches(0.3), rw - Inches(0.6), Inches(0.3),
             val, font_size=Pt(18), font_color=WHITE, bold=True)

add_source(s, 'Source: 战略咨询项目终期报告，2026年Q1')
add_page_number(s, 11, 12)
```

---

#### #70 — Stacked Area Chart

**Use case**: Cumulative trends over time — market composition, revenue streams, resource allocation showing both individual and total trends.

```
┌──────────────────────────────────────────────┐
│ [Action Title — full width, NAVY bg]         │
├──────────────────────────────────────────────┤
│  Y ─                                         │
│     ████████████████████████████   Total      │
│     ████████████████████████  Series C        │
│     ██████████████████  Series B              │
│     ██████████  Series A                      │
│  0 ──────────────────────────────────        │
│     2020  2021  2022  2023  2024  2025       │
│  Takeaway text...                            │
├──────────────────────────────────────────────┤
│ Source | Page N/Total                         │
└──────────────────────────────────────────────┘
```

```python
s = prs.slides.add_slide(prs.slide_layouts[6])

# ── Title bar ──
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.75), NAVY)
add_text(s, LM, Inches(0), CONTENT_W, Inches(0.75),
         '营收构成趋势（2021-2026E）',
         font_size=TITLE_SIZE, font_color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_hline(s, LM, Inches(0.75), CONTENT_W, BLACK, Pt(0.5))

# ── Chart setup ──
chart_l = LM + Inches(1.0)
chart_b = Inches(4.8)
chart_t = Inches(1.3)
chart_w = Inches(9.5)
chart_h = chart_b - chart_t
max_val = 12  # ¥12亿

years = ['2021', '2022', '2023', '2024', '2025', '2026E']
# Series data (stacked from bottom): values in 亿
series_data = [
    ('线上直营', [1.5, 2.0, 2.8, 3.5, 4.2, 5.0], NAVY),
    ('经销渠道', [2.0, 2.3, 2.5, 2.8, 3.0, 3.2], ACCENT_BLUE),
    ('企业客户', [0.5, 0.8, 1.2, 1.5, 1.8, 2.5], ACCENT_GREEN),
]

n_pts = len(years)
col_w = chart_w // n_pts

# Draw stacked columns for each year
for yi in range(n_pts):
    cumulative = 0
    for si, (name, values, color) in enumerate(series_data):
        val = values[yi]
        bar_h = int(chart_h * val / max_val)
        base_h = int(chart_h * cumulative / max_val)
        bx = chart_l + int(chart_w * yi / n_pts)
        by = chart_b - base_h - bar_h
        add_rect(s, bx + Inches(0.05), by, col_w - Inches(0.1), bar_h, color)
        cumulative += val

    # Total label above
    total = sum(s_data[1][yi] for s_data in series_data)
    total_h = int(chart_h * total / max_val)
    add_text(s, chart_l + int(chart_w * yi / n_pts), chart_b - total_h - Inches(0.25),
             col_w, Inches(0.2),
             f'¥{total:.1f}亿', font_size=Pt(10), font_color=DARK_GRAY, bold=True,
             alignment=PP_ALIGN.CENTER)

    # Year label
    add_text(s, chart_l + int(chart_w * yi / n_pts), chart_b + Inches(0.05),
             col_w, Inches(0.2),
             years[yi], font_size=Pt(10), font_color=MED_GRAY,
             alignment=PP_ALIGN.CENTER)

# Baseline
add_hline(s, chart_l, chart_b, chart_w, BLACK, Pt(0.5))

# Y-axis labels
for i in range(5):
    val = max_val * i / 4
    yy = chart_b - int(chart_h * i / 4)
    add_text(s, LM, yy - Inches(0.1), Inches(0.8), Inches(0.2),
             f'¥{int(val)}亿', font_size=Pt(9), font_color=MED_GRAY,
             alignment=PP_ALIGN.RIGHT)
    if i > 0:
        add_hline(s, chart_l, yy, chart_w, RGBColor(0xE8, 0xE8, 0xE8), Pt(0.25))

# ── Legend ──
legend_x = LM + Inches(9.0)
for i, (name, _, color) in enumerate(series_data):
    ly = Inches(1.5) + i * Inches(0.4)
    add_rect(s, legend_x, ly + Inches(0.05), Inches(0.25), Inches(0.2), color)
    add_text(s, legend_x + Inches(0.35), ly, Inches(2.0), Inches(0.3),
             name, font_size=Pt(11), font_color=DARK_GRAY)

# ── Takeaway ──
add_rect(s, LM, Inches(5.2), CONTENT_W, Inches(0.8), BG_GRAY)
add_text(s, LM + Inches(0.3), Inches(5.2), Inches(1.5), Inches(0.8),
         '趋势分析', font_size=BODY_SIZE, font_color=NAVY, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, LM + Inches(2.0), Inches(5.2), CONTENT_W - Inches(2.3), Inches(0.8),
         '总营收6年CAGR达19%，线上直营增速最快（22% CAGR），企业客户板块2026年有望成为第二大收入来源',
         font_size=BODY_SIZE, font_color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)

add_source(s, 'Source: 财务部 & 战略规划部，2026年')
add_page_number(s, 5, 12)
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
             alignment=PP_ALIGN.LEFT, ea_font='KaiTi', anchor=MSO_ANCHOR.TOP,
             line_spacing=Pt(6)):
    """Unified text helper. Pass str for single line, list for multi-line."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    anchor_map = {MSO_ANCHOR.MIDDLE: 'ctr', MSO_ANCHOR.BOTTOM: 'b', MSO_ANCHOR.TOP: 't'}
    bodyPr.set('anchor', anchor_map.get(anchor, 't'))
    for attr in ['lIns','tIns','rIns','bIns']:
        bodyPr.set(attr, '45720')
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = font_size
        p.font.name = font_name
        p.font.color.rgb = font_color
        p.font.bold = bold
        p.alignment = alignment
        p.space_before = line_spacing if i > 0 else Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 0.93 if font_size.pt >= 18 else Pt(font_size.pt * 1.35)  # Titles (>=18pt): 0.93x multiple; Body: 135% fixed Pt
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


def add_image_placeholder(slide, left, top, width, height, label='Image'):
    """Draw a gray placeholder box with crosshair + label for image positions.
    Users replace these with real images after PPT generation."""
    PLACEHOLDER_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
    # Background rect
    rect = add_rect(slide, left, top, width, height, PLACEHOLDER_GRAY)
    # Horizontal center line
    add_hline(slide, left, top + height // 2, width, RGBColor(0xBB, 0xBB, 0xBB), Pt(0.5))
    # Vertical center line as thin rect
    vw = Pt(0.5)
    add_rect(slide, left + width // 2 - vw // 2, top, vw, height, RGBColor(0xBB, 0xBB, 0xBB))
    # Label
    add_text(slide, left, top + height // 2 - Inches(0.2), width, Inches(0.4),
             f'[ {label} ]', font_size=Pt(12), font_color=RGBColor(0x99, 0x99, 0x99),
             alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return rect
```

---

## Common Issues & Solutions

### Problem 1: PPT Won't Open / "File Needs Repair"

**Cause**: Shapes or connectors carry `<p:style>` with `effectRef idx="2"`, referencing theme effects (shadows/3D)

**Solution** (three-layer defense):
1. **Never use connectors** — use `add_hline()` (thin rectangle) instead of `add_connector()`
2. **Inline cleanup** — every `add_rect()` and `add_oval()` calls `_clean_shape()` to remove `p:style`
3. **Post-save cleanup** — `full_cleanup()` removes ALL `<p:style>` from every slide XML + theme effects

### Problem 2: Text Not Displaying Correctly in PowerPoint

**Cause**: Chinese characters rendered as English font instead of KaiTi

**Solution**:
- Use `set_ea_font(run, 'KaiTi')` in every paragraph with Chinese text
- Call it inside the loop that creates runs:
  ```python
  for run in p.runs:
      set_ea_font(run, 'KaiTi')
  ```

### Problem 3: Font Sizes Inconsistent Across Slides

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

### Problem 4: Columns/Lists Not Aligning Vertically

**Cause**: Mixing different line spacing or not accounting for text height

**Solution**:
- Use consistent `line_spacing=Pt(N)` in `add_text()` calls
- Calculate row heights in tables based on actual text size:
  - For 14pt text with spacing: use 1.0" height minimum
  - For lists with bullets: use 0.35" height per line + 8pt spacing
- Test by saving and opening in PowerPoint to verify alignment

### Problem 5: Chinese Multi-Line Text Overlapping (v1.5.0 Fix)

**Cause**: `add_text()` only set `space_before` (paragraph spacing) but did NOT set `p.line_spacing` (the actual line height / `<a:lnSpc>` in OOXML). When Chinese text wraps within a paragraph, lines overlap because PowerPoint has no explicit line height to follow.

**Solution** (fixed in v1.5.0, refined in v1.10.3):
- `add_text()` sets `p.line_spacing` for every paragraph with a **two-tier strategy**:
  - **Titles (font_size ≥ 18pt)**: `p.line_spacing = 0.93` — multiple spacing for tighter, more professional title rendering
  - **Body text (font_size < 18pt)**: `p.line_spacing = Pt(font_size.pt * 1.35)` — fixed Pt spacing to prevent CJK overlap
- Title multiple spacing (`0.93`) maps to `<a:lnSpc><a:spcPct val="93000"/>` in OOXML
- Body fixed spacing maps to `<a:lnSpc><a:spcPts>` in OOXML

### Problem 6: Content Overflowing Container Boxes (v1.9.0)

**Cause**: Text placed inside a colored rectangle (`add_rect`) with identical coordinates to the box itself, so text runs to the very edge and may visually overflow, especially with CJK characters that have wider natural widths.

**Solution**: Always inset text boxes by at least 0.15" on left/right within their container:
```python
# Box at (box_x, box_y, box_w, box_h)
add_rect(s, box_x, box_y, box_w, box_h, BG_GRAY)
# Text inset by 0.3" on each side
add_text(s, box_x + Inches(0.3), box_y, box_w - Inches(0.6), box_h, text, ...)
```
For tight spaces, reduce font_size by 1-2pt rather than reducing padding below 0.15".

### Problem 7: Chart Legend Colors Mismatch (v1.9.0)

**Cause**: Legend text uses Unicode "■" character in black, while actual chart bars/areas use NAVY/ACCENT_RED/ACCENT_GREEN — creating confusion about which color maps to which series.

**Solution**: Replace text-only legends with `add_rect()` color squares. See **Production Guard Rails Rule 4** for the standard pattern. Each legend item = colored square (0.15" × 0.15") + label text.

### Problem 8: Inconsistent Title Bar Styles (v1.9.0)

**Cause**: Mixing `add_navy_title_bar()` (navy background + white text) and `add_action_title()` (white background + black text + underline) on different slides within the same deck, creating visual inconsistency.

**Solution**: Use `add_action_title()` exclusively for all content slides. Remove `add_navy_title_bar()` usage. See **Production Guard Rails Rule 5**.

**Migration**: When converting `add_navy_title_bar()` → `add_action_title()`, adjust content start position from `Inches(1.0)` to `Inches(1.25)` since `add_action_title()` occupies slightly more vertical space.

### Problem 9: Axis Labels Off-Center in Matrix Charts (v1.9.0)

**Cause**: Y-axis label positioned at a fixed left offset, X-axis label at a fixed bottom offset — neither centered on the actual grid dimensions when grid position/size changes.

**Solution**: Calculate axis label positions from actual grid dimensions. See **Production Guard Rails Rule 6** for the centering formula.

### Problem 10: Bottom Whitespace Under Charts (v1.9.0)

**Cause**: Chart height calculated independently of the bottom summary bar position, leaving 0.5-1.0" of dead space between chart bottom and the summary bar.

**Solution**: Either extend chart height to fill the gap or move the bottom bar up. Target maximum 0.3" gap. See **Production Guard Rails Rule 3**.

### Problem 11: Cover Slide Title/Subtitle Overlap (v1.10.4)

**Cause**: Cover slide title textbox height is fixed (e.g. `Inches(1.0)`), but when the title contains `\n` (multi-line), two lines of 44pt text require ~1.66" of vertical space. The subtitle is positioned at a fixed `y` coordinate (e.g. `Inches(3.5)`), so the title overflows its textbox and visually overlaps the subtitle.

**Solution**: Calculate title height **dynamically** based on line count, then position subtitle/author/date relative to title bottom:

```python
# ✅ CORRECT: Dynamic title height on cover slides
lines = title.split('\n') if isinstance(title, str) else title
n_lines = len(lines) if isinstance(lines, list) else title.count('\n') + 1
title_h = Inches(0.8 + 0.65 * max(n_lines - 1, 0))  # ~0.65" per extra line

add_text(s, Inches(1), Inches(1.2), Inches(11), title_h,
         title, font_size=Pt(44), font_color=NAVY, bold=True, font_name='Georgia')

# Position subtitle BELOW the title dynamically
sub_y = Inches(1.2) + title_h + Inches(0.3)
if subtitle:
    add_text(s, Inches(1), sub_y, Inches(11), Inches(0.8),
             subtitle, font_size=Pt(24), font_color=DARK_GRAY)
    sub_y += Inches(1.0)
```

**Rule**: Never use fixed `y` coordinates for cover slide elements below the title. Always compute positions relative to title bottom.

### Problem 12: Action Title Text Not Flush Against Separator Line (v1.10.4)

**Cause**: `add_action_title()` uses `anchor=MSO_ANCHOR.MIDDLE` (vertical center alignment), so single-line titles float in the middle of the title bar, leaving a visible gap between the text baseline and the separator line at `Inches(1.05)`.

**Solution**: Change the text anchor from `MSO_ANCHOR.MIDDLE` to **`MSO_ANCHOR.BOTTOM`** so the text sits flush against the bottom of the textbox, right above the separator line:

```python
# ✅ CORRECT: Bottom-anchored action title — text sits flush against separator
def add_action_title(slide, text, title_size=Pt(22)):
    add_text(s, Inches(0.8), Inches(0.15), Inches(11.7), Inches(0.9), text,
             font_size=title_size, font_color=BLACK, bold=True, font_name='Georgia',
             anchor=MSO_ANCHOR.BOTTOM)  # ← BOTTOM, not MIDDLE
    add_hline(s, Inches(0.8), Inches(1.05), Inches(11.7), BLACK, Pt(0.5))
```

### Problem 13: Checklist Rows Overflowing Page Bottom (v1.10.4)

**Cause**: `#61 Checklist / Status` uses a fixed `row_h = Inches(0.55)` or `Inches(0.85)`. With 7+ rows, total height = `0.85 * 7 = 5.95"`, starting from `~Inches(1.45)` extends to `Inches(7.4)` — exceeding page height (7.5") and overlapping with source/page number areas.

**Solution**: Calculate `row_h` dynamically based on available vertical space, and switch to smaller font when rows are tight:

```python
# ✅ CORRECT: Dynamic row height for checklist
bottom_limit = BOTTOM_BAR_Y - Inches(0.1) if bottom_bar else SOURCE_Y - Inches(0.05)
available_h = bottom_limit - (header_y + Inches(0.5))
row_h = min(Inches(0.85), available_h / max(len(rows), 1))  # cap at 0.85" max

# Use smaller font when rows are tight
row_font = SMALL_SIZE if row_h < Inches(0.65) else BODY_SIZE
```

**Rule**: For any layout with a variable number of rows/items, ALWAYS compute item height dynamically: `item_h = min(MAX_ITEM_H, available_space / n_items)`. Never use a fixed height that assumes a specific item count.

### Problem 14: Value Chain Stages Not Filling Content Area (v1.10.4)

**Cause**: `#67 Value Chain` uses a fixed `stage_w = Inches(2.0)` and centers stages. With 4 stages, total width = `4*2.0 + 3*0.4 = 9.2"`, centered in `CW=11.73"` leaves ~1.27" empty on each side. Stage height is also fixed at `Inches(2.8)`, leaving ~3.3" of dead space below.

**Solution**: Calculate stage width and height dynamically to fill the entire content area:

```python
# ✅ CORRECT: Dynamic stage sizing — fills full content width and height
n = len(stages)
arrow_w = Inches(0.35)
stage_w = (CW - arrow_w * (n - 1)) / n  # fill entire content width
stage_y = CONTENT_TOP + Inches(0.1)
# Fill down to bottom_bar or source area
stage_h = (BOTTOM_BAR_Y - Inches(0.15) - stage_y) if bottom_bar else (SOURCE_Y - Inches(0.15) - stage_y)
```

**Rule**: For layouts with N equally-sized elements arranged horizontally, compute width as `(CW - gap * (N-1)) / N`, not a fixed `Inches(2.0)`. For vertical space, fill down to the bottom bar or source line.

### Problem 15: Closing Slide Bottom Line Too Short (v1.10.4)

**Cause**: The closing slide's bottom decorative line uses a fixed width like `Inches(3)`, which only spans a small portion of the slide — looking unfinished and asymmetric.

**Solution**: Use `CW` (content width) as the line width, and `LM` (left margin) as the starting x, so the line spans the full content area:

```python
# ❌ WRONG: Fixed short width
add_hline(s, Inches(1), Inches(6.8), Inches(3), NAVY, Pt(2))

# ✅ CORRECT: Full content width
add_hline(s, LM, Inches(6.8), CW, NAVY, Pt(2))
```

**Rule**: Decorative horizontal lines on structural slides (cover, closing) should span the full content width (`CW`), not arbitrary fixed widths.

### Problem 16: Donut/Pie Charts Made of Hundreds of Tiny Rect Blocks (v2.0)

**Cause**: Using nested loops with `math.cos/sin` + `add_rect()` to approximate circles/arcs creates 100-2800 shapes per chart. This inflates PPTX file size by 60-80%, causes generation timeouts (2+ minutes), and produces visible gaps and jagged edges.

**Solution**: Use `BLOCK_ARC` preset shapes with XML `adj` parameter control. Each segment = 1 shape:

```python
# ❌ WRONG: Hundreds of tiny blocks (slow, large file, jagged)
for deg in range(0, 360, 2):
    rad = math.radians(deg)
    for r in range(0, int(radius), int(block_sz)):
        bx = cx + int(r * math.cos(rad))
        add_rect(s, bx, by, block_sz, block_sz, color)  # → 2000+ shapes!

# ✅ CORRECT: One BLOCK_ARC per segment (fast, clean, 4 shapes total)
add_block_arc(s, cx - r, cy - r, r * 2, r * 2,
              start_deg, end_deg, inner_ratio, color)
```

See **Production Guard Rails Rule 9** for the complete `add_block_arc()` helper and usage patterns.

### Problem 17: Gauge Arc Renders Vertically Instead of Horizontally (v2.0)

**Cause**: Using math convention angles (0°=right, 90°=top, CCW) instead of PPT convention (0°=top, 90°=right, CW). A "horizontal rainbow" gauge using `math.radians(0)` to `math.radians(180)` renders as a **vertical** arc in PowerPoint because the coordinate systems are incompatible.

**Solution**: Use PPT's native clockwise-from-12-o'clock coordinate system directly:

```python
# PPT angle mapping for horizontal rainbow (opening upward ⌢):
#   Left  = 270° PPT
#   Top   = 0° (or 360°) PPT
#   Right = 90° PPT
# Total sweep: 270° → 0° → 90° = 180° clockwise

# ❌ WRONG: Math convention angles
ppt_angle = (90 - math_angle) % 360  # Fragile, error-prone conversion

# ✅ CORRECT: Think directly in PPT coordinates
ppt_cum = 270  # start at left
for pct, color in gauge_segs:
    sweep = pct * 180
    add_block_arc(s, ..., ppt_cum % 360, (ppt_cum + sweep) % 360, ...)
    ppt_cum += sweep
```

### Problem 18: Donut Center Text Unreadable Against Colored Ring (v2.0)

**Cause**: Center labels (e.g., "¥7,013亿", "总营收") use NAVY or MED_GRAY font color, which is invisible or low-contrast against the colored BLOCK_ARC ring segments behind them.

**Solution**: Use **WHITE** for center labels inside donut charts. The colored ring provides enough contrast:

```python
# ❌ WRONG: Navy text on navy/blue ring — invisible
add_text(s, ..., '¥7,013亿', font_color=NAVY, ...)

# ✅ CORRECT: White text, visible against any ring color
add_text(s, ..., '¥7,013亿', font_color=WHITE, bold=True,
         font_name='Georgia', ...)
add_text(s, ..., '总营收', font_color=WHITE, ...)
```

### Problem 19: Chart Elements Overlapping Title Bar — Body Content Too High (v2.0)

**Cause**: Chart area `chart_top` set to `Inches(1.0)` or `Inches(1.2)`, which places chart elements above the title separator line at `Inches(1.05)`. Applies to waterfall charts, line charts, bar charts, and other data visualization layouts.

**Solution**: All chart/content body areas must start at or below `Inches(1.3)`:

```python
# ❌ WRONG: Content starts above title separator
chart_top = Inches(1.0)   # overlaps title!

# ✅ CORRECT: Content respects title bar space
chart_top = Inches(1.3)   # safe start below title + separator + gap
```

**Rule**: Apply `Inches(1.3)` as minimum content start for ALL content slides (charts, tables, text blocks). The title bar occupies `Inches(0) → Inches(1.05)`, and `Inches(0.25)` gap is mandatory.

### Problem 20: Waterfall Chart Connector Lines Look Like Dots (v2.0)

**Cause**: Connector lines between waterfall bars are drawn using `add_hline()` with very short length (< 0.1"), making them appear as small dots instead of visible connection lines.

**Solution**: Ensure connector lines span the full gap between bars, and use consistent thin styling:

```python
# Between bar[i] and bar[i+1]:
connector_x = bx + bar_w  # start at right edge of current bar
connector_w = gap          # span the full gap to next bar
connector_y = running_top  # at the running total level
add_hline(s, connector_x, connector_y, connector_w, LINE_GRAY, Pt(0.75))
```

**Rule**: Waterfall connector lines must have `width >= gap_between_bars` and use `Pt(0.75)` line weight for visibility.

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

## Best Practices

1. **Always start from scratch** - Don't try to edit existing .pptx files with python-pptx; regenerate
2. **Test early** - Save and open in PowerPoint after every 2-3 slides to catch issues
3. **Use constants** - Define all colors, sizes, positions as named constants at the top
4. **Keep code DRY** - Use helper functions like `add_text()`, `add_hline()`, `add_oval()`, etc.
5. **Never use connectors** - Always draw lines as thin rectangles via `add_hline()`
6. **Validate XML** - After `full_cleanup()`, verify zero `p:style` and zero shadows remain
7. **Document decisions** - Comment code explaining why specific colors/sizes are chosen
8. **Version control** - Save Python generation script alongside .pptx output

### Code Efficiency Guidelines (v1.9)

The generated Python scripts can become 500+ lines for 15+ slide presentations. Follow these patterns to reduce code size, improve readability, and minimize LLM token consumption:

#### 1. Extract Repeated Layout Constants

Instead of recalculating positions inline, define named constants at the top:

```python
# ✅ Define once, reuse everywhere
CONTENT_TOP = Inches(1.25)      # Below action title
CONTENT_BOTTOM = Inches(6.9)    # Above source line
BOTTOM_BAR_Y = Inches(6.2)     # Standard bottom bar position
BOTTOM_BAR_H = Inches(0.65)    # Standard bottom bar height
BOTTOM_BAR_GAP = Inches(0.15)  # Minimum gap above bottom bar
LEGEND_Y = Inches(1.15)        # Standard legend line Y position
LEGEND_SQ = Inches(0.15)       # Legend color square size
```

#### 2. Use Helper Functions for Repeated Patterns

When the same visual pattern appears across multiple slides, create a reusable function:

```python
# ✅ Reusable legend builder
def add_color_legend(slide, x, y, items):
    """items: list of (label, color) tuples"""
    cx = x
    for label, color in items:
        add_rect(slide, cx, y + Inches(0.03), LEGEND_SQ, LEGEND_SQ, color)
        add_text(slide, cx + Inches(0.2), y, Inches(1.2), Inches(0.25),
                 label, font_size=Pt(10), font_color=MED_GRAY)
        cx += Inches(0.2) + Inches(len(label) * 0.12 + 0.3)  # dynamic spacing

# ✅ Reusable bottom bar
def add_bottom_bar(slide, label, text, y=BOTTOM_BAR_Y):
    add_rect(slide, LM, y, CW, BOTTOM_BAR_H, BG_GRAY)
    add_text(slide, LM + Inches(0.3), y, Inches(1.5), BOTTOM_BAR_H,
             label, font_size=BODY_SIZE, font_color=NAVY, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, LM + Inches(2), y, CW - Inches(2.3), BOTTOM_BAR_H,
             text, font_size=BODY_SIZE, font_color=DARK_GRAY,
             anchor=MSO_ANCHOR.MIDDLE)
```

#### 3. Use Short Variable Names (Approved Abbreviations)

To keep code compact, use these standard abbreviations consistently:

| Short | Full Name | Example |
|-------|-----------|---------|
| `s` | slide | `s = prs.slides.add_slide(BL)` |
| `at` | add_text | `at(s, x, y, w, h, text, ...)` |
| `ar` | add_rect | `ar(s, x, y, w, h, color)` |
| `ahl` | add_hline | `ahl(s, x, y, length, color)` |
| `ao` | add_oval | `ao(s, x, y, label)` |
| `aat` | add_action_title | `aat(s, 'Title Text')` |
| `asrc` | add_source | `asrc(s, 'Source: ...')` |
| `apn` | add_page_number | `apn(s, 5, 19)` |

#### 4. Batch Data as Lists of Tuples

Instead of separate variables for each element, organize data as iteration-ready structures:

```python
# ❌ WRONG: Separate variables
card1_title = 'Agent化'; card1_value = '95%'; card1_color = ACCENT_RED
card2_title = '架构层'; card2_value = '88%'; card2_color = ACCENT_ORANGE
# ... 8 more lines

# ✅ CORRECT: Compact data structure
cards = [
    ('Agent化', '95%', ACCENT_RED),
    ('架构层', '88%', ACCENT_ORANGE),
    ('安全危机', '82%', ACCENT_ORANGE),
]
for i, (title, value, color) in enumerate(cards):
    x = LM + i * (card_w + gap)
    # ... render card
```

#### 5. Page Number Auto-Tracking

Use a global counter instead of hardcoding page numbers:

```python
TT = 19  # Total slide count (set after planning)
_pn = 0  # Auto-incrementing page counter

def next_slide(prs):
    global _pn
    _pn += 1
    return prs.slides.add_slide(BL)

# Usage:
s = next_slide(prs)
aat(s, 'Title Here')
# ... content ...
asrc(s, 'Source: ...')
apn(s, _pn, TT)
```

This eliminates the need to manually renumber all slides when inserting or removing a page.

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

## Channel Delivery (v1.10)

When users interact via a **messaging channel** (Feishu/飞书, Telegram, WhatsApp, Discord, Slack, etc.), the generated PPTX file **MUST** be sent back to the chat — not just saved to disk.

### Why This Matters

Users on mobile or messaging channels cannot access server file paths. Saving a file to `./output/` is invisible to them. The file must be delivered through the same channel the user is talking on.

### Delivery Method

After `prs.save(outpath)` and `full_cleanup(outpath)`, use the OpenClaw media pipeline to send the file:

```bash
openclaw message send --media <outpath> --message "✅ PPT generated — <N> slides, <size> bytes"
```

### Python Helper

```python
import subprocess, shutil

def deliver_to_channel(outpath, slide_count):
    """Send generated PPTX back to user's chat channel via OpenClaw media pipeline.
    Falls back gracefully if not running in a channel context."""
    if not shutil.which('openclaw'):
        print(f'[deliver] openclaw CLI not found, skipping channel delivery')
        print(f'[deliver] File saved locally: {outpath}')
        return False
    
    size_kb = os.path.getsize(outpath) / 1024
    caption = f'✅ PPT generated — {slide_count} slides, {size_kb:.0f} KB'
    
    try:
        result = subprocess.run(
            ['openclaw', 'message', 'send',
             '--media', outpath,
             '--message', caption],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0:
            print(f'[deliver] Sent to channel: {outpath}')
            return True
        else:
            print(f'[deliver] Channel send failed: {result.stderr}')
            print(f'[deliver] File saved locally: {outpath}')
            return False
    except Exception as e:
        print(f'[deliver] Error: {e}')
        print(f'[deliver] File saved locally: {outpath}')
        return False
```

### Integration with Generation Flow

The complete post-generation sequence is:

```python
# 1. Save
prs.save(outpath)

# 2. Clean (mandatory)
full_cleanup(outpath)

# 3. Deliver to channel (if available)
slide_count = len(prs.slides)
deliver_to_channel(outpath, slide_count)

# 4. Confirm
print(f'Created: {outpath} ({os.path.getsize(outpath):,} bytes)')
```

### Rules

1. **Always attempt delivery** — after every successful generation, call `deliver_to_channel()`
2. **Graceful fallback** — if `openclaw` CLI is not available (e.g., running in IDE or CI), skip silently and print the local path
3. **Caption required** — always include slide count and file size so the user knows what they received
4. **No duplicate sends** — call `deliver_to_channel()` exactly once per generation
5. **File type** — `.pptx` is classified as "document" in OpenClaw's media pipeline (max 100MB), well within limits for any presentation

### Channel-Specific Notes

| Channel | File Support | Max Size | Notes |
|---------|-------------|----------|-------|
| Feishu/飞书 | ✅ Document | 100MB | Renders as downloadable file card |
| Telegram | ✅ Document | 100MB | Shows as file attachment |
| WhatsApp | ✅ Document | 100MB | Delivered as document message |
| Discord | ✅ Attachment | 100MB | Appears in chat as file |
| Slack | ✅ File | 100MB | Shared as file snippet |
| Signal | ✅ Attachment | 100MB | Sent as generic attachment |
| Others | ✅ Document | 100MB | All OpenClaw channels support document type |

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 2.0.0 | 2026-03-19 | **BLOCK_ARC Chart Engine**: Donut (#48), Pie (#64), and Gauge (#55) charts rewritten from hundreds of `add_rect()` blocks to native BLOCK_ARC shapes — 3-4 shapes per chart instead of 100-2800. File size reduced 60-80%. New `add_block_arc()` helper function with PPT coordinate system documentation. **Guard Rail Rule 9**: mandatory BLOCK_ARC for all circular charts. **5 new Common Issues** (Problems 16-20): rect-block charts, vertical gauge, unreadable donut center text, body content above title bar, waterfall connector dots. Donut center labels changed to WHITE for contrast. Gauge uses correct PPT angle mapping (270°→0°→90° for horizontal rainbow). |
| 1.10.4 | 2026-03-19 | **5 New Bug Fixes + Guard Rail Rule 8**: (1) Cover slide title/subtitle overlap — dynamic title height from line count; (2) Action title anchor changed to `MSO_ANCHOR.BOTTOM` for flush separator alignment; (3) Checklist `#61` dynamic `row_h` prevents page overflow with 7+ rows; (4) Value Chain `#67` dynamic `stage_w` and `stage_h` fill content area instead of fixed 2.0" width; (5) Closing `#36` bottom line changed from `Inches(3)` to `CW` for full-width. New **Production Guard Rails Rule 8**: dynamic sizing for variable-count layouts. **5 new Common Issues** (Problems 11-15). Updated code examples for #1, #36, #61, #67. |
| 1.10.3 | 2026-03-18 | **Title Line Spacing Optimization**: Titles (≥18pt) now use `0.93` multiple spacing instead of fixed `Pt(fs*1.35)`, producing tighter, more professional title rendering. Body text (<18pt) retains fixed Pt spacing. Updated Problem 5 documentation. Thanks to **冯梓航 Denzel** for detailed feedback. |
| 1.10.2 | 2026-03-18 | **#54 Matrix Side Panel Variant**: Added compact grid + side panel layout variant for Pattern #54 (Risk/Heat Matrix). When matrix needs a companion insight panel, `cell_w` shrinks from 3.0" to 2.15" and `axis_label_w` from 1.8" to 0.65", yielding ~4.2" panel width. Includes layout math, ASCII wireframe, code example, and minimum-width rule. |
| 1.10.1 | 2026-03-18 | **Frontmatter Fix**: Fixed "malformed YAML frontmatter" error on Claude install. Removed unsupported fields (`license`, `version`, `metadata` with emoji, etc.) — Claude only supports `name` + `description`. Used YAML folded block scalar (`>-`) for description. Metadata relocated to document body. |
| 1.10.0 | 2026-03-18 | **Channel Delivery**: New `deliver_to_channel()` helper sends generated PPTX back to user's chat via `openclaw message send --media`. Supports Feishu/飞书, Telegram, WhatsApp, Discord, Slack, Signal and all OpenClaw channels. Graceful fallback when not in channel context. Updated example scripts. |
| 1.9.0 | 2026-03-15 | **Production Guard Rails**: 7 mandatory rules derived from real-world feedback — spacing/overflow protection, legend color consistency, title style uniformity (`add_action_title()` only), axis label centering, image placeholder page requirement, bottom whitespace elimination, content overflow detection. **Code Efficiency Guidelines**: variable reuse, helper function patterns, short abbreviation table, batch data structures, auto page numbering. **5 new Common Issues** (Problems 6-10). |
| 1.8.0 | 2026-03-15 | **Massive layout expansion**: 39 → **70 patterns** across 8 → **12 categories**. Added Category I (Image+Content, #40-#47), Category J (Advanced Data Viz, #48-#56), Category K (Dashboards, #57-#58), Category L (Visual Storytelling, #59-#70). New `add_image_placeholder()` helper. Image Priority Rule added. Layout Diversity table expanded. Based on McKinsey PowerPoint Template 2023 analysis. |
| 1.7.0 | 2026-03-13 | **Category H: Data Charts**: Added 3 new chart layout patterns (#37 Grouped Bar, #38 Stacked Bar, #39 Horizontal Bar) using pure `add_rect()` drawing. Added Chart Priority Rule to Layout Diversity table — when data contains dates + values/percentages, chart patterns are mandatory. Total patterns: 39. |
| 1.6.0 | 2026-03-08 | **Cross-model quality alignment**: Added Accent Color System (4 accent + 4 light BG colors), Presentation Planning section (structure templates, layout diversity rules, content density requirements, mandatory slide elements, page number helper). Based on comparative analysis across Opus/Minimax/Hunyuan/GLM5 outputs. |
| 1.5.0 | 2026-03-08 | **Critical fix**: `add_text()` now sets `p.line_spacing = Pt(font_size.pt * 1.35)` to prevent Chinese multi-line text overlap. Added Problem 5 to Common Issues. |
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
| | | - `add_text()` bullet param removed; use `'\u2022 '` prefix in text |
| 1.0.0 | 2026-03-02 | Initial complete specification, all refinements documented |
| | | - Color palette finalized (NAVY primary) |
| | | - Typography hierarchy locked (22pt title, 14pt body) |
| | | - Line treatment standardized (no shadows) |
| | | - Theme cleanup process documented |
| | | - All helper functions optimized |

