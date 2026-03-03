---
name: workbuddy-ppt-design
description: "McKinsey-style PowerPoint presentation design for Tencent WorkBuddy product introductions. This skill provides comprehensive design guidelines, typography standards, color palettes, layout principles, and Python code patterns for creating professional, consistent presentations using python-pptx from scratch. Includes all refinements and user feedback from iterative design process."
license: Apache-2.0
version: "1.1.0"
user-invocable: true
allowed-tools:
  - Read
  - Write
  - Bash
---

# WorkBuddy PPT Design Framework

## Overview

This skill encodes the complete design specification for **Tencent WorkBuddy product introductions** - a professional, consultant-grade PowerPoint presentation framework based on McKinsey design principles. It includes:

- **Color systems** and typography hierarchy
- **Layout patterns** for different slide types
- **Design principles** (minimalism, consistency, hierarchy)
- **Font specifications** (English vs. Chinese character handling)
- **Line and shape treatments** (no shadows, clean/flat design)
- **Python-pptx code patterns** ready to customize

All specifications have been refined through **iterative user feedback** to ensure visual consistency and professional polish.

---

## When to Use This Skill

Use this skill when you need to:

1. **Create a new WorkBuddy product presentation from scratch** - Use the provided Python code templates
2. **Maintain design consistency** - Reference color codes, font sizes, spacing, and layout patterns
3. **Refine existing presentations** - Apply font adjustments, line treatments, or layout improvements
4. **Document design decisions** - Explain color choices, typography rationale, or layout structures
5. **Troubleshoot common issues**:
   - PPT files won't open in PowerPoint (shadow/3D effects causing corruption)
   - Text formatting inconsistencies across slides
   - Lines appearing with unwanted shadows or effects
   - Font size hierarchy problems

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
| **44pt** | Cover Title | "腾讯WorkBuddy" | Cover slide only, Georgia |
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
         '腾讯WorkBuddy', font_size=Pt(44), font_name='Georgia',
         font_color=NAVY, bold=True, ea_font='KaiTi')
add_text(s1, Inches(1), Inches(3.5), Inches(11), Inches(0.6),
         'AI驱动的新一代办公效率助手', font_size=Pt(24),
         font_color=DARK_GRAY, ea_font='KaiTi')
add_text(s1, Inches(1), Inches(4.5), Inches(11), Inches(0.5),
         '产品功能介绍  |  2026年3月', font_size=BODY_SIZE,
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
    """Add a circle label with a letter (e.g. 'A', '1')."""
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    c.fill.solid()
    c.fill.fore_color.rgb = bg
    c.line.fill.background()
    tf = c.text_frame
    tf.paragraphs[0].text = letter
    tf.paragraphs[0].font.size = Pt(14)
    tf.paragraphs[0].font.color.rgb = fg
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
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

See `examples/minimal_workbuddy_ppt.py` for a complete, working example that generates:
- Cover slide
- Table of contents
- Content slide with title + body text
- Source attribution
- Proper theme cleanup

---

## File References

Generated presentations are typically saved to:
```
/Users/kaku/.workbuddy/workspace/default_project/WorkBuddy产品介绍v2.pptx
```

All colors, fonts, and dimensions referenced in code should match this document exactly.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
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

