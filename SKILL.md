---
name: workbuddy-ppt-design
description: "McKinsey-style PowerPoint presentation design for Tencent WorkBuddy product introductions. This skill provides comprehensive design guidelines, typography standards, color palettes, layout principles, and Python code patterns for creating professional, consistent presentations using python-pptx from scratch. Includes all refinements and user feedback from iterative design process."
license: Apache-2.0
version: "1.0.0"
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

#### Code Implementation

```python
def add_line(slide, x1, y1, x2, y2, color=BLACK, width=Pt(0.5)):
    """Add a simple flat line with NO theme effects/shadows."""
    c = slide.shapes.add_connector(1, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = width
    # CRITICAL: Remove the theme style reference that causes shadow
    sp = c._element
    style = sp.find(qn('p:style'))
    if style is not None:
        sp.remove(style)
    return c
```

**Why this matters**: python-pptx automatically attaches `<p:style>` to connectors, which references theme effects. The theme by default includes `outerShdw` (outer shadow). We must remove this to prevent:
- File corruption when saving/reopening in PowerPoint
- Unwanted shadows appearing on lines
- "File needs repair" errors in Microsoft Office

#### Post-Save Theme Cleanup

After `prs.save(outpath)`, ALWAYS clean the theme XML:

```python
import zipfile, os
from lxml import etree

tmppath = outpath + '.tmp'
with zipfile.ZipFile(outpath, 'r') as zin:
    with zipfile.ZipFile(tmppath, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if 'theme' in item.filename.lower() and item.filename.endswith('.xml'):
                root = etree.fromstring(data)
                ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                # Remove all shadow effects from effectStyleLst
                for eff_list in root.findall('.//a:effectStyleLst/a:effectStyle/a:effectLst', ns):
                    for shadow in eff_list.findall('a:outerShdw', ns):
                        eff_list.remove(shadow)
                    for shadow in eff_list.findall('a:innerShdw', ns):
                        eff_list.remove(shadow)
                # Remove 3D effects
                for eff_style in root.findall('.//a:effectStyleLst/a:effectStyle', ns):
                    for scene3d in eff_style.findall('a:scene3d', ns):
                        eff_style.remove(scene3d)
                    for sp3d in eff_style.findall('a:sp3d', ns):
                        eff_style.remove(sp3d)
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
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
- No shadow effects (automatically removed by theme cleanup)

```python
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
shape.fill.solid()
shape.fill.fore_color.rgb = BG_GRAY
shape.line.fill.background()  # CRITICAL: removes border
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
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

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
                  alignment=PP_ALIGN.LEFT, ea_font='KaiTi', bullet=False, line_spacing_pt=6):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    for attr in ['lIns','tIns','rIns','bIns']:
        bodyPr.set(attr, '45720')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ('• ' if bullet else '') + line
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
    return shape

def add_line(slide, x1, y1, x2, y2, color=RGBColor(0, 0, 0), width=Pt(0.5)):
    """Add a simple flat line with NO theme effects/shadows."""
    c = slide.shapes.add_connector(1, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = width
    sp = c._element
    style = sp.find(qn('p:style'))
    if style is not None:
        sp.remove(style)
    return c

def add_circle_label(slide, x, y, letter, size=Inches(0.5), 
                     color=RGBColor(0x05, 0x1C, 0x2C)):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    c.fill.solid()
    c.fill.fore_color.rgb = color
    c.line.fill.background()
    tf = c.text_frame
    tf.paragraphs[0].text = letter
    tf.paragraphs[0].font.size = Pt(14)
    tf.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    sp = c._element
    style = sp.find(qn('p:style'))
    if style is not None:
        sp.remove(style)
    return c

def add_action_title(slide, text, title_size=Pt(22)):
    """White bg, black text, thin black line below — NO shadow."""
    add_text(slide, Inches(0.8), Inches(0.15), Inches(11.7), Inches(0.9),
             text, font_size=title_size, font_color=RGBColor(0, 0, 0), bold=True,
             font_name='Georgia', ea_font='KaiTi', anchor=MSO_ANCHOR.MIDDLE)
    add_line(slide, Inches(0.8), Inches(1.05), Inches(12.5), Inches(1.05),
             color=RGBColor(0, 0, 0), width=Pt(0.5))

def add_source(slide, text, y=Inches(7.05)):
    add_text(slide, Inches(0.8), y, Inches(11), Inches(0.3),
             text, font_size=Pt(9), font_color=RGBColor(0x66, 0x66, 0x66))
```

---

## Common Issues & Solutions

### Problem 1: PPT Won't Open / "File Needs Repair"

**Cause**: Connectors have `<p:style>` referencing theme effects with shadows

**Solution**:
1. Ensure all `add_line()` calls use the provided function (which removes `p:style`)
2. After `prs.save()`, run theme cleanup code (see "Post-Save Theme Cleanup" section)
3. Verify with:
   ```python
   from pptx import Presentation
   from pptx.oxml.ns import qn
   prs = Presentation('file.pptx')
   for slide in prs.slides:
       for shape in slide.shapes:
           if shape._element.tag.endswith('}cxnSp'):
               if shape._element.find(qn('p:style')) is not None:
                   print(f"ERROR: {shape.name} still has p:style!")
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

**Cause**: Theme's default effects being applied

**Solution**:
1. Ensure `add_line()` removes `p:style` immediately after creating connector
2. Run post-save theme cleanup (see section above)
3. Do NOT rely on PowerPoint to "fix" shadows - they indicate underlying corruption

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
- [ ] **Line width changes** → Update calls to `add_line()` with new `width=Pt(X)`
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
4. **Keep code DRY** - Use helper functions like `add_text()`, `add_line()`, etc.
5. **Validate XML** - After adding theme cleanup, verify no shadows remain
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
| 1.0.0 | 2026-03-02 | Initial complete specification, all refinements documented |
| | | - Color palette finalized (NAVY primary) |
| | | - Typography hierarchy locked (22pt title, 14pt body) |
| | | - Line treatment standardized (no shadows) |
| | | - Theme cleanup process documented |
| | | - All helper functions optimized |

