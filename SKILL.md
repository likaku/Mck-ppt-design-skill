---
name: mck-ppt-design
description: >-
  Create professional, consultant-grade PowerPoint presentations from scratch
  using MckEngine (python-pptx wrapper) with McKinsey-style design. Use when
  user asks to create slides, pitch decks, business presentations, strategy
  decks, quarterly reviews, board meeting slides, or any professional PPTX.
  AI calls eng.cover(), eng.donut(), eng.timeline() etc — 67 high-level methods
  across 12 categories (structure, data, framework, comparison, narrative,
  timeline, team, charts, images, advanced viz, dashboards, visual storytelling),
  consistent typography, zero file-corruption issues, BLOCK_ARC native shapes
  for circular charts (donut, pie, gauge), production-hardened guard rails
  for spacing, overflow, legend consistency, title style uniformity,
  dynamic sizing for variable-count layouts, chart rendering, and
  AI-generated cover images via Tencent Hunyuan 2.0 with professional cutout,
  cool grey-blue tint, and McKinsey-style Bézier ribbon decoration.
---

# McKinsey PPT Design Framework

> **Copyright © 2024-2026 Kaku Li (https://github.com/likaku).** Licensed under Apache 2.0. See [NOTICE](NOTICE).

> **Version**: 2.2.0 · **License**: Apache-2.0 · **Author**: [likaku](https://github.com/likaku/Mck-ppt-design-skill)
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
- **AI-generated cover images** — Tencent Hunyuan 2.0 text-to-image → rembg professional cutout → cool grey-blue tint + 50% lighten → McKinsey-style Bézier ribbon curves → transparent RGBA PNG full-bleed background (v2.2)

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


---

## MckEngine Quick Start (v2.0)

v2.0 introduces a **Python runtime engine** (`mck_ppt/`) that encapsulates all 67 layout methods. Instead of writing raw `add_shape()` / `add_text()` coordinate code, the AI calls high-level methods like `eng.cover()`, `eng.donut()`, `eng.timeline()`.

### Why Use MckEngine

| | v1.x (inline code) | v2.0 (MckEngine) |
|---|---|---|
| Code generation | AI writes `add_shape()` + coordinate math per slide | AI calls `eng.cover()`, `eng.donut()` etc. |
| Output tokens per 30-slide deck | 40,000–60,000 | 9,000–12,000 |
| Rounds per deck | 10–15 | 3–4 |
| Chart rendering | `add_rect()` stacking (100–2,800 shapes) | `BLOCK_ARC` native arcs (3–4 shapes) |
| File corruption risk | Manual cleanup needed | Automatic three-layer defense |

### Setup

```bash
pip install python-pptx lxml
```

The `mck_ppt/` package lives inside the skill directory. Before generating any presentation, the AI MUST:

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
```

### Complete Generation Pattern

Every presentation script follows this exact pattern:

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
from mck_ppt.constants import *  # NAVY, ACCENT_BLUE, etc.
from pptx.util import Inches

eng = MckEngine(total_slides=12)  # Set total for page numbering

# ── Structure ──
eng.cover(title='Q1 2026 战略回顾', subtitle='董事会汇报', author='战略部', date='2026年3月', cover_image='auto')
eng.toc(items=[('1', '市场概览', '当前竞争格局'),
               ('2', '财务分析', '营收与利润趋势'),
               ('3', '战略建议', '三大核心行动')])

# ── Content slides ──
eng.section_divider(section_label='第一部分', title='市场概览')
eng.big_number(title='市场规模', number='¥850亿', description='2026年预估市场总量',
    detail_items=['同比增长23%', '线上渠道占比突破60%'], source='Source: 行业报告 2026')
eng.donut(title='市场份额分布',
    segments=[(0.35, NAVY, '我们'), (0.25, ACCENT_BLUE, '竞品A'),
              (0.20, ACCENT_GREEN, '竞品B'), (0.20, ACCENT_ORANGE, '其他')],
    center_label='35%', center_sub='市占率', source='Source: 市场调研 2026')

eng.section_divider(section_label='第二部分', title='财务分析')
eng.grouped_bar(title='季度营收趋势',
    categories=['Q1', 'Q2', 'Q3', 'Q4'],
    series=[('产品', NAVY), ('服务', ACCENT_BLUE)],
    data=[[120, 80], [145, 95], [160, 110], [180, 130]],
    max_val=200, source='Source: 财务部')

eng.section_divider(section_label='第三部分', title='战略建议')
eng.table_insight(title='三大战略方向对比',
    headers=['战略方向', '核心举措', '预期成效'],
    rows=[['产品创新', 'AI赋能 + 用户体验升级', '市场份额+15%'],
          ['市场拓展', '进入3个新行业 + 海外布局', '营收增长30%'],
          ['运营卓越', '成本优化20% + 数字化覆盖85%', '利润率+8pp']],
    insights=['三大方向协同发力，形成增长飞轮', '产品创新为引擎，市场拓展为杠杆', '运营卓越为底座，确保可持续性'],
    source='Source: 战略部')
eng.timeline(title='执行路线图',
    milestones=[('Q1', '方案设计'), ('Q2', '试点验证'),
                ('Q3', '规模推广'), ('Q4', '效果评估')],
    source='Source: PMO')

# ── Closing ──
eng.closing(title='谢谢', message='期待与您进一步交流')

# ── Save (auto cleanup) ──
eng.save('output/q1_strategy_review.pptx')
print('Done! 12 slides generated.')
```

### Key Rules

1. **One method = one slide**. `eng.cover()` creates slide 1, `eng.toc()` creates slide 2, etc.
2. **`eng.save()` handles everything** — XML cleanup, shadow removal, p:style sanitization. No manual `full_cleanup()` needed.
3. **Page numbers are automatic** — the engine tracks `_page` internally.
4. **All guard rails are built-in** — dynamic sizing, overflow protection, CJK font handling.
5. **Use constants from `mck_ppt.constants`** — `NAVY`, `ACCENT_BLUE`, `BG_GRAY`, `BODY_SIZE`, etc.

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
 Slides 3-5: Core Content (vary layouts: #8, #71, #19, #33)
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
| 3-4 parallel concepts | Table+Insight (#71), Four-Column (#27), Metric Cards (#10), Icon Grid (#63) | Bullet list |
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
| **Opening analysis / key arguments** | **Table+Insight (#71), Key Takeaway (#25)** | **Bullet list, Plain text** |

**NEVER** use Two-Column Text (#26) for more than 1 slide per deck. It is the least visually engaging layout.

**OPENING SLIDE PRIORITY RULE**: For **Slides 2–5** (the first few content slides after cover/TOC), **strongly prefer high-impact editorial layouts** that set the tone for the entire presentation. Prioritized layouts for opening slides (in order of preference):
1. **Table+Insight (`table_insight`, #71)** — structured arguments with gray-bg right-panel takeaways + chevron icon
2. **Big Number (#8)** — single impactful statistic
3. **Key Takeaway (#25)** — left detail + right summary

These layouts create a strong visual opening that hooks the audience. Avoid starting presentations with plain text or simple bullet lists.

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
# Without cover image (default, classic layout)
eng.cover(title='Q1 2026 战略回顾', subtitle='董事会汇报', author='战略部', date='2026年3月')

# With AI-generated cover image (McKinsey style)
eng.cover(title='Q1 2026 战略回顾', subtitle='董事会汇报', author='战略部', date='2026年3月', cover_image='auto')

# With custom image file
eng.cover(title='Q1 2026 战略回顾', subtitle='董事会汇报', author='战略部', date='2026年3月', cover_image='assets/my_cover.png')
```

**Cover Image Generation Pipeline** (`cover_image='auto'`):

When `cover_image='auto'`, the system automatically generates a McKinsey-style cover illustration:

1. **Keyword → Real Product Mapping**: Title keywords are matched to real-world product descriptions via `_METAPHOR_MAP` in `mck_ppt/cover_image.py` (e.g. `'AI'` → high-end triple-fan GPU, `'医药'` → stethoscope + capsules, `'金融'` → metal chip bank card, `'建筑'` → 3D-printed architectural model)
2. **Hunyuan 2.0 Generation**: Tencent Hunyuan 2.0 (`SubmitHunyuanImageJob` async API) generates a 1024×1024 product photo with prompt: "真实产品摄影照片，{product}，纯白色背景，轮廓清晰锐利，影棚灯光，超高清"
3. **Professional Cutout**: `rembg` removes background completely — only the product subject remains with transparent background
4. **Cool Grey-Blue Tint**: Desaturation (30%), channel shift (R×0.85, G×0.92, B×1.18), then 50% lighten by blending with white
5. **Canvas Placement**: Subject scaled to ~66% canvas height, placed at right-center of 1920×1080 transparent canvas
6. **Bézier Ribbon Curves**: 24 parallel cubic Bézier curves drawn from bottom-left to top-right, with a gentle twist/fold at center (lines cross over like a folded silk ribbon). Inner lines thicker+darker, outer lines thinner+lighter
7. **Full-bleed Embed**: Image added as first shape (bottom layer), all text rendered on top

**Requirements**: `pip install tencentcloud-sdk-python rembg pillow numpy` + env vars `TENCENT_SECRET_ID`, `TENCENT_SECRET_KEY`

**Standalone usage**:
```python
from mck_ppt.cover_image import generate_cover_image
path = generate_cover_image('AI的能力边界', output_path='cover.png')
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
eng.section_divider(section_label='第一部分', title='市场分析', subtitle='当前格局与核心机会')
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
eng.toc(items=[('1', '市场概览', '当前竞争格局'), ('2', '战略方向', '三大核心举措'), ('3', '执行路线', '时间表与责任人')])
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
eng.appendix_title(title='附录', subtitle='补充数据与参考资料')
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
eng.big_number(title='关键发现标题', number='95%', unit='', description='描述数据含义',
    detail_items=['洞见要点一', '洞见要点二', '洞见要点三'], source='Source: ...')
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
eng.two_stat(title='对比标题',
    stats=[('$2.4B', '2026年目标', True), ('$1.8B', '2025年实际', False)],
    source='Source: ...')
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
eng.three_stat(title='核心运营指标',
    stats=[('98.5%', '系统可用性', True), ('12ms', '平均响应时间', False), ('4.8', '用户满意度', True)],
    source='Source: ...')
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
eng.data_table(title='五大核心能力',
    headers=['模块', '功能描述', '应用场景'],
    rows=[['AI Agent', '自主决策与执行', '客服自动化'],
          ['数据引擎', '实时数据处理', '风控决策']],
    source='Source: ...')
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
eng.metric_cards(title='月度运营仪表盘',
    cards=[('98.5%', '可用性', '超越SLA目标'),
           ('¥2.3亿', '月营收', '同比+18%'),
           ('4.8/5', '满意度', '连续3月提升')],
    source='Source: ...')
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
eng.matrix_2x2(title='战略优先级矩阵',
    quadrants=[('高价值 / 低难度', ['快速推进项目A', '立即启动项目B']),
               ('高价值 / 高难度', ['重点攻关项目C']),
               ('低价值 / 低难度', ['委托执行项目D']),
               ('低价值 / 高难度', ['暂缓或放弃'])],
    source='Source: ...')
```

#### ~~14. Three-Pillar Framework~~ → RETIRED (v2.0.4)

> **已退役**：#14 布局已被 **#71 Table+Insight** 取代。请使用 `eng.table_insight()` 实现类似的三列对比需求，数据表达更清晰、视觉层级更强。

#### 15. Staircase Evolution (阶梯进化图) ⭐ v2.0.5

**适用场景**: 展示阶段性演进路径（如品牌进化、能力成长、战略升级），从左下到右上阶梯式上升。具有明显的阶梯台阶轮廓线。可选底部结构化详情表格。

```
┌──────────────────────────────────────────────────────┐
│ ▌ Action Title                                       │
├──────────────────────────────────────────────────────┤
│                              ●3● 远期标题             │
│                    ┌─────────┘                        │
│         ●2● 中期标题│  远期描述...                     │
│    ┌────┘          │                                  │
│  ●1● 近期标题 │ 中期描述...                            │
│  ─────────── │                                       │
│  近期描述... │                                        │
│──────────────────────────────────────────────────────│
│ 行标题1  │ • bullet1    │ • bullet1    │ • bullet1    │
│──────────────────────────────────────────────────────│
│ 行标题2  │ • bullet1    │ • bullet1    │ • bullet1    │
└──────────────────────────────────────────────────────┘
```

**Engine method**: `eng.pyramid()`

```python
eng.pyramid(
    title='品牌心智三层进化路径',
    levels=[
        ('层次一：2-3年', '"冰淇淋品类的操作系统"\nB端采购决策者', '1'),
        ('层次二：3-5年', '"冰淇淋界的Dolby"\n品质认知消费者', '2'),
        ('层次三：5-10年', '"美好甜蜜时刻的守护者"\n所有消费者', '3'),
    ],
    detail_rows=[
        ('核心策略', [
            ['AI数据+品类白皮书', '建立B端操作系统壁垒'],
            ['"日世品质"认证标识', '品质认知渗透消费端'],
            ['品牌情感化升级', '全渠道消费者心智占领'],
        ]),
        ('对标案例', [
            ['Intel Inside', 'B端技术标准→消费认知'],
            ['利乐 / 杜比', '品质认证→行业标配'],
            ['Gore-Tex', '专业品牌→大众信赖'],
        ]),
    ],
    source='对标: Intel Inside / 利乐 / 杜比 / Gore-Tex',
)
```

**Parameters**: `title`, `levels` (label, description, icon_text), `detail_rows` (row_label, [[col_texts]...]), `source`, `bottom_bar` (optional)

**设计特点**:
- **阶梯台阶轮廓线**：NAVY 颜色的水平台面线+垂直阶梯线，形成清晰的台阶外形
- **Icon + 标题同行**：NAVY 圆形 icon 在左，粗体标题在右，位于台面线上方
- **描述在台面线下方**：每阶段的详细描述文字在水平台面线下面
- **可选底部结构化表格**：行标题+各阶段 bullet 详情，居中对齐
- 兼容有/无 detail_rows 两种模式

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
eng.process_chevron(title='客户旅程五步法',
    steps=['需求识别', '方案设计', '实施交付', '运营优化', '持续创新'],
    bottom_bar=('关键洞见', '端到端数字化覆盖率从30%提升至85%'),
    source='Source: ...')
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
eng.venn(title='能力交叉模型',
    circles=[('技术', ['云计算', 'AI/ML'], 0.8, 1.5, 4.5, 3.5),
             ('业务', ['行业Know-how', '流程优化'], 3.5, 1.5, 4.5, 3.5)],
    overlap_label='数字化创新',
    source='Source: ...')
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
eng.temple(title='企业架构框架',
    roof_text='企业愿景：成为全球领先的科技公司',
    pillar_names=['产品创新', '客户体验', '运营卓越', '人才发展'],
    foundation_text='数据驱动 · 敏捷组织 · 开放生态',
    source='Source: ...')
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
eng.side_by_side(title='方案A vs 方案B',
    options=[('方案A：自建', ['完全自主可控', '前期投入¥500万', '上线周期6个月']),
             ('方案B：SaaS', ['快速上线2周', '年费¥80万', '依赖供应商'])],
    source='Source: ...')
```

#### 20. Before / After (前后对比页) _(v2.0.1 rewrite)_

**适用场景**: 展示变革前后的对比（如行业退潮 vs 生存公式、流程优化、组织变革）。

**设计特征**:
- **白底清洁布局** — 无背景色块，纯白底
- **黑色竖线 + 圆圈箭头** — 中间细竖线分隔，竖线正中放黑色实心圆圈内含 `>` 箭头（Arial 字体，内边距0）
- **结构化数据行**（左侧）— 每行有 label/brand/value/extra，数值红色大字
- **公式卡片**（右侧）— 每条有 title/desc/cases，案例数字黑色+下划线
- **支持简单文字列表后备** — 如传入 `list[str]` 自动退化为简单 bullet 模式
- **可选虚线角标** — 右上角虚线框（如 `Part II > 退潮`）
- **可选底部总结条** — bottom_bar

```
┌──────────────────────────────────────────────┐
│ ▌ Action Title                   ┊Part II >┊ │
├──────────────────┬────┬──────────────────────┤
│  左侧标题        │    │  右侧标题            │
│                  │ ● │                      │
│  标签 品牌 数值   │ > │  1. 公式标题          │
│  ────────────── │ ● │     描述 + 案例下划线  │
│  标签 品牌 数值   │    │  ─────────────      │
│  ────────────── │    │  2. 公式标题          │
│  总结(灰粗体)    │    │     总结(红色粗体)    │
├──────────────────────────────────────────────┤
│ [关键洞察] 底部总结条                         │
└──────────────────────────────────────────────┘
```

**Engine method**: `eng.before_after()`

```python
eng.before_after(title='流程优化效果',
    before_title='优化前', before_points=['人工审批 3-5天', '错误率 8%', '满意度 65分'],
    after_title='优化后', after_points=['自动审批 2小时', '错误率 0.5%', '满意度 92分'],
    source='Source: ...')
```

**Parameters**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `title` | str | 页面标题 |
| `before_title` | str | 左侧标题 |
| `before_points` | list[dict] 或 list[str] | 左侧数据行（dict: label/brand1/val1/brand2/val2/extra）或简单文字列表 |
| `after_title` | str | 右侧标题 |
| `after_points` | list[dict] 或 list[str] | 右侧公式卡片（dict: title/desc/cases）或简单文字列表 |
| `corner_label` | str | 右上角虚线角标文字（可选） |
| `bottom_bar` | tuple(str,str) | 底部条 (标签, 文字)（可选） |
| `left_summary` | str | 左侧底部总结文字（可选，灰色粗体） |
| `right_summary` | str | 右侧底部总结文字（可选，默认红色粗体） |
| `right_summary_color` | RGBColor | 右侧总结文字颜色（默认 ACCENT_RED） |
| `source` | str | 数据来源 |

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
eng.pros_cons(title='并购方案评估',
    pros_title='优势', pros=['快速获取市场份额', '技术团队整合', '品牌协同效应'],
    cons_title='风险', cons=['整合成本高', '文化冲突', '监管审批不确定'],
    source='Source: ...')
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
eng.rag_status(title='项目健康度仪表盘',
    headers=['项目', '进度', '预算', '质量', '负责人'],
    rows=[('CRM升级', '🟢', '🟡', '🟢', '张三'),
          ('ERP迁移', '🟡', '🔴', '🟢', '李四')],
    source='Source: ...')
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
eng.scorecard(title='数字化成熟度评估',
    items=[('数据治理', 85, '已建立完整数据标准'),
           ('流程自动化', 62, 'RPA覆盖40%核心流程'),
           ('AI应用', 45, '试点阶段，3个场景落地')],
    source='Source: ...')
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
eng.executive_summary(title='执行摘要',
    headline='本季度实现营收¥8.5亿，同比增长23%，超额完成年度目标的52%',
    items=[('市场拓展', '新增客户127家，覆盖3个新行业'),
           ('产品迭代', '发布V3.0版本，NPS提升15个百分点'),
           ('运营效率', '人效比提升18%，交付周期缩短40%')],
    source='Source: ...')
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
eng.key_takeaway(title='核心洞见',
    takeaway='数字化转型的关键不在技术，而在组织能力的重塑',
    details=['技术是enabler，组织是driver', '自上而下的战略共识是前提', '小步快跑优于大规模重构'],
    source='Source: ...')
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
eng.quote(title='客户反馈',
    quote_text='这是我们见过的最专业、最高效的数字化转型方案。',
    attribution='张总, XX集团CEO',
    source='Source: ...')
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
eng.two_column_text(title='能力对比分析',
    columns=[('A', '核心能力', ['云原生架构设计', '大规模分布式系统', 'AI/ML工程化']),
             ('B', '待提升领域', ['前端体验设计', '国际化运营', '生态合作伙伴'])],
    source='Source: ...')
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
eng.four_column(title='四大战略方向',
    items=[('产品创新', '持续迭代核心产品线'),
           ('市场拓展', '进入3个新垂直行业'),
           ('运营卓越', '全流程数字化覆盖'),
           ('人才战略', '关键岗位100%到位')],
    source='Source: ...')
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
eng.timeline(title='项目路线图 2026',
    milestones=[('Q1', '需求调研与方案设计'),
                ('Q2', '核心功能开发与测试'),
                ('Q3', '灰度发布与用户反馈'),
                ('Q4', '全量上线与效果评估')],
    source='Source: ...')
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
eng.vertical_steps(title='实施五步法',
    steps=[('1', '诊断评估', '全面了解现状与痛点'),
           ('2', '方案设计', '制定分阶段实施方案'),
           ('3', '试点验证', '选取2-3个场景快速验证'),
           ('4', '规模推广', '复制成功经验至全业务线'),
           ('5', '持续优化', '建立数据驱动的迭代机制')],
    source='Source: ...')
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
eng.cycle(title='敏捷开发循环',
    phases=[('规划', 1.0, 2.0), ('开发', 5.0, 1.0),
            ('测试', 9.0, 2.0), ('发布', 5.0, 4.0)],
    source='Source: ...')
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
eng.funnel(title='销售漏斗分析',
    stages=[('线索获取', '10,000', '100%'),
            ('需求确认', '3,500', '35%'),
            ('方案报价', '1,200', '12%'),
            ('合同签署', '480', '4.8%')],
    source='Source: ...')
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
eng.meet_the_team(title='核心团队',
    members=[('张三', 'CEO', '15年行业经验'),
             ('李四', 'CTO', '前Google高级工程师'),
             ('王五', 'VP Sales', '年销售额¥5亿+')],
    source='Source: ...')
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
eng.case_study(title='XX银行数字化转型案例',
    sections=[('S', '背景', '传统核心系统老化，无法支撑业务增长'),
              ('A', '行动', '分阶段微服务改造 + 数据中台建设'),
              ('R', '成果', '交易处理能力提升10倍，故障率下降90%')],
    result_box=('关键指标', 'ROI 380% | 12个月回本'),
    source='Source: ...')
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
eng.action_items(title='下一步行动计划',
    actions=[('启动数据治理项目', 'Q2 2026', '建立统一数据标准与质量体系', '张三'),
             ('招聘AI工程师团队', 'Q1-Q2 2026', '组建10人ML工程团队', '李四'),
             ('完成ERP云迁移', 'Q3 2026', '核心ERP系统迁移至云原生架构', '王五')],
    source='Source: ...')
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
eng.closing(title='谢谢', message='期待与您进一步交流')
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
eng.grouped_bar(title='季度营收趋势（按产品线）',
    categories=['Q1', 'Q2', 'Q3', 'Q4'],
    series=[('产品A', NAVY), ('产品B', ACCENT_BLUE)],
    data=[[120, 80], [145, 95], [160, 110], [180, 130]],
    y_max=200, y_step=50, y_unit='万',
    source='Source: ...')
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
eng.stacked_bar(title='营收构成变化趋势',
    periods=['2023', '2024', '2025', '2026E'],
    series=[('产品', NAVY), ('服务', ACCENT_BLUE), ('订阅', ACCENT_GREEN)],
    data=[[40, 35, 25], [35, 35, 30], [30, 35, 35], [25, 35, 40]],
    source='Source: ...')
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
eng.horizontal_bar(title='各部门数字化成熟度排名',
    items=[('研发部', 92, NAVY), ('市场部', 78, ACCENT_BLUE), ('运营部', 65, ACCENT_GREEN),
           ('财务部', 58, ACCENT_ORANGE), ('HR', 45, MED_GRAY)],
    source='Source: ...')
```

---

### Category I: Image + Content Layouts

> **Image Placeholder Convention**: Since python-pptx cannot embed web images at generation time, all image positions use a **gray placeholder rectangle** with crosshair lines and a label. The user replaces these with real images after generation.

#### Helper: `add_image_placeholder()`

The `add_image_placeholder()` helper is available via `from mck_ppt.core import add_image_placeholder`. Image layouts in MckEngine call it automatically — you do not need to invoke it directly.

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
eng.content_right_image(title='产品核心功能',
    subtitle='AI智能分析引擎',
    bullets=['实时数据处理', '自动异常检测', '智能决策推荐'],
    takeaway='准确率达到98.5%，领先行业平均水平20个百分点',
    image_label='产品截图',
    source='Source: ...')
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
eng.three_images(title='办公环境',
    items=[('总部大楼', '位于科技园区核心位置', '总部外景'),
           ('开放办公', '敏捷协作空间设计', '办公区域'),
           ('创新实验室', '前沿技术研发基地', '实验室')],
    source='Source: ...')
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
eng.image_four_points(title='核心优势',
    image_label='产品示意图',
    points=[('高性能', '毫秒级响应'),
            ('高可用', '99.99% SLA'),
            ('安全', '多层防护体系'),
            ('弹性', '自动扩缩容')],
    source='Source: ...')
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
eng.full_width_image(title='全球布局',
    image_label='世界地图',
    overlay_text='覆盖全球32个国家和地区',
    source='Source: ...')
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
eng.case_study_image(title='XX银行案例',
    sections=[('背景', '传统系统老化', ACCENT_BLUE),
              ('方案', '微服务改造', ACCENT_GREEN),
              ('成果', '效率提升10倍', ACCENT_ORANGE)],
    image_label='系统架构图',
    kpis=[('10x', '处理能力'), ('90%', '故障减少')],
    source='Source: ...')
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
eng.quote_bg_image(title='',
    quote_text='创新不是选择，而是生存的必需。',
    attribution='CEO, 2026年全员大会',
    image_label='背景图',
    source='')
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
eng.goals_illustration(title='2026年度目标',
    goals=[('营收翻倍', '达到¥20亿年营收'),
           ('全球化', '进入5个海外市场'),
           ('IPO准备', '完成合规与治理升级')],
    image_label='目标愿景图',
    source='Source: ...')
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
eng.donut(title='2026年上半年营收渠道构成',
    segments=[(0.45, NAVY, '线上直营'), (0.28, ACCENT_BLUE, '经销商'),
              (0.15, ACCENT_GREEN, '企业客户'), (0.12, ACCENT_ORANGE, '其他')],
    center_label='¥8.5亿', center_sub='总营收',
    summary='线上直营渠道占比同比提升12个百分点',
    source='Source: ...')
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
eng.waterfall(title='利润变动瀑布图',
    items=[('2025年利润', 100, 'base'), ('营收增长', 35, 'up'),
           ('成本节省', 15, 'up'), ('新投资', -20, 'down'),
           ('2026年利润', 130, 'base')],
    source='Source: ...')
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
eng.line_chart(title='月活用户趋势',
    x_labels=['1月','2月','3月','4月','5月','6月'],
    y_labels=['0','100万','200万','300万'],
    values=[0.4, 0.45, 0.5, 0.6, 0.7, 0.82],
    legend_label='月活用户',
    source='Source: ...')
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
eng.pareto(title='缺陷类型分析（帕累托）',
    items=[('UI问题', 45), ('性能', 28), ('兼容性', 15), ('安全', 8), ('其他', 4)],
    max_val=50,
    source='Source: ...')
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
eng.kpi_tracker(title='OKR进度追踪',
    kpis=[('营收目标', 0.78, '¥7.8亿 / ¥10亿', 'on'),
          ('客户增长', 0.92, '184家 / 200家', 'on'),
          ('NPS提升', 0.65, '72分 / 80分', 'risk')],
    source='Source: ...')
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
eng.bubble(title='业务组合分析',
    bubbles=[(30, 70, 1.2, '产品A', NAVY),
             (60, 50, 0.8, '产品B', ACCENT_BLUE),
             (80, 30, 0.5, '产品C', ACCENT_GREEN)],
    x_label='市场份额 →', y_label='增长率 ↑',
    source='Source: ...')
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
eng.risk_matrix(title='项目风险评估矩阵',
    grid_colors=[[ACCENT_GREEN,ACCENT_ORANGE,ACCENT_RED],
                 [ACCENT_GREEN,ACCENT_ORANGE,ACCENT_ORANGE],
                 [ACCENT_GREEN,ACCENT_GREEN,ACCENT_ORANGE]],
    grid_lights=[[None]*3]*3,
    risks=[('数据泄露', 2, 2), ('系统宕机', 1, 2)],
    y_labels=['高','中','低'], x_labels=['低','中','高'],
    source='Source: ...')
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
eng.risk_matrix(title='项目风险评估矩阵',
    grid_colors=[[ACCENT_GREEN,ACCENT_ORANGE,ACCENT_RED],
                 [ACCENT_GREEN,ACCENT_ORANGE,ACCENT_ORANGE],
                 [ACCENT_GREEN,ACCENT_GREEN,ACCENT_ORANGE]],
    grid_lights=[[None]*3]*3,
    risks=[('数据泄露', 2, 2), ('系统宕机', 1, 2)],
    y_labels=['高','中','低'], x_labels=['低','中','高'],
    source='Source: ...')
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
eng.gauge(title='客户满意度',
    score=78,
    benchmarks=[('行业平均', '65分', ACCENT_ORANGE), ('目标', '85分', ACCENT_GREEN)],
    source='Source: ...')
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
eng.harvey_ball_table(title='供应商能力评估',
    headers=['供应商', '技术', '交付', '价格', '服务'],
    rows=[('供应商A', 100, 75, 50, 100),
          ('供应商B', 75, 100, 75, 50)],
    source='Source: ...')
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
eng.dashboard_kpi_chart(title='月度运营概览',
    kpi_cards=[('¥2.3亿', '月营收', '+18%', NAVY),
               ('98.5%', '可用性', '超SLA', ACCENT_GREEN),
               ('4.8', 'NPS', '+0.3', ACCENT_BLUE)],
    chart_data={'labels':['1月','2月','3月'], 'values':[180,210,230], 'color':NAVY},
    source='Source: ...')
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
eng.dashboard_table_chart(title='季度业务仪表盘',
    table_data={'headers':['指标','Q1','Q2','Q3'],
                'rows':[('营收','¥1.8亿','¥2.1亿','¥2.5亿'),
                        ('利润','¥0.3亿','¥0.4亿','¥0.5亿')]},
    source='Source: ...')
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
eng.stakeholder_map(title='干系人矩阵',
    quadrants=[('高关注 / 高影响', ['CEO', 'CTO']),
               ('高关注 / 低影响', ['项目经理', '产品经理']),
               ('低关注 / 高影响', ['监管机构']),
               ('低关注 / 低影响', ['普通员工'])],
    source='Source: ...')
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
eng.decision_tree(title='技术选型决策树',
    root='是否需要实时处理？',
    branches=[('是', ['流式计算', 'Kafka + Flink']),
              ('否', ['批处理', 'Spark + Hive'])],
    source='Source: ...')
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
eng.checklist(title='上线准备检查清单',
    columns=['检查项', '负责人', '截止日', '状态'],
    col_widths=[Inches(4.5), Inches(2.0), Inches(2.0), Inches(2.0)],
    rows=[('安全审计完成', '张三', '3/15', 'done'),
          ('性能压测通过', '李四', '3/20', 'wip'),
          ('文档更新', '王五', '3/25', 'todo')],
    status_map={'done': ('✅ 完成', ACCENT_GREEN, LIGHT_GREEN),
                'wip': ('🔄 进行中', ACCENT_ORANGE, LIGHT_ORANGE),
                'todo': ('⏳ 待开始', MED_GRAY, BG_GRAY)},
    source='Source: ...')
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
eng.metric_comparison(title='数字化转型前后关键指标对比',
    metrics=[('营收规模', '¥5.2亿', '¥8.5亿', '+63%'),
             ('库存周转', '45天', '28天', '–38%'),
             ('客户NPS', '52', '78', '+50%')],
    source='Source: ...')
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
eng.icon_grid(title='产品功能矩阵',
    items=[('🔍', '智能搜索', '毫秒级全文检索'),
           ('📊', '数据分析', '实时BI仪表盘'),
           ('🤖', 'AI助手', '自然语言交互'),
           ('🔒', '安全防护', '多层加密体系')],
    source='Source: ...')
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
eng.pie(title='市场份额分布',
    segments=[(0.35, NAVY, '我们', ''), (0.25, ACCENT_BLUE, '竞品A', ''),
              (0.20, ACCENT_GREEN, '竞品B', ''), (0.20, ACCENT_ORANGE, '其他', '')],
    source='Source: ...')
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
eng.swot(title='SWOT分析',
    quadrants=[('优势 Strengths', ['技术领先', '团队优秀', '资金充裕']),
               ('劣势 Weaknesses', ['品牌知名度低', '销售网络有限']),
               ('机会 Opportunities', ['市场快速增长', '政策利好']),
               ('威胁 Threats', ['巨头入场', '人才竞争激烈'])],
    source='Source: ...')
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
eng.agenda(title='季度回顾会议议程',
    headers=[('议题', Inches(5.0)), ('时间', Inches(2.0)),
             ('负责人', Inches(2.0)), ('备注', Inches(2.5))],
    items=[('Q3业绩回顾', '09:00-09:30', '张总', '', 'key'),
           ('产品路线图更新', '09:30-10:00', '李总', '', 'normal'),
           ('茶歇', '10:00-10:15', '', '', 'break'),
           ('2026规划讨论', '10:15-11:00', '全员', '', 'key')],
    source='Source: ...')
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
eng.value_chain(title='价值链分析',
    stages=[('研发', '产品设计与技术开发', ACCENT_BLUE),
            ('生产', '精益制造与质量管控', ACCENT_GREEN),
            ('营销', '品牌建设与渠道管理', ACCENT_ORANGE),
            ('服务', '客户支持与持续运营', NAVY)],
    source='Source: ...')
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
eng.two_col_image_grid(title='解决方案矩阵',
    items=[('智能客服', '7×24小时AI客服', ACCENT_BLUE, '客服系统截图'),
           ('数据分析', '实时BI仪表盘', ACCENT_GREEN, '分析面板'),
           ('流程自动化', 'RPA机器人流程', ACCENT_ORANGE, '自动化流程'),
           ('安全合规', '一站式合规管理', NAVY, '安全架构')],
    source='Source: ...')
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
eng.numbered_list_panel(title='战略建议',
    items=[('优先推进AI能力建设', '短期内集中资源打造AI核心竞争力'),
           ('建立数据治理体系', '统一数据标准，打通数据孤岛'),
           ('构建开放生态', '与合作伙伴共建行业解决方案')],
    panel=('战略目标', ['2026年AI渗透率达到60%', '数据资产价值提升3倍', '合作伙伴超过100家']),
    source='Source: ...')
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
eng.stacked_area(title='用户增长趋势',
    years=['2022', '2023', '2024', '2025', '2026E'],
    series_data=[('企业用户', [20, 35, 55, 80, 120], NAVY),
                 ('个人用户', [50, 80, 120, 160, 200], ACCENT_BLUE)],
    max_val=350, source='Source: ...')
```

---


---

## MckEngine API Reference

All 67 public methods on `MckEngine`. Every method creates exactly one slide (except `save()`).

### Initialization

```python
eng = MckEngine(total_slides=N)  # N = planned slide count (for page numbering)
```

### Save

```python
eng.save('output/deck.pptx')  # Auto-runs full_cleanup (p:style, shadow, 3D removal)
```

### Structure

**`eng.cover(title, subtitle='', author='', date='', cover_image=None)`**
> #1 Cover Slide — title, subtitle, author, date, accent line, optional AI-generated cover image.
> cover_image: `None` (no image, default), `'auto'` (AI-generated via Hunyuan 2.0), or `'path.png'` (custom image file).
> When `cover_image='auto'`: generates a McKinsey-style cover illustration automatically.
> Requires env vars: `TENCENT_SECRET_ID`, `TENCENT_SECRET_KEY`.

**`eng.toc(title='目录', items=None, source='')`**
> #6 Table of Contents — numbered items with descriptions.
> items: list of (num, title, description)

**`eng.section_divider(section_label, title, subtitle='')`**
> #5 Section Divider — navy left bar, large title.

**`eng.appendix_title(title, subtitle='')`**
> #7 Appendix Title — centered title with accent lines.

**`eng.closing(title, message='', source_text='')`**
> #36 Closing / Thank You slide.


### Data & Stats

**`eng.big_number(title, number, unit='', description='', detail_items=None, source='', bottom_bar=None)`**
> #8 Big Number — large stat with context.
> detail_items: list[str] bullet points shown below.
> bottom_bar: (label, text) or None.

**`eng.two_stat(title, stats, detail_items=None, source='')`**
> #9 Two-Stat Comparison — two big numbers side by side.
> stats: list of (number, label, is_navy:bool)

**`eng.three_stat(title, stats, detail_items=None, source='')`**
> #10 Three-Stat — three big numbers in a row.
> stats: list of 3 (number, label, is_navy:bool)

**`eng.data_table(title, headers, rows, col_widths=None, source='', bottom_bar=None)`**
> #11 Data Table — header row + data rows with separators.
> headers: list[str], rows: list[list[str]], col_widths: list[Inches] or auto.

**`eng.metric_cards(title, cards, source='')`**
> #12 Metric Cards — 3-4 accent-colored cards.
> cards: list of (letter, card_title, description, accent_color, light_bg)
> or (letter, card_title, description) — auto-colors from ACCENT_PAIRS.

**`eng.metric_comparison(title, metrics, source='')`**
> #62 Metric Comparison — before/after row cards with delta badges.
> metrics: list of (label, before_val, after_val, delta_str).


### Frameworks & Matrices

**`eng.matrix_2x2(title, quadrants, axis_labels=None, source='', bottom_bar=None)`**
> #13 2×2 Matrix — four quadrants.
> quadrants: list of 4 (label, bg_color, description).
> axis_labels: (x_label, y_label) or None.

**`eng.table_insight(title, headers, rows, insights, col_widths=None, insight_title='启示：', source='', bottom_bar=None)`**
> #71 Table+Insight — left data table + right insight panel with chevron icon.
> headers: list[str], rows: list[list[str]], insights: list[str].
> Supports **bold** markup in cell text. Replaces retired #14 Three-Pillar.

**`eng.pyramid(title, levels, source='', bottom_bar=None)`**
> #15 Pyramid — top-down widening layers.
> levels: list of (label, description, width_inches:float).

**`eng.process_chevron(title, steps, source='', bottom_bar=None)`**
> #16 Process Chevron — horizontal step flow with arrows.
> steps: list of (label, step_title, description).

**`eng.venn(title, circles, overlap_label='', right_text=None, source='')`**
> #17 Venn Diagram — overlapping rectangles for 2-3 sets.
> circles: list of (label, points:list[str], x, y, w, h) positioned rects.
> overlap_label: text for overlap zone.
> right_text: list[str] explanation on the right side.

**`eng.temple(title, roof_text, pillar_names, foundation_text, pillar_colors=None, source='')`**
> #18 Temple / House Framework — roof + pillars + foundation.
> pillar_names: list[str], pillar_colors: list[RGBColor] or auto.


### Comparison & Evaluation

**`eng.side_by_side(title, options, source='')`**
> #19 Side-by-Side Comparison — two columns with navy headers.
> options: list of 2 (option_title, points:list[str]).

**`eng.before_after(title, before_title, before_points, after_title, after_points, source='')`**
> #20 Before/After — gray (before) + navy (after) with arrow.

**`eng.pros_cons(title, pros_title, pros, cons_title, cons, conclusion=None, source='')`**
> #21 Pros/Cons — two-column layout.
> pros, cons: list[str].
> conclusion: (label, text) or None.

**`eng.rag_status(title, headers, rows, source='')`**
> #22 RAG Status — table with red/amber/green status dots.
> headers: list[str], rows: list of (name, status_color, *values, note).

**`eng.scorecard(title, items, source='')`**
> #23 Scorecard — items with progress bars.
> items: list of (name, score_str, pct_float_0_to_1)

**`eng.checklist(title, columns, col_widths, rows, status_map=None, source='', bottom_bar=None)`**
> #61 Checklist / Status table.
> columns: list[str] header labels.
> col_widths: list[Inches].
> rows: list of tuples — last element is status key.
> status_map: dict of status_key → (label, color, bg_color).

**`eng.swot(title, quadrants, source='')`**
> #65 SWOT Analysis — 2×2 colored grid.
> quadrants: list of 4 (label, accent_color, light_bg, points:list[str]).


### Narrative

**`eng.executive_summary(title, headline, items, source='')`**
> #24 Executive Summary — navy headline + numbered items.
> headline: str, items: list of (num, item_title, description).

**`eng.key_takeaway(title, left_text, takeaways, source='')`**
> #25 Key Takeaway — left analysis + right gray panel.
> left_text: list[str], takeaways: list[str].

**`eng.quote(quote_text, attribution='')`**
> #26 Quote Slide — centered quote with accent lines.

**`eng.two_column_text(title, columns, source='')`**
> #27 Two-Column Text — lettered columns with bullet lists.
> columns: list of 2 (letter, col_title, points:list[str]).

**`eng.four_column(title, items, source='')`**
> #28 Four-Column Overview — 4 vertical cards.
> items: list of (num, col_title, description:str_or_list).

**`eng.numbered_list_panel(title, items, panel=None, source='')`**
> #69 Numbered List + Side Panel — left numbered list + right accent panel.
> items: list of (item_title, description).
> panel: dict with 'subtitle','big_number','big_label','metrics':list[(label,value)].


### Timeline & Process

**`eng.timeline(title, milestones, source='', bottom_bar=None)`**
> #29 Timeline / Roadmap — horizontal line with milestone nodes.
> milestones: list of (label, description).

**`eng.vertical_steps(title, steps, source='', bottom_bar=None)`**
> #30 Vertical Steps — top-down numbered steps.
> steps: list of (num, step_title, description).

**`eng.cycle(title, phases, right_panel=None, source='')`**
> #31 Cycle Diagram — rectangular nodes with arrows in a loop.
> phases: list of (label, x_inches, y_inches) — positioned boxes.
> right_panel: (panel_title, points:list[str]) or None.

**`eng.funnel(title, stages, source='')`**
> #32 Funnel — top-down narrowing bars.
> stages: list of (name, count_label, pct_float).

**`eng.value_chain(title, stages, source='', bottom_bar=None)`**
> #67 Value Chain / Horizontal Flow — stages with arrows.
> stages: list of (stage_title, description, accent_color).
> Stages fill the full content width; height fills available vertical space.


### Team & Cases

**`eng.meet_the_team(title, members, source='')`**
> #33 Meet the Team — profile cards in a row.
> members: list of (name, role, bio:str_or_list).

**`eng.case_study(title, sections, result_box=None, source='')`**
> #34 Case Study — S/A/R or custom sections.
> sections: list of (letter, section_title, description).
> result_box: (label, text) or None.

**`eng.action_items(title, actions, source='')`**
> #35 Action Items — cards with timeline + owner.
> actions: list of (action_title, timeline, description, owner).

**`eng.case_study_image(title, sections, image_label, kpis=None, source='')`**
> #45 Case Study with Image — left text sections + right image + KPIs.
> sections: list of (label, text, accent_color).
> kpis: list of (value, label) or None.


### Charts (BLOCK_ARC)

**`eng.donut(title, segments, center_label='', center_sub='', legend_x=None, summary=None, source='')`**
> #48 Donut Chart — BLOCK_ARC ring segments.
> segments: list of (pct_float, color, label).

**`eng.pie(title, segments, legend_x=None, summary=None, source='')`**
> #64 Pie Chart — BLOCK_ARC with inner_ratio=0 (solid).
> segments: list of (pct_float, color, label, sub_label).

**`eng.gauge(title, score, benchmarks=None, source='')`**
> #55 Gauge — semicircle rainbow arc with center score.
> score: int 0-100.
> benchmarks: list of (label, value_str, color) shown below gauge.


### Charts (Bar/Line)

**`eng.grouped_bar(title, categories, series, data, max_val=None, y_ticks=None, summary=None, source='')`**
> #37 Grouped Bar Chart — vertical bars grouped by category.
> categories: list[str] x-labels. series: list of (name, color).
> data: list[list[int]] — data[cat_idx][series_idx].
> summary: (label, text) or None.

**`eng.stacked_bar(title, periods, series, data, summary=None, source='')`**
> #38 Stacked Bar Chart — 100% stacked vertical bars.
> periods: list[str] x-labels. series: list of (name, color).
> data: list[list[int]] — percentages, data[period_idx][series_idx].
> summary: (label, text) or None.

**`eng.horizontal_bar(title, items, summary=None, source='')`**
> #39 Horizontal Bar Chart — labeled bars with percentage.
> items: list of (name, pct_int_0_to_100, bar_color).
> summary: (label, text) or None.

**`eng.line_chart(title, x_labels, y_labels, values, legend_label='', summary=None, source='')`**
> #50 Line Chart — single line with dot approximation.
> x_labels: list[str], y_labels: list[str], values: list[float] 0.0-1.0 normalized.

**`eng.waterfall(title, items, max_val=None, legend_items=None, summary=None, source='')`**
> #49 Waterfall Chart — bridge from start to end.
> items: list of (label, value, type) — type: 'base'|'up'|'down'.

**`eng.pareto(title, items, max_val=None, summary=None, source='')`**
> #51 Pareto Chart — descending bars with value/pct labels.
> items: list of (label, value).

**`eng.stacked_area(title, years, series_data, max_val=None, summary=None, source='')`**
> #70 Stacked Area Chart — stacked columns for area approximation.
> years: list[str] x-labels.
> series_data: list of (name, values:list[int], color).


### Charts (Advanced)

**`eng.bubble(title, bubbles, x_label='', y_label='', legend_items=None, summary=None, source='')`**
> #53 Bubble / Scatter — positioned circles on XY plane.
> bubbles: list of (x_pct, y_pct, size_inches, label, color).

**`eng.kpi_tracker(title, kpis, summary=None, source='')`**
> #52 KPI Tracker — progress bars with status dots.
> kpis: list of (name, pct_float, detail, status_key).
> status_key: 'on'|'risk'|'off'.

**`eng.risk_matrix(title, grid_colors, grid_lights, risks, y_labels=None, x_labels=None, notes=None, source='')`**
> #54 Risk Matrix — 3×3 heatmap grid with risk labels.
> grid_colors: 3×3 list[list[RGBColor]] dot colors.
> grid_lights: 3×3 list[list[RGBColor]] cell backgrounds.
> risks: list of (row, col, name).
> notes: list[str] or None for bottom panel.

**`eng.harvey_ball_table(title, criteria, options, scores, legend_text=None, summary=None, source='')`**
> #56 Harvey Ball Table — matrix with Harvey Ball indicators.
> criteria: list[str] row labels. options: list[str] column headers.
> scores: list[list[int]] — scores[row][col], 0-4.


### Dashboards

**`eng.dashboard_kpi_chart(title, kpi_cards, chart_data=None, summary=None, source='')`**
> #57 Dashboard KPI + Chart — top KPI cards + bottom mini chart.
> kpi_cards: list of (value, label, detail, accent_color).
> chart_data: dict with 'labels','actual','target','max_val','legend'.

**`eng.dashboard_table_chart(title, table_data, chart_data=None, factoids=None, source='')`**
> #58 Dashboard Table + Chart — left table + right mini chart + bottom facts.
> table_data: dict with 'headers','col_widths','rows'.
> chart_data: dict with 'title','items':(name, value, max_val).
> factoids: list of (value, label, color).


### Image Layouts

**`eng.content_right_image(title, subtitle, bullets, takeaway='', image_label='Image', source='')`**
> #40 Content + Right Image.

**`eng.three_images(title, items, source='')`**
> #42 Three Images — three image+caption columns.
> items: list of (caption_title, description, image_label).

**`eng.image_four_points(title, image_label, points, source='')`**
> #43 Image + 4 Points — center image with 4 corner cards.
> points: list of 4 (point_title, description, accent_color).

**`eng.full_width_image(title, image_label, overlay_text='', attribution='', source='')`**
> #44 Full-Width Image — edge-to-edge image with text overlay.

**`eng.quote_bg_image(image_label, quote_text, attribution='', source='')`**
> #46 Quote with Background Image — image top + quote bottom.

**`eng.goals_illustration(title, goals, image_label, source='')`**
> #47 Goals with Illustration — left numbered goals + right image.
> goals: list of (goal_title, description, accent_color).

**`eng.two_col_image_grid(title, items, source='')`**
> #68 Two-Column Image + Text Grid — 2×2 image-text cards.
> items: list of (card_title, description, accent_color, image_label).


### Special

**`eng.icon_grid(title, items, cols=3, source='')`**
> #63 Icon Grid — grid of icon cards.
> items: list of (item_title, description, accent_color).

**`eng.stakeholder_map(title, quadrants, x_label='影响力 →', y_label='关注度 ↑', summary=None, source='')`**
> #59 Stakeholder Map — 2×2 quadrant with stakeholder lists.
> quadrants: list of 4 (label_cn, label_en, bg_color, members:list[str]).

**`eng.decision_tree(title, root, branches, right_panel=None, source='')`**
> #60 Decision Tree — root → L1 → L2 hierarchy with connector lines.
> root: (label,).
> branches: list of (L1_title, L1_metric, L1_color, children:list[(name, metric)]).
> right_panel: (panel_title, points:list[str]) or None.

**`eng.agenda(title, headers, items, footer_text='', source='')`**
> #66 Agenda — table-style meeting agenda.
> headers: list of (label, width).
> items: list of (*values, item_type) — type: 'key'|'normal'|'break'.


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

1. **Use MckEngine** — Never write raw `add_shape()` / coordinate code. Call `eng.xxx()` methods.
2. **One script, all slides** — Generate ALL planned slides in a single script run. Do not truncate.
3. **Set `total_slides` accurately** — This controls page number display (e.g., "Page 5/12").
4. **Use constants** — Import from `mck_ppt.constants`: `NAVY`, `ACCENT_BLUE`, `BG_GRAY`, `Inches`, etc.
5. **Layout diversity** — Each content slide MUST use a DIFFERENT layout from its neighbors.
6. **Chart priority** — When data has dates + values, use chart methods (`eng.grouped_bar`, `eng.donut`, etc.).
7. **Image priority** — For case studies / product showcases, use image layouts (`eng.content_right_image`, etc.).
8. **TOC completeness** — The TOC slide must list ALL content sections by number and title.
9. **`eng.save()` is sufficient** — It auto-runs `full_cleanup()`. No manual XML processing needed.

### Code Efficiency with MckEngine

MckEngine already handles constants, helpers, and cleanup internally. Your script only needs:

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
from mck_ppt.constants import *
from pptx.util import Inches

eng = MckEngine(total_slides=N)
# ... eng.cover() / eng.toc() / eng.xxx() calls ...
eng.save('output/deck.pptx')
```

No need to define `add_text()`, `add_rect()`, `add_hline()`, `_clean_shape()`, `full_cleanup()` — they are all encapsulated in the engine.

---

## Dependencies

- **python-pptx** >= 0.6.21
- **lxml** — XML processing for theme cleanup
- Python 3.8+

```bash
pip install python-pptx lxml
```

---

## Example: Complete Minimal Presentation

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
from mck_ppt.constants import *

eng = MckEngine(total_slides=5)
eng.cover(title='示例演示', subtitle='McKinsey Design Framework', date='2026')
eng.toc(items=[('1', '数据概览', '核心指标'), ('2', '分析', '趋势洞见')])
eng.big_number(title='核心发现', number='42%', description='年增长率',
    source='Source: 内部数据')
eng.table_insight(title='核心发现',
    headers=['维度', '现状', '目标'],
    rows=[['创新', '产品迭代中', 'AI赋能'],
          ['增长', '市场扩张中', '客户深耕'],
          ['效率', '流程优化中', '全面自动化']],
    insights=['创新驱动增长', '效率保障可持续'],
    source='Source: 战略部')
eng.closing(title='谢谢')
eng.save('output/demo.pptx')
```

---

## File References

```
~/.workbuddy/skills/mck-ppt-design/
├── SKILL.md                 # This file (design spec + API reference)
├── mck_ppt/
│   ├── __init__.py          # from mck_ppt import MckEngine
│   ├── engine.py            # 67 layout methods
│   ├── core.py              # Drawing primitives + XML cleanup
│   └── constants.py          # Colors, typography, grid constants
└── output/                  # Default output directory
```

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
| 2.2.0 | 2026-03-22 | **AI Cover Image Generation**: New `mck_ppt/cover_image.py` module. `eng.cover()` gains `cover_image` parameter (`None`/`'auto'`/`'path.png'`). When `'auto'`: Tencent Hunyuan 2.0 async API (`SubmitHunyuanImageJob`) generates 1024×1024 product photo → `rembg` professional background removal → cool grey-blue tint (desat 30%, R×0.85/G×0.92/B×1.18) + 50% lighten → subject placed at right-center of 1920×1080 transparent canvas → 24 McKinsey-style cubic Bézier ribbon curves with silk-fold twist at center → full-bleed RGBA PNG embedded as bottom layer. `_METAPHOR_MAP` maps 24 industry keywords to realistic product descriptions (GPU, capsules, bank card, solar panel, etc.). Prompt enforces: real product photography, sharp edges, white background, studio lighting. `__init__.py` exports `generate_cover_image`. Dependencies: `tencentcloud-sdk-python`, `rembg`, `pillow`, `numpy`. |
| 2.0.5 | 2026-03-21 | **#15 Staircase Evolution v3**: PNG icon support (auto-detect `.png` paths, overlay on navy circle with 0.08" inset). Single-line detail_rows = no bullet; multi-line = bullet. Icon library (6 icons in `assets/icons/`). New example: `staircase_civilization.py`. Unified release: merged v2.0.4 engine + v2.1 SKILL.md rewrite + #14→#71 cleanup. |
| 2.0.4 | 2026-03-19 | **#14 Three-Pillar RETIRED**: Removed `three_pillar` method from engine.py and its documentation from SKILL.md. All former #14 use cases now served by **#71 Table+Insight Panel** (`table_insight`). Updated Layout Diversity table, Opening Slide Priority Rule, and recommended slide structure. |
| 2.0.3 | 2026-03-19 | **3 Template Updates — Category M: Editorial Narrative**: (1) **#20 Before/After rewrite** (v2.0.1) — replaced BG_GRAY+NAVY color blocks with clean white-bg + black vertical divider + black circle `>` arrow + structured data rows (dict: label/brand/val/extra) + formula cards (dict: title/desc/cases with underline), new params: `corner_label`, `bottom_bar`, `left_summary`, `right_summary`, `right_summary_color`; (2) **#71 Table+Insight Panel** (NEW) — left data table (~60%) + middle CHEVRON shape icon (0.7") + right gray-bg (#F2F2F2) insight panel (~32%) with "启示：" title + `•` bullet points, supports `**bold**` markup in cells, self-adaptive row height; (3) **#72 Multi-Bar Panel Chart** (NEW) — 2-3 side-by-side bar panels with auto-numbered titles, CAGR trend arrows following actual bar-top slopes (RIGHT_ARROW shape, ~1.5px shaft, ±0.27" offset), per-bar value labels, green/red CAGR coloring. **OPENING SLIDE PRIORITY RULE** added: Slides 2-5 strongly prefer #71, #8, #14, #25. New Category M in layout-catalog.md. Total patterns: **72**. |
| 2.0.2 | 2026-03-19 | **Adaptive Row Height**: `data_table` / `vertical_steps` dynamically calculate row_h to prevent overflow. Font shrinks to Pt(10) for compact rows. |
| 2.0.1 | 2026-03-19 | **Before/After Rewrite**: White editorial layout with structured data rows (label/brand/val/extra), formula cards, left/right summaries. Fixed `set_ea_font` import in core.py. |
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
