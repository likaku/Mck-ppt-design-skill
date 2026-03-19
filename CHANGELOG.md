# Changelog

All notable changes to this project will be documented in this file.

## [2.0.0] - 2026-03-19

### Breaking Changes — Chart Rendering Engine
- **Donut Chart (#48)**: Rewritten from `math.cos/sin` + `add_rect()` block loops to **BLOCK_ARC** native shapes. Each segment = 1 shape (total: 4 shapes vs. hundreds of tiny squares). Center labels changed from NAVY/MED_GRAY to **WHITE** for readability against colored ring.
- **Pie Chart (#64)**: Rewritten from 2000+ `add_rect()` blocks to **BLOCK_ARC** with `inner_ratio=0` (solid sectors). Each segment = 1 shape (total: 4 shapes vs. 2778 blocks).
- **Gauge/Dial (#55)**: Rewritten from `add_rect()` arc + white circle overlay to **3 BLOCK_ARC** shapes (red/orange/green zones). Fixed horizontal orientation using PPT coordinate system directly (270°→0°→90° CW). Ring width controlled via `inner_ratio` parameter.

### Added
- **`add_block_arc()` helper function** — draws BLOCK_ARC preset shapes with precise angle/ring-width control via XML `adj` parameters (`adj1`=start angle, `adj2`=end angle, `adj3`=inner radius ratio). Documented in Guard Rails Rule 9.
- **Production Guard Rails Rule 9**: BLOCK_ARC Native Shapes for Circular Charts — mandatory for all donut/pie/gauge charts. Includes PPT angle convention documentation (CW from 12 o'clock, in 60000ths of degree), three usage patterns (donut ring, solid pie sector, horizontal rainbow gauge), and explicit anti-patterns.
- **5 new Common Issues** (Problems 16-20):
  - Problem 16: Donut/Pie charts made of hundreds of tiny rect blocks → use BLOCK_ARC
  - Problem 17: Gauge arc renders vertically instead of horizontally → use PPT CW coordinates
  - Problem 18: Donut center text unreadable → use WHITE font color
  - Problem 19: Chart elements overlapping title bar → minimum `Inches(1.3)` content start
  - Problem 20: Waterfall connector lines look like dots → ensure full-gap width

### Changed
- Updated wireframe diagrams for #48 and #64 to show "BLOCK_ARC ×4" instead of "rects" / "blocks"
- Georgia font used consistently for numeric values in charts (donut center labels, gauge scores, benchmark values)
- Guard Rails section header updated to "v1.9 / v2.0"

### Performance Impact
- **File size**: Donut chart slide ~80% smaller (eliminated hundreds of shapes)
- **Generation time**: Circular charts generate in <1s (was 30s-2min with block loops)
- **Shape count per chart**: 3-4 (was 100-2800)
- **Total shape count** for a 67-slide deck with donut+pie+gauge: reduced by ~3000 shapes

### Context
- All changes derived from 5 rounds of production testing on a 67-slide/70-template Tencent Annual Report PPT
- BLOCK_ARC approach validated through XML inspection: correct `adj1`/`adj2`/`adj3` values verified in generated PPTX
- Horizontal gauge verified: 270°(left) → 342° → 36° → 90°(right) produces correct rainbow arc

### Stats
- Common Issues: 15 → **20** (+5 new)
- Production Guard Rails: 8 → **9** (+1 new)
- Rewritten layout patterns: **3** (#48, #55, #64)
- New helper function: **1** (`add_block_arc()`)
- Estimated net SKILL.md change: ~+200 lines (new content) / ~-100 lines (removed block-loop code) = ~+100 net

## [1.10.4] - 2026-03-19

### Fixed
- **Cover Slide Title/Subtitle Overlap** (Problem 11) — When cover title contains `\n`, fixed-height textbox caused title to overflow into subtitle. Fix: compute title height dynamically from line count (`Inches(0.8 + 0.65 * (n_lines - 1))`), position subtitle/author/date relative to title bottom
- **Action Title Not Flush Against Separator** (Problem 12) — `add_action_title()` anchor changed from `MSO_ANCHOR.MIDDLE` to `MSO_ANCHOR.BOTTOM` so single-line titles sit flush against the separator line at `Inches(1.05)`
- **Checklist #61 Rows Overflowing Page** (Problem 13) — Fixed `row_h` caused 7+ rows to extend past page boundary. Fix: `row_h = min(Inches(0.85), available_h / n_rows)` with auto font-size reduction when rows are tight
- **Value Chain #67 Not Filling Content Area** (Problem 14) — Fixed `stage_w = Inches(2.0)` left ~2.5" whitespace on each side. Fix: `stage_w = (CW - arrow_w * (n-1)) / n` fills entire content width; stage height computed dynamically to fill vertical space
- **Closing #36 Bottom Line Too Short** (Problem 15) — Bottom decorative line changed from `Inches(3)` (partial) to `CW` (full content width)

### Added
- **Production Guard Rails Rule 8**: Dynamic Sizing for Variable-Count Layouts — mandatory dynamic dimension computation for layouts with variable item counts (horizontal: `(CW - gap*(N-1)) / N`, vertical: `min(max_h, available/N)`, multi-line: height from line count)
- **5 new Common Issues** (Problems 11-15) with code examples and anti-pattern warnings

### Changed
- Updated code examples for Pattern #1 (Cover), #36 (Closing), #61 (Checklist), #67 (Value Chain) to reflect the fixes
- Updated `add_action_title()` code example in Rule 5 to show `anchor=MSO_ANCHOR.BOTTOM`

### Stats
- Common Issues: 10 → **15** (+5 new)
- Production Guard Rails: 7 → **8** (+1 new)
- Updated code in 4 layout patterns + 1 helper function

## [1.10.3] - 2026-03-18

### Improved
- **Title Line Spacing Optimization** — titles (font_size ≥ 18pt) now use `p.line_spacing = 0.93` (multiple spacing) instead of `Pt(font_size.pt * 1.35)` (fixed point spacing)
  - Produces tighter, more professional title rendering — matches PowerPoint's "多倍行距 0.93" setting
  - Applies to: 22pt page titles, 28pt section divider titles, 18pt sub-headers
  - Body text (< 18pt) retains the existing `Pt(font_size.pt * 1.35)` fixed spacing to prevent CJK overlap
  - In OOXML: titles now emit `<a:lnSpc><a:spcPct val="93000"/>` instead of `<a:spcPts>`
- Updated Problem 5 documentation to reflect the two-tier line spacing strategy

### Thanks
- **冯梓航 Denzel** — for detailed feedback on title spacing refinement

## [1.10.2] - 2026-03-18

### Added
- **#54 Matrix Side Panel Variant** — new layout variant for Pattern #54 (Risk/Heat Matrix) that pairs a compact 3×3 grid with an adjacent insight panel
  - Compact grid dimensions: `cell_w=2.15"` (down from 3.0"), `axis_label_w=0.65"` (down from 1.8")
  - Side panel gains ~4.2" width — sufficient for 6+ bullet items with comfortable reading
  - Includes ASCII wireframe, layout math code, and dark summary box at panel bottom
  - New rule: panel must never shrink below `Inches(2.5)`

### Context
- When using #54 with a side panel for "Key Changes" or "Action Items", the default 3.0" cell width consumed ~9" of the 11.7" content area, crushing the panel to ~1.4" — unreadable
- The new compact variant allocates ~60% to grid + ~38% to panel with ~2% gap, achieving zero whitespace waste

## [1.10.1] - 2026-03-18

### Fixed
- **YAML Frontmatter compatibility** — fixed "malformed YAML frontmatter in SKILL.md" error when installing via Claude's "Copy to your skills" button
  - Claude's SKILL.md parser only supports `name` and `description` fields in frontmatter
  - Removed unsupported fields from frontmatter: `license`, `version`, `author`, `homepage`, `user-invocable`, `allowed-tools`, `metadata`
  - Moved `metadata` field (contained emoji `📊` and inline JSON) which caused YAML parse failures
  - Changed `description` from single-line quoted string (673 chars) to YAML folded block scalar (`>-`) for better readability and compatibility
  - Relocated project metadata (version, license, author, tools) to a blockquote section in the document body — no information lost

## [1.10.0] - 2026-03-18

### Added
- **Channel Delivery** — new section in SKILL.md enabling automatic file delivery to messaging channels (Feishu/飞书, Telegram, WhatsApp, Discord, Slack, Signal, etc.)
  - `deliver_to_channel()` Python helper function — sends generated PPTX back to user's chat via `openclaw message send --media`
  - Graceful fallback — if `openclaw` CLI is not available (IDE, CI), skips silently and prints local path
  - Caption includes slide count and file size
  - Channel-specific compatibility table (all channels support document type, max 100MB)
- **Updated `minimal_example.py`** — now calls `deliver_to_channel()` after save + cleanup
  - Added `shutil` and `subprocess` imports
  - Post-generation flow: save → cleanup → deliver → confirm

### Context
- Users interacting via OpenClaw channels (Feishu, Telegram, etc.) could not receive generated PPTX files — they were only saved to server disk
- OpenClaw's media pipeline supports document files up to 100MB via `openclaw message send --media <path>`
- `.pptx` is classified as "document" type in OpenClaw, well within size limits for any presentation

### Stats
- Net addition: ~80 lines in SKILL.md (Channel Delivery section)
- Updated 2 example scripts (scripts/ and examples/)

## [1.9.0] - 2026-03-15

### Added
- **Production Guard Rails** — 7 mandatory rules derived from iterative production feedback, added to Presentation Planning section:
  - **Rule 1: Spacing Between Content and Bottom Bars** — minimum 0.15" gap between last content block and any bottom summary/action bar
  - **Rule 2: Content Overflow Protection** — right margin + bottom margin boundary checks, text inset within container boxes
  - **Rule 3: Bottom Whitespace Elimination** — charts/content must fill vertical space, maximum 0.3" gap above bottom bar
  - **Rule 4: Legend Color Consistency** — colored `add_rect()` squares mandatory for chart legends, no text-only "■" symbols
  - **Rule 5: Title Style Consistency** — `add_action_title()` is the ONLY approved title style; `add_navy_title_bar()` deprecated for content slides
  - **Rule 6: Axis Label Centering** — axis labels in 2×2 matrices must be centered on actual grid dimensions
  - **Rule 7: Image Placeholder Slide Requirement** — presentations with 8+ slides must include at least 1 image placeholder slide; triple-border placeholder style documented

- **Code Efficiency Guidelines** — new subsection in Best Practices:
  - Extracted layout constants pattern (CONTENT_TOP, BOTTOM_BAR_Y, etc.)
  - Reusable helper functions for legends and bottom bars
  - Standard short variable name table (at/ar/ahl/ao/aat/asrc/apn)
  - Batch data as lists-of-tuples pattern
  - Auto-incrementing page number counter pattern

- **5 new Common Issues** (Problems 6-10):
  - Problem 6: Content overflowing container boxes
  - Problem 7: Chart legend color mismatch
  - Problem 8: Inconsistent title bar styles
  - Problem 9: Axis labels off-center in matrix charts
  - Problem 10: Bottom whitespace under charts

### Context
- All rules derived from actual production feedback during generation of a 19-slide AI industry report
- Issues discovered across multiple review iterations: spacing between table/bottom bars, text overflow beyond margins, legend colors not matching chart colors, inconsistent title styles (navy bar vs white background), axis labels misaligned in 2×2 matrices, lack of image placeholder pages, and excessive bottom whitespace
- Code efficiency guidelines extracted from analysis of a 535-line production script, identifying common patterns that can reduce code by ~15-20%

### Stats
- Net addition: ~250 lines in SKILL.md (Production Guard Rails + Code Efficiency + Common Issues)
- Common Issues: 5 → **10** (+5 new)
- New section: Production Guard Rails (7 rules with code examples)
- New subsection: Code Efficiency Guidelines (5 patterns with code examples)

## [1.8.0] - 2026-03-15

### Added
- **Category I: Image + Content Layouts** — 8 new layout patterns for slides containing images:
  - **#40 Content + Right Image** — text left, image placeholder right
  - **#41 Left Image + Content** — image left, text right
  - **#42 Three Images + Descriptions** — three-column image comparison
  - **#43 Image + Four Key Points** — central image with surrounding callout points
  - **#44 Full-Width Image with Overlay** — hero image with semi-transparent text overlay
  - **#45 Case Study with Image** — SAR format with supporting visual + KPI boxes
  - **#46 Quote with Background Image** — keynote-style quote slide
  - **#47 Goals with Illustration** — OKR/target list with supporting illustration

- **Category J: Advanced Data Visualization** — 9 new chart patterns drawn with pure `add_rect()`:
  - **#48 Donut Chart** — part-of-whole with center label (block approximation)
  - **#49 Waterfall Chart** — revenue/profit bridge analysis
  - **#50 Line / Trend Chart** — multi-series time-series trends
  - **#51 Pareto Chart** — 80/20 analysis with cumulative line + threshold
  - **#52 Progress Bars / KPI Tracker** — OKR tracking with status indicators
  - **#53 Bubble / Scatter Plot** — two-variable comparison with size encoding
  - **#54 Risk / Heat Matrix** — impact vs likelihood grid
  - **#55 Gauge / Dial Chart** — single KPI health indicator
  - **#56 Harvey Ball Status Table** — multi-criteria evaluation matrix

- **Category K: Dashboard Layouts** — 2 data-dense executive dashboard patterns:
  - **#57 Dashboard: KPIs + Chart + Takeaways** — top KPI cards, middle chart, bottom insights
  - **#58 Dashboard: Table + Chart + Factoids** — left table, right chart, bottom factoid cards

- **Category L: Visual Storytelling & Special** — 12 new visual narrative patterns:
  - **#59 Stakeholder Map** — influence vs interest matrix
  - **#60 Issue / Decision Tree** — MECE logic tree with 3 levels
  - **#61 Five-Row Checklist / Status** — task completion tracker with progress bar
  - **#62 Metric Comparison Row** — before/after with delta badges
  - **#63 Icon Grid** — 3×2 capability/feature grid with icon placeholders
  - **#64 Pie Chart** — simple part-of-whole (block approximation)
  - **#65 SWOT Analysis** — 2×2 color-coded strategic analysis
  - **#66 Agenda / Meeting Outline** — timed agenda with speaker assignments
  - **#67 Value Chain / Horizontal Flow** — end-to-end pipeline with KPI boxes
  - **#68 Two-Column Image + Text Grid** — 2×2 visual catalog
  - **#69 Numbered List with Side Panel** — recommendations + highlight panel
  - **#70 Stacked Area Chart** — cumulative trends over time

- **`add_image_placeholder()` helper function** — draws gray placeholder rectangle with crosshair lines and label for image positions; users replace with real images after generation
- **Image Priority Rule** — when content involves case studies, product showcases, or visual storytelling, prefer Image+Content layouts (#40-#47, #68)
- **Expanded Layout Diversity table** — 7 new content-type-to-layout mappings for image, composition, risk, recommendations, dashboard, stakeholder, and agenda content types

### Context
- New layouts based on systematic analysis of McKinsey PowerPoint Template 2023 (679 pages) — extracted text from all pages, searched by keywords (chart: 91 pages, image: many, pie: 41, flow: 57, process: 61), exported 30 key pages as screenshots for visual reference
- All new chart patterns maintain the pure `add_rect()` / `add_oval()` drawing approach — no matplotlib, no chart objects, no connectors
- Image placeholders use a standardized gray rect + crosshair + label convention for consistency
- New layouts are fully additive — zero modifications to existing patterns #1-#39

### Stats
- Layout patterns: 39 → **70** (+31 new)
- Categories: 8 → **12** (+4 new: I, J, K, L)
- Net addition: ~3000+ lines in SKILL.md

## [1.7.0] - 2026-03-13

### Added
- **Category H: Data Charts** — 3 new chart layout patterns drawn with pure `add_rect()`, no matplotlib dependency:
  - **#37 Grouped Bar Chart** — multi-category comparison across time points (e.g., sentiment distribution, multi-product sales)
  - **#38 Stacked Bar Chart** — part-to-whole composition over time (e.g., channel revenue mix, sentiment ratio)
  - **#39 Horizontal Bar Chart** — category ranking with long labels (e.g., feature usage rate, department performance)
- **Chart Priority Rule** in Layout Diversity table — when data contains dates + numeric values/percentages, chart patterns (#37/#38/#39) are mandatory over text-based layouts
- **Chart trigger signals** — automatic detection rules: date+percentage combos, `████` progress bars, trend-related keywords, ≥3 rows with categories and values
- Color assignment for charts uses the four-color scheme: NAVY / LINE_GRAY / MED_GRAY / ACCENT_BLUE

### Stats
- Layout patterns: 36 → **39** (3 new)
- Categories: 7 → **8** (added Category H)
- Net addition: ~372 lines in SKILL.md

## [1.6.0] - 2026-03-08

### Added
- **Accent Color System** — 4 accent colors (Blue #006BA6, Green #007A53, Orange #D46A00, Red #C62828) with paired light backgrounds for multi-item visual differentiation. Includes usage rules, constants, and code snippets
- **Presentation Planning section** — comprehensive guidance for slide structure, layout selection, and content density:
  - **Recommended Slide Structures**: Standard (10-12 slides) and Short (6-8 slides) templates with specific pattern assignments
  - **Layout Diversity Requirement**: content-type-to-layout matching table; consecutive slides must use different patterns
  - **Content Density Requirements**: minimum 3 visual blocks per slide, ≥50% area utilization, full-sentence Action Titles
  - **Mandatory Slide Elements**: every content slide must include Action Title, source attribution, and page number
  - **`add_page_number()` helper function**: displays "N/Total" at bottom-right
- Minimum slide count rule: 8 slides for any substantive presentation

### Context
- Based on comparative analysis of the same Skill prompt across 4 LLM models (Opus 4.6 / Minimax 2.5 / Hunyuan 2 think / GLM5)
- Opus produced 402 shapes, 15 colors, diverse layouts; other models produced 65-145 shapes, 7 colors, repetitive layouts
- These additions target the structural gaps that caused weaker models to produce sparse, monotonous output
- Expected to close ~70% of the quality gap between Opus and other models

## [1.5.0] - 2026-03-08

### Fixed
- **Critical: Chinese multi-line text overlapping** — `add_text()` now sets `p.line_spacing = Pt(font_size.pt * 1.35)` for every paragraph, mapping to `<a:lnSpc><a:spcPts>` in OOXML
- Previously only `space_before` (paragraph spacing) was set, but the actual line height (`lnSpc`) was unset, causing word-wrapped Chinese lines to render on top of each other

### Added
- Problem 5 in Common Issues: documents the CJK line overlap root cause and fix

### Stats
- Net addition: ~10 lines — 1 line of code fix + documentation

## [1.4.0] - 2026-03-06

### Changed
- **Merged `add_text()` and `add_multiline()` into a single unified function** — pass `str` for single line, `list` for multi-line. Adds `line_spacing=Pt(6)` and `anchor` parameters
- Updated all 36 layout template examples to use the unified `add_text()` function
- Parameter renamed: `line_spacing_pt=N` → `line_spacing=Pt(N)` for consistency with python-pptx API

### Removed
- `add_multiline()` function (replaced by `add_text()` with list support)
- DEPRECATED connector explanation and old code sample (lines 173-179)
- v1.1 improvement paragraph (line 214)
- Refining Existing Presentations section (was 22 lines of generic guidance)
- Error Handling section (4 items consolidated into Common Issues, removing 3 duplicates)
- Problem 3 "Lines Appearing With Shadows" from Common Issues (duplicate of "never use connectors" rule)
- Verification code block from Problem 1 (full_cleanup already well-documented above)

### Stats
- Net reduction: **109 lines, ~4.2KB** → lower token consumption per generation

## [1.3.0] - 2026-03-04

### Added
- ClawHub-compatible `metadata` field (declares `python3`/`pip` dependencies)
- `homepage` field pointing to GitHub repository
- `Edge Cases` section: large presentations, font availability, slide dimensions, LibreOffice compatibility
- `Error Handling` section: file repair, Chinese rendering, module errors, alignment issues
- `references/color-palette.md` — quick color & font-size reference
- `references/layout-catalog.md` — all 36 layout types at a glance
- `scripts/` directory mirroring example code for ClawHub convention

### Changed
- Rewrote `description` for ClawHub discoverability: verb-first, `Use when` trigger pattern, keyword coverage (pitch deck, strategy, quarterly review, board meeting, etc.)
- Expanded `When to Use` with business scenario keywords
- Version bumped to 1.3.0

## [1.2.0] - 2026-03-04

### Fixed
- Circle shape (`add_oval()`) number font now matches body text — added `font_name='Arial'` and `set_ea_font()` for consistent typography
- Circle numbers simplified from `01, 02, 03` to `1, 2, 3` (no leading zeros)

### Changed
- Removed product-specific references from skill description; skill is now fully generic for any professional PPT
- Skill name updated to `mck-ppt-design` for generic usage

## [1.1.0] - 2026-03-03

### Breaking Changes
- `add_line()` **deprecated** — replaced by `add_hline()` (thin rectangle, no connector)
- `add_circle_label()` **renamed** to `add_oval()` with `bg`/`fg` color parameters
- `cleanup_theme()` **replaced** by `full_cleanup()` (sanitizes all slide + theme XML)
- `add_multiline()` removed `bullet` parameter; use `'• '` prefix in text directly

### Added
- `_clean_shape()` — inline p:style removal, called automatically by `add_rect()` and `add_oval()`
- `add_hline()` — draws horizontal lines as thin rectangles (zero connector usage)
- `full_cleanup()` — nuclear XML sanitization: removes ALL `<p:style>` from every slide + theme effects
- Three-layer defense against file corruption documented

### Fixed
- **Critical**: 62+ shapes carrying `effectRef idx="2"` caused "File needs repair" in PowerPoint
- Connectors' `<p:style>` could not be reliably removed; eliminated connectors entirely

## [1.0.0] - 2026-03-02

### Added
- Initial release of McKinsey-style PPT Design Skill
- Complete color palette specification (NAVY, BLACK, DARK_GRAY, MED_GRAY, LINE_GRAY, BG_GRAY, WHITE)
- Typography hierarchy system (44pt cover to 9pt footnote)
- Line treatment standards with shadow removal
- Post-save theme cleanup for removing OOXML shadow/3D effects
- Layout patterns: Cover, Action Title, Table, Three-Column Overview
- Complete Python helper functions (add_text, add_line, add_rect, add_circle_label, etc.)
- Common issues & solutions documentation
- Minimal example script
