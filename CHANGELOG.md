# Changelog

All notable changes to this project will be documented in this file.

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
