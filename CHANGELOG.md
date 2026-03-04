# Changelog

All notable changes to this project will be documented in this file.

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
