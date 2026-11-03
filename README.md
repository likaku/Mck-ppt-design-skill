<div align="center">

# MCK PPT Design Skill

**Your AI agent writes the story. This skill builds the deck: consulting-grade, fully editable, checked by code.**

[![Stars](https://img.shields.io/github/stars/likaku/Mck-ppt-design-skill?style=flat&logo=github)](https://github.com/likaku/Mck-ppt-design-skill/stargazers)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![CI](https://github.com/likaku/Mck-ppt-design-skill/actions/workflows/ci.yml/badge.svg)](https://github.com/likaku/Mck-ppt-design-skill/actions)
[![Release](https://img.shields.io/github/v/release/likaku/Mck-ppt-design-skill)](https://github.com/likaku/Mck-ppt-design-skill/releases)

[English](README.md) · [简体中文](README.zh-CN.md) · [Layout gallery](docs/GALLERY.md) · [Architecture](docs/ARCHITECTURE.md) · [Harness mode](docs/HARNESS.md)

</div>

<table>
<tr>
<td><img alt="Cover" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" /></td>
<td><img alt="Strategy analysis" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" /></td>
<td><img alt="Data table" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" /></td>
</tr>
<tr>
<td><img alt="Four-column framework" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" /></td>
<td><img alt="Color system" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" /></td>
<td><img alt="Executive summary" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" /></td>
</tr>
</table>

## Why this exists

Ask an AI agent for a deck and you usually get one of two things: text boxes stacked on a white page, or a pretty picture you cannot edit. This skill gives the agent a third path.

- **Real .pptx, every shape editable.** Native shapes and text, no screenshots. Open in PowerPoint, WPS or Keynote and keep working.
- **62 layouts that look like a consulting deck.** Action titles, source lines, restrained navy and grey, flat design, no shadows.
- **The agent decides, Python draws.** `eng.donut(...)` is about 20 tokens of agent output and a deterministic rendering step. A 30-slide deck drops from 40,000-60,000 output tokens to 9,000-12,000 (measured in our v1 to v2 rewrite).
- **Checked by code, not by the agent's opinion.** Overflow, collisions, legend and font consistency are tested by scripts. Delivery is blocked until `gate_result.json` says `passed: true`.
- **Chinese works.** East Asian fonts are injected on every run, so 楷体 and 微软雅黑 do not turn into boxes.

## Try it in 60 seconds

No API key, no network:

```bash
pip install mck-ppt            # or: pip install python-pptx lxml && git clone this repo
python -m mck_ppt demo         # writes mck_demo.pptx (13 slides, passes built-in QA)
python -m mck_ppt layouts      # list every layout with a one-line description
```

## Use it from your agent

| Agent | Install |
|---|---|
| **Claude Code** | `/plugin marketplace add likaku/Mck-ppt-design-skill` then `/plugin install mck-ppt-design@mck-ppt` |
| **Any agent that reads skills** | `npx skills add likaku/Mck-ppt-design-skill` |
| **WorkBuddy / CodeBuddy** | Search `mck-ppt-design` in the skill market, or copy this folder to `~/.workbuddy/skills/` |
| **Cursor / Continue / any LLM** | Add `SKILL.md` as a project rule and run `pip install mck-ppt` |

Then ask in plain language:

> Make a 12-slide board deck on why our Q1 growth now depends on one channel. Use action titles, one chart per argument, and a source line on every slide.

> 用麦肯锡风格做一份 15 页的年度战略汇报：先给执行摘要，再分三章，最后列出需要决策的四件事。

The agent follows five stages: brief, structure (`outline.json`), content (`content.json`), render, deliver. Two gates are scripts, not self-assessment. See [Harness mode](docs/HARNESS.md).

## Or use it as a library

```python
from mck_ppt import MckEngine
from mck_ppt.constants import *

eng = MckEngine(total_slides=4)
eng.cover(title='Q1 strategy review', subtitle='Board presentation', date='2026')
eng.table_insight(
    title='Three shifts are restructuring the market',
    headers=['Dimension', 'Before', 'After'],
    rows=[['Distribution', 'Offline-first', 'Digital-first'], ['Pricing', 'Cost-plus', 'Value-based']],
    insights=['Digital channels now drive 60% of acquisition', 'Value pricing unlocks about 3x margin'])
eng.donut(title='Revenue mix', segments=[(0.45, NAVY, 'Product'), (0.35, ACCENT_BLUE, 'Service'), (0.20, ACCENT_GREEN, 'Other')])
eng.closing(title='Thank you')
eng.save('deck.pptx')
```

## How it compares

| | This skill | Free-form python-pptx / PptxGenJS skills | HTML-slide skills |
|---|---|---|---|
| Output | Editable .pptx | Editable .pptx | HTML / PDF |
| Layout source | 62 tested layout methods | The model draws coordinates | The model writes CSS |
| Quality gate | Code (`gate_check.py`), blocks delivery | Usually visual check by the model | Usually none |
| Chinese typography | Handled in engine | Per prompt | Per CSS |
| Best for | Strategy, board, consulting-style decks | Custom one-off designs | Talks you present in a browser |

Pick an HTML-slide skill if you present from a browser. Pick this one if the file has to end up in someone's PowerPoint.

## Layouts

62 active layouts. Full list with signatures: `python -m mck_ppt layouts`, or the [gallery](docs/GALLERY.md).

| Category | Layouts |
|---|---|
| **Structure** | `cover`, `section_divider`, `toc`, `agenda`, `appendix_title`, `closing` |
| **Key numbers** | `big_number`, `two_stat`, `three_stat`, `metric_cards`, `scorecard`, `kpi_tracker`, `rag_status`, `data_table`, `harvey_ball_table` |
| **Frameworks** | `table_insight`, `matrix_2x2`, `pyramid`, `process_chevron`, `temple`, `swot`, `value_chain`, `risk_matrix`, `stakeholder_map`, `decision_tree` |
| **Comparison** | `side_by_side`, `before_after`, `pros_cons`, `metric_comparison` |
| **Narrative** | `executive_summary`, `key_takeaway`, `four_column`, `two_column_text`, `case_study`, `action_items`, `checklist`, `quote`, `meet_the_team`, `numbered_list_panel`, `icon_grid` |
| **Timeline** | `timeline`, `vertical_steps` |
| **Charts** | `donut`, `waterfall`, `line_chart`, `pareto`, `bubble`, `grouped_bar`, `stacked_bar`, `horizontal_bar`, `stacked_area`, `multi_bar_panel`, `dashboard_kpi_chart`, `dashboard_table_chart` |
| **Image layouts** | `content_right_image`, `three_images`, `image_four_points`, `full_width_image`, `case_study_image`, `quote_bg_image`, `goals_illustration`, `two_col_image_grid` |

## Guard rails (selected)

Hard rules learned from production use, enforced in code. Full list in [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md#-production-guard-rails).

- No connectors (they trigger PowerPoint repair prompts); lines are thin rectangles.
- Shadows and 3D are stripped after save, so the file never opens in "repair" mode.
- Same-row shapes keep the same font size; text never touches separator lines.
- Circular charts use native `BLOCK_ARC`, 3-4 shapes instead of thousands of rectangles.

## Project layout

```
SKILL.md          entry for agents (about 250 lines)
mck_ppt/          engine, QA, review pipeline, DeckBuilder, demo, CLI
references/       per-stage knowledge, layout matrix, gate scripts
experiences/      lessons the agent appends after fixing a pattern-level bug
docs/             architecture, harness, gallery, growth playbook
```

## Roadmap

- [ ] Rendered preview for every layout in the gallery
- [ ] `mck-ppt from-outline outline.json` one-shot CLI
- [ ] Theme switch (navy / graphite / warm) without touching layouts
- [ ] Import an existing .pptx and restyle it
- [ ] More layouts from community requests: [open a request](https://github.com/likaku/Mck-ppt-design-skill/issues/new?template=layout_request.yml)

## Community

- Made a deck with it? [Show it](https://github.com/likaku/Mck-ppt-design-skill/issues/new?template=showcase.yml). Good ones go in the README.
- Questions: [Discussions](https://github.com/likaku/Mck-ppt-design-skill/discussions)
- Discord: [join](https://discord.gg/SaFybFAT) · WeChat group (200+ members): add `Kashinmoto`
- If it saved you an evening, a star helps other people find it.

## Version history

See [CHANGELOG.md](CHANGELOG.md). Latest: **v2.4.0** adds `python -m mck_ppt demo`, `pip install mck-ppt`, CI, and a slimmer README.

---

<div align="center"><sub>Apache 2.0 · © 2024-2026 <strong>Kaku Li</strong> · <a href="NOTICE">NOTICE</a></sub></div>
