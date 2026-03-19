<div align="center">

# McKinsey PPT Design Skill

**A complete McKinsey-style PowerPoint design system for AI agents**
<br/>Generate professional, consultant-grade presentations from scratch using `python-pptx` | v1.10.3

[English](#overview) · [中文说明](#中文说明)

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21+-orange.svg)](https://python-pptx.readthedocs.io)
[![GitHub stars](https://img.shields.io/github/stars/likaku/Mck-ppt-design-skill?style=social)](https://github.com/likaku/Mck-ppt-design-skill)

</div>

---

## Overview

**McKinsey PPT Design Skill** encodes a complete consulting-firm PowerPoint design specification into a single document (`SKILL.md`). When loaded by AI agents (Claude, GPT, Cursor, Codebuddy, etc.), it enables consistent, professional slide generation — every time.

**Keywords**: McKinsey PowerPoint template, consulting slide design, AI presentation generator, python-pptx automation, professional PPT design system, business slide template, pitch deck generator, strategy deck, quarterly review slides, board meeting presentation

### What It Does

- 🎨 **70 layout patterns** across 12 categories — from title slides to dashboards, SWOT analyses, waterfall charts, and more
- 📐 **Strict McKinsey design system** — flat design, no shadows, no 3D, consistent typography hierarchy
- 🛡️ **Three-layer file corruption defense** — eliminates `p:style`, shadow, and 3D artifacts that cause PowerPoint repair prompts
- 🔤 **CJK + Latin font handling** — proper Chinese/Japanese/Korean rendering with KaiTi / Georgia / Arial
- 📊 **Hand-drawn charts** — donut, waterfall, line, Pareto, bubble, Harvey Ball, and more — all built with `add_rect()`, no chart XML needed
- 🖼️ **Image placeholder system** — gray placeholder boxes with crosshairs for easy replacement
- 🚀 **Production guard rails** — 7 mandatory rules preventing common AI generation mistakes
- 📨 **Channel delivery** — auto-send generated PPTX via Feishu, Telegram, Slack, Discord, WhatsApp

### Sample Output

| Cover Page | Content Page | Table Page |
|:------:|:------:|:------:|
| <img width="600" alt="Cover" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" /> | <img width="600" alt="Content" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" /> | <img width="600" alt="Table" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" /> |
| **4-Column Layout** | **Color System** | **Summary Page** |
| <img width="600" alt="4-Column" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" /> | <img width="600" alt="Colors" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" /> | <img width="600" alt="Summary" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" /> |

---

## Quick Start

```bash
# Install dependencies
pip install python-pptx lxml

# Run the minimal example
cd scripts && python minimal_example.py

# Install from ClawHub (recommended)
npx clawhub@latest install mck-ppt-design

# Or manual install for Claude
mkdir -p ~/.claude/skills/mck-ppt-design
cp SKILL.md ~/.claude/skills/mck-ppt-design/
```

### Compatibility

| AI Agent | Status | Install Method |
|----------|--------|----------------|
| **Claude** (Anthropic) | ✅ Fully supported | ClawHub or manual SKILL.md |
| **Cursor** | ✅ Fully supported | Add as project rule |
| **Codebuddy** | ✅ Fully supported | Load as skill |
| **GPT / ChatGPT** | ✅ Works with system prompt | Paste SKILL.md content |
| **Any LLM** | ✅ Universal | Feed SKILL.md as context |

---

## Design Principles

| Principle | Description |
|:----------|:------------|
| **Minimalism** | Remove all non-essential visual elements — no gradients, no decoration |
| **Flat Design** | No shadows, no 3D, no reflections — pure solid fills |
| **Strict Hierarchy** | Title 22pt → Subtitle 18pt → Body 14pt → Footnote 9pt |
| **Global Consistency** | Unified color palette, fonts, spacing across every slide |

---

## Color System

| Name | Swatch | Hex | Usage |
|:-----|:------:|:---:|:------|
| **NAVY** | ![](docs/colors/navy.png) | `#051C2C` | Primary — titles, circular indicators, TOC highlights |
| **BLACK** | ![](docs/colors/black.png) | `#000000` | Dividers, title underlines, table header lines |
| **DARK_GRAY** | ![](docs/colors/dark-gray.png) | `#333333` | Body text |
| **MED_GRAY** | ![](docs/colors/med-gray.png) | `#666666` | Secondary text, labels, source annotations |
| **LINE_GRAY** | ![](docs/colors/line-gray.png) | `#CCCCCC` | Table row separators |
| **BG_GRAY** | ![](docs/colors/bg-gray.png) | `#F2F2F2` | Background panels, takeaway areas |

**Accent colors** (v1.6.0+) — for 3+ parallel items:

| Name | Hex | Light BG | Usage |
|:-----|:---:|:--------:|:------|
| **ACCENT_BLUE** | `#006BA6` | `#E3F2FD` | 1st item emphasis |
| **ACCENT_GREEN** | `#007A53` | `#E8F5E9` | 2nd item emphasis |
| **ACCENT_ORANGE** | `#D46A00` | `#FFF3E0` | 3rd item emphasis |
| **ACCENT_RED** | `#C62828` | `#FFEBEE` | 4th item / alert |

---

## Layout Categories (70 Patterns)

| # | Category | Patterns | Examples |
|---|----------|----------|----------|
| A | Structure | #1–#7 | Title, divider, TOC, agenda, executive summary |
| B | Data Display | #8–#15 | Tables, KPI cards, comparison panels |
| C | Frameworks | #16–#23 | Process flows, pyramids, matrices, Venn |
| D | Comparison | #24–#29 | Before/after, side-by-side, scorecard |
| E | Narrative | #30–#33 | Case study, quote, key findings |
| F | Timeline | #34–#39 | Horizontal, vertical, milestone, roadmap |
| G | Team/Org | — | Org chart, team profiles |
| H | Charts | — | Bar, stacked bar, grouped bar |
| I | Image+Content | #40–#47 | Photo+text, 3-photo comparison, full-bleed |
| J | Advanced Viz | #48–#56 | Donut, waterfall, line, Pareto, bubble, Harvey Ball |
| K | Dashboard | #57–#58 | Executive dashboards |
| L | Visual Story | #59–#70 | Stakeholder map, decision tree, SWOT, pie chart |

See [references/layout-catalog.md](references/layout-catalog.md) for the full catalog with ASCII wireframes.

---

## Core Technology

### Three-Layer File Corruption Defense (v1.1)

python-pptx auto-attaches `<p:style>` elements referencing `outerShdw`, `effectRef`, etc., causing PowerPoint to prompt for repair. This skill eliminates the issue with three defenses:

1. **No connectors** — all lines drawn as ultra-thin rectangles (`add_hline()`), preventing connector `p:style`
2. **Inline cleanup** — every `add_rect()` / `add_oval()` immediately calls `_clean_shape()` to remove `p:style`
3. **Post-save full wash** — `full_cleanup()` traverses all slide XML + theme XML, stripping all `p:style`, shadow, and 3D nodes

### CJK Font Handling

All paragraphs containing Chinese characters call `set_ea_font(run, 'KaiTi')` to set the East Asian font — otherwise Chinese renders with default Latin fonts.

---

## Project Structure

```
├── SKILL.md                 # Core design specification (268KB)
├── LICENSE                  # Apache 2.0
├── CHANGELOG.md             # Version history
├── scripts/
│   ├── minimal_example.py   # 2-page demo
│   └── requirements.txt     # Dependencies
├── references/
│   ├── color-palette.md     # Color quick-reference
│   └── layout-catalog.md    # 70 layout catalog
└── examples/
    ├── minimal_example.py   # 2-page demo (legacy path)
    └── requirements.txt
```

---

## Recent Updates

> ### v2.0.0 — BLOCK_ARC Chart Engine 🎉
>
> - **Major chart rendering rewrite** — Donut (#48), Pie (#64), and Gauge (#55) charts now use native BLOCK_ARC shapes
>   - 3-4 shapes per chart instead of 100-2800 tiny rect blocks
>   - File size reduced 60-80%, generation time from 2min → <1s
>   - Pixel-perfect arcs with no gaps or jagged edges
> - **New `add_block_arc()` helper** — precise control over arc angles and ring width via XML adj parameters
> - **Guard Rail Rule 9**: Mandatory BLOCK_ARC for all circular charts
> - **5 new Common Issues** (Problems 16-20): block-chart migration, vertical gauge fix, center text readability, title overlap, waterfall connectors
> - Derived from 5 rounds of production testing on a 67-slide Tencent Annual Report PPT
>
> See [CHANGELOG.md](CHANGELOG.md)

> ### v1.10.3 — Title Line Spacing Optimization
>
> - **Title spacing refined** — Titles (≥18pt) now use `0.93` multiple spacing instead of fixed `Pt(fs*1.35)`, producing tighter, more professional rendering
>   - Applies to: 22pt page titles, 28pt section dividers, 18pt sub-headers
>   - Body text (<18pt) retains fixed Pt spacing for CJK overlap prevention
>   - Maps to PowerPoint's "Multiple spacing 0.93" setting
> - Thanks to **冯梓航 Denzel** for detailed feedback 🙏
>
> See [CHANGELOG.md](CHANGELOG.md)

<details>
<summary><b>Earlier versions</b></summary>

> ### v1.10.4 — Bug Fixes + Dynamic Sizing
> - 5 bug fixes: cover overlap, title anchor, checklist overflow, value chain fill, closing line width
> - Guard Rail Rule 8: dynamic sizing for variable-count layouts

> ### v1.10.2 — #54 Matrix Side Panel Variant
>
> - **Added side-panel layout for Heat Matrix** — compact 3×3 grid (~60% width) + insight panel (~38% width)
>   - Grid cell width from `3.0"` to `2.15"`, Y-axis label area from `1.8"` to `0.65"`
>   - Side panel expands from ~1.4" (unreadable) to ~4.2" (fits 6+ entries)
>
> See [CHANGELOG.md](CHANGELOG.md)

<details>
<summary><b>Earlier versions</b></summary>

> ### v1.10.1 — YAML Frontmatter Fix
> - Fixed Claude installation error — SKILL.md parser only supports `name` + `description` frontmatter fields

> ### v1.10.0 — Channel File Delivery
> - Added `deliver_to_channel()` — auto-send PPTX via Feishu, Telegram, WhatsApp, Discord, Slack

> ### v1.9.0 — Production Guard Rails
> - 7 mandatory rules from real production feedback: spacing protection, overflow detection, legend consistency, etc.

> ### v1.8.0 — Layout Expansion (39 → 70)
> - 31 new professional layouts across 4 new categories (Image+Content, Advanced Viz, Dashboard, Visual Story)

See [CHANGELOG.md](CHANGELOG.md) for the complete history.

</details>

---

## Community

<table>
<tr>
    <td align="center" width="50%" valign="top">
      <strong>WeChat Group / 微信交流群</strong><br/><br/>
      <img width="180" src="https://github.com/user-attachments/assets/d4eb704e-3825-4380-ac54-2fbbe4c993ce" alt="WeChat Group" />
    </td>
    <td align="center" width="50%" valign="top">
      <strong>Discord</strong><br/><br/>
      <a href="https://discord.gg/SaFybFAT">
        <img src="https://img.shields.io/badge/Discord-Join_Community-5865F2?style=for-the-badge&logo=discord&logoColor=white" alt="Discord" />
      </a>
      <br/><br/>
      <span>Click above to join</span>
    </td>
  </tr>
</table>

---

## Requirements

Python 3.8+ · python-pptx ≥ 0.6.21 · lxml ≥ 4.9.0

---

## Contributing

Issues and PRs welcome! Contribution ideas:

- New layout patterns (timeline variants, 2×2 matrices, etc.)
- Extended color themes (dark mode, brand customization)
- Additional examples and documentation translations

---

## 中文说明

<details>
<summary><b>点击展开中文文档</b></summary>

### 简介

**McKinsey PPT Design Skill**（麦肯锡 PPT 设计技能）是一套完整的咨询公司风格 PowerPoint 设计体系。将完整的麦肯锡设计规范编码为一份文档（`SKILL.md`），AI 读取后即可持续输出风格统一的专业 PPT。

### 它解决什么问题

- 手动排版 PPT 耗时耗力，团队间设计风格难以统一
- `python-pptx` 默认生成的文件带阴影 / 3D / `p:style` 引用，PowerPoint 打开报错或提示修复
- 中文字体渲染需要特殊处理，否则显示异常
- AI 生成的 PPT 缺乏专业设计感，每次产出质量不一致

### 设计原则

| 原则 | 说明 |
|:-----|:-----|
| **极简主义** | 移除一切非必要视觉元素，无渐变、无装饰 |
| **扁平设计** | 无阴影、无 3D、无反射，纯实色填充 |
| **严格层次** | 标题 22pt → 子标题 18pt → 正文 14pt → 脚注 9pt |
| **全局一致** | 统一色板、字体、间距，贯穿每一页 |

### 快速上手

```bash
# 1. 安装依赖
pip install python-pptx lxml

# 2. 运行示例
cd scripts && python minimal_example.py

# 3. 从 ClawHub 安装（推荐）
npx clawhub@latest install mck-ppt-design

# 或手动安装
mkdir -p ~/.claude/skills/mck-ppt-design
cp SKILL.md ~/.claude/skills/mck-ppt-design/
```

### 参与贡献

欢迎提交 Issue 和 Pull Request。贡献方向：

- 新增布局模式（时间轴页、2x2矩阵等）
- 扩展色彩主题（深色模式、品牌定制）
- 补充示例代码与文档翻译

</details>

---

<div align="center">
<sub>Apache 2.0 · Copyright © 2026 <strong>likaku</strong> · <a href="https://github.com/likaku/Mck-ppt-design-skill">GitHub</a> · <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">Issues & Feedback</a></sub>
</div>
