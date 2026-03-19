<div align="center">

# McKinsey PPT Design Skill

**The most advanced AI-native PowerPoint design system — built for performance**
<br/>70 layouts · BLOCK_ARC chart engine · Python runtime · Save $2-4 per deck on Opus

[English](#-performance-architecture) · [中文说明](#中文说明)

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21+-orange.svg)](https://python-pptx.readthedocs.io)
[![GitHub stars](https://img.shields.io/github/stars/likaku/Mck-ppt-design-skill?style=social)](https://github.com/likaku/Mck-ppt-design-skill)

</div>

---

## ⚡ Performance Architecture

> **Core thesis**: Every token your AI spends figuring out *how* to build a PPT is a token not spent on *what* to put in it.

This skill takes a fundamentally different approach from typical AI PPT generation. Instead of letting the AI improvise slide-by-slide (10-15 rounds of trial-and-error), we provide a **complete, pre-optimized design system** that reduces generation to **3-4 deterministic rounds**.

### Token Economics — Real Numbers

| Metric | Without This Skill | With This Skill | Savings |
|--------|-------------------|-----------------|---------|
| **Rounds per 30-slide deck** | 10–15 | 3–4 | **~70% fewer rounds** |
| **Output tokens per deck** | 40,000–60,000 | 9,000–12,000 | **~75% reduction** |
| **Claude Opus cost per deck** | $3.00–4.50 | $0.67–0.90 | **Save $2–4 per deck** |
| **Claude Sonnet cost per deck** | $0.60–0.90 | $0.14–0.18 | **Save ~$0.50 per deck** |
| **GPT-4o cost per deck** | $0.60–0.75 | $0.14–0.15 | **Save ~$0.50 per deck** |
| **Chart rendering (donut/pie)** | 100–2,800 shapes | 3–4 shapes | **99% fewer shapes** |
| **Chart generation time** | ~2 min | <1 sec | **120x faster** |
| **File size (chart-heavy)** | 2–5 MB | 0.5–1 MB | **60–80% smaller** |

> 💡 **For heavy Opus users**: If you generate 10 decks/month, this skill saves you **$20–40/month** in API costs alone. That's before counting the hours saved on manual fixes.

### Why It's Faster

```
Traditional AI PPT Generation:
  Round 1: "Create a cover slide" → AI guesses layout, colors, fonts
  Round 2: "The font is wrong, fix it" → AI patches
  Round 3: "Add a table slide" → AI improvises table design
  Round 4: "The alignment is off" → AI tries to fix
  Round 5–15: Repeat for every slide type...
  → 40,000-60,000 output tokens, inconsistent results

With McKinsey PPT Design Skill:
  Round 1: AI reads SKILL.md → knows ALL 70 layouts, colors, fonts, guard rails
  Round 2: AI generates complete deck using pre-defined patterns
  Round 3: Minor adjustments (if any)
  → 9,000-12,000 output tokens, consistent professional results
```

---

## 🏗️ Architecture

### Three-Tier Design

```
┌─────────────────────────────────────────────────────────┐
│  Tier 1: SKILL.md (Design Specification)                │
│  ├── 70 layout patterns with exact coordinates          │
│  ├── Color system + typography hierarchy                 │
│  ├── Production guard rails (9 mandatory rules)         │
│  ├── Common issues + solutions (20 documented)          │
│  └── BLOCK_ARC chart rendering spec                     │
├─────────────────────────────────────────────────────────┤
│  Tier 2: mck_ppt/ (Python Runtime Engine)      [NEW]   │
│  ├── engine.py — 70 high-level layout methods           │
│  ├── core.py — Drawing primitives + XML cleanup         │
│  ├── constants.py — Colors, typography, grid constants  │
│  └── __init__.py — Clean public API                     │
├─────────────────────────────────────────────────────────┤
│  Tier 3: Post-Processing Pipeline                       │
│  ├── Three-layer file corruption defense                │
│  ├── Full XML sanitization (p:style, shadow, 3D)        │
│  └── CJK font injection (KaiTi for East Asian text)     │
└─────────────────────────────────────────────────────────┘
```

**Tier 1** (SKILL.md) is the brain — it tells the AI agent *what* to build and *how*. Every layout has pixel-perfect coordinates, every color has a hex code, every edge case has a documented solution.

**Tier 2** (mck_ppt/) is the muscle — a complete Python library with 70+ high-level layout methods. Instead of the AI writing `add_shape()` calls from scratch, it calls `eng.cover()`, `eng.toc()`, `eng.three_column()`, `eng.donut_chart()` and gets production-ready slides.

**Tier 3** is the immune system — automatic cleanup that prevents the #1 cause of "file needs repair" errors in AI-generated PPTs.

### Python Runtime Engine (NEW in v2.0)

The `mck_ppt/` package provides a complete, importable Python engine:

```python
from mck_ppt import MckEngine

eng = MckEngine(total_slides=30)
eng.cover(title='Q1 2026 Strategy Review', subtitle='Board Presentation', author='Strategy Team')
eng.toc(items=[('1', 'Market Overview', 'Current landscape'), ...])
eng.three_column(title='Key Findings', items=[...])
eng.donut_chart(title='Revenue Mix', segments=[('Product A', 45, NAVY), ...])
eng.save('output/strategy_deck.pptx')
```

**70 layout methods** — one method per layout pattern. No boilerplate, no coordinate math, no XML wrestling.

---

## 🎯 What It Does

- 🎨 **70 layout patterns** across 12 categories — structure, data, framework, comparison, narrative, timeline, team, charts, images, advanced viz, dashboards, visual storytelling
- 📊 **BLOCK_ARC native chart engine** — donut, pie, gauge rendered with 3-4 native shapes instead of hundreds of rect blocks (v2.0)
- 📐 **Strict McKinsey design system** — flat design, no shadows, no 3D, consistent typography
- 🛡️ **Three-layer file corruption defense** — eliminates `p:style`, shadow, and 3D artifacts
- 🔤 **CJK + Latin font handling** — KaiTi / Georgia / Arial
- 🖼️ **Image placeholder system** — gray boxes with crosshairs for easy replacement
- 🚀 **9 production guard rails** — preventing common AI generation mistakes
- 📨 **Channel delivery** — auto-send PPTX via Feishu, Telegram, Slack, Discord, WhatsApp

### Sample Output

| Cover Page | Content Page | Table Page |
|:------:|:------:|:------:|
| <img width="600" alt="Cover" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" /> | <img width="600" alt="Content" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" /> | <img width="600" alt="Table" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" /> |
| **4-Column Layout** | **Color System** | **Summary Page** |
| <img width="600" alt="4-Column" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" /> | <img width="600" alt="Colors" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" /> | <img width="600" alt="Summary" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" /> |

---

## 🚀 Quick Start

```bash
# Install dependencies
pip install python-pptx lxml

# Option 1: Use the Python engine directly
pip install -e .  # or just copy mck_ppt/ to your project

# Option 2: Install from ClawHub (for AI agents)
npx clawhub@latest install mck-ppt-design

# Option 3: Manual install for Claude
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

## 📐 Layout Categories (70 Patterns)

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

## 🔧 BLOCK_ARC Chart Engine (v2.0)

The biggest performance win in v2.0. Previously, donut/pie/gauge charts were drawn with hundreds of tiny `add_rect()` blocks — each block = 1 XML shape = ~500 bytes. A complex donut chart could generate **2,800 shapes** and a **5 MB file**.

Now we use PowerPoint's native `BLOCK_ARC` shape with precise angle calculations:

```python
# Before (v1.x): 200+ shapes, 2 minutes, 3 MB
for angle in range(0, 360):
    add_rect(slide, x + cos(angle)*r, y + sin(angle)*r, ...)

# After (v2.0): 4 shapes, <1 second, 0.5 MB
add_block_arc(slide, cx, cy, outer_r, start_deg=0, sweep_deg=120, fill_color=NAVY)
add_block_arc(slide, cx, cy, outer_r, start_deg=120, sweep_deg=80, fill_color=ACCENT_BLUE)
add_block_arc(slide, cx, cy, outer_r, start_deg=200, sweep_deg=160, fill_color=BG_GRAY)
```

**Impact**: A 67-slide Tencent Annual Report that previously took 15 minutes to generate now completes in under 3 minutes.

---

## 🛡️ Production Guard Rails (9 Rules)

Hard-won from 50+ production PPT generations:

| # | Rule | What It Prevents |
|---|------|------------------|
| 1 | Never use connectors | File corruption from connector `p:style` |
| 2 | Always call `_clean_shape()` | Shadow/3D artifacts leaking into slides |
| 3 | Run `full_cleanup()` after save | Residual theme effects |
| 4 | Set `set_ea_font()` on CJK text | Chinese characters rendering as boxes |
| 5 | Use `add_hline()` not `add_line()` | Connector-based lines causing repair prompts |
| 6 | Validate spacing before save | Overlapping text boxes |
| 7 | Check overflow on long content | Text truncation in fixed-height boxes |
| 8 | Dynamic sizing for variable-count | 3 items vs 7 items need different spacing |
| 9 | **Mandatory BLOCK_ARC for circular charts** | Hundreds of rect-blocks bloating file size |

---

## 📁 Project Structure

```
├── SKILL.md                 # Core design specification (290KB, 6100 lines)
│                            # ↑ The complete brain — every layout, color, rule
├── mck_ppt/                 # Python runtime engine (140KB)     [NEW v2.0]
│   ├── __init__.py          # Public API: from mck_ppt import MckEngine
│   ├── engine.py            # 70 high-level layout methods (2,359 lines)
│   ├── core.py              # Drawing primitives + XML cleanup (295 lines)
│   └── constants.py         # Colors, typography, grid constants (78 lines)
├── LICENSE                  # Apache 2.0
├── CHANGELOG.md             # Version history (v1.0 → v2.0)
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

## 📊 Version History

| Version | Date | Highlights |
|---------|------|------------|
| **v2.0** | 2026-03-19 | **BLOCK_ARC chart engine** — 99% fewer shapes, 60-80% smaller files, 120x faster chart rendering |
| v1.10.x | 2026-03-15 | Channel delivery, dynamic sizing, title spacing optimization, YAML fix |
| v1.9 | 2026-03-12 | **Production guard rails** — 7 mandatory rules from 50+ real-world generations |
| v1.8 | 2026-03-10 | **Layout expansion** — 39 → 70 patterns, 4 new categories |
| v1.7 | 2026-03-08 | **Data charts** — grouped bar, stacked bar, horizontal bar |
| v1.5 | 2026-03-08 | CJK line spacing fix |
| v1.4 | 2026-03-06 | **P0 optimization** — merged functions, 109 lines removed |
| v1.1 | 2026-03-03 | **Three-layer defense** — eliminated file corruption |
| v1.0 | 2026-03-02 | Initial release |

See [CHANGELOG.md](CHANGELOG.md) for full details.

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

- New layout patterns
- Extended color themes (dark mode, brand customization)
- Additional chart types
- Examples and documentation translations

---

## 中文说明

<details>
<summary><b>点击展开中文文档</b></summary>

### 简介

**McKinsey PPT Design Skill** 是目前最先进的 AI 原生 PPT 设计系统。将完整的麦肯锡设计规范 + Python 运行时引擎编码为一个仓库，AI 加载后即可持续输出风格统一的专业 PPT。

### 🔥 性能优化 — 核心卖点

**每次生成 PPT 帮你省 $2-4（Opus）**

| 指标 | 无此技能 | 有此技能 | 节省 |
|------|---------|---------|------|
| 30页 PPT 交互轮数 | 10-15 轮 | 3-4 轮 | **减少 ~70%** |
| 输出 tokens | 40,000-60,000 | 9,000-12,000 | **减少 ~75%** |
| Opus 费用/次 | $3.00-4.50 | $0.67-0.90 | **省 $2-4** |
| 图表渲染(甜甜圈/饼图) | 100-2,800 个形状 | 3-4 个形状 | **减少 99%** |
| 图表生成时间 | ~2 分钟 | <1 秒 | **快 120 倍** |

> 💡 如果你每月生成 10 个 PPT，仅 API 费用就节省 **$20-40/月**。

### 架构

三层架构设计：

1. **SKILL.md**（设计规范层）— 70 种布局模式 + 色彩体系 + 9 条生产规则
2. **mck_ppt/**（Python 运行时引擎）— 70+ 高级布局方法，一行代码一页幻灯片
3. **后处理流水线** — 三层文件防腐、XML 清洗、中文字体注入

### 快速上手

```bash
# 安装依赖
pip install python-pptx lxml

# 从 ClawHub 安装（推荐）
npx clawhub@latest install mck-ppt-design

# 或手动安装
mkdir -p ~/.claude/skills/mck-ppt-design
cp SKILL.md ~/.claude/skills/mck-ppt-design/
```

### 参与贡献

欢迎提交 Issue 和 Pull Request。

</details>

---

<div align="center">
<sub>Apache 2.0 · Copyright © 2026 <strong>likaku</strong> · <a href="https://github.com/likaku/Mck-ppt-design-skill">GitHub</a> · <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">Issues & Feedback</a></sub>
</div>
