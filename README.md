<div align="center">

# MCK PPT Design Skill

**AI-native PowerPoint design system — 67 layouts · Harness Engineering · BLOCK_ARC charts · QA pipeline · Python runtime**

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21+-orange.svg)](https://python-pptx.readthedocs.io)
[![GitHub stars](https://img.shields.io/github/stars/likaku/Mck-ppt-design-skill?style=social)](https://github.com/likaku/Mck-ppt-design-skill)

> **Copyright © 2024-2026 Kaku Li.** Licensed under [Apache 2.0](LICENSE). See [NOTICE](NOTICE) for details.

[English](#-quick-start) · [中文说明](#中文说明) · [Harness Guide](#-harness-engineering-mode)

</div>

---

## 🖼️ Sample Output

| Cover Page | Strategy Analysis | Data Dashboard |
|:---:|:---:|:---:|
| <img width="420" alt="Cover" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" /> | <img width="420" alt="Content" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" /> | <img width="420" alt="Table" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" /> |
| **4-Column Framework** | **Color System** | **Executive Summary** |
| <img width="420" alt="4-Column" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" /> | <img width="420" alt="Colors" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" /> | <img width="420" alt="Summary" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" /> |

---

## ⚡ Quick Start

```bash
pip install python-pptx lxml
```

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
from mck_ppt.constants import *

eng = MckEngine(total_slides=12)
eng.cover(title='Q1 2026 Strategy Review', subtitle='Board Presentation', date='2026')
eng.toc(items=[('1', 'Market Overview', 'Current landscape'), ('2', 'Strategy', 'Key actions')])
eng.table_insight(title='Three shifts driving market restructuring',
    headers=['Dimension', 'Before', 'After'],
    rows=[['Distribution', 'Offline-first', 'Digital-first'],
          ['Pricing', 'Cost-plus', 'Value-based']],
    insights=['Digital channels now control 60% of CAC', 'Value pricing unlocks 3× margins'])
eng.donut(title='Revenue Mix', segments=[(0.45, NAVY, 'Product'), (0.35, ACCENT_BLUE, 'Service'), (0.20, ACCENT_GREEN, 'Other')])
eng.timeline(title='12-month roadmap', milestones=[('Q1', 'Foundation'), ('Q2', 'Pilot'), ('Q3', 'Scale'), ('Q4', 'Review')])
eng.closing(title='Thank You')
eng.save('output/deck.pptx')
```

### AI Agent Compatibility

| AI Agent | Status |
|----------|--------|
| **WorkBuddy / Codebuddy** | ✅ Native skill (`mck-ppt-design`) |
| **Claude / Claude Code** | ✅ Load SKILL.md as skill |
| **Cursor / Continue** | ✅ Add as project rule |
| **Any LLM** | ✅ Feed SKILL.md as context |

---

## 🔱 Harness Engineering Mode

> **New in v2.3.3-harness** — Transforms the skill from vibe-coding into structured 5-stage generation with machine-readable gates.

### The Problem: Vibe Coding vs. Harness

Without Harness, the entire 3,967-line SKILL.md was loaded every time, AI jumped straight to code, and errors were only discovered *after* rendering. With Harness, context is loaded progressively, structure is locked before content, and gates run as Python scripts — not AI self-evaluation.

| Dimension | Without Harness | With Harness |
|-----------|-----------------|--------------|
| Context load | 3,967 lines every time (~6k tokens) | 245-line entry → load per-stage (~800 tokens) |
| Structure design | AI guesses layouts directly | `outline.json` locks every slide's layout + insight |
| Action titles | Topic labels ("Market Overview") | Full insight sentences ("Three shifts driving restructuring") |
| Error discovery | After rendering (expensive to fix) | S3 gate catches API format errors before render |
| Gate enforcement | AI self-declaration ("looks good") | `gate_check.py` outputs `gate_result.json` — `passed` is a Python bool |
| Experience accumulation | Discarded after fix | `experiences/*.md` → AI reads on next deck, avoids repeat mistakes |
| QA score (A/B test) | 95/100 (simple layouts, fewer QA rules triggered) | 92/100 (complex layouts, richer content) |
| Insight titles (A/B test) | **0/7** content slides | **8/12** content slides |

> **Note on QA scores**: The no-Harness deck scored higher *because* it used simpler layouts (fewer QA rules triggered). The Harness deck used richer layouts (donut, matrix_2×2, value_chain) that expose more QA checks. QA score measures *layout health*, not *content quality* — they are different dimensions.

### 5-Stage Generation Flow

```
┌─────────┐   ┌────────────────┐   ┌───────────────────┐   ┌──────────────────┐   ┌──────────┐
│ S1      │──▶│ S2 ⭐           │──▶│ S3 ⭐              │──▶│ S4 ⭐⭐           │──▶│ S5       │
│ Brief   │   │ Structure      │   │ Content           │   │ Render + QA      │   │ Deliver  │
│ brief.md│   │ outline.json   │   │ content.json      │   │ .pptx            │   │ + Learn  │
└─────────┘   └────────────────┘   └───────────────────┘   └──────────────────┘   └──────────┘
                   ⭐ = gate (FAIL → fix in current stage, never skip)
                   S3/S4 gates run Python scripts → read JSON result → no AI self-eval
```

| Stage | Input | Output | Gate |
|-------|-------|--------|------|
| S1 Brief | User request | `brief.md` (audience, goal, duration, messages) | AI self-check: 3 fields non-empty |
| S2 Structure | `brief.md` | `outline.json` (layout + insight per slide) | AI self-check: cover exists, layouts valid, titles are sentences |
| S3 Content | `outline.json` | `content.json` (copy, data, sources) | **`gate_check_s3.py`** → `gate_s3.json` |
| S4 Render+QA | `content.json` | `.pptx` + `gate_result.json` | **`gate_check.py`** → `gate_result.json` |
| S5 Deliver | `gate_result.json` (`passed: true`) | Final `.pptx` + `experiences/` update | Read `gate_result.json` before declaring done |

### Machine-Readable Gates (the key innovation)

**S3 Gate — Content audit before rendering:**

```bash
python ~/.workbuddy/skills/mck-ppt-design/references/scripts/gate_check_s3.py \
    ./ppt-project-foo/content.json  ./ppt-project-foo/
```

Checks (all automatic, no AI judgment):
- `four_column` / `executive_summary` / `meet_the_team` items are 3-tuples `(num, title, desc)`
- `matrix_2x2` quadrants are 3-tuples `(label, bg_color, desc)` — exactly 4
- `process_chevron` steps ≤ 5, labels contain no `\n`, desc ≤ 50 chars
- `donut` / `pie` segments ≤ 6
- `grouped_bar` categories ≤ 6, series ≤ 3
- Every content slide has a non-empty `source`
- Action titles are full sentences (> 10 chars)

Output `gate_s3.json`:
```json
{
  "passed": true,
  "verdict": "PASS — 可进入 S4 渲染",
  "fail_items": [],
  "pass_items": [...]
}
```

**S4 Gate — QA after rendering:**

```bash
python ~/.workbuddy/skills/mck-ppt-design/references/scripts/gate_check.py \
    ./output.pptx  ./ppt-project-foo/
```

Classifies every QA error as `user_code` (must fix) or `engine_bug` (whitelisted, with written evidence). `passed = user_code_errors == 0`. AI cannot override this.

Output `gate_result.json`:
```json
{
  "passed": true,
  "overall_score": 92,
  "checklist": {
    "user_code_errors": 0,
    "engine_bug_errors": 7,
    "warnings": 1
  },
  "verdict": "PASS — 可进入 S5 交付",
  "user_code_error_detail": []
}
```

### Anti-Patterns (from `SKILL.md` — read before every generation)

The three most common failure modes observed in real execution:

1. **"Gate passed" without running the script** — AI declares pass based on self-assessment. Fix: read `gate_result.json`. If file doesn't exist, gate hasn't run.
2. **S3 "checked in my head"** — Mental review misses API format errors every time. Fix: run `gate_check_s3.py`. Today's 3 API format bugs were all caught by script, none by self-review.
3. **"engine_bug" as escape hatch** — AI reclassifies user errors as engine bugs to pass the gate. Fix: `ENGINE_BUG_WHITELIST` in `gate_check.py` is code, not AI judgment. Add to whitelist only with written evidence.

### Knowledge Structure

```
~/.workbuddy/skills/mck-ppt-design/
├── SKILL.md                     # Entry (~245 lines): flow + hard rules + index
│                                  (was 3,967 lines before Harness)
├── references/
│   ├── INDEX.md                 # Per-stage load router
│   ├── layout-matrix.yaml       # Layout × theme × char_budget (source of truth)
│   ├── team/
│   │   ├── brand-guide.md       # Colors, fonts, design principles
│   │   └── presentation-convention.md
│   ├── framework/
│   │   ├── engine-api.md        # 67-method API quick reference
│   │   ├── guard-rails.md       # 10 production rules
│   │   └── planning-guide.md    # Layout selection, content density
│   └── scripts/
│       ├── gate_check.py        # S4 gate: QA classification → gate_result.json
│       └── gate_check_s3.py     # S3 gate: API format + count → gate_s3.json
└── experiences/
    ├── overflow.md              # Text overflow patterns (6 entries)
    ├── chart-limits.md          # Chart data limits (2 entries)
    ├── layout-pitfalls.md       # Layout-specific traps (5 entries)
    └── cjk-issues.md            # CJK rendering (2 entries)
```

---

## 🏗️ Architecture

### v1.x → v2.0: GPU → CPU Shift

> The fundamental change in v2.0: **moving ~80% of compute from GPU (LLM inference) to CPU (deterministic Python execution).**

In v1.x, every coordinate, every color, every spacing value was produced by the model token-by-token. A single donut chart required **2,800 `add_rect()` calls** — each one a GPU-computed token. A 30-slide deck burned 40,000–60,000 output tokens and ~2 minutes per chart.

v2.0 extracts that deterministic work into Python: `eng.donut()` → 20 AI tokens → 2,745 lines of CPU execution.

| | v1.x (Pure GPU) | v2.0 (GPU + CPU) |
|---|---|---|
| Compute split | ~100% GPU | ~20% GPU (decisions) + ~80% CPU (execution) |
| Chart rendering | `add_rect()` stacking (100–2,800 shapes) | `BLOCK_ARC` native arcs (3–4 shapes) |
| Output tokens / 30-slide deck | 40,000–60,000 | 9,000–12,000 |
| Chart generation time | ~2 min | <1 sec |
| File size (chart-heavy) | 2–5 MB | 0.5–1 MB |

### 5-Layer Architecture (v2.3.3-harness)

```
┌──────────────────────────────────────────────────────────┐
│  L5  Harness Engineering (v2.3.3-harness)                │
│      5-stage flow · 4 gates · 3-tier knowledge           │
│      machine-readable gates · self-refinement loop       │
├──────────────────────────────────────────────────────────┤
│  L4  Post-Processing Pipeline                            │
│      Three-layer XML corruption defense                  │
│      Full p:style / shadow / 3D sanitization             │
│      CJK font injection (KaiTi for all East Asian text)  │
├──────────────────────────────────────────────────────────┤
│  L3  Review + Auto-fix Pipeline (v2.3)                   │
│      Dual QA (layout + narrative) → AutoFixPipeline      │
│      Peer font harmonization · Gate: 0 ERROR = PASS      │
├──────────────────────────────────────────────────────────┤
│  L2  Python Runtime Engine (mck_ppt/)                    │
│      67 high-level layout methods                        │
│      Drawing primitives · XML cleanup · Constants        │
│      BLOCK_ARC chart engine · AI cover image (v2.2)      │
├──────────────────────────────────────────────────────────┤
│  L1  Design Specification                                │
│      Color system + typography + layout patterns         │
│      (Split into references/ for progressive loading)    │
└──────────────────────────────────────────────────────────┘
```

---

## 🛡️ Production Guard Rails

13 rules hard-won from production use:

| # | Rule | What It Prevents |
|---|------|-----------------|
| 1 | Never use connectors | File corruption from connector `p:style` |
| 2 | Always call `_clean_shape()` | Shadow/3D artifacts |
| 3 | Run `full_cleanup()` after save | Residual theme effects |
| 4 | Set `set_ea_font()` on CJK text | Chinese characters rendering as boxes |
| 5 | Use `add_hline()` not `add_line()` | Connector-based lines causing repair prompts |
| 6–8 | Overflow + spacing + dynamic sizing | Overlapping text, missing content |
| 9 | Mandatory BLOCK_ARC for circular charts | Rect-block bloat |
| 10 | Peer font consistency | Same-row shapes with different font sizes |
| 11 | Text-line collision check | Text overlapping separator lines |
| 12 | Post-generation QA gate | Silent defects |
| 13 | Chart legend overflow check | Legend exceeding content area |

---

## 📁 Project Structure

```
├── SKILL.md                     # Harness entry (~245 lines, was 3,967)
├── mck_ppt/                     # Python runtime engine
│   ├── engine.py                # 67 layout methods
│   ├── core.py                  # Drawing primitives + XML cleanup
│   ├── constants.py             # Colors, typography, grid
│   ├── qa.py                    # Layout QA engine
│   ├── review.py                # Review + auto-fix pipeline
│   ├── deck_builder.py          # Storyline-driven deck generator
│   ├── cover_image.py           # AI cover image (Hunyuan 2.0)
│   └── storylines/              # Pre-built storyline templates
├── references/
│   ├── INDEX.md                 # Per-stage knowledge router
│   ├── layout-matrix.yaml       # Layout × char_budget (source of truth)
│   ├── team/                    # Brand guide, presentation conventions
│   ├── framework/               # Engine API, guard rails, planning guide
│   └── scripts/
│       ├── gate_check.py        # S4 QA gate → gate_result.json
│       └── gate_check_s3.py     # S3 content gate → gate_s3.json
├── experiences/
│   ├── overflow.md              # Accumulated overflow lessons
│   ├── chart-limits.md          # Chart data limits
│   ├── layout-pitfalls.md       # Layout-specific pitfalls
│   └── cjk-issues.md            # CJK rendering issues
├── assets/icons/                # 6 pre-built PNG icons
├── examples/                    # Minimal example + staircase demo
└── CHANGELOG.md
```

---

## 📊 Version History

| Version | Date | Highlights |
|---------|------|------------|
| **v2.3.3-harness** | 2026-05-03 | **Harness Engineering**: SKILL.md 3,967→245 lines; 3-tier knowledge split; machine-readable S3/S4 gates (`gate_check.py`, `gate_check_s3.py`); `experiences/` self-refinement loop; anti-pattern warnings |
| **v2.3.3** | 2026-04-09 | Layout polish, unified color palette, retired 5 legacy layouts |
| **v2.3.2** | 2026-03-25 | DeckBuilder storyline generator; stacked_bar fix; `chart_legend_overflow` QA rule; 33-slide AI enterprise demo |
| **v2.3.1** | 2026-03-24 | Dynamic row height for `numbered_list_panel`; `text_line_collision` QA rule |
| **v2.3.0** | 2026-03-24 | Post-generation review + auto-fix pipeline; peer font harmonization; 14 errors → 0 |
| **v2.2.0** | 2026-03-23 | AI cover image pipeline (Hunyuan 2.0 + rembg) |
| **v2.0.5** | 2026-03-21 | `table_insight` (#71), icon library, unified release |
| **v2.0** | 2026-03-19 | BLOCK_ARC charts, Python runtime engine, three-tier architecture |
| v1.9 | 2026-03-12 | 9 production guard rails |
| v1.0 | 2026-03-02 | Initial release |

---

## Community

<table width="100%">
  <tr>
    <td align="center" width="50%" valign="top">
      <p><strong>WeChat Group / 微信交流群</strong></p>
      <img width="160" alt="WeChat QR" src="https://github.com/user-attachments/assets/71373a9b-7611-46e0-801f-6105446dde54" />
      <br/>
      <sub>扫描二维码加入社群</sub>
    </td>
    <td align="center" width="50%" valign="top">
      <p><strong>Discord Server</strong></p>
      <br/>
      <a href="https://discord.gg/SaFybFAT">
        <img src="https://img.shields.io/badge/Discord-Join_Community-5865F2?style=for-the-badge&logo=discord&logoColor=white" alt="Discord" />
      </a>
      <br/><br/>
      <p>Click above to join our global community</p>
    </td>
  </tr>
</table>

---

## 中文说明

<details>
<summary><b>点击展开中文文档</b></summary>

### Harness 工程化模式（v2.3.3-harness 新增）

这次更新的核心不是新增版式，而是**把 AI 做 PPT 的方式从"感觉驱动"改成"工程化驱动"**。

#### 主要变化

**SKILL.md 从 3,967 行瘦身到 245 行**。原来所有内容全量加载（~6000 tokens），现在入口只有 245 行，每个阶段按需加载对应文件（~800 tokens 起）。

**五阶段流程 + 机读化门禁**。生成一个 PPT 不再是"一句话 → 直接写代码"，而是：
1. S1 需求定义 → 写 `brief.md`
2. S2 结构设计 → 写 `outline.json`（每页确定版式和洞见句）
3. S3 内容填充 → 写 `content.json` → **运行 `gate_check_s3.py`**
4. S4 渲染 + QA → 生成 `.pptx` → **运行 `gate_check.py`**
5. S5 交付 → 读 `gate_result.json` 确认 `passed: true` 再交付

**门禁是 Python 脚本，不是 AI 口头判断**。这是最关键的改变。之前 AI 可以说"7个错误都是 engine 设计行为，门禁通过"——这是 AI 在给自己写完成证书。现在 `gate_check.py` 运行后输出 `gate_result.json`，`passed` 字段由 `user_code_errors == 0` 这个 Python 布尔值决定，AI 无法绕过。

#### A/B 对比（相同提示词，真实测量）

| 维度 | 无 Harness | 有 Harness |
|------|-----------|-----------|
| 洞见标题比例 | **0/7**（全是"主题词"标签） | **8/12**（含数字判断的完整句子）|
| 版式多样性 | ~5 种 | 9 种（donut/matrix_2x2/value_chain 等）|
| 章节结构 | 10 页平铺 | 15 页三章节+分隔页 |
| QA 分数 | 95（简单版式，触发规则少）| 92（复杂版式，更多内容） |

> QA 分数高不代表质量好——无 Harness 版用了更简单的版式，触发更少 QA 规则，所以分更高。Harness 版用了更丰富的版式（donut、四象限矩阵、价值链），洞见密度高得多。

#### 经验沉淀（Self-Refinement）

每次修正"模式性问题"后，AI 会写入 `experiences/` 目录对应文件。下次做 PPT，AI 在 S3 阶段会读这些文件，主动避免同类错误：

- `experiences/overflow.md` — 文字溢出经验（6 条）
- `experiences/chart-limits.md` — 图表数据段数限制（2 条）
- `experiences/layout-pitfalls.md` — 版式特定踩坑（5 条）
- `experiences/cjk-issues.md` — 中文渲染问题（2 条）

### v2.0 核心架构变化

v2.0 的本质是**将约 80% 的算力从 GPU（大模型推理）迁移到 CPU（确定性 Python 执行）**。

一个甜甜圈图从 2,800 个 `add_rect()` 调用变成了 1 行 `eng.donut()`，从 ~2 分钟 GPU 推理变成 <1 秒 CPU 执行，文件从 5MB 缩小到 0.5MB。

```python
# v1.x：AI 输出每个方块的坐标（2800+ tokens）
for angle in range(0, 360):
    add_rect(slide, x + cos(angle)*r, y + sin(angle)*r, ...)

# v2.0：AI 输出 20 个 token → CPU 执行 2745 行确定性代码
eng.donut(title='市场份额', segments=[(0.35, NAVY, '我们'), ...])
```

### 快速上手

```bash
pip install python-pptx lxml
# 安装到 WorkBuddy / Codebuddy：在 skill 市场搜索 mck-ppt-design
# 或手动：cp -r /path/to/mck-ppt-design ~/.workbuddy/skills/
```

</details>

---

<div align="center">
<sub>Apache 2.0 · © 2024-2026 <strong>Kaku Li</strong> · <a href="https://github.com/likaku/Mck-ppt-design-skill">GitHub</a> · <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">Issues</a></sub>
</div>
