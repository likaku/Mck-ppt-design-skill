<div align="center">

# McKinsey PPT Design Skill

**AI-native PowerPoint design system — 70 layouts · BLOCK_ARC chart engine · Python runtime**

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21+-orange.svg)](https://python-pptx.readthedocs.io)
[![GitHub stars](https://img.shields.io/github/stars/likaku/Mck-ppt-design-skill?style=social)](https://github.com/likaku/Mck-ppt-design-skill)

[English](#-v1x--v20--whats-changed-and-why) · [中文说明](#中文说明)

</div>

---

## Community

<table>
<tr>
    <td align="center" width="50%" valign="top">
      <strong>WeChat Group / 微信交流群</strong><br/><br/>
<img width="180" alt="Clipboard_Screenshot_1773906668" src="https://github.com/user-attachments/assets/b9d24976-91fc-4d1d-91b6-b14b7e910da2" />
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

## ⚡ v1.x → v2.0 — What's Changed and Why

> v2.0 is a direct response to community feedback: **v1.x consumes too many tokens on chart-heavy decks.**
>
> We rewrote the chart rendering pipeline and added a Python runtime engine to address this. v2.0 is still being battle-tested — **if you have production workloads depending on v1.x, take your time upgrading.** Report issues in the WeChat group or Discord above.

### The Core Problem in v1.x

v1.x rendered donut / pie / gauge charts by stacking hundreds of tiny `add_rect()` blocks — each block = 1 XML shape = ~500 bytes of output tokens. A single complex donut could generate **2,800 shapes**, a **5 MB file**, and cost **~2 minutes** of generation time. The AI had to output the coordinates for every single block.

### How v2.0 Fixes It

v2.0 replaces the rect-block approach with PowerPoint's native `BLOCK_ARC` shape. Each chart segment is now **1 shape** instead of hundreds:

```python
# v1.x: 200+ shapes per chart, AI outputs every coordinate
for angle in range(0, 360):
    add_rect(slide, x + cos(angle)*r, y + sin(angle)*r, ...)

# v2.0: 3-4 shapes per chart, AI calls one function
add_block_arc(slide, cx, cy, outer_r, start_deg=0, sweep_deg=120, fill_color=NAVY)
add_block_arc(slide, cx, cy, outer_r, start_deg=120, sweep_deg=80, fill_color=ACCENT_BLUE)
add_block_arc(slide, cx, cy, outer_r, start_deg=200, sweep_deg=160, fill_color=BG_GRAY)
```

### v1.x vs v2.0 — Technical Comparison

| | v1.x | v2.0 |
|---|------|------|
| **Chart rendering** | `add_rect()` block stacking (100–2,800 shapes/chart) | `BLOCK_ARC` native arcs (3–4 shapes/chart) |
| **Code generation** | AI writes raw `add_shape()` / coordinate math per slide | AI calls `eng.donut_chart()`, `eng.cover()` etc. — 70 high-level methods |
| **Rounds per 30-slide deck** | 10–15 (trial-and-error) | 3–4 (deterministic) |
| **Output tokens per deck** | 40,000–60,000 | 9,000–12,000 |
| **Chart generation time** | ~2 min | <1 sec |
| **File size (chart-heavy)** | 2–5 MB | 0.5–1 MB |
| **File corruption defense** | Basic XML cleanup | Three-layer defense (p:style, shadow, 3D sanitization) |
| **CJK handling** | Manual font setting | Automatic `set_ea_font()` on all CJK text runs |
| **Architecture** | Single-tier (SKILL.md only) | Three-tier (SKILL.md + Python engine + post-processing) |

### v2.0 Three-Tier Architecture

```
┌─────────────────────────────────────────────────────────┐
│  Tier 1: SKILL.md (Design Specification)                │
│  ├── 70 layout patterns with exact coordinates          │
│  ├── Color system + typography hierarchy                 │
│  ├── 9 production guard rails                           │
│  └── BLOCK_ARC chart rendering spec                     │
├─────────────────────────────────────────────────────────┤
│  Tier 2: mck_ppt/ (Python Runtime Engine)      [NEW]   │
│  ├── engine.py — 70 high-level layout methods           │
│  ├── core.py — Drawing primitives + XML cleanup         │
│  ├── constants.py — Colors, typography, grid constants  │
│  └── __init__.py — Clean public API                     │
├─────────────────────────────────────────────────────────┤
│  Tier 3: Post-Processing Pipeline              [NEW]    │
│  ├── Three-layer file corruption defense                │
│  ├── Full XML sanitization (p:style, shadow, 3D)        │
│  └── CJK font injection (KaiTi for East Asian text)     │
└─────────────────────────────────────────────────────────┘
```

**Tier 1** tells the AI *what* to build — every layout has pixel-perfect coordinates, every color has a hex code, every edge case has a documented solution.

**Tier 2** [NEW] is a complete Python library. Instead of the AI writing `add_shape()` from scratch, it calls `eng.cover()`, `eng.toc()`, `eng.donut_chart()` — one method per layout pattern, 2,745 lines of production-tested code.

**Tier 3** [NEW] automatically prevents the #1 cause of "file needs repair" errors in AI-generated PPTs.

---

## 🛡️ Production Guard Rails

9 rules hard-won from 50+ production generations:

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
| 9 | Mandatory BLOCK_ARC for circular charts | Rect-block bloat (v1.x legacy problem) |

---

## 🚀 Quick Start

```bash
pip install python-pptx lxml

# Option 1: ClawHub (for AI agents)
npx clawhub@latest install mck-ppt-design

# Option 2: Manual
mkdir -p ~/.claude/skills/mck-ppt-design
cp SKILL.md ~/.claude/skills/mck-ppt-design/

# Option 3: Python engine directly
pip install -e .  # or copy mck_ppt/ to your project
```

```python
from mck_ppt import MckEngine

eng = MckEngine(total_slides=30)
eng.cover(title='Q1 2026 Strategy Review', subtitle='Board Presentation')
eng.toc(items=[('1', 'Market Overview', 'Current landscape'), ...])
eng.donut_chart(title='Revenue Mix', segments=[('Product A', 45, NAVY), ...])
eng.save('output/deck.pptx')
```

### Compatibility

| AI Agent | Status | Install Method |
|----------|--------|----------------|
| **Claude** (Anthropic) | ✅ Supported | ClawHub or manual SKILL.md |
| **Cursor** | ✅ Supported | Add as project rule |
| **Codebuddy** | ✅ Supported | Load as skill |
| **GPT / ChatGPT** | ✅ Works | Paste SKILL.md as system prompt |
| **Any LLM** | ✅ Universal | Feed SKILL.md as context |

### Sample Output

| Cover Page | Content Page | Table Page |
|:------:|:------:|:------:|
| <img width="600" alt="Cover" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" /> | <img width="600" alt="Content" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" /> | <img width="600" alt="Table" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" /> |
| **4-Column Layout** | **Color System** | **Summary Page** |
| <img width="600" alt="4-Column" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" /> | <img width="600" alt="Colors" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" /> | <img width="600" alt="Summary" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" /> |

---

## 📁 Project Structure

```
├── SKILL.md                 # Design specification (290KB, 6100 lines)
├── mck_ppt/                 # Python runtime engine (140KB)     [NEW v2.0]
│   ├── __init__.py          # Public API
│   ├── engine.py            # 70 layout methods (2,359 lines)
│   ├── core.py              # Drawing primitives + XML cleanup (295 lines)
│   └── constants.py         # Colors, typography, grid (78 lines)
├── CHANGELOG.md
├── scripts/
│   ├── minimal_example.py
│   └── requirements.txt
└── references/
    ├── color-palette.md
    └── layout-catalog.md
```

---

## 📊 Version History

| Version | Date | Highlights |
|---------|------|------------|
| **v2.0** | 2026-03-19 | BLOCK_ARC chart engine, Python runtime engine, three-tier architecture |
| v1.10.x | 2026-03-15 | Channel delivery, dynamic sizing |
| v1.9 | 2026-03-12 | Production guard rails (9 rules) |
| v1.8 | 2026-03-10 | Layout expansion: 39 → 70 patterns |
| v1.7 | 2026-03-08 | Data charts: grouped/stacked/horizontal bar |
| v1.0 | 2026-03-02 | Initial release |

---

## 中文说明

<details>
<summary><b>点击展开中文文档</b></summary>

### v2.0 更新说明

v2.0 是对社区反馈的直接回应：**v1.x 在图表密集型 PPT 上 token 消耗过高。**

我们重写了图表渲染管线，新增了 Python 运行时引擎。v2.0 仍在持续验证中——**如果你的生产工作依赖 v1.x，建议稳妥升级，不必急。** 遇到问题请在微信群或 Discord 反馈。

### v1.x 的问题

v1.x 用几百个 `add_rect()` 小方块堆叠渲染甜甜圈/饼图/仪表盘。一个复杂甜甜圈可能产生 **2,800 个形状**、**5 MB 文件**、**~2 分钟**生成时间。AI 需要逐个输出每个方块的坐标。

### v2.0 怎么解决的

用 PowerPoint 原生 `BLOCK_ARC` 形状替代方块堆叠。每个图表段现在是 **1 个形状**而非几百个。

### 技术对比

| | v1.x | v2.0 |
|---|------|------|
| 图表渲染 | `add_rect()` 堆叠（100–2,800 形状/图表） | `BLOCK_ARC` 原生弧（3–4 形状/图表） |
| 代码生成 | AI 手写 `add_shape()` + 坐标计算 | AI 调用 `eng.donut_chart()` 等 70 个高级方法 |
| 30 页 PPT 交互轮数 | 10–15 轮 | 3–4 轮 |
| 输出 tokens | 40,000–60,000 | 9,000–12,000 |
| 图表生成时间 | ~2 分钟 | <1 秒 |
| 文件大小（图表密集） | 2–5 MB | 0.5–1 MB |
| 架构 | 单层（仅 SKILL.md） | 三层（SKILL.md + Python 引擎 + 后处理） |

### 快速上手

```bash
pip install python-pptx lxml
npx clawhub@latest install mck-ppt-design
```

</details>

---

<div align="center">
<sub>Apache 2.0 · © 2026 <strong>likaku</strong> · <a href="https://github.com/likaku/Mck-ppt-design-skill">GitHub</a> · <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">Issues</a></sub>
</div>
