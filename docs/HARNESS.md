# Harness Engineering Mode

> Back to [README](../README.md)


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
