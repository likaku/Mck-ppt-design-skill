# Architecture & Guard Rails

> Back to [README](../README.md)


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
