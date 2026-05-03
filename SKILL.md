---
name: mck-ppt-design
description: >-
  Create professional, consultant-grade PowerPoint presentations from scratch
  using MckEngine (python-pptx wrapper) with McKinsey-style design. Use when
  user asks to create slides, pitch decks, business presentations, strategy
  decks, quarterly reviews, board meeting slides, or any professional PPTX.
  AI calls eng.cover(), eng.donut(), eng.timeline() etc — 67 high-level methods
  across 12 categories (structure, data, framework, comparison, narrative,
  timeline, team, charts, images, advanced viz, dashboards, visual storytelling),
  consistent typography, zero file-corruption issues, BLOCK_ARC native shapes
  for circular charts (donut, pie, gauge), production-hardened guard rails
  for spacing, overflow, legend consistency, title style uniformity,
  dynamic sizing for variable-count layouts, horizontal item overflow
  protection, chart rendering, and
  AI-generated cover images via Tencent Hunyuan 2.0 with professional cutout,
  cool grey-blue tint, and McKinsey-style Bézier ribbon decoration.
---

# McKinsey PPT Design — Harness Framework

> **Version**: 2.3.3-harness-v2 · **Engine**: MckEngine (python-pptx) · **Author**: [likaku](https://github.com/likaku/Mck-ppt-design-skill)
>
> **Required tools**: Read, Write, Bash · **Requires**: `pip install python-pptx lxml`

---

## 常见崩坏模式（每次生成前必读，先于 HARD RULES）

> 告诉 AI 要做什么效果一般，告诉它前人最容易在哪里塌效果更好。
> 以下三种反模式均已在真实执行中被验证，每次必须主动警惕。

### 反模式 1：口头宣布"门禁通过"（最常见）

**错误做法**：
> 「S4 QA 共 7 个 errors，判断均为 engine 设计行为，门禁通过，进入 S5」

**问题所在**：`passed` 是由 AI 口头判断的，不是由程序派生的。无论理由多充分，这句话都是 AI 在给自己写完成证书。

**正确做法**：
1. 执行 `python references/scripts/gate_check.py <pptx路径> <项目目录>`
2. 读取 `<项目目录>/gate_result.json`
3. 只有 `gate_result.json` 中 `"passed": true` 时，才能进入 S5
4. 如果 `"passed": false`，修复 `user_code_errors` 列表中的问题，重新渲染，再次执行 gate_check

---

### 反模式 2：S3 门禁"脑子里过一遍"就算通过

**错误做法**：
> 「S3 内容门禁预检：API 格式正确，字数在限制内，通过，进入 S4」（没有执行任何代码）

**问题所在**：今天真实发生的 3 个 API 格式错误（`four_column`/`matrix_2x2`/`executive_summary` 参数格式），靠脑子过是过不出来的，必须靠代码检查。

**正确做法**：
1. 执行 `python references/scripts/gate_check_s3.py <content.json路径> <项目目录>`
2. 读取 `<项目目录>/gate_s3.json`
3. 只有 `"passed": true` 时，才能进入 S4
4. 如果有 `fail` 项，修正 `content.json`，重新执行 gate_check_s3

---

### 反模式 3：engine_bug 分类作为软话逃生口

**错误做法**：
> 「peer_font_inconsistency 是 engine 内部设计行为，不是用户代码问题，可以放行」

**问题所在**：这个分类本身是正确的，但由 AI 在对话里口头做出，等于把豁免权交给了 AI 自己——而 AI 有动机让自己通过。

**正确做法**：
`gate_check.py` 里有硬编码的 `ENGINE_BUG_WHITELIST` 枚举。只有在白名单里的 error category，才会被豁免。白名单是代码，不是 AI 的判断。如果你认为某类 error 应该加入白名单，修改 `gate_check.py` 里的枚举，而不是口头声明豁免。

---

## HARD RULES（必须遵守，不可绕过）

1. **每次生成必须走五阶段流程**，不允许"一句话直接生成"跳过前置阶段
2. **TaskCreate 驱动进度**：开始前创建5个 task，每阶段完成后立即 TaskUpdate completed
3. **按需加载上下文**：每个阶段只读对应文件，不要全量加载旧版 SKILL.md 里的内容
4. **门禁必须机读化**：S3 和 S4 门禁必须执行对应的 gate_check 脚本，读 JSON 结果，不得口头判断
5. **Self-Refinement 必做**：每次修正 pattern-level 问题后，写入 `experiences/` 对应文件
6. **引擎路径固定**：`sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))`

---

## 知识路由（上下文加载索引）

> **在每个阶段开始时，读对应文件**。不要一次性全读。

| 阶段 | 必须读 | 说明 |
|------|--------|------|
| S1 需求 | `references/team/brand-guide.md` | 了解设计约束 |
| S2 结构 | `references/framework/engine-api.md` + `references/layout-matrix.yaml` | 版式选择和能力边界 |
| S3 内容 | `references/framework/guard-rails.md` + `experiences/*.md`（存在时）| 防溢出规则 + 历史踩坑 |
| S4 渲染 | 用到的 `references/layouts/*.md`（只读实际用到的版式）| 版式实现细节 |
| S5 交付 | 无 | — |

**完整路由表**：`references/INDEX.md`

---

## 五阶段生成流程

```
┌─────────┐   ┌─────────────┐   ┌─────────────┐   ┌──────────────┐   ┌──────────┐
│ S1 需求  │──▶│ S2 结构 ⭐   │──▶│ S3 内容 ⭐   │──▶│ S4 渲染+QA ⭐⭐│──▶│ S5 交付   │
│ brief.md│   │ outline.json│   │content.json │   │  .pptx       │   │ + 沉淀   │
└─────────┘   └─────────────┘   └─────────────┘   └──────────────┘   └──────────┘
                    ⭐ = 门禁（FAIL 则在本阶段修正，不得跳过）
    S3/S4 门禁必须运行 gate_check 脚本，读 JSON 结果 — 不得口头宣布通过
```

### Stage 1: 需求定义

**读文件**：`references/team/brand-guide.md`

**收集信息**：
- 受众（职位/决策角色）
- 目标（决策/汇报/说服）
- 时长（分钟数 → 约 1 分钟/页）
- 关键信息（最多 5 条核心 message）
- 数据来源（有哪些数据可用）

**产出**：在工作目录创建 `ppt-project-{slug}/brief.md`

**门禁**：`audience` + `goal` + `key_messages` 三项非空（AI 自评即可）

---

### Stage 2: 结构设计

**读文件**：`references/framework/engine-api.md`, `references/layout-matrix.yaml`

**任务**：
1. 根据时长确定页数（1分钟/页）
2. 为每张幻灯片选定 `layout`（查 engine-api.md 速查表）
3. 每页写一句核心 `key_point`（完整句子，不是标签）
4. 确认版式在能力边界内（查 layout-matrix.yaml）

**产出**：`ppt-project-{slug}/outline.json`

```json
{
  "brief": {"audience": "董事会", "goal": "战略汇报", "duration_minutes": 15},
  "slides": [
    {"idx": 1, "layout": "cover", "title": "Q1 2026 战略回顾", "key_point": ""},
    {"idx": 2, "layout": "toc", "title": "目录", "key_point": ""},
    {"idx": 3, "layout": "table_insight", "title": "市场格局三大转变驱动战略重构", "key_point": ""}
  ]
}
```

**⭐ 门禁 S2**（AI 自评）：
- [ ] `cover` 幻灯片存在
- [ ] 幻灯片数量在时长约束内（`count <= duration_minutes * 1.2`）
- [ ] 所有 layout 在 `layout-matrix.yaml` 中有定义
- [ ] Action Title 均为完整句子（len > 10，包含动词）
- [ ] `two_column_text` 数量 ≤ 1

---

### Stage 3: 内容填充

**读文件**：`references/framework/guard-rails.md`, `experiences/*.md`（全部存在的文件）

**任务**：
1. 为每张幻灯片填充具体文案、数字、图表数据
2. 确保每页有 `source` 出处
3. 按 `layout-matrix.yaml` 的 `char_budget` 控制文字量

**产出**：`ppt-project-{slug}/content.json`

**⭐ 门禁 S3**（必须机读化，不得 AI 自评）：

```bash
python ~/.workbuddy/skills/mck-ppt-design/references/scripts/gate_check_s3.py \
    <项目目录>/content.json  <项目目录>
```

读取 `<项目目录>/gate_s3.json`：
- `"passed": true` → 进入 S4
- `"passed": false` → 修正 `content.json` 中 `fail_items` 列出的问题，重新执行

---

### Stage 4: 渲染 + QA

**读文件**：用到的版式对应的 `references/layouts/*.md`

**任务**：
1. 根据 `content.json` 生成 Python 渲染脚本
2. 执行脚本生成 `.pptx`
3. 运行 QA 门禁脚本

**⭐⭐ 门禁 S4**（必须机读化，不得口头宣布通过）：

```bash
python ~/.workbuddy/skills/mck-ppt-design/references/scripts/gate_check.py \
    <pptx路径>  <项目目录>
```

读取 `<项目目录>/gate_result.json`：
- `"passed": true` → 进入 S5
- `"passed": false` → 查看 `user_code_errors`，修复渲染代码，重新渲染，再次执行

**注意**：`engine_bug` 类 errors（`peer_font_inconsistency` 等白名单条目）由脚本自动豁免，不需要 AI 口头判断。

**产出**：`<项目目录>/gate_result.json` + `.pptx`

---

### Stage 5: 交付 + Self-Refinement

**任务**：
1. 读取 `gate_result.json` 确认 `passed: true`（不得在无此文件时声称通过）
2. 交付 `.pptx` 文件
3. **Self-Refinement**：判断本次是否有 pattern-level 修正

**⭐ Self-Refinement 协议**：

```
修正后判断：
  ONE-TIME（单次特定调整）→ 不需要沉淀
  PATTERN（跨 deck 可重现的问题）→ 必须写入 experiences/
    - 溢出类 → experiences/overflow.md
    - 图表限制 → experiences/chart-limits.md
    - 版式踩坑 → experiences/layout-pitfalls.md
    - 中文渲染 → experiences/cjk-issues.md

格式：
  ## Experience NNN: {title}
  **Date**: YYYY-MM-DD
  **Problem**: ...
  **Root Cause**: ...
  **Fix**: ...
  **Rule**: ...（门禁层面如何预防）
```

---

## Fast Track（简单需求跳过 S2/S3 门禁）

满足**全部**条件时，AI 可跳过 S2 结构评审和 S3 内容审查：
- 总页数 ≤ 5
- 无数据图表（donut/bar/table 类）
- 用户明确表达"快速"/"马上要"/"简单"

**仍然必须**：S1 需求收集 + S4 QA 门禁（gate_check.py 脚本执行）+ S5 交付

---

## Checkpoint / 断点恢复

当用户说"继续做那个 PPT"时：

```python
import os, json, glob

projects = glob.glob('ppt-project-*/')
for proj in projects:
    brief   = os.path.exists(f'{proj}brief.md')
    outline = os.path.exists(f'{proj}outline.json')
    content = os.path.exists(f'{proj}content.json')
    gate_s3 = os.path.exists(f'{proj}gate_s3.json')
    gate_s4 = os.path.exists(f'{proj}gate_result.json')
    pptx    = bool(glob.glob(f'{proj}*.pptx'))

    if not brief:           stage = 1
    elif not outline:       stage = 2
    elif not content:       stage = 3
    elif not gate_s3:       stage = '3-gate'   # content 有了但 gate 还没跑
    elif not pptx:          stage = 4
    elif not gate_s4:       stage = '4-gate'   # pptx 有了但 gate 还没跑
    else:                   stage = 5

    print(f"项目 {proj}: 当前处于 Stage {stage}")
```

---

## MckEngine 快速模板

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
from mck_ppt.constants import *

eng = MckEngine(total_slides=12)
eng.cover(title='标题', subtitle='副标题', author='作者', date='2026年')
eng.toc(items=[('1', '章节一', '一句话描述'), ('2', '章节二', '一句话描述')])
# ... 内容页 ...
eng.closing(title='谢谢', message='期待进一步交流')
eng.save('output/deck.pptx')  # 自动 full_cleanup
```

**详细 API**：`references/framework/engine-api.md`
**版式规范**：`references/layouts/`
**10条护栏**：`references/framework/guard-rails.md`
**版式能力边界**：`references/layout-matrix.yaml`
**历史踩坑**：`experiences/`
**S3 门禁脚本**：`references/scripts/gate_check_s3.py`
**S4 门禁脚本**：`references/scripts/gate_check.py`
