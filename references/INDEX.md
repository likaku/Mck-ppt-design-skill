# Knowledge Index — PPT Skill Reference Router

> AI 在每个阶段应该读哪些文件。按需加载，不要全量读取。

## Stage → Load Map

| 阶段 | 必须读 | 可选读（按需） |
|------|--------|--------------|
| S1 需求定义 | `team/brand-guide.md`（了解约束）| — |
| S2 结构设计 | `framework/engine-api.md`, `references/layout-matrix.yaml` | 涉及具体版式时读对应 `layouts/*.md` |
| S3 内容填充 | `framework/guard-rails.md`, `experiences/*.md`（若存在）| `framework/planning-guide.md`（复杂结构时）|
| S4 渲染 | 对应版式的 `layouts/*.md`（只读用到的）| `team/presentation-convention.md`（确认边距）|
| S5 交付 | — | — |

## File Directory

```
references/
├── INDEX.md                      # 本文件，知识路由表
├── layout-matrix.yaml            # 版式×能力边界×字符预算（真相源）
│
├── team/                         # 团队级（最稳定，所有项目继承）
│   ├── brand-guide.md            # 配色、字体、设计原则
│   └── presentation-convention.md # 边距、必须元素、文件交付
│
├── framework/                    # 框架级（中频更新）
│   ├── engine-api.md             # 67方法 API 速查表
│   ├── guard-rails.md            # 10条生产护栏（防崩溃规则）
│   └── planning-guide.md        # 版式选择、内容密度、结构规划
│
├── layouts/                      # 版式级（按需加载）
│   ├── structure.md              # cover, toc, section_divider, closing
│   ├── data-stats.md             # big_number, two_stat, three_stat, data_table, metric_cards
│   ├── frameworks.md             # matrix_2x2, table_insight, process_chevron, temple, venn
│   ├── comparison.md             # side_by_side, before_after, pros_cons, rag_status, scorecard
│   ├── narrative.md              # executive_summary, key_takeaway, quote, four_column
│   ├── timeline.md               # timeline, vertical_steps, cycle, funnel, value_chain
│   ├── team-cases.md             # meet_the_team, case_study, action_items
│   ├── charts-circular.md        # donut, pie, gauge (BLOCK_ARC 规范)
│   ├── charts-bar-line.md        # grouped_bar, stacked_bar, horizontal_bar, line_chart
│   ├── charts-advanced.md        # waterfall, pareto, progress_bars, bubble, risk_matrix
│   ├── dashboards.md             # dashboard_kpi, dashboard_table
│   ├── images.md                 # content_right_image, left_image_content, three_images
│   └── special.md                # stakeholder_map, decision_tree, icon_grid, swot
│
├── color-palette.md              # 已有，颜色色板参考
├── layout-catalog.md             # 已有，版式目录索引
│
└── scripts/                      # 门禁脚本（机读化执行，不可用 AI 自评替代）
    ├── gate_check.py             # S4 门禁：QA 分类，输出 gate_result.json
    └── gate_check_s3.py          # S3 门禁：API格式+数量+source检查，输出 gate_s3.json
```

## experiences/ Directory

```
experiences/                      # Self-Refinement 经验沉淀（AI 自动维护）
├── overflow.md                   # 文字溢出类经验
├── chart-limits.md               # 图表数据段数限制经验
├── layout-pitfalls.md            # 版式特定踩坑
└── cjk-issues.md                 # 中文渲染问题
```

## Quick Load Cheatsheet

**只需要做一个封面** → 读 `team/brand-guide.md` 即可

**做完整 10+ 页汇报** → S2 读 `engine-api.md` + `layout-matrix.yaml`；S3 读 `guard-rails.md`；S4 只读用到版式的文件

**有图表** → S4 还要读 `layouts/charts-*.md`

**之前做过同类汇报有报错** → S3 先读 `experiences/*.md`
