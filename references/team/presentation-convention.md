# Presentation Convention — McKinsey PPT Design

> 团队级汇报体例规范。适用于所有 PPT 项目。

## Slide Dimensions

```python
prs.slide_width  = Inches(13.333)  # 宽屏 16:9
prs.slide_height = Inches(7.5)
```

## Standard Margins

| Position | Size | Usage |
|----------|------|-------|
| Left margin (LM) | 0.8" | 默认左边距 |
| Right margin | 0.8" | 默认右边距 |
| Content width (CW) | 11.733" | LM 到右边距 |
| Content start (below title) | 1.3–1.4" | 内容区上边界 |
| Source line | 7.05" | 出处文字基线 |
| Action title top | 0.15" | 标题顶部 |
| Action title height | 0.9" | 标题栏高度 |

## Mandatory Slide Elements

每个内容页（除 Cover 和 Closing 外）**必须**包含：

| 元素 | 方法 | 位置 |
|------|------|------|
| Action Title | `add_action_title()` / `aat()` | 顶部 0.15" |
| 标题分隔线 | Action title 自动包含 | 1.05" |
| 内容区 | 版式相关 | 1.3"–6.5" |
| Source 出处 | `add_source()` | 7.05" |
| 页码 | `add_page_number(n, total)` | 右下角 |

## Title Style Rule

**唯一标准**：所有内容页使用 `add_action_title()` (白底+黑字+下划线)。

**已废弃**：`add_navy_title_bar()` (深蓝底+白字) — 禁止在内容页使用。

## Source Attribution

每个内容页必须有 source，格式：`Source: [来源机构/报告 年份]`

## Page Numbering

```python
def add_page_number(slide, num, total):
    add_text(slide, Inches(12.2), Inches(7.1), Inches(1), Inches(0.3),
             f"{num}/{total}", font_size=Pt(9), font_color=MED_GRAY,
             alignment=PP_ALIGN.RIGHT)
```

## Slide Count Guidelines

- **标准汇报**：10–12页
- **简短汇报**：6–8页
- **最少**：8页（实质性主题）
- **时长约束**：约 1 分钟/页

## File Delivery

Python 引擎路径：`~/.workbuddy/skills/mck-ppt-design`

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
```

`eng.save()` 自动执行三层 XML 清理，无需手动调用 `full_cleanup()`。
