<div align="center">

# Mck PPT Design Skill

一套完整的咨询公司风格 PowerPoint 设计体系
<br/>基于 `python-pptx` 从零生成专业级演示文稿 | v1.10.2

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21+-orange.svg)](https://python-pptx.readthedocs.io)

</div>

---

### 社区

<table>
<tr>
    <td align="center" width="50%" valign="top">
      <strong>微信交流群</strong><br/><br/>
      <img width="180" src="https://github.com/user-attachments/assets/d4eb704e-3825-4380-ac54-2fbbe4c993ce" alt="WeChat Group" />
    </td>
    <td align="center" width="50%" valign="top">
      <strong>Discord</strong><br/><br/>
      <a href="https://discord.gg/SaFybFAT">
        <img src="https://img.shields.io/badge/Discord-加入社区-5865F2?style=for-the-badge&logo=discord&logoColor=white" alt="Discord" />
      </a>
      <br/><br/>
      <span>点击上方按钮加入</span>
    </td>
  </tr>
</table>

---


> ### v1.10.2 更新 — #54 矩阵侧面板变体
>
> - **新增 #54 Heat Matrix 侧面板布局变体** — 当矩阵需要搭配洞察面板（如"竞争关键变化"、"行动项"）时，使用紧凑九宫格（~60%宽）+ 侧面板（~38%宽）的组合布局
>   - 九宫格单元格宽从 `3.0"` 缩至 `2.15"`，Y轴标签区从 `1.8"` 缩至 `0.65"`
>   - 侧面板从 ~1.4"（被挤压不可读）扩至 ~4.2"（可容纳 6+ 条目）
>   - 含 ASCII 线框图、布局数学代码、深色底部小结栏、最小宽度规则
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.10.1 更新 — YAML Frontmatter 修复
>
> - **修复 Claude 安装报错** — Claude 的 SKILL.md 解析器仅支持 `name` + `description` 两个 frontmatter 字段
>   - 移除 7 个不支持的字段（`license`, `version`, `author`, `homepage` 等）
>   - `description` 改用 YAML folded block scalar (`>-`)，兼容性更好
>   - 元数据信息迁移至正文 blockquote，无信息丢失
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.10.0 更新 — 频道文件投递
>
> - **新增 Channel Delivery 能力** — 在飞书、Telegram、WhatsApp、Discord、Slack 等频道中直接回传生成的 PPTX 文件
>   - `deliver_to_channel()` 辅助函数 — 通过 `openclaw message send --media` 自动投递
>   - 环境感知 — IDE / CI 等非频道环境下静默跳过，打印本地路径
>   - 支持所有文档类型，最大 100MB
> - **更新 `minimal_example.py`** — 生成后自动调用投递流程
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.9.0 更新 — 生产级质量护栏
>
> - **新增 Production Guard Rails** — 7 条从实际生产反馈提炼的强制规则：
>   - 内容块与底部栏间距保护（最小 0.15"）
>   - 内容溢出检测（右边距 + 底部边距 + 容器内文字内缩）
>   - 底部留白消除（图表/内容填满纵向空间）
>   - 图例颜色一致性（必须用 `add_rect()` 色块，禁止纯文本"■"）
>   - 标题风格统一（仅使用 `add_action_title()`，`add_navy_title_bar()` 已废弃）
>   - 矩阵图轴标签居中（基于实际网格尺寸计算）
>   - 图片占位页强制要求（8+ 页 PPT 至少含 1 页图片占位）
> - **新增 Code Efficiency Guidelines** — 常量提取、辅助函数复用、标准缩写表、批量数据结构、自动页码计数
> - **新增 5 个 Common Issues**（Problem 6-10）：容器溢出、图例色差、标题风格混乱、轴标签偏移、底部留白
> - 基于 19 页 AI 行业报告的多轮迭代反馈提炼
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.8.0 更新 — 大规模布局扩展
>
> - **布局总数 39 → 70**，类别 8 → 12，新增 31 个专业布局模板
> - **新增 Category I：图片+内容布局** — 8 种含图片占位符的布局（#40-#47）：
>   - 内容+右侧图片、左侧图片+内容、三图对比、图片+四要点、全幅图片叠加文字、带图案例研究、引言+背景图、目标+配图
> - **新增 Category J：高级数据可视化** — 9 种纯 `add_rect()` 手绘图表（#48-#56）：
>   - 环形图、瀑布图、折线/趋势图、帕累托图、进度条/KPI追踪、气泡/散点图、风险热力矩阵、仪表盘、Harvey Ball 评估表
> - **新增 Category K：仪表盘布局** — 2 种数据密集型执行仪表盘（#57-#58）
> - **新增 Category L：视觉叙事** — 12 种视觉叙事模板（#59-#70）：
>   - 利益相关者地图、决策树、检查清单、指标对比行、图标网格、饼图、SWOT分析、议程、价值链、双列图文、编号列表+面板、堆积面积图
> - **新增 `add_image_placeholder()` 辅助函数** — 灰色占位矩形+十字线+标签，用户生成后替换为真实图片
> - **新增 Image Priority Rule** — 涉及案例、产品展示等内容时优先使用图片布局
> - 基于 McKinsey PowerPoint Template 2023（679页）系统分析，提取关键模板模式
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.7.0 更新 — 数据图表类别
>
> - **新增 Category H：数据图表** — 3 种纯 `add_rect()` 手绘图表布局，无需 matplotlib 依赖：
>   - **#37 Grouped Bar Chart（分组柱状图）** — 多类别在不同时间点的数值对比
>   - **#38 Stacked Bar Chart（堆叠柱状图）** — 各部分占比随时间变化
>   - **#39 Horizontal Bar Chart（水平柱状图）** — 类别排名对比
> - **新增 Chart Priority Rule** — 当数据包含日期 + 数值/百分比时，必须优先使用图表模式
> - **布局总数 36 → 39**，类别 7 → 8
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.6.0 更新 — 跨模型质量对齐
>
> - **新增强调色系统** — 4 组强调色（Blue #006BA6 / Green #007A53 / Orange #D46A00 / Red #C62828）+ 配套浅色背景，用于多项目视觉区分
> - **新增 Presentation Planning 章节** — 幻灯片结构模板、布局多样性规则、内容密度要求、必备元素规范
> - **新增 `add_page_number()` 辅助函数** — 自动在右下角显示 "N/Total" 页码
> - 基于 Opus 4.6 / Minimax 2.5 / Hunyuan 2 think / GLM5 四模型对比分析，针对性缩小弱模型与 Opus 的输出质量差距
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.5.0 更新 — 行距修复
>
> - **修复中文多行文字重叠问题** — `add_text()` 新增 `p.line_spacing = Pt(font_size.pt * 1.35)` 显式设置行距
> - 此前仅设置段前距（`space_before`），未设置行高（`lnSpc`），导致自动换行的中文行互相叠压
> - Common Issues 新增 Problem 5 详细说明根因与修复
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.4.0 更新 — P0 优化
>
> - 合并 `add_text()` / `add_multiline()` 为统一函数，传 `str` 单行，传 `list` 多行
> - 更新全部 36 个布局模板调用，参数名 `line_spacing_pt` → `line_spacing=Pt(N)`
> - 删除 DEPRECATED connector 说明、v1.1 改进备注等冗余内容
> - 合并 Common Issues + Error Handling，移除重复项
> - 删除 Refining Existing Presentations 章节
> - 净减 109 行 (~4.2KB)，降低每次生成的 token 消耗
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.3.0 更新 — ClawHub 首发
>
> - 适配 ClawHub 发布标准，优化 description 提升搜索可发现性
> - 新增 `Edge Cases` 和 `Error Handling` 章节
> - 新增 `scripts/` 和 `references/` 目录，文件结构更规范
> - 添加 `metadata`（依赖声明）和 `homepage`（GitHub 链接）
>
> 详见 [CHANGELOG.md](CHANGELOG.md)

> ### v1.1.0 更新 — 新增 32 种布局模式
>
> 基于模板研究，新增 7 大类共 32 种页面布局，覆盖咨询项目全场景：
>
> | 类别 | 新增布局 |
> |:-----|:---------|
> | **结构导航** | 章节分隔页、目录/议程页、附录标题页 |
> | **数据统计** | 大数据展示、双数据对比、三指标仪表盘、数据表格、指标卡片行 |
> | **框架矩阵** | 四象限矩阵、三支柱框架、金字塔层级、流程箭头、维恩图概念、殿堂框架 |
> | **对比评估** | 左右对比、前后对比、优劣分析、红绿灯状态、计分卡 |
> | **内容叙事** | 执行摘要、核心洞见、引言/洞见页、双栏文本、四栏概览 |
> | **时间流程** | 时间轴/路线图、垂直步骤、循环图、漏斗图 |
> | **团队专题** | 团队介绍、案例研究、行动计划、结束页 |
>
> 同时重构底层渲染引擎，通过三层防御体系（矩形画线 + 内联清理 + 全量XML清洗）彻底解决文件损坏问题。

---

### 样例展示

| 封面页 | 内容页 | 表格页 |
|:------:|:------:|:------:|
| <img width="600" alt="封面页" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" /> | <img width="600" alt="内容页" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" /> | <img width="600" alt="表格页" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" /> |
| **四栏布局** | **色彩体系** | **总结页** |
| <img width="600" alt="四栏布局" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" /> | <img width="600" alt="色彩体系" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" /> | <img width="600" alt="总结页" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" /> |

---

### 它解决什么问题

- 手动排版 PPT 耗时耗力，团队间设计风格难以统一
- `python-pptx` 默认生成的文件带阴影 / 3D / `p:style` 引用，PowerPoint 打开报错或提示修复
- 中文字体渲染需要特殊处理，否则显示异常
- AI 生成的 PPT 缺乏专业设计感，每次产出质量不一致

**本 Skill 将完整的麦肯锡设计规范编码为一份文档**，AI 读取后即可持续输出风格统一的专业 PPT。

---

### 设计原则

| 原则 | 说明 |
|:-----|:-----|
| **极简主义** | 移除一切非必要视觉元素，无渐变、无装饰 |
| **扁平设计** | 无阴影、无 3D、无反射，纯实色填充 |
| **严格层次** | 标题 22pt → 子标题 18pt → 正文 14pt → 脚注 9pt |
| **全局一致** | 统一色板、字体、间距，贯穿每一页 |

---

### 色彩体系

| 名称 | 色块 | Hex | 用途 |
|:-----|:----:|:---:|:-----|
| **NAVY** | ![](docs/colors/navy.png) | `#051C2C` | 主色调 — 标题、圆形指标、TOC 高亮 |
| **BLACK** | ![](docs/colors/black.png) | `#000000` | 分隔线、标题下划线、表头线 |
| **DARK_GRAY** | ![](docs/colors/dark-gray.png) | `#333333` | 正文文本 |
| **MED_GRAY** | ![](docs/colors/med-gray.png) | `#666666` | 次要文本、标签、来源注释 |
| **LINE_GRAY** | ![](docs/colors/line-gray.png) | `#CCCCCC` | 表格行分隔线 |
| **BG_GRAY** | ![](docs/colors/bg-gray.png) | `#F2F2F2` | 背景面板、Takeaway 区域 |

**强调色（v1.6.0 新增）** — 用于 3+ 并列项的视觉区分：

| 名称 | Hex | 配套浅色背景 | 用途 |
|:-----|:---:|:---:|:-----|
| **ACCENT_BLUE** | `#006BA6` | `#E3F2FD` | 第一项强调 |
| **ACCENT_GREEN** | `#007A53` | `#E8F5E9` | 第二项强调 |
| **ACCENT_ORANGE** | `#D46A00` | `#FFF3E0` | 第三项强调 |
| **ACCENT_RED** | `#C62828` | `#FFEBEE` | 第四项 / 警示 |

---

### 核心技术

**文件兼容性保障（v1.1 三层防御）**

python-pptx 自动为形状附加 `<p:style>` 元素，引用主题中的 `outerShdw`、`effectRef` 等效果，导致文件在 PowerPoint 中无法打开或提示修复。本 Skill 通过三道防线彻底解决：

1. **不使用 connector** — 所有线条用极细矩形（`add_hline()`）绘制，从源头杜绝 connector 的 `p:style`
2. **内联清理** — 每个 `add_rect()` 和 `add_oval()` 创建后立即调用 `_clean_shape()` 移除 `p:style`
3. **保存后全量清洗** — `full_cleanup()` 遍历所有 slide XML + theme XML，移除全部 `p:style`、阴影和 3D 节点

**中文字体处理**

所有含中文的段落需调用 `set_ea_font(run, 'KaiTi')` 设置东亚字体，否则中文将以默认英文字体渲染。

---

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

---

### 项目结构

```
├── SKILL.md                 # 核心设计规范
├── LICENSE                  # Apache 2.0
├── CHANGELOG.md             # 版本记录
├── scripts/
│   ├── minimal_example.py   # 2 页 Demo
│   └── requirements.txt     # 依赖列表
├── references/
│   ├── color-palette.md     # 色彩速查
│   └── layout-catalog.md    # 70 种布局目录
└── examples/
    ├── minimal_example.py   # 2 页 Demo（兼容旧路径）
    └── requirements.txt
```

---

### 环境要求

Python 3.8+ · python-pptx ≥ 0.6.21 · lxml ≥ 4.9.0

---



### 参与贡献

欢迎提交 Issue 和 Pull Request。贡献方向：

- 新增布局模式（时间轴页、2x2矩阵等）
- 扩展色彩主题（深色模式、品牌定制）
- 补充示例代码与文档翻译

---

<div align="center">
<sub>Apache 2.0 · Copyright © 2026 <strong>likaku</strong> · <a href="https://github.com/likaku/Mck-ppt-design-skill">GitHub</a> · <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">反馈建议</a></sub>
</div>
