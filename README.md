<div align="center">

# Mck PPT Design Skill

一套完整的麦肯锡风格 PowerPoint 设计体系
<br/>基于 `python-pptx` 从零生成专业级演示文稿 | v1.3.0

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21+-orange.svg)](https://python-pptx.readthedocs.io)

</div>

---

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
> 基于麦肯锡 2023 官方模板研究，新增 7 大类共 32 种页面布局，覆盖咨询项目全场景：
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
│   └── layout-catalog.md    # 36 种布局目录
└── examples/
    ├── minimal_example.py   # 2 页 Demo（兼容旧路径）
    └── requirements.txt
```

---

### 环境要求

Python 3.8+ · python-pptx ≥ 0.6.21 · lxml ≥ 4.9.0

---

### 社区

<table>
  <tr>
    <td align="center" width="50%">
      <strong>微信交流群</strong><br/><br/>
      <img width="400" height="400" alt="Clipboard_Screenshot_1772519369" src="https://github.com/user-attachments/assets/7c60a36d-9752-4321-aa7b-8cd93580086e" />
    </td>
    <td align="center" width="50%">
      <strong>Discord</strong><br/><br/>
      <!-- 替换 YOUR_INVITE_LINK -->
      <a href="https://discord.gg/YOUR_INVITE_LINK"><img src="https://img.shields.io/badge/Discord-加入社区-5865F2?style=for-the-badge&logo=discord&logoColor=white" alt="Discord" /></a>
    </td>
  </tr>
</table>

---

### 参与贡献

欢迎提交 Issue 和 Pull Request。贡献方向：

- 新增布局模式（时间轴页、数据图表页、2x2矩阵等）
- 扩展色彩主题（深色模式、品牌定制）
- 补充示例代码与文档翻译

---

<div align="center">
<sub>Apache 2.0 · <a href="https://github.com/likaku/Mck-ppt-design-skill">GitHub</a> · <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">反馈建议</a></sub>
</div>
