<div align="center">

# MCK PPT Design Skill

**让 AI 智能体直接交付咨询公司风格、每个元素都能编辑、并且由代码检查过的 PPT。**

[![Stars](https://img.shields.io/github/stars/likaku/Mck-ppt-design-skill?style=flat&logo=github)](https://github.com/likaku/Mck-ppt-design-skill/stargazers)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)

[English](README.md) · [简体中文](README.zh-CN.md) · [版式清单](docs/GALLERY.md) · [架构](docs/ARCHITECTURE.md) · [Harness 工程化模式](docs/HARNESS.md)

</div>

<table>
<tr>
<td><img alt="封面" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" /></td>
<td><img alt="战略分析" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" /></td>
<td><img alt="数据表" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" /></td>
</tr>
<tr>
<td><img alt="四栏框架" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" /></td>
<td><img alt="配色" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" /></td>
<td><img alt="执行摘要" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" /></td>
</tr>
</table>

## 它解决什么问题

让 AI 做 PPT，常见结果有两种：要么是白底上堆文字框，要么是一张好看但改不了的图片。这个项目给出第三种做法：

- **真正的 .pptx，每个元素都能改。** 全部是原生形状和文本，不是截图。在 PowerPoint、WPS、Keynote 里打开就能继续编辑。
- **62 种咨询风格版式。** 标题写结论、每页有来源行、深蓝加灰的克制配色、扁平无阴影。
- **智能体负责决策，Python 负责绘制。** `eng.donut(...)` 智能体只输出约 20 个 token，其余由确定性代码完成。30 页的 PPT 输出 token 从 4 到 6 万降到 0.9 到 1.2 万（v1 到 v2 重写时的实测）。
- **质量由代码检查，不由智能体自己判断。** 文字溢出、元素碰撞、图例、字体一致性都用脚本测。`gate_result.json` 里 `passed` 不是 true，就不允许交付。
- **中文可用。** 每段文字都写入东亚字体，楷体、微软雅黑不会变成方块。

## 60 秒试用

不需要 API Key，不需要联网：

```bash
pip install mck-ppt
python -m mck_ppt demo         # 生成 mck_demo.pptx，13 页，通过内置质检
python -m mck_ppt layouts      # 列出全部版式
```

## 在智能体里使用

| 智能体 | 安装方式 |
|---|---|
| **Claude Code** | `/plugin marketplace add likaku/Mck-ppt-design-skill`，再 `/plugin install mck-ppt-design@mck-ppt` |
| **支持 skill 的其他智能体** | `npx skills add likaku/Mck-ppt-design-skill` |
| **WorkBuddy / CodeBuddy** | 在 skill 市场搜索 `mck-ppt-design`，或把本目录复制到 `~/.workbuddy/skills/` |
| **Cursor / Continue / 任意大模型** | 把 `SKILL.md` 作为项目规则，并 `pip install mck-ppt` |

然后直接用中文提需求：

> 用麦肯锡风格做一份 15 页的年度战略汇报：先给执行摘要，再分三章，最后列出需要决策的四件事。每页标题写结论，每页有数据来源。

智能体会按五个阶段工作：需求、结构（`outline.json`）、内容（`content.json`）、渲染、交付。其中两道门禁是脚本，不是智能体口头宣布。详见 [Harness 工程化模式](docs/HARNESS.md)。

## 和其他做法的区别

| | 本项目 | 让模型自由写 python-pptx / PptxGenJS | HTML 幻灯片类 skill |
|---|---|---|---|
| 输出 | 可编辑 .pptx | 可编辑 .pptx | HTML / PDF |
| 版式来源 | 62 个经过测试的版式方法 | 模型自己算坐标 | 模型写 CSS |
| 质量门禁 | 代码执行，不通过不交付 | 通常由模型目测 | 通常没有 |
| 中文排版 | 引擎内置处理 | 靠提示词 | 靠 CSS |
| 适合 | 战略、董事会、咨询风格汇报 | 一次性的自定义设计 | 在浏览器里演讲 |

要在浏览器里演讲，选 HTML 类。文件最后要交给别人用 PowerPoint 打开修改，选本项目。

## 版式

共 62 种在用版式，完整清单和调用签名见 [docs/GALLERY.md](docs/GALLERY.md)，或运行 `python -m mck_ppt layouts`。

## 参与

- 用它做了 PPT？[发出来看看](https://github.com/likaku/Mck-ppt-design-skill/issues/new?template=showcase.yml)，好的会放进 README。
- 缺版式？[提需求](https://github.com/likaku/Mck-ppt-design-skill/issues/new?template=layout_request.yml)。
- 微信群（200 多人）：加微信 `Kashinmoto` 拉你进群。Discord：[加入](https://discord.gg/SaFybFAT)。
- 如果它帮你省了一个晚上，点个 Star，能让更多人搜到它。

---

<div align="center"><sub>Apache 2.0 · © 2024-2026 <strong>Kaku Li</strong></sub></div>
