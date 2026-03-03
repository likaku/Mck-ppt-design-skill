<div align="center">

# Mck PPT Design Skill

一套完整的麦肯锡风格 PowerPoint 设计体系
<br/>基于 `python-pptx` 从零生成专业级演示文稿

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21+-orange.svg)](https://python-pptx.readthedocs.io)

</div>

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
- `python-pptx` 默认生成的文件带阴影 / 3D 瑕疵，PowerPoint 打开报错
- 中文字体渲染需要特殊处理，否则显示异常
- AI 生成的 PPT 缺乏专业设计感，每次产出质量不一致

**本 Skill 将完整的麦肯锡设计规范编码为一份 685 行的文档**，AI 读取后即可持续输出风格统一的专业 PPT。

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

**阴影消除（双重防御）**

python-pptx 自动为连接器附加 `<p:style>`，引用主题中的 `outerShdw` 效果，导致线条出现阴影甚至文件损坏。本 Skill 通过两道防线彻底解决：

1. **内联移除** — 创建连接器后立即删除 `<p:style>`，从源头阻断
2. **主题清理** — 保存后用 `zipfile` + `lxml` 处理 theme XML，移除全部阴影和 3D 节点

**中文字体处理**

所有含中文的段落需调用 `set_ea_font(run, 'KaiTi')` 设置东亚字体，否则中文将以默认英文字体渲染。

---

### 快速上手

```bash
# 1. 安装依赖
pip install python-pptx lxml

# 2. 运行示例
cd examples && python minimal_example.py

# 3. 安装为 Skill（可选）
mkdir -p ~/.workbuddy/skills/workbuddy-ppt-design
cp SKILL.md ~/.workbuddy/skills/workbuddy-ppt-design/
```

---

### 项目结构

```
├── SKILL.md                 # 核心设计规范（685 行）
├── LICENSE                  # Apache 2.0
├── CHANGELOG.md             # 版本记录
└── examples/
    ├── minimal_example.py   # 2 页 Demo
    └── requirements.txt     # 依赖列表
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
      <img width="600" height="600" alt="Clipboard_Screenshot_1772519369" src="https://github.com/user-attachments/assets/7c60a36d-9752-4321-aa7b-8cd93580086e" />
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

- 新增布局模式（时间轴页、数据图表页）
- 扩展色彩主题（深色模式、品牌定制）
- 补充示例代码与文档翻译

---

<div align="center">
<sub>Apache 2.0 · <a href="https://github.com/likaku/Mck-ppt-design-skill">GitHub</a> · <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">反馈建议</a></sub>
</div>
