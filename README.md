<div align="center">

# Mck PPT Design Skill

**麦肯锡风格 PowerPoint 设计体系**

基于 `python-pptx` 从零生成专业级、顾问级演示文稿

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python&logoColor=white)](https://python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21+-orange.svg)](https://python-pptx.readthedocs.io)

*让 AI 像麦肯锡顾问一样输出 PPT，每次生成都保持一致的设计水准*

</div>

<br/>

## 核心特性

|  | 特性 | 说明 |
|:--|:-----|:-----|
| :art: | **完整设计体系** | 色彩、字体、线条、布局，一套文档覆盖全部设计决策 |
| :rocket: | **即插即用代码** | 复制 Helper 函数即可开始，无需从零摸索 API |
| :sparkles: | **无阴影渲染** | 双重防御方案，彻底消除 OOXML 阴影与 3D 伪影 |
| :cn: | **中文字体支持** | 楷体东亚字体完美渲染，告别中文乱码 |
| :jigsaw: | **5 种布局模式** | 封面、Action Title、表格、三栏概览、目录页 |
| :wrench: | **主题后处理** | 自动化 XML 清理，确保 PowerPoint 无任何瑕疵 |

<br/>

## PPT 样例展示

使用本 Skill 生成的演示文稿效果：

| 封面页 | 内容页 | 表格页 |
|:------:|:------:|:------:|
| <img width="600" alt="封面页" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" /> | <img width="600" alt="内容页" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" /> | <img width="600" alt="表格页" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" /> |

| 四栏布局 | 色彩体系 | 总结页 |
|:--------:|:--------:|:------:|
| <img width="600" alt="四栏布局" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" /> | <img width="600" alt="色彩体系" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" /> | <img width="600" alt="总结页" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" /> |

<br/>

## 设计原则

<table>
  <tr>
    <td width="25%" align="center"><strong>极简主义</strong><br/><sub>移除一切非必要视觉元素<br/>无渐变、无装饰</sub></td>
    <td width="25%" align="center"><strong>扁平设计</strong><br/><sub>无阴影、无 3D、无反射<br/>纯实色填充</sub></td>
    <td width="25%" align="center"><strong>严格层次</strong><br/><sub>标题 22pt → 子标题 18pt<br/>正文 14pt → 脚注 9pt</sub></td>
    <td width="25%" align="center"><strong>全局一致</strong><br/><sub>统一色板、字体、间距<br/>贯穿每一页</sub></td>
  </tr>
</table>

<br/>

## 色彩体系

```
  NAVY        BLACK       DARK_GRAY   MED_GRAY    LINE_GRAY   BG_GRAY
  #051C2C     #000000     #333333     #666666     #CCCCCC     #F2F2F2
  ██████      ██████      ██████      ██████      ██████      ██████
  主色调       分隔线       正文文本     次要文本     表格行线     背景面板
```

| 名称 | Hex 值 | 用途 |
|:-----|:------:|:-----|
| **NAVY** | `#051C2C` | 标题、圆形指标、TOC 高亮、封面装饰线 |
| **BLACK** | `#000000` | 分隔线、标题下划线、表头线 |
| **DARK_GRAY** | `#333333` | 正文文本、主要描述内容 |
| **MED_GRAY** | `#666666` | 次要文本、标签、来源注释 |
| **LINE_GRAY** | `#CCCCCC` | 表格行分隔线、时间轴线 |
| **BG_GRAY** | `#F2F2F2` | 背景面板、Key Takeaway 区域 |

<br/>

## 快速上手

```bash
# Step 1: 安装依赖
pip install python-pptx lxml

# Step 2: 运行示例，生成一份 2 页的 Demo PPT
cd examples && python minimal_example.py

# Step 3（可选）: 安装为 AI Agent Skill
mkdir -p ~/.workbuddy/skills/workbuddy-ppt-design
cp SKILL.md ~/.workbuddy/skills/workbuddy-ppt-design/
```

将 `SKILL.md` 放入 Skill 目录后，创建 PPT 时会自动加载本设计体系。

<br/>

## 项目结构

```
Mck-ppt-design-skill/
├── SKILL.md                    # 核心设计规范文档（685 行完整规范）
├── README.md                   # 本文件
├── LICENSE                     # Apache 2.0 许可证
├── CHANGELOG.md                # 版本变更记录
├── .gitignore                  # Git 排除配置
├── docs/
│   └── screenshots/            # PPT 样例截图
└── examples/
    ├── minimal_example.py      # 可运行的 2 页 Demo
    └── requirements.txt        # Python 依赖列表
```

<br/>

## 关键技术细节

### 阴影消除（核心难点）

python-pptx 会自动为连接器附加 `<p:style>` 元素，引用主题中包含 `outerShdw`（外阴影）的效果样式，导致线条出现阴影甚至文件损坏。

**双重防御方案：**

> **防御一：内联移除** — 创建连接器后立即删除 `<p:style>`，从源头阻止阴影生效
>
> **防御二：主题清理** — 保存后通过 `zipfile` + `lxml` 处理 theme XML，移除全部 `outerShdw` / `innerShdw` / `scene3d` / `sp3d`

### 中文字体处理

所有含中文的段落必须调用 `set_ea_font(run, 'KaiTi')` 设置东亚字体，通过操作 `<a:ea typeface="KaiTi">` XML 节点实现，需在每个 text run 上调用。

<br/>

## 环境要求

| 依赖 | 最低版本 |
|:-----|:--------:|
| Python | 3.8+ |
| python-pptx | 0.6.21+ |
| lxml | 4.9.0+ |

<br/>

## 社区与交流

欢迎加入社区，交流设计体系、分享作品、提出建议！

<table>
  <tr>
    <td align="center" width="50%">
      <strong>微信交流群</strong><br/><br/>
      <!-- 将微信群二维码图片替换到 docs/wechat-qr.png 即可显示 -->
      <code>📱 二维码即将上线</code><br/><br/>
      <sub>如需加群请提 <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">Issue</a></sub>
    </td>
    <td align="center" width="50%">
      <strong>Discord</strong><br/><br/>
      <!-- 替换 YOUR_INVITE_LINK 为实际的 Discord 邀请链接 -->
      <a href="https://discord.gg/YOUR_INVITE_LINK">
        <img src="https://img.shields.io/badge/Discord-加入社区-5865F2?style=for-the-badge&logo=discord&logoColor=white" alt="Discord" />
      </a><br/><br/>
      <sub>讨论设计规范 · 分享生成效果 · 获取最新更新</sub>
    </td>
  </tr>
</table>

<br/>

## 参与贡献

欢迎提交 Issue 和 Pull Request！请确保贡献内容遵循 `SKILL.md` 中定义的设计原则。

**贡献方向：**
- :jigsaw: 新增布局模式（时间轴页、数据图表页）
- :art: 扩展色彩主题（深色模式、品牌定制色）
- :page_facing_up: 补充更多示例代码
- :globe_with_meridians: 完善文档和翻译

<br/>

## 许可证

[Apache License 2.0](LICENSE)

<br/>

<div align="center">

**Mck PPT Design Skill** — 让每一份 PPT 都有麦肯锡的质感

[GitHub](https://github.com/likaku/Mck-ppt-design-skill) · [反馈建议](https://github.com/likaku/Mck-ppt-design-skill/issues) · [Discord](https://discord.gg/YOUR_INVITE_LINK)

</div>
