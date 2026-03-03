# Mck PPT Design Skill

一套完整的**麦肯锡风格 PowerPoint 设计体系**，基于 `python-pptx` 从零生成专业级、顾问级演示文稿。

> 让 AI 像麦肯锡顾问一样输出 PPT，每次生成都保持一致的设计水准。

---

## 核心特性

- **完整设计体系** — 色彩体系、字体层级、线条规范、布局模式，一套文档覆盖全部设计决策
- **即插即用代码** — 复制 Helper 函数即可开始，无需从零摸索 python-pptx API
- **无阴影渲染** — 双重防御方案，彻底消除 OOXML 阴影与 3D 伪影
- **中文字体支持** — 楷体（KaiTi）东亚字体完美渲染，告别中文乱码
- **多种布局模式** — 封面、Action Title、表格、三栏概览、目录页等 5 种标准版式
- **主题后处理清理** — 自动化 XML 处理，确保输出文件在 PowerPoint 中无任何瑕疵

---

## PPT 样例展示

使用本 Skill 生成的演示文稿效果：

<!-- 
  在此处添加 PPT 样例截图
  建议将截图放入 docs/screenshots/ 目录，然后引用：
  ![封面页](docs/screenshots/cover.png)
  ![内容页](docs/screenshots/content.png)
  ![表格页](docs/screenshots/table.png)
-->

| 封面页 | 内容页 | 表格页 |
|--------|--------|--------|
| <img width="2192" height="1228" alt="Clipboard_Screenshot_1772517587" src="https://github.com/user-attachments/assets/075ec46d-dd73-4454-92d0-84184b78d276" />
| <img width="2188" height="1226" alt="Clipboard_Screenshot_1772517603" src="https://github.com/user-attachments/assets/3b25f071-8a81-48e3-a62b-9d9be9026f2e" />
| <img width="2188" height="1228" alt="Clipboard_Screenshot_1772517625" src="https://github.com/user-attachments/assets/be327c14-aff9-459f-89b0-d4a8bffaabfc" />
|

| 四栏布局 | 色彩体系 | 总结页 |
|----------|----------|--------|
|<img width="2188" height="1230" alt="Clipboard_Screenshot_1772517657" src="https://github.com/user-attachments/assets/687cee47-13bb-4d6b-840f-77f8e001a62b" />
| <img width="2190" height="1232" alt="Clipboard_Screenshot_1772517677" src="https://github.com/user-attachments/assets/41371c47-608f-4857-9bfe-791121ec1579" />
|<img width="2192" height="1234" alt="Clipboard_Screenshot_1772517694" src="https://github.com/user-attachments/assets/c5b6e52a-fd91-4c28-88a4-82fdfedfd956" />
|

> **提示**：将截图放入 `docs/screenshots/` 目录即可自动显示。推荐尺寸：1920x1080 或 2560x1440。

---

## 设计原则

| 原则 | 说明 |
|------|------|
| **极简主义** | 移除一切非必要视觉元素，无渐变、无装饰 |
| **扁平设计** | 无阴影、无 3D、无反射，纯实色填充 |
| **严格层次** | 标题 (22pt) > 子标题 (18pt) > 正文 (14pt) > 脚注 (9pt) |
| **全局一致** | 统一色板、字体、间距贯穿每一页幻灯片 |

---

## 色彩体系

| 名称 | Hex 值 | 用途 |
|------|--------|------|
| **NAVY** | `#051C2C` | 主色调：标题、圆形指标、TOC 高亮 |
| **BLACK** | `#000000` | 分隔线、标题下划线、表头线 |
| **DARK_GRAY** | `#333333` | 正文文本、主要描述内容 |
| **MED_GRAY** | `#666666` | 次要文本、标签、来源注释 |
| **LINE_GRAY** | `#CCCCCC` | 表格行分隔线、时间轴线 |
| **BG_GRAY** | `#F2F2F2` | 背景面板、Key Takeaway 区域 |

---

## 快速上手

```bash
# 1. 安装依赖
pip install python-pptx lxml

# 2. 运行示例
cd examples
python minimal_example.py

# 3. 安装为 Skill（可选）
cp SKILL.md ~/.workbuddy/skills/workbuddy-ppt-design/SKILL.md
```

将 `SKILL.md` 放入 Skill 目录后，创建 PPT 时会自动加载本设计体系。

---

## 项目结构

```
.
├── SKILL.md                    # 核心设计规范文档（685 行）
├── README.md                   # 本文件
├── LICENSE                     # Apache 2.0 许可证
├── CHANGELOG.md                # 版本变更记录
├── .gitignore                  # Git 排除配置
├── docs/
│   └── screenshots/            # PPT 样例截图
└── examples/
    ├── minimal_example.py      # 可运行的 2 页示例
    └── requirements.txt        # Python 依赖列表
```

---

## 关键技术细节

### 阴影消除（核心）

python-pptx 会自动为连接器附加 `<p:style>` 元素，引用主题中包含 `outerShdw`（外阴影）的效果样式。本 Skill 提供双重防御：

1. **内联移除** — 创建连接器后立即删除 `<p:style>`，阻止阴影生效
2. **主题清理** — 保存后通过 zipfile + lxml 处理 theme XML，移除所有 `outerShdw` / `innerShdw` / `scene3d` / `sp3d`

### 中文字体处理

所有含中文的段落必须调用 `set_ea_font(run, 'KaiTi')` 设置东亚字体。该函数通过操作 `<a:ea typeface="KaiTi">` XML 节点实现，需在每个 text run 上调用。

---

## 环境要求

- Python 3.8+
- python-pptx >= 0.6.21
- lxml >= 4.9.0

---

## 社区与交流

欢迎加入社区，交流 PPT 设计体系、分享作品、提出建议！

### 微信交流群

<!--
  在此处放置微信群二维码图片
  建议将二维码放入 docs/ 目录：
  ![微信群二维码](docs/wechat-qr.png)
-->

<p align="center">
  <img src="docs/wechat-qr.png" alt="微信交流群" width="200" />
  <br/>
  <em>扫码加入微信交流群（如二维码过期，请提 Issue 获取最新链接）</em>
</p>

### Discord

<!--
  替换为实际的 Discord 邀请链接
-->

[![Discord](https://img.shields.io/badge/Discord-加入社区-5865F2?style=for-the-badge&logo=discord&logoColor=white)](https://discord.gg/YOUR_INVITE_LINK)

> 加入 Discord 社区讨论设计规范、分享生成效果、获取最新更新。

---

## 许可证

Apache License 2.0 — 详见 [LICENSE](LICENSE)。

---

## 参与贡献

欢迎提交 Issue 和 Pull Request！请确保贡献内容遵循 `SKILL.md` 中定义的设计原则。

**贡献方向**：
- 新增布局模式（如时间轴页、数据图表页）
- 扩展色彩主题（深色模式、品牌定制色）
- 补充更多示例代码
- 完善文档和翻译

---

<p align="center">
  <strong>Mck PPT Design Skill</strong> — 让每一份 PPT 都有麦肯锡的质感
  <br/>
  <a href="https://github.com/likaku/Mck-ppt-design-skill">GitHub</a> · 
  <a href="https://github.com/likaku/Mck-ppt-design-skill/issues">反馈建议</a> · 
  <a href="https://discord.gg/YOUR_INVITE_LINK">Discord</a>
</p>
