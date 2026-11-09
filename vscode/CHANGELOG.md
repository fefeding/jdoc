# Changelog

## 0.1.25

- **SEO / Marketplace 优化**：显示名改为 `jdoc Office Viewer & Editor`（保留 `Office` 与 `Editor` 关键词）；简短描述与 `README` 补入 Word / Excel / PowerPoint / Office 自然语言词；`keywords` 扩充至 20 个（新增 `office` / `word` / `excel` / `powerpoint` / `converter` 等）；类目由 `Visualization, Other` 调整为 `Editing, Visualization`；新增 `galleryBanner` / `homepage` / `qna` / `bugs`；`README` 增加 badges、安装指引与评分引导，并修复中文 README 的截图 `/blob/` 链接。

## 0.1.24

- **PPTX → HTML 改为幻灯片版式还原**：每页一张固定画布、元素绝对定位，与 PDF / 扩展预览高度一致（背景图、形状、渐变、连接线箭头、圆角、镜像、虚线全部还原；文本按预览端 flex 对齐，富文本 / 项目符号 / 内边距 / 缩字均保留）。此前走文本块抽取会丢失版面。
- **PPTX → Markdown 加分页与小标题层级**：页与页之间用 `---` 分隔；页内「小节标题」文本框升级为二级标题；图片不再写幻灯片画布坐标的宽高到 title（其它来源的 md 仍保留尺寸以维持往返）。
