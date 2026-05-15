# md2word — Markdown 转 Word 文档工具

> English version: [README-en.md](README-en.md)

将 Markdown 文档转换为格式精美的 Word (.docx) 文件，支持自定义样式模板。支持图片（本地/URL）、表格、代码块、任务列表等。

## 安装

```bash
pip install -e .
```

需要 Python 3.10+，依赖：python-docx、markdown、Pillow、requests。

## 用法

```bash
# 基本转换 — md-example.md → md-example.docx
md2word md-example.md -o output.docx

# 使用自定义模板
md2word md-example.md -t 模板文件.docx -o output.docx

# 列出可用主题
md2word --list-themes

# 生成模板以便自定义
md2word --create-template 我的模板.docx --theme official
```

也可读取标准输入：

```bash
cat document.md | md2word -o output.docx
```

## 工作原理

工具通过模板 docx 中的**引导段落**来识别样式。每个引导段落包含一个关键词，告诉工具它代表什么样式元素。工具读取该段落的**格式**（字体、字号、颜色、对齐、缩进等），并应用到对应的 Markdown 内容上。

| 关键词 | 样式槽 | Markdown 元素 |
|--------|--------|---------------|
| 一级标题 | h1 | `# 标题` |
| 二级标题 | h2 | `## 标题` |
| 三级标题 | h3 | `### 标题` |
| 正文 | body | 普通段落 |
| 首行缩进 | body_indent | 首行缩进的正文 |
| 图片 | image | 图片容器 |
| 图注 | figcaption | 图片说明（alt 文本） |
| 引用 | quote | `> 引用块` |
| 代码 | code | `` `行内代码` `` / 代码块 |
| 无序列表 | bullet_list | `- 项目` |
| 有序列表 | number_list | `1. 项目` |

**自定义方法**：在 Word 中打开生成的模板，修改引导段落的格式，保存即可。工具会自动检测您的更改。

## 内置主题

### 官方公文
仿照《党政机关公文格式》标准：
- 仿宋三号正文，黑体标题
- 大字号（三号=16pt），1.5 倍行距
- 页边距上 3.7cm/下 3.5cm/左 2.8cm/右 2.6cm

### 学术论文
严谨学术排版风格：
- 全宋体系列，黑体标题居中
- 首行缩进 2 字符，1.5 倍行距
- 标准页边距（上下 2.54cm、左右 3.17cm）

### 技术文档
现代清晰技术文档风格：
- 微软雅黑字体，深蓝色层级标题
- 代码等宽字体（Consolas + 等线）
- 紧凑排版，信息密度高

### 自媒体排版
视觉系自媒体排版风格：
- 大号标题（32pt），品牌橙色点缀
- 1.8 倍超高行距，宽松版心
- 楷体大字号引用，舒适阅读体验

## 功能特性

- **图片**：自动下载本地和 URL 图片，缩放至页面合适宽度
- **图注**：图片 alt 文本自动置于图片下方居中显示
- **任务列表**：`[x]` 和 `[ ]` 渲染为 ☑/☐ 复选框
- **表格**：Markdown 表格渲染为浅色边框的干净表格
- **代码块**：等宽字体，逐行渲染
- **引用块**：完整保留，支持行内格式
- **列表**：无序和有序列表，正确缩进
- **行内格式**：**粗体**、*斜体*、`代码`、[链接](/) 完整保留
- **中文适配**：全部显式设置东亚字体，杜绝日文字体回退（MS Mincho 等）
- **段落换行**：Markdown 中的换行自动转为段落标记（^p），而非手动换行符

## 项目结构

```
src/md2word/
├── cli.py         — CLI 入口：参数解析、4 套主题模板
├── converter.py   — 核心转换：MD → HTML → docx，含样式提取
├── template.py    — 从引导段落提取样式格式
├── image_utils.py — 图片下载、缩放、嵌入
template/
├── 官方公文.docx
├── 学术论文.docx
├── 技术文档.docx
└── 自媒体排版.docx
```
