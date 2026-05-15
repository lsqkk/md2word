# md2word — Markdown to Word Converter

> Chinese version: [README.md](README.md)

Convert Markdown documents to professionally formatted Word (.docx) files using customizable style templates. Supports images (local/URL), tables, code blocks, task lists, and more.

## Installation

```bash
pip install -e .
```

Requires Python 3.10+, python-docx, markdown, Pillow, requests.

## Usage

```bash
# Basic — converts md-example.md → md-example.docx
md2word md-example.md -o output.docx

# Custom template
md2word md-example.md -t my-template.docx -o output.docx

# List available themes
md2word --list-themes

# Generate a template for customization
md2word --create-template my-template.docx --theme official
```

## How It Works

The tool uses **guide paragraphs** in the template docx. Each guide paragraph has a keyword that tells the tool what style element it represents. The tool reads the **formatting** (font, size, color, alignment, indentation) and applies it to matching markdown content.

| Keyword | Slot | Markdown Element |
|---------|------|-----------------|
| 一级标题 | h1 | `# heading` |
| 二级标题 | h2 | `## heading` |
| 三级标题 | h3 | `### heading` |
| 正文 | body | Normal paragraph |
| 首行缩进 | body_indent | Body with first-line indent |
| 图片 | image | Image container |
| 图注 | figcaption | Image caption (alt text) |
| 引用 | quote | `> blockquote` |
| 代码 | code | `` `code` `` / fenced code blocks |
| 无序列表 | bullet_list | `- item` (unordered) |
| 有序列表 | number_list | `1. item` (ordered) |

**To customize**: open a generated template in Word, format the guide paragraphs, and save. The tool detects your changes automatically.

## Built-in Themes

### 官方公文 (Official Document)
- FangSong body, SimHei headings, 16pt, 1.5 line spacing
- Margins: top 3.7cm/bottom 3.5cm/left 2.8cm/right 2.6cm

### 学术论文 (Academic Paper)
- SimSong throughout, SimHei centered headings
- First-line indent 2 chars, 1.5 line spacing
- Standard academic margins

### 技术文档 (Technical Documentation)
- Microsoft YaHei, deep blue hierarchical headings
- Monospace code (Consolas + DengXian)
- Compact layout, high information density

### 自媒体排版 (Social Media / We-Media)
- Large headings (32pt), brand orange accents
- 1.8x line spacing, generous margins
- KaiTi large-size quotes

## Features

- **Images**: Local and URL images auto-downloaded, resized to fit page
- **Captions**: Image alt text placed as centered caption below images
- **Task lists**: `[x]` and `[ ]` rendered as ☑ / ☐ checkboxes
- **Tables**: Markdown tables with clean light borders
- **Code blocks**: Monospace font, line-by-line rendering
- **Blockquotes**: Complete preservation with inline formatting
- **Lists**: Bullet and numbered lists with proper indentation
- **Inline formatting**: Bold, italic, code, links preserved
- **Chinese fonts**: All East-Asian fonts explicitly set (no MS Mincho fallback)
- **Paragraph breaks**: Newlines in Markdown become paragraph marks (^p), not soft breaks

## Project Structure

```
src/md2word/
├── cli.py         — CLI: argument parsing, 4 theme templates
├── converter.py   — Core: MD → HTML → docx with style extraction
├── template.py    — Style extraction from guide paragraphs
├── image_utils.py — Image download, resize, embed
template/
├── 官方公文.docx
├── 学术论文.docx
├── 技术文档.docx
└── 自媒体排版.docx
```

## License

MIT
