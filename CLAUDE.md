# md2word — Markdown to Word Converter

## Technology
- Python 3.11+, python-docx, markdown, Pillow, requests
- CLI entry via `md2word` command (registered system-wide)

## Structure
```
src/md2word/
├── cli.py         — CLI: arg parsing, 4 theme templates, template generation
├── converter.py   — Core: MD → HTML → docx with template-derived styles
├── template.py    — Style extraction from guide paragraphs (incl. East-Asian fonts)
├── image_utils.py — Image download (URL/local), resize, embed
template/
├── 官方公文.docx
├── 学术论文.docx
├── 技术文档.docx
└── 自媒体排版.docx
```

## How It Works
The template docx contains **guide paragraphs** with marker keywords:
- "一级标题" → H1
- "二级标题" → H2
- "三级标题" → H3
- "正文" → Body
- "首行缩进" → Body with first-line indent
- "图片" → Centered image container
- "图注" → Image caption (alt text placed below image)
- "引用" → Blockquote
- "代码" → Code block (DengXian/Consolas monospace)
- "无序列表" → Bullet list
- "有序列表" → Numbered list

All fonts are explicitly set with both Western (`w:ascii`/`w:hAnsi`) and
East-Asian (`w:eastAsia`) font names to prevent fallback to Japanese fonts.

Newlines in markdown are automatically converted to paragraph breaks (^p).

## Commands
```bash
md2word input.md -o output.docx                             # default template
md2word input.md -t template/技术文档.docx -o out.docx       # custom theme
md2word --create-template my-template.docx --theme media     # generate themed template
md2word --list-themes                                        # show available themes
md2word --list-styles -t template/学术论文.docx               # show detected styles
```

## Themes
| Name      | Label         | Description                      |
|-----------|---------------|----------------------------------|
| official  | 官方公文      | 仿宋正文+黑体标题+GB标准页边距     |
| academic  | 学术论文      | 全宋体系列+黑体标题居中+标准学术排版 |
| tech      | 技术文档      | 微软雅黑+深蓝色层级标题+紧凑排版    |
| media     | 自媒体排版    | 视觉系大字报+高行距+品牌橙色        |
