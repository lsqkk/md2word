# md2word — Markdown to Word Converter

## Technology
- Python 3.10+, python-docx, markdown, Pillow, requests
- Optional: Pygments (highlight), matplotlib (math), resvg (SVG)
- CLI entry via `md2word` command (registered system-wide)

## Structure
```
src/md2word/
├── cli.py              — CLI: arg parsing, 4 theme templates
├── converter.py        — Core: MD → HTML → docx with template-derived styles
├── template.py         — Style extraction, validation, guide paragraph matching
├── image_utils.py      — Image download (URL/local), resize, embed
├── syntax.py           — Code syntax highlighting via Pygments
├── math_renderer.py    — LaTeX → SVG rendering via matplotlib
├── mermaid_renderer.py — Mermaid → SVG via API or mmdc CLI
template/
├── 官方公文.docx
├── 学术论文.docx
├── 技术文档.docx
└── 自媒体排版.docx
```

## Template Guide Paragraphs
Keywords in template docx define styles: 一级标题(h1), 二级标题(h2), 三级标题(h3), 正文(body), 首行缩进(body_indent), 图片(image), 图注(figcaption), 引用(quote), 代码(code), 无序列表(bullet_list), 有序列表(number_list), 目录标题(toc_title).

## Commands
```bash
md2word input.md -o output.docx                                           # basic
md2word file1.md file2.md                                                 # batch
md2word input.md -t template/技术文档.docx -o out.docx                    # custom theme
md2word --create-template my-template.docx --theme media                  # generate themed template
md2word --validate-template my-template.docx                              # validate template
md2word --list-themes                                                     # show themes
md2word --list-styles -t template/学术论文.docx                           # show detected styles
md2word input.md --toc --toc-depth 1-3 --number-headings --page-break     # full features
md2word input.md --no-highlight --no-math --no-mermaid                    # disable optional features
```

## Optional Dependencies
| Extra | Package | Feature |
|-------|---------|---------|
| highlight | pygments | Code syntax highlighting |
| math | matplotlib | LaTeX formula rendering |
| svg | resvg | Native SVG embedding in docx |

## Themes
| Name      | Label         | Description                      |
|-----------|---------------|----------------------------------|
| official  | 官方公文      | 仿宋正文+黑体标题+GB标准页边距     |
| academic  | 学术论文      | 全宋体系列+黑体标题居中+标准学术排版 |
| tech      | 技术文档      | 微软雅黑+深蓝色层级标题+紧凑排版    |
| media     | 自媒体排版    | 视觉系大字报+高行距+品牌橙色        |
