# md2word — Markdown to Word Converter

## Technology
- Python 3.10+, python-docx, markdown, Pillow, requests
- Optional: Pygments (highlight), matplotlib (math), resvg (SVG), watchdog (watch)
- CLI entry via `md2word` command (registered system-wide)

## Structure
```
src/md2word/
├── __init__.py          — Version info (v1.2.0)
├── __main__.py          — python -m md2word entry
├── cli.py               — CLI: arg parsing, orchestration, config merge, watch mode
├── converter.py         — Core: MD → HTML → docx with template styles
├── config.py            — Config file support (md2word.yaml / pyproject.toml)
├── template.py          — Style extraction, validation, guide paragraph matching
├── themes.py            — Theme specs & template builder (4 themes)
├── footnotes.py         — Footnote extraction & Word native footnote insertion
├── image_utils.py       — Image download (URL/local), resize, SVG dimension parsing
├── syntax.py            — Code syntax highlighting via Pygments
├── math_omml.py         — LaTeX → Word OMML (native editable formulas)
├── math_renderer.py     — LaTeX → SVG rendering via matplotlib (fallback)
├── mermaid_renderer.py  — Mermaid → SVG via API or mmdc CLI
├── py.typed             — PEP 561 marker
template/
├── 官方公文.docx, 学术论文.docx, 技术文档.docx, 自媒体排版.docx
MD2WORD-TOOL.md          — Copy of global tool doc (sync with ~/.claude/tools/md2word-tool.md)
```

## Template Guide Paragraphs
Keywords in template docx define styles: 一级标题(h1), 二级标题(h2), 三级标题(h3), 正文(body), 首行缩进(body_indent), 图片(image), 图注(figcaption), 引用(quote), 代码(code), 无序列表(bullet_list), 有序列表(number_list), 目录标题(toc_title). See `template.py` `_STYLE_KEYWORDS` for full list.

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
md2word input.md --three-line-table                                       # academic table style
md2word input.md --watch                                                  # auto-convert on change
md2word input.md --config md2word.yaml                                    # explicit config file
md2word --check-deps                                                      # check optional deps
```

## Config File
Auto-detects `md2word.yaml` / `md2word.yml` / `md2word.json` or `[tool.md2word]` in pyproject.toml. CLI args override config. See `config.py` for details.

## Features Added in v1.2.0
- Config file: auto-load md2word.yaml, CLI takes priority
- Watch mode: `--watch` polls .md files and auto-converts
- Footnotes: `[^1]` syntax → Word native footnotes
- Three-line table: `--three-line-table` for academic 三线表
- Cross-reference: `[text](#heading)` → Word REF field bookmarks
- Header/footer: auto-inherited from template
- Formula numbering: SEQ field auto-numbering for block math
- Deps check: `--check-deps` verifies optional dependencies
- Better errors: template/file missing friendly messages

## Optional Dependencies
| Extra | Package | Feature |
|-------|---------|---------|
| highlight | pygments | Code syntax highlighting |
| math | matplotlib, latex2mathml | LaTeX formula rendering (OMML + SVG fallback) |
| svg | resvg | Native SVG embedding in docx |
| watch | watchdog | Watch mode file monitoring |

## Themes
| Name      | Label         | Description                      |
|-----------|---------------|----------------------------------|
| official  | 官方公文      | 仿宋正文+黑体标题+GB标准页边距     |
| academic  | 学术论文      | 全宋体系列+黑体标题居中+标准学术排版 |
| tech      | 技术文档      | 微软雅黑+深蓝色层级标题+紧凑排版    |
| media     | 自媒体排版    | 视觉系大字报+高行距+品牌橙色        |

## Tool Doc Sync
The global tool doc at `~/.claude/tools/md2word-tool.md` and the project copy `MD2WORD-TOOL.md` must be kept in sync. When adding CLI features, update both files.
