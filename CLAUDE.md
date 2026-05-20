# md2word — Markdown to Word Converter

## Technology
- Python 3.10+, python-docx, markdown, Pillow, requests
- Optional: Pygments (highlight), matplotlib (math), resvg (SVG), watchdog (watch)
- CLI entry via `md2word` command (registered system-wide)
- v1.3.0: 107 tests, all passing

## Structure
```
src/md2word/
├── __init__.py          — Version info (v1.3.0)
├── __main__.py          — python -m md2word entry
├── cli.py               — CLI: arg parsing, orchestration, config merge, watch mode
├── converter.py         — Core: MD → HTML → docx with template styles
├── config.py            — Config file support (md2word.yaml / pyproject.toml)
├── template.py          — Style extraction, validation, guide paragraph matching
├── themes.py            — Theme specs & template builder (6 themes: official, academic, academic-plus, tech, media, redhead)
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
md2word 通知.md --redhead "XX市人民政府" --theme redhead                   # 红头文件
md2word input.md --theme official --gb-check                               # GB合规检查
md2word input.md --page-number "-- %d --"                                  # 页码格式
md2word file1.md file2.md --incremental                                    # 增量转换
md2word input.md --project-dir /path/to/project                            # 项目目录
```

## Config File
Auto-detects `.md2word/config.yaml` > `md2word.yaml` / `md2word.yml` / `md2word.json` or `[tool.md2word]` in pyproject.toml. Also auto-detects `.md2word/template.docx` for project-level templates. CLI args override config. See `config.py` for details.

## Features Added in v1.3.0
- Red-head document: `--redhead AUTHORITY` generates 红头文件 with red authority name + separator + document number
- GB compliance check: `--gb-check` validates margins/fonts against GB/T 9704-2012
- Page number: `--page-number FMT` inserts custom page number format in footer
- Incremental conversion: `--incremental` uses MD5 hash cache to skip unchanged files
- Project-level config: `.md2word/config.yaml` + `.md2word/template.docx` auto-detection
- Project dir: `--project-dir DIR` explicitly sets project root
- New themes: academic-plus (enhanced academic with 摘要/关键词/参考文献) and redhead (红头文件专用)
- ConversionReport: structured return value from convert() with error/warning tracking
- 107 tests: comprehensive coverage across all modules

## Optional Dependencies
| Extra | Package | Feature |
|-------|---------|---------|
| highlight | pygments | Code syntax highlighting |
| math | matplotlib, latex2mathml | LaTeX formula rendering (OMML + SVG fallback) |
| svg | resvg | Native SVG embedding in docx |
| watch | watchdog | Watch mode file monitoring |

## Themes
| Name          | Label              | Description                              |
|---------------|--------------------|------------------------------------------|
| official      | 官方公文           | 仿宋正文+黑体标题+GB标准页边距            |
| academic      | 学术论文           | 全宋体系列+黑体标题居中+标准学术排版       |
| academic-plus | 学术论文（增强版）  | 增加摘要/关键词/参考文献引导段落           |
| tech          | 技术文档           | 微软雅黑+深蓝色层级标题+紧凑排版          |
| media         | 自媒体排版         | 视觉系大字报+高行距+品牌橙色              |
| redhead       | 红头文件           | GB/T 9704-2012 标准 + 红头样式            |

## Converter v1.3.0 features
- `_insert_redhead_header()` — injects red authority name + "文件" + red separator + doc number
- `_set_page_number_format()` — centered PAGE field in footer
- `_check_gb_compliance()` — validates margins against GB standards
- `_file_hash()` / `_load_cache()` / `_save_cache()` — incremental conversion via MD5 hash
- `ConversionReport` dataclass — structured error/warning/info tracking
- `.md2word/` directory support in config search

## Tool Doc Sync
The global tool doc at `~/.claude/tools/md2word-tool.md` and the project copy `MD2WORD-TOOL.md` must be kept in sync. When adding CLI features, update both files.
