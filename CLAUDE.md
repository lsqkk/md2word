# md2word — Markdown to Word Converter

## Technology
- Python 3.10+, python-docx, markdown, Pillow, requests, PyYAML
- Optional: Pygments (highlight), matplotlib (math), resvg (SVG), watchdog (watch)
- CLI entry via `md2word` command (registered system-wide)
- v1.9.0: 251 tests, all passing

## Structure
```
src/md2word/
├── __init__.py          — Version info (v1.9.0)
├── __main__.py          — python -m md2word entry
├── cli.py               — CLI: arg parsing, orchestration, config merge, watch mode
├── converter.py         — Orchestration: MD → HTML → docx pipeline (reduced)
├── handlers.py          — Block processors (_handle_*) + inline run builders
├── ooxml_helpers.py     — OOXML ops (bookmarks, fields, SVG, numbering, TOC)
├── metadata.py          — Post-processing, GB compliance, red-head, page numbers
├── context.py           — ConversionContext + ConversionReport with severity
├── frontmatter.py       — YAML frontmatter ← → docx properties
├── cache.py             — Incremental conversion cache (MD5 hash)
├── config.py            — Config file support (md2word.yaml / pyproject.toml)
├── template.py          — Style extraction, validation, guide paragraph matching
├── themes.py            — Theme specs & template builder (6 themes)
├── footnotes.py         — Footnote extraction & Word native footnote insertion
├── image_utils.py       — Image download (URL/local), resize, SVG dimension parsing
├── syntax.py            — Code syntax highlighting via Pygments
├── math_omml.py         — LaTeX → Word OMML (native editable formulas)
├── math_renderer.py     — LaTeX → SVG rendering via matplotlib (fallback)
├── mermaid_renderer.py  — Mermaid → SVG via API or mmdc CLI
├── options.py           — ConvertOptions dataclass (convert() configuration)
├── update_check.py      — GitHub version update check with local caching
├── py.typed             — PEP 561 marker
template/
├── 官方公文.docx, 学术论文.docx, 技术文档.docx, 自媒体排版.docx
MD2WORD-TOOL.md          — Copy of global tool doc (sync with ~/.claude/tools/md2word-tool.md)
```

## Template Guide Paragraphs
Keywords in template docx define styles: 一级标题(h1), 二级标题(h2), 三级标题(h3), 正文(body), 首行缩进(body_indent), 图片(image), 图注(figcaption), 引用(quote), 代码(code), 无序列表(bullet_list), 有序列表(number_list), 目录标题(toc_title), 摘要(abstract), 关键词(keywords), 参考文献(references). See `template.py` `_STYLE_KEYWORDS` for full list.

## Commands
```bash
md2word input.md -o output.docx                                           # basic
md2word file1.md file2.md                                                 # batch
md2word "docs/**/*.md"                                                    # glob input patterns
md2word file1.md file2.md --out-dir ./output                              # output directory
md2word input.md -t template/技术文档.docx -o out.docx                    # custom theme
md2word --create-template my-template.docx --theme media                  # generate themed template
md2word --validate-template my-template.docx                              # validate template
md2word --list-themes                                                     # show themes
md2word --list-styles -t template/学术论文.docx                           # show detected styles
md2word input.md --toc --toc-depth 1-3 --number-headings --page-break     # full features
md2word input.md --no-highlight --no-math --no-mermaid                    # disable optional features
md2word input.md --no-update-check                                       # disable version update check
md2word input.md --three-line-table                                       # academic table style
md2word input.md --watch                                                  # auto-convert on change
md2word input.md --config md2word.yaml                                    # explicit config file
md2word --check-deps                                                      # check optional deps
md2word 通知.md --redhead "XX市人民政府" --theme redhead                   # 红头文件
md2word 通知.md --redhead "XX市人民政府" --redhead-year 2026 --redhead-number 12  # 红头文件(自定义文号)
md2word input.md --theme official --gb-check                               # GB合规检查
md2word input.md --page-number "-- %d --"                                  # 页码格式
md2word file1.md file2.md --incremental                                    # 增量转换
md2word input.md --project-dir /path/to/project                            # 项目目录
```

## Config File
Auto-detects `.md2word/config.yaml` > `md2word.yaml` / `md2word.yml` / `md2word.json` or `[tool.md2word]` in pyproject.toml. Also auto-detects `.md2word/template.docx` for project-level templates. CLI args override config. See `config.py` for details.

## Features Added in v1.9.0
- **Version update check**: After successful conversion, silently checks GitHub for newer releases. Results cached locally for 24h to avoid API rate limits.
- **CLI**: `--no-update-check` disables the check. Config equivalent: `update-check: false` in `md2word.yaml`.
- **New module**: `update_check.py` — `check_for_update()` / `fetch_latest_version()` / `format_update_message()`
- **25 new tests**: version comparison, cache, format, edge cases

## Features Added in v1.8.0
- **ConvertOptions dataclass**: `convert()` now accepts a single `ConvertOptions` object instead of 14+ keyword arguments — cleaner API with documented defaults
- **Config system upgrade**: PyYAML is now a hard dependency — nested config (`style_map`, lists) supported in `md2word.yaml`. Unknown config keys emit warnings. Backward-compatible with flat configs via fallback parser.
- **Template custom XML markers**: Generated templates inject machine-readable slot identifiers via `<w:customXml>`. `_guess_slot()` reads markers first, falls back to text matching — eliminates false positives from substring matching.
- **Verbose mode**: `--verbose` now emits stage-by-stage progress during conversion (template, front matter, footnotes, math, blocks 1/N, etc.)
- **Error context**: Block processing errors include inline text preview (first 80 chars) — easier to locate the source of conversion issues
- **39 new tests**: 2450 total test lines, covering ConvertOptions, config validation, slot markers, verbose mode, error context

## Features Added in v1.6.0
- **Abstract/keywords rendering**: frontmatter `abstract`/`keywords` now inserts styled paragraphs using template style slots (摘要/关键词)
- **Bookmark conflict fix**: duplicate heading slugs get `-1`, `-2` suffixes — no more Word "bookmark name duplicate" errors
- **`_push_detect` refactored**: extracted `_detect_task_prefix()` helper eliminating duplicate logic
- **`build_runs`/`build_runs_skip` unified**: `_build_inline_runs_core()` shared core with configurable flags
- **Code block strip fix**: `text.strip()` → `text.strip('\n')` preserves intentional whitespace in code
- **`converter.py` cleaned**: removed `_strip_front_matter()`, moved `ensure_list_blank_lines` to handlers.py, created `cache.py` module
- **`style_map` implemented**: `style_map={"code": "CustomStyle"}` overrides default style slots
- **Red-head number configurable**: `--redhead-year YEAR` and `--redhead-number NUM` for document number
- **Template keyword expansion**: added 摘要→abstract, 关键词→keywords, 参考文献→references to `_STYLE_KEYWORDS`
- **61 new tests**: metadata.py, ooxml_helpers.py coverage, v1.6 feature integration tests

## Features Added in v1.5.0
- **Architecture refactor**: `converter.py` split into 6 modules (handlers, ooxml_helpers, metadata, context, frontmatter)
- **ConversionContext**: replaces global state (_BOOKMARK_COUNTER, _HEADING_COUNTERS) — batch-safe
- **Severity levels**: ConversionReport now supports info/warning/error/critical
- **Strikethrough**: `~~text~~` → ~~text~~ in Word
- **Highlight**: `==text==` → yellow highlight
- **Super/Subscript**: `^sup^` / `~sub~` → superscript/subscript
- **Task lists**: `- [x]` / `- [ ]` supported in nested lists
- **YAML frontmatter**: title/author/date → docx core_properties; keywords/abstract support
- **Cross-reference fix**: `[text](#heading-slug)` uses slug as bookmark name — REF fields work
- **Out-dir**: `--out-dir DIR` for batch output to a directory
- **Glob patterns**: `md2word "docs/**/*.md"` expands wildcards

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

## Converter modules (v1.6.0+)
- `convert()` in converter.py — orchestration pipeline
- `handlers.py` — all `_handle_*` block processors + `_build_inline_runs_core()`
- `ooxml_helpers.py` — OOXML ops: bookmarks, REF/SEQ fields, SVG embed, numbering, TOC
- `metadata.py` — `fix_ooxml_metadata()` / `check_gb_compliance()` / `insert_redhead_header()` / `set_page_number_format()` / `remove_guide_paragraphs()`
- `context.py` — `ConversionContext` (replaces globals, now includes `used_bookmark_slugs`) + `ConversionReport` with severity
- `frontmatter.py` — YAML frontmatter parsing and docx property application
- `cache.py` — `file_hash()` / `load_cache()` / `save_cache()` — incremental conversion via MD5 hash
- `update_check.py` — `check_for_update()` / `fetch_latest_version()` / `format_update_message()` — version update check

## Known issues & fixes

### White thumbnail (blank preview) in Windows File Explorer

**Root cause**: python-docx embeds a built-in `docProps/thumbnail.jpeg` that is a blank/white image. When Windows finds a thumbnail inside the docx, it displays it instead of generating a preview from the document content.

**Fix applied in `fix_ooxml_metadata()`** (metadata.py):
1. Strip `docProps/thumbnail.jpeg` from the ZIP during post-processing
2. Remove the corresponding `<Relationship>` entry from `_rels/.rels` so the OPC stays valid
3. Also fix the Application name from "Microsoft Macintosh Word" to "Microsoft Office Word"

**If the white thumbnail reappears** — suspect that `fix_ooxml_metadata()` is not being called (check the `convert()` function's epilogue), or that python-docx changed the default thumbnail relationship path. Verify by extracting ZIP entries and checking for `thumbnail` in filenames.

## Tool Doc Sync
The global tool doc at `~/.claude/tools/md2word-tool.md` and the project copy `MD2WORD-TOOL.md` must be kept in sync. When adding CLI features, update both files.
