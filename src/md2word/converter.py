"""Core converter — markdown → Word document using template-derived styles.

This module now serves as the orchestration layer.  Implementation details
have been split into focused sub-modules:

- ``handlers.py``: Block handlers and inline run builders
- ``ooxml_helpers.py``: OOXML-level operations (bookmarks, fields, SVG, etc.)
- ``metadata.py``: Post-processing, GB compliance, red-head, page numbers
- ``context.py``: ConversionContext, ConversionReport with severity levels
- ``frontmatter.py``: YAML front matter parsing and application
"""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path

import markdown
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .config import load_config as _load_config
from .context import ConversionContext, ConversionReport
from .footnotes import add_footnotes_to_document, extract_footnotes
from .frontmatter import apply_front_matter, parse_front_matter
from .handlers import (
    build_runs_skip,
    handle_blockquote,
    handle_code_block,
    handle_heading,
    handle_horizontal_rule,
    handle_image,
    handle_ordered_list,
    handle_paragraph,
    handle_table,
    handle_unordered_list,
    html_to_blocks,
    postprocess_math_html,
    preprocess_extended_syntax,
    prepare_html,
    replace_fn_placeholders,
)
from .math_renderer import extract_and_placeholder
from .metadata import (
    check_gb_compliance,
    fix_ooxml_metadata,
    insert_redhead_header,
    remove_guide_paragraphs,
    set_page_number_format,
)
from .template import extract_template_styles, _STYLE_KEYWORDS


# ── Regex ────────────────────────────────────────────────────────────

_FN_PLACEHOLDER_RE = re.compile(r"\x00FN_([^\x00]+)\x00")
_FN_TAG_RE = re.compile(r'<fn\s+id="([^"]+)"\s*/>')
_LIST_MARKER_RE = re.compile(r"^[\-\*\+] ")
_NUMBERED_MARKER_RE = re.compile(r"^\d+[\.\)] ")


# ── Markdown pre-processing ──────────────────────────────────────────


def _strip_front_matter(text: str) -> str:
    """Remove YAML front matter (---…---) from the beginning of markdown."""
    from .frontmatter import parse_front_matter
    _, body = parse_front_matter(text)
    return body


def _ensure_list_blank_lines(text: str) -> str:
    """Insert blank lines before list items when missing."""
    lines = text.split("\n")
    result: list[str] = []
    for i, line in enumerate(lines):
        stripped = line.strip()
        is_list = bool(
            _LIST_MARKER_RE.match(line) or _NUMBERED_MARKER_RE.match(line)
        )
        if is_list and i > 0:
            prev = lines[i - 1].strip()
            if (
                prev
                and not _LIST_MARKER_RE.match(lines[i - 1])
                and not _NUMBERED_MARKER_RE.match(lines[i - 1])
                and not prev.startswith("#")
            ):
                if result and result[-1] != "":
                    result.append("")
        result.append(line)
    return "\n".join(result)


# ── Incremental cache ────────────────────────────────────────────────


def _file_hash(path: str | Path) -> str:
    """Return MD5 hex digest of file content."""
    h = hashlib.md5()
    h.update(Path(path).read_bytes())
    return h.hexdigest()


def _load_cache(cache_path: str | Path) -> dict:
    """Load incremental conversion cache."""
    try:
        return json.loads(Path(cache_path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_cache(cache_path: str | Path, cache: dict) -> None:
    """Save incremental conversion cache."""
    Path(cache_path).write_text(
        json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── Main convert function ────────────────────────────────────────────


def convert(
    markdown_text: str,
    template_path: str | Path,
    output_path: str | Path,
    *,
    image_max_width: float = 5.5,
    toc: bool = True,
    toc_depth: str = "1-3",
    highlight_enabled: bool = True,
    math_enabled: bool = True,
    mermaid_enabled: bool = True,
    number_headings: bool = False,
    page_break_h1: bool = False,
    three_line_table: bool = False,
    footnotes_enabled: bool = True,
    formula_numbering: bool = False,
    redhead_authority: str | None = None,
    page_number_fmt: str | None = None,
    gb_check: bool = False,
    style_map: dict[str, str] | None = None,
) -> ConversionReport:
    """Convert *markdown_text* to a Word document.

    Args:
        markdown_text: Raw markdown content.
        template_path: Path to the .docx template with guide paragraphs.
        output_path: Where to write the result.
        image_max_width: Maximum image width in inches (default 5.5).
        toc: Whether to generate a Table of Contents.
        toc_depth: Heading levels to include, e.g. "1-3".
        highlight_enabled: Enable code syntax highlighting.
        math_enabled: Enable math formula rendering ($...$ / $$...$$).
        mermaid_enabled: Enable mermaid diagram rendering.
        number_headings: Add auto-numbering to headings.
        page_break_h1: Add page break before each H1 heading.
        three_line_table: Use academic three-line table style.
        footnotes_enabled: Process footnote syntax [^id].
        formula_numbering: Add SEQ equation numbers to block math.
        redhead_authority: Issuing authority name for red-head document.
        page_number_fmt: Page number format, e.g. "-- %d --".
        gb_check: Check formatting against GB standards.
        style_map: Optional mapping of element type → Word style name.
    """
    ctx = ConversionContext()
    ctx.report = ConversionReport()
    ctx.styles = extract_template_styles(template_path)
    ctx.styles["_report"] = ctx.report

    # ── Parse YAML front matter ────────────────────────────────────────
    front_matter, markdown_text = parse_front_matter(markdown_text)

    # ── Pre-process extended syntax (~~strikethrough~~, ==highlight==, etc.) ─
    markdown_text = preprocess_extended_syntax(markdown_text)

    # ── Footnote preprocessing ──────────────────────────────────────────
    fn_list: list = []
    if footnotes_enabled:
        markdown_text, fn_list = extract_footnotes(markdown_text)

    # ── Math preprocessing: extract LaTeX expressions ────────────────────
    if math_enabled:
        markdown_text, math_exprs = extract_and_placeholder(markdown_text)
    else:
        math_exprs = []

    # ── Ensure list items have preceding blank lines ─────────────────
    markdown_text = _ensure_list_blank_lines(markdown_text)

    # ── Markdown → HTML ────────────────────────────────────────────────
    html = markdown.markdown(
        markdown_text,
        extensions=[
            "extra",
            "tables",
            "fenced_code",
            "sane_lists",
        ],
    )

    html = prepare_html(html)

    # ── Math postprocessing: render LaTeX to OMML/SVG, embed in HTML ─────
    if math_exprs:
        html = postprocess_math_html(html, math_exprs)

    # ── Replace footnote placeholders with safe HTML tags ──────────
    if fn_list:
        html = replace_fn_placeholders(html)

    blocks = html_to_blocks(html)

    doc = Document(str(template_path))

    # Apply front matter to docx properties
    apply_front_matter(doc, front_matter)

    # Ensure first section doesn't force a new page
    if doc.sections:
        doc.sections[0].start_type = 0

    # Remove guide paragraphs AND empty paragraph at top
    remove_guide_paragraphs(doc)
    if doc.paragraphs and not doc.paragraphs[0].text.strip():
        p = doc.paragraphs[0]._p
        p.getparent().remove(p)

    # ── Red-head header (insert at very beginning) ──────────────────────
    if redhead_authority:
        insert_redhead_header(doc, redhead_authority, ctx.styles)

    # ── Page number format ─────────────────────────────────────────────
    if page_number_fmt:
        set_page_number_format(doc, page_number_fmt)

    # ── GB standards compliance check ──────────────────────────────────
    if gb_check:
        check_gb_compliance(doc, ctx.styles, ctx.report)

    # ── Table of Contents ──────────────────────────────────────────────
    from .ooxml_helpers import add_toc as _add_toc

    _toc_pending = toc
    check_blocks = list(blocks)
    if _toc_pending and not any(b.tag == "h1" for b in check_blocks):
        _add_toc(doc, ctx.styles, depth=toc_depth)
        _toc_pending = False

    # Set Word to update fields (e.g. TOC) on open
    settings_el = doc.settings._element
    up = settings_el.find(qn("w:updateFields"))
    if up is None:
        up = OxmlElement("w:updateFields")
        up.set(qn("w:val"), "true")
        settings_el.append(up)
    else:
        up.set(qn("w:val"), "true")

    # ── Heading numbering ───────────────────────────────────────────────
    if number_headings:
        ctx.reset_heading_counters()

    # ── Add footnotes to doc ────────────────────────────────────────────
    if fn_list:
        ctx.fn_map = add_footnotes_to_document(doc, fn_list)

    # ── Process blocks ──────────────────────────────────────────────────
    table_counter = [0]

    for block in blocks:
        tag = block.tag
        try:
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                handle_heading(doc, block, int(tag[1]), ctx.styles, ctx,
                               number_headings=number_headings,
                               page_break=page_break_h1)
                if _toc_pending and tag == "h1":
                    _add_toc(doc, ctx.styles, depth=toc_depth)
                    _toc_pending = False
            elif tag == "p":
                handle_paragraph(doc, block, ctx.styles, ctx,
                                 image_max_width=image_max_width)
            elif tag == "img":
                handle_image(doc, block, ctx.styles, ctx,
                             image_max_width=image_max_width)
                if formula_numbering and block.get("class", "").startswith("math-block"):
                    from .ooxml_helpers import add_seq_number as _add_seq_number
                    p_num = doc.add_paragraph()
                    p_num.paragraph_format.alignment = 3
                    _add_seq_number(p_num, "Equation")
            elif tag == "blockquote":
                handle_blockquote(doc, block, ctx.styles, ctx)
            elif tag == "pre":
                handle_code_block(doc, block, ctx.styles, ctx,
                                  highlight_enabled=highlight_enabled,
                                  mermaid_enabled=mermaid_enabled)
            elif tag == "ul":
                handle_unordered_list(doc, block, ctx.styles, ctx)
            elif tag == "ol":
                handle_ordered_list(doc, block, ctx.styles, ctx)
            elif tag == "table":
                handle_table(doc, block, ctx.styles, ctx,
                             three_line=three_line_table,
                             table_idx=table_counter[0])
                table_counter[0] += 1
            elif tag == "hr":
                handle_horizontal_rule(doc, ctx.styles)
            else:
                handle_paragraph(doc, block, ctx.styles, ctx,
                                 image_max_width=image_max_width)
        except Exception as e:
            ctx.report.warn(f"处理 <{tag}> 块时出错: {e}")

    doc.save(str(output_path))

    # ── Fix OOXML metadata for proper Windows preview ──────────────────
    fix_ooxml_metadata(output_path)

    ctx.report.info_msg("输出文件: " + str(output_path))
    print(f"  {ctx.report.summary()}")
    return ctx.report
