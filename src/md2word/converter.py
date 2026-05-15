"""Core converter — markdown → Word document using template-derived styles."""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET

import markdown
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from .image_utils import get_image_dimensions, resolve_image
from .template import (
    ParagraphFormat,
    extract_template_styles,
    set_run_font,
    _STYLE_KEYWORDS,
)


# ── HTML parsing ────────────────────────────────────────────────────────────


def _prepare_html(html: str) -> str:
    """Convert bare \\n in text content to <br/>, preserving <pre> blocks."""
    import re

    # Protect <pre> blocks from conversion
    pres: list[str] = []
    def _save(m):
        pres.append(m.group(0))
        return f"\x00PRE{len(pres)-1}\x00"
    html = re.sub(r'<pre[^>]*>.*?</pre>', _save, html, flags=re.DOTALL)

    # Replace all \n with <br/>
    html = html.replace('\n', '<br/>')

    # Remove <br/> that ended up between block tags (was inter-element whitespace)
    html = re.sub(r'>\s*<br/>\s*<', '><', html)

    # Remove leading/trailing <br/>
    html = re.sub(r'^(<br/>\s*)+', '', html)
    html = re.sub(r'(<br/>\s*)+$', '', html)

    # Restore <pre> blocks
    for i, block in enumerate(pres):
        html = html.replace(f"\x00PRE{i}\x00", block)

    return html


def _html_to_blocks(html: str) -> list[ET.Element]:
    root = ET.fromstring(f"<root>{html}</root>")
    return list(root)


def _inline_text(elem: ET.Element) -> str:
    parts: list[str] = []
    if elem.text:
        parts.append(elem.text)
    for child in elem:
        parts.append(_inline_text(child))
        if child.tail:
            parts.append(child.tail)
    return "".join(parts)


# ── run/paragraph helpers ───────────────────────────────────────────────────


def _set_run_color(run, color: RGBColor) -> None:
    run.font.color.rgb = color


def _push_run(p, text: str, base_fmt: ParagraphFormat, **overrides) -> None:
    if not text:
        return
    run = p.add_run(text)
    base_fmt.apply_to_run(run)
    for k, v in overrides.items():
        if k == "color":
            _set_run_color(run, v)
        elif k == "name":
            set_run_font(run, v, v)
        elif k == "underline":
            run.font.underline = v
        else:
            setattr(run.font, k, v)


def _build_runs(p, elem: ET.Element, base_fmt: ParagraphFormat) -> None:
    """Populate a docx paragraph with runs reflecting inline HTML tags."""

    def _push_detect(text: str, **overrides):
        """Push text, detecting markdown checkboxes [x]/[ ]."""
        if not text:
            return
        if text.startswith("[x] ") or text.startswith("[X] "):
            _push_run(p, "☑ ", base_fmt, color=RGBColor(0x2E, 0x7D, 0x32))
            text = text[4:]
        elif text.startswith("[ ] "):
            _push_run(p, "☐ ", base_fmt, color=RGBColor(0x99, 0x99, 0x99))
            text = text[4:]
        _push_run(p, text, base_fmt, **overrides)

    if elem.text:
        _push_detect(elem.text)

    for child in elem:
        tag = child.tag
        if tag == "br":
            continue  # handled by caller via _split_by_br
        elif tag in ("strong", "b"):
            _push_run(p, child.text or "", base_fmt, bold=True)
        elif tag in ("em", "i"):
            _push_run(p, child.text or "", base_fmt, italic=True)
        elif tag == "code":
            _push_run(p, child.text or "", base_fmt, name="DengXian")
        elif tag == "a":
            txt = child.text or child.get("href", "")
            _push_run(p, txt, base_fmt, underline=True, color=RGBColor(0x00, 0x52, 0xCC))
        elif tag == "img":
            src = child.get("src", "")
            alt = child.get("alt", "")
            stream = resolve_image(src)
            if stream is not None:
                w, h = get_image_dimensions(stream, max_width_inches=4.0)
                run = p.add_run()
                run.add_picture(stream, width=w, height=h)
            else:
                _push_run(p, f"[Image: {alt}]", base_fmt, italic=True)
        else:
            _push_run(p, child.text or "", base_fmt)

        if child.tail:
            _push_detect(child.tail)


def _split_by_br(elem: ET.Element) -> list[str]:
    """Split element text by <br/> tags into segments."""
    segments: list[str] = []
    buf: list[str] = []

    def _flush():
        nonlocal buf
        text = "".join(buf).strip()
        if text:
            segments.append(text)
        buf = []

    if elem.text:
        buf.append(elem.text)
    for child in elem:
        if child.tag == "br":
            _flush()
        elif child.tag in ("strong", "b", "em", "i", "a", "code"):
            buf.append(_inline_text(child))
        if child.tail:
            buf.append(child.tail)
    _flush()
    return segments if segments else [""]


# ── oxml helpers ────────────────────────────────────────────────────────────


def _set_outline_level(p, level: int) -> None:
    pPr = p._p.get_or_add_pPr()
    for existing in pPr.findall(qn("w:outlineLvl")):
        pPr.remove(existing)
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), str(level - 1))
    pPr.append(el)


def _add_bottom_border(p, color: str = "CCCCCC", sz: str = "6") -> None:
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def _style_table(table) -> None:
    tbl = table._tbl
    tblPr = tbl.tblPr
    for existing in list(tblPr.findall(qn("w:tblBorders"))):
        tblPr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "999999")
        borders.append(el)
    tblPr.append(borders)


# ── block handlers ──────────────────────────────────────────────────────────


def _handle_heading(
    doc: Document, elem: ET.Element, level: int, styles: dict
) -> None:
    slot = styles.get(f"h{level}") or styles.get("body", ParagraphFormat())
    p = doc.add_paragraph()
    slot.apply_to_paragraph(p)
    _build_runs(p, elem, slot)
    _set_outline_level(p, level)


def _handle_paragraph(doc: Document, elem: ET.Element, styles: dict) -> None:
    """Normal paragraph — detects lone images and embedded <br>."""
    imgs = elem.findall("img")
    if len(imgs) == 1 and not elem.text and not any(
        c.tag != "img" for c in elem
    ):
        return _handle_image(doc, imgs[0], styles)

    fmt = styles.get("body", ParagraphFormat())

    # <br> → new paragraph
    brs = elem.findall(".//br")
    if brs:
        for part in _split_by_br(elem):
            p = doc.add_paragraph()
            fmt.apply_to_paragraph(p)
            _push_run(p, part, fmt)
        return

    p = doc.add_paragraph()
    fmt.apply_to_paragraph(p)
    _build_runs(p, elem, fmt)


def _handle_image(doc: Document, elem: ET.Element, styles: dict) -> None:
    """Image with alt-text caption below."""
    src = elem.get("src", "")
    alt = elem.get("alt", "")
    stream = resolve_image(src)
    if stream is None:
        p = doc.add_paragraph()
        _push_run(p, f"[Image: {alt}]", ParagraphFormat(), italic=True)
        return

    img_fmt = styles.get("image", ParagraphFormat())
    p = doc.add_paragraph()
    img_fmt.apply_to_paragraph(p)
    width, height = get_image_dimensions(stream, max_width_inches=5.5)
    run = p.add_run()
    run.add_picture(stream, width=width, height=height)

    if alt:
        cap_fmt = styles.get("figcaption", ParagraphFormat())
        if cap_fmt.font_name is None:
            p_cap = doc.add_paragraph()
            p_cap.paragraph_format.alignment = 1  # center
            r = p_cap.add_run(alt)
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            set_run_font(r, "SimSun", "SimSun")
        else:
            p_cap = doc.add_paragraph()
            cap_fmt.apply_to_paragraph(p_cap)
            r = p_cap.add_run(alt)
            cap_fmt.apply_to_run(r)


def _handle_blockquote(doc: Document, elem: ET.Element, styles: dict) -> None:
    """Blockquote — process child <p> elements (fixes truncation bug)."""
    fmt = styles.get("quote", styles.get("body", ParagraphFormat()))
    for child in elem:
        if child.tag == "p":
            p = doc.add_paragraph()
            fmt.apply_to_paragraph(p)
            _build_runs(p, child, fmt)
        else:
            p = doc.add_paragraph()
            fmt.apply_to_paragraph(p)
            _build_runs(p, child if child.tag else elem, fmt)


def _handle_code_block(doc: Document, elem: ET.Element, styles: dict) -> None:
    """Code block — each line is a paragraph."""
    fmt = styles.get("code", ParagraphFormat())
    text = _inline_text(elem)
    for line in text.split("\n"):
        p = doc.add_paragraph()
        fmt.apply_to_paragraph(p)
        run = p.add_run(line if line else " ")
        fmt.apply_to_run(run)
        set_run_font(run, fmt.font_name or "DengXian", fmt.east_asia_font or "DengXian")


def _handle_unordered_list(
    doc: Document, elem: ET.Element, styles: dict
) -> None:
    """Unordered list with bullet prefix."""
    fmt = styles.get("bullet_list", styles.get("body", ParagraphFormat()))
    for li in elem.findall("li"):
        p = doc.add_paragraph()
        fmt.apply_to_paragraph(p)
        if fmt.left_indent_emu is None:
            p.paragraph_format.left_indent = Inches(0.5)
        bullet = p.add_run("• ")
        fmt.apply_to_run(bullet)
        _build_runs(p, li, fmt)


def _handle_ordered_list(
    doc: Document, elem: ET.Element, styles: dict
) -> None:
    """Ordered list with number prefix."""
    fmt = styles.get("number_list", styles.get("body", ParagraphFormat()))
    for idx, li in enumerate(elem.findall("li"), 1):
        p = doc.add_paragraph()
        fmt.apply_to_paragraph(p)
        if fmt.left_indent_emu is None:
            p.paragraph_format.left_indent = Inches(0.5)
        _push_run(p, f"{idx}. ", fmt)
        _build_runs(p, li, fmt)


def _handle_table(doc: Document, elem: ET.Element, styles: dict) -> None:
    """Markdown table."""
    rows_elem = elem.findall(".//tr")
    if not rows_elem:
        return
    cols = max(
        len(tr.findall("th")) + len(tr.findall("td")) for tr in rows_elem
    )
    if cols == 0:
        return

    table = doc.add_table(rows=len(rows_elem), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    body_fmt = styles.get("body", ParagraphFormat())

    for ri, tr in enumerate(rows_elem):
        cells = tr.findall("th") + tr.findall("td")
        for ci, cell in enumerate(cells):
            if ci >= cols:
                break
            wc = table.rows[ri].cells[ci]
            wc.text = ""
            p = wc.paragraphs[0]
            if cell.tag == "th":
                body_fmt.apply_to_paragraph(p)
                _push_run(p, _inline_text(cell), body_fmt, bold=True)
            else:
                body_fmt.apply_to_paragraph(p)
                _build_runs(p, cell, body_fmt)
    _style_table(table)


def _handle_horizontal_rule(doc: Document, styles: dict) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    _add_bottom_border(p)


# ── guide-paragraph removal ──────────────────────────────────────────────


def _remove_guide_paragraphs(doc: Document) -> None:
    """Remove all guide paragraphs (style markers) from a template."""
    from .template import _STYLE_KEYWORDS

    to_remove = []
    for p in doc.paragraphs:
        text = p.text.strip().lower()
        if not text:
            continue
        for keyword in sorted(_STYLE_KEYWORDS, key=len, reverse=True):
            if keyword in text:
                to_remove.append(p._p)
                break
    for p_elem in to_remove:
        p_elem.getparent().remove(p_elem)


# ── main ────────────────────────────────────────────────────────────────────


def convert(
    markdown_text: str,
    template_path: str | Path,
    output_path: str | Path,
    *,
    image_max_width: float = 5.5,
) -> None:
    """Convert *markdown_text* to a Word document.

    Args:
        markdown_text: Raw markdown content.
        template_path: Path to the .docx template with guide paragraphs.
        output_path: Where to write the result.
        image_max_width: Maximum image width in inches.
    """
    styles = extract_template_styles(template_path)

    html = markdown.markdown(
        markdown_text,
        extensions=[
            "extra",
            "tables",
            "fenced_code",
            "codehilite",
            "sane_lists",
        ],
    )

    html = _prepare_html(html)
    blocks = _html_to_blocks(html)

    doc = Document(str(template_path))

    # Remove guide paragraphs AND empty paragraph at top
    _remove_guide_paragraphs(doc)
    if doc.paragraphs and not doc.paragraphs[0].text.strip():
        p = doc.paragraphs[0]._p
        p.getparent().remove(p)

    for block in blocks:
        tag = block.tag
        try:
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                _handle_heading(doc, block, int(tag[1]), styles)
            elif tag == "p":
                _handle_paragraph(doc, block, styles)
            elif tag == "img":
                _handle_image(doc, block, styles)
            elif tag == "blockquote":
                _handle_blockquote(doc, block, styles)
            elif tag == "pre":
                _handle_code_block(doc, block, styles)
            elif tag == "ul":
                _handle_unordered_list(doc, block, styles)
            elif tag == "ol":
                _handle_ordered_list(doc, block, styles)
            elif tag == "table":
                _handle_table(doc, block, styles)
            elif tag == "hr":
                _handle_horizontal_rule(doc, styles)
            else:
                _handle_paragraph(doc, block, styles)
        except Exception as e:
            print(f"  [WARN] Failed to process <{tag}> block: {e}")

    doc.save(str(output_path))
