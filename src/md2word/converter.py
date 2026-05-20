"""Core converter — markdown → Word document using template-derived styles."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET


# ── GB Standards reference ─────────────────────────────────────────────────


_GB_STANDARDS: dict[str, dict] = {
    "9704_2012": {
        # GB/T 9704-2012 党政机关公文格式
        "margins_cm": (3.7, 3.5, 2.8, 2.6),  # 上 37mm, 下 35mm, 左 28mm, 右 26mm
        "body_font": "FangSong",
        "body_size_pt": 16,  # 三号
        "heading1_font": "SimHei",
        "heading1_size_pt": 22,  # 二号
        "line_spacing": 1.5,  # 一倍半行距
    },
    "7713_2015": {
        # GB/T 7713-2015 学术论文格式
        "margins_cm": (2.54, 2.54, 3.17, 3.17),
        "body_font": "SimSun",
        "body_size_pt": 12,  # 小四号
        "heading1_font": "SimHei",
        "heading1_size_pt": 16,  # 三号
        "heading_align": 1,  # 居中
        "line_spacing": 1.5,
    },
}

_GB_CHECKLIST: dict[str, list[str]] = {
    "margins": [
        ('GB/T 9704-2012 公文上边距应为 37mm，当前 %.1fmm', _GB_STANDARDS["9704_2012"]["margins_cm"][0]),
        ('GB/T 9704-2012 公文下边距应为 35mm，当前 %.1fmm', _GB_STANDARDS["9704_2012"]["margins_cm"][1]),
        ('GB/T 9704-2012 公文左边距应为 28mm，当前 %.1fmm', _GB_STANDARDS["9704_2012"]["margins_cm"][2]),
        ('GB/T 9704-2012 公文右边距应为 26mm，当前 %.1fmm', _GB_STANDARDS["9704_2012"]["margins_cm"][3]),
    ],
}


# ── Incremental cache ──────────────────────────────────────────────────────


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
    Path(cache_path).write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


# ── Conversion report ───────────────────────────────────────────────────────


@dataclass
class ConversionReport:
    """Tracks warnings and errors during conversion.

    Provides structured output that can be queried programmatically.
    """
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    info: list[str] = field(default_factory=list)

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)
        print(f"  ⚠  {msg}")

    def error(self, msg: str) -> None:
        self.errors.append(msg)
        print(f"  ✖  {msg}")

    def info_msg(self, msg: str) -> None:
        self.info.append(msg)

    def summary(self) -> str:
        """Return a structured human-readable summary."""
        lines = []
        if self.info:
            for msg in self.info:
                lines.append(f"  ℹ  {msg}")
        if self.warnings:
            lines.append(f"  ⚠  {len(self.warnings)} 个警告")
        if self.errors:
            lines.append(f"  ✖  {len(self.errors)} 个错误")
        if not self.errors and not self.warnings:
            lines.append("  ✅ 转换完成，无警告或错误")
        else:
            lines.append("  📋 请查看以上详细信息")
        return "\n".join(lines)

    def has_errors(self) -> bool:
        return len(self.errors) > 0

import markdown
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from .image_utils import (
    get_image_dimensions,
    get_svg_size,
    is_svg_source,
    resolve_image,
    resolve_svg_raw,
)
from .footnotes import (
    add_footnotes_to_document,
    create_footnote_reference_run,
    extract_footnotes,
)
from .math_omml import latex_to_omml as _latex_to_omml
from .math_renderer import extract_and_placeholder, has_math, render_math_svg
from .mermaid_renderer import render_mermaid
from .syntax import extract_language, highlight
from .template import (
    ParagraphFormat,
    extract_template_styles,
    set_run_font,
    _STYLE_KEYWORDS,
)

# Regex for footnote placeholder \x00FN_id\x00
import re
_FN_PLACEHOLDER_RE = re.compile(r"\x00FN_([^\x00]+)\x00")

# XML-safe footnote tag
_FN_TAG_RE = re.compile(r'<fn\s+id="([^"]+)"\s*/>')


def _replace_fn_placeholders(html: str) -> str:
    """Replace \\x00FN_id\\x00 → XML-safe <fn id="id"/> tags."""
    return _FN_PLACEHOLDER_RE.sub(r'<fn id="\1"/>', html)


# ── Front matter ───────────────────────────────────────────────────────────────


_LIST_MARKER_RE = re.compile(r"^[\-\*\+] ")
_NUMBERED_MARKER_RE = re.compile(r"^\d+[\.\)] ")


def _ensure_list_blank_lines(text: str) -> str:
    """Insert blank lines before list items when missing.

    Standard Markdown requires a blank line before ``- ``, ``* ``, ``1. ``
    etc. for them to be parsed as lists.  This pre-processor inserts the
    missing blank lines so that ``-`` at the start of a line always becomes
    a list item.
    """
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


def _strip_front_matter(text: str) -> str:
    """Remove YAML front matter (---…---) from the beginning of markdown."""
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            return text[end + 3 :].lstrip()
    return text


def _postprocess_math_html(html: str, math_exprs: list[dict]) -> str:
    """Render math LaTeX to OMML (preferred) or SVG, replacing placeholders."""
    import base64

    result = html
    for expr in math_exprs:
        latex = expr["latex"]
        kind = expr["kind"]
        alt = f"[{latex[:60]}]"

        # Try OMML first — yields editable native Word formulas
        omml = _latex_to_omml(latex, kind == "block")
        if omml:
            omml_b64 = base64.b64encode(omml.encode("utf-8")).decode("ascii")
            replacement = (
                f'<img class="math-{kind}" src="" '
                f'data-omml-base64="{omml_b64}" alt="{alt}"/>'
            )
        else:
            # Fall back to SVG rendering
            svg_bytes, w_px, h_px = render_math_svg(latex, kind)
            if svg_bytes:
                svg_b64 = base64.b64encode(svg_bytes).decode("ascii")
                replacement = (
                    f'<img class="math-{kind}" src="" '
                    f'data-svg-base64="{svg_b64}" alt="{alt}"/>'
                )
            else:
                replacement = f'<span class="math-fallback">{alt}</span>'
        result = result.replace(expr["placeholder"], replacement)
    return result


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


# ── globals for cross-ref bookmarks ─────────────────────────────────────────

_BOOKMARK_COUNTER: int = 0


def _next_bookmark_id() -> int:
    global _BOOKMARK_COUNTER
    _BOOKMARK_COUNTER += 1
    return _BOOKMARK_COUNTER


def _slugify(text: str) -> str:
    """Convert text to a bookmark-friendly ASCII slug."""
    text = text.strip().lower()
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"[\s_]+", "-", text)
    text = re.sub(r"-+", "-", text)
    return text.strip("-") or "ref"


def _add_bookmark(paragraph, name: str) -> str:
    """Add a bookmark (start + end) around a paragraph.

    Returns the unique bookmark name (may be suffixed if name conflicts).
    """
    bid = _next_bookmark_id()
    bm_name = f"_Ref{bid}"
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), bm_name)
    paragraph._p.insert(0, start)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    paragraph._p.append(end)
    return bm_name


def _make_ref_field(display_text: str, bookmark_name: str) -> list:
    """Create OMML runs for a REF field pointing to a bookmark.

    Returns a list of OxmlElement runs.
    """
    runs = []
    # begin
    r1 = OxmlElement("w:r")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r1.append(fld_begin)
    runs.append(r1)
    # instrText
    r2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' REF {bookmark_name} \\h '
    r2.append(instr)
    runs.append(r2)
    # separate
    r3 = OxmlElement("w:r")
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r3.append(fld_sep)
    runs.append(r3)
    # display text (what Word shows when fields are not updated)
    r4 = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = display_text
    r4.append(t)
    runs.append(r4)
    # end
    r5 = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r5.append(fld_end)
    runs.append(r5)
    return runs


def _add_seq_number(p, seq_name: str = "Equation") -> None:
    """Add a SEQ field to a paragraph for auto-numbering."""
    # begin
    r1 = OxmlElement("w:r")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r1.append(fld_begin)
    p._p.append(r1)
    # instrText
    r2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" SEQ {seq_name} \\* ARABIC"
    r2.append(instr)
    p._p.append(r2)
    # separate
    r3 = OxmlElement("w:r")
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r3.append(fld_sep)
    p._p.append(r3)
    # display text
    r4 = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r4.append(t)
    p._p.append(r4)
    # end
    r5 = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r5.append(fld_end)
    p._p.append(r5)


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


def _push_text_with_footnotes(
    p, text: str, base_fmt: ParagraphFormat,
    fn_map: dict[str, int],
    **overrides,
) -> None:
    """Push *text*, splitting on footnote placeholders."""
    parts = _FN_PLACEHOLDER_RE.split(text)
    for i, part in enumerate(parts):
        if not part:
            continue
        if i % 2 == 0:
            # Regular text
            _push_run(p, part, base_fmt, **overrides)
        else:
            # Footnote reference
            fn_id = fn_map.get(part)
            if fn_id is not None:
                p._p.append(create_footnote_reference_run(None, fn_id))
            else:
                _push_run(p, f"[{part}]", base_fmt, **overrides)


def _build_runs(
    doc: Document,
    p,
    elem: ET.Element,
    base_fmt: ParagraphFormat,
    fn_map: dict[str, int] | None = None,
) -> None:
    """Populate a docx paragraph with runs reflecting inline HTML tags.

    If *fn_map* is provided, footnote placeholders are converted to
    Word footnote reference marks.
    """
    _fn_map = fn_map or {}

    def _push_detect(text: str, **overrides):
        """Push text, detecting checkboxes."""  # footnotes handled as <fn> tags
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
            href = child.get("href", "")
            txt = child.text or href
            if href.startswith("#"):
                # Cross-reference to a bookmark
                target = href[1:]
                for ref_run in _make_ref_field(txt, target):
                    p._p.append(ref_run)
            else:
                _push_run(p, txt, base_fmt, underline=True, color=RGBColor(0x00, 0x52, 0xCC))
        elif tag == "fn":
            fn_id_attr = child.get("id", "")
            if _fn_map and fn_id_attr in _fn_map:
                p._p.append(create_footnote_reference_run(None, _fn_map[fn_id_attr]))
        elif tag == "img":
            src = child.get("src", "")
            alt = child.get("alt", "")
            omml_b64 = child.get("data-omml-base64", "")
            if omml_b64:
                import base64
                omml_data = base64.b64decode(omml_b64).decode("utf-8")
                if omml_data:
                    _insert_inline_omml(p, omml_data)
                else:
                    _push_run(p, f"[Math: {alt}]", base_fmt, italic=True)
            elif svg_b64 := child.get("data-svg-base64", ""):
                import base64
                svg_data = base64.b64decode(svg_b64)
                if svg_data:
                    _embed_svg_in_paragraph(doc, p, svg_data, alt, 3.5)
                else:
                    _push_run(p, f"[Math: {alt}]", base_fmt, italic=True)
            elif is_svg_source(src):
                svg_data = resolve_svg_raw(src)
                if svg_data:
                    _embed_svg_in_paragraph(doc, p, svg_data, alt, 4.0)
                else:
                    _push_run(p, f"[SVG: {alt}]", base_fmt, italic=True)
            else:
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


def _split_by_br(elem: ET.Element) -> list[tuple[str, ET.Element | None]]:
    """Split element content by <br/> tags.

    Returns list of (text, img_element) pairs — *img_element* is ``None``
    for plain-text segments and the ``<img>`` element for image-only segments.
    """
    segments: list[tuple[str, ET.Element | None]] = []
    buf: list[str] = []

    def _flush():
        nonlocal buf
        text = "".join(buf).strip()
        if text:
            segments.append((text, None))
        buf = []

    if elem.text:
        buf.append(elem.text)
    for child in elem:
        if child.tag == "br":
            _flush()
        elif child.tag == "img":
            _flush()
            segments.append(("", child))
        else:
            buf.append(_inline_text(child))
        if child.tail:
            buf.append(child.tail)
    _flush()
    if not segments:
        segments = [("", None)]
    return segments


# ── Table of Contents ───────────────────────────────────────────────────────


def _add_toc(doc: Document, styles: dict, depth: str = "1-3") -> None:
    """Insert a TOC field at the beginning of the document."""
    # Title
    toc_fmt = styles.get("toc_title", styles.get("h1", ParagraphFormat()))
    p_title = doc.add_paragraph()
    toc_fmt.apply_to_paragraph(p_title)
    r = p_title.add_run("目录")
    toc_fmt.apply_to_run(r)

    # TOC field
    p = doc.add_paragraph()

    def _mk_r(*children) -> OxmlElement:
        r = OxmlElement("w:r")
        for c in children:
            r.append(c)
        return r

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    p._p.append(_mk_r(begin))

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' TOC \\o "{depth}" \\h \\z \\u '
    p._p.append(_mk_r(instr))

    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    p._p.append(_mk_r(separate))

    p_ph = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "（更新域以生成目录，Ctrl+A → F9）"
    p_ph.append(t)
    p._p.append(p_ph)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    p._p.append(_mk_r(end))


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


# ── SVG embedder ──────────────────────────────────────────────────────────────


def _embed_svg_in_paragraph(
    doc: Document, p, svg_data: bytes, alt: str = "",
    max_width_inches: float = 5.5,
) -> None:
    """Embed SVG directly in *p* via OPC-level ``asvg:svgBlip``."""
    from docx.oxml import parse_xml
    from docx.package import ImagePart

    size = get_svg_size(svg_data)
    if size is None:
        _push_run(p, f"[SVG: {alt}]", ParagraphFormat(), italic=True)
        return
    w_px, h_px = size

    max_px = max_width_inches * 96
    if w_px > max_px:
        scale = max_px / w_px
        w_px = int(w_px * scale)
        h_px = int(h_px * scale)

    emu_w = w_px * 914400 // 96
    emu_h = h_px * 914400 // 96

    partname = doc.part.package.next_partname("/word/media/image%d.svg")
    svg_part = ImagePart(partname, "image/svg+xml", svg_data)
    svg_part._package = doc.part.package

    rId = doc.part.relate_to(
        svg_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    )

    xml = (
        f'<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f'  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        f'  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        f'  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        f'  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
        f'  xmlns:asvg="http://schemas.microsoft.com/office/drawing/2020/SVG">'
        f'  <wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'    <wp:extent cx="{emu_w}" cy="{emu_h}"/>'
        f'    <wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'    <wp:docPr id="0" name="SVG" descr="{alt}"/>'
        f'    <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'    <a:graphic>'
        f'      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'        <pic:pic>'
        f'          <pic:nvPicPr>'
        f'            <pic:cNvPr id="0" name="SVG"/>'
        f'            <pic:cNvPicPr/>'
        f'          </pic:nvPicPr>'
        f'          <pic:blipFill>'
        f'            <a:blip r:embed="{rId}">'
        f'              <a:extLst>'
        f'                <a:ext uri="{{96DAC541-7B7A-43D3-8B79-5D63384654D6}}">'
        f'                  <asvg:svgBlip r:embed="{rId}"/>'
        f'                </a:ext>'
        f'              </a:extLst>'
        f'            </a:blip>'
        f'            <a:srcRect/>'
        f'            <a:stretch><a:fillRect/></a:stretch>'
        f'          </pic:blipFill>'
        f'          <pic:spPr>'
        f'            <a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu_w}" cy="{emu_h}"/></a:xfrm>'
        f'            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'          </pic:spPr>'
        f'        </pic:pic>'
        f'      </a:graphicData>'
        f'    </a:graphic>'
        f'  </wp:inline>'
        f'</w:drawing>'
    )

    run = p.add_run()
    run._r.append(parse_xml(xml))


# ── OMML helpers ──────────────────────────────────────────────────────────


def _insert_inline_omml(p, omml_xml: str) -> None:
    """Insert an inline OMML formula inside a run in *p*."""
    from docx.oxml import parse_xml

    run = p.add_run()
    run._r.append(parse_xml(omml_xml))


def _insert_block_omml(doc: Document, omml_xml: str) -> None:
    """Insert a display OMML formula as a centered paragraph."""
    from docx.oxml import parse_xml

    p = doc.add_paragraph()
    p.paragraph_format.alignment = 1  # center
    p._p.append(parse_xml(omml_xml))


def _style_table(table, three_line: bool = False) -> None:
    tbl = table._tbl
    tblPr = tbl.tblPr
    for existing in list(tblPr.findall(qn("w:tblBorders"))):
        tblPr.remove(existing)
    borders = OxmlElement("w:tblBorders")

    if three_line:
        # Academic three-line table: thick top, medium header-bottom, thick bottom
        for edge, sz, color in [
            ("top", "12", "000000"),
            ("bottom", "12", "000000"),
            ("insideH", "6", "000000"),
        ]:
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), sz)
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
            borders.append(el)
        # No vertical or side borders
        for edge in ("left", "right", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "none")
            borders.append(el)
    else:
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "999999")
            borders.append(el)
    tblPr.append(borders)


# ── List numbering helpers ────────────────────────────────────────────────


def _ensure_list_abstract_nums(doc: Document) -> dict[str, int]:
    """Create abstractNum definitions for bullet and ordered lists if missing.

    Returns dict ``{"bullet": id, "ordered": id}``.
    """
    key = "_md2word_list_abstracts"
    if hasattr(doc, key):
        return getattr(doc, key)

    from docx.oxml import parse_xml

    numbering_part = doc.part.numbering_part
    numbering_xml = numbering_part._element

    # Next available abstractNumId
    existing_abstract = numbering_xml.findall(qn("w:abstractNum"))
    ids = [
        int(a.get(qn("w:abstractNumId")))
        for a in existing_abstract
        if a.get(qn("w:abstractNumId")) is not None
    ]
    next_id = max(ids) + 1 if ids else 0

    bullet_abstract = (
        f'<w:abstractNum w:abstractNumId="{next_id}"'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'  <w:multiLevelType w:val="hybridMultilevel"/>'
        f'  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'    <w:lvlText w:val="●"/><w:lvlJc w:val="left"/>'
        f'    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        f'  </w:lvl>'
        f'  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'    <w:lvlText w:val="○"/><w:lvlJc w:val="left"/>'
        f'    <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
        f'  </w:lvl>'
        f'  <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'    <w:lvlText w:val="▪"/><w:lvlJc w:val="left"/>'
        f'    <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>'
        f'  </w:lvl>'
        f'  <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'    <w:lvlText w:val="◆"/><w:lvlJc w:val="left"/>'
        f'    <w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr>'
        f'  </w:lvl>'
        f'</w:abstractNum>'
    )
    numbering_xml.append(parse_xml(bullet_abstract))
    bullet_id = next_id
    next_id += 1

    ordered_abstract = (
        f'<w:abstractNum w:abstractNumId="{next_id}"'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'  <w:multiLevelType w:val="hybridMultilevel"/>'
        f'  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        f'    <w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
        f'    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        f'  </w:lvl>'
        f'  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>'
        f'    <w:lvlText w:val="%2."/><w:lvlJc w:val="left"/>'
        f'    <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
        f'  </w:lvl>'
        f'  <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/>'
        f'    <w:lvlText w:val="%3."/><w:lvlJc w:val="left"/>'
        f'    <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>'
        f'  </w:lvl>'
        f'  <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        f'    <w:lvlText w:val="%4."/><w:lvlJc w:val="left"/>'
        f'    <w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr>'
        f'  </w:lvl>'
        f'</w:abstractNum>'
    )
    numbering_xml.append(parse_xml(ordered_abstract))
    ordered_id = next_id

    result = {"bullet": bullet_id, "ordered": ordered_id}
    setattr(doc, key, result)
    return result


def _create_list_num_id(doc: Document, list_type: str = "bullet") -> int:
    """Create a new ``<w:num>`` instance and return its numId."""
    from docx.oxml import parse_xml

    abstracts = _ensure_list_abstract_nums(doc)
    abstract_id = abstracts[list_type]

    numbering_xml = doc.part.numbering_part._element
    existing_nums = numbering_xml.findall(qn("w:num"))
    num_ids = [
        int(n.get(qn("w:numId")))
        for n in existing_nums
        if n.get(qn("w:numId")) is not None
    ]
    next_num_id = max(num_ids) + 1 if num_ids else 1

    num_xml = (
        f'<w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' w:numId="{next_num_id}">'
        f'  <w:abstractNumId w:val="{abstract_id}"/>'
        f'</w:num>'
    )
    numbering_xml.append(parse_xml(num_xml))
    return next_num_id


def _add_num_pr(p, num_id: int, ilvl: int = 0) -> None:
    """Add ``<w:numPr>`` to a paragraph, replacing any existing one."""
    pPr = p._p.get_or_add_pPr()
    for existing in pPr.findall(qn("w:numPr")):
        pPr.remove(existing)
    numPr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl")
    ilvl_el.set(qn("w:val"), str(ilvl))
    num_id_el = OxmlElement("w:numId")
    num_id_el.set(qn("w:val"), str(num_id))
    numPr.append(ilvl_el)
    numPr.append(num_id_el)
    pPr.insert(0, numPr)


# ── block handlers ──────────────────────────────────────────────────────────


_HEADING_COUNTERS: list[int] = []


def _reset_heading_counters(max_level: int = 6) -> None:
    global _HEADING_COUNTERS
    _HEADING_COUNTERS = [0] * (max_level + 1)


def _next_heading_number(level: int) -> str:
    """Return heading number with offset: h2 → ``1``, h3 → ``1.1``, etc.

    ``h1`` (#) is treated as a document title and never numbered.
    """
    global _HEADING_COUNTERS
    if level <= 1:
        return ""
    # Shift: h2 → display level 1, h3 → level 2, etc.
    dl = level - 1
    _HEADING_COUNTERS[dl] += 1
    for i in range(dl + 1, len(_HEADING_COUNTERS)):
        _HEADING_COUNTERS[i] = 0
    return ".".join(str(_HEADING_COUNTERS[l]) for l in range(1, dl + 1))


def _add_page_break(doc: Document) -> None:
    """Add a page break paragraph."""
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)


def _handle_heading(
    doc: Document, elem: ET.Element, level: int, styles: dict,
    number_headings: bool = False,
    page_break: bool = False,
    fn_map: dict[str, int] | None = None,
) -> None:
    slot = styles.get(f"h{level}") or styles.get("body", ParagraphFormat())
    if page_break and level == 1:
        _add_page_break(doc)
    p = doc.add_paragraph()
    slot.apply_to_paragraph(p)
    if number_headings:
        num = _next_heading_number(level)
        if num:
            _push_run(p, f"{num} ", slot)
    _build_runs(doc, p, elem, slot, fn_map=fn_map)
    _set_outline_level(p, level)
    # Add bookmark from heading text for cross-references
    heading_text = _inline_text(elem)
    if heading_text:
        _add_bookmark(p, _slugify(heading_text))


def _handle_paragraph(
    doc: Document, elem: ET.Element, styles: dict,
    image_max_width: float = 5.5,
    fn_map: dict[str, int] | None = None,
) -> None:
    """Normal paragraph — detects lone images and embedded <br>."""
    imgs = elem.findall("img")
    if len(imgs) == 1 and not elem.text and not any(
        c.tag != "img" for c in elem
    ):
        return _handle_image(doc, imgs[0], styles, image_max_width=image_max_width)

    fmt = styles.get("body", ParagraphFormat())

    # <br> → new paragraph (handle text+img mixed content)
    brs = elem.findall(".//br")
    if brs:
        for text_part, img_part in _split_by_br(elem):
            if img_part is not None:
                _handle_image(doc, img_part, styles, image_max_width=image_max_width)
            elif text_part:
                p = doc.add_paragraph()
                fmt.apply_to_paragraph(p)
                _push_run(p, text_part, fmt)
        return

    p = doc.add_paragraph()
    fmt.apply_to_paragraph(p)
    _build_runs(doc, p, elem, fmt, fn_map=fn_map)


def _handle_image(doc: Document, elem: ET.Element, styles: dict, image_max_width: float = 5.5) -> None:
    """Image with alt-text caption below. Supports raster and SVG."""
    src = elem.get("src", "")
    alt = elem.get("alt", "")

    # OMML formula (block / display math) — native Word math
    omml_b64 = elem.get("data-omml-base64", "")
    if omml_b64:
        import base64
        omml_data = base64.b64decode(omml_b64).decode("utf-8")
        if omml_data:
            _insert_block_omml(doc, omml_data)
        else:
            p = doc.add_paragraph()
            _push_run(p, f"[Math: {alt}]", ParagraphFormat(), italic=True)
        return

    img_fmt = styles.get("image", ParagraphFormat())

    # Detect math formula (generated by _postprocess_math_html)
    is_math = bool(elem.get("data-svg-base64", ""))

    # Math SVG via embedded data
    svg_b64 = elem.get("data-svg-base64", "")
    if svg_b64:
        import base64
        svg_data = base64.b64decode(svg_b64)
        if not svg_data:
            p = doc.add_paragraph()
            _push_run(p, f"[Math: {alt}]", ParagraphFormat(), italic=True)
            return
        p = doc.add_paragraph()
        if not is_math:
            img_fmt.apply_to_paragraph(p)
        _embed_svg_in_paragraph(doc, p, svg_data, alt, max_width_inches=image_max_width)
    elif is_svg_source(src):
        svg_data = resolve_svg_raw(src)
        if svg_data is None:
            p = doc.add_paragraph()
            _push_run(p, f"[SVG: {alt}]", ParagraphFormat(), italic=True)
            return
        p = doc.add_paragraph()
        img_fmt.apply_to_paragraph(p)
        _embed_svg_in_paragraph(doc, p, svg_data, alt, max_width_inches=5.5)
    else:
        stream = resolve_image(src)
        if stream is None:
            p = doc.add_paragraph()
            _push_run(p, f"[Image: {alt}]", ParagraphFormat(), italic=True)
            return
        p = doc.add_paragraph()
        img_fmt.apply_to_paragraph(p)
        width, height = get_image_dimensions(stream, max_width_inches=5.5)
        run = p.add_run()
        run.add_picture(stream, width=width, height=height)

    if alt and not is_math:
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
        # Add bookmark for figure cross-references
        _add_bookmark(p_cap, f"fig-{_slugify(alt)}")


def _handle_blockquote(
    doc: Document, elem: ET.Element, styles: dict,
    fn_map: dict[str, int] | None = None,
) -> None:
    """Blockquote — process child <p> elements (fixes truncation bug)."""
    fmt = styles.get("quote", styles.get("body", ParagraphFormat()))
    for child in elem:
        if child.tag == "p":
            p = doc.add_paragraph()
            fmt.apply_to_paragraph(p)
            _build_runs(doc, p, child, fmt, fn_map=fn_map)
        else:
            p = doc.add_paragraph()
            fmt.apply_to_paragraph(p)
            _build_runs(doc, p, child if child.tag else elem, fmt, fn_map=fn_map)


def _handle_code_block(doc: Document, elem: ET.Element, styles: dict,
                       highlight_enabled: bool = True,
                       mermaid_enabled: bool = True) -> None:
    """Code block — each line is a paragraph, with optional syntax highlighting."""
    fmt = styles.get("code", ParagraphFormat())
    code_elem = elem.find("code")
    lang = extract_language(code_elem) if code_elem is not None else ""

    text = _inline_text(elem)

    # ── Mermaid diagram ──────────────────────────────────────────────────
    if mermaid_enabled and lang == "mermaid" and text.strip():
        result = render_mermaid(text.strip())
        if result:
            img_fmt = styles.get("image", ParagraphFormat())
            p = doc.add_paragraph()
            img_fmt.apply_to_paragraph(p)
            _embed_svg_in_paragraph(doc, p, result.svg_bytes, "mermaid", max_width_inches=5.5)
            return
        else:
            print("  [WARN] Falling back to plain text for mermaid block")

    tokens = None
    if highlight_enabled and lang != "mermaid":
        tokens = highlight(text, lang)

    for line in text.strip().split('\n'):
        p = doc.add_paragraph()
        fmt.apply_to_paragraph(p)
        if not line:
            run = p.add_run(" ")
            fmt.apply_to_run(run)
            set_run_font(run, fmt.font_name or "DengXian", fmt.east_asia_font or "DengXian")
            continue

        if tokens:
            line_tokens = _tokens_for_line(tokens, line)
            for tok in line_tokens:
                run = p.add_run(tok.text)
                fmt.apply_to_run(run)
                set_run_font(run, fmt.font_name or "Consolas", fmt.east_asia_font or "DengXian")
                if tok.color:
                    run.font.color.rgb = tok.color
                if tok.bold:
                    run.font.bold = True
                if tok.italic:
                    run.font.italic = True
        else:
            run = p.add_run(line)
            fmt.apply_to_run(run)
            set_run_font(run, fmt.font_name or "DengXian", fmt.east_asia_font or "DengXian")


def _tokens_for_line(tokens: list, line_text: str) -> list:
    """Filter tokens to match a single line."""
    line = line_text.rstrip("\n")
    result = []
    offset = 0
    for tok in tokens:
        if offset >= len(line):
            break
        # Find this token's text in the line starting at offset
        idx = line.find(tok.text, offset)
        if idx == -1:
            # Token not found in remaining line — skip
            continue
        if idx > offset:
            # Gap before this token — add as plain text
            result.append(type(tok)(text=line[offset:idx], color=None, bold=False, italic=False))
        result.append(tok)
        offset = idx + len(tok.text)
    if offset < len(line):
        result.append(type(tok)(text=line[offset:], color=None, bold=False, italic=False))
    return result


def _handle_unordered_list(
    doc: Document, elem: ET.Element, styles: dict, depth: int = 0,
    num_id: int | None = None,
) -> None:
    """Unordered list with Word native bullet numbering. Supports nesting."""
    fmt = styles.get("bullet_list", styles.get("body", ParagraphFormat()))
    if num_id is None:
        num_id = _create_list_num_id(doc, "bullet")
    for li in elem:
        if li.tag != "li":
            continue
        nested = li.findall("ul") + li.findall("ol")
        p = doc.add_paragraph()
        fmt.apply_to_paragraph(p)
        _add_num_pr(p, num_id, ilvl=depth)
        _build_runs_skip(doc, p, li, fmt, skip_tags={"ul", "ol"})
        if nested:
            for nest in nested:
                if nest.tag == "ul":
                    _handle_unordered_list(doc, nest, styles, depth + 1, num_id)
                else:
                    _handle_ordered_list(doc, nest, styles, depth + 1)


def _handle_ordered_list(
    doc: Document, elem: ET.Element, styles: dict, depth: int = 0,
    num_id: int | None = None,
) -> None:
    """Ordered list with Word native decimal numbering. Supports nesting."""
    fmt = styles.get("number_list", styles.get("body", ParagraphFormat()))
    if num_id is None:
        num_id = _create_list_num_id(doc, "ordered")
    for child in elem:
        if child.tag != "li":
            continue
        nested = child.findall("ul") + child.findall("ol")
        p = doc.add_paragraph()
        fmt.apply_to_paragraph(p)
        _add_num_pr(p, num_id, ilvl=depth)
        _build_runs_skip(doc, p, child, fmt, skip_tags={"ul", "ol"})
        if nested:
            for nest in nested:
                if nest.tag == "ul":
                    _handle_unordered_list(doc, nest, styles, depth + 1)
                else:
                    _handle_ordered_list(doc, nest, styles, depth + 1, num_id)


def _build_runs_skip(doc, p, elem, base_fmt, skip_tags=None):
    """Like _build_runs but skips child tags in *skip_tags*."""
    def _inner(e):
        text_parts = []
        if e.text:
            text_parts.append(e.text)
        for child in e:
            if skip_tags and child.tag in skip_tags:
                if child.tail:
                    text_parts.append(child.tail)
                continue
            if child.tag in ("strong", "b"):
                _push_run(p, child.text or "", base_fmt, bold=True)
            elif child.tag in ("em", "i"):
                _push_run(p, child.text or "", base_fmt, italic=True)
            elif child.tag == "code":
                _push_run(p, child.text or "", base_fmt, name="DengXian")
            elif child.tag == "a":
                txt = child.text or child.get("href", "")
                _push_run(p, txt, base_fmt, underline=True, color=RGBColor(0x00, 0x52, 0xCC))
            elif child.tag == "img":
                _build_runs_img(doc, p, child, base_fmt)
            else:
                _push_run(p, child.text or "", base_fmt)
            if child.tail:
                text_parts.append(child.tail)
        return "".join(text_parts)
    text = _inner(elem)
    if text:
        _push_run(p, text, base_fmt)


def _build_runs_img(doc, p, child, base_fmt):
    """Handle an <img> inside a list item."""
    src = child.get("src", "")
    alt = child.get("alt", "")
    svg_b64 = child.get("data-svg-base64", "")
    if svg_b64:
        import base64
        svg_data = base64.b64decode(svg_b64)
        if svg_data:
            _embed_svg_in_paragraph(doc, p, svg_data, alt, 3.5)
    elif is_svg_source(src):
        svg_data = resolve_svg_raw(src)
        if svg_data:
            _embed_svg_in_paragraph(doc, p, svg_data, alt, 3.5)
    else:
        stream = resolve_image(src)
        if stream is not None:
            w, h = get_image_dimensions(stream, max_width_inches=3.5)
            run = p.add_run()
            run.add_picture(stream, width=w, height=h)


def _handle_table(
    doc: Document, elem: ET.Element, styles: dict,
    three_line: bool = False,
    fn_map: dict[str, int] | None = None,
    table_idx: int = 0,
) -> None:
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
                _build_runs(doc, p, cell, body_fmt, fn_map=fn_map)
    _style_table(table, three_line=three_line)
    # Add bookmark for table cross-references
    _add_bookmark(table.rows[0].cells[0].paragraphs[0], f"tbl-{table_idx}")


def _handle_horizontal_rule(doc: Document, styles: dict) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    _add_bottom_border(p)


# ── guide-paragraph removal ──────────────────────────────────────────────


def _remove_guide_paragraphs(doc: Document) -> None:
    """Remove all guide paragraphs (style markers) from a template."""
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


# ── OOXML metadata fix ──────────────────────────────────────────────────


def _fix_ooxml_metadata(output_path: str | Path) -> None:
    """Post-process docx ZIP to fix thumbnail + application name.

    python-docx embeds a blank ``docProps/thumbnail.jpeg`` that makes
    Windows show a white box instead of a content preview.  We strip
    it so Windows generates a preview from the actual document content.

    Also fixes the Application name from "Microsoft Macintosh Word"
    to "Microsoft Office Word".
    """
    import io
    import re
    import zipfile

    path = Path(output_path)
    buf = path.read_bytes()
    out_buf = io.BytesIO()
    changed = False

    with zipfile.ZipFile(io.BytesIO(buf)) as zin:
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                # Strip blank thumbnail — Windows will preview real content
                if "thumbnail" in item.filename:
                    changed = True
                    continue

                raw = zin.read(item.filename)
                if item.filename == "docProps/app.xml":
                    text = raw.decode("utf-8")
                    fixed = text.replace(
                        "<Application>Microsoft Macintosh Word</Application>",
                        "<Application>Microsoft Office Word</Application>",
                    )
                    if fixed != text:
                        changed = True
                    raw = fixed.encode("utf-8")
                elif item.filename == "_rels/.rels":
                    # Remove the thumbnail relationship so OPC stays valid
                    without_thumb = re.sub(
                        r'<Relationship[^>]*thumbnail[^>]*/>',
                        '',
                        raw.decode("utf-8"),
                    )
                    if without_thumb != raw.decode("utf-8"):
                        changed = True
                    raw = without_thumb.encode("utf-8")
                zout.writestr(item, raw)

    if changed:
        path.write_bytes(out_buf.getvalue())


# ── GB compliance check ──────────────────────────────────────────────────


def _check_gb_compliance(doc, styles: dict, report: ConversionReport) -> None:
    """Check document formatting against Chinese GB standards, emit warnings."""
    if not doc.sections:
        return
    s = doc.sections[0]
    from docx.shared import Cm, Mm

    top = s.top_margin / 914400 * 25.4  # EMU → mm
    bottom = s.bottom_margin / 914400 * 25.4
    left = s.left_margin / 914400 * 25.4
    right = s.right_margin / 914400 * 25.4

    # Check against GB/T 9704-2012 (公文)
    gb = _GB_STANDARDS["9704_2012"]
    gb_t, gb_b, gb_l, gb_r = [v * 10 for v in gb["margins_cm"]]  # cm → mm
    checks = [
        (top, gb_t, 2.0, f"上边距 ({top:.0f}mm) 偏离 GB/T 9704-2012 标准 ({gb_t:.0f}mm)"),
        (bottom, gb_b, 2.0, f"下边距 ({bottom:.0f}mm) 偏离 GB/T 9704-2012 标准 ({gb_b:.0f}mm)"),
        (left, gb_l, 2.0, f"左边距 ({left:.0f}mm) 偏离 GB/T 9704-2012 标准 ({gb_l:.0f}mm)"),
        (right, gb_r, 2.0, f"右边距 ({right:.0f}mm) 偏离 GB/T 9704-2012 标准 ({gb_r:.0f}mm)"),
    ]
    for actual, expected, tol, msg in checks:
        if abs(actual - expected) > tol:
            report.info_msg(f"[GB] 非公文模板: {msg}")

    # Check body font
    body_fmt = styles.get("body")
    if body_fmt and body_fmt.font_name:
        if body_fmt.font_name not in ("FangSong", "SimSun", "SimHei"):
            report.info_msg(f"[GB] 正文字体 '{body_fmt.font_name}' — 非标准公文/学术字体（推荐仿宋/宋体）")


# ── Red-head document header ────────────────────────────────────────────


def _insert_redhead_header(doc, authority_name: str, styles: dict) -> None:
    """Insert 红头文件 header elements at the document start.

    Structure:
        1. Red authority name (centered, large, bold, red)
        2. Full-width red separator line
        3. Optional document number line (empty, user fills in Word)
    """
    from docx.shared import Pt, RGBColor, Cm

    # ── Red header: authority name ────────────────────────────────────
    p_red = doc.add_paragraph()
    p_red.paragraph_format.alignment = 1  # center
    p_red.paragraph_format.space_after = Pt(4)
    run = p_red.add_run(authority_name)
    run.font.name = "SimHei"
    run.font.size = Pt(28)  # 一号
    run.font.bold = True
    from .template import set_run_font
    set_run_font(run, "SimHei", "SimHei")
    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)  # 标准红

    # Add "文件" suffix on next line
    p_suffix = doc.add_paragraph()
    p_suffix.paragraph_format.alignment = 1
    p_suffix.paragraph_format.space_after = Pt(6)
    run2 = p_suffix.add_run("文件")
    run2.font.name = "SimHei"
    run2.font.size = Pt(28)
    run2.font.bold = True
    set_run_font(run2, "SimHei", "SimHei")
    run2.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

    # ── Red separator line ────────────────────────────────────────────
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(2)
    p_line.paragraph_format.space_after = Pt(12)
    _add_bottom_border(p_line, color="CC0000", sz="16")

    # ── Document number placeholder ───────────────────────────────────
    p_num = doc.add_paragraph()
    p_num.paragraph_format.alignment = 1
    p_num.paragraph_format.space_after = Pt(12)
    run3 = p_num.add_run("〔2024〕 号")
    run3.font.name = "FangSong"
    run3.font.size = Pt(16)
    set_run_font(run3, "FangSong", "FangSong")


# ── Page number formatting ─────────────────────────────────────────────


def _set_page_number_format(doc, fmt: str = "-- %d --") -> None:
    """Set page number in document footer.

    *fmt* uses ``%d`` as placeholder for the page number.
    Example: ``-- %d --`` → ``-- 1 --``
    """
    if not doc.sections:
        return

    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.paragraph_format.alignment = 1  # center

    # Split format around %d
    parts = fmt.split("%d", 1)
    prefix = parts[0]
    suffix = parts[1] if len(parts) > 1 else ""

    if prefix:
        run_pre = p.add_run(prefix)
        run_pre.font.size = Pt(10)

    # PAGE field
    run_field = OxmlElement("w:r")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run_field.append(fld_begin)
    p._p.append(run_field)

    run_instr = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    run_instr.append(instr)
    p._p.append(run_instr)

    run_sep = OxmlElement("w:r")
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run_sep.append(fld_sep)
    p._p.append(run_sep)

    run_disp = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    run_disp.append(t)
    p._p.append(run_disp)

    run_end = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run_end.append(fld_end)
    p._p.append(run_end)

    if suffix:
        run_suf = p.add_run(suffix)
        run_suf.font.size = Pt(10)


# ── main ────────────────────────────────────────────────────────────────────


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
) -> None:
    """Convert *markdown_text* to a Word document.

    Args:
        markdown_text: Raw markdown content.
        template_path: Path to the .docx template with guide paragraphs.
        output_path: Where to write the result.
        image_max_width: Maximum image width in inches.
        toc: Whether to generate a Table of Contents.
        toc_depth: Heading levels to include, e.g. "1-3".
        three_line_table: Use academic three-line table style.
        footnotes_enabled: Process footnote syntax [^id].
        formula_numbering: Add SEQ equation numbers to block math.
        redhead_authority: Issuing authority name for red-head document.
        page_number_fmt: Page number format, e.g. "-- %d --".
        gb_check: Check formatting against GB standards.
    """
    styles = extract_template_styles(template_path)
    report = ConversionReport()
    styles["_report"] = report

    markdown_text = _strip_front_matter(markdown_text)

    # ── Footnote preprocessing ────────────────────────────────────────────
    fn_list: list = []
    fn_map: dict[str, int] = {}
    if footnotes_enabled:
        markdown_text, fn_list = extract_footnotes(markdown_text)

    # ── Math preprocessing: extract LaTeX expressions ──────────────────────
    if math_enabled:
        markdown_text, math_exprs = extract_and_placeholder(markdown_text)
    else:
        math_exprs = []

    # ── Ensure list items have preceding blank lines ───────────────────
    markdown_text = _ensure_list_blank_lines(markdown_text)

    html = markdown.markdown(
        markdown_text,
        extensions=[
            "extra",
            "tables",
            "fenced_code",
            "sane_lists",
        ],
    )

    html = _prepare_html(html)

    # ── Math postprocessing: render LaTeX to SVG, embed in HTML ───────────
    if math_exprs:
        html = _postprocess_math_html(html, math_exprs)

    # ── Replace footnote placeholders with safe HTML tags ────────────
    if fn_list:
        html = _replace_fn_placeholders(html)

    blocks = _html_to_blocks(html)

    doc = Document(str(template_path))

    # Ensure first section doesn't force a new page (prevents blank first page)
    if doc.sections:
        doc.sections[0].start_type = 0  # WD_SECTION_START.CONTINUOUS

    # Remove guide paragraphs AND empty paragraph at top
    _remove_guide_paragraphs(doc)
    if doc.paragraphs and not doc.paragraphs[0].text.strip():
        p = doc.paragraphs[0]._p
        p.getparent().remove(p)

    # ── Red-head header (insert at very beginning) ──────────────────────
    if redhead_authority:
        _insert_redhead_header(doc, redhead_authority, styles)

    # ── Page number format ─────────────────────────────────────────────
    if page_number_fmt:
        _set_page_number_format(doc, page_number_fmt)

    # ── GB standards compliance check ──────────────────────────────────
    if gb_check:
        _check_gb_compliance(doc, styles, report)

    # ── Table of Contents (inserted after the first h1 heading) ─────────
    _toc_pending = toc
    check_blocks = list(blocks)
    if _toc_pending and not any(b.tag == "h1" for b in check_blocks):
        _add_toc(doc, styles, depth=toc_depth)
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

    # ── Heading numbering ─────────────────────────────────────────────────
    if number_headings:
        _reset_heading_counters()

    # ── Add footnotes to doc (before block processing for fn_map) ─────────
    if fn_list:
        fn_map = add_footnotes_to_document(doc, fn_list)

    # ── Process blocks ────────────────────────────────────────────────────
    table_counter = [0]

    for block in blocks:
        tag = block.tag
        try:
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                _handle_heading(doc, block, int(tag[1]), styles,
                                number_headings=number_headings,
                                page_break=page_break_h1,
                                fn_map=fn_map if fn_list else None)
                if _toc_pending and tag == "h1":
                    _add_toc(doc, styles, depth=toc_depth)
                    _toc_pending = False
            elif tag == "p":
                _handle_paragraph(doc, block, styles,
                                  image_max_width=image_max_width,
                                  fn_map=fn_map if fn_list else None)
            elif tag == "img":
                _handle_image(doc, block, styles,
                              image_max_width=image_max_width)
                # Formula numbering for block-level math
                if formula_numbering and block.get("class", "").startswith("math-block"):
                    p_num = doc.add_paragraph()
                    p_num.paragraph_format.alignment = 3  # right
                    _add_seq_number(p_num, "Equation")
            elif tag == "blockquote":
                _handle_blockquote(doc, block, styles,
                                   fn_map=fn_map if fn_list else None)
            elif tag == "pre":
                _handle_code_block(doc, block, styles,
                                   highlight_enabled=highlight_enabled,
                                   mermaid_enabled=mermaid_enabled)
            elif tag == "ul":
                _handle_unordered_list(doc, block, styles)
            elif tag == "ol":
                _handle_ordered_list(doc, block, styles)
            elif tag == "table":
                _handle_table(doc, block, styles,
                              three_line=three_line_table,
                              fn_map=fn_map if fn_list else None,
                              table_idx=table_counter[0])
                table_counter[0] += 1
            elif tag == "hr":
                _handle_horizontal_rule(doc, styles)
            else:
                _handle_paragraph(doc, block, styles,
                                  fn_map=fn_map if fn_list else None)
        except Exception as e:
            report.warn(f"Failed to process <{tag}> block: {e}")

    doc.save(str(output_path))

    # ── Fix OOXML metadata for proper Windows icon/thumbnail ────────────
    _fix_ooxml_metadata(output_path)

    report.info_msg("输出文件: " + str(output_path))
    print(f"  {report.summary()}")
    return report
