"""Footnote processing — extract Markdown footnotes, insert Word native footnotes.

Supports the standard Markdown footnote syntax:

    Here is some text with a footnote[^1].

    [^1]: The footnote content goes here.
          It can span multiple lines.
"""

from __future__ import annotations

import re
from typing import NamedTuple

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class FootnoteDef(NamedTuple):
    """A single footnote definition extracted from markdown."""
    id: str
    content: str


# ── Extraction ───────────────────────────────────────────────────────────────


# Footnote definition: [^id]: content (possibly multi-line)
_DEF_PATTERN = re.compile(
    r"^\[\^([^\]]+)\]:\s*(.*?)(?=\n(?:\[\^|$|\n(?!\s{2,}))|$)",
    re.MULTILINE | re.DOTALL,
)

# Inline footnote reference: [^id]
_REF_PATTERN = re.compile(r"\[\^([^\]]+)\]")


def extract_footnotes(text: str) -> tuple[str, list[FootnoteDef]]:
    """Extract footnote definitions from *text*.

    Returns (cleaned_text, footnotes) where cleaned_text has all footnote
    definitions removed and inline references replaced with placeholders.
    """
    # Find definitions
    defs: dict[str, str] = {}
    for m in _DEF_PATTERN.finditer(text):
        fid = m.group(1).strip()
        content = m.group(2).strip().replace("\n", " ")
        defs[fid] = content

    if not defs:
        return text, []

    # Remove definitions from text
    cleaned = _DEF_PATTERN.sub("", text)

    # Replace inline references with placeholders
    footnotes: list[FootnoteDef] = []
    seen: set[str] = set()

    def _replace_ref(m: re.Match) -> str:
        fid = m.group(1)
        if fid in defs:
            if fid not in seen:
                seen.add(fid)
                footnotes.append(FootnoteDef(id=fid, content=defs[fid]))
            return "\x00FN_" + fid + "\x00"
        return m.group(0)

    cleaned = _REF_PATTERN.sub(_replace_ref, cleaned)
    return cleaned, footnotes


# ── Word footnote insertion ──────────────────────────────────────────────────


def add_footnotes_to_document(doc, footnotes: list[FootnoteDef]) -> dict[str, int]:
    """Add native Word footnotes to a python-docx Document.

    Returns a dict mapping footnote ID → Word footnote ID (integer).
    """
    if not footnotes:
        return {}

    from lxml import etree
    from docx.opc.part import Part
    from docx.oxml import parse_xml

    fn_id_map: dict[str, int] = {}

    doc_part = doc.part

    # Build footnotes XML
    root_xml = (
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships">'
        '  <w:footnote w:type="separator" w:id="-1">'
        '    <w:p><w:r><w:separator/></w:r></w:p>'
        '  </w:footnote>'
        '  <w:footnote w:type="continuationSeparator" w:id="0">'
        '    <w:p><w:r><w:continuationSeparator/></w:r></w:p>'
        '  </w:footnote>'
        '</w:footnotes>'
    )
    root = parse_xml(root_xml)

    for i, fn in enumerate(footnotes):
        fn_id = i + 1
        fn_id_map[fn.id] = fn_id
        _add_footnote_to_xml(root, fn_id, fn.content)

    # Serialize to bytes and create Part (use standard partname)
    fn_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    content_type = (
        "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.footnotes+xml"
    )
    from docx.opc.packuri import PackURI
    fn_part = Part(
        PackURI("/word/footnotes.xml"),
        content_type,
        fn_bytes,
        doc_part.package,
    )
    # Remove any existing footnotes relationship, then add ours
    existing_fn = [
        r for r in doc_part.rels.values()
        if r.reltype.endswith("/footnotes")
    ]
    for rel in existing_fn:
        del doc_part.rels[rel.rId]
    doc_part.relate_to(
        fn_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    )

    return fn_id_map


def _add_footnote_to_xml(footnotes_elem, fn_id: int, content: str) -> None:
    """Add a single footnote to the footnotes XML element."""
    fn = OxmlElement("w:footnote")
    fn.set(qn("w:id"), str(fn_id))

    p = OxmlElement("w:p")

    # Footnote reference run
    r1 = OxmlElement("w:r")
    rPr1 = OxmlElement("w:rPr")
    rStyle1 = OxmlElement("w:rStyle")
    rStyle1.set(qn("w:val"), "FootnoteReference")
    rPr1.append(rStyle1)
    r1.append(rPr1)
    fn_ref = OxmlElement("w:footnoteRef")
    r1.append(fn_ref)
    p.append(r1)

    # Space after footnote number
    r_space = OxmlElement("w:r")
    rPr_space = OxmlElement("w:rPr")
    rStyle_space = OxmlElement("w:rStyle")
    rStyle_space.set(qn("w:val"), "FootnoteReference")
    rPr_space.append(rStyle_space)
    r_space.append(rPr_space)
    t_space = OxmlElement("w:t")
    t_space.set(qn("xml:space"), "preserve")
    t_space.text = " "
    r_space.append(t_space)
    p.append(r_space)

    # Footnote text run
    r2 = OxmlElement("w:r")
    t2 = OxmlElement("w:t")
    t2.set(qn("xml:space"), "preserve")
    t2.text = content
    r2.append(t2)
    p.append(r2)

    fn.append(p)
    footnotes_elem.append(fn)


def create_footnote_reference_run(doc, fn_word_id: int):
    """Create a run containing a footnote reference mark."""
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "FootnoteReference")
    rPr.append(rStyle)
    run.append(rPr)
    fn_ref = OxmlElement("w:footnoteReference")
    fn_ref.set(qn("w:id"), str(fn_word_id))
    run.append(fn_ref)
    return run
