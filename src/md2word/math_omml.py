"""LaTeX → Word OMML (Office Math Markup Language) conversion.

Converts LaTeX math expressions to native Word formulas that are
fully editable in Word's equation editor and MathType.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape


# Large operators that become <m:nary> in OMML
_LARGE_OPERATORS = frozenset(
    chr(c)
    for c in [
        0x2211,  # ∑ SUM
        0x220F,  # ∏ PRODUCT
        0x2210,  # ∐ COPRODUCT
        0x222B,  # ∫ INTEGRAL
        0x222C,  # ∬ DOUBLE INTEGRAL
        0x222D,  # ∭ TRIPLE INTEGRAL
        0x222E,  # ∮ CONTOUR INTEGRAL
        0x22C2,  # ⋂ INTERSECTION
        0x22C3,  # ⋃ UNION
        0x2A01,  # ⨁ CIRCLED PLUS
        0x2A02,  # ⨂ CIRCLED TIMES
    ]
)

# Function names that look better as <m:limLow> than <m:sSub>
_FUNC_NAMES = frozenset(
    "lim liminf limsup sup max min inf det Pr gcd deg".split()
)

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


# ── Public API ─────────────────────────────────────────────────────────────


def latex_to_omml(latex: str, display: bool = False) -> str | None:
    """Convert LaTeX to an OMML XML fragment.

    Returns an ``<m:oMath>`` (inline) or ``<m:oMathPara>`` (display) XML
    string ready for insertion into a python-docx paragraph, or ``None``
    if conversion fails.
    """
    try:
        from latex2mathml.converter import convert as latex_to_mathml

        mathml = latex_to_mathml(latex).strip()
        inner = _convert_mathml(mathml)
        if inner is None:
            return None
        if display:
            return (
                f'<m:oMathPara xmlns:m="{MATH_NS}">'
                f"<m:oMath>{inner}</m:oMath>"
                f"</m:oMathPara>"
            )
        return f'<m:oMath xmlns:m="{MATH_NS}">{inner}</m:oMath>'
    except Exception as exc:
        import logging

        logging.getLogger(__name__).warning(
            "OMML conversion failed for '%s': %s", latex[:60], exc
        )
        return None


# ── MathML → OMML conversion ──────────────────────────────────────────────


def _convert_mathml(mathml: str) -> str | None:
    """Convert MathML presentation markup to OMML inner content."""
    try:
        root = ET.fromstring(mathml)
    except ET.ParseError:
        return None
    parts = [_convert_node(child) for child in root]
    return "".join(p for p in parts if p)


def _tag(elem: ET.Element) -> str:
    """Get the local tag name (strip namespace)."""
    return elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag


def _text(elem: ET.Element) -> str:
    """Get the full text content of an element."""
    return elem.text or ""


def _make_run(text: str, *, style: str | None = None) -> str:
    """Create an OMML run (<m:r><m:t>…</m:t></m:r>)."""
    safe = xml_escape(text)
    rpr = f'<m:rPr><m:sty m:val="{style}"/></m:rPr>' if style else ""
    return f"<m:r>{rpr}<m:t>{safe}</m:t></m:r>"


def _convert_children(elem: ET.Element) -> str:
    """Convert all children of *elem* and concatenate."""
    return "".join(_convert_node(child) for child in elem)


def _convert_node(elem: ET.Element | None) -> str:
    """Recursively convert a MathML element to OMML."""
    if elem is None:
        return ""

    tag = _tag(elem)
    text = _text(elem)
    children = list(elem)

    # ── atomic text nodes ──────────────────────────────────────────────
    if tag in ("mi", "mo", "mn", "mtext"):
        return _make_run(text)

    # ── mspace (ignore) ────────────────────────────────────────────────
    if tag == "mspace":
        return ""

    # ── mrow (flatten children) ────────────────────────────────────────
    if tag == "mrow":
        return _convert_children(elem)

    # ── mfrac (fraction) ───────────────────────────────────────────────
    if tag == "mfrac":
        if len(children) >= 2:
            num = _convert_node(children[0])
            den = _convert_node(children[1])
            lt = elem.get("linethickness", "")
            if lt in ("0", "0pt"):
                return (
                    f'<m:f><m:fPr><m:type m:val="noBar"/>'
                    f"</m:fPr><m:num>{num}</m:num><m:den>{den}</m:den></m:f>"
                )
            return f"<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>"
        return ""

    # ── msqrt (square root) ────────────────────────────────────────────
    if tag == "msqrt":
        inner = _convert_children(elem)
        return f"<m:rad><m:e>{inner}</m:e></m:rad>"

    # ── mroot (n-th root) ──────────────────────────────────────────────
    if tag == "mroot":
        if len(children) >= 2:
            deg = _convert_node(children[1])
            base = _convert_node(children[0])
            return f"<m:rad><m:deg>{deg}</m:deg><m:e>{base}</m:e></m:rad>"
        return ""

    # ── msup (superscript) ─────────────────────────────────────────────
    if tag == "msup":
        if len(children) >= 2:
            e = _convert_node(children[0])
            sup = _convert_node(children[1])
            return f"<m:sSup><m:e>{e}</m:e><m:sup>{sup}</m:sup></m:sSup>"
        return ""

    # ── msub (subscript) ───────────────────────────────────────────────
    if tag == "msub":
        if len(children) >= 2:
            base = _convert_node(children[0])
            sub = _convert_node(children[1])
            if _is_large_operator(children[0]):
                op_char = _get_base_text(children[0]) or "∫"
                return (
                    f"<m:nary>"
                    f'<m:naryPr><m:chr m:val="{xml_escape(op_char)}"/>'
                    f'<m:limLoc m:val="undOvr"/></m:naryPr>'
                    f"<m:sub>{sub}</m:sub>"
                    f"<m:e></m:e>"
                    f"</m:nary>"
                )
            base_text = _get_base_text(children[0])
            if base_text and base_text in _FUNC_NAMES:
                return (
                    f"<m:limLow><m:e>{base}</m:e>"
                    f"<m:lim>{sub}</m:lim></m:limLow>"
                )
            return f"<m:sSub><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>"
        return ""

    # ── msubsup (sub + superscript) ───────────────────────────────────
    if tag == "msubsup":
        if len(children) >= 3:
            base = _convert_node(children[0])
            sub = _convert_node(children[1])
            sup = _convert_node(children[2])
            if _is_large_operator(children[0]):
                op_char = _get_base_text(children[0]) or "∑"
                return (
                    f"<m:nary>"
                    f'<m:naryPr><m:chr m:val="{xml_escape(op_char)}"/>'
                    f'<m:limLoc m:val="undOvr"/></m:naryPr>'
                    f"<m:sub>{sub}</m:sub>"
                    f"<m:sup>{sup}</m:sup>"
                    f"<m:e></m:e>"
                    f"</m:nary>"
                )
            return (
                f"<m:sSubSup>"
                f"<m:e>{base}</m:e>"
                f"<m:sub>{sub}</m:sub>"
                f"<m:sup>{sup}</m:sup>"
                f"</m:sSubSup>"
            )
        return ""

    # ── mover (overscript / accent) ────────────────────────────────────
    if tag == "mover":
        if len(children) >= 2:
            base = _convert_node(children[0])
            over = _convert_node(children[1])
            accent = elem.get("accent", "")
            if accent == "true" or _is_accent(children[1]):
                acc_char = _get_base_text(children[1])
                return (
                    f'<m:acc><m:e>{base}</m:e>'
                    f'<m:accPr><m:chr m:val="{xml_escape(acc_char)}"/>'
                    f"</m:accPr></m:acc>"
                )
            return f"<m:limUpp><m:e>{base}</m:e><m:lim>{over}</m:lim></m:limUpp>"
        return ""

    # ── munder (underscript) ───────────────────────────────────────────
    if tag == "munder":
        if len(children) >= 2:
            base = _convert_node(children[0])
            under = _convert_node(children[1])
            return (
                f"<m:limLow><m:e>{base}</m:e>"
                f"<m:lim>{under}</m:lim></m:limLow>"
            )
        return ""

    # ── munderover ─────────────────────────────────────────────────────
    if tag == "munderover":
        if len(children) >= 3:
            base = _convert_node(children[0])
            under = _convert_node(children[1])
            over = _convert_node(children[2])
            if _is_large_operator(children[0]):
                op_char = _get_base_text(children[0]) or "∑"
                return (
                    f"<m:nary>"
                    f'<m:naryPr><m:chr m:val="{xml_escape(op_char)}"/>'
                    f'<m:limLoc m:val="undOvr"/></m:naryPr>'
                    f"<m:sub>{under}</m:sub>"
                    f"<m:sup>{over}</m:sup>"
                    f"<m:e></m:e>"
                    f"</m:nary>"
                )
            return f"<m:limLow><m:e>{base}</m:e><m:lim>{under}</m:lim></m:limLow>{over}"
        return ""

    # ── mpadded, menclose, etc – skip wrapper, process children ───────
    if tag in ("mpadded", "menclose", "mphantom", "mstyle"):
        return _convert_children(elem)

    # ── mmultiscripts, mtable, etc – too complex, fallback ────────────
    # Return text content as an OMML run
    full = _full_text(elem)
    return _make_run(full) if full else ""


# ── helpers ────────────────────────────────────────────────────────────────


def _get_base_text(elem: ET.Element) -> str:
    """Extract the text content from a MathML element for operator detection."""
    tag = _tag(elem)
    if tag in ("mi", "mo", "mn", "mtext"):
        return (elem.text or "").strip()
    if tag == "mrow":
        for child in elem:
            bt = _get_base_text(child)
            if bt:
                return bt
    return ""


def _full_text(elem: ET.Element) -> str:
    """Get all text content recursively."""
    parts = []
    if elem.text:
        parts.append(elem.text)
    for child in elem:
        parts.append(_full_text(child))
        if child.tail:
            parts.append(child.tail)
    return "".join(parts)


def _is_large_operator(elem: ET.Element) -> bool:
    """Check if *elem* represents a large operator (∑, ∫, ∏, …)."""
    tag = _tag(elem)
    if tag == "mo":
        t = (elem.text or "").strip()
        return t in _LARGE_OPERATORS
    if tag == "mrow":
        for child in elem:
            if _is_large_operator(child):
                return True
    return False


def _is_accent(elem: ET.Element) -> bool:
    """Check if *elem* is an accent character (ˆ, ¯, ˙, ¨, →, ~)."""
    tag = _tag(elem)
    if tag == "mo":
        t = (elem.text or "").strip()
        # Check common accent characters
        cp = ord(t) if len(t) == 1 else 0
        return cp in (0x005E, 0x02C6, 0x02DC, 0x007E, 0x00AF, 0x00B4, 0x00A8, 0x02D9, 0x2192)
    return False
