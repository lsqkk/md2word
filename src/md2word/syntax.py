"""Code syntax highlighting via Pygments."""

from __future__ import annotations

import re
from typing import NamedTuple

from docx.shared import RGBColor


class HighlightToken(NamedTuple):
    text: str
    color: RGBColor | None
    bold: bool
    italic: bool


# Pygments token → docx color mapping (default style approximation)
_TOKEN_COLORS: dict[str, RGBColor] = {
    "Token.Keyword": RGBColor(0x00, 0x00, 0xFF),           # blue
    "Token.Keyword.Type": RGBColor(0x00, 0x00, 0xFF),
    "Token.Keyword.Constant": RGBColor(0x00, 0x00, 0xFF),
    "Token.Keyword.Declaration": RGBColor(0x00, 0x00, 0xFF),
    "Token.Keyword.Namespace": RGBColor(0x00, 0x80, 0x80),  # teal
    "Token.Name.Builtin": RGBColor(0x00, 0x80, 0x80),
    "Token.Name.Function": RGBColor(0x80, 0x00, 0x80),     # purple
    "Token.Name.Class": RGBColor(0x00, 0x80, 0x80),
    "Token.Name.Decorator": RGBColor(0x80, 0x40, 0x00),    # brown
    "Token.Name.Exception": RGBColor(0x00, 0x80, 0x00),    # green
    "Token.Literal.String": RGBColor(0x00, 0x80, 0x00),    # green
    "Token.Literal.String.Doc": RGBColor(0x80, 0x80, 0x80), # gray
    "Token.Literal.Number": RGBColor(0x00, 0x00, 0x80),    # dark blue
    "Token.Operator": RGBColor(0x66, 0x66, 0x66),          # gray
    "Token.Punctuation": RGBColor(0x33, 0x33, 0x33),       # dark gray
    "Token.Comment": RGBColor(0x00, 0x80, 0x00),           # green
    "Token.Comment.Special": RGBColor(0x00, 0x80, 0x00),
}


def highlight(code: str, lang: str = "") -> list[HighlightToken] | None:
    """Tokenize *code* with Pygments and return colored runs.

    Returns None if Pygments is not available or language is not supported.
    """
    try:
        from pygments import lex
        from pygments.lexers import get_lexer_by_name, guess_lexer
        from pygments.token import Token
    except ImportError:
        return None

    try:
        if lang:
            lexer = get_lexer_by_name(lang, stripall=False, stripnl=False)
        else:
            lexer = guess_lexer(code)
    except Exception:
        return None

    result: list[HighlightToken] = []
    for token_type, text in lex(code, lexer):
        if not text:
            continue
        # Walk the token type hierarchy to find a color
        color = None
        for cls in _walk_token_hierarchy(token_type, Token):
            key = str(cls)
            if key in _TOKEN_COLORS:
                color = _TOKEN_COLORS[key]
                break
        bold = token_type in (Token.Keyword, Token.Name.Function, Token.Name.Class)
        italic = token_type in (Token.Comment, Token.Literal.String.Doc)
        result.append(HighlightToken(text=text, color=color, bold=bold, italic=italic))

    return result


def _walk_token_hierarchy(ttype, root):
    """Yield the token type and all its parents up to *root*."""
    yield ttype
    parent = ttype.parent
    while parent is not None and parent != root:
        yield parent
        parent = parent.parent
    yield root


def extract_language(elem) -> str:
    """Extract language from a <code> element's class attribute."""
    cls = elem.get("class", "")
    m = re.search(r"language-(\w+)", cls)
    return m.group(1) if m else ""
