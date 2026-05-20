"""Math formula rendering — LaTeX → SVG."""

from __future__ import annotations

import io
import re
import tempfile
from pathlib import Path


_INLINE_PATTERN = re.compile(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)")
_BLOCK_PATTERN = re.compile(r"\$\$(.+?)\$\$", re.DOTALL)


def has_math(text: str) -> bool:
    """Check if *text* contains any LaTeX math expressions."""
    return bool(_INLINE_PATTERN.search(text) or _BLOCK_PATTERN.search(text))


def extract_and_placeholder(text: str) -> tuple[str, list[dict]]:
    """Replace math expressions with placeholders.

    Returns (modified_text, expressions) where each expression has:
        kind: "inline" or "block"
        latex: the raw LaTeX string
        placeholder: the placeholder token
    """
    expressions: list[dict] = []
    idx = [0]

    def _save_block(m):
        latex = m.group(1).strip()
        ph = f"\x00BLOCKMATH{idx[0]}\x00"
        idx[0] += 1
        expressions.append({"kind": "block", "latex": latex, "placeholder": ph})
        return ph

    def _save_inline(m):
        latex = m.group(1).strip()
        ph = f"\x00INLINEMATH{idx[0]}\x00"
        idx[0] += 1
        expressions.append({"kind": "inline", "latex": latex, "placeholder": ph})
        return ph

    text = _BLOCK_PATTERN.sub(_save_block, text)
    text = _INLINE_PATTERN.sub(_save_inline, text)
    return text, expressions


# Shorthand LaTeX → long-form mappings that matplotlib's mathtext requires.
# Use word-boundary regex to avoid collisions (e.g. \left should not become
# \leqft when normalizing \le).
_LATEX_NORMALIZE: list[tuple[re.Pattern, str]] = [
    (re.compile(r"\\ge(?!\w)"), r"\geq"),
    (re.compile(r"\\le(?!\w)"), r"\leq"),
    (re.compile(r"\\ne(?!\w)"), r"\neq"),
    (re.compile(r"\\lvert(?!\w)"), r"|"),
    (re.compile(r"\\rvert(?!\w)"), r"|"),
]


def _normalize_latex(latex: str) -> str:
    """Expand shorthand LaTeX macros matplotlib doesn't understand."""
    for pattern, replacement in _LATEX_NORMALIZE:
        latex = pattern.sub(lambda m: replacement, latex)
    return latex


def render_math_svg(latex: str, kind: str = "inline") -> tuple[bytes | None, int, int]:
    """Render LaTeX to SVG using matplotlib's mathtext.

    Returns (svg_bytes, width_px, height_px) or (None, 0, 0) on failure.

    Notes:
        matplotlib 3.10+ removed ``MathTextParser("svg")``, so we render via
        a temporary figure and export its text bounding box as SVG.
    """
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return None, 0, 0

    try:
        latex = _normalize_latex(latex)
        fontsize = 14 if kind == "block" else 11
        dpi = 120

        fig, ax = plt.subplots(figsize=(0.01, 0.01))
        # matplotlib.text adds $…$ automatically for mathtext rendering
        ax.text(0, 0, f"${latex}$", fontsize=fontsize, va="bottom", ha="left")
        ax.axis("off")

        buf = io.BytesIO()
        fig.savefig(buf, format="svg", dpi=dpi, bbox_inches="tight",
                    pad_inches=0.01, transparent=True)
        plt.close(fig)
        svg_str = buf.getvalue().decode("utf-8")

        # Parse SVG dimensions using shared utility
        from .image_utils import parse_svg_size as _parse_svg_size

        size = _parse_svg_size(svg_str.encode("utf-8"))
        if size:
            w_pt, h_pt = size
        else:
            w_pt, h_pt = 100.0, 30.0

        # Convert pt → px at given DPI  (1 pt = 1/72 inch)
        scale = dpi / 72.0
        w_px = max(1, int(w_pt * scale))
        h_px = max(1, int(h_pt * scale))

        svg_bytes = svg_str.encode("utf-8")
        return svg_bytes, w_px, h_px
    except Exception as e:
        print(f"  [WARN] Failed to render math: {e}")
        print(f"  [WARN]   LaTeX: {latex[:80]}")
        return None, 0, 0
