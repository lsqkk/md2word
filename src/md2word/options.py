"""ConvertOptions dataclass — configuration for convert().

Replaces the long keyword-argument list in ``convert()`` with a single
dataclass, making the API cleaner and easier to document.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ConvertOptions:
    """Configuration options for :func:`convert`.

    All fields have sensible defaults — you only need to set what
    differs from the default.
    """

    # ── Content / display ────────────────────────────────────────────────
    image_max_width: float = 5.5
    """Maximum image width in inches (default 5.5)."""

    toc: bool = True
    """Whether to generate a Table of Contents."""

    toc_depth: str = "1-3"
    """Heading levels to include in the TOC, e.g. ``"1-3"``."""

    number_headings: bool = False
    """Add auto-numbering to headings (1, 1.1, 1.1.1, etc.)."""

    page_break_h1: bool = False
    """Add a page break before each H1 heading."""

    three_line_table: bool = False
    """Use academic three-line table style (三线表)."""

    footnotes_enabled: bool = True
    """Process footnote syntax ``[^id]``."""

    formula_numbering: bool = False
    """Add SEQ equation numbers to block math."""

    # ── Optional rendering engines ────────────────────────────────────────
    highlight_enabled: bool = True
    """Enable code syntax highlighting via Pygments."""

    math_enabled: bool = True
    """Enable math formula rendering (``$...$`` / ``$$...$$``)."""

    mermaid_enabled: bool = True
    """Enable Mermaid diagram rendering."""

    # ── Document production ──────────────────────────────────────────────
    redhead_authority: str | None = None
    """Issuing authority name for red-head official document."""

    redhead_year: int | None = None
    """Year in the red-head document number (default 2024)."""

    redhead_number: str | None = None
    """Document number string for red-head (e.g. ``"12"``)."""

    page_number_fmt: str | None = None
    """Page number format string, e.g. ``"-- %d --"``."""

    gb_check: bool = False
    """Check formatting against GB/T standards and report violations."""

    # ── Style / mapping ──────────────────────────────────────────────────
    style_map: dict[str, str] | None = None
    """Optional mapping of element type → Word style name override."""

    # ── Update check ────────────────────────────────────────────────────
    update_check: bool = True
    """Check GitHub for newer versions after successful conversion."""

    # ── Diagnostics ──────────────────────────────────────────────────────
    verbose: bool = False
    """Emit detailed progress to stderr."""

    # ── Factory helpers ──────────────────────────────────────────────────

    @classmethod
    def from_cli_args(cls, args: dict, cfg: dict | None = None) -> ConvertOptions:
        """Build a ConvertOptions from a CLI args dict (and optional config).

        CLI values take priority over config values; config defaults are
        used when CLI args are ``None``.
        """
        merged = dict(cfg or {})
        merged.update({k: v for k, v in args.items() if v is not None})
        return cls(**{k: v for k, v in merged.items()
                      if k in cls.__dataclass_fields__})
