"""CLI entry point for md2word."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .converter import convert
from .template import list_template_styles
from .themes import build_theme, get_theme, list_themes as _list_themes


# ── CLI ──────────────────────────────────────────────────────────────────────


def _create_template(output_path: Path, theme: str = "official") -> None:
    spec = get_theme(theme)
    if spec is None:
        names = ", ".join(n for n, _ in _list_themes())
        print(f"Unknown theme '{theme}'. Available: {names}")
        print("Falling back to 'official'.")
        spec = get_theme("official")

    build_theme(spec, output_path)
    print(f"Template created: {output_path}")
    print(f"  Theme: {spec.label} — {spec.desc}")
    print("Open it in Word, adjust any style, then save.")
    print("The tool detects styles by marker keywords (一级标题, 正文, etc.).")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    theme_names = [n for n, _ in _list_themes()]

    parser = argparse.ArgumentParser(
        prog="md2word",
        description="Convert Markdown to Word (.docx) using a custom template.",
    )

    parser.add_argument(
        "inputs", nargs="*", type=str, default=[], help="Input Markdown file(s) (omit to read stdin)"
    )
    parser.add_argument("-o", "--output", type=str, default=None, help="Output .docx path")
    parser.add_argument(
        "-t", "--template", type=str, default=None, help="Template .docx with guide paragraphs"
    )
    parser.add_argument(
        "--image-width", type=float, default=5.5, help="Max image width in inches (default: 5.5)"
    )
    parser.add_argument(
        "--toc", action="store_true", default=None, help="Generate table of contents"
    )
    parser.add_argument(
        "--no-toc", action="store_true", dest="no_toc", default=None, help="Skip table of contents"
    )
    parser.add_argument(
        "--toc-depth", type=str, default="1-3", help="TOC heading depth range (default: 1-3)"
    )
    parser.add_argument(
        "--no-highlight", action="store_true", dest="no_highlight", default=False,
        help="Disable code syntax highlighting"
    )
    parser.add_argument(
        "--no-math", action="store_true", dest="no_math", default=False,
        help="Disable math formula rendering ($...$ / $$...$$)"
    )
    parser.add_argument(
        "--no-mermaid", action="store_true", dest="no_mermaid", default=False,
        help="Disable mermaid diagram rendering"
    )
    parser.add_argument(
        "--number-headings", action="store_true", dest="number_headings", default=False,
        help="Add auto-numbering to headings (1, 1.1, 1.1.1, etc.)"
    )
    parser.add_argument(
        "--page-break", action="store_true", dest="page_break_h1", default=False,
        help="Add page break before each H1 heading"
    )
    parser.add_argument(
        "--list-styles", action="store_true", help="List detected guide styles in a template"
    )
    parser.add_argument(
        "--create-template",
        type=str,
        default=None,
        metavar="PATH",
        help="Generate a sample template (use --theme to choose style)",
    )
    parser.add_argument(
        "--theme",
        type=str,
        default="official",
        choices=theme_names,
        help="Template theme (use --list-themes to see all)",
    )
    parser.add_argument(
        "--list-themes", action="store_true", help="List available template themes"
    )
    parser.add_argument("--version", action="store_true", help="Show version")
    parser.add_argument(
        "--validate-template", type=str, default=None, metavar="PATH",
        help="Check a template for missing required guide paragraphs"
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    if sys.stdout.encoding and sys.stdout.encoding.upper() in ("GBK", "GB2312", "GB18030"):
        sys.stdout.reconfigure(errors="replace")
    args = parse_args(argv)

    if args.version:
        print(f"md2word v{__version__}")
        return 0

    if args.list_themes:
        print("Available template themes:\n")
        for name, spec in _list_themes():
            print(f"  {name:<12} {spec.label:<14} {spec.filename:<16}  {spec.desc}")
        return 0

    if args.create_template:
        _create_template(Path(args.create_template), theme=args.theme)
        return 0

    if args.validate_template:
        from .template import validate_template as _validate
        result = _validate(args.validate_template)
        print(f"Template: {args.validate_template}")
        print(f"  Found: {', '.join(result['found']) or '(none)'}")
        if result["missing_required"]:
            print(f"  ❌ MISSING (required): {', '.join(result['missing_required'])}")
        else:
            print(f"  ✅ All required slots present")
        if result["missing_recommended"]:
            print(f"  ⚠️  MISSING (recommended): {', '.join(result['missing_recommended'])}")
        return 1 if result["missing_required"] else 0

    if args.list_styles:
        tpl = args.template or _find_default_template()
        if not tpl:
            print("No template specified and no default found.", file=sys.stderr)
            return 1
        items = list_template_styles(tpl)
        if not items:
            print(f"No guide paragraphs found in {tpl}")
            return 0
        print(f"Styles in: {tpl}\n")
        print(f"{'Slot':<14} {'Guide Text':<12} {'Font':<18} {'EA Font':<14} {'Size':<7} {'Bold':<6}")
        print("-" * 80)
        for item in items:
            print(
                f"{item['slot']:<14} {item['guide_text'][:10]:<12} "
                f"{item['font']:<18} {item['ea_font']:<14} "
                f"{str(item['size_pt']):<7} {'Yes' if item['bold'] else 'No':<6}"
            )
        return 0

    # ── Batch conversion ────────────────────────────────────────────────────
    tpl = args.template or _find_default_template(args.theme)
    if not tpl:
        print("No template specified and no default found.", file=sys.stderr)
        print("Generate one: md2word --create-template template.docx")
        return 1

    # Determine TOC
    use_toc = True
    if args.no_toc:
        use_toc = False
    elif args.toc is not None:
        use_toc = args.toc

    inputs = args.inputs
    if inputs:
        ok = True
        for i, in_str in enumerate(inputs):
            in_path = Path(in_str)
            if not in_path.exists():
                print(f"Input file not found: {in_path}", file=sys.stderr)
                ok = False
                continue
            md_text = in_path.read_text(encoding="utf-8")
            out_path = (
                Path(args.output) if args.output and len(inputs) == 1
                else in_path.with_suffix(".docx")
            )
            if i > 0:
                print()
            print(f"[{i+1}/{len(inputs)}] Converting: {in_path}")
            print(f"  Template: {tpl}")
            print(f"  Output:   {out_path}")
            convert(
                md_text, tpl, out_path,
                image_max_width=args.image_width,
                toc=use_toc,
                toc_depth=args.toc_depth,
                highlight_enabled=not args.no_highlight,
                math_enabled=not args.no_math,
                mermaid_enabled=not args.no_mermaid,
                number_headings=args.number_headings,
                page_break_h1=args.page_break_h1,
            )
            print(f"  Done → {out_path}")
        return 0 if ok else 1

    # Single conversion from stdin
    md_text = sys.stdin.read()
    if not args.output:
        print("Output path required when reading from stdin (use -o)", file=sys.stderr)
        return 1
    out_path = Path(args.output)

    print(f"Converting: (stdin)")
    print(f"Template:   {tpl}")
    print(f"Output:     {out_path}")

    convert(
        md_text, tpl, out_path,
        image_max_width=args.image_width,
        toc=use_toc,
        toc_depth=args.toc_depth,
        highlight_enabled=not args.no_highlight,
        math_enabled=not args.no_math,
        mermaid_enabled=not args.no_mermaid,
        number_headings=args.number_headings,
        page_break_h1=args.page_break_h1,
    )
    print(f"Done → {out_path}")
    return 0


def _find_default_template(theme: str | None = None) -> Path | None:
    """Find default template — prefers theme-named files in template/."""
    if theme:
        spec = get_theme(theme)
        if spec:
            for base in [Path("template"), Path(__file__).parent.parent.parent / "template"]:
                candidate = base / spec.filename
                if candidate.exists():
                    return candidate

    search = [
        Path("template/template1.docx"),
        Path("template/官方公文.docx"),
        Path("template/学术论文.docx"),
        Path("template/技术文档.docx"),
        Path("template/自媒体排版.docx"),
    ]
    # Also look relative to package
    pkg_tpl = Path(__file__).parent.parent.parent / "template"
    search.extend(pkg_tpl / n for n in [
        "template1.docx", "官方公文.docx", "学术论文.docx",
        "技术文档.docx", "自媒体排版.docx",
    ])
    for c in search:
        if c.exists():
            return c
    return None


if __name__ == "__main__":
    sys.exit(main())
