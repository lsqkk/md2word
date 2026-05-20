"""CLI entry point for md2word."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .config import load_config, merge_with_args
from .converter import convert
from .template import list_template_styles, validate_template
from .themes import build_theme, get_theme, list_themes as _list_themes


def _check_deps() -> dict[str, bool]:
    """Check availability of optional dependencies."""
    status: dict[str, bool] = {}

    try:
        import pygments  # noqa: F401
        status["highlight"] = True
    except ImportError:
        status["highlight"] = False

    try:
        import matplotlib  # noqa: F401
        import latex2mathml  # noqa: F401
        status["math"] = True
    except ImportError:
        status["math"] = False

    try:
        import resvg  # noqa: F401
        status["svg"] = True
    except ImportError:
        status["svg"] = False

    try:
        import yaml  # noqa: F401
        status["yaml"] = True
    except ImportError:
        status["yaml"] = False

    try:
        import watchdog  # noqa: F401
        status["watch"] = True
    except ImportError:
        status["watch"] = False

    return status


def _check_deps_cmd() -> int:
    """``--check-deps`` sub-command."""
    deps = _check_deps()
    print("md2word optional dependency status:\n")
    rows = [
        ("Code highlight", "pygments", deps["highlight"]),
        ("Math formulas", "matplotlib + latex2mathml", deps["math"]),
        ("SVG parsing", "resvg", deps["svg"]),
        ("YAML config", "pyyaml", deps["yaml"]),
        ("Watch mode", "watchdog", deps["watch"]),
    ]
    for label, pkg, ok in rows:
        icon = "✅" if ok else "❌"
        print(f"  {icon} {label:<20} {pkg}")
    missing = [pkg for _, pkg, ok in rows if not ok]
    if missing:
        print("\nInstall missing:  pip install md2word[all]")
    return 0 if all(ok for _, _, ok in rows) else 0


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
        epilog="See https://github.com/lsqkk/md2word for full documentation.",
    )

    parser.add_argument(
        "inputs", nargs="*", type=str, default=[],
        help="Input Markdown file(s) (omit to read stdin)"
    )
    parser.add_argument("-o", "--output", type=str, default=None,
                        help="Output .docx path")
    parser.add_argument(
        "-t", "--template", type=str, default=None,
        help="Template .docx with guide paragraphs"
    )
    parser.add_argument(
        "--image-width", type=float, default=None,
        help="Max image width in inches (default: 5.5)"
    )
    parser.add_argument(
        "--toc", action="store_true", default=None, help="Generate table of contents"
    )
    parser.add_argument(
        "--no-toc", action="store_true", dest="no_toc", default=None,
        help="Skip table of contents"
    )
    parser.add_argument(
        "--toc-depth", type=str, default=None,
        help="TOC heading depth range (default: 1-3)"
    )
    parser.add_argument(
        "--no-highlight", action="store_true", dest="no_highlight", default=None,
        help="Disable code syntax highlighting"
    )
    parser.add_argument(
        "--no-math", action="store_true", dest="no_math", default=None,
        help="Disable math formula rendering ($...$ / $$...$$)"
    )
    parser.add_argument(
        "--no-mermaid", action="store_true", dest="no_mermaid", default=None,
        help="Disable mermaid diagram rendering"
    )
    parser.add_argument(
        "--number-headings", action="store_true", dest="number_headings", default=None,
        help="Add auto-numbering to headings (1, 1.1, 1.1.1, etc.)"
    )
    parser.add_argument(
        "--page-break", action="store_true", dest="page_break_h1", default=None,
        help="Add page break before each H1 heading"
    )
    parser.add_argument(
        "--list-styles", action="store_true",
        help="List detected guide styles in a template"
    )
    parser.add_argument(
        "--create-template", type=str, default=None, metavar="PATH",
        help="Generate a sample template (use --theme to choose style)",
    )
    parser.add_argument(
        "--theme", type=str, default=None, choices=theme_names,
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
    parser.add_argument(
        "--watch", action="store_true", default=None,
        help="Watch directory for changes and auto-convert"
    )
    parser.add_argument(
        "--config", type=str, default=None, metavar="PATH",
        help="Path to config file (default: auto-detect md2word.yaml)"
    )
    parser.add_argument(
        "--check-deps", action="store_true",
        help="Check availability of optional dependencies"
    )
    parser.add_argument(
        "--three-line-table", action="store_true", default=None,
        help="Use academic three-line table style (三线表)"
    )
    parser.add_argument(
        "--no-footnotes", action="store_true", default=None,
        help="Disable footnote processing"
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    if sys.stdout.encoding and sys.stdout.encoding.upper() in ("GBK", "GB2312", "GB18030"):
        sys.stdout.reconfigure(errors="replace")

    raw_args = parse_args(argv)
    args = vars(raw_args)

    # ── Version ────────────────────────────────────────────────────────────
    if args["version"]:
        print(f"md2word v{__version__}")
        return 0

    # ── Check deps ─────────────────────────────────────────────────────────
    if args["check_deps"]:
        return _check_deps_cmd()

    # ── Load config ────────────────────────────────────────────────────────
    cfg = {}
    if args.get("config"):
        from .config import load_config as _load
        try:
            cfg = _load(Path(args["config"]).parent)
        except Exception as e:
            print(f"  [ERROR] Failed to load config '{args['config']}': {e}",
                  file=sys.stderr)
            return 1
    else:
        cfg = load_config()

    # ── Merge: CLI args override config ────────────────────────────────────
    merged = merge_with_args(cfg, args)
    raw_args = argparse.Namespace(**merged)

    # ── List themes ────────────────────────────────────────────────────────
    if raw_args.list_themes:
        print("Available template themes:\n")
        for name, spec in _list_themes():
            print(f"  {name:<12} {spec.label:<14} {spec.filename:<16}  {spec.desc}")
        return 0

    # ── Create template ────────────────────────────────────────────────────
    if raw_args.create_template:
        theme = raw_args.theme or "official"
        _create_template(Path(raw_args.create_template), theme=theme)
        return 0

    # ── Validate template ──────────────────────────────────────────────────
    if raw_args.validate_template:
        result = validate_template(raw_args.validate_template)
        print(f"Template: {raw_args.validate_template}")
        print(f"  Found: {', '.join(result['found']) or '(none)'}")
        if result["missing_required"]:
            print(f"  ❌ MISSING (required): {', '.join(result['missing_required'])}")
        else:
            print(f"  ✅ All required slots present")
        if result["missing_recommended"]:
            print(f"  ⚠️  MISSING (recommended): {', '.join(result['missing_recommended'])}")
        return 1 if result["missing_required"] else 0

    # ── List styles ────────────────────────────────────────────────────────
    if raw_args.list_styles:
        tpl = _resolve_template(raw_args.template, raw_args.theme)
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

    # ── Resolve template ───────────────────────────────────────────────────
    tpl = _resolve_template(raw_args.template, raw_args.theme)
    if not tpl:
        print("No template specified and no default found.", file=sys.stderr)
        print("Generate one:  md2word --create-template template.docx")
        return 1

    # ── Determine TOC ──────────────────────────────────────────────────────
    use_toc = True
    if raw_args.no_toc is True:
        use_toc = False
    elif raw_args.toc is not None:
        use_toc = raw_args.toc
    elif "toc" in cfg:
        use_toc = bool(cfg["toc"])

    # ── Build conversion kwargs ────────────────────────────────────────────
    conv_kwargs = {
        "image_max_width": raw_args.image_width or 5.5,
        "toc": use_toc,
        "toc_depth": raw_args.toc_depth or "1-3",
        "highlight_enabled": not (raw_args.no_highlight or False),
        "math_enabled": not (raw_args.no_math or False),
        "mermaid_enabled": not (raw_args.no_mermaid or False),
        "number_headings": raw_args.number_headings or False,
        "page_break_h1": raw_args.page_break_h1 or False,
        "three_line_table": raw_args.three_line_table or False,
        "footnotes_enabled": not (raw_args.no_footnotes or False),
    }

    # ── Watch mode ─────────────────────────────────────────────────────────
    if raw_args.watch:
        _watch_mode(raw_args.inputs, tpl, conv_kwargs)
        return 0

    # ── Batch conversion ───────────────────────────────────────────────────
    inputs = raw_args.inputs
    if inputs:
        ok = True
        for i, in_str in enumerate(inputs):
            in_path = Path(in_str)
            if not in_path.exists():
                print(f"  ❌ Input file not found: {in_path}", file=sys.stderr)
                ok = False
                continue
            md_text = in_path.read_text(encoding="utf-8")
            out_path = (
                Path(raw_args.output) if raw_args.output and len(inputs) == 1
                else in_path.with_suffix(".docx")
            )
            if i > 0:
                print()
            print(f"[{i+1}/{len(inputs)}] Converting: {in_path}")
            print(f"  Template: {tpl}")
            print(f"  Output:   {out_path}")
            convert(md_text, tpl, out_path, **conv_kwargs)
        return 0 if ok else 1

    # ── Stdin ──────────────────────────────────────────────────────────────
    md_text = sys.stdin.read()
    if not raw_args.output:
        print("  ❌ Output path required when reading from stdin (use -o)",
              file=sys.stderr)
        return 1
    out_path = Path(raw_args.output)
    print(f"Converting: (stdin)")
    print(f"Template:   {tpl}")
    print(f"Output:     {out_path}")
    convert(md_text, tpl, out_path, **conv_kwargs)
    return 0


# ── helpers ──────────────────────────────────────────────────────────────────


def _resolve_template(template_arg: str | None, theme: str | None) -> Path | None:
    """Resolve template path from user arg / config theme hint / default search."""
    pkg_tpl = Path(__file__).parent.parent.parent / "template"

    if template_arg:
        p = Path(template_arg)
        if p.exists():
            return p
        # Also try relative to package template dir (supports config files elsewhere)
        pkg_candidate = pkg_tpl / p.name
        if pkg_candidate.exists():
            return pkg_candidate
        # Not an error if it's from config — we'll try defaults below
        if p.is_absolute():
            print(f"  [WARN] Template not found: {p}", file=sys.stderr)
            return None

    # Try theme -> filename mapping
    if theme:
        spec = get_theme(theme)
        if spec:
            for base in [Path("template"), pkg_tpl]:
                candidate = base / spec.filename
                if candidate.exists():
                    return candidate

    # Fallback: search known paths
    search: list[Path] = []
    for name in ["template1.docx", "官方公文.docx", "学术论文.docx",
                  "技术文档.docx", "自媒体排版.docx"]:
        search.append(Path(f"template/{name}"))
    search += [pkg_tpl / name for name in
               ["template1.docx", "官方公文.docx", "学术论文.docx",
                "技术文档.docx", "自媒体排版.docx"]]
    for c in search:
        if c.exists():
            return c
    return None


def _watch_mode(
    inputs: list[str],
    template_path: Path,
    conv_kwargs: dict,
) -> None:
    """Watch input file(s) or directory and auto-convert on changes."""
    watch_paths: list[Path] = []
    if inputs:
        watch_paths = [Path(p) for p in inputs]
    else:
        watch_paths = [Path.cwd()]

    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
        _HAS_WATCHDOG = True
    except ImportError:
        _HAS_WATCHDOG = False

    if _HAS_WATCHDOG:
        _watch_with_watchdog(watch_paths, template_path, conv_kwargs)
    else:
        _watch_polling(watch_paths, template_path, conv_kwargs)


def _watch_with_watchdog(
    paths: list[Path],
    template_path: Path,
    conv_kwargs: dict,
) -> None:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler

    class _Handler(FileSystemEventHandler):
        def on_modified(self, event):
            if event.src_path.endswith(".md"):
                _convert_if_md(Path(event.src_path), template_path, conv_kwargs)

    observer = Observer()
    for p in paths:
        target = p if p.is_dir() else p.parent
        observer.schedule(_Handler(), str(target), recursive=False)
    print(f"  👁️  Watching {len(paths)} path(s) for .md changes (Ctrl+C to stop)")
    try:
        observer.start()
        import time
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


def _watch_polling(
    paths: list[Path],
    template_path: Path,
    conv_kwargs: dict,
) -> None:
    import time

    # Gather initial files
    targets: dict[Path, float] = {}
    for p in paths:
        if p.is_dir():
            for f in p.iterdir():
                if f.suffix == ".md":
                    targets[f] = f.stat().st_mtime
        elif p.suffix == ".md":
            targets[p] = p.stat().st_mtime

    print(f"  👁️  Watching {len(targets)} .md file(s) for changes (Ctrl+C to stop)")
    try:
        while True:
            changed = False
            for p in list(targets):
                try:
                    new_mtime = p.stat().st_mtime
                    if new_mtime > targets[p]:
                        targets[p] = new_mtime
                        _convert_if_md(p, template_path, conv_kwargs)
                        changed = True
                except FileNotFoundError:
                    targets.pop(p, None)
            if not changed:
                time.sleep(2)
    except KeyboardInterrupt:
        print("  Watch stopped.")


def _convert_if_md(
    path: Path,
    template_path: Path,
    conv_kwargs: dict,
) -> None:
    if path.suffix.lower() != ".md":
        return
    try:
        md_text = path.read_text(encoding="utf-8")
        out_path = path.with_suffix(".docx")
        print(f"\n  🔄 Changed: {path.name}")
        convert(md_text, template_path, out_path, **conv_kwargs)
        print(f"  ✅ {out_path.name}")
    except Exception as e:
        print(f"  ❌ Failed: {e}")


if __name__ == "__main__":
    sys.exit(main())
