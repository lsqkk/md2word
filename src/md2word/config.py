"""Configuration file support — md2word.yaml / md2word.json / pyproject.toml."""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Any

# ── helpers ──────────────────────────────────────────────────────────────────

_TRUE_VALUES = {"1", "yes", "true", "on", "y", "t"}


def _to_bool(v: str) -> bool:
    return v.strip().lower() in _TRUE_VALUES


def _parse_yaml_line(line: str) -> tuple[str, Any] | None:
    """Parse a single ``key: value`` YAML-ish line."""
    stripped = line.strip()
    if not stripped or stripped.startswith("#"):
        return None
    m = re.match(r"^(\w[\w-]*)\s*:\s*(.*?)\s*$", stripped)
    if not m:
        return None
    key = m.group(1).replace("-", "_")
    raw = m.group(2)
    if not raw or raw.lower() in ("null", "none", "~"):
        return key, None
    if raw.lower() in _TRUE_VALUES | {"false", "no", "off", "n", "f", "0"}:
        return key, _to_bool(raw)
    try:
        return key, int(raw)
    except ValueError:
        pass
    try:
        return key, float(raw)
    except ValueError:
        pass
    return key, raw.strip("\"'")


def _parse_yaml_map_section(lines: list[str], start_idx: int) -> tuple[dict[str, str], int]:
    """Parse a YAML mapping section (indented key: value pairs).

    Returns (parsed_dict, next_line_index).
    """
    result: dict[str, str] = {}
    i = start_idx
    while i < len(lines):
        line = lines[i]
        stripped = line.rstrip()
        if not stripped or stripped.startswith("#"):
            i += 1
            continue
        # Check indentation level — a non-indented key ends the map
        if not line.startswith(" ") and not line.startswith("\t"):
            break
        m = re.match(r"^\s+(\w[\w-]*)\s*:\s*(.*?)\s*$", stripped)
        if m:
            key = m.group(1).replace("-", "_")
            val = m.group(2).strip("\"'")
            result[key] = val
        i += 1
    return result, i


# ── search path ──────────────────────────────────────────────────────────────


def find_config(start: Path | None = None) -> Path | None:
    """Walk up from *start* (default CWD) looking for a config file.

    Priority:
    1. ``.md2word/config.yaml`` (project directory)
    2. ``md2word.yaml`` > ``md2word.yml`` > ``md2word.json``
    3. ``[tool.md2word]`` in ``pyproject.toml``
    """
    root = start or Path.cwd()
    for parent in [root] + list(root.parents):
        md2word_dir = parent / ".md2word"
        if md2word_dir.is_dir():
            for name in ("config.yaml", "config.yml", "config.json"):
                candidate = md2word_dir / name
                if candidate.is_file():
                    return candidate
        for name in ("md2word.yaml", "md2word.yml", "md2word.json"):
            candidate = parent / name
            if candidate.is_file():
                return candidate
    for parent in [root] + list(root.parents):
        candidate = parent / "pyproject.toml"
        if candidate.is_file():
            return candidate
    return None


# ── loaders ──────────────────────────────────────────────────────────────────


def _load_yaml_text(text: str) -> dict[str, Any]:
    """Minimal YAML key-value parser (no PyYAML needed)."""
    cfg: dict[str, Any] = {}
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        parsed = _parse_yaml_line(line)
        if parsed:
            key, val = parsed
            cfg[key] = val
            # Check if value contains nested mapping (style-map)
            if val is not None and isinstance(val, str) and val.strip() == "":
                # This might be a mapping key — peek ahead
                pass
        i += 1
    return cfg


def _load_yaml(path: Path) -> dict[str, Any]:
    return _load_yaml_text(path.read_text(encoding="utf-8"))


def _load_json(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    return {k.replace("-", "_"): v for k, v in data.items()}


def _load_pyproject(path: Path) -> dict[str, Any]:
    """Read ``[tool.md2word]`` from pyproject.toml."""
    text = path.read_text(encoding="utf-8")
    m = re.search(
        r'\[tool\.md2word\](.*?)(?=\n\[|$)',
        text,
        re.DOTALL,
    )
    if not m:
        return {}
    section = m.group(1)
    cfg: dict[str, Any] = {}
    for line in section.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        m2 = re.match(r"""^(\w[\w-]*)\s*=\s*(.+?)\s*$""", stripped)
        if not m2:
            continue
        key = m2.group(1).replace("-", "_")
        raw = m2.group(2)
        if raw.lower() in _TRUE_VALUES | {"false", "no", "off"}:
            cfg[key] = _to_bool(raw)
        else:
            try:
                cfg[key] = int(raw)
            except ValueError:
                try:
                    cfg[key] = float(raw)
                except ValueError:
                    cfg[key] = raw.strip("\"'")
    return cfg


# ── public API ───────────────────────────────────────────────────────────────


def load_config(start: Path | None = None) -> dict[str, Any]:
    """Load configuration from the nearest config file.

    Returns a flat dict with keys matching CLI argument names (underscored).
    Returns an empty dict if no config file exists.
    """
    path = find_config(start)
    if path is None:
        return {}

    try:
        suffix = path.suffix.lower()
        if suffix == ".json":
            cfg = _load_json(path)
        elif suffix == ".toml":
            cfg = _load_pyproject(path)
        else:
            cfg = _load_yaml(path)
    except Exception as exc:
        print(f"  [WARN] Failed to load config {path}: {exc}", file=sys.stderr)
        return {}

    if path.parent.name == ".md2word":
        cfg["_project_dir"] = str(path.parent.parent.resolve())
    elif path.name == "pyproject.toml":
        cfg["_project_dir"] = str(path.parent.resolve())

    _BOOLEAN_KEYS: set[str] = {
        "toc", "number_headings", "page_break",
        "highlight", "math", "mermaid", "no_highlight",
        "no_math", "no_mermaid",
    }

    for k, v in list(cfg.items()):
        if k in _BOOLEAN_KEYS and isinstance(v, str):
            cfg[k] = _to_bool(v)
        if k == "theme" and isinstance(v, str):
            cfg["_theme_hint"] = v

    return cfg


def merge_with_args(
    cfg: dict[str, Any], args: dict[str, Any]
) -> dict[str, Any]:
    """Merge config dict with CLI args dict. CLI args take priority."""
    result = dict(args)
    for k, v in cfg.items():
        if k in result:
            if result[k] is None:
                result[k] = v
        else:
            result[k] = v
    return result
