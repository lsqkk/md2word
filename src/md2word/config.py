"""Configuration file support — md2word.yaml / md2word.json / pyproject.toml.

Supports YAML (primary, via PyYAML), JSON, and TOML (pyproject.toml)
config formats.  PyYAML is preferred for nested structures like
``style_map``; a minimal fallback parser handles simple key-value
configs when PyYAML is not available.
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Any

from .options import ConvertOptions

# ── helpers ──────────────────────────────────────────────────────────────────

_TRUE_VALUES = {"1", "yes", "true", "on", "y", "t"}


def _to_bool(v: str) -> bool:
    return v.strip().lower() in _TRUE_VALUES


# ── PyYAML wrapper ───────────────────────────────────────────────────────────

_HAS_PYYAML = False
try:
    import yaml as _yaml  # type: ignore[import-untyped]
    _HAS_PYYAML = True
except ImportError:
    pass


def _yaml_load(text: str) -> dict[str, Any] | None:
    """Try to parse *text* with PyYAML; return ``None`` if unavailable."""
    if not _HAS_PYYAML:
        return None
    try:
        data = _yaml.safe_load(text)
        if isinstance(data, dict):
            return {k.replace("-", "_"): v for k, v in data.items()}
        return {}
    except Exception:
        return None


# ── Minimal fallback YAML parser (flat key-value only) ───────────────────


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


def _load_yaml_text_fallback(text: str) -> dict[str, Any]:
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
        i += 1
    return cfg


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


def _load_yaml(path: Path) -> dict[str, Any]:
    """Load a YAML file, trying PyYAML first, then the fallback parser."""
    text = path.read_text(encoding="utf-8")
    result = _yaml_load(text)
    if result is not None:
        return _normalize_yaml(result)
    return _load_yaml_text_fallback(text)


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


def _normalize_yaml(data: dict) -> dict[str, Any]:
    """Normalize YAML-loaded dict: convert list values, fix key naming."""
    result: dict[str, Any] = {}
    for k, v in data.items():
        key = k.replace("-", "_")
        if isinstance(v, list):
            # e.g. style_map might be passed as a list of pairs, but
            # dict is preferred — store as-is for the consumer to handle
            result[key] = v
        elif isinstance(v, dict):
            result[key] = v
        elif isinstance(v, str):
            result[key] = v
        elif isinstance(v, bool):
            result[key] = v
        elif isinstance(v, (int, float)):
            result[key] = v
        elif v is None:
            result[key] = None
    return result


# ── public API ───────────────────────────────────────────────────────────────


def load_config(start: Path | None = None, validate: bool = True) -> dict[str, Any]:
    """Load configuration from the nearest config file.

    Returns a flat dict with keys matching CLI argument names (underscored).
    Returns an empty dict if no config file exists.

    When *validate* is ``True`` (default), unknown keys emit a warning.
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

    # Normalise boolean string values
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

    # Validate config keys against ConvertOptions fields
    if validate:
        _validate_config_keys(cfg, path)

    return cfg


def _validate_config_keys(cfg: dict[str, Any], path: Path) -> None:
    """Warn about unknown config keys."""
    valid_keys = set(ConvertOptions.__dataclass_fields__)
    valid_keys.update({
        "_project_dir", "_theme_hint", "output", "template",
        "out_dir", "image_width", "no_toc", "no_highlight", "no_math",
        "no_mermaid", "no_footnotes", "page_break", "redhead",
        "page_number", "incremental", "project_dir", "config",
        "theme", "watch",
    })
    for k in cfg:
        if k.startswith("_"):
            continue
        if k not in valid_keys:
            print(f"  [WARN] 未知配置键 '{k}' 在 {path}", file=sys.stderr)


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
