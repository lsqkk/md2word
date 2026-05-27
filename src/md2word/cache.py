"""Incremental conversion cache — MD5-based content hash tracking."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path


def file_hash(path: str | Path) -> str:
    """Return MD5 hex digest of file content."""
    h = hashlib.md5()
    h.update(Path(path).read_bytes())
    return h.hexdigest()


def load_cache(cache_path: str | Path) -> dict:
    """Load incremental conversion cache."""
    try:
        return json.loads(Path(cache_path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_cache(cache_path: str | Path, cache: dict) -> None:
    """Save incremental conversion cache."""
    Path(cache_path).write_text(
        json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8"
    )
