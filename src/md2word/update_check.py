"""Periodic update check against GitHub releases.

Checks the latest release tag from ``lsqqk/md2word`` on GitHub and compares
it with the installed version.  Results are cached locally for 24 h to
avoid hitting API rate limits on every conversion.
"""

from __future__ import annotations

import json
import time
from dataclasses import dataclass
from pathlib import Path
from typing import NamedTuple

import requests

# ── Constants ──────────────────────────────────────────────────────────────

_CACHE_DIR = Path.home() / ".md2word"
_CACHE_FILE = _CACHE_DIR / "update_cache.json"
_CACHE_TTL = 86_400  # 24 hours

_GITHUB_API = "https://api.github.com/repos/lsqkk/md2word/releases/latest"
_HEADERS = {
    "Accept": "application/vnd.github.v3+json",
    "User-Agent": "md2word",
}
_REQUEST_TIMEOUT = 5  # seconds


# ── Public types ────────────────────────────────────────────────────────────


class UpdateInfo(NamedTuple):
    """Result of an update availability check."""

    latest_version: str | None
    """The latest version available on GitHub, or ``None`` if the check
    could not be performed."""

    has_update: bool
    """``True`` when *latest_version* is newer than the installed version."""

    error: str | None
    """Human-readable error description, or ``None`` on success."""


# ── Version helpers ────────────────────────────────────────────────────────


def parse_version(text: str) -> str | None:
    """Extract ``__version__`` from Python source text."""
    import re

    m = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', text)
    return m.group(1) if m else None


def compare_versions(v1: str, v2: str) -> int:
    """Compare two semver strings.

    Returns:
        -1 if *v1* < *v2*, 0 if equal, 1 if *v1* > *v2*.
    """
    parts1 = _split(v1)
    parts2 = _split(v2)
    max_len = max(len(parts1), len(parts2))
    parts1 += [0] * (max_len - len(parts1))
    parts2 += [0] * (max_len - len(parts2))
    for a, b in zip(parts1, parts2):
        if a < b:
            return -1
        if a > b:
            return 1
    return 0


def _split(v: str) -> list[int]:
    """Split a version string into integer parts, ignoring non-numeric suffixes."""
    result = []
    for part in v.split("."):
        num = ""
        for ch in part:
            if ch.isdigit():
                num += ch
            else:
                break
        result.append(int(num) if num else 0)
    return result


# ── Cache ──────────────────────────────────────────────────────────────────


def _load_cache() -> dict:
    try:
        if _CACHE_FILE.exists():
            return json.loads(_CACHE_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        pass
    return {}


def _save_cache(data: dict) -> None:
    try:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        _CACHE_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass


# ── Network ────────────────────────────────────────────────────────────────


def fetch_latest_version() -> str | None:
    """Fetch the latest version tag from GitHub Releases.

    Returns the version string (e.g. ``"1.9.0"``) or ``None`` on failure.
    """
    try:
        resp = requests.get(_GITHUB_API, headers=_HEADERS, timeout=_REQUEST_TIMEOUT)
        if resp.status_code == 200:
            tag = resp.json().get("tag_name", "")
            return tag.lstrip("v")
    except (requests.RequestException, ValueError, KeyError):
        pass

    return None


def check_for_update(
    current_version: str, *, force: bool = False
) -> UpdateInfo:
    """Check whether a newer version of md2word is available.

    Results are cached for 24 h (configurable via ``_CACHE_TTL``).  Pass
    ``force=True`` to bypass the cache.

    The call is safe — all network and parse errors are caught and reported
    in the returned ``UpdateInfo.error`` field.
    """
    cache = _load_cache()
    now = time.time()

    # ── Return cached result if still fresh ──────────────────────────────
    if not force:
        cached_latest = cache.get("latest_version")
        checked_at = cache.get("checked_at", 0)
        if cached_latest and (now - checked_at) < _CACHE_TTL:
            latest: str = cached_latest
            return UpdateInfo(
                latest_version=latest,
                has_update=compare_versions(current_version, latest) < 0,
                error=None,
            )

    # ── Fetch from GitHub ────────────────────────────────────────────────
    latest = fetch_latest_version()
    if latest is None:
        stale = cache.get("latest_version")
        if stale:
            return UpdateInfo(
                latest_version=stale,
                has_update=compare_versions(current_version, stale) < 0,
                error="无法连接 GitHub，使用缓存数据",
            )
        return UpdateInfo(
            latest_version=None,
            has_update=False,
            error="无法检查更新（网络不可用或 GitHub 限速）",
        )

    _save_cache({"latest_version": latest, "checked_at": now})
    return UpdateInfo(
        latest_version=latest,
        has_update=compare_versions(current_version, latest) < 0,
        error=None,
    )


# ── Formatting ─────────────────────────────────────────────────────────────


def format_update_message(info: UpdateInfo, current_version: str) -> str | None:
    """Return a human-readable update notice, or ``None`` if none needed."""
    if info.has_update and info.latest_version:
        return (
            f"📦 新版本可用: v{current_version} → v{info.latest_version}\n"
            f"   更新: pip install --upgrade md2word\n"
            f"   https://github.com/lsqkk/md2word/releases"
        )
    return None
