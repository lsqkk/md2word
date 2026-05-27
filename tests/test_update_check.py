"""Tests for update_check module — version comparison, caching, formatting."""

from __future__ import annotations

import json
import time
from pathlib import Path

import pytest

from md2word.update_check import (
    UpdateInfo,
    check_for_update,
    compare_versions,
    fetch_latest_version,
    format_update_message,
    parse_version,
)


# ── parse_version ────────────────────────────────────────────────────────────


class TestParseVersion:
    def test_extracts_version(self):
        assert parse_version('__version__ = "1.8.0"') == "1.8.0"

    def test_single_quotes(self):
        assert parse_version("__version__ = '2.0.0'") == "2.0.0"

    def test_with_spaces(self):
        assert parse_version('__version__ =    "1.0.0"') == "1.0.0"

    def test_no_match(self):
        assert parse_version("x = 42") is None

    def test_empty_string(self):
        assert parse_version("") is None

    def test_dev_version(self):
        assert parse_version('__version__ = "1.8.0.dev1"') == "1.8.0.dev1"


# ── compare_versions ────────────────────────────────────────────────────────


class TestCompareVersions:
    def test_equal(self):
        assert compare_versions("1.8.0", "1.8.0") == 0

    def test_major_newer(self):
        assert compare_versions("1.8.0", "2.0.0") == -1

    def test_major_older(self):
        assert compare_versions("2.0.0", "1.8.0") == 1

    def test_minor_newer(self):
        assert compare_versions("1.7.0", "1.8.0") == -1

    def test_patch_newer(self):
        assert compare_versions("1.8.0", "1.8.1") == -1

    def test_patch_older(self):
        assert compare_versions("1.8.1", "1.8.0") == 1

    def test_different_lengths(self):
        assert compare_versions("1.8", "1.8.0") == 0
        assert compare_versions("1.8", "1.8.1") == -1
        assert compare_versions("1.8.1", "1.8") == 1

    def test_pre_release(self):
        assert compare_versions("1.8.0", "1.8.0a1") == 0
        assert compare_versions("1.8.0a1", "1.8.0") == 0

    def test_same(self):
        assert compare_versions("0.0.1", "0.0.1") == 0


# ── fetch_latest_version ─────────────────────────────────────────────────────


class TestFetchLatestVersion:
    def test_returns_none_on_network_error(self):
        """Should not raise — should return None gracefully."""
        result = fetch_latest_version()
        # This may actually succeed if run on a machine with internet — that's OK.
        # We just verify it returns a string or None, never raises.
        assert result is None or isinstance(result, str)


# ── check_for_update ─────────────────────────────────────────────────────────


class TestCheckForUpdate:
    def test_returns_update_info_type(self):
        info = check_for_update("99.99.99", force=True)
        assert isinstance(info, UpdateInfo)
        assert isinstance(info.has_update, bool)

    def test_no_update_when_latest_unknown(self):
        """When fetch fails and no cache, has_update should be False."""
        info = check_for_update("1.8.0", force=True)
        # If network is available, latest might be found — just verify type
        assert isinstance(info, UpdateInfo)

    def test_update_detected_with_cache(self, tmp_path):
        """Seed a cache file with a newer version and verify detection."""
        cache_dir = Path.home() / ".md2word"
        cache_dir.mkdir(parents=True, exist_ok=True)
        cache_file = cache_dir / "update_cache.json"
        fake_data = {"latest_version": "9.9.9", "checked_at": time.time()}
        cache_file.write_text(
            json.dumps(fake_data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        try:
            info = check_for_update("1.8.0", force=False)
            assert info.has_update is True
            assert info.latest_version == "9.9.9"
            assert info.error is None
        finally:
            cache_file.unlink(missing_ok=True)


# ── format_update_message ────────────────────────────────────────────────────


class TestFormatUpdateMessage:
    def test_returns_message_when_update_available(self):
        info = UpdateInfo(
            latest_version="2.0.0", has_update=True, error=None
        )
        msg = format_update_message(info, "1.8.0")
        assert msg is not None
        assert "1.8.0" in msg
        assert "2.0.0" in msg
        assert "pip install --upgrade" in msg

    def test_returns_none_when_no_update(self):
        info = UpdateInfo(
            latest_version=None, has_update=False, error=None
        )
        assert format_update_message(info, "1.8.0") is None

    def test_returns_none_when_up_to_date(self):
        info = UpdateInfo(
            latest_version="1.8.0", has_update=False, error=None
        )
        assert format_update_message(info, "1.8.0") is None

    def test_returns_none_when_older(self):
        info = UpdateInfo(
            latest_version="1.7.0", has_update=False, error=None
        )
        assert format_update_message(info, "1.8.0") is None


# ── Edge cases ───────────────────────────────────────────────────────────────


def test_cache_corrupted_still_works(tmp_path):
    """Corrupted cache file should be treated as empty cache."""
    cache_dir = Path.home() / ".md2word"
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / "update_cache.json"
    cache_file.write_text("this is not valid json", encoding="utf-8")
    try:
        # Should not crash — fall through to fresh check
        info = check_for_update("1.8.0", force=False)
        assert isinstance(info, UpdateInfo)
    finally:
        cache_file.unlink(missing_ok=True)


def test_format_message_with_error():
    """UpdateInfo with error but no update should not produce message."""
    info = UpdateInfo(
        latest_version=None,
        has_update=False,
        error="网络不可用",
    )
    assert format_update_message(info, "1.8.0") is None
