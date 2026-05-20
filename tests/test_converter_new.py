"""Tests for new converter features: incremental cache, GB check, red-head, page numbers."""

from pathlib import Path

from md2word.converter import _file_hash, _load_cache, _save_cache


class TestFileHash:
    def test_consistent_hash(self, tmp_path):
        f = tmp_path / "test.md"
        f.write_text("hello world", encoding="utf-8")
        assert _file_hash(f) == _file_hash(f)

    def test_different_content_different_hash(self, tmp_path):
        f1 = tmp_path / "a.md"
        f2 = tmp_path / "b.md"
        f1.write_text("content a", encoding="utf-8")
        f2.write_text("content b", encoding="utf-8")
        assert _file_hash(f1) != _file_hash(f2)


class TestCachePersistence:
    def test_roundtrip(self, tmp_path):
        cache_path = tmp_path / ".cache.json"
        data = {"file1.md": "abc123", "file2.md": "def456"}
        _save_cache(cache_path, data)
        loaded = _load_cache(cache_path)
        assert loaded == data

    def test_missing_cache(self, tmp_path):
        assert _load_cache(tmp_path / "nonexistent.json") == {}

    def test_corrupted_cache(self, tmp_path):
        f = tmp_path / "bad.json"
        f.write_text("not json", encoding="utf-8")
        assert _load_cache(f) == {}
