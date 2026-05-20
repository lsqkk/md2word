"""Tests for configuration loading and merging."""

from pathlib import Path

import pytest

from md2word.config import (
    load_config,
    merge_with_args,
    find_config,
    _to_bool,
    _parse_yaml_line,
)


class TestToBool:
    def test_true_values(self):
        assert _to_bool("1") is True
        assert _to_bool("yes") is True
        assert _to_bool("true") is True
        assert _to_bool("on") is True
        assert _to_bool("Y") is True
        assert _to_bool("T") is True

    def test_false_values(self):
        assert _to_bool("0") is False
        assert _to_bool("no") is False
        assert _to_bool("false") is False
        assert _to_bool("off") is False
        assert _to_bool("n") is False
        assert _to_bool("f") is False


class TestParseYamlLine:
    def test_comment_line(self):
        assert _parse_yaml_line("# comment") is None
        assert _parse_yaml_line("  # indented comment") is None

    def test_empty_line(self):
        assert _parse_yaml_line("") is None
        assert _parse_yaml_line("   ") is None

    def test_bool_values(self):
        key, val = _parse_yaml_line("toc: true")
        assert key == "toc"
        assert val is True

        key, val = _parse_yaml_line("highlight: false")
        assert key == "highlight"
        assert val is False

    def test_int_values(self):
        key, val = _parse_yaml_line("image_width: 5")
        assert key == "image_width"
        assert val == 5

    def test_float_values(self):
        key, val = _parse_yaml_line("image_width: 5.5")
        assert key == "image_width"
        assert val == 5.5

    def test_string_values(self):
        key, val = _parse_yaml_line("template: 官方公文.docx")
        assert key == "template"
        assert val == "官方公文.docx"

    def test_null_values(self):
        key, val = _parse_yaml_line("theme: null")
        assert key == "theme"
        assert val is None

        key, val = _parse_yaml_line("theme:")
        assert key == "theme"
        assert val is None

    def test_hyphen_to_underscore(self):
        key, val = _parse_yaml_line("page-break: true")
        assert key == "page_break"
        assert val is True


class TestFindConfig:
    def test_no_config_returns_none(self, tmp_path):
        result = find_config(tmp_path)
        assert result is None

    def test_finds_yaml_in_directory(self, tmp_path):
        cfg = tmp_path / "md2word.yaml"
        cfg.write_text("toc: true", encoding="utf-8")
        result = find_config(tmp_path)
        assert result == cfg

    def test_yaml_preferred_over_json(self, tmp_path):
        (tmp_path / "md2word.yaml").write_text("toc: true", encoding="utf-8")
        (tmp_path / "md2word.json").write_text('{"toc": true}')
        result = find_config(tmp_path)
        assert result.suffix == ".yaml"

    def test_finds_yml_fallback(self, tmp_path):
        (tmp_path / "md2word.yml").write_text("toc: true", encoding="utf-8")
        result = find_config(tmp_path)
        assert result.suffix == ".yml"

    def test_dot_md2word_directory(self, tmp_path):
        """Finds config in .md2word/ project directory."""
        config_dir = tmp_path / ".md2word"
        config_dir.mkdir()
        cfg = config_dir / "config.yaml"
        cfg.write_text("toc: true", encoding="utf-8")
        result = find_config(tmp_path)
        assert result == cfg

    def test_dot_md2word_directory_deep(self, tmp_path):
        """Finds .md2word/config.yaml in parent directories."""
        subdir = tmp_path / "a" / "b" / "c"
        subdir.mkdir(parents=True)
        config_dir = tmp_path / ".md2word"
        config_dir.mkdir()
        (config_dir / "config.yaml").write_text("toc: true", encoding="utf-8")
        result = find_config(subdir)
        assert result is not None
        assert result.parent.name == ".md2word"

    def test_dot_md2word_preferred_over_flat(self, tmp_path):
        """.md2word/config.yaml takes priority over flat md2word.yaml."""
        config_dir = tmp_path / ".md2word"
        config_dir.mkdir()
        (config_dir / "config.yaml").write_text("theme: academic", encoding="utf-8")
        (tmp_path / "md2word.yaml").write_text("theme: tech", encoding="utf-8")
        result = find_config(tmp_path)
        assert result.parent.name == ".md2word"
        assert "academic" in result.read_text()

    def test_walks_up_directory_tree(self, tmp_path):
        subdir = tmp_path / "a" / "b" / "c"
        subdir.mkdir(parents=True)
        (tmp_path / "md2word.yaml").write_text("toc: true", encoding="utf-8")
        result = find_config(subdir)
        assert result is not None

    def test_pyproject_toml(self, tmp_path):
        content = """[tool.md2word]
toc = true
"""
        (tmp_path / "pyproject.toml").write_text(content, encoding="utf-8")
        result = find_config(tmp_path)
        assert result is not None
        assert result.name == "pyproject.toml"


class TestLoadConfig:
    def test_returns_empty_dict_when_no_config(self, tmp_path):
        cfg = load_config(tmp_path)
        assert cfg == {}

    def test_loads_yaml(self, tmp_path):
        (tmp_path / "md2word.yaml").write_text(
            "toc: true\nnumber_headings: yes\n",
            encoding="utf-8",
        )
        cfg = load_config(tmp_path)
        assert cfg.get("toc") is True
        assert cfg.get("number_headings") is True

    def test_loads_json(self, tmp_path):
        (tmp_path / "md2word.json").write_text(
            '{"toc": true, "number_headings": true}',
            encoding="utf-8",
        )
        cfg = load_config(tmp_path)
        assert cfg.get("toc") is True
        assert cfg.get("number_headings") is True

    def test_loads_pyproject(self, tmp_path):
        content = """[tool.md2word]
toc = true
number-headings = true
"""
        (tmp_path / "pyproject.toml").write_text(content, encoding="utf-8")
        cfg = load_config(tmp_path)
        assert cfg.get("toc") is True
        assert cfg.get("number_headings") is True

    def test_malformed_file_returns_empty(self, tmp_path):
        (tmp_path / "md2word.yaml").write_text(": : malformed", encoding="utf-8")
        cfg = load_config(tmp_path)
        # Should not crash, return {}
        assert isinstance(cfg, dict)


class TestMergeWithArgs:
    def test_cli_overrides_config(self):
        cfg = {"toc": False, "theme": "academic"}
        args = {"toc": True, "theme": None, "output": None}
        merged = merge_with_args(cfg, args)
        assert merged["toc"] is True  # CLI wins
        assert merged["theme"] == "academic"  # config fills None

    def test_config_fills_cli_defaults(self):
        cfg = {"number_headings": True}
        args = {"toc": None, "number_headings": None}
        merged = merge_with_args(cfg, args)
        assert merged["number_headings"] is True
        assert merged["toc"] is None  # Not in config

    def test_cli_explicit_none_does_not_override(self):
        cfg = {"theme": "official"}
        args = {"theme": None}
        merged = merge_with_args(cfg, args)
        assert merged["theme"] == "official"

    def test_empty_config(self):
        cfg = {}
        args = {"toc": True, "output": "out.docx"}
        merged = merge_with_args(cfg, args)
        assert merged == args
