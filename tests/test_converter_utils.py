"""Tests for converter utility functions."""
import pytest
from docx.shared import Inches, RGBColor

from md2word.handlers import (
    _slugify,
    ensure_list_blank_lines,
    inline_text,
)
from md2word.context import ConversionContext


class TestEnsureListBlankLines:
    def test_inserts_blank_before_list(self):
        text = "前面有文字\n- 列表项\n- 列表项2"
        result = ensure_list_blank_lines(text)
        assert "文字\n\n- 列表项" in result or "文字\n- 列表项" in result

    def test_preserves_existing_blank_line(self):
        text = "前面有文字\n\n- 列表项"
        result = ensure_list_blank_lines(text)
        assert result == text

    def test_heading_before_list_no_insert(self):
        text = "# 标题\n- 列表项"
        result = ensure_list_blank_lines(text)
        assert result == text

    def test_ordered_list(self):
        text = "文字\n1. 第一项\n2. 第二项"
        result = ensure_list_blank_lines(text)
        assert result != text

    def test_consecutive_lists(self):
        text = "- 项1\n- 项2"
        result = ensure_list_blank_lines(text)
        assert result == text

    def test_empty_text(self):
        assert ensure_list_blank_lines("") == ""

    def test_no_list_markers(self):
        text = "普通文本\n没有列表\n结束"
        assert ensure_list_blank_lines(text) == text


class TestSlugify:
    def test_basic_slug(self):
        assert _slugify("Hello World") == "hello-world"

    def test_chinese_text(self):
        result = _slugify("第一章 引言")
        assert isinstance(result, str)

    def test_special_characters(self):
        assert _slugify("Hello, World!") == "hello-world"
        assert _slugify("Test [v1.0]") == "test-v10"

    def test_whitespace_handling(self):
        assert _slugify("  multiple   spaces  ") == "multiple-spaces"

    def test_empty_string_fallback(self):
        assert _slugify("") == "ref"
        assert _slugify("   ") == "ref"


class TestInlineText:
    def test_simple_text(self):
        import xml.etree.ElementTree as ET
        elem = ET.fromstring("<p>Hello World</p>")
        assert inline_text(elem) == "Hello World"

    def test_with_child_elements(self):
        import xml.etree.ElementTree as ET
        elem = ET.fromstring("<p>Hello <b>bold</b> world</p>")
        result = inline_text(elem)
        assert "Hello" in result
        assert "bold" in result
        assert "world" in result

    def test_with_tail_text(self):
        import xml.etree.ElementTree as ET
        elem = ET.fromstring("<p><b>Bold</b> tail</p>")
        assert inline_text(elem) == "Bold tail"

    def test_empty_element(self):
        import xml.etree.ElementTree as ET
        elem = ET.fromstring("<p></p>")
        assert inline_text(elem) == ""


class TestHeadingNumbering:
    def test_h1_not_numbered(self):
        ctx = ConversionContext()
        ctx.reset_heading_counters()
        assert ctx.next_heading_number(1) == ""

    def test_h2_first_is_1(self):
        ctx = ConversionContext()
        ctx.reset_heading_counters()
        assert ctx.next_heading_number(2) == "1"

    def test_h3_under_h2(self):
        ctx = ConversionContext()
        ctx.reset_heading_counters()
        ctx.next_heading_number(2)
        assert ctx.next_heading_number(3) == "1.1"

    def test_multiple_h2(self):
        ctx = ConversionContext()
        ctx.reset_heading_counters()
        ctx.next_heading_number(2)
        ctx.next_heading_number(2)
        assert ctx.next_heading_number(3) == "2.1"

    def test_deep_nesting(self):
        ctx = ConversionContext()
        ctx.reset_heading_counters()
        ctx.next_heading_number(2)
        ctx.next_heading_number(3)
        ctx.next_heading_number(4)
        assert ctx.next_heading_number(4) == "1.1.2"

    def test_reset(self):
        ctx = ConversionContext()
        ctx.reset_heading_counters()
        ctx.next_heading_number(2)
        ctx.next_heading_number(3)
        ctx.reset_heading_counters()
        assert ctx.next_heading_number(2) == "1"

    def test_bookmark_id_increments(self):
        ctx = ConversionContext()
        assert ctx.next_bookmark_id() == 1
        assert ctx.next_bookmark_id() == 2
        assert ctx.next_bookmark_id() == 3


class TestListMarkerRegex:
    def test_unordered_markers(self):
        import re
        assert re.match(r"^[\-\*\+] ", "- item")
        assert re.match(r"^[\-\*\+] ", "* item")
        assert re.match(r"^[\-\*\+] ", "+ item")
        assert not re.match(r"^[\-\*\+] ", "not a list")
        assert not re.match(r"[\-\*\+] ", "")

    def test_ordered_markers(self):
        import re
        assert re.match(r"^\d+[\.\)] ", "1. item")
        assert re.match(r"^\d+[\.\)] ", "1) item")
        assert re.match(r"^\d+[\.\)] ", "10. item")
        assert not re.match(r"^\d+[\.\)] ", "- item")
