"""Tests for syntax highlighting helpers."""

import xml.etree.ElementTree as ET

from md2word.syntax import extract_language


class TestExtractLanguage:
    def _make_code_elem(self, cls: str = "") -> ET.Element:
        elem = ET.Element("code")
        if cls:
            elem.set("class", cls)
        return elem

    def test_with_language(self):
        elem = self._make_code_elem("language-python")
        assert extract_language(elem) == "python"

    def test_with_multiple_classes(self):
        elem = self._make_code_elem("language-javascript hljs")
        assert extract_language(elem) == "javascript"

    def test_no_class(self):
        elem = self._make_code_elem()
        assert extract_language(elem) == ""

    def test_no_language_prefix(self):
        elem = self._make_code_elem("nohighlight")
        assert extract_language(elem) == ""
