"""Tests for image utility functions."""

from md2word.image_utils import (
    is_svg_source,
    is_svg_data,
    parse_svg_size,
    _strip_svg_unit,
)


class TestStripSvgUnit:
    def test_removes_units(self):
        assert _strip_svg_unit("100pt") == "100"
        assert _strip_svg_unit("100px") == "100"
        assert _strip_svg_unit("2.5cm") == "2.5"
        assert _strip_svg_unit("10mm") == "10"
        assert _strip_svg_unit("1in") == "1"

    def test_no_unit(self):
        assert _strip_svg_unit("100") == "100"
        assert _strip_svg_unit("0.5") == "0.5"

    def test_empty_string(self):
        assert _strip_svg_unit("") == ""


class TestIsSvgSource:
    def test_extension_check(self):
        assert is_svg_source("image.svg") is True
        assert is_svg_source("path/to/image.SVG") is True
        assert is_svg_source("image.png") is False
        assert is_svg_source("image.jpg") is False
        assert is_svg_source("") is False


class TestIsSvgData:
    def test_svg_with_tag(self):
        assert is_svg_data(b"<svg xmlns='http://www.w3.org/2000/svg'></svg>") is True

    def test_svg_with_xml_declaration(self):
        data = b'<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg">'
        assert is_svg_data(data) is True

    def test_not_svg(self):
        assert is_svg_data(b"PNG data here") is False
        assert is_svg_data(b"") is False

    def test_svg_with_attributes(self):
        data = b'<svg width="100" height="100" viewBox="0 0 100 100">'
        assert is_svg_data(data) is True


class TestParseSvgSize:
    def test_from_viewBox(self):
        svg = b'<svg viewBox="0 0 200 100"></svg>'
        size = parse_svg_size(svg)
        assert size == (200.0, 100.0)

    def test_from_viewBox_with_commas(self):
        svg = b'<svg viewBox="0,0,200,100"></svg>'
        size = parse_svg_size(svg)
        assert size == (200.0, 100.0)

    def test_from_width_height(self):
        svg = b'<svg width="300" height="150"></svg>'
        size = parse_svg_size(svg)
        assert size == (300.0, 150.0)

    def test_no_dimensions(self):
        svg = b"<svg></svg>"
        assert parse_svg_size(svg) is None

    def test_invalid_xml(self):
        assert parse_svg_size(b"not xml") is None
