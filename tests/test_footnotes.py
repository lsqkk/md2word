"""Tests for footnote extraction and insertion."""

from md2word.footnotes import extract_footnotes, FootnoteDef


class TestExtractFootnotes:
    def test_no_footnotes(self):
        text = "这是一段没有脚注的文字。"
        cleaned, fns = extract_footnotes(text)
        assert cleaned == text
        assert fns == []

    def test_single_footnote(self):
        text = "正文[^1]内容。\n\n[^1]: 脚注说明"
        cleaned, fns = extract_footnotes(text)
        assert "\x00FN_1\x00" in cleaned
        assert "[^1]" not in cleaned
        assert len(fns) == 1
        assert fns[0] == FootnoteDef(id="1", content="脚注说明")

    def test_multiple_footnotes(self):
        text = "先[^a]后[^b]。\n\n[^a]: 甲\n[^b]: 乙"
        cleaned, fns = extract_footnotes(text)
        assert "\x00FN_a\x00" in cleaned
        assert "\x00FN_b\x00" in cleaned
        assert len(fns) == 2
        assert fns[0] == FootnoteDef(id="a", content="甲")
        assert fns[1] == FootnoteDef(id="b", content="乙")

    def test_multiline_footnote(self):
        text = "内容[^x]。\n\n[^x]: 第一行\n    第二行\n    第三行"
        cleaned, fns = extract_footnotes(text)
        assert len(fns) == 1
        # Multi-line content should be joined with spaces
        assert fns[0].id == "x"

    def test_undefined_reference_preserved(self):
        text = "未定义的脚注[^undefined]。"
        cleaned, fns = extract_footnotes(text)
        assert "[^undefined]" in cleaned
        assert fns == []

    def test_definitions_removed(self):
        text = "正文[^1]。\n\n[^1]: 说明\n\n更多正文。"
        cleaned, fns = extract_footnotes(text)
        assert "[^1]:" not in cleaned
        assert len(fns) == 1
        assert fns[0] == FootnoteDef(id="1", content="说明")

    def test_special_chars_in_id(self):
        text = "引用[^ref_1]。\n\n[^ref_1]: 带下划线的ID"
        cleaned, fns = extract_footnotes(text)
        assert "\x00FN_ref_1\x00" in cleaned
        assert len(fns) == 1
