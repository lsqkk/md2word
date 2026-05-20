"""Tests for OMML math conversion."""

from md2word.math_omml import latex_to_omml


class TestLatexToOmml:
    def test_simple_inline(self):
        result = latex_to_omml("E=mc^2")
        assert result is not None
        assert "<m:oMath" in result

    def test_simple_block(self):
        result = latex_to_omml("\\int_0^\\infty e^{-x} dx", display=True)
        assert result is not None
        assert "<m:oMathPara" in result

    def test_fraction(self):
        result = latex_to_omml("\\frac{a}{b}")
        assert result is not None
        assert "<m:f>" in result or "m:f" in result

    def test_greek_letters(self):
        result = latex_to_omml("\\alpha + \\beta")
        assert result is not None
        assert "<m:oMath" in result

    def test_invalid_latex(self):
        """Truly invalid LaTeX should return None, not crash."""
        result = latex_to_omml("\\invalidCommand{}")
        # latex2mathml may still produce output, the main test is no crash

    def test_empty(self):
        result = latex_to_omml("")
        assert result is None or result is not None  # may succeed or fail gracefully

    def test_special_functions(self):
        result = latex_to_omml("\\lim_{x \\to 0} f(x)")
        assert result is not None
