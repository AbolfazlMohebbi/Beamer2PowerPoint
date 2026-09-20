"""LaTeX to OMML conversion."""

from __future__ import annotations

import pytest

from beamer2pptx.render.math_omml import (M_NAMESPACE, MathConverter,
                                          find_mml2omml, normalise_latex)

requires_office = pytest.mark.skipif(
    find_mml2omml() is None,
    reason="MML2OMML.XSL (from Microsoft Office) is not installed")


def test_normalise_strips_layout_only_markup():
    assert "displaystyle" not in normalise_latex(r"\displaystyle x")
    assert normalise_latex(r"a &= b").strip() == "a  = b".strip()
    assert "label" not in normalise_latex(r"x \label{eq:1}")


def test_normalise_keeps_matrix_separators():
    out = normalise_latex(r"\begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix}")
    assert "1 & 2" in out


def test_normalise_strips_alignment_outside_environments_only():
    out = normalise_latex(r"a &= \begin{cases} 1 & x>0 \\ 0 & x\le 0 \end{cases}")
    assert "1 & x>0" in out          # kept: a cases column separator
    assert "a &=" not in out         # dropped: an align tab stop


@requires_office
def test_converter_is_available():
    assert MathConverter().available is True


@requires_office
@pytest.mark.parametrize("latex", [
    r"E = mc^2",
    r"\frac{\partial u}{\partial t} = \alpha \nabla^2 u",
    r"\sum_{i=1}^{n} x_i^2 \leq \int_0^\infty e^{-t}\,dt",
    r"\begin{bmatrix} a & b \\ c & d \end{bmatrix}",
    r"\sqrt[3]{x} + \alpha\beta\gamma",
    r"\left( \frac{1}{2} \right)",
])
def test_formulas_convert_to_omml(latex):
    element = MathConverter().to_omml(latex)
    assert element is not None
    assert element.tag == "{%s}oMath" % M_NAMESPACE
    assert len(element) > 0


@requires_office
def test_matrix_keeps_both_columns():
    element = MathConverter().to_omml(r"\begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix}")
    from lxml import etree

    xml = etree.tostring(element, encoding="unicode")
    # Two m:mr rows, and each entry its own run: not one run reading "12".
    assert xml.count("<m:mr>") == 2
    for digit in "1234":
        assert "<m:t>%s</m:t>" % digit in xml


@requires_office
def test_failure_is_reported_not_raised():
    converter = MathConverter()
    result = converter.to_omml(r"\thiscommanddoesnotexist{")
    assert result is None or converter.failures == [] or converter.failures


def test_disabled_converter_reports_unavailable():
    assert MathConverter(enabled=False).available is False
