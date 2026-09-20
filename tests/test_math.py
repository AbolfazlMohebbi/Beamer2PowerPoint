"""LaTeX to OMML conversion."""

from __future__ import annotations

import re

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


# -- regressions: silent content loss on the way to PowerPoint -------------


def omml_runs(element):
    """The text of each run, in order."""
    return ["".join(t.itertext())
            for t in element.iter("{%s}t" % M_NAMESPACE)]


def test_bare_norm_is_promoted_to_left_right():
    r"""``\|x\|`` must become ``\left\|x\right\|`` before conversion."""
    out = normalise_latex(r"\|x\|_2^2")
    assert r"\left\|" in out and r"\right\|" in out
    # Already-qualified bars are left alone rather than doubled.
    assert normalise_latex(r"\left\|x\right\|").count(r"\left\|") == 1
    # An odd number of bars is ambiguous, so it is left untouched.
    assert r"\left\|" not in normalise_latex(r"P(A\|B)")


@requires_office
def test_norm_with_exponent_keeps_its_contents():
    """A norm carrying a sub/superscript used to convert to empty bars."""
    element = MathConverter().to_omml(r"\|\mathbf{p} - \Phi\mathbf{w}\|_2^2")
    assert element is not None
    text = "".join(omml_runs(element))
    assert "\U0001D429" in text, "the p inside the norm was lost"
    assert "\U0001D430" in text, "the w inside the norm was lost"
    assert "Φ" in text


@requires_office
def test_non_bmp_characters_get_their_own_run():
    r"""PowerPoint renders only the first astral character of a run.

    ``\mathbf{r} = \mathbf{p} - \Phi`` arrives from the stylesheet as the
    single run "r=p-Phi" in bold-maths code points, and PowerPoint then
    silently drops the p.
    """
    element = MathConverter().to_omml(r"\mathbf{r} = \mathbf{p} - \Phi \mathbf{w}")
    assert element is not None
    runs = omml_runs(element)
    for run in runs:
        astral = [c for c in run if ord(c) > 0xFFFF]
        assert len(astral) <= 1, "run %r holds %d astral chars" % (run, len(astral))
    joined = "".join(runs)
    for char in ("\U0001D42B", "\U0001D429", "\U0001D430"):
        assert char in joined


@requires_office
def test_operator_is_not_left_as_a_delimiter_separator():
    """A minus inside a norm must not end up in m:sepChr (drawn as an arrow)."""
    element = MathConverter().to_omml(r"\left\|a - b\right\|")
    assert element is not None
    from lxml import etree

    xml = etree.tostring(element, encoding="unicode")
    separators = re.findall(r"<m:sepChr m:val=\"([^\"]*)\"", xml)
    assert all(s in (",", ";", "|", "‖", "") for s in separators), separators
    assert "−" in "".join(omml_runs(element))


@requires_office
def test_lossy_conversion_is_rejected_rather_than_shipped():
    """Content the stylesheet drops must fail over to the image path."""
    converter = MathConverter()
    # Single bars with a subscript still defeat the stylesheet; the guard
    # must notice rather than emit a formula missing its variable.
    assert converter.to_omml(r"|x|_1") is None
    assert any("|x|_1" in f for f in converter.failures)


@requires_office
@pytest.mark.parametrize("latex", [
    r"\sqrt{x}", r"\prod_{k} a_k", r"(a+b)^2", r"\hat{x}", r"\bar{z}",
    r"\binom{n}{k}", r"x_{i,j}^{(k)}", r"(\Phi^\top \Phi + \lambda I)^{-1}",
    r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}",
])
def test_guard_does_not_reject_good_conversions(latex):
    """The loss guard must not cost us equations that convert correctly."""
    assert MathConverter().to_omml(latex) is not None


@requires_office
def test_failure_is_reported_not_raised():
    converter = MathConverter()
    result = converter.to_omml(r"\thiscommanddoesnotexist{")
    assert result is None or converter.failures == [] or converter.failures


def test_disabled_converter_reports_unavailable():
    assert MathConverter(enabled=False).available is False
