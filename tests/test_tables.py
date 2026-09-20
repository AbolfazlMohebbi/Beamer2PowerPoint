"""Table parsing: column specs, spans and rules."""

from __future__ import annotations

import pytest

from beamer2pptx.parse.inline import InlineContext
from beamer2pptx.parse.tables import (parse_colspec, parse_length, parse_table,
                                      split_top_level)


@pytest.fixture
def ctx():
    return InlineContext()


def text(cell):
    return "".join(r.text for r in cell.runs)


def test_parse_colspec_basics():
    aligns, widths, vlines = parse_colspec("|l|c|r|")
    assert aligns == ["l", "c", "r"]
    assert widths == [0.0, 0.0, 0.0]
    assert vlines[:4] == [True, True, True, True]


def test_parse_colspec_paragraph_and_flexible_columns():
    aligns, widths, _ = parse_colspec(r"l p{3cm} X")
    assert aligns == ["l", "l", "l"]
    assert widths[1] == pytest.approx(3 * 28.45276, rel=1e-3)
    assert widths[2] == -1.0            # tabularx shares the leftover width


def test_parse_colspec_expands_repeats():
    aligns, _, _ = parse_colspec(r"*{3}{c}l")
    assert aligns == ["c", "c", "c", "l"]


def test_parse_colspec_skips_decorations():
    aligns, _, _ = parse_colspec(r"@{}l>{\bfseries}c@{\extracolsep{2pt}}r@{}")
    assert aligns == ["l", "c", "r"]


def test_parse_length():
    assert parse_length("10pt") == pytest.approx(10.0)
    assert parse_length("1in") == pytest.approx(72.27)
    assert parse_length(r"0.5\textwidth", 300.0) == pytest.approx(150.0)
    assert parse_length("nonsense") is None


def test_split_top_level_respects_braces_and_environments():
    assert split_top_level(r"a & b", "&") == ["a ", " b"]
    assert split_top_level(r"\textbf{a & b} & c", "&") == [r"\textbf{a & b} ", " c"]
    assert len(split_top_level(r"x \\ y \\ z", "\\\\")) == 3
    # A nested tabular keeps its own separators.
    nested = r"\begin{tabular}{cc}p & q\end{tabular} & r"
    assert len(split_top_level(nested, "&")) == 2


def test_simple_table(ctx):
    table = parse_table(r"a & b \\ c & d \\", "ll", ctx)
    assert len(table.rows) == 2
    assert [text(c) for c in table.rows[0]] == ["a", "b"]
    assert [text(c) for c in table.rows[1]] == ["c", "d"]


def test_booktabs_rules_and_header(ctx):
    body = r"\toprule A & B \\ \midrule 1 & 2 \\ \bottomrule"
    table = parse_table(body, "lr", ctx)
    assert table.header_rows == 1
    assert all(c.bold for c in table.rows[0])
    assert all(c.top_rule for c in table.rows[0])
    assert all(c.bottom_rule for c in table.rows[-1])


def test_multicolumn_creates_a_span(ctx):
    table = parse_table(r"\multicolumn{2}{c}{wide} & z \\", "lll", ctx)
    row = table.rows[0]
    assert row[0].colspan == 2
    assert row[0].align == "c"
    assert row[1].merged is True
    assert text(row[2]) == "z"


def test_multirow_creates_a_vertical_span(ctx):
    table = parse_table(r"\multirow{2}{*}{tall} & a \\ b \\", "ll", ctx)
    assert table.rows[0][0].rowspan == 2
    assert table.rows[1][0].merged is True
    assert text(table.rows[1][1]) == "b"


def test_cell_inline_formatting_and_math(ctx):
    table = parse_table(r"\textbf{bold} & $x^2$ \\", "ll", ctx)
    assert table.rows[0][0].runs[0].bold is True
    assert table.rows[0][1].runs[0].math == "x^2"


def test_rows_are_padded_to_the_widest(ctx):
    table = parse_table(r"a & b & c \\ d \\", "lll", ctx)
    assert all(len(row) == 3 for row in table.rows)


def test_cmidrule_trim_specifier_is_consumed(ctx):
    r"""``\cmidrule(lr){3-4}`` must not leak "(lr)3-4" into a cell."""
    body = (r"\toprule A & B \\ \cmidrule(lr){1-2} 1 & 2 \\ \bottomrule")
    table = parse_table(body, "ll", ctx)
    rendered = " ".join(text(c) for row in table.rows for c in row)
    assert "lr" not in rendered
    assert "1-2" not in rendered
    assert [text(c) for c in table.rows[1]] == ["1", "2"]


def test_multirow_continuation_does_not_shift_columns(ctx):
    r"""A continuation row supplies an empty cell for the spanned column.

    That placeholder must be consumed by the span rather than pushing the
    rest of the row one column to the right.
    """
    body = (r"\multirow{2}{*}{Group} & Alpha & 1 \\"
            r"                       & Beta  & 2 \\")
    table = parse_table(body, "llr", ctx)
    assert len(table.col_aligns) == 3, table.col_aligns
    assert all(len(row) == 3 for row in table.rows)
    assert table.rows[0][0].rowspan == 2
    assert table.rows[1][0].merged is True
    assert [text(c) for c in table.rows[1][1:]] == ["Beta", "2"]


def test_nested_header_with_multirow_and_multicolumn(ctx):
    """The full booktabs header idiom: stacked spans in both directions."""
    body = (
        r"\toprule"
        r"\multirow{2}{*}{Cohort} & \multicolumn{2}{c}{Error} & "
        r"\multirow{2}{*}{Rate} \\"
        r"\cmidrule(lr){2-3} "
        r" & Peak & RMS & \\"
        r"\midrule "
        r"Healthy & 4.1 & 2.3 & 118 \\"
        r"\bottomrule"
    )
    table = parse_table(body, "lrrr", ctx)
    assert len(table.col_aligns) == 4
    assert all(len(row) == 4 for row in table.rows)
    assert table.rows[0][0].rowspan == 2 and text(table.rows[0][0]) == "Cohort"
    assert table.rows[0][1].colspan == 2 and text(table.rows[0][1]) == "Error"
    assert [text(c) for c in table.rows[1]] == ["", "Peak", "RMS", ""]
    assert [text(c) for c in table.rows[2]] == ["Healthy", "4.1", "2.3", "118"]


def test_escaped_ampersand_is_not_a_separator(ctx):
    table = parse_table(r"AT\&T & other \\", "ll", ctx)
    assert len(table.rows[0]) == 2
    assert text(table.rows[0][0]) == "AT&T"
