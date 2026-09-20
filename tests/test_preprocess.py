r"""Flattening: comments, verbatim, \input and macro expansion."""

from __future__ import annotations

from beamer2pptx.preprocess import (expand_macros, find_group, find_optional,
                                    find_overlay, harvest_macros, preprocess,
                                    protect_verbatim, strip_comments)


def test_find_group_handles_nesting_and_escapes():
    assert find_group("{a{b}c}rest", 0) == ("a{b}c", 7)
    assert find_group(r"{a\}b}", 0) == (r"a\}b", 6)
    assert find_group("no group", 0) == ("", 0)


def test_find_optional_and_overlay():
    assert find_optional("[opt]{x}", 0) == ("opt", 5)
    assert find_optional("{x}", 0) == (None, 0)
    assert find_overlay("<2->{x}", 0) == ("2-", 4)


def test_strip_comments_keeps_escaped_percent():
    text = "a % comment\nb \\% literal\n"
    out = strip_comments(text)
    assert "comment" not in out
    assert "\\%" in out
    assert out.startswith("a b")


def test_strip_comments_removes_comment_environment():
    assert "hidden" not in strip_comments(
        "keep \\begin{comment}hidden\\end{comment} keep")


def test_protect_verbatim_isolates_content():
    store = []
    out = protect_verbatim(
        "before \\begin{verbatim}x % y $z$\\end{verbatim} after", store)
    assert store[0][0] == "verbatim"
    assert store[0][1] == "x % y $z$"
    assert "verbatim}" not in out
    # The placeholder must survive comment stripping intact.
    assert "%" not in strip_comments(out)


def test_protect_verb_inline():
    store = []
    out = protect_verbatim(r"see \verb|a % b| done", store)
    assert store[0] == ("verb", "a % b")
    assert "beamerverbinline" in out


def test_harvest_and_expand_macros():
    source = r"\newcommand{\R}{\mathbb{R}}" "\n" r"\newcommand{\hi}[2][world]{Hi #1 and #2}"
    macros = harvest_macros(source)
    assert macros["R"].body == r"\mathbb{R}"
    assert macros["hi"].nargs == 2
    assert macros["hi"].default == "world"

    assert expand_macros(r"$\R$", macros) == r"$\mathbb{R}$"
    assert expand_macros(r"\hi{you}", macros) == "Hi world and you"
    assert expand_macros(r"\hi[all]{you}", macros) == "Hi all and you"


def test_expand_macros_leaves_structural_macros_alone():
    macros = harvest_macros(r"\renewcommand{\item}{BROKEN}")
    assert expand_macros(r"\item text", macros) == r"\item text"


def test_preprocess_inlines_inputs_and_reads_graphicspath(tmp_path):
    (tmp_path / "main.tex").write_text(
        "\\documentclass{beamer}\n"
        "\\graphicspath{{figs/}{img/}}\n"
        "\\begin{document}\n\\input{part}\n\\end{document}\n",
        encoding="utf-8")
    (tmp_path / "part.tex").write_text("INCLUDED BODY", encoding="utf-8")

    source = preprocess(str(tmp_path / "main.tex"), str(tmp_path))
    assert "INCLUDED BODY" in source.body
    assert source.graphicspaths == ["figs/", "img/"]
    assert len(source.files) == 2


def test_preprocess_survives_a_circular_input(tmp_path):
    (tmp_path / "main.tex").write_text(
        "\\documentclass{beamer}\\begin{document}\\input{a}\\end{document}",
        encoding="utf-8")
    (tmp_path / "a.tex").write_text("A\\input{b}", encoding="utf-8")
    (tmp_path / "b.tex").write_text("B\\input{a}", encoding="utf-8")

    source = preprocess(str(tmp_path / "main.tex"), str(tmp_path))
    assert "A" in source.body and "B" in source.body
