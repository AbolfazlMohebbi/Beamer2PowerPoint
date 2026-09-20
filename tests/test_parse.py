"""Parsing frames into the intermediate representation."""

from __future__ import annotations

import pytest

from beamer2pptx.ir import (BeamerBlock, BulletList, Columns, Equation, Figure,
                            Paragraph, Verbatim, runs_to_text)
from beamer2pptx.parse.document import ParseOptions


def find(blocks, kind):
    """Depth-first search for the first block of *kind*."""
    for block in blocks:
        if isinstance(block, kind):
            return block
        for nested in _children(block):
            found = find(nested, kind)
            if found is not None:
                return found
    return None


def _children(block):
    if isinstance(block, Columns):
        yield from (c.blocks for c in block.columns)
    elif isinstance(block, BeamerBlock):
        yield block.blocks
    elif isinstance(block, BulletList):
        yield from (i.blocks for i in block.items)


def text_of(runs):
    return runs_to_text(runs)


# -- frames and metadata ---------------------------------------------------


def test_metadata_and_title_slide(parse_body):
    deck = parse_body(r"\begin{frame}\titlepage\end{frame}")
    assert text_of(deck.title) == "Test deck"
    assert text_of(deck.author) == "Tester"
    assert len(deck.slides) == 1
    assert deck.slides[0].kind == "title"
    # A \titlepage frame inherits the document metadata.
    assert text_of(deck.slides[0].title) == "Test deck"


def test_frame_title_forms(parse_body):
    deck = parse_body(
        r"\begin{frame}{Braced}a\end{frame}"
        r"\begin{frame}\frametitle{Macro}b\end{frame}"
        r"\begin{frame}{Both}{Sub}c\end{frame}")
    assert [text_of(s.title) for s in deck.slides] == ["Braced", "Macro", "Both"]
    assert text_of(deck.slides[2].subtitle) == "Sub"


def test_frame_option_does_not_become_a_title(parse_body):
    deck = parse_body(r"\begin{frame}[fragile]{Real}body\end{frame}")
    assert text_of(deck.slides[0].title) == "Real"


def test_section_slides_can_be_switched_off(parse_body):
    body = r"\section{S}\begin{frame}{F}x\end{frame}"
    assert len(parse_body(body).slides) == 2
    off = ParseOptions(section_slides=False)
    assert len(parse_body(body, off).slides) == 1


def test_note_becomes_speaker_notes(parse_body):
    deck = parse_body(r"\begin{frame}{F}body\note{Remember this}\end{frame}")
    assert "Remember this" in deck.slides[0].notes


# -- inline content --------------------------------------------------------


def test_inline_styles_and_colours(parse_body):
    deck = parse_body(r"\begin{frame}{F}"
                      r"\textbf{b} \emph{i} \texttt{m} \alert{a} "
                      r"\textcolor{blue}{c}\end{frame}")
    runs = find(deck.slides[0].blocks, Paragraph).runs
    styles = {r.text.strip(): (r.bold, r.italic, r.mono, r.color) for r in runs}
    assert styles["b"][0] is True
    assert styles["i"][1] is True
    assert styles["m"][2] is True
    assert styles["a"][3] == "C00000"
    assert styles["c"][3] == "0000FF"


def test_accents_and_escapes(parse_body):
    deck = parse_body(r"\begin{frame}{F}Montr\'eal 50\% caf\'{e} na\"ive\end{frame}")
    text = text_of(find(deck.slides[0].blocks, Paragraph).runs)
    assert "Montréal" in text
    assert "50%" in text
    assert "café" in text
    assert "naïve" in text


def test_links(parse_body):
    deck = parse_body(r"\begin{frame}{F}\url{https://a.example} "
                      r"\href{https://b.example}{named}\end{frame}")
    runs = find(deck.slides[0].blocks, Paragraph).runs
    targets = {r.hyperlink for r in runs if r.hyperlink}
    assert targets == {"https://a.example", "https://b.example"}


def test_unknown_macro_keeps_its_text(parse_body):
    deck = parse_body(r"\begin{frame}{F}\mysterymacro{kept text}\end{frame}")
    assert "kept text" in text_of(find(deck.slides[0].blocks, Paragraph).runs)


# -- lists -----------------------------------------------------------------


def test_nested_lists_get_levels(parse_body):
    deck = parse_body(r"""
\begin{frame}{F}
\begin{itemize}\item one
  \begin{itemize}\item two
    \begin{itemize}\item three\end{itemize}
  \end{itemize}
\end{itemize}
\end{frame}""")
    items = find(deck.slides[0].blocks, BulletList).items
    assert [i.level for i in items] == [0, 1, 2]
    assert [text_of(i.runs) for i in items] == ["one", "two", "three"]


def test_enumerate_is_marked_ordered(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{enumerate}\item a\end{enumerate}\end{frame}")
    assert find(deck.slides[0].blocks, BulletList).items[0].ordered is True


def test_description_labels(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{description}"
                      r"\item[Key] value\end{description}\end{frame}")
    assert find(deck.slides[0].blocks, BulletList).items[0].marker == "Key"


def test_math_symbol_marker(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{itemize}"
                      r"\item[$\star$] starred\end{itemize}\end{frame}")
    assert find(deck.slides[0].blocks, BulletList).items[0].marker == "★"


# -- overlays --------------------------------------------------------------


def test_pause_and_explicit_overlays(parse_body):
    deck = parse_body(r"""
\begin{frame}{F}
\begin{itemize}\item a \pause \item b \item<3-> c\end{itemize}
\end{frame}""")
    items = find(deck.slides[0].blocks, BulletList).items
    assert items[0].overlay.ranges == ()
    assert items[1].overlay.ranges == ((2, None),)
    assert items[2].overlay.ranges == ((3, None),)
    assert deck.slides[0].max_overlay == 3


def test_expand_produces_one_slide_per_step(parse_body):
    body = r"""
\begin{frame}{F}
\begin{itemize}\item a \pause \item b\end{itemize}
\end{frame}"""
    deck = parse_body(body, ParseOptions(overlays="expand"))
    assert len(deck.slides) == 2
    first = find(deck.slides[0].blocks, BulletList)
    second = find(deck.slides[1].blocks, BulletList)
    assert len(first.items) == 1
    assert len(second.items) == 2


def test_flatten_keeps_everything(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{itemize}"
                      r"\item a \pause \item b\end{itemize}\end{frame}")
    assert len(deck.slides) == 1
    assert len(find(deck.slides[0].blocks, BulletList).items) == 2


def test_only_takes_the_visible_branch(parse_body):
    deck = parse_body(r"\begin{frame}{F}\only<2->{shown}\end{frame}")
    assert "shown" in text_of(find(deck.slides[0].blocks, Paragraph).runs)


# -- structure -------------------------------------------------------------


def test_columns(parse_body):
    deck = parse_body(r"""
\begin{frame}{F}
\begin{columns}
\begin{column}{0.3\textwidth}left\end{column}
\begin{column}{0.7\textwidth}right\end{column}
\end{columns}
\end{frame}""")
    columns = find(deck.slides[0].blocks, Columns)
    assert [round(c.width_frac, 2) for c in columns.columns] == [0.3, 0.7]
    assert "left" in text_of(find(columns.columns[0].blocks, Paragraph).runs)


def test_beamer_blocks(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{alertblock}{Heads up}"
                      r"body\end{alertblock}\end{frame}")
    block = find(deck.slides[0].blocks, BeamerBlock)
    assert block.kind == "alertblock"
    assert text_of(block.title) == "Heads up"


def test_theorem_becomes_a_block(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{theorem}[Name]"
                      r"statement\end{theorem}\end{frame}")
    assert text_of(find(deck.slides[0].blocks, BeamerBlock).title) == "Theorem (Name)"


def test_verbatim_is_preserved_literally(parse_body):
    deck = parse_body("\\begin{frame}[fragile]{F}\n"
                      "\\begin{verbatim}\nx = 1  % not a comment\n"
                      "\\end{verbatim}\n\\end{frame}")
    verbatim = find(deck.slides[0].blocks, Verbatim)
    assert verbatim.text == "x = 1  % not a comment"


# -- figures and equations -------------------------------------------------


def test_includegraphics_options(parse_body):
    deck = parse_body(r"\begin{frame}{F}"
                      r"\includegraphics[width=0.6\textwidth,angle=90]{pic.png}"
                      r"\end{frame}")
    figure = find(deck.slides[0].blocks, Figure)
    assert figure.source_path == "pic.png"
    assert figure.width_frac == pytest.approx(0.6)
    assert figure.angle == 90


def test_figure_caption_attaches(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{figure}"
                      r"\includegraphics{a.png}\caption{A caption}"
                      r"\end{figure}\end{frame}")
    assert text_of(find(deck.slides[0].blocks, Figure).caption) == "A caption"


def test_tikz_is_captured_verbatim(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{tikzpicture}"
                      r"\draw (0,0)--(1,1);\end{tikzpicture}\end{frame}")
    figure = find(deck.slides[0].blocks, Figure)
    assert figure.kind == "tikz"
    assert r"\draw (0,0)--(1,1);" in figure.tex


def test_display_and_inline_math(parse_body):
    deck = parse_body(r"\begin{frame}{F}inline $a^2$ then \[ b^2 \]\end{frame}")
    paragraph = find(deck.slides[0].blocks, Paragraph)
    assert any(r.math == "a^2" for r in paragraph.runs)
    assert find(deck.slides[0].blocks, Equation).latex == "b^2"


def test_align_splits_into_lines(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{align}a &= b \\ c &= d\end{align}\end{frame}")
    equation = find(deck.slides[0].blocks, Equation)
    assert equation.lines == ["a &= b", "c &= d"]


def test_align_does_not_split_inside_a_matrix(parse_body):
    deck = parse_body(r"\begin{frame}{F}\begin{align}"
                      r"A &= \begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix} \\ b &= c"
                      r"\end{align}\end{frame}")
    equation = find(deck.slides[0].blocks, Equation)
    assert len(equation.lines) == 2
    assert "bmatrix" in equation.lines[0]
