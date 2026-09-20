r"""Resolving every asset a deck needs before it is laid out.

Walks the whole deck once: figures are found and converted into ``figures/``,
TikZ snippets are compiled, and equations are turned into OMML (or rendered as
images when that fails).  Doing it up front means the layout stage knows the
real size of everything it has to place.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import List

from ..ir import (BeamerBlock, Block, BulletList, Columns, Deck, Equation,
                  Figure, Table)
from .figures import FigureStore
from .latexrun import LatexRunner
from .math_omml import MathConverter

log = logging.getLogger(__name__)


@dataclass
class PrepareReport:
    figures_written: int = 0
    figures_vector: int = 0
    tikz_compiled: int = 0
    equations_native: int = 0
    equations_image: int = 0
    equations_text: int = 0
    missing_figures: List[str] = field(default_factory=list)
    failed_tikz: List[str] = field(default_factory=list)
    messages: List[str] = field(default_factory=list)


@dataclass
class AssetPreparer:
    store: FigureStore
    runner: LatexRunner
    math: MathConverter
    preamble: str = ""
    math_mode: str = "native"           # native | image
    equation_dpi: int = 600
    report: PrepareReport = field(default_factory=PrepareReport)

    def prepare(self, deck: Deck) -> PrepareReport:
        for index, slide in enumerate(deck.slides, start=1):
            self._walk(slide.blocks, slide, index)
        return self.report

    # -- traversal --------------------------------------------------------

    def _walk(self, blocks: List[Block], slide, slide_number: int) -> None:
        for block in blocks:
            if isinstance(block, Figure):
                self._prepare_figure(block, slide, slide_number)
            elif isinstance(block, Equation):
                self._prepare_equation(block, slide)
            elif isinstance(block, BulletList):
                for item in block.items:
                    self._walk(item.blocks, slide, slide_number)
            elif isinstance(block, Columns):
                for column in block.columns:
                    self._walk(column.blocks, slide, slide_number)
            elif isinstance(block, BeamerBlock):
                self._walk(block.blocks, slide, slide_number)
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row:
                        self._walk(cell.blocks, slide, slide_number)

    # -- figures ----------------------------------------------------------

    def _prepare_figure(self, figure: Figure, slide, slide_number: int) -> None:
        if figure.kind == "tikz":
            self._prepare_tikz(figure, slide, slide_number)
            return

        reference = figure.source_path or ""
        resolved = self.store.resolve(reference)
        if resolved is None:
            message = "Figure not found: %s" % reference
            self.report.missing_figures.append(reference)
            slide.warnings.append(message)
            log.warning("Slide %d: %s", slide_number, message)
            return

        placed = self.store.place(resolved, prefix="fig%02d" % slide_number)
        if placed is None:
            slide.warnings.append("Could not convert figure: %s" % reference)
            return
        figure.rendered_path = placed.png_path
        figure.native_size = placed.size
        figure.vector_path = placed.svg_path
        self.report.figures_written += 1
        if placed.is_vector:
            self.report.figures_vector += 1

    def _prepare_tikz(self, figure: Figure, slide, slide_number: int) -> None:
        snippet = figure.tex or ""
        if not self.runner.available:
            slide.warnings.append(
                "TikZ picture skipped: %s" % (self.runner.unavailable_reason
                                              or "no LaTeX engine"))
            return

        pdf = self.runner.render_tikz(snippet, self.preamble)
        if pdf is None:
            reason = self.runner.failures[-1] if self.runner.failures else "compile failed"
            self.report.failed_tikz.append(reason)
            slide.warnings.append("TikZ picture failed to compile: %s" % reason)
            log.warning("Slide %d: TikZ compile failed (%s)", slide_number, reason)
            return

        placed = self.store.place_pdf(pdf, prefix="tikz%02d" % slide_number,
                                      transparent=True, crop=True)
        if placed is None:
            slide.warnings.append("TikZ picture could not be rasterised")
            return
        figure.rendered_path = placed.png_path
        figure.native_size = placed.size
        figure.vector_path = placed.svg_path
        self.report.tikz_compiled += 1
        self.report.figures_written += 1
        if placed.is_vector:
            self.report.figures_vector += 1

    # -- equations --------------------------------------------------------

    def _prepare_equation(self, equation: Equation, slide) -> None:
        if self.math_mode == "native":
            omml = self._convert_native(equation)
            if omml is not None:
                equation.omml = omml
                self.report.equations_native += 1
                return

        if self._render_equation_image(equation):
            self.report.equations_image += 1
            return

        self.report.equations_text += 1
        slide.warnings.append("Equation kept as LaTeX text: %s"
                              % _short(equation.latex))

    def _convert_native(self, equation: Equation):
        """Convert to OMML, returning one element per rendered line."""
        if not self.math.available:
            return None
        if equation.lines and len(equation.lines) > 1:
            # align/gather: one m:oMath per row, stacked by the oMathPara.
            elements = [self.math.to_omml(line) for line in equation.lines]
            if any(e is None for e in elements):
                return None
            return elements
        element = self.math.to_omml(equation.latex)
        return [element] if element is not None else None

    def _render_equation_image(self, equation: Equation) -> bool:
        if not self.runner.available:
            return False
        pdf = self.runner.render_math(equation.latex, display=equation.display,
                                      preamble=self.preamble)
        if pdf is None:
            return False
        placed = self.store.place_pdf(pdf, prefix="eq", transparent=True,
                                      crop=True, dpi=self.equation_dpi)
        if placed is None:
            return False
        equation.rendered_path = placed.png_path
        equation.rendered_size = placed.size
        equation.vector_path = placed.svg_path
        if equation.rendered_size and equation.rendered_size[1]:
            equation.rendered_height_pt = (equation.rendered_size[1] * 72.0
                                           / self.equation_dpi)
        return True


def _short(text: str, limit: int = 50) -> str:
    collapsed = " ".join(text.split())
    return collapsed if len(collapsed) <= limit else collapsed[:limit] + "..."
