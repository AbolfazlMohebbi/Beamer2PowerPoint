"""Intermediate representation shared by the parsers and the PowerPoint builder.

Everything the parsers understand is expressed with these dataclasses; the
builder never sees LaTeX.  Keeping the two halves apart means a new Beamer
construct only ever needs a parser change plus, at most, one new block type.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, List, Optional, Tuple


# --------------------------------------------------------------------------
# overlays
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Overlay:
    """A Beamer overlay specification such as ``<2->`` or ``<1,3-4>``."""

    ranges: Tuple[Tuple[int, Optional[int]], ...] = ()

    def visible_at(self, step: int) -> bool:
        if not self.ranges:
            return True
        for start, end in self.ranges:
            if step >= start and (end is None or step <= end):
                return True
        return False

    def max_step(self) -> int:
        top = 1
        for start, end in self.ranges:
            top = max(top, start if end is None else end)
        return top


ALWAYS = Overlay()


# --------------------------------------------------------------------------
# inline content
# --------------------------------------------------------------------------


@dataclass
class Run:
    """A styled span of text, or a piece of inline math when ``math`` is set."""

    text: str = ""
    bold: bool = False
    italic: bool = False
    mono: bool = False
    color: Optional[str] = None          # "RRGGBB"
    hyperlink: Optional[str] = None
    math: Optional[str] = None           # inline LaTeX, ``text`` then unused
    size_scale: float = 1.0
    overlay: Overlay = ALWAYS

    @property
    def is_math(self) -> bool:
        return self.math is not None


def plain(text: str) -> List[Run]:
    return [Run(text=text)]


def runs_to_text(runs: List[Run]) -> str:
    return "".join(r.text if not r.is_math else f"${r.math}$" for r in runs)


# --------------------------------------------------------------------------
# blocks
# --------------------------------------------------------------------------


@dataclass
class Block:
    overlay: Overlay = ALWAYS


@dataclass
class Bullet:
    runs: List[Run] = field(default_factory=list)
    level: int = 0
    marker: Optional[str] = None         # explicit label (description lists)
    ordered: bool = False
    overlay: Overlay = ALWAYS
    blocks: List[Block] = field(default_factory=list)   # nested non-text content


@dataclass
class BulletList(Block):
    items: List[Bullet] = field(default_factory=list)


@dataclass
class Paragraph(Block):
    runs: List[Run] = field(default_factory=list)
    align: str = "left"                  # left | center | right
    space_before: bool = True


@dataclass
class Equation(Block):
    latex: str = ""
    display: bool = True
    number: Optional[str] = None
    lines: List[str] = field(default_factory=list)   # align/gather rows
    align: str = "center"
    #: The converted ``m:oMath`` element, when native conversion succeeded.
    omml: Any = None
    # Filled in by the render stage when the formula falls back to an image.
    rendered_path: Optional[str] = None
    vector_path: Optional[str] = None
    rendered_size: Optional[Tuple[int, int]] = None
    rendered_height_pt: Optional[float] = None


@dataclass
class Figure(Block):
    """An image to place on the slide.

    ``source_path`` is the file named in the document (already resolved when
    ``kind == 'image'``); ``tex`` holds the snippet to compile for TikZ.
    """

    kind: str = "image"                  # image | tikz
    source_path: Optional[str] = None
    tex: Optional[str] = None
    width_frac: Optional[float] = None   # fraction of the text width
    height_frac: Optional[float] = None
    scale: Optional[float] = None
    angle: float = 0.0
    caption: List[Run] = field(default_factory=list)
    # filled in by the render stage
    rendered_path: Optional[str] = None
    #: Vector twin of ``rendered_path``; PowerPoint displays this and
    #: keeps the PNG only as a fallback.
    vector_path: Optional[str] = None
    native_size: Optional[Tuple[int, int]] = None


@dataclass
class Cell:
    runs: List[Run] = field(default_factory=list)
    colspan: int = 1
    rowspan: int = 1
    align: str = "l"
    bold: bool = False
    fill: Optional[str] = None
    merged: bool = False                 # continuation of a span, not emitted
    top_rule: bool = False
    bottom_rule: bool = False
    blocks: List[Block] = field(default_factory=list)


@dataclass
class Table(Block):
    rows: List[List[Cell]] = field(default_factory=list)
    col_aligns: List[str] = field(default_factory=list)
    col_weights: List[float] = field(default_factory=list)
    caption: List[Run] = field(default_factory=list)
    header_rows: int = 0
    vlines: List[bool] = field(default_factory=list)


@dataclass
class Column:
    width_frac: float = 0.5
    blocks: List[Block] = field(default_factory=list)


@dataclass
class Columns(Block):
    columns: List[Column] = field(default_factory=list)


@dataclass
class BeamerBlock(Block):
    kind: str = "block"                  # block | alertblock | exampleblock
    title: List[Run] = field(default_factory=list)
    blocks: List[Block] = field(default_factory=list)


@dataclass
class Verbatim(Block):
    text: str = ""
    language: Optional[str] = None


@dataclass
class Placeholder(Block):
    """Something we could not convert; rendered as a visible marker."""

    label: str = "Unconverted content"
    detail: str = ""


# --------------------------------------------------------------------------
# slides and deck
# --------------------------------------------------------------------------


@dataclass
class Slide:
    kind: str = "content"                # title | section | content
    title: List[Run] = field(default_factory=list)
    subtitle: List[Run] = field(default_factory=list)
    blocks: List[Block] = field(default_factory=list)
    notes: str = ""
    source_line: Optional[int] = None
    warnings: List[str] = field(default_factory=list)
    max_overlay: int = 1
    overlay_step: Optional[int] = None   # set when frames are expanded
    shrink: float = 1.0                  # explicit [shrink=N] hint


@dataclass
class Deck:
    title: List[Run] = field(default_factory=list)
    subtitle: List[Run] = field(default_factory=list)
    author: List[Run] = field(default_factory=list)
    institute: List[Run] = field(default_factory=list)
    date: List[Run] = field(default_factory=list)
    slides: List[Slide] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
