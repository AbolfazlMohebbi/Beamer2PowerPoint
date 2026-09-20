r"""Document-level parsing: metadata, sections and frames.

Turns the flattened source produced by :mod:`beamer2pptx.preprocess` into a
:class:`~beamer2pptx.ir.Deck`.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import List, Optional, Sequence

from ..ir import Deck, Run, Slide
from ..preprocess import Source
from .blocks import BuildState, build_blocks
from .inline import InlineContext, render_runs, render_text
from .nodes import EnvNode, GroupNode, MacroNode, MathNode, Node, TextNode, scan
from .overlays import expand_slide, flatten_slide

log = logging.getLogger(__name__)

METADATA_MACROS = {"title", "subtitle", "author", "date", "institute"}


@dataclass
class ParseOptions:
    section_slides: bool = True
    subsection_slides: bool = False
    overlays: str = "flatten"           # flatten | expand
    alert_color: str = "C00000"


def parse_document(source: Source, options: Optional[ParseOptions] = None,
                   ctx: Optional[InlineContext] = None) -> Deck:
    """Parse a preprocessed :class:`Source` into a :class:`Deck`."""
    options = options or ParseOptions()
    ctx = ctx or InlineContext(verbatims=source.verbatims,
                               alert_color=options.alert_color)
    ctx.verbatims = source.verbatims

    deck = Deck()
    _read_metadata(scan(source.preamble), deck, ctx)

    body_nodes = scan(source.body)
    _read_metadata(body_nodes, deck, ctx)

    slides: List[Slide] = []
    for node in body_nodes:
        _walk_top_level(node, deck, slides, options, ctx)

    # A frame holding \titlepage inherits the document metadata.
    for slide in slides:
        if slide.kind == "title":
            if not slide.title:
                slide.title = list(deck.title)
            if not slide.subtitle:
                slide.subtitle = list(deck.subtitle)

    # Overlay handling happens once the whole slide is built.
    final: List[Slide] = []
    for slide in slides:
        if options.overlays == "expand":
            final.extend(expand_slide(slide))
        else:
            final.append(flatten_slide(slide))
    deck.slides = final

    if ctx.unknown_macros:
        top = sorted(ctx.unknown_macros.items(), key=lambda kv: -kv[1])[:12]
        deck.warnings.append(
            "Macros rendered as plain text: "
            + ", ".join("\\%s (x%d)" % (name, count) for name, count in top)
        )
    deck.warnings.extend(ctx.warnings)
    return deck


# --------------------------------------------------------------------------
# metadata
# --------------------------------------------------------------------------


def _read_metadata(nodes: Sequence[Node], deck: Deck, ctx: InlineContext) -> None:
    for node in nodes:
        if isinstance(node, MacroNode) and node.name in METADATA_MACROS:
            runs = render_runs(scan(node.arg(0) or ""), ctx)
            if not runs:
                continue
            if node.name == "title" and not deck.title:
                deck.title = runs
            elif node.name == "subtitle" and not deck.subtitle:
                deck.subtitle = runs
            elif node.name == "author" and not deck.author:
                deck.author = runs
            elif node.name == "date" and not deck.date:
                deck.date = runs
            elif node.name == "institute" and not deck.institute:
                deck.institute = runs


# --------------------------------------------------------------------------
# top-level walk
# --------------------------------------------------------------------------


def _walk_top_level(node: Node, deck: Deck, slides: List[Slide],
                    options: ParseOptions, ctx: InlineContext) -> None:
    if isinstance(node, EnvNode):
        if node.name in ("frame", "frame*"):
            slides.extend(_build_frame(node, ctx, options))
            return
        # Content wrapped in a group-like environment: keep looking for frames.
        for child in node.nodes:
            _walk_top_level(child, deck, slides, options, ctx)
        return

    if isinstance(node, MacroNode):
        if node.name == "frame":
            body = node.arg(0) or ""
            pseudo = EnvNode(name="frame", body=body, nodes=scan(body),
                             raw=node.raw, spec="", args=[])
            slides.extend(_build_frame(pseudo, ctx, options))
            return

        if node.name in ("section", "subsection", "section*", "subsection*"):
            title = render_runs(scan(node.arg(0) or ""), ctx)
            want = (options.section_slides if node.name.startswith("section")
                    else options.subsection_slides)
            if want and title:
                slides.append(Slide(kind="section", title=title,
                                    source_line=node.pos))
            return

        if node.name == "note" and slides:
            slides[-1].notes = (slides[-1].notes + "\n"
                                + render_text(scan(node.arg(0) or ""), ctx)).strip()
            return

        if node.name in ("maketitle", "titlepage"):
            slides.append(_title_slide(deck))
            return

        if node.name == "againframe":
            return

    if isinstance(node, GroupNode):
        for child in node.nodes:
            _walk_top_level(child, deck, slides, options, ctx)


def _title_slide(deck: Deck) -> Slide:
    return Slide(kind="title", title=list(deck.title), subtitle=list(deck.subtitle))


# --------------------------------------------------------------------------
# frames
# --------------------------------------------------------------------------


def _parse_frame_options(raw: Optional[str]) -> dict:
    out: dict = {}
    if not raw:
        return out
    for item in raw.split(","):
        item = item.strip()
        if not item:
            continue
        if "=" in item:
            key, _, value = item.partition("=")
            out[key.strip().lower()] = value.strip()
        else:
            out[item.lower()] = True
    return out


def _build_frame(node: EnvNode, ctx: InlineContext,
                 options: ParseOptions) -> List[Slide]:
    frame_options = _parse_frame_options(node.opt(0))
    children = list(node.nodes)

    title: List[Run] = []
    subtitle: List[Run] = []
    notes: List[str] = []

    has_frametitle = any(isinstance(c, MacroNode) and c.name == "frametitle"
                         for c in children)

    # \begin{frame}{Title}{Subtitle}: leading groups carry the titles.
    if not has_frametitle:
        while children and isinstance(children[0], TextNode) and not children[0].text.strip():
            children.pop(0)
        if children and isinstance(children[0], GroupNode) and _looks_like_title(children[0]):
            title = render_runs(children.pop(0).nodes, ctx)
            while children and isinstance(children[0], TextNode) and not children[0].text.strip():
                children.pop(0)
            if children and isinstance(children[0], GroupNode) and _looks_like_title(children[0]):
                subtitle = render_runs(children.pop(0).nodes, ctx)

    # Pull out \frametitle, \framesubtitle, \note and a bare \titlepage.
    is_title_slide = False
    remaining: List[Node] = []
    for child in children:
        if isinstance(child, MacroNode):
            if child.name == "frametitle":
                title = render_runs(scan(child.arg(0) or ""), ctx)
                continue
            if child.name == "framesubtitle":
                subtitle = render_runs(scan(child.arg(0) or ""), ctx)
                continue
            if child.name == "note":
                notes.append(render_text(scan(child.arg(0) or ""), ctx))
                continue
            if child.name in ("titlepage", "maketitle"):
                is_title_slide = True
                continue
        remaining.append(child)

    state = BuildState(ctx=ctx)
    blocks = build_blocks(remaining, state)

    slide = Slide(kind="title" if is_title_slide else "content",
                  title=title, subtitle=subtitle, blocks=blocks,
                  notes="\n".join(n for n in notes if n).strip(),
                  source_line=node.pos)

    shrink = frame_options.get("shrink")
    if shrink is True:
        slide.shrink = 0.9
    elif isinstance(shrink, str):
        try:
            slide.shrink = max(0.4, 1.0 - float(shrink) / 100.0)
        except ValueError:
            pass

    if frame_options.get("plain"):
        slide.kind = "content"

    if not slide.blocks and not slide.title and not is_title_slide:
        return []

    return [slide]


def _looks_like_title(group: GroupNode) -> bool:
    """Guard the ``\\begin{frame}{Title}`` heuristic against real content."""
    for child in group.nodes:
        if isinstance(child, EnvNode):
            return False
        if isinstance(child, MacroNode) and child.name in ("includegraphics", "item",
                                                           "beamerverbblock"):
            return False
        if isinstance(child, MathNode) and child.display:
            return False
    return len(group.raw) <= 300
