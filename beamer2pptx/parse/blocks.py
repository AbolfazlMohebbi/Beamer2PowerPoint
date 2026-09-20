r"""Frame contents into IR blocks.

Walks the node list of one frame and produces the block structure the builder
consumes: bullet lists, tables, figures, equations, columns and Beamer blocks.
Anything it does not recognise still contributes its text, so no content is
silently lost.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import List, Optional, Sequence

from ..ir import (ALWAYS, BeamerBlock, Block, Bullet, BulletList, Column,
                  Columns, Equation, Figure, Overlay, Paragraph, Placeholder,
                  Run, Table, Verbatim)
from .inline import InlineContext, render_runs
from .nodes import (EnvNode, GroupNode, MacroNode, MathNode, Node, TextNode,
                    TABLE_ENVS, TIKZ_ENVS, scan)
from .overlays import parse_overlay
from .tables import parse_length, parse_table, split_top_level

log = logging.getLogger(__name__)

#: Default Beamer text width, used to turn absolute lengths into fractions.
TEXTWIDTH_PT = 307.0

LIST_ENVS = {"itemize", "enumerate", "description"}
FLOAT_ENVS = {"figure", "figure*", "table", "table*", "wrapfigure", "subfigure",
              "subtable", "minipage", "threeparttable", "adjustbox"}
THEOREM_ENVS = {"theorem": "Theorem", "lemma": "Lemma", "corollary": "Corollary",
                "definition": "Definition", "proposition": "Proposition",
                "example": "Example", "examples": "Examples", "proof": "Proof",
                "remark": "Remark", "claim": "Claim", "fact": "Fact",
                "observation": "Observation", "note": "Note"}
ALIGN_ENVS = {"center": "center", "flushleft": "left", "flushright": "right",
              "centering": "center"}


@dataclass
class BuildState:
    """Per-frame state threaded through the block builder."""

    ctx: InlineContext
    step: int = 1                       # advanced by \pause
    plus_counter: List[int] = field(default_factory=lambda: [0])
    level: int = 0
    align: str = "left"
    figures: List[Figure] = field(default_factory=list)
    equations: List[Equation] = field(default_factory=list)

    def current_overlay(self) -> Overlay:
        return ALWAYS if self.step <= 1 else Overlay(((self.step, None),))


def build_blocks(nodes: Sequence[Node], state: BuildState) -> List[Block]:
    """Convert *nodes* into a list of blocks."""
    blocks: List[Block] = []
    buffer: List[Node] = []

    def flush() -> None:
        if not buffer:
            return
        display_math: List[MathNode] = []
        saved = state.ctx.collect_display_math
        state.ctx.collect_display_math = display_math
        try:
            runs = render_runs(buffer, state.ctx)
        finally:
            state.ctx.collect_display_math = saved
        buffer.clear()
        if runs:
            blocks.append(_tag(Paragraph(runs=runs, align=state.align), state))
        for math in display_math:
            blocks.append(_tag(_equation(math), state))

    for node in nodes:
        if isinstance(node, TextNode):
            # A blank line ends the current paragraph.
            pieces = re.split(r"\n[ \t]*\n", node.text)
            for index, piece in enumerate(pieces):
                if index:
                    flush()
                if piece.strip():
                    buffer.append(TextNode(pos=node.pos, end=node.end,
                                           raw=piece, text=piece))
            continue

        if isinstance(node, MathNode):
            if node.display:
                flush()
                blocks.append(_tag(_equation(node), state))
            else:
                buffer.append(node)
            continue

        if isinstance(node, MacroNode):
            handled = _macro_block(node, state, blocks, flush)
            if not handled:
                buffer.append(node)
            continue

        if isinstance(node, EnvNode):
            handled = _environment_block(node, state, blocks, flush)
            if not handled:
                buffer.append(node)
            continue

        if isinstance(node, GroupNode):
            # A brace group may hold structure, e.g. {\centering \includegraphics...}
            if _has_structure(node.nodes):
                flush()
                blocks.extend(build_blocks(node.nodes, state))
            else:
                buffer.append(node)
            continue

    flush()
    return blocks


def _tag(block: Block, state: BuildState) -> Block:
    if block.overlay is ALWAYS:
        block.overlay = state.current_overlay()
    return block


def _has_structure(nodes: Sequence[Node]) -> bool:
    for node in nodes:
        if isinstance(node, EnvNode) and (node.name in LIST_ENVS or node.name in FLOAT_ENVS
                                          or node.name in TABLE_ENVS or node.name in TIKZ_ENVS
                                          or node.name in ("columns", "block", "alertblock",
                                                           "exampleblock")):
            return True
        if isinstance(node, MacroNode) and node.name in ("includegraphics", "beamerverbblock"):
            return True
        if isinstance(node, MathNode) and node.display:
            return True
    return False


# --------------------------------------------------------------------------
# macros that produce blocks
# --------------------------------------------------------------------------


def _macro_block(node: MacroNode, state: BuildState, blocks: List[Block],
                 flush) -> bool:
    name = node.name

    if name == "pause":
        flush()
        state.step += 1
        return True

    if name == "includegraphics":
        flush()
        blocks.append(_tag(_figure_from_includegraphics(node), state))
        return True

    if name == "beamerverbblock":
        flush()
        index = _int_or_none(node.arg(0))
        if index is not None and 0 <= index < len(state.ctx.verbatims):
            language, text = state.ctx.verbatims[index]
            blocks.append(_tag(Verbatim(text=text, language=language), state))
        return True

    if name in ("centering", "raggedright", "raggedleft"):
        flush()
        state.align = {"centering": "center", "raggedright": "left",
                       "raggedleft": "right"}[name]
        return True

    if name in ("only", "uncover", "visible", "action", "alt", "temporal"):
        inner = node.arg(0)
        sub_nodes = scan(inner or "")
        if not _has_structure(sub_nodes):
            return False
        flush()
        overlay = parse_overlay(node.overlay, state.plus_counter)
        sub_state = BuildState(ctx=state.ctx, step=state.step,
                               plus_counter=state.plus_counter, level=state.level,
                               align=state.align)
        for block in build_blocks(sub_nodes, sub_state):
            if block.overlay is ALWAYS or block.overlay == state.current_overlay():
                block.overlay = overlay
            blocks.append(block)
        return True

    if name == "onslide":
        flush()
        overlay = parse_overlay(node.overlay, state.plus_counter)
        if overlay.ranges:
            state.step = overlay.ranges[0][0]
        return True

    if name == "invisible":
        return True

    if name == "column":
        # A bare \column outside our columns handler: ignore the width.
        return True

    if name in ("vfill", "vspace", "newline", "medskip", "bigskip", "smallskip"):
        return False

    if name == "caption":
        flush()
        runs = render_runs(scan(node.arg(0) or ""), state.ctx)
        _attach_caption(blocks, runs)
        return True

    if name == "tableofcontents":
        flush()
        blocks.append(_tag(Placeholder(label="Table of contents",
                                       detail="Beamer \\tableofcontents"), state))
        return True

    return False


# --------------------------------------------------------------------------
# environments that produce blocks
# --------------------------------------------------------------------------


def _environment_block(node: EnvNode, state: BuildState, blocks: List[Block],
                       flush) -> bool:
    name = node.name

    if name in LIST_ENVS:
        flush()
        # Capture the overlay before building: a \pause inside the list
        # advances the step for what follows, not for the list itself.
        entry_overlay = state.current_overlay()
        bullet_list = _bullet_list(node, state)
        if bullet_list.overlay is ALWAYS:
            bullet_list.overlay = entry_overlay
        blocks.append(bullet_list)
        return True

    if name in TABLE_ENVS:
        flush()
        colspec = _table_colspec(node)
        blocks.append(_tag(parse_table(node.body, colspec, state.ctx), state))
        return True

    if name in TIKZ_ENVS or (name in ("axis", "semilogxaxis", "semilogyaxis",
                                      "loglogaxis", "groupplot")):
        flush()
        blocks.append(_tag(Figure(kind="tikz", tex=node.raw), state))
        return True

    if name in ("columns", "columns*"):
        flush()
        blocks.append(_tag(_columns(node, state), state))
        return True

    if name in ("block", "alertblock", "exampleblock"):
        flush()
        title = render_runs(scan(node.arg(0) or ""), state.ctx)
        inner = BuildState(ctx=state.ctx, step=state.step,
                           plus_counter=state.plus_counter, level=state.level)
        body = build_blocks(node.nodes, inner)
        overlay = parse_overlay(node.overlay, state.plus_counter) if node.overlay else ALWAYS
        block = BeamerBlock(kind=name, title=title, blocks=body)
        if overlay.ranges:
            block.overlay = overlay
        blocks.append(_tag(block, state))
        return True

    if name.rstrip("*") in THEOREM_ENVS:
        flush()
        label = THEOREM_ENVS[name.rstrip("*")]
        extra = node.opt(0)
        title_text = "%s (%s)" % (label, extra) if extra else label
        inner = BuildState(ctx=state.ctx, step=state.step,
                           plus_counter=state.plus_counter, level=state.level)
        blocks.append(_tag(BeamerBlock(kind="block", title=[Run(text=title_text)],
                                       blocks=build_blocks(node.nodes, inner)), state))
        return True

    if name in FLOAT_ENVS:
        flush()
        inner = BuildState(ctx=state.ctx, step=state.step,
                           plus_counter=state.plus_counter, level=state.level,
                           align=state.align)
        produced = build_blocks(node.nodes, inner)
        blocks.extend(produced)
        return True

    if name in ALIGN_ENVS:
        flush()
        inner = BuildState(ctx=state.ctx, step=state.step,
                           plus_counter=state.plus_counter, level=state.level,
                           align=ALIGN_ENVS[name])
        blocks.extend(build_blocks(node.nodes, inner))
        return True

    if name in ("overprint", "overlayarea"):
        flush()
        inner = BuildState(ctx=state.ctx, step=state.step,
                           plus_counter=state.plus_counter, level=state.level)
        blocks.extend(build_blocks(node.nodes, inner))
        return True

    if name in ("small", "footnotesize", "scriptsize", "tiny", "large", "Large"):
        return False

    if name in ("beamercolorbox", "quote", "quotation", "abstract"):
        flush()
        inner = BuildState(ctx=state.ctx, step=state.step,
                           plus_counter=state.plus_counter, level=state.level)
        blocks.extend(build_blocks(node.nodes, inner))
        return True

    return False


# --------------------------------------------------------------------------
# lists
# --------------------------------------------------------------------------


def _bullet_list(node: EnvNode, state: BuildState) -> BulletList:
    ordered = node.name == "enumerate"
    items: List[Bullet] = []
    level = state.level
    local_step = state.step

    # Split the environment body on \item.
    groups: List[tuple] = []           # (marker, nodes)
    current: List[Node] = []
    marker: Optional[str] = None
    overlay: Optional[str] = None
    started = False

    for child in node.nodes:
        if isinstance(child, MacroNode) and child.name == "item":
            if started:
                groups.append((marker, overlay, current))
            current = []
            marker = child.opt(0)
            overlay = child.overlay
            started = True
            continue
        if started:
            current.append(child)
    if started:
        groups.append((marker, overlay, current))

    for marker_text, overlay_spec, children in groups:
        item_overlay = (parse_overlay(overlay_spec, state.plus_counter)
                        if overlay_spec else ALWAYS)

        # \pause inside an item list advances the step for the items after it.
        leading_pauses = 0
        filtered: List[Node] = []
        for child in children:
            if isinstance(child, MacroNode) and child.name == "pause":
                leading_pauses += 1
                continue
            filtered.append(child)

        inner_state = BuildState(ctx=state.ctx, step=1,
                                 plus_counter=state.plus_counter, level=level + 1)
        nested_blocks: List[Block] = []
        inline_nodes: List[Node] = []
        for child in filtered:
            if _is_block_level(child):
                nested_blocks.append(child)
            else:
                inline_nodes.append(child)

        display_math: List[MathNode] = []
        saved = state.ctx.collect_display_math
        state.ctx.collect_display_math = display_math
        try:
            runs = render_runs(inline_nodes, state.ctx)
        finally:
            state.ctx.collect_display_math = saved

        bullet = Bullet(runs=runs, level=level, ordered=ordered,
                        marker=_marker_text(marker_text, state))
        if item_overlay.ranges:
            bullet.overlay = item_overlay
        elif local_step > 1:
            bullet.overlay = Overlay(((local_step, None),))

        for math in display_math:
            bullet.blocks.append(_equation(math))
        if nested_blocks:
            bullet.blocks.extend(build_blocks(nested_blocks, inner_state))

        items.append(bullet)
        local_step += leading_pauses

        # Nested lists become further items at a deeper level.
        promoted: List[Bullet] = []
        remaining: List[Block] = []
        for block in bullet.blocks:
            if isinstance(block, BulletList):
                promoted.extend(block.items)
            else:
                remaining.append(block)
        bullet.blocks = remaining
        items.extend(promoted)

    state.step = local_step
    return BulletList(items=items)


def _is_block_level(node: Node) -> bool:
    if isinstance(node, EnvNode):
        return (node.name in LIST_ENVS or node.name in TABLE_ENVS
                or node.name in TIKZ_ENVS or node.name in FLOAT_ENVS
                or node.name in ("columns", "block", "alertblock", "exampleblock")
                or node.name.rstrip("*") in THEOREM_ENVS)
    if isinstance(node, MacroNode):
        return node.name in ("includegraphics", "beamerverbblock")
    return False


#: Symbols commonly used as custom \item markers, as their Unicode glyph.
MARKER_SYMBOLS = {
    "star": "★", "bullet": "•", "cdot": "·", "circ": "◦", "ast": "∗",
    "dagger": "†", "ddagger": "‡", "checkmark": "✓", "times": "×",
    "rightarrow": "→", "Rightarrow": "⇒", "triangleright": "▹",
    "blacksquare": "■", "square": "□", "diamond": "◇", "heartsuit": "♥",
    "to": "→", "gg": "≫", "dash": "–", "minus": "−",
}


def _marker_text(marker: Optional[str], state: BuildState) -> Optional[str]:
    """Render an ``\\item[...]`` label, including a symbol given as math."""
    if not marker:
        return None
    pieces: List[str] = []
    for run in render_runs(scan(marker), state.ctx):
        if run.is_math:
            key = (run.math or "").strip().lstrip("\\")
            pieces.append(MARKER_SYMBOLS.get(key, key))
        else:
            pieces.append(run.text)
    return "".join(pieces).strip() or None


# --------------------------------------------------------------------------
# columns, figures, equations
# --------------------------------------------------------------------------


def _columns(node: EnvNode, state: BuildState) -> Columns:
    columns: List[Column] = []
    current: Optional[Column] = None
    pending: List[Node] = []

    def close() -> None:
        nonlocal pending
        if current is not None:
            inner = BuildState(ctx=state.ctx, step=state.step,
                               plus_counter=state.plus_counter, level=state.level)
            current.blocks = build_blocks(pending, inner)
            columns.append(current)
        pending = []

    for child in node.nodes:
        if isinstance(child, EnvNode) and child.name == "column":
            close()
            current = Column(width_frac=_width_fraction(child.arg(0)) or 0.5)
            inner = BuildState(ctx=state.ctx, step=state.step,
                               plus_counter=state.plus_counter, level=state.level)
            current.blocks = build_blocks(child.nodes, inner)
            columns.append(current)
            current = None
            continue
        if isinstance(child, MacroNode) and child.name == "column":
            close()
            current = Column(width_frac=_width_fraction(child.arg(0)) or 0.5)
            continue
        if current is not None:
            pending.append(child)
        elif not isinstance(child, TextNode) or child.text.strip():
            pending.append(child)
            if columns or current:
                continue
    close()

    if not columns:
        inner = BuildState(ctx=state.ctx, step=state.step,
                           plus_counter=state.plus_counter, level=state.level)
        columns = [Column(width_frac=1.0, blocks=build_blocks(node.nodes, inner))]

    total = sum(c.width_frac for c in columns) or 1.0
    if total > 1.05 or total < 0.95:
        for column in columns:
            column.width_frac /= total
    return Columns(columns=columns)


def _width_fraction(spec: str) -> Optional[float]:
    r"""Interpret a width argument such as ``0.45\textwidth`` as a fraction."""
    if not spec:
        return None
    spec = spec.strip()
    m = re.match(r"^\s*([\d.]*)\s*\\(?:textwidth|linewidth|columnwidth|paperwidth)\s*$", spec)
    if m:
        return float(m.group(1)) if m.group(1) else 1.0
    points = parse_length(spec, TEXTWIDTH_PT)
    if points:
        return max(0.05, min(1.0, points / TEXTWIDTH_PT))
    return None


def _figure_from_includegraphics(node: MacroNode) -> Figure:
    options = node.opt(0) or ""
    figure = Figure(kind="image", source_path=(node.arg(0) or "").strip())
    for key, value in _parse_keyvals(options):
        if key == "width":
            figure.width_frac = _width_fraction(value)
        elif key == "height":
            points = parse_length(value, TEXTWIDTH_PT)
            if points:
                figure.height_frac = points / (TEXTWIDTH_PT * 0.75)
        elif key == "scale":
            try:
                figure.scale = float(value)
            except ValueError:
                pass
        elif key == "angle":
            try:
                figure.angle = float(value)
            except ValueError:
                pass
    return figure


def _parse_keyvals(options: str) -> List[tuple]:
    pairs: List[tuple] = []
    depth = 0
    current = ""
    for ch in options:
        if ch in "{[":
            depth += 1
        elif ch in "}]":
            depth -= 1
        if ch == "," and depth == 0:
            pairs.append(current)
            current = ""
            continue
        current += ch
    pairs.append(current)
    out = []
    for item in pairs:
        if "=" in item:
            key, _, value = item.partition("=")
            out.append((key.strip().lower(), value.strip()))
        elif item.strip():
            out.append((item.strip().lower(), ""))
    return out


def _equation(node: MathNode) -> Equation:
    latex = node.latex.strip()
    lines: List[str] = []
    if node.env and node.env.rstrip("*") in ("align", "gather", "eqnarray",
                                             "flalign", "alignat", "multline"):
        # Split only at the top level: a matrix inside the formula has \\ of
        # its own that must stay where it is.
        lines = [part.strip() for part in split_top_level(latex, "\\\\")
                 if part.strip()]
    return Equation(latex=latex, display=True, lines=lines, number=None)


def _attach_caption(blocks: List[Block], runs: List[Run]) -> None:
    """Attach a caption to the most recent figure or table."""
    for block in reversed(blocks):
        if isinstance(block, (Figure, Table)) and not block.caption:
            block.caption = runs
            return
    blocks.append(Paragraph(runs=runs, align="center"))


def _table_colspec(node: EnvNode) -> str:
    mandatory = [a for a, s in zip(node.args, node.spec) if s == "m"]
    return (mandatory[-1] if mandatory else "") or ""


def _int_or_none(value: str) -> Optional[int]:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None
