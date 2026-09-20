r"""Beamer overlay specifications.

Parses the ``<...>`` syntax (``<2>``, ``<2->``, ``<-3>``, ``<1,3-4>``,
``<+->``) into an :class:`~beamer2pptx.ir.Overlay`, and expands a slide into
one slide per overlay step when the user asks for it.
"""

from __future__ import annotations

import copy
import re
from typing import List, Optional, Tuple

from ..ir import (ALWAYS, BeamerBlock, Block, Bullet, BulletList, Columns,
                  Overlay, Run, Slide)

_RANGE_RE = re.compile(r"^\s*(\d*)\s*(-?)\s*(\d*)\s*$")


def parse_overlay(spec: Optional[str], counter: Optional[List[int]] = None) -> Overlay:
    """Parse an overlay specification into an :class:`Overlay`.

    *counter* is a one-element list holding the running value of Beamer's
    incremental ``+`` specifier; pass the same list across one frame so
    ``<+->`` advances the way Beamer does.
    """
    if spec is None:
        return ALWAYS
    spec = spec.strip()
    if not spec:
        return ALWAYS

    # Drop beamer action prefixes such as "alert@2-" or "@1".
    spec = re.sub(r"^[A-Za-z]+@", "", spec)

    ranges: List[Tuple[int, Optional[int]]] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "+" in part:
            if counter is not None:
                counter[0] += 1
                current = counter[0]
            else:
                current = 1
            part = part.replace("+", str(current))
        m = _RANGE_RE.match(part)
        if not m:
            continue
        start_s, dash, end_s = m.group(1), m.group(2), m.group(3)
        if not dash:
            if start_s:
                ranges.append((int(start_s), int(start_s)))
            continue
        start = int(start_s) if start_s else 1
        end = int(end_s) if end_s else None
        ranges.append((start, end))

    return Overlay(tuple(ranges)) if ranges else ALWAYS


def shift_overlay(overlay: Overlay, offset: int) -> Overlay:
    """Move every range in *overlay* later by *offset* steps."""
    if offset == 0 or not overlay.ranges:
        return overlay
    return Overlay(tuple((s + offset, None if e is None else e + offset)
                         for s, e in overlay.ranges))


# --------------------------------------------------------------------------
# slide expansion
# --------------------------------------------------------------------------


def max_step(slide: Slide) -> int:
    """Highest overlay step referenced anywhere in *slide*."""
    top = 1

    def visit_runs(runs: List[Run]) -> None:
        nonlocal top
        for run in runs:
            top = max(top, run.overlay.max_step())

    def visit_blocks(blocks: List[Block]) -> None:
        nonlocal top
        for block in blocks:
            top = max(top, block.overlay.max_step())
            if isinstance(block, BulletList):
                for item in block.items:
                    top = max(top, item.overlay.max_step())
                    visit_runs(item.runs)
                    visit_blocks(item.blocks)
            elif isinstance(block, Columns):
                for column in block.columns:
                    visit_blocks(column.blocks)
            elif isinstance(block, BeamerBlock):
                visit_runs(block.title)
                visit_blocks(block.blocks)
            elif hasattr(block, "runs"):
                visit_runs(getattr(block, "runs"))

    visit_blocks(slide.blocks)
    visit_runs(slide.title)
    return top


def _filter_runs(runs: List[Run], step: int) -> List[Run]:
    return [r for r in runs if r.overlay.visible_at(step)]


def _filter_blocks(blocks: List[Block], step: int) -> List[Block]:
    kept: List[Block] = []
    for block in blocks:
        if not block.overlay.visible_at(step):
            continue
        block = copy.deepcopy(block)
        if isinstance(block, BulletList):
            items: List[Bullet] = []
            for item in block.items:
                if not item.overlay.visible_at(step):
                    continue
                item.runs = _filter_runs(item.runs, step)
                item.blocks = _filter_blocks(item.blocks, step)
                items.append(item)
            block.items = items
            if not items:
                continue
        elif isinstance(block, Columns):
            for column in block.columns:
                column.blocks = _filter_blocks(column.blocks, step)
        elif isinstance(block, BeamerBlock):
            block.title = _filter_runs(block.title, step)
            block.blocks = _filter_blocks(block.blocks, step)
        elif hasattr(block, "runs"):
            filtered = _filter_runs(getattr(block, "runs"), step)
            if not filtered:
                continue
            setattr(block, "runs", filtered)
        kept.append(block)
    return kept


def expand_slide(slide: Slide) -> List[Slide]:
    """Return one slide per overlay step, or ``[slide]`` when there are none."""
    steps = max_step(slide)
    slide.max_overlay = steps
    if steps <= 1:
        return [slide]

    out: List[Slide] = []
    for step in range(1, steps + 1):
        copy_slide = copy.deepcopy(slide)
        copy_slide.blocks = _filter_blocks(slide.blocks, step)
        copy_slide.title = _filter_runs(copy.deepcopy(slide.title), step)
        copy_slide.overlay_step = step
        copy_slide.max_overlay = steps
        out.append(copy_slide)
    return out


def flatten_slide(slide: Slide) -> Slide:
    """Drop overlay information, keeping every piece of content visible."""
    slide.max_overlay = max_step(slide)
    slide.overlay_step = None
    return slide
