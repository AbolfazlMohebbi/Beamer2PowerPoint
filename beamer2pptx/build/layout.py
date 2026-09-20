r"""Measuring content so it fits the template's content region.

Beamer frames are usually denser than a PowerPoint template expects, so the
builder needs a reasonable estimate of how tall each block will render.  The
estimate uses real font metrics via Pillow, wrapping text at the placeholder
width exactly as PowerPoint will.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from typing import Dict, List, Sequence, Tuple

from ..ir import (BeamerBlock, Block, BulletList, Columns, Equation, Figure,
                  Paragraph, Placeholder, Run, Table, Verbatim)

log = logging.getLogger(__name__)

EMU_PER_POINT = 12700
EMU_PER_INCH = 914400

#: Extra indent per outline level, in points.
LEVEL_INDENT_PT = 27.0
#: Vertical gap between stacked blocks, in points.
BLOCK_GAP_PT = 10.0
#: Line spacing the template asks for (overridden from the layout when known).
DEFAULT_LINE_SPACING = 1.2
#: Space above each paragraph, as a fraction of its font size.  The builder
#: writes this value onto every paragraph it creates, so the estimate here and
#: what PowerPoint renders cannot drift apart.
SPACE_BEFORE_RATIO = 0.35
#: Safety margin on measured text: wrapping is estimated, not exact, and
#: overlapping shapes are a far worse failure than a little slack.
TEXT_SAFETY = 1.06


def pt_to_emu(points: float) -> int:
    return int(round(points * EMU_PER_POINT))


def emu_to_pt(emu: float) -> float:
    return emu / EMU_PER_POINT


# --------------------------------------------------------------------------
# font metrics
# --------------------------------------------------------------------------

#: Typeface name -> font file, for the fonts a Windows box actually ships.
_FONT_FILES = {
    "calibri": "calibri.ttf", "calibri light": "calibril.ttf",
    "arial": "arial.ttf", "helvetica": "arial.ttf",
    "times new roman": "times.ttf", "times": "times.ttf",
    "cambria": "cambria.ttc", "georgia": "georgia.ttf",
    "verdana": "verdana.ttf", "tahoma": "tahoma.ttf",
    "segoe ui": "segoeui.ttf", "consolas": "consola.ttf",
    "courier new": "cour.ttf", "trebuchet ms": "trebuc.ttf",
    "garamond": "GARA.TTF", "book antiqua": "BKANT.TTF",
}

_MONO_FONT = "consola.ttf"


@dataclass
class FontMetrics:
    """Width measurement for one typeface, cached at a reference size."""

    typeface: str = "Calibri"
    reference_px: int = 100
    _regular = None
    _bold = None
    _mono = None
    _cache: Dict[Tuple[str, bool, bool], float] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self._regular = self._load(self.typeface, bold=False)
        self._bold = self._load(self.typeface, bold=True) or self._regular
        self._mono = self._load("consolas", bold=False) or self._regular

    def _load(self, typeface: str, bold: bool):
        try:
            from PIL import ImageFont
        except ImportError:
            return None
        name = (typeface or "").strip().lower()
        candidates = []
        if name in _FONT_FILES:
            base = _FONT_FILES[name]
            if bold:
                stem, ext = os.path.splitext(base)
                candidates.append(stem + "b" + ext)
            candidates.append(base)
        collapsed = name.replace(" ", "")
        candidates += [collapsed + ("b" if bold else "") + ".ttf",
                       collapsed + ".ttf", "calibri.ttf", "arial.ttf"]
        fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        for candidate in candidates:
            for path in (os.path.join(fonts_dir, candidate), candidate):
                try:
                    return ImageFont.truetype(path, self.reference_px)
                except (OSError, IOError):
                    continue
        try:
            return ImageFont.load_default()
        except Exception:
            return None

    def width_pt(self, text: str, size_pt: float, bold: bool = False,
                 mono: bool = False) -> float:
        """Rendered width of *text* at *size_pt*, in points."""
        if not text:
            return 0.0
        font = self._mono if mono else (self._bold if bold else self._regular)
        if font is None:
            return len(text) * size_pt * 0.5
        key = (text, bold, mono)
        reference = self._cache.get(key)
        if reference is None:
            try:
                reference = float(font.getlength(text))
            except Exception:
                reference = len(text) * self.reference_px * 0.5
            if len(self._cache) < 20000:
                self._cache[key] = reference
        return reference * size_pt / self.reference_px


# --------------------------------------------------------------------------
# measuring
# --------------------------------------------------------------------------


@dataclass
class MeasureContext:
    metrics: FontMetrics
    font_sizes: Dict[int, float]
    line_spacing: float = DEFAULT_LINE_SPACING
    scale: float = 1.0
    #: Width, in points, available to a figure at full text width.
    text_width_pt: float = 400.0

    def size_for_level(self, level: int) -> float:
        base = self.font_sizes.get(min(level, 8), 18.0)
        return base * self.scale


def _run_text(run: Run) -> str:
    if run.is_math:
        # Rough: a formula renders about as wide as its source minus markup.
        stripped = run.math or ""
        for token in ("\\left", "\\right", "\\displaystyle", "{", "}", "\\,"):
            stripped = stripped.replace(token, "")
        return stripped.replace("\\", "")
    return run.text


def wrapped_line_count(runs: Sequence[Run], width_pt: float, size_pt: float,
                       metrics: FontMetrics) -> int:
    """Greedy word wrap, counting the lines a run sequence needs."""
    if width_pt <= 0:
        return 1
    lines = 1
    used = 0.0
    for run in runs:
        text = _run_text(run)
        if not text:
            continue
        run_size = size_pt * (run.size_scale or 1.0)
        for token in _tokenise(text):
            if token == "\n":
                lines += 1
                used = 0.0
                continue
            width = metrics.width_pt(token, run_size, run.bold, run.mono)
            if used > 0 and used + width > width_pt:
                lines += 1
                used = width if token.strip() else 0.0
            else:
                used += width
    return max(1, lines)


def _tokenise(text: str) -> List[str]:
    """Split into wrap units, keeping the spaces attached to their word."""
    tokens: List[str] = []
    current = ""
    for ch in text:
        if ch == "\n":
            if current:
                tokens.append(current)
                current = ""
            tokens.append("\n")
        elif ch == " ":
            current += ch
            tokens.append(current)
            current = ""
        else:
            current += ch
    if current:
        tokens.append(current)
    return tokens


def measure_block(block: Block, width_emu: int, ctx: MeasureContext) -> int:
    """Estimated rendered height of one block, in EMU."""
    width_pt = emu_to_pt(width_emu)

    if isinstance(block, BulletList):
        total = 0.0
        for item in block.items:
            size = ctx.size_for_level(item.level)
            indent = LEVEL_INDENT_PT * (item.level + 1)
            lines = wrapped_line_count(item.runs, max(40.0, width_pt - indent),
                                       size, ctx.metrics)
            total += lines * size * ctx.line_spacing * TEXT_SAFETY
            total += size * SPACE_BEFORE_RATIO
            for nested in item.blocks:
                total += emu_to_pt(measure_block(nested, width_emu - pt_to_emu(indent), ctx))
        return pt_to_emu(total)

    if isinstance(block, Paragraph):
        size = ctx.size_for_level(0)
        lines = wrapped_line_count(block.runs, width_pt, size, ctx.metrics)
        return pt_to_emu(lines * size * ctx.line_spacing * TEXT_SAFETY
                         + size * SPACE_BEFORE_RATIO)

    if isinstance(block, Equation):
        size = ctx.size_for_level(0)
        if block.rendered_height_pt:
            return pt_to_emu(block.rendered_height_pt + 6)
        lines = max(1, len(block.lines))
        return pt_to_emu(lines * size * 1.7 + 8)

    if isinstance(block, Figure):
        return pt_to_emu(figure_height_pt(block, width_pt, ctx))

    if isinstance(block, Table):
        return pt_to_emu(table_height_pt(block, width_pt, ctx))

    if isinstance(block, Columns):
        tallest = 0
        for column in block.columns:
            column_width = int(width_emu * column.width_frac)
            height = sum(measure_block(b, column_width, ctx) for b in column.blocks)
            height += pt_to_emu(BLOCK_GAP_PT) * max(0, len(column.blocks) - 1)
            tallest = max(tallest, height)
        return tallest

    if isinstance(block, BeamerBlock):
        size = ctx.size_for_level(0)
        title_height = size * 1.35 + 6 if block.title else 0.0
        inner_width = width_emu - pt_to_emu(20)
        body = sum(measure_block(b, inner_width, ctx) for b in block.blocks)
        body += pt_to_emu(BLOCK_GAP_PT) * max(0, len(block.blocks) - 1)
        return pt_to_emu(title_height + 10) + body

    if isinstance(block, Verbatim):
        size = ctx.size_for_level(0) * 0.7
        lines = block.text.count("\n") + 1
        return pt_to_emu(lines * size * 1.15 + 12)

    if isinstance(block, Placeholder):
        return pt_to_emu(ctx.size_for_level(0) * 2.2)

    return pt_to_emu(ctx.size_for_level(0) * ctx.line_spacing)


def figure_height_pt(figure: Figure, width_pt: float, ctx: MeasureContext) -> float:
    """Height a figure will occupy once scaled to fit *width_pt*."""
    native = figure.native_size or (0, 0)
    aspect = (native[1] / native[0]) if native[0] and native[1] else 0.7
    draw_width = width_pt
    if figure.width_frac:
        draw_width = min(width_pt, ctx.text_width_pt * figure.width_frac)
    elif figure.scale and native[0]:
        draw_width = min(width_pt, native[0] * 72.0 / 300.0 * figure.scale)
    elif native[0]:
        draw_width = min(width_pt, native[0] * 72.0 / 150.0)
    height = draw_width * aspect
    if figure.caption:
        height += ctx.size_for_level(0) * 0.8 * 1.3 + 4
    return height + 6


def table_height_pt(table: Table, width_pt: float, ctx: MeasureContext) -> float:
    """Height of a native table, allowing for wrapped cell text."""
    if not table.rows:
        return 0.0
    size = table_font_size(table, ctx)
    widths = column_widths_pt(table, width_pt, ctx)
    total = 0.0
    for row in table.rows:
        lines = 1
        for index, cell in enumerate(row):
            if cell.merged or not cell.runs:
                continue
            span = sum(widths[index:index + cell.colspan]) if widths else width_pt
            lines = max(lines, wrapped_line_count(cell.runs, max(20.0, span - 8),
                                                  size, ctx.metrics))
        total += lines * size * 1.25 + 6
    if table.caption:
        total += ctx.size_for_level(0) * 0.8 * 1.3 + 4
    return total + 4


def table_font_size(table: Table, ctx: MeasureContext) -> float:
    """Tables need to be a little smaller than body text to stay readable."""
    base = ctx.size_for_level(1)
    columns = max(1, len(table.col_aligns))
    if columns >= 7:
        base *= 0.65
    elif columns >= 5:
        base *= 0.8
    if len(table.rows) >= 10:
        base *= 0.85
    return max(8.0, base)


def column_widths_pt(table: Table, width_pt: float, ctx: MeasureContext) -> List[float]:
    """Distribute the available width across the table's columns."""
    count = len(table.col_aligns) or (len(table.rows[0]) if table.rows else 1)
    if count == 0:
        return []

    explicit = list(table.col_weights) + [0.0] * (count - len(table.col_weights))
    size = table_font_size(table, ctx)

    natural: List[float] = []
    for index in range(count):
        longest = 0.0
        for row in table.rows:
            if index >= len(row):
                continue
            cell = row[index]
            if cell.merged or cell.colspan > 1 or not cell.runs:
                continue
            text = "".join(_run_text(r) for r in cell.runs)
            longest = max(longest, ctx.metrics.width_pt(text, size, cell.bold, False))
        natural.append(max(24.0, min(longest + 12.0, width_pt * 0.6)))

    fixed_total = sum(w for w in explicit if w > 0)
    flexible = [i for i, w in enumerate(explicit) if w <= 0]
    widths = [0.0] * count
    for index, weight in enumerate(explicit):
        if weight > 0:
            widths[index] = weight

    remaining = max(0.0, width_pt - fixed_total)
    natural_total = sum(natural[i] for i in flexible) or 1.0
    for index in flexible:
        widths[index] = remaining * natural[index] / natural_total

    total = sum(widths) or 1.0
    return [w * width_pt / total for w in widths]


# --------------------------------------------------------------------------
# fitting
# --------------------------------------------------------------------------


def total_height(blocks: Sequence[Block], width_emu: int,
                 ctx: MeasureContext) -> int:
    if not blocks:
        return 0
    height = sum(measure_block(b, width_emu, ctx) for b in blocks)
    return height + pt_to_emu(BLOCK_GAP_PT) * (len(blocks) - 1)


def fit_scale(blocks: Sequence[Block], region_width: int, region_height: int,
              metrics: FontMetrics, font_sizes: Dict[int, float],
              min_font_pt: float = 14.0, line_spacing: float = DEFAULT_LINE_SPACING,
              text_width_pt: float = 400.0) -> Tuple[float, bool]:
    """Find the largest scale at which *blocks* fit the region.

    Returns the scale and whether it bottomed out at *min_font_pt*.
    """
    base = max(font_sizes.get(0, 28.0), 1.0)
    floor = max(0.3, min(1.0, min_font_pt / base))

    scale = 1.0
    while scale >= floor - 1e-9:
        ctx = MeasureContext(metrics=metrics, font_sizes=font_sizes,
                             line_spacing=line_spacing, scale=scale,
                             text_width_pt=text_width_pt)
        if total_height(blocks, region_width, ctx) <= region_height:
            return scale, False
        scale -= 0.05

    return floor, True
