r"""Emitting the PowerPoint deck.

Takes the IR (with every asset already resolved) and writes slides onto the
user's template: text into the template's own placeholders wherever possible,
native tables, pictures from ``figures/``, and native equations.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

from ..ir import (BeamerBlock, Block, Bullet, BulletList, Cell, Columns,
                  Deck, Equation, Figure, Paragraph, Placeholder, Run, Slide,
                  Table, Verbatim)
from ..render.math_omml import MathConverter
from . import oxml
from .layout import (BLOCK_GAP_PT, EMU_PER_POINT, SPACE_BEFORE_RATIO,
                     FontMetrics, MeasureContext, column_widths_pt, emu_to_pt,
                     fit_scale, measure_block, pt_to_emu, table_font_size)
from .template import BODY_TYPES, Region, TITLE_TYPES, TemplateBinding

log = logging.getLogger(__name__)

TEXT_BLOCKS = (Paragraph, BulletList, Verbatim, Placeholder)

#: Fallback colours when the theme does not provide enough accents.
DEFAULT_ACCENTS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]


@dataclass
class BuildOptions:
    min_font_pt: float = 14.0
    line_spacing: float = 1.2
    alert_color: str = "C00000"
    caption_scale: float = 0.72
    table_style_id: Optional[str] = None
    number_equations: bool = False
    figures_relative_to: Optional[str] = None


@dataclass
class BuildReport:
    slides: int = 0
    tables: int = 0
    pictures: int = 0
    vector_pictures: int = 0
    native_equations: int = 0
    image_equations: int = 0
    text_equations: int = 0
    shrunk: List[Tuple[int, float]] = field(default_factory=list)
    overflowed: List[int] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


class SlideBuilder:
    """Writes a :class:`Deck` onto a template."""

    def __init__(self, binding: TemplateBinding, math: MathConverter,
                 options: Optional[BuildOptions] = None):
        self.binding = binding
        self.math = math
        self.options = options or BuildOptions()
        self.metrics = FontMetrics(typeface=binding.theme_font)
        self.accents = binding.accent_colors or DEFAULT_ACCENTS
        self.report = BuildReport()
        self._equation_number = 0
        #: SVG path -> the package part holding it, shared across slides.
        self._svg_parts: dict = {}

    # -- entry point ------------------------------------------------------

    def build(self, deck: Deck):
        for index, slide in enumerate(deck.slides, start=1):
            try:
                self._build_slide(slide, index, deck)
            except Exception as exc:                 # keep converting
                log.exception("Slide %d failed to build", index)
                self.report.warnings.append("Slide %d failed: %s" % (index, exc))
                self._build_failure_slide(slide, exc)
            self.report.slides += 1
        return self.binding.presentation

    # -- slide kinds ------------------------------------------------------

    def _build_slide(self, slide: Slide, index: int, deck: Deck) -> None:
        if slide.kind == "title":
            self._build_title_slide(slide, deck)
        elif slide.kind == "section":
            self._build_section_slide(slide)
        else:
            self._build_content_slide(slide, index)

    def _build_title_slide(self, slide: Slide, deck: Deck) -> None:
        pptx_slide = self.binding.presentation.slides.add_slide(self.binding.title_layout)
        title_shape = _placeholder(pptx_slide, TITLE_TYPES)
        subtitle_shape = _placeholder(pptx_slide, ("SUBTITLE",)) or _placeholder(
            pptx_slide, BODY_TYPES)

        if title_shape is not None:
            self._fill_simple_text(title_shape, slide.title or deck.title,
                                   size=None)
        if subtitle_shape is not None:
            lines: List[List[Run]] = []
            if slide.subtitle:
                lines.append(slide.subtitle)
            if deck.author:
                lines.append(deck.author)
            if deck.institute:
                lines.append(deck.institute)
            if deck.date:
                lines.append(deck.date)
            if lines:
                self._fill_multiline(subtitle_shape, lines)
            else:
                oxml.remove_shape(subtitle_shape)
        self._write_notes(pptx_slide, slide)

    def _build_section_slide(self, slide: Slide) -> None:
        layout = self.binding.section_layout
        pptx_slide = self.binding.presentation.slides.add_slide(layout)
        title_shape = _placeholder(pptx_slide, TITLE_TYPES)
        if title_shape is not None:
            self._fill_simple_text(title_shape, slide.title, size=None)
        for shape in list(pptx_slide.placeholders):
            if shape is title_shape:
                continue
            if not shape.has_text_frame or not shape.text_frame.text.strip():
                oxml.remove_shape(shape)
        self._write_notes(pptx_slide, slide)

    def _build_failure_slide(self, slide: Slide, exc: Exception) -> None:
        pptx_slide = self.binding.presentation.slides.add_slide(self.binding.content_layout)
        title_shape = _placeholder(pptx_slide, TITLE_TYPES)
        if title_shape is not None:
            self._fill_simple_text(title_shape, slide.title or [Run(text="Slide")],
                                   size=None)
        body = _placeholder(pptx_slide, BODY_TYPES)
        if body is not None:
            frame = body.text_frame
            frame.clear()
            paragraph = frame.paragraphs[0]
            run = paragraph.add_run()
            run.text = "This frame could not be converted (%s)." % exc
            run.font.italic = True
            run.font.size = Pt(16)

    # -- content slides ---------------------------------------------------

    def _build_content_slide(self, slide: Slide, index: int) -> None:
        pptx_slide = self.binding.presentation.slides.add_slide(self.binding.content_layout)

        title_shape = _placeholder(pptx_slide, TITLE_TYPES)
        if title_shape is not None:
            if slide.title or slide.subtitle:
                self._fill_title(title_shape, slide)
            else:
                oxml.remove_shape(title_shape)

        body_shape = _placeholder(pptx_slide, BODY_TYPES)
        region = self.binding.content_region
        if body_shape is not None:
            region = Region(body_shape.left, body_shape.top,
                            body_shape.width, body_shape.height)

        blocks = [b for b in slide.blocks if not _is_empty(b)]
        if not blocks:
            if body_shape is not None:
                oxml.remove_shape(body_shape)
            self._write_notes(pptx_slide, slide)
            return

        scale, bottomed_out = fit_scale(
            blocks, region.width, region.height, self.metrics,
            self.binding.body_font_sizes, self.options.min_font_pt,
            self.options.line_spacing, emu_to_pt(region.width))
        scale *= slide.shrink
        if bottomed_out:
            self.report.overflowed.append(index)
            slide.warnings.append(
                "Content is denser than the template; text reduced to the "
                "%.0f pt floor." % self.options.min_font_pt)
        if scale < 0.999:
            self.report.shrunk.append((index, scale))

        ctx = MeasureContext(metrics=self.metrics,
                             font_sizes=self.binding.body_font_sizes,
                             line_spacing=self.options.line_spacing,
                             scale=scale, text_width_pt=emu_to_pt(region.width))

        self._place_blocks(pptx_slide, blocks, region, ctx, body_shape)
        self._write_notes(pptx_slide, slide)

    # -- block placement --------------------------------------------------

    def _place_blocks(self, pptx_slide, blocks: Sequence[Block], region: Region,
                      ctx: MeasureContext, body_shape=None) -> None:
        """Stack *blocks* down *region*, reusing the body placeholder once."""
        groups = _group_blocks(blocks)
        heights = [self._group_height(group, region.width, ctx) for group in groups]
        gap = pt_to_emu(BLOCK_GAP_PT)
        total = sum(heights) + gap * max(0, len(groups) - 1)

        top = region.top
        if total < region.height and len(groups) == 1 and groups[0][0] == "object":
            # A lone figure or table looks better centred.
            top = region.top + (region.height - total) // 2

        placeholder_used = False
        for (kind, group_blocks), height in zip(groups, heights):
            available = min(height, region.bottom - top)
            if available <= 0:
                break
            sub_region = Region(region.left, top, region.width, available)
            if kind == "text" and body_shape is not None and not placeholder_used:
                self._fill_body_placeholder(body_shape, group_blocks, sub_region, ctx)
                placeholder_used = True
            elif kind == "text":
                self._add_text_box(pptx_slide, group_blocks, sub_region, ctx)
            else:
                self._place_object(pptx_slide, group_blocks[0], sub_region, ctx)
            top += available + gap

        if body_shape is not None and not placeholder_used:
            oxml.remove_shape(body_shape)

    def _group_height(self, group: Tuple[str, List[Block]], width: int,
                      ctx: MeasureContext) -> int:
        blocks = group[1]
        height = sum(measure_block(b, width, ctx) for b in blocks)
        return height + pt_to_emu(BLOCK_GAP_PT) * max(0, len(blocks) - 1)

    def _place_object(self, pptx_slide, block: Block, region: Region,
                      ctx: MeasureContext) -> None:
        if isinstance(block, Figure):
            self._add_figure(pptx_slide, block, region, ctx)
        elif isinstance(block, Table):
            self._add_table(pptx_slide, block, region, ctx)
        elif isinstance(block, Equation):
            self._add_equation(pptx_slide, block, region, ctx)
        elif isinstance(block, Columns):
            self._add_columns(pptx_slide, block, region, ctx)
        elif isinstance(block, BeamerBlock):
            self._add_beamer_block(pptx_slide, block, region, ctx)
        else:
            self._add_text_box(pptx_slide, [block], region, ctx)

    def _add_columns(self, pptx_slide, block: Columns, region: Region,
                     ctx: MeasureContext) -> None:
        gap = pt_to_emu(10)
        fractions = [c.width_frac for c in block.columns]
        regions = region.split_horizontal(fractions, gap=gap)
        for column, sub_region in zip(block.columns, regions):
            inner = [b for b in column.blocks if not _is_empty(b)]
            if inner:
                self._place_blocks(pptx_slide, inner, sub_region, ctx, None)

    # -- text -------------------------------------------------------------

    def _fill_body_placeholder(self, shape, blocks: Sequence[Block],
                               region: Region, ctx: MeasureContext) -> None:
        shape.left, shape.top = Emu(region.left), Emu(region.top)
        shape.width, shape.height = Emu(region.width), Emu(region.height)
        frame = shape.text_frame
        frame.clear()
        frame.word_wrap = True
        self._write_blocks_into_frame(frame, blocks, ctx, inherit_bullets=True)
        oxml.set_autofit_scale(frame, min(1.0, ctx.scale))

    def _add_text_box(self, pptx_slide, blocks: Sequence[Block], region: Region,
                      ctx: MeasureContext):
        box = pptx_slide.shapes.add_textbox(Emu(region.left), Emu(region.top),
                                            Emu(region.width), Emu(region.height))
        frame = box.text_frame
        frame.word_wrap = True
        oxml.set_text_frame_margins(frame, 0, 0, 0, 0)
        self._write_blocks_into_frame(frame, blocks, ctx, inherit_bullets=False)
        return box

    def _write_blocks_into_frame(self, frame, blocks: Sequence[Block],
                                 ctx: MeasureContext, inherit_bullets: bool,
                                 default_color: Optional[str] = None) -> None:
        first = True
        for block in blocks:
            if isinstance(block, BulletList):
                for item in block.items:
                    paragraph = self._next_paragraph(frame, first)
                    first = False
                    self._write_bullet(paragraph, item, ctx, inherit_bullets,
                                       default_color)
                    for nested in item.blocks:
                        self._write_nested_text(frame, nested, ctx, item.level + 1,
                                                default_color)
            elif isinstance(block, Paragraph):
                paragraph = self._next_paragraph(frame, first)
                first = False
                oxml.set_bullet_none(paragraph)
                paragraph.alignment = _alignment(block.align)
                size = ctx.size_for_level(0)
                self._set_spacing(paragraph, size, ctx)
                self._add_runs(paragraph, block.runs, size, ctx, default_color)
            elif isinstance(block, Verbatim):
                size = max(8.0, ctx.size_for_level(0) * 0.7)
                for line in block.text.splitlines() or [""]:
                    paragraph = self._next_paragraph(frame, first)
                    first = False
                    oxml.set_bullet_none(paragraph)
                    paragraph.line_spacing = 1.0
                    paragraph.space_before = Pt(0)
                    paragraph.space_after = Pt(0)
                    run = paragraph.add_run()
                    run.text = line or " "
                    run.font.name = "Consolas"
                    run.font.size = Pt(size)
                    if default_color:
                        _set_color(run, default_color)
            elif isinstance(block, Placeholder):
                paragraph = self._next_paragraph(frame, first)
                first = False
                oxml.set_bullet_none(paragraph)
                run = paragraph.add_run()
                run.text = "[%s]" % block.label
                run.font.italic = True
                run.font.size = Pt(max(10.0, ctx.size_for_level(0) * 0.8))
                _set_color(run, "808080")
            else:
                paragraph = self._next_paragraph(frame, first)
                first = False
                oxml.set_bullet_none(paragraph)
                self._add_runs(paragraph, getattr(block, "runs", []),
                               ctx.size_for_level(0), ctx, default_color)

    def _write_nested_text(self, frame, block: Block, ctx: MeasureContext,
                           level: int, default_color: Optional[str] = None) -> None:
        """Nested equations and paragraphs inside a bullet."""
        paragraph = frame.add_paragraph()
        paragraph.level = min(level, 8)
        oxml.set_bullet_none(paragraph)
        if isinstance(block, Equation):
            self._write_equation_paragraph(paragraph, block, ctx)
        else:
            self._add_runs(paragraph, getattr(block, "runs", []),
                           ctx.size_for_level(level), ctx, default_color)

    def _next_paragraph(self, frame, first: bool):
        if first and frame.paragraphs and not frame.paragraphs[0].runs:
            return frame.paragraphs[0]
        return frame.add_paragraph()

    def _set_spacing(self, paragraph, size_pt: float, ctx: MeasureContext) -> None:
        """Pin the spacing the layout estimate assumed."""
        paragraph.line_spacing = ctx.line_spacing
        paragraph.space_before = Pt(size_pt * SPACE_BEFORE_RATIO)
        paragraph.space_after = Pt(0)

    def _write_bullet(self, paragraph, item: Bullet, ctx: MeasureContext,
                      inherit_bullets: bool,
                      default_color: Optional[str] = None) -> None:
        paragraph.level = min(item.level, 8)
        size = ctx.size_for_level(item.level)
        self._set_spacing(paragraph, size, ctx)

        if item.marker:
            oxml.set_bullet_none(paragraph)
            if not inherit_bullets:
                oxml.set_hanging_indent(paragraph, item.level, 0)
            run = paragraph.add_run()
            run.text = item.marker + "  "
            run.font.bold = True
            run.font.size = Pt(size)
            _set_color(run, self.accents[0])
        elif item.ordered:
            oxml.set_bullet_autonumber(paragraph, level=item.level,
                                       hanging=not inherit_bullets)
        elif not inherit_bullets:
            oxml.set_bullet_char(paragraph, "•" if item.level == 0 else "–",
                                 level=item.level)

        self._add_runs(paragraph, item.runs, size, ctx, default_color)

    # -- runs -------------------------------------------------------------

    def _add_runs(self, paragraph, runs: Sequence[Run], size_pt: float,
                  ctx: MeasureContext, default_color: Optional[str] = None) -> None:
        for run in runs:
            if run.is_math:
                self._add_math_run(paragraph, run, size_pt, default_color)
                continue
            text = run.text
            if not text:
                continue
            for index, piece in enumerate(text.split("\n")):
                if index:
                    _add_line_break(paragraph)
                if not piece:
                    continue
                pptx_run = paragraph.add_run()
                pptx_run.text = piece
                self._style_run(pptx_run, run, size_pt, default_color)

    def _add_math_run(self, paragraph, run: Run, size_pt: float,
                      default_color: Optional[str] = None) -> None:
        omml = self.math.to_omml(run.math or "") if self.math.available else None
        if omml is not None:
            oxml.style_omml(omml, color=run.color or default_color,
                            size_pt=size_pt)
            oxml.append_math(paragraph, omml, fallback_text=run.math or "",
                             display=False)
            self.report.native_equations += 1
            return
        pptx_run = paragraph.add_run()
        pptx_run.text = run.math or ""
        pptx_run.font.italic = True
        pptx_run.font.name = "Cambria Math"
        pptx_run.font.size = Pt(size_pt)
        _set_color(pptx_run, run.color or default_color)
        self.report.text_equations += 1

    def _style_run(self, pptx_run, run: Run, size_pt: float,
                   default_color: Optional[str] = None) -> None:
        font = pptx_run.font
        font.size = Pt(max(6.0, size_pt * (run.size_scale or 1.0)))
        font.bold = run.bold
        font.italic = run.italic
        if run.mono:
            font.name = "Consolas"
        # An autoshape inherits light text from the theme's shape style, so
        # text placed inside one needs its colour stated outright.
        _set_color(pptx_run, run.color or default_color)
        if run.hyperlink:
            try:
                pptx_run.hyperlink.address = run.hyperlink
            except Exception as exc:
                log.debug("Could not set hyperlink %s: %s", run.hyperlink, exc)

    # -- titles -----------------------------------------------------------

    def _fill_title(self, shape, slide: Slide) -> None:
        frame = shape.text_frame
        frame.clear()
        paragraph = frame.paragraphs[0]
        self._add_runs(paragraph, slide.title, _inherited_size(shape, 32.0),
                       MeasureContext(metrics=self.metrics,
                                      font_sizes=self.binding.body_font_sizes))
        if slide.subtitle:
            second = frame.add_paragraph()
            self._add_runs(second, slide.subtitle,
                           _inherited_size(shape, 32.0) * 0.6,
                           MeasureContext(metrics=self.metrics,
                                          font_sizes=self.binding.body_font_sizes))
        oxml.set_autofit_scale(frame, 1.0)

    def _fill_simple_text(self, shape, runs: Sequence[Run],
                          size: Optional[float]) -> None:
        frame = shape.text_frame
        frame.clear()
        paragraph = frame.paragraphs[0]
        for run in runs:
            if run.is_math:
                self._add_math_run(paragraph, run, size or 28.0)
                continue
            pptx_run = paragraph.add_run()
            pptx_run.text = run.text
            if size:
                pptx_run.font.size = Pt(size)
            pptx_run.font.bold = run.bold
            pptx_run.font.italic = run.italic
            if run.color:
                try:
                    pptx_run.font.color.rgb = RGBColor.from_string(run.color.upper())
                except ValueError:
                    pass

    def _fill_multiline(self, shape, lines: Sequence[Sequence[Run]]) -> None:
        # Read the inherited geometry before writing any of it: a placeholder
        # that inherits from the layout loses the rest of its position as soon
        # as one dimension is set explicitly.
        left, top = shape.left, shape.top
        width, height = shape.width, shape.height

        frame = shape.text_frame
        frame.clear()
        frame.word_wrap = True
        ctx = MeasureContext(metrics=self.metrics,
                             font_sizes=self.binding.body_font_sizes)
        base = _inherited_size(shape, 20.0)
        needed = 0.0
        for index, runs in enumerate(lines):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            oxml.set_bullet_none(paragraph)
            paragraph.alignment = PP_ALIGN.CENTER
            size = base if index == 0 else base * 0.85
            self._add_runs(paragraph, runs, size, ctx)
            needed += size * 1.3

        # A title-slide subtitle placeholder is usually sized for one line;
        # grow it downwards so author, institute and date are not clipped.
        wanted = pt_to_emu(needed + 8)
        if wanted > height and None not in (left, top, width, height):
            room = self.binding.slide_height - top - pt_to_emu(8)
            shape.left, shape.top = Emu(left), Emu(top)
            shape.width = Emu(width)
            shape.height = Emu(max(height, min(wanted, max(room, 1))))

    # -- figures ----------------------------------------------------------

    def _add_figure(self, pptx_slide, figure: Figure, region: Region,
                    ctx: MeasureContext) -> None:
        if not figure.rendered_path or not os.path.isfile(figure.rendered_path):
            self._add_missing_figure(pptx_slide, figure, region, ctx)
            return

        caption_height = 0
        if figure.caption:
            caption_height = pt_to_emu(ctx.size_for_level(0)
                                       * self.options.caption_scale * 1.35 + 4)

        box = Region(region.left, region.top, region.width,
                     max(pt_to_emu(20), region.height - caption_height))
        width, height = self._figure_size(figure, box, ctx)
        left = box.left + (box.width - width) // 2
        top = box.top

        picture = self._add_picture(pptx_slide, figure.rendered_path,
                                    figure.vector_path, left, top, width, height)
        if figure.angle:
            oxml.set_shape_rotation(picture, -figure.angle)

        if figure.caption:
            self._add_caption(pptx_slide, figure.caption,
                              Region(region.left, top + height, region.width,
                                     caption_height), ctx)

    def _add_picture(self, pptx_slide, png_path: str, svg_path: Optional[str],
                     left: int, top: int, width: int, height: int):
        """Place a picture, preferring its vector version when there is one."""
        picture = pptx_slide.shapes.add_picture(png_path, Emu(left), Emu(top),
                                                Emu(width), Emu(height))
        self.report.pictures += 1

        if svg_path and os.path.isfile(svg_path):
            try:
                with open(svg_path, "rb") as fh:
                    svg_bytes = fh.read()
            except OSError as exc:
                log.debug("Could not read %s: %s", svg_path, exc)
                return picture
            if oxml.add_svg_to_picture(picture, pptx_slide.part, svg_bytes,
                                       self._svg_parts, svg_path):
                self.report.vector_pictures += 1
        return picture

    def _figure_size(self, figure: Figure, box: Region,
                     ctx: MeasureContext) -> Tuple[int, int]:
        native = figure.native_size or (0, 0)
        aspect = (native[1] / native[0]) if native[0] and native[1] else 0.7

        if figure.width_frac:
            width = int(ctx.text_width_pt * figure.width_frac * EMU_PER_POINT)
        elif figure.scale and native[0]:
            width = pt_to_emu(native[0] * 72.0 / 300.0 * figure.scale)
        elif native[0]:
            width = pt_to_emu(native[0] * 72.0 / 150.0)
        else:
            width = box.width

        width = min(width, box.width)
        height = int(width * aspect)
        if height > box.height:
            height = box.height
            width = int(height / aspect) if aspect else box.width
        return max(1, width), max(1, height)

    def _add_missing_figure(self, pptx_slide, figure: Figure, region: Region,
                            ctx: MeasureContext) -> None:
        label = figure.source_path or ("TikZ picture" if figure.kind == "tikz"
                                       else "figure")
        height = min(region.height, pt_to_emu(64))
        shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Emu(region.left), Emu(region.top),
            Emu(region.width), Emu(height))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string("F2F2F2")
        shape.line.color.rgb = RGBColor.from_string("BFBFBF")
        shape.line.width = Pt(0.75)
        frame = shape.text_frame
        frame.word_wrap = True
        paragraph = frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        run = paragraph.add_run()
        run.text = "Missing graphic: %s" % label
        run.font.size = Pt(12)
        run.font.italic = True
        run.font.color.rgb = RGBColor.from_string("808080")

    def _add_caption(self, pptx_slide, runs: Sequence[Run], region: Region,
                     ctx: MeasureContext) -> None:
        box = pptx_slide.shapes.add_textbox(Emu(region.left), Emu(region.top),
                                            Emu(region.width), Emu(region.height))
        frame = box.text_frame
        frame.word_wrap = True
        oxml.set_text_frame_margins(frame, 0, 0, 0, 0)
        paragraph = frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        oxml.set_bullet_none(paragraph)
        size = ctx.size_for_level(0) * self.options.caption_scale
        for run in runs:
            if run.is_math:
                self._add_math_run(paragraph, run, size)
                continue
            pptx_run = paragraph.add_run()
            pptx_run.text = run.text
            pptx_run.font.size = Pt(max(8.0, size))
            pptx_run.font.italic = True

    # -- tables -----------------------------------------------------------

    def _add_table(self, pptx_slide, table: Table, region: Region,
                   ctx: MeasureContext) -> None:
        if not table.rows:
            return

        caption_height = 0
        if table.caption:
            caption_height = pt_to_emu(ctx.size_for_level(0)
                                       * self.options.caption_scale * 1.35 + 4)

        rows = len(table.rows)
        columns = max(len(r) for r in table.rows)
        body_height = max(pt_to_emu(20), region.height - caption_height)

        widths_pt = column_widths_pt(table, emu_to_pt(region.width), ctx)
        graphic_frame = pptx_slide.shapes.add_table(
            rows, columns, Emu(region.left), Emu(region.top),
            Emu(region.width), Emu(body_height))
        pptx_table = graphic_frame.table

        for index, width in enumerate(widths_pt[:columns]):
            pptx_table.columns[index].width = Emu(pt_to_emu(width))

        oxml.set_table_banding(pptx_table, first_row=table.header_rows > 0,
                               band_row=False)
        if self.options.table_style_id:
            oxml.set_table_style(pptx_table, self.options.table_style_id)

        size = table_font_size(table, ctx)
        row_height = pt_to_emu(size * 1.25 + 6)

        for row_index, row in enumerate(table.rows):
            pptx_table.rows[row_index].height = Emu(row_height)
            for column_index in range(columns):
                cell = row[column_index] if column_index < len(row) else Cell()
                pptx_cell = pptx_table.cell(row_index, column_index)
                self._fill_table_cell(pptx_cell, cell, size, ctx,
                                      header=row_index < table.header_rows)

        self._merge_table_cells(pptx_table, table)
        self._apply_table_rules(pptx_table, table)

        self.report.tables += 1

        if table.caption:
            self._add_caption(pptx_slide, table.caption,
                              Region(region.left, region.top + body_height,
                                     region.width, caption_height), ctx)

    def _fill_table_cell(self, pptx_cell, cell: Cell, size: float,
                         ctx: MeasureContext, header: bool) -> None:
        oxml.set_cell_margins(pptx_cell, 5.0, 5.0, 2.0, 2.0)
        pptx_cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        frame = pptx_cell.text_frame
        frame.word_wrap = True
        paragraph = frame.paragraphs[0]
        paragraph.alignment = _alignment({"l": "left", "c": "center",
                                          "r": "right"}.get(cell.align, "left"))
        oxml.set_bullet_none(paragraph)

        runs = list(cell.runs)
        if not runs:
            return
        for run in runs:
            if run.is_math:
                self._add_math_run(paragraph, run, size)
                continue
            pptx_run = paragraph.add_run()
            pptx_run.text = run.text
            pptx_run.font.size = Pt(size)
            pptx_run.font.bold = run.bold or header
            pptx_run.font.italic = run.italic
            if run.mono:
                pptx_run.font.name = "Consolas"
            if run.color:
                try:
                    pptx_run.font.color.rgb = RGBColor.from_string(run.color.upper())
                except ValueError:
                    pass

        if cell.fill:
            oxml.set_cell_fill(pptx_cell, cell.fill)

    def _merge_table_cells(self, pptx_table, table: Table) -> None:
        for row_index, row in enumerate(table.rows):
            for column_index, cell in enumerate(row):
                if cell.merged or (cell.colspan == 1 and cell.rowspan == 1):
                    continue
                last_row = min(row_index + cell.rowspan - 1, len(table.rows) - 1)
                last_column = min(column_index + cell.colspan - 1,
                                  len(pptx_table.columns) - 1)
                if last_row == row_index and last_column == column_index:
                    continue
                try:
                    pptx_table.cell(row_index, column_index).merge(
                        pptx_table.cell(last_row, last_column))
                except Exception as exc:
                    log.debug("Could not merge cells (%d,%d): %s",
                              row_index, column_index, exc)

    def _apply_table_rules(self, pptx_table, table: Table) -> None:
        """Draw the booktabs rules the source asked for."""
        columns = len(pptx_table.columns)
        for row_index, row in enumerate(table.rows):
            top = any(c.top_rule for c in row)
            bottom = any(c.bottom_rule for c in row)
            for column_index in range(columns):
                pptx_cell = pptx_table.cell(row_index, column_index)
                if top:
                    width = 1.5 if row_index == 0 else 0.75
                    oxml.set_cell_border(pptx_cell, "T", width, "404040")
                if bottom:
                    oxml.set_cell_border(pptx_cell, "B", 1.5, "404040")

    # -- equations --------------------------------------------------------

    def _add_equation(self, pptx_slide, equation: Equation, region: Region,
                      ctx: MeasureContext) -> None:
        if equation.omml is None and equation.rendered_path:
            self._add_equation_image(pptx_slide, equation, region, ctx)
            return

        box = pptx_slide.shapes.add_textbox(Emu(region.left), Emu(region.top),
                                            Emu(region.width), Emu(region.height))
        frame = box.text_frame
        frame.word_wrap = True
        oxml.set_text_frame_margins(frame, 0, 0, 0, 0)
        self._write_equation_paragraph(frame.paragraphs[0], equation, ctx)

    def _write_equation_paragraph(self, paragraph, equation: Equation,
                                  ctx: MeasureContext) -> None:
        oxml.set_bullet_none(paragraph)
        paragraph.alignment = PP_ALIGN.CENTER
        size = ctx.size_for_level(0)

        if equation.omml is not None:
            oxml.append_math(paragraph, equation.omml,
                             fallback_text=equation.latex, display=True,
                             align=equation.align)
            self.report.native_equations += 1
            _set_paragraph_font_size(paragraph, size)
            return

        run = paragraph.add_run()
        run.text = equation.latex
        run.font.italic = True
        run.font.name = "Cambria Math"
        run.font.size = Pt(size * 0.9)
        self.report.text_equations += 1

    def _add_equation_image(self, pptx_slide, equation: Equation, region: Region,
                            ctx: MeasureContext) -> None:
        size = equation.rendered_size or (0, 0)
        if not size[0]:
            return
        aspect = size[1] / size[0]
        # Render at the source resolution so the formula matches the body text.
        width = pt_to_emu(size[0] * 72.0 / 600.0)
        width = min(width, region.width)
        height = int(width * aspect)
        if height > region.height:
            height = region.height
            width = int(height / aspect)
        left = region.left + (region.width - width) // 2
        self._add_picture(pptx_slide, equation.rendered_path,
                          equation.vector_path, left, region.top, width, height)
        self.report.image_equations += 1

    # -- beamer blocks ----------------------------------------------------

    def _add_beamer_block(self, pptx_slide, block: BeamerBlock, region: Region,
                          ctx: MeasureContext) -> None:
        color = {"alertblock": self.options.alert_color,
                 "exampleblock": self.accents[min(5, len(self.accents) - 1)],
                 }.get(block.kind, self.accents[0])

        size = ctx.size_for_level(0)
        title_height = pt_to_emu(size * 1.35 + 6) if block.title else 0
        body_height = max(pt_to_emu(16), region.height - title_height)

        body = pptx_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Emu(region.left), Emu(region.top + title_height),
            Emu(region.width), Emu(body_height))
        body.fill.solid()
        body.fill.fore_color.rgb = RGBColor.from_string(_tint(color, 0.90))
        body.line.fill.background()
        body.shadow.inherit = False

        frame = body.text_frame
        frame.word_wrap = True
        frame.vertical_anchor = MSO_ANCHOR.TOP
        oxml.set_text_frame_margins(frame, pt_to_emu(10), pt_to_emu(10),
                                    pt_to_emu(4), pt_to_emu(4))
        frame.clear()
        self._write_blocks_into_frame(frame, block.blocks, ctx,
                                      inherit_bullets=False,
                                      default_color="1A1A1A")

        if block.title:
            bar = pptx_slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Emu(region.left), Emu(region.top),
                Emu(region.width), Emu(title_height))
            bar.fill.solid()
            bar.fill.fore_color.rgb = RGBColor.from_string(color)
            bar.line.fill.background()
            bar.shadow.inherit = False
            bar_frame = bar.text_frame
            oxml.set_text_frame_margins(bar_frame, pt_to_emu(10), pt_to_emu(10), 0, 0)
            bar_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            paragraph = bar_frame.paragraphs[0]
            oxml.set_bullet_none(paragraph)
            paragraph.alignment = PP_ALIGN.LEFT
            for run in block.title:
                if run.is_math:
                    self._add_math_run(paragraph, run, size, "FFFFFF")
                    continue
                pptx_run = paragraph.add_run()
                pptx_run.text = run.text
                pptx_run.font.bold = True
                pptx_run.font.size = Pt(size * 0.95)
                pptx_run.font.color.rgb = RGBColor.from_string("FFFFFF")

    # -- notes ------------------------------------------------------------

    def _write_notes(self, pptx_slide, slide: Slide) -> None:
        pieces = [slide.notes] if slide.notes else []
        pieces += slide.warnings
        text = "\n".join(p for p in pieces if p).strip()
        if not text:
            return
        try:
            pptx_slide.notes_slide.notes_text_frame.text = text
        except Exception as exc:
            log.debug("Could not write notes: %s", exc)


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------


def _group_blocks(blocks: Sequence[Block]) -> List[Tuple[str, List[Block]]]:
    """Run consecutive text blocks together; everything else stands alone."""
    groups: List[Tuple[str, List[Block]]] = []
    for block in blocks:
        if isinstance(block, TEXT_BLOCKS):
            if groups and groups[-1][0] == "text":
                groups[-1][1].append(block)
            else:
                groups.append(("text", [block]))
        else:
            groups.append(("object", [block]))
    return groups


def _placeholder(pptx_slide, wanted_types: Sequence[str]):
    for shape in pptx_slide.placeholders:
        try:
            name = str(shape.placeholder_format.type).split()[0]
        except Exception:
            continue
        if name in wanted_types:
            return shape
    return None


def _set_color(pptx_run, color: Optional[str]) -> None:
    if not color:
        return
    try:
        pptx_run.font.color.rgb = RGBColor.from_string(color.upper())
    except (ValueError, AttributeError):
        pass


def _alignment(align: str):
    return {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)


def _add_line_break(paragraph) -> None:
    from pptx.oxml.ns import qn

    paragraph._p.append(paragraph._p.makeelement(qn("a:br"), {}))


def _set_paragraph_font_size(paragraph, size_pt: float) -> None:
    """Set the size an equation inherits, via the paragraph end properties."""
    from pptx.oxml.ns import qn

    p_pr = paragraph._p.get_or_add_pPr()
    def_rpr = p_pr.find(qn("a:defRPr"))
    if def_rpr is None:
        def_rpr = p_pr.makeelement(qn("a:defRPr"), {})
        p_pr.append(def_rpr)
    def_rpr.set("sz", str(int(round(size_pt * 100))))


def _inherited_size(shape, default: float) -> float:
    """Best guess at the font size a placeholder will render at."""
    try:
        from pptx.oxml.ns import qn

        body = shape._element.find(qn("p:txBody"))
        if body is not None:
            for level_pr in body.iter(qn("a:lvl1pPr")):
                def_rpr = level_pr.find(qn("a:defRPr"))
                if def_rpr is not None and def_rpr.get("sz"):
                    return int(def_rpr.get("sz")) / 100.0
    except Exception:
        pass
    return default


def _is_empty(block: Block) -> bool:
    if isinstance(block, Paragraph):
        return not any(r.is_math or r.text.strip() for r in block.runs)
    if isinstance(block, BulletList):
        return not block.items
    if isinstance(block, Columns):
        return all(not c.blocks for c in block.columns)
    if isinstance(block, Table):
        return not block.rows
    return False


def _tint(color: str, amount: float) -> str:
    """Blend *color* towards white by *amount* (0 = unchanged, 1 = white)."""
    try:
        r = int(color[0:2], 16)
        g = int(color[2:4], 16)
        b = int(color[4:6], 16)
    except (ValueError, IndexError):
        return "F2F2F2"
    blend = lambda v: int(round(v + (255 - v) * amount))
    return "%02X%02X%02X" % (blend(r), blend(g), blend(b))
