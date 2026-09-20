r"""Turning a node list into styled :class:`~beamer2pptx.ir.Run` spans.

This is the one place that knows how ``\textbf``, ``\alert``, colours, links
and inline math become PowerPoint text runs, so table cells, bullets, captions
and titles all inherit the same behaviour for free.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field, replace
from typing import Dict, List, Optional, Sequence

from ..ir import ALWAYS, Overlay, Run
from .nodes import (EnvNode, GroupNode, MacroNode, MathNode, Node, TextNode)
from .overlays import parse_overlay

log = logging.getLogger(__name__)

#: LaTeX/xcolor names we can resolve without reading the document preamble.
NAMED_COLORS: Dict[str, str] = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "00A000",
    "blue": "0000FF", "cyan": "00FFFF", "magenta": "FF00FF", "yellow": "FFD700",
    "orange": "FF7F00", "purple": "800080", "violet": "8000FF", "brown": "996633",
    "gray": "808080", "grey": "808080", "darkgray": "404040", "darkgrey": "404040",
    "lightgray": "BFBFBF", "lightgrey": "BFBFBF", "olive": "808000",
    "teal": "008080", "pink": "FFC0CB", "lime": "BFFF00", "navy": "000080",
    "maroon": "800000", "darkblue": "00008B", "darkgreen": "006400",
    "darkred": "8B0000",
}

#: Font-size macros, as a multiple of the surrounding size.
SIZE_MACROS: Dict[str, float] = {
    "tiny": 0.6, "scriptsize": 0.7, "footnotesize": 0.8, "small": 0.9,
    "normalsize": 1.0, "large": 1.15, "Large": 1.3, "LARGE": 1.45,
    "huge": 1.7, "Huge": 2.0,
}

#: Declarative style switches that apply to the rest of their group.
SWITCHES = {
    "bf": {"bold": True}, "bfseries": {"bold": True},
    "it": {"italic": True}, "itshape": {"italic": True},
    "sl": {"italic": True}, "slshape": {"italic": True},
    "tt": {"mono": True}, "ttfamily": {"mono": True},
    "em": {"italic": True}, "normalfont": {"bold": False, "italic": False, "mono": False},
    "rmfamily": {"mono": False}, "sffamily": {"mono": False},
    "mdseries": {"bold": False}, "upshape": {"italic": False},
}

#: Macros that wrap content in one style.
STYLE_MACROS = {
    "textbf": {"bold": True}, "textit": {"italic": True},
    "textsl": {"italic": True}, "emph": {"italic": True},
    "texttt": {"mono": True}, "textsc": {},
    "textrm": {"mono": False}, "textsf": {"mono": False},
    "textnormal": {}, "mbox": {}, "fbox": {}, "boldmath": {"bold": True},
    "underline": {"italic": False}, "uline": {}, "structure": {},
}

#: Macros whose sole mandatory argument is just unwrapped.
TRANSPARENT = {
    "mbox", "makebox", "text", "ensuremath", "protect", "relax",
    "centering", "raggedright", "raggedleft", "noindent", "leavevmode",
    "hfill", "hfil", "vfill", "medskip", "bigskip", "smallskip", "newpage",
    "clearpage", "toprule", "midrule", "bottomrule", "hline", "arraybackslash",
}

#: Dropped outright, arguments included.
DROPPED = {
    "label", "hypertarget", "index", "vspace", "hspace", "setlength",
    "addtolength", "captionsetup", "phantom", "vphantom", "hphantom",
    "usetheme", "usecolortheme", "setbeamertemplate", "setbeamercolor",
    "setbeamerfont", "addtocounter", "setcounter", "refstepcounter",
    "pagestyle", "thispagestyle", "nolinebreak", "linebreak", "allowbreak",
}


@dataclass
class Style:
    bold: bool = False
    italic: bool = False
    mono: bool = False
    color: Optional[str] = None
    size_scale: float = 1.0
    hyperlink: Optional[str] = None
    overlay: Overlay = ALWAYS

    def with_(self, **kwargs) -> "Style":
        return replace(self, **kwargs)


@dataclass
class InlineContext:
    """Shared state the inline renderer needs from the rest of the pipeline."""

    verbatims: Sequence = ()
    alert_color: str = "C00000"
    structure_color: Optional[str] = None
    warnings: List[str] = field(default_factory=list)
    unknown_macros: Dict[str, int] = field(default_factory=dict)
    #: Display-math nodes lifted out of the inline flow by the caller.
    collect_display_math: Optional[List[MathNode]] = None

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)


def resolve_color(spec: str, model: Optional[str] = None) -> Optional[str]:
    """Map an xcolor argument onto ``RRGGBB``."""
    if not spec:
        return None
    spec = spec.strip()
    if model:
        model = model.strip().upper()
        if model in ("HTML", "HTML6"):
            value = spec.lstrip("#").upper()
            return value if re.fullmatch(r"[0-9A-F]{6}", value) else None
        if model in ("RGB", "RGB256"):
            parts = [p.strip() for p in spec.split(",")]
            if len(parts) == 3:
                try:
                    return "".join("%02X" % max(0, min(255, int(float(p)))) for p in parts)
                except ValueError:
                    return None
        if model == "rgb" or model == "RGB1":
            parts = [p.strip() for p in spec.split(",")]
            if len(parts) == 3:
                try:
                    return "".join("%02X" % max(0, min(255, round(float(p) * 255))) for p in parts)
                except ValueError:
                    return None
        if model in ("GRAY", "gray"):
            try:
                v = max(0, min(255, round(float(spec) * 255)))
                return "%02X%02X%02X" % (v, v, v)
            except ValueError:
                return None
    # "red!40!black" and similar blends: take the dominant component.
    base = spec.split("!")[0].strip().lower()
    if base in NAMED_COLORS:
        return NAMED_COLORS[base]
    if re.fullmatch(r"[0-9A-Fa-f]{6}", spec):
        return spec.upper()
    return None


def render_runs(nodes: Sequence[Node], ctx: InlineContext,
                style: Optional[Style] = None) -> List[Run]:
    """Render *nodes* into runs, honouring nested styling."""
    style = style or Style()
    runs: List[Run] = []
    _render(nodes, ctx, style, runs)
    return merge_runs(runs)


def render_text(nodes: Sequence[Node], ctx: InlineContext) -> str:
    """Plain-text rendering, used for notes, alt text and log messages."""
    return "".join(r.text if not r.is_math else r.math or ""
                   for r in render_runs(nodes, ctx))


def _emit(runs: List[Run], text: str, style: Style) -> None:
    if not text:
        return
    runs.append(Run(text=text, bold=style.bold, italic=style.italic,
                    mono=style.mono, color=style.color,
                    hyperlink=style.hyperlink, size_scale=style.size_scale,
                    overlay=style.overlay))


def _render(nodes: Sequence[Node], ctx: InlineContext, style: Style,
            runs: List[Run]) -> Style:
    """Render *nodes*; returns the style in force at the end of the list."""
    for node in nodes:
        if isinstance(node, TextNode):
            text = _clean_text(node.text)
            _emit(runs, text, style)

        elif isinstance(node, MathNode):
            if node.display and ctx.collect_display_math is not None:
                ctx.collect_display_math.append(node)
            else:
                runs.append(Run(math=node.latex.strip(), bold=style.bold,
                                color=style.color, size_scale=style.size_scale,
                                overlay=style.overlay))

        elif isinstance(node, GroupNode):
            # A group scopes declarative switches.
            _render(node.nodes, ctx, style, runs)

        elif isinstance(node, MacroNode):
            style = _render_macro(node, ctx, style, runs)

        elif isinstance(node, EnvNode):
            style = _render_environment(node, ctx, style, runs)

    return style


def _clean_text(text: str) -> str:
    """Collapse LaTeX whitespace and convert quote ligatures."""
    text = text.replace("``", "“").replace("''", "”")
    text = re.sub(r"(?<![-])---(?![-])", "—", text)
    text = re.sub(r"(?<![-])--(?![-])", "–", text)
    # Blank lines are paragraph breaks; other newlines are plain spaces.
    text = re.sub(r"[ \t]*\n[ \t]*\n[ \t\n]*", "\n", text)
    text = re.sub(r"[ \t]*\n[ \t]*", " ", text)
    return re.sub(r"[ \t]{2,}", " ", text)


def _render_macro(node: MacroNode, ctx: InlineContext, style: Style,
                  runs: List[Run]) -> Style:
    name = node.name
    overlay = parse_overlay(node.overlay) if node.overlay else style.overlay

    if name == "\\":
        _emit(runs, "\n", style)
        return style

    if name in DROPPED:
        return style

    if name in SIZE_MACROS:
        return style.with_(size_scale=SIZE_MACROS[name])

    if name in SWITCHES:
        return style.with_(**SWITCHES[name])

    if name in ("beamerverbinline", "beamerverbblock"):
        index = _as_int(node.arg(0))
        if index is not None and 0 <= index < len(ctx.verbatims):
            _emit(runs, ctx.verbatims[index][1], style.with_(mono=True))
        return style

    if name == "alert":
        _render_arg(node, 0, ctx, style.with_(color=ctx.alert_color, bold=True,
                                              overlay=overlay), runs)
        return style

    if name == "structure":
        colour = ctx.structure_color or style.color
        _render_arg(node, 0, ctx, style.with_(color=colour), runs)
        return style

    if name in STYLE_MACROS:
        _render_arg(node, 0, ctx, style.with_(**STYLE_MACROS[name]), runs)
        return style

    if name == "textcolor":
        # \textcolor{name}{text} or \textcolor[model]{spec}{text}
        colour = resolve_color(node.arg(0), node.opt(0)) or style.color
        _render_arg(node, 1, ctx, style.with_(color=colour), runs)
        return style

    if name in ("color", "colorlet"):
        colour = resolve_color(node.arg(0), node.opt(0))
        return style.with_(color=colour or style.color)

    if name in ("colorbox", "fcolorbox"):
        _render_arg(node, len(node.spec.replace("o", "")) - 1, ctx, style, runs)
        return style

    if name == "url" or name == "nolinkurl":
        target = node.arg(0)
        _emit(runs, target, style.with_(mono=True, hyperlink=target))
        return style

    if name == "href":
        target = node.arg(0)
        inner = node.arg(1)
        sub = _scan_inline(inner)
        _render(sub, ctx, style.with_(hyperlink=target), runs)
        return style

    if name == "hyperlink":
        _render_arg(node, 1, ctx, style, runs)
        return style

    if name in ("ref", "eqref", "pageref", "autoref", "cref", "Cref"):
        _emit(runs, "[%s]" % node.arg(0).strip(), style)
        return style

    if name in ("cite", "citep", "citet", "citeauthor", "citeyear"):
        keys = node.arg(0).strip()
        note = node.opt(0)
        label = keys.replace(",", ", ")
        if note:
            label = "%s, %s" % (label, note)
        _emit(runs, "[%s]" % label, style)
        return style

    if name == "footnote":
        inner = _scan_inline(node.arg(0))
        _emit(runs, " (", style.with_(size_scale=style.size_scale * 0.75))
        _render(inner, ctx, style.with_(size_scale=style.size_scale * 0.75), runs)
        _emit(runs, ")", style.with_(size_scale=style.size_scale * 0.75))
        return style

    if name in ("only", "uncover", "visible", "action"):
        _render_arg(node, 0, ctx, style.with_(overlay=overlay), runs)
        return style

    if name == "invisible":
        return style

    if name == "alt":
        _render_arg(node, 0, ctx, style.with_(overlay=overlay), runs)
        return style

    if name == "temporal":
        _render_arg(node, 1, ctx, style.with_(overlay=overlay), runs)
        return style

    if name == "onslide":
        return style.with_(overlay=overlay)

    if name in ("textsuperscript", "textsubscript"):
        _render_arg(node, 0, ctx, style.with_(size_scale=style.size_scale * 0.7), runs)
        return style

    if name in ("resizebox", "scalebox", "raisebox", "rotatebox", "parbox",
                "framebox", "adjustbox", "makebox"):
        mandatory = [a for a, s in zip(node.args, node.spec) if s == "m"]
        if mandatory:
            _render(_scan_inline(mandatory[-1] or ""), ctx, style, runs)
        return style

    if name in TRANSPARENT:
        if "m" in node.spec:
            _render_arg(node, 0, ctx, style, runs)
        return style

    # Unknown macro: keep its argument text so no content silently vanishes.
    mandatory = [a for a, s in zip(node.args, node.spec) if s == "m"]
    if mandatory:
        for arg in mandatory:
            _render(_scan_inline(arg or ""), ctx, style, runs)
    ctx.unknown_macros[name] = ctx.unknown_macros.get(name, 0) + 1
    return style


def _render_environment(node: EnvNode, ctx: InlineContext, style: Style,
                        runs: List[Run]) -> Style:
    """Inline fallback for environments met in a text-only context."""
    if node.name in ("center", "flushleft", "flushright", "quote", "quotation"):
        _render(node.nodes, ctx, style, runs)
    elif node.name in ("small", "footnotesize", "scriptsize", "tiny", "large"):
        _render(node.nodes, ctx, style.with_(size_scale=SIZE_MACROS.get(node.name, 1.0)), runs)
    else:
        _render(node.nodes, ctx, style, runs)
    return style


def _render_arg(node: MacroNode, index: int, ctx: InlineContext, style: Style,
                runs: List[Run]) -> None:
    mandatory = [a for a, s in zip(node.args, node.spec) if s == "m"]
    if index < len(mandatory):
        _render(_scan_inline(mandatory[index] or ""), ctx, style, runs)


def _scan_inline(text: str) -> List[Node]:
    from .nodes import scan
    return scan(text)


def _as_int(value: str) -> Optional[int]:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def merge_runs(runs: List[Run]) -> List[Run]:
    """Join adjacent runs that share formatting, and trim the edges."""
    merged: List[Run] = []
    for run in runs:
        if run.is_math:
            merged.append(run)
            continue
        if not run.text:
            continue
        prev = merged[-1] if merged else None
        if (prev is not None and not prev.is_math
                and (prev.bold, prev.italic, prev.mono, prev.color, prev.hyperlink,
                     prev.size_scale, prev.overlay)
                == (run.bold, run.italic, run.mono, run.color, run.hyperlink,
                    run.size_scale, run.overlay)):
            prev.text += run.text
        else:
            merged.append(run)

    while merged and not merged[0].is_math and not merged[0].text.strip():
        merged.pop(0)
    while merged and not merged[-1].is_math and not merged[-1].text.strip():
        merged.pop()
    if merged and not merged[0].is_math:
        merged[0].text = merged[0].text.lstrip()
    if merged and not merged[-1].is_math:
        merged[-1].text = merged[-1].text.rstrip()
    return [r for r in merged if r.is_math or r.text]
