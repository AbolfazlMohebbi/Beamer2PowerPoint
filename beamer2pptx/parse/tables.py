r"""LaTeX ``tabular`` bodies into :class:`~beamer2pptx.ir.Table`.

Handles the column specification, ``\\`` row breaks, ``&`` cell separators,
``\multicolumn``/``\multirow`` spans, and the ``\hline``/booktabs rules, all
without touching the inline renderer used for the cell contents.
"""

from __future__ import annotations

import logging
import re
from typing import List, Optional, Tuple

from ..ir import Cell, Run, Table
from ..preprocess import find_group, find_optional
from .inline import InlineContext, render_runs, resolve_color
from .nodes import scan

log = logging.getLogger(__name__)

#: Rule macros recognised at the start of a row.
RULE_MACROS = ("hline", "toprule", "midrule", "bottomrule", "cmidrule",
               "cline", "specialrule", "addlinespace", "morecmidrules",
               "hdashline", "cdashline")

#: A TeX length: an optional number, then a unit that may be a macro
#: (``0.5\textwidth``, ``3cm``, or a bare ``\textwidth`` meaning 1.0).
_LENGTH_RE = re.compile(
    r"(-?[\d.]*)\s*\\?(pt|mm|cm|in|em|ex|bp|sp|textwidth|linewidth"
    r"|columnwidth|paperwidth|textheight|paperheight)\b")

#: TeX lengths in points; relative units resolve against the text width later.
_UNIT_POINTS = {"pt": 1.0, "bp": 1.00375, "mm": 2.845276, "cm": 28.45276,
                "in": 72.27, "em": 10.0, "ex": 4.5, "sp": 1.0 / 65536}


def parse_length(text: str, text_width_pt: float = 345.0) -> Optional[float]:
    """Convert a TeX length to points; ``None`` when it is not a length."""
    if not text:
        return None
    m = _LENGTH_RE.search(text.replace(" ", ""))
    if not m:
        return None
    raw = m.group(1)
    try:
        value = float(raw) if raw not in ("", "-", ".") else 1.0
    except ValueError:
        return None
    unit = m.group(2)
    if unit in ("textwidth", "linewidth", "columnwidth", "paperwidth"):
        return value * text_width_pt
    if unit in ("textheight", "paperheight"):
        return value * text_width_pt * 0.75
    return value * _UNIT_POINTS.get(unit, 1.0)


def parse_colspec(spec: str) -> Tuple[List[str], List[float], List[bool]]:
    """Split a column specification into alignments, widths and vertical rules.

    Returns ``(aligns, widths, vlines)`` where a width of ``0`` means "size
    this column from its content" and *vlines* has one entry per gap
    (``len(aligns) + 1`` entries).
    """
    spec = _expand_star(spec or "")
    aligns: List[str] = []
    widths: List[float] = []
    vlines: List[bool] = [False]
    i = 0
    while i < len(spec):
        ch = spec[i]
        if ch == "|":
            vlines[-1] = True
            i += 1
            continue
        if ch in " \t\n":
            i += 1
            continue
        if ch in "@!>< ":
            # @{...}, >{...}, <{...}, !{...} decorations: skip their group.
            _, after = find_group(spec, i + 1)
            i = after if after != i + 1 else i + 1
            continue
        if ch in "lcr":
            aligns.append(ch)
            widths.append(0.0)
            vlines.append(False)
            i += 1
            continue
        if ch in "pmb":
            body, after = find_group(spec, i + 1)
            aligns.append("l" if ch == "p" else "l")
            widths.append(parse_length(body) or 0.0)
            vlines.append(False)
            i = after if after != i + 1 else i + 1
            continue
        if ch in "XY":                      # tabularx / tabulary
            aligns.append("l")
            widths.append(-1.0)             # flexible: share the leftover width
            vlines.append(False)
            i += 1
            continue
        if ch in "SdD":                     # siunitx / dcolumn
            _, after = find_group(spec, i + 1)
            aligns.append("r")
            widths.append(0.0)
            vlines.append(False)
            i = after if after != i + 1 else i + 1
            continue
        i += 1
    return aligns, widths, vlines


def _expand_star(spec: str) -> str:
    r"""Expand ``*{3}{c}`` repetitions in a column specification."""
    out = spec
    for _ in range(8):
        m = re.search(r"\*\s*\{\s*(\d+)\s*\}", out)
        if not m:
            break
        body, after = find_group(out, m.end())
        if after == m.end():
            break
        out = out[:m.start()] + body * int(m.group(1)) + out[after:]
    return out


def split_top_level(text: str, separator: str) -> List[str]:
    """Split *text* on *separator* at brace/bracket depth zero."""
    parts: List[str] = []
    depth = 0
    env_depth = 0
    i = 0
    start = 0
    n = len(text)
    sep_len = len(separator)
    while i < n:
        ch = text[i]
        if ch == "\\":
            if text.startswith(separator, i) and depth == 0 and env_depth == 0:
                parts.append(text[start:i])
                i += sep_len
                # Absorb a trailing [2pt] spacing argument.
                _, after = find_optional(text, i)
                i = after
                start = i
                continue
            if text.startswith("\\begin", i):
                env_depth += 1
                i += 6
                continue
            if text.startswith("\\end", i):
                env_depth = max(0, env_depth - 1)
                i += 4
                continue
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth = max(0, depth - 1)
        elif depth == 0 and env_depth == 0 and text.startswith(separator, i):
            parts.append(text[start:i])
            i += sep_len
            start = i
            continue
        i += 1
    parts.append(text[start:])
    return parts


def _skip_trim_spec(text: str, start: int) -> int:
    r"""Skip a booktabs trim specifier such as the ``(lr)`` in ``\cmidrule(lr)``."""
    i = start
    while i < len(text) and text[i] in " \t":
        i += 1
    if i < len(text) and text[i] == "(":
        close = text.find(")", i)
        if close >= 0:
            return close + 1
    return start


def _strip_rules(text: str) -> Tuple[str, List[str]]:
    """Pull leading rule macros off a row, returning the rest and their names."""
    found: List[str] = []
    i = 0
    while i < len(text):
        while i < len(text) and text[i] in " \t\n\r":
            i += 1
        m = re.match(r"\\(" + "|".join(RULE_MACROS) + r")\b\*?", text[i:])
        if not m:
            break
        found.append(m.group(1))
        i += m.end()
        _, i = find_optional(text, i)
        i = _skip_trim_spec(text, i)
        if m.group(1) in ("cmidrule", "cline", "cdashline"):
            _, after = find_group(text, i)
            i = after
        elif m.group(1) == "specialrule":
            for _ in range(3):
                _, i = find_group(text, i)
        elif m.group(1) == "addlinespace":
            _, i = find_optional(text, i)
    return text[i:], found


def parse_table(body: str, colspec: str, ctx: InlineContext,
                caption: Optional[List[Run]] = None) -> Table:
    """Parse a ``tabular`` body into a :class:`Table`."""
    aligns, widths, vlines = parse_colspec(colspec)

    raw_rows = split_top_level(body, "\\\\")
    parsed: List[Tuple[List[str], List[str]]] = []
    trailing_rules: List[str] = []
    for raw in raw_rows:
        stripped, rules = _strip_rules(raw)
        if not stripped.strip():
            # Rules with no content belong to the row above.
            trailing_rules.extend(rules)
            continue
        parsed.append((split_top_level(stripped, "&"), rules))

    rows: List[List[Cell]] = []
    pending_rowspans: dict = {}
    header_rows = 0

    for row_index, (raw_cells, rules) in enumerate(parsed):
        if any(r in ("midrule", "hline") for r in rules) and row_index > 0 and header_rows == 0:
            header_rows = row_index
        top_rule = bool(rules)
        cells: List[Cell] = []
        col = 0
        for raw_cell in raw_cells:
            # Close out any column held open by a \multirow above.  Correct
            # LaTeX puts an empty placeholder cell there, which the span
            # should swallow; a document that omits it still lines up,
            # because a cell with content is carried past the span instead.
            consumed = False
            while pending_rowspans.get(col, 0) > 0:
                pending_rowspans[col] -= 1
                cells.append(Cell(merged=True, align=_align_at(aligns, col)))
                col += 1
                if not raw_cell.strip():
                    consumed = True
                    break
            if consumed:
                continue

            cell = _parse_cell(raw_cell, aligns, col, ctx)
            cell.top_rule = top_rule
            cells.append(cell)
            for _extra in range(1, cell.colspan):
                cells.append(Cell(merged=True, align=cell.align))
            if cell.rowspan > 1:
                pending_rowspans[col] = cell.rowspan - 1
            col += cell.colspan
        # A row may simply stop early; carry any remaining spans across.
        while pending_rowspans.get(col, 0) > 0:
            pending_rowspans[col] -= 1
            cells.append(Cell(merged=True, align=_align_at(aligns, col)))
            col += 1
        rows.append(cells)

    if rows and trailing_rules:
        for cell in rows[-1]:
            cell.bottom_rule = True

    ncols = max((len(r) for r in rows), default=len(aligns))
    if ncols > len(aligns):
        aligns += ["l"] * (ncols - len(aligns))
        widths += [0.0] * (ncols - len(widths))
    for row in rows:
        while len(row) < ncols:
            row.append(Cell(align=_align_at(aligns, len(row))))

    if header_rows == 0 and len(rows) > 1 and _looks_like_header(rows[0]):
        header_rows = 1
    for row in rows[:header_rows]:
        for cell in row:
            cell.bold = True

    return Table(rows=rows, col_aligns=aligns[:ncols], col_weights=widths[:ncols],
                 caption=caption or [], header_rows=header_rows,
                 vlines=vlines[:ncols + 1])


def _align_at(aligns: List[str], index: int) -> str:
    return aligns[index] if index < len(aligns) else "l"


def _looks_like_header(row: List[Cell]) -> bool:
    """A first row of short, non-numeric labels is almost certainly a header."""
    texts = ["".join(r.text for r in c.runs).strip() for c in row if not c.merged]
    if not texts or not all(texts):
        return False
    numeric = sum(1 for t in texts if re.fullmatch(r"[-+]?[\d.,%\s]+", t))
    return numeric == 0 and all(len(t) <= 40 for t in texts)


def _parse_cell(raw: str, aligns: List[str], col: int, ctx: InlineContext) -> Cell:
    text = raw.strip()
    align = _align_at(aligns, col)
    colspan = 1
    rowspan = 1
    fill: Optional[str] = None

    m = re.match(r"\s*\\cellcolor\s*(\[[^\]]*\])?\s*", text)
    if m:
        model = (m.group(1) or "").strip("[]") or None
        value, after = find_group(text, m.end())
        fill = resolve_color(value, model)
        text = text[after:].strip()

    m = re.match(r"\s*\\rowcolor\s*(\[[^\]]*\])?\s*", text)
    if m:
        _, after = find_group(text, m.end())
        text = text[after:].strip()

    m = re.match(r"\s*\\multicolumn\s*", text)
    if m:
        count, i = find_group(text, m.end())
        spec, i = find_group(text, i)
        content, i = find_group(text, i)
        try:
            colspan = max(1, int(count.strip()))
        except ValueError:
            colspan = 1
        sub_aligns, _, _ = parse_colspec(spec)
        align = sub_aligns[0] if sub_aligns else align
        text = content

    m = re.match(r"\s*\\multirow\s*", text)
    if m:
        _, i = find_optional(text, m.end())
        count, i = find_group(text, i)
        _, i = find_group(text, i)            # width argument
        content, i = find_group(text, i)
        try:
            rowspan = max(1, int(count.strip().lstrip("*") or 1))
        except ValueError:
            rowspan = 1
        text = content

    runs = render_runs(scan(text), ctx)
    return Cell(runs=runs, colspan=colspan, rowspan=rowspan, align=align, fill=fill)
