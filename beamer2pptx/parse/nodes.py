r"""A small, forgiving LaTeX scanner.

Produces a node tree over the flattened source.  Every node keeps the exact
slice of source it came from, which is what lets later stages hand raw LaTeX
(math, TikZ, tabular bodies) straight to a converter or to ``pdflatex``.

The scanner never raises on malformed input: unbalanced groups and unclosed
environments degrade to text rather than aborting a conversion.
"""

from __future__ import annotations

import datetime
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from ..preprocess import find_group, find_optional, find_overlay

# --------------------------------------------------------------------------
# node types
# --------------------------------------------------------------------------


@dataclass
class Node:
    pos: int = 0
    end: int = 0
    raw: str = ""


@dataclass
class TextNode(Node):
    text: str = ""


@dataclass
class MacroNode(Node):
    name: str = ""
    overlay: Optional[str] = None
    args: List[Optional[str]] = field(default_factory=list)   # in signature order
    spec: str = ""

    def arg(self, index: int) -> str:
        """Mandatory-argument accessor that tolerates missing arguments."""
        mandatory = [a for a, s in zip(self.args, self.spec) if s == "m"]
        return mandatory[index] if index < len(mandatory) else ""

    def opt(self, index: int = 0) -> Optional[str]:
        optional = [a for a, s in zip(self.args, self.spec) if s == "o"]
        return optional[index] if index < len(optional) else None


@dataclass
class GroupNode(Node):
    nodes: List[Node] = field(default_factory=list)


@dataclass
class MathNode(Node):
    latex: str = ""
    display: bool = False
    env: Optional[str] = None       # equation, align, ... when it came from one
    numbered: bool = False


@dataclass
class EnvNode(Node):
    name: str = ""
    overlay: Optional[str] = None
    args: List[Optional[str]] = field(default_factory=list)
    spec: str = ""
    body: str = ""                  # raw body source
    nodes: List[Node] = field(default_factory=list)

    def arg(self, index: int) -> str:
        mandatory = [a for a, s in zip(self.args, self.spec) if s == "m"]
        return mandatory[index] if index < len(mandatory) else ""

    def opt(self, index: int = 0) -> Optional[str]:
        optional = [a for a, s in zip(self.args, self.spec) if s == "o"]
        return optional[index] if index < len(optional) else None


# --------------------------------------------------------------------------
# signatures
# --------------------------------------------------------------------------

#: ``o`` = optional ``[...]``, ``m`` = mandatory ``{...}``.  Macros absent from
#: this table take no arguments; a group that follows one is scanned normally,
#: so ``\unknown{text}`` still renders "text".
MACRO_SIGNATURES: Dict[str, str] = {
    # sectioning and metadata
    "section": "om", "subsection": "om", "subsubsection": "om",
    "section*": "m", "subsection*": "m",
    "title": "om", "subtitle": "om", "author": "om", "date": "om",
    "institute": "om", "titlegraphic": "om", "logo": "m",
    "frametitle": "om", "framesubtitle": "om", "note": "om",
    # text styling
    "textbf": "m", "textit": "m", "textsl": "m", "textsc": "m", "textrm": "m",
    "textsf": "m", "texttt": "m", "textnormal": "m", "emph": "m",
    "underline": "m", "uline": "m", "sout": "m", "textsuperscript": "m",
    "textsubscript": "m", "mbox": "m", "makebox": "oom", "fbox": "m",
    "alert": "m", "structure": "m", "textcolor": "mm", "color": "m",
    "colorbox": "mm", "fcolorbox": "mmm", "boldmath": "", "st": "m",
    # links and references
    "href": "mm", "url": "m", "nolinkurl": "m", "hyperlink": "mm",
    "hypertarget": "mm", "label": "m", "ref": "m", "eqref": "m", "pageref": "m",
    "cite": "om", "citep": "om", "citet": "om", "citeauthor": "m",
    "citeyear": "m", "footnote": "om", "footnotemark": "o",
    # graphics
    "includegraphics": "om", "caption": "om", "subcaption": "om",
    "resizebox": "mmm", "scalebox": "om", "raisebox": "mm",
    "rotatebox": "om", "adjustbox": "mm",
    # tables
    "multicolumn": "mmm", "multirow": "ommm", "cmidrule": "om",
    "rowcolor": "om", "cellcolor": "om", "columncolor": "om",
    "specialrule": "mmm", "addlinespace": "o",
    # overlays
    "only": "m", "uncover": "m", "visible": "m", "invisible": "m",
    "onslide": "o", "alt": "mm", "temporal": "mmm", "item": "o",
    "action": "m",
    # spacing and layout
    "hspace": "m", "vspace": "m", "hspace*": "m", "vspace*": "m",
    "rule": "omm", "parbox": "omm", "framebox": "oom",
    "setlength": "mm", "addtolength": "mm", "renewcommand": "mm",
    "usetheme": "om", "usecolortheme": "om", "setbeamertemplate": "omm",
    "setbeamercolor": "mm", "setbeamerfont": "mm", "captionsetup": "om",
    "tableofcontents": "o", "titlepage": "o", "insertsection": "",
    "documentclass": "om", "usepackage": "om",
    "frame": "m", "againframe": "om",
    # internal placeholders from the preprocessor
    "beamerverbblock": "m", "beamerverbinline": "m",
}

#: Environments and the arguments that follow ``\begin{name}``.
ENV_SIGNATURES: Dict[str, str] = {
    "frame": "o", "columns": "o", "column": "m", "minipage": "oom",
    "block": "m", "alertblock": "m", "exampleblock": "m",
    "itemize": "o", "enumerate": "o", "description": "o",
    "tabular": "om", "tabular*": "mom", "tabularx": "mom", "tabulary": "mom",
    "array": "om", "longtable": "om", "supertabular": "m", "threeparttable": "o",
    "figure": "o", "figure*": "o", "table": "o", "table*": "o",
    "subfigure": "om", "subtable": "om", "wrapfigure": "omm",
    "tikzpicture": "o", "axis": "o", "semilogxaxis": "o", "semilogyaxis": "o",
    "loglogaxis": "o", "groupplot": "o", "scope": "o", "tikzcd": "o",
    "theorem": "o", "lemma": "o", "corollary": "o", "definition": "o",
    "proposition": "o", "example": "o", "proof": "o", "remark": "o",
    "beamercolorbox": "om", "overlayarea": "mm", "overprint": "o",
    "center": "", "flushleft": "", "flushright": "", "quote": "", "quotation": "",
}

#: Environments whose body is mathematics.
MATH_ENVS = {
    "equation", "equation*", "align", "align*", "alignat", "alignat*",
    "gather", "gather*", "multline", "multline*", "eqnarray", "eqnarray*",
    "displaymath", "math", "flalign", "flalign*", "IEEEeqnarray",
    "IEEEeqnarray*", "split", "dmath", "dmath*",
}

#: Environments drawn by TikZ/PGF, which we compile rather than parse.
TIKZ_ENVS = {
    "tikzpicture", "pgfpicture", "tikzcd", "circuitikz", "forest",
}

#: Environments whose body must be handed to the table parser verbatim.
TABLE_ENVS = {
    "tabular", "tabular*", "tabularx", "tabulary", "array", "longtable",
    "supertabular", "xtabular", "tabu",
}

#: Accent macros, as the Unicode combining mark they apply to the next
#: character or group.  ``\'e`` becomes "e" + U+0301, normalised to "é".
ACCENTS = {
    "'": "́", "`": "̀", "^": "̂", '"': "̈",
    "~": "̃", "=": "̄", ".": "̇", "H": "̋",
    "u": "̆", "v": "̌", "c": "̧", "k": "̨",
    "r": "̊",
}

#: Single characters LaTeX escapes with a backslash.
ESCAPES = {
    "%": "%", "$": "$", "&": "&", "#": "#", "_": "_", "{": "{", "}": "}",
    " ": " ", ",": "\u2009", ";": "\u2002", ":": "\u2005", "!": "",
    "-": "", "/": "", "@": "",
}

#: Accents and symbol macros that map straight onto a character.
SYMBOLS = {
    "ldots": "\u2026", "dots": "\u2026", "cdots": "\u22ef", "textellipsis": "\u2026",
    "textbackslash": "\\", "textasciitilde": "~", "textasciicircum": "^",
    "textbar": "|", "textless": "<", "textgreater": ">", "textquotedblleft": "\u201c",
    "textquotedblright": "\u201d", "textquoteleft": "\u2018", "textquoteright": "\u2019",
    "textemdash": "\u2014", "textendash": "\u2013", "texttrademark": "\u2122",
    "textregistered": "\u00ae", "copyright": "\u00a9", "textcopyright": "\u00a9",
    "degree": "\u00b0", "textdegree": "\u00b0", "pounds": "\u00a3", "euro": "\u20ac",
    "texteuro": "\u20ac", "S": "\u00a7", "P": "\u00b6", "dag": "\u2020",
    "ddag": "\u2021", "bullet": "\u2022", "LaTeX": "LaTeX", "TeX": "TeX",
    "BibTeX": "BibTeX", "quad": "\u2003", "qquad": "\u2003\u2003",
    "thinspace": "\u2009", "enspace": "\u2002", "nobreakspace": "\u00a0",
    "space": " ", "newline": "\n", "par": "\n\n", "ae": "\u00e6", "oe": "\u0153",
    "AA": "\u00c5", "aa": "\u00e5", "ss": "\u00df", "o": "\u00f8", "O": "\u00d8",
    "l": "\u0142", "L": "\u0141", "i": "i", "j": "j",
}

_MACRO_RE = re.compile(r"\\([A-Za-z@]+\*?|.)", re.S)


# --------------------------------------------------------------------------
# scanner
# --------------------------------------------------------------------------


def _read_args(text: str, pos: int, spec: str) -> Tuple[List[Optional[str]], int]:
    args: List[Optional[str]] = []
    i = pos
    for kind in spec:
        if kind == "o":
            value, i = find_optional(text, i)
            args.append(value)
        else:
            value, after = find_group(text, i)
            if after == i:           # no group present: signature did not match
                args.append(None)
            else:
                args.append(value)
                i = after
    return args, i


def scan(text: str, offset: int = 0) -> List[Node]:
    """Scan *text* into a flat list of nodes (groups and environments nest)."""
    nodes: List[Node] = []
    i = 0
    n = len(text)
    buf: List[str] = []
    buf_start = 0

    def flush(end: int) -> None:
        if buf:
            joined = "".join(buf)
            if joined:
                nodes.append(TextNode(pos=offset + buf_start, end=offset + end,
                                      raw=joined, text=joined))
            buf.clear()

    while i < n:
        ch = text[i]

        if ch == "\\":
            m = _MACRO_RE.match(text, i)
            if not m:
                buf.append(ch)
                i += 1
                continue
            name = m.group(1)

            # \\ line break
            if name == "\\":
                flush(i)
                _, after = find_optional(text, m.end())
                nodes.append(MacroNode(pos=offset + i, end=offset + after,
                                       raw=text[i:after], name="\\"))
                i = after
                buf_start = i
                continue

            # escaped character
            if len(name) == 1 and name in ESCAPES:
                buf.append(ESCAPES[name])
                i = m.end()
                continue

            # \[ ... \] display math
            if name == "[":
                close = _find_close(text, m.end(), "\\]")
                flush(i)
                body = text[m.end():close]
                nodes.append(MathNode(pos=offset + i, end=offset + close + 2,
                                      raw=text[i:close + 2], latex=body, display=True))
                i = close + 2
                buf_start = i
                continue

            # \( ... \) inline math
            if name == "(":
                close = _find_close(text, m.end(), "\\)")
                flush(i)
                body = text[m.end():close]
                nodes.append(MathNode(pos=offset + i, end=offset + close + 2,
                                      raw=text[i:close + 2], latex=body, display=False))
                i = close + 2
                buf_start = i
                continue

            if name == "begin":
                env_name, after = find_group(text, m.end())
                if after != m.end():
                    flush(i)
                    node, i = _scan_environment(text, i, after, env_name.strip(), offset)
                    nodes.append(node)
                    buf_start = i
                    continue

            if name == "end":
                # A stray \end: skip its argument and carry on.
                _, after = find_group(text, m.end())
                i = after if after != m.end() else m.end()
                continue

            if name in ACCENTS:
                base, after = _read_accent_base(text, m.end())
                buf.append(unicodedata.normalize("NFC", base + ACCENTS[name]))
                i = after
                continue

            if name == "today":
                buf.append(datetime.date.today().strftime("%B %d, %Y"))
                i = m.end()
                continue

            if name in SYMBOLS:
                buf.append(SYMBOLS[name])
                i = m.end()
                continue

            flush(i)
            overlay, after = find_overlay(text, m.end())
            spec = MACRO_SIGNATURES.get(name, "")
            args, after = _read_args(text, after, spec)
            nodes.append(MacroNode(pos=offset + i, end=offset + after,
                                   raw=text[i:after], name=name,
                                   overlay=overlay, args=args, spec=spec))
            i = after
            buf_start = i
            continue

        if ch == "{":
            body, after = find_group(text, i)
            flush(i)
            nodes.append(GroupNode(pos=offset + i, end=offset + after,
                                   raw=text[i:after],
                                   nodes=scan(body, offset + i + 1)))
            i = after
            buf_start = i
            continue

        if ch == "}":
            # Unbalanced closing brace: ignore it.
            i += 1
            continue

        if ch == "$":
            display = text.startswith("$$", i)
            delim = "$$" if display else "$"
            close = text.find(delim, i + len(delim))
            if close < 0:
                buf.append(ch)
                i += 1
                continue
            flush(i)
            body = text[i + len(delim):close]
            nodes.append(MathNode(pos=offset + i, end=offset + close + len(delim),
                                  raw=text[i:close + len(delim)],
                                  latex=body, display=display))
            i = close + len(delim)
            buf_start = i
            continue

        if ch == "~":
            buf.append("\u00a0")
            i += 1
            continue

        if not buf:
            buf_start = i
        buf.append(ch)
        i += 1

    flush(n)
    return nodes


def _read_accent_base(text: str, pos: int) -> Tuple[str, int]:
    r"""Read the letter an accent applies to: ``\'e``, ``\'{e}`` or ``\c{c}``."""
    i = pos
    while i < len(text) and text[i] in " \t":
        i += 1
    if i < len(text) and text[i] == "{":
        body, after = find_group(text, i)
        # \'{\i} and \'{\j} are dotless letters carrying an accent.
        body = body.strip().replace("\\i", "i").replace("\\j", "j")
        return body[:1], after
    if i < len(text):
        if text[i] == "\\":
            m = re.match(r"\\(i|j)\b", text[i:])
            if m:
                return m.group(1), i + m.end()
        return text[i], i + 1
    return "", pos


def _find_close(text: str, start: int, token: str) -> int:
    idx = text.find(token, start)
    return idx if idx >= 0 else len(text)


def _scan_environment(text: str, start: int, after_name: int, name: str,
                      offset: int) -> Tuple[Node, int]:
    """Scan ``\\begin{name} ... \\end{name}`` starting at *start*."""
    end_re = re.compile(r"\\end\s*\{" + re.escape(name) + r"\}")

    # Find the matching \end, skipping nested environments of the same name.
    depth = 1
    search = after_name
    begin_re = re.compile(r"\\begin\s*\{" + re.escape(name) + r"\}")
    body_end = len(text)
    end_after = len(text)
    while search < len(text):
        m_begin = begin_re.search(text, search)
        m_end = end_re.search(text, search)
        if not m_end:
            break
        if m_begin and m_begin.start() < m_end.start():
            depth += 1
            search = m_begin.end()
            continue
        depth -= 1
        if depth == 0:
            body_end = m_end.start()
            end_after = m_end.end()
            break
        search = m_end.end()

    overlay, i = find_overlay(text, after_name)
    spec = ENV_SIGNATURES.get(name, "")
    args, i = _read_args(text, i, spec)
    body = text[i:body_end]

    if name in MATH_ENVS:
        node: Node = MathNode(pos=offset + start, end=offset + end_after,
                              raw=text[start:end_after], latex=body, display=True,
                              env=name, numbered=not name.endswith("*"))
        return node, end_after

    children: List[Node] = []
    if name not in TIKZ_ENVS and name not in TABLE_ENVS:
        children = scan(body, offset + i)

    return EnvNode(pos=offset + start, end=offset + end_after,
                   raw=text[start:end_after], name=name, overlay=overlay,
                   args=args, spec=spec, body=body, nodes=children), end_after
