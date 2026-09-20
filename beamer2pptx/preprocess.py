r"""Flattening a LaTeX project into a single expanded source string.

Runs before any structural parsing and handles the text-level concerns:
protecting verbatim regions, dropping comments, inlining ``\input``-ed files,
harvesting ``\newcommand`` definitions and ``\graphicspath``, and splitting the
preamble from the document body.
"""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .project import read_text

log = logging.getLogger(__name__)

VERBATIM_ENVS = (
    "verbatim", "Verbatim", "BVerbatim", "LVerbatim", "semiverbatim",
    "lstlisting", "minted", "alltt", "listing", "pyglist",
)

# Placeholder macro names must be letters only: that is all a LaTeX control
# word may contain, and the scanner tokenises them the same way.
VERB_PLACEHOLDER = "\\beamerverbblock"
VERBINLINE_PLACEHOLDER = "\\beamerverbinline"

_INPUT_RE = re.compile(r"\\(input|include|subfile|subfileinclude)\s*\{([^}]*)\}")
_GRAPHICSPATH_RE = re.compile(r"\\graphicspath\s*\{((?:\s*\{[^}]*\}\s*)+)\}")
_NEWCOMMAND_RE = re.compile(
    r"\\(newcommand|renewcommand|providecommand)\s*\*?\s*"
    r"(?:\{\s*\\([A-Za-z@]+)\s*\}|\\([A-Za-z@]+))"
    r"\s*(?:\[\s*(\d+)\s*\])?"
    r"\s*(?:\[([^\]]*)\])?"
    r"\s*(?=\{)"
)
_DEF_RE = re.compile(r"\\def\s*\\([A-Za-z@]+)\s*(?=\{)")
_DECLARE_OP_RE = re.compile(r"\\DeclareMathOperator\s*\*?\s*\{\s*\\([A-Za-z@]+)\s*\}\s*(?=\{)")


@dataclass
class MacroDef:
    name: str
    nargs: int = 0
    default: Optional[str] = None
    body: str = ""


@dataclass
class Source:
    """The flattened document."""

    text: str = ""
    preamble: str = ""
    body: str = ""
    root: str = "."
    graphicspaths: List[str] = field(default_factory=list)
    macros: Dict[str, MacroDef] = field(default_factory=dict)
    verbatims: List[Tuple[str, str]] = field(default_factory=list)
    files: List[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# balanced-group helpers (used by every parsing stage)
# --------------------------------------------------------------------------


def find_group(text: str, start: int) -> Tuple[str, int]:
    """Read a balanced ``{...}`` group at *start*.

    Returns the contents and the index past the closing brace; ``("", start)``
    when there is no group there.
    """
    i = start
    while i < len(text) and text[i].isspace():
        i += 1
    if i >= len(text) or text[i] != "{":
        return "", start
    depth = 0
    j = i
    while j < len(text):
        ch = text[j]
        if ch == "\\":
            j += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[i + 1:j], j + 1
        j += 1
    return text[i + 1:], len(text)


def find_optional(text: str, start: int) -> Tuple[Optional[str], int]:
    """Read a ``[...]`` optional argument at *start*."""
    j = start
    while j < len(text) and text[j] in " \t":
        j += 1
    if j >= len(text) or text[j] != "[":
        return None, start
    depth = 0
    i = j
    while i < len(text):
        ch = text[i]
        if ch == "\\":
            i += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return text[j + 1:i], i + 1
        i += 1
    return None, start


def find_overlay(text: str, start: int) -> Tuple[Optional[str], int]:
    """Read a Beamer ``<...>`` overlay specification at *start*."""
    j = start
    while j < len(text) and text[j] in " \t":
        j += 1
    if j >= len(text) or text[j] != "<":
        return None, start
    close = text.find(">", j)
    if close < 0:
        return None, start
    return text[j + 1:close], close + 1


# --------------------------------------------------------------------------
# verbatim protection
# --------------------------------------------------------------------------


def protect_verbatim(text: str, store: List[Tuple[str, str]]) -> str:
    r"""Replace verbatim environments and ``\verb`` with opaque placeholders."""
    out: List[str] = []
    i = 0
    env_alt = "|".join(re.escape(e) for e in VERBATIM_ENVS)
    begin_re = re.compile(r"\\begin\s*\{(" + env_alt + r")\*?\}(\[[^\]]*\])?(\{[^}]*\})?")
    verb_re = re.compile(r"\\(?:verb|lstinline|mintinline)\*?(?:\[[^\]]*\])?(?:\{[A-Za-z]+\})?(.)")

    while i < len(text):
        m_env = begin_re.search(text, i)
        m_verb = verb_re.search(text, i)
        if m_env and (not m_verb or m_env.start() <= m_verb.start()):
            name = m_env.group(1)
            m_end = re.compile(r"\\end\s*\{" + re.escape(name) + r"\*?\}").search(text, m_env.end())
            if not m_end:
                out.append(text[i:])
                return "".join(out)
            body = text[m_env.end():m_end.start()].strip("\n")
            store.append((name, body))
            out.append(text[i:m_env.start()])
            out.append("%s{%d}" % (VERB_PLACEHOLDER, len(store) - 1))
            i = m_end.end()
        elif m_verb:
            delim = m_verb.group(1)
            if delim in "*\n":
                out.append(text[i:m_verb.end()])
                i = m_verb.end()
                continue
            close = text.find(delim, m_verb.end())
            if close < 0:
                out.append(text[i:])
                return "".join(out)
            store.append(("verb", text[m_verb.end():close]))
            out.append(text[i:m_verb.start()])
            out.append("%s{%d}" % (VERBINLINE_PLACEHOLDER, len(store) - 1))
            i = close + 1
        else:
            out.append(text[i:])
            break
    return "".join(out)


# --------------------------------------------------------------------------
# comments
# --------------------------------------------------------------------------


def strip_comments(text: str) -> str:
    r"""Remove ``%`` comments, mimicking the way TeX swallows the newline."""
    out: List[str] = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and i + 1 < n:
            out.append(text[i:i + 2])
            i += 2
            continue
        if ch == "%":
            j = text.find("\n", i)
            if j < 0:
                break
            k = j + 1
            while k < n and text[k] in " \t":
                k += 1
            i = k
            continue
        out.append(ch)
        i += 1
    text = "".join(out)
    text = re.sub(r"\\iffalse\b.*?\\fi\b", "", text, flags=re.S)
    text = re.sub(r"\\begin\s*\{comment\}.*?\\end\s*\{comment\}", "", text, flags=re.S)
    return text


# --------------------------------------------------------------------------
# \input inlining
# --------------------------------------------------------------------------


def _resolve_input(name: str, base: str, root: str) -> Optional[str]:
    name = name.strip().replace("\\", "/")
    names = [name]
    if not name.lower().endswith((".tex", ".ltx")):
        names.append(name + ".tex")
    for b in (base, root):
        for candidate in names:
            path = os.path.join(b, candidate)
            if os.path.isfile(path):
                return os.path.abspath(path)
    return None


def inline_inputs(text: str, base: str, root: str, seen: set, files: List[str],
                  store: List[Tuple[str, str]], depth: int = 0) -> str:
    r"""Recursively splice ``\input``/``\include``/``\subfile`` files in place."""
    if depth > 24:
        log.warning("Stopping \\input recursion at depth %d", depth)
        return text

    def repl(match) -> str:
        target = _resolve_input(match.group(2), base, root)
        if target is None:
            log.warning("Could not resolve \\%s{%s}", match.group(1), match.group(2))
            return ""
        if target in seen:
            log.warning("Skipping circular \\input of %s", target)
            return ""
        seen.add(target)
        files.append(target)
        child = strip_comments(protect_verbatim(read_text(target), store))
        if match.group(1).startswith("subfile"):
            m = re.search(r"\\begin\s*\{document\}(.*)\\end\s*\{document\}", child, re.S)
            if m:
                child = m.group(1)
        return inline_inputs(child, os.path.dirname(target), root, seen, files, store, depth + 1)

    return _INPUT_RE.sub(repl, text)


# --------------------------------------------------------------------------
# macro harvesting and expansion
# --------------------------------------------------------------------------


def harvest_macros(text: str) -> Dict[str, MacroDef]:
    r"""Collect ``\newcommand``/``\def``/``\DeclareMathOperator`` definitions."""
    macros: Dict[str, MacroDef] = {}

    for m in _NEWCOMMAND_RE.finditer(text):
        name = m.group(2) or m.group(3)
        body, _ = find_group(text, m.end())
        macros[name] = MacroDef(name, int(m.group(4) or 0), m.group(5), body)

    for m in _DEF_RE.finditer(text):
        if m.group(1) in macros:
            continue
        body, _ = find_group(text, m.end())
        macros[m.group(1)] = MacroDef(m.group(1), 0, None, body)

    for m in _DECLARE_OP_RE.finditer(text):
        body, _ = find_group(text, m.end())
        macros[m.group(1)] = MacroDef(m.group(1), 0, None, "\\operatorname{%s}" % body)

    return macros


#: Never expanded even if redefined: the parser gives these structural meaning.
PROTECTED = {
    "item", "frametitle", "framesubtitle", "section", "subsection", "subsubsection",
    "title", "subtitle", "author", "date", "institute", "titlepage", "maketitle",
    "includegraphics", "caption", "label", "pause", "only", "onslide", "uncover",
    "visible", "invisible", "alt", "column", "textbf", "textit", "emph", "texttt",
    "alert", "beamerverbblock", "beamerverbinline", "tableofcontents", "note", "multicolumn",
    "multirow", "begin", "end", "hline", "toprule", "midrule", "bottomrule",
    "textwidth", "linewidth", "textheight", "columnwidth",
}


def _split_args(text: str, pos: int, nargs: int,
                default: Optional[str]) -> Tuple[List[str], int]:
    args: List[str] = []
    i = pos
    remaining = nargs
    if default is not None and remaining > 0:
        opt, i = find_optional(text, i)
        args.append(opt if opt is not None else default)
        remaining -= 1
    for _ in range(remaining):
        while i < len(text) and text[i] in " \t\n":
            i += 1
        if i >= len(text):
            args.append("")
        elif text[i] == "{":
            arg, i = find_group(text, i)
            args.append(arg)
        elif text[i] == "\\":
            m = re.match(r"\\[A-Za-z@]+|\\.", text[i:])
            args.append(m.group(0) if m else "")
            i += m.end() if m else 1
        else:
            args.append(text[i])
            i += 1
    return args, i


def expand_macros(text: str, macros: Dict[str, MacroDef], max_passes: int = 8) -> str:
    """Textually expand user macros so later stages only see standard LaTeX."""
    active = {k: v for k, v in macros.items() if k not in PROTECTED}
    if not active:
        return text

    name_re = re.compile(r"\\([A-Za-z@]+)")
    for _ in range(max_passes):
        out: List[str] = []
        i = 0
        changed = False
        while True:
            m = name_re.search(text, i)
            if not m:
                out.append(text[i:])
                break
            macro = active.get(m.group(1))
            if macro is None:
                out.append(text[i:m.end()])
                i = m.end()
                continue
            args, end = _split_args(text, m.end(), macro.nargs, macro.default)
            body = macro.body
            for idx in range(len(args), 0, -1):
                body = body.replace("#%d" % idx, args[idx - 1])
            out.append(text[i:m.start()])
            out.append(body)
            i = end
            changed = True
        text = "".join(out)
        if not changed:
            break
    return text


# --------------------------------------------------------------------------
# entry point
# --------------------------------------------------------------------------


def preprocess(main_tex: str, root: str, expand: bool = True) -> Source:
    """Flatten *main_tex* into a single expanded source ready for parsing."""
    store: List[Tuple[str, str]] = []
    files = [os.path.abspath(main_tex)]

    text = strip_comments(protect_verbatim(read_text(main_tex), store))
    text = inline_inputs(text, os.path.dirname(main_tex) or ".", root,
                         {os.path.abspath(main_tex)}, files, store)

    graphicspaths: List[str] = []
    for m in _GRAPHICSPATH_RE.finditer(text):
        for g in re.findall(r"\{([^{}]*)\}", m.group(1)):
            if g.strip():
                graphicspaths.append(g.strip())

    macros = harvest_macros(text)

    m = re.search(r"\\begin\s*\{document\}", text)
    if m:
        preamble = text[:m.start()]
        rest = text[m.end():]
        m_end = re.search(r"\\end\s*\{document\}", rest)
        body = rest[:m_end.start()] if m_end else rest
    else:
        log.warning("No \\begin{document} found; treating the whole file as the body.")
        preamble, body = "", text

    if expand and macros:
        body = expand_macros(body, macros)

    return Source(text=text, preamble=preamble, body=body, root=root,
                  graphicspaths=graphicspaths, macros=macros,
                  verbatims=store, files=files)
