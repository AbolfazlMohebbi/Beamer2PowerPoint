r"""LaTeX mathematics into native PowerPoint equations.

The path is LaTeX -> MathML (``latex2mathml``) -> OMML, using the
``MML2OMML.XSL`` stylesheet that ships with Microsoft Office.  The result is a
real, editable PowerPoint equation rather than a picture.  When a formula
defeats the converter the caller can fall back to a rendered image.
"""

from __future__ import annotations

import logging
import os
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import List, Optional

log = logging.getLogger(__name__)

M_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"

#: Standard install locations of the Office MathML-to-OMML stylesheet.
_XSL_CANDIDATES = (
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office15\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office15\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office14\MML2OMML.XSL",
)

#: LaTeX that latex2mathml does not know but that has a plain equivalent.
_PREPROCESS = (
    (re.compile(r"\\(?:displaystyle|textstyle|scriptstyle|scriptscriptstyle)\b"), ""),
    (re.compile(r"\\(?:limits|nolimits)\b"), ""),
    (re.compile(r"\\(?:,|;|:|!)"), " "),
    (re.compile(r"\\(?:bigl|bigr|Bigl|Bigr|biggl|biggr|Biggl|Biggr|big|Big|bigg|Bigg)\b"), ""),
    (re.compile(r"\\(?:left|right)\s*\."), ""),
    (re.compile(r"\\(?:qquad|quad)\b"), r"\\ "),
    (re.compile(r"\\(?:notag|nonumber)\b"), ""),
    (re.compile(r"\\label\s*\{[^}]*\}"), ""),
    (re.compile(r"\\(?:vspace|hspace)\*?\s*\{[^}]*\}"), ""),
    (re.compile(r"\\intertext\s*\{[^}]*\}"), ""),
    (re.compile(r"&\s*=\s*&"), "="),
)


def find_mml2omml() -> Optional[str]:
    """Locate ``MML2OMML.XSL`` from an Office installation."""
    for path in _XSL_CANDIDATES:
        if os.path.isfile(path):
            return path
    for root in (r"C:\Program Files\Microsoft Office",
                 r"C:\Program Files (x86)\Microsoft Office"):
        if not os.path.isdir(root):
            continue
        for dirpath, _dirnames, filenames in os.walk(root):
            for name in filenames:
                if name.upper() == "MML2OMML.XSL":
                    return os.path.join(dirpath, name)
    return None


_BEGIN_END_RE = re.compile(r"\\begin\s*\{([^}]*)\}|\\end\s*\{([^}]*)\}")


def normalise_latex(latex: str) -> str:
    """Make a formula more digestible for ``latex2mathml``."""
    text = latex.strip()
    for pattern, replacement in _PREPROCESS:
        text = pattern.sub(replacement, text)
    text = _promote_norm_delimiters(text)
    return _strip_alignment_ampersands(text).strip()


_BARE_NORM_RE = re.compile(r"(?<!\\)\\\|")


def _promote_norm_delimiters(text: str) -> str:
    r"""Turn ``\|x\|`` into ``\left\|x\right\|``.

    Office's MathML-to-OMML stylesheet drops the entire contents of a norm
    that carries a sub- or superscript — ``\|x\|_2^2`` converts to an empty
    delimiter — because latex2mathml emits the bars as loose operators rather
    than a fenced group.  The ``\left``/``\right`` form produces the nested
    structure the stylesheet expects, and renders identically.
    """
    text = (text.replace(r"\lVert", r"\left\|").replace(r"\rVert", r"\right\|")
                .replace(r"\lvert", r"\left|").replace(r"\rvert", r"\right|"))

    # Leave alone any bar that is already qualified.
    masked = text.replace(r"\left\|", "\x01").replace(r"\right\|", "\x02")
    positions = [m.start() for m in _BARE_NORM_RE.finditer(masked)]
    if positions and len(positions) % 2 == 0:
        pieces = []
        previous = 0
        for index, start in enumerate(positions):
            pieces.append(masked[previous:start])
            pieces.append(r"\left\|" if index % 2 == 0 else r"\right\|")
            previous = start + 2
        pieces.append(masked[previous:])
        masked = "".join(pieces)

    return masked.replace("\x01", r"\left\|").replace("\x02", r"\right\|")


def _strip_alignment_ampersands(text: str) -> str:
    r"""Drop alignment ``&`` but keep the ones that separate matrix columns.

    Outside an environment an ``&`` is an ``align`` tab stop with no MathML
    equivalent; inside ``bmatrix``, ``array``, ``cases`` and friends it is a
    real column separator that latex2mathml needs.
    """
    out = []
    depth = 0
    index = 0
    for match in _BEGIN_END_RE.finditer(text):
        segment = text[index:match.start()]
        out.append(segment if depth else re.sub(r"(?<!\\)&", " ", segment))
        out.append(match.group(0))
        depth += 1 if match.group(1) else -1
        depth = max(0, depth)
        index = match.end()
    tail = text[index:]
    out.append(tail if depth else re.sub(r"(?<!\\)&", " ", tail))
    return "".join(out)


@dataclass
class MathConverter:
    """Converts LaTeX formulas to OMML elements."""

    xsl_path: Optional[str] = None
    enabled: bool = True
    #: Formulas the converter could not handle, for the run report.
    failures: List[str] = field(default_factory=list)
    _transform = None
    _checked: bool = field(default=False, init=False)

    def __post_init__(self) -> None:
        if self.xsl_path is None:
            self.xsl_path = find_mml2omml()

    @property
    def available(self) -> bool:
        if not self.enabled:
            return False
        self._ensure_transform()
        return self._transform is not None

    def _ensure_transform(self) -> None:
        if self._checked:
            return
        self._checked = True
        if not self.xsl_path or not os.path.isfile(self.xsl_path):
            log.warning("MML2OMML.XSL not found; equations will be rendered as images.")
            return
        try:
            from lxml import etree

            self._transform = etree.XSLT(etree.parse(self.xsl_path))
        except Exception as exc:
            log.warning("Could not load %s: %s", self.xsl_path, exc)
            self._transform = None

    def to_omml(self, latex: str):
        """Convert one formula to an ``m:oMath`` element, or ``None``."""
        if not self.available:
            return None

        from lxml import etree

        try:
            import latex2mathml.converter as converter
        except ImportError:
            log.warning("latex2mathml is not installed; equations become images.")
            return None

        source = normalise_latex(latex)
        if not source:
            return None
        try:
            mathml = converter.convert(source)
        except Exception as exc:
            self.failures.append("%s (%s)" % (_short(latex), exc.__class__.__name__))
            return None

        try:
            tree = etree.fromstring(mathml.encode("utf-8"))
            result = self._transform(tree)
            root = result.getroot()
        except Exception as exc:
            self.failures.append("%s (%s)" % (_short(latex), exc.__class__.__name__))
            return None

        if root is None or not root.tag.endswith("}oMath"):
            self.failures.append(_short(latex))
            return None
        if not len(root) and not (root.text or "").strip():
            # An empty conversion is worse than no conversion.
            self.failures.append(_short(latex))
            return None

        # The stylesheet can silently drop a sub-expression rather than fail.
        # A formula that is quietly missing a term is far worse than one
        # rendered as a picture, so verify nothing was lost before accepting.
        missing = _lost_content(tree, root)
        if missing:
            self.failures.append("%s [%s]" % (_short(latex), missing))
            return None

        _flatten_operator_separators(root)
        _isolate_astral_runs(root)
        return root


#: Structural slots that carry an expression and must never come back empty.
#: Deliberately excludes ``deg`` (a plain square root has no degree) and
#: ``sub``/``sup`` (an n-ary operator may carry only one limit).
_REQUIRED_SLOTS = ("e", "num", "den", "fName")


def _alnum(text: str) -> str:
    """The characters whose loss would actually change the formula.

    Punctuation is excluded on purpose: OMML stores fences, accents and
    n-ary symbols as attributes rather than text, and omits them entirely
    when they are the default for the construct, so comparing them would
    flag perfectly good conversions.
    """
    return "".join(c for c in (text or "") if c.isalnum())


def _lost_content(mathml_root, omml_root) -> Optional[str]:
    """Describe what the stylesheet dropped, or ``None`` if nothing was."""
    for tag in _REQUIRED_SLOTS:
        for element in omml_root.iter("{%s}%s" % (M_NAMESPACE, tag)):
            if len(element) == 0 and not (element.text or "").strip():
                return "empty <m:%s>" % tag

    source = _alnum("".join(mathml_root.itertext()))
    produced = _alnum("".join(
        "".join(e.itertext()) for e in omml_root.iter("{%s}t" % M_NAMESPACE)))

    missing = Counter(source) - Counter(produced)
    if missing:
        return "dropped %r" % "".join(sorted(missing.elements()))[:12]
    return None


#: Characters that really do separate the slots of a delimiter, as in
#: ``(a, b)``.  Anything else the stylesheet puts in ``m:sepChr`` is an
#: operator it mis-filed, not a separator.
_TRUE_SEPARATORS = {",", ";", "|", "‖", ""}


def _flatten_operator_separators(root) -> None:
    r"""Rebuild delimiters whose "separator" is really an operator.

    For ``\left\|p - \Phi w\right\|`` the stylesheet emits a delimiter with
    two slots and ``m:sepChr`` set to the minus sign.  PowerPoint draws that
    separator as an arrow rather than a minus, and the structure is wrong in
    any case: the minus belongs inside the norm, not between two arguments.
    Merging the slots back into one and inserting the operator as an ordinary
    run gives what the author wrote.
    """
    d_tag = "{%s}d" % M_NAMESPACE
    for delimiter in list(root.iter(d_tag)):
        properties = delimiter.find("{%s}dPr" % M_NAMESPACE)
        if properties is None:
            continue
        separator = properties.find("{%s}sepChr" % M_NAMESPACE)
        if separator is None:
            continue
        value = separator.get("{%s}val" % M_NAMESPACE) or ""
        if value in _TRUE_SEPARATORS:
            continue

        slots = delimiter.findall("{%s}e" % M_NAMESPACE)
        if len(slots) < 2:
            continue

        first = slots[0]
        for slot in slots[1:]:
            run = first.makeelement("{%s}r" % M_NAMESPACE, {})
            text = first.makeelement("{%s}t" % M_NAMESPACE, {})
            text.text = value
            run.append(text)
            first.append(run)
            for child in list(slot):
                first.append(child)
            delimiter.remove(slot)
        separator.set("{%s}val" % M_NAMESPACE, "")


def split_on_astral(text: str) -> List[str]:
    """Split *text* so each non-BMP character stands alone."""
    pieces: List[str] = []
    buffer = ""
    for char in text:
        if ord(char) > 0xFFFF:
            if buffer:
                pieces.append(buffer)
                buffer = ""
            pieces.append(char)
        else:
            buffer += char
    if buffer:
        pieces.append(buffer)
    return pieces


def _isolate_astral_runs(root) -> None:
    r"""Give every non-BMP character in the expression its own run.

    PowerPoint renders only the *first* non-BMP character of a maths run and
    silently drops any that follow, so ``\mathbf{r} = \mathbf{p} - \Phi``
    arrives as one run "𝐫=𝐩−Φ" and displays as "𝐫=−Φ".  Bold, italic,
    script and fraktur letters all live above U+FFFF, so this hits ordinary
    vector notation constantly.

    Splitting the run changes nothing but structure — the characters and
    their order are identical, and it is how PowerPoint stores such runs
    itself.
    """
    from copy import deepcopy

    run_tag = "{%s}r" % M_NAMESPACE
    text_tag = "{%s}t" % M_NAMESPACE

    for run in list(root.iter(run_tag)):
        text_element = run.find(text_tag)
        if text_element is None or not text_element.text:
            continue
        text = text_element.text
        if len(text) < 2 or all(ord(c) <= 0xFFFF for c in text):
            continue

        pieces = split_on_astral(text)
        if len(pieces) < 2:
            continue

        parent = run.getparent()
        if parent is None:
            continue
        position = list(parent).index(run)
        for offset, piece in enumerate(pieces):
            clone = deepcopy(run)
            clone.find(text_tag).text = piece
            parent.insert(position + offset, clone)
        parent.remove(run)


def _short(latex: str, limit: int = 60) -> str:
    text = " ".join(latex.split())
    return text if len(text) <= limit else text[:limit] + "..."
