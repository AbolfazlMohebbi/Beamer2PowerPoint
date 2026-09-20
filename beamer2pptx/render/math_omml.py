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
    return _strip_alignment_ampersands(text).strip()


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
        return root


def _short(latex: str, limit: int = 60) -> str:
    text = " ".join(latex.split())
    return text if len(text) <= limit else text[:limit] + "..."
