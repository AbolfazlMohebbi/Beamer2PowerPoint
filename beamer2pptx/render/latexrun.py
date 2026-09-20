r"""Compiling LaTeX snippets and turning PDFs into images.

Everything that needs a real TeX run — TikZ pictures, pgfplots axes, equations
that the OMML converter cannot handle — goes through here.  Results are cached
on a hash of the source, so re-converting a deck costs nothing.
"""

from __future__ import annotations

import hashlib
import logging
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

log = logging.getLogger(__name__)

#: Where MiKTeX and TeX Live usually put their binaries on Windows.
_SEARCH_DIRS = (
    r"C:\Program Files\MiKTeX\miktex\bin\x64",
    r"C:\Program Files (x86)\MiKTeX\miktex\bin",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\MiKTeX\miktex\bin\x64"),
    r"C:\texlive\2025\bin\windows",
    r"C:\texlive\2024\bin\windows",
    r"C:\texlive\2023\bin\win32",
    "/usr/bin", "/usr/local/bin", "/Library/TeX/texbin",
)

#: Packages a standalone snippet needs that the harvested preamble may not load.
_TIKZ_PACKAGES = ("tikz", "pgfplots")

#: Preamble lines that must not be carried into a standalone document.
_PREAMBLE_DROP = re.compile(
    r"^\s*\\(documentclass|usetheme|usecolortheme|usefonttheme|useinnertheme"
    r"|useoutertheme|setbeamer\w+|title|subtitle|author|date|institute"
    r"|titlegraphic|logo|AtBegin\w+|input|include|bibliography\w*|addbibresource"
    r"|graphicspath|hypersetup|setbeamercovered|beamertemplate\w*)\b"
)


def find_executable(name: str) -> Optional[str]:
    """Locate a TeX binary on PATH or in a standard installation directory."""
    found = shutil.which(name)
    if found:
        return found
    for directory in _SEARCH_DIRS:
        for candidate in (name, name + ".exe"):
            path = os.path.join(directory, candidate)
            if os.path.isfile(path):
                return path
    return None


@dataclass
class LatexRunner:
    """A cached ``pdflatex`` driver."""

    engine: str = "pdflatex"
    cache_dir: Optional[str] = None
    dpi: int = 300
    enabled: bool = True
    timeout: int = 180
    #: Populated with a human-readable reason when LaTeX is unavailable.
    unavailable_reason: Optional[str] = None
    _executable: Optional[str] = field(default=None, init=False, repr=False)
    _resolved: bool = field(default=False, init=False, repr=False)
    failures: List[str] = field(default_factory=list)

    # -- availability -----------------------------------------------------

    @property
    def executable(self) -> Optional[str]:
        if not self._resolved:
            self._resolved = True
            self._executable = find_executable(self.engine)
            if self._executable is None and self.engine != "pdflatex":
                self._executable = find_executable("pdflatex")
            if self._executable is None:
                self.unavailable_reason = (
                    "No LaTeX engine found. Install MiKTeX or TeX Live, or pass "
                    "--latex with the path to pdflatex."
                )
                log.warning("%s", self.unavailable_reason)
        return self._executable

    @property
    def available(self) -> bool:
        return self.enabled and self.executable is not None

    # -- compilation ------------------------------------------------------

    def compile_to_pdf(self, source: str, key: str = "") -> Optional[str]:
        """Compile a complete LaTeX document, returning the PDF path."""
        if not self.available:
            return None

        digest = hashlib.sha256((key + "\n" + source).encode("utf-8")).hexdigest()[:20]
        cached = self._cached(digest + ".pdf")
        if cached and os.path.isfile(cached):
            log.debug("LaTeX cache hit %s", digest)
            return cached

        work = tempfile.mkdtemp(prefix="b2p-tex-")
        try:
            tex_path = os.path.join(work, "snippet.tex")
            with open(tex_path, "w", encoding="utf-8") as fh:
                fh.write(source)

            command = [
                self.executable,
                "-interaction=nonstopmode",
                "-halt-on-error",
                "-file-line-error",
                "snippet.tex",
            ]
            # MiKTeX installs missing packages on demand when asked to.
            if "miktex" in (self.executable or "").lower():
                command.insert(1, "--enable-installer")

            try:
                result = subprocess.run(command, cwd=work, capture_output=True,
                                        text=True, timeout=self.timeout,
                                        errors="replace")
            except subprocess.TimeoutExpired:
                self.failures.append("LaTeX timed out after %ds" % self.timeout)
                return None

            pdf_path = os.path.join(work, "snippet.pdf")
            if not os.path.isfile(pdf_path):
                self.failures.append(_first_error(result.stdout or "")
                                     or "pdflatex produced no output")
                log.debug("LaTeX failed:\n%s", (result.stdout or "")[-2000:])
                return None

            if cached:
                os.makedirs(os.path.dirname(cached), exist_ok=True)
                shutil.copy(pdf_path, cached)
                return cached

            keep = os.path.join(tempfile.gettempdir(), "b2p-%s.pdf" % digest)
            shutil.copy(pdf_path, keep)
            return keep
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def _cached(self, name: str) -> Optional[str]:
        if not self.cache_dir:
            return None
        return os.path.join(self.cache_dir, name)

    # -- snippet documents ------------------------------------------------

    def standalone_document(self, body: str, preamble: str = "",
                            border: str = "2pt", packages: Sequence[str] = (),
                            class_options: str = "") -> str:
        """Wrap *body* in a ``standalone`` document using the deck preamble."""
        options = "border=%s" % border
        if class_options:
            options += "," + class_options
        cleaned = clean_preamble(preamble)
        missing = [p for p in packages if ("\\usepackage" not in cleaned
                                           or p not in cleaned)]
        extra = "".join("\\usepackage{%s}\n" % p for p in missing)
        pgfcompat = "\\pgfplotsset{compat=newest}\n" if "pgfplots" in (cleaned + extra) else ""
        return (
            "\\documentclass[%s]{standalone}\n" % options
            + "\\usepackage[utf8]{inputenc}\n"
            + "\\usepackage{amsmath,amssymb,amsfonts}\n"
            + extra
            + cleaned
            + pgfcompat
            + "\\begin{document}\n"
            + body
            + "\n\\end{document}\n"
        )

    def render_tikz(self, snippet: str, preamble: str = "") -> Optional[str]:
        """Compile a TikZ/pgfplots snippet, returning a PDF path."""
        source = self.standalone_document(snippet, preamble,
                                          packages=_TIKZ_PACKAGES)
        pdf = self.compile_to_pdf(source, key="tikz")
        if pdf:
            return pdf
        # Retry without the document preamble: a Beamer-only macro in there is
        # the most common reason a standalone build fails.
        source = self.standalone_document(snippet, "", packages=_TIKZ_PACKAGES)
        return self.compile_to_pdf(source, key="tikz-bare")

    def render_math(self, latex: str, display: bool = True,
                    preamble: str = "") -> Optional[str]:
        """Compile a single formula, returning a PDF path."""
        body = ("\\(\\displaystyle %s\\)" % latex) if display else ("\\(%s\\)" % latex)
        source = self.standalone_document(body, preamble, border="1pt",
                                          class_options="varwidth=15cm")
        pdf = self.compile_to_pdf(source, key="math")
        if pdf:
            return pdf
        source = self.standalone_document(body, "", border="1pt")
        return self.compile_to_pdf(source, key="math-bare")


def clean_preamble(preamble: str) -> str:
    r"""Strip the Beamer-only parts of a preamble so ``standalone`` can use it."""
    if not preamble:
        return ""
    kept: List[str] = []
    for line in preamble.splitlines():
        if _PREAMBLE_DROP.match(line):
            continue
        if "\\newcommand" in line or "\\renewcommand" in line:
            # Keep user macros: snippets often depend on them.
            kept.append(line)
            continue
        if line.strip().startswith("\\") or line.strip().startswith("%"):
            kept.append(line)
    text = "\n".join(kept)
    # Beamer-specific commands that break outside a presentation.
    text = re.sub(r"\\mode\s*<[^>]*>\s*\{[^{}]*\}", "", text)
    return text.strip() + "\n" if text.strip() else ""


def _first_error(log_text: str) -> str:
    for line in log_text.splitlines():
        if line.startswith("!") or re.match(r"^.+:\d+: ", line):
            return line.strip()[:200]
    return ""


# --------------------------------------------------------------------------
# PDF rasterisation
# --------------------------------------------------------------------------


def pdf_to_png(pdf_path: str, out_path: str, dpi: int = 300,
               transparent: bool = False, crop: bool = False,
               page: int = 0) -> Optional[Tuple[int, int]]:
    """Rasterise one PDF page to PNG; returns the pixel size."""
    try:
        import pypdfium2 as pdfium
    except ImportError:  # pragma: no cover - dependency is declared
        log.error("pypdfium2 is required to rasterise PDF figures")
        return None

    try:
        document = pdfium.PdfDocument(pdf_path)
    except Exception as exc:
        log.warning("Could not open %s: %s", pdf_path, exc)
        return None

    try:
        if len(document) == 0:
            return None
        pdf_page = document[min(page, len(document) - 1)]
        fill = (255, 255, 255, 0) if transparent else (255, 255, 255, 255)
        bitmap = pdf_page.render(scale=dpi / 72.0, fill_color=fill,
                                 draw_annots=False)
        image = bitmap.to_pil()
        if transparent:
            image = image.convert("RGBA")
        if crop:
            image = _autocrop(image)
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        image.save(out_path, "PNG", dpi=(dpi, dpi))
        return image.size
    except Exception as exc:
        log.warning("Could not rasterise %s: %s", pdf_path, exc)
        return None
    finally:
        try:
            document.close()
        except Exception:
            pass


def pdf_to_svg(pdf_path: str, out_path: str, page: int = 0) -> bool:
    """Convert one PDF page to SVG, keeping it vector.

    A compiled TikZ picture is vector art; rasterising it throws that away.
    Embedded as SVG instead, it stays sharp at any zoom and PowerPoint can
    turn it into native shapes ("Convert to Shape").
    """
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)

    cairo = find_executable("pdftocairo") or find_executable("miktex-pdftocairo")
    if cairo:
        command = [cairo, "-svg", "-f", str(page + 1), "-l", str(page + 1),
                   pdf_path, out_path]
        if _run_quiet(command) and _is_usable_svg(out_path):
            return True

    # dvisvgm reads PDF too; --no-fonts converts glyphs to paths, which keeps
    # the result independent of the fonts installed on the viewer's machine.
    dvisvgm = find_executable("dvisvgm") or find_executable("miktex-dvisvgm")
    if dvisvgm:
        command = [dvisvgm, "--pdf", "--no-fonts", "--page=%d" % (page + 1),
                   "--output=%s" % out_path, pdf_path]
        if _run_quiet(command) and _is_usable_svg(out_path):
            return True

    return False


def _run_quiet(command: List[str], timeout: int = 120) -> bool:
    try:
        subprocess.run(command, capture_output=True, timeout=timeout, check=False)
        return True
    except (OSError, subprocess.TimeoutExpired) as exc:
        log.debug("%s failed: %s", command[0], exc)
        return False


def _is_usable_svg(path: str) -> bool:
    try:
        if os.path.getsize(path) < 64:
            return False
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return "<svg" in fh.read(4096)
    except OSError:
        return False


def _autocrop(image, padding: int = 4):
    """Trim uniform margins, keeping a small padding."""
    if image.mode == "RGBA":
        bbox = image.getchannel("A").getbbox()
    else:
        grey = image.convert("L")
        inverted = grey.point(lambda v: 255 - v)
        bbox = inverted.getbbox()
    if not bbox:
        return image
    left = max(0, bbox[0] - padding)
    top = max(0, bbox[1] - padding)
    right = min(image.width, bbox[2] + padding)
    bottom = min(image.height, bbox[3] + padding)
    return image.crop((left, top, right, bottom))
