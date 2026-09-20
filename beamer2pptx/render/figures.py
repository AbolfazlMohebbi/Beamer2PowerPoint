r"""Resolving ``\includegraphics`` targets and normalising them to PNG.

Every figure that ends up on a slide is written into the output ``figures/``
directory, whatever format it started in: rasters are copied, PDF and EPS are
rendered through the TeX toolchain, SVG goes through ``svglib``.
"""

from __future__ import annotations

import hashlib
import logging
import os
import re
import shutil
import subprocess
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .latexrun import LatexRunner, find_executable, pdf_to_png, pdf_to_svg

log = logging.getLogger(__name__)

#: Extensions tried, in order, when the document omits one.
SEARCH_EXTENSIONS = (".pdf", ".png", ".jpg", ".jpeg", ".eps", ".ps", ".svg",
                     ".gif", ".bmp", ".tif", ".tiff", ".webp")

#: Formats PowerPoint can embed directly.
NATIVE_RASTER = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif"}

#: Formats that carry vector art worth preserving as SVG.
VECTOR_SOURCES = {".pdf", ".eps", ".ps", ".svg"}

_SAFE_NAME = re.compile(r"[^A-Za-z0-9_.-]+")


@dataclass
class PlacedImage:
    """A figure written into ``figures/``.

    ``png_path`` is always present — PowerPoint needs a raster fallback even
    when the vector version is the one it displays.
    """

    png_path: str
    size: Tuple[int, int] = (0, 0)
    svg_path: Optional[str] = None

    @property
    def is_vector(self) -> bool:
        return bool(self.svg_path and os.path.isfile(self.svg_path))


@dataclass
class FigureStore:
    """Owns the output ``figures/`` directory and the conversion cache."""

    out_dir: str
    root: str = "."
    search_paths: List[str] = field(default_factory=list)
    runner: Optional[LatexRunner] = None
    dpi: int = 300
    #: Keep vector art as SVG rather than rasterising it.
    vector: bool = True
    #: Source path -> what we wrote, so a figure reused across slides is
    #: converted once and stored once.
    _written: Dict[str, PlacedImage] = field(default_factory=dict, init=False)
    _used_names: set = field(default_factory=set, init=False)
    failures: List[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        os.makedirs(self.out_dir, exist_ok=True)

    # -- resolution -------------------------------------------------------

    def resolve(self, reference: str) -> Optional[str]:
        r"""Find the file a ``\includegraphics`` argument refers to."""
        if not reference:
            return None
        reference = reference.strip().strip('"').replace("\\", "/")
        # LaTeX ignores a leading ./ and tolerates a missing extension.
        candidates: List[str] = []
        bases = [self.root] + [os.path.join(self.root, p) for p in self.search_paths]
        bases += [p for p in self.search_paths if os.path.isabs(p)]

        if os.path.isabs(reference):
            bases = [""]

        for base in bases:
            path = os.path.normpath(os.path.join(base, reference)) if base else reference
            candidates.append(path)
            if not os.path.splitext(path)[1]:
                candidates.extend(path + ext for ext in SEARCH_EXTENSIONS)
            else:
                # A named .eps may really be a .pdf next to it, and vice versa.
                stem = os.path.splitext(path)[0]
                candidates.extend(stem + ext for ext in SEARCH_EXTENSIONS)

        for candidate in candidates:
            if os.path.isfile(candidate):
                return os.path.abspath(candidate)

        # Last resort: match on basename anywhere under the project root.
        target = os.path.basename(reference).lower()
        stem = os.path.splitext(target)[0]
        for dirpath, _dirnames, filenames in os.walk(self.root):
            for name in filenames:
                if name.lower() == target or os.path.splitext(name)[0].lower() == stem:
                    if os.path.splitext(name)[1].lower() in SEARCH_EXTENSIONS:
                        return os.path.abspath(os.path.join(dirpath, name))
        return None

    # -- conversion -------------------------------------------------------

    def place(self, source_path: str, prefix: str = "fig") -> Optional[PlacedImage]:
        """Copy or convert *source_path* into ``figures/``."""
        if source_path in self._written:
            return self._written[source_path]

        extension = os.path.splitext(source_path)[1].lower()
        target = self._unique_name(prefix, source_path,
                                   extension if extension in NATIVE_RASTER else ".png")
        svg_target: Optional[str] = None

        if extension in NATIVE_RASTER:
            try:
                shutil.copy2(source_path, target)
            except OSError as exc:
                self.failures.append("Could not copy %s: %s" % (source_path, exc))
                return None
        elif extension == ".pdf":
            if pdf_to_png(source_path, target, dpi=self.dpi) is None:
                self.failures.append("Could not rasterise %s" % source_path)
                return None
            svg_target = self._vector_from_pdf(source_path, target)
        elif extension in (".eps", ".ps"):
            pdf = self._eps_to_pdf(source_path, target)
            if pdf is None or pdf_to_png(pdf, target, dpi=self.dpi) is None:
                if not self._eps_to_png(source_path, target):
                    self.failures.append("Could not convert %s" % source_path)
                    return None
            elif pdf:
                svg_target = self._vector_from_pdf(pdf, target)
            if pdf and os.path.isfile(pdf):
                os.remove(pdf)
        elif extension == ".svg":
            if not self._svg_to_png(source_path, target):
                self.failures.append("Could not convert %s" % source_path)
                return None
            if self.vector:
                # The source is already vector: carry it through untouched.
                svg_target = os.path.splitext(target)[0] + ".svg"
                shutil.copy2(source_path, svg_target)
        else:
            if not self._pillow_convert(source_path, target):
                self.failures.append("Unsupported image format: %s" % source_path)
                return None

        placed = PlacedImage(target, _image_size(target), svg_target)
        self._written[source_path] = placed
        return placed

    def place_pdf(self, pdf_path: str, prefix: str, transparent: bool = False,
                  crop: bool = True, dpi: Optional[int] = None,
                  vector: bool = True) -> Optional[PlacedImage]:
        """Place a generated PDF (TikZ or an equation) into ``figures/``.

        The PNG is always written — PowerPoint uses it as the fallback that
        older viewers and thumbnails display — and the SVG alongside it is
        what PowerPoint actually renders when it can.
        """
        target = self._unique_name(prefix, pdf_path, ".png")
        size = pdf_to_png(pdf_path, target, dpi=dpi or self.dpi,
                          transparent=transparent, crop=crop)
        if size is None:
            return None
        svg_target = self._vector_from_pdf(pdf_path, target) if vector else None
        return PlacedImage(target, size, svg_target)

    def _vector_from_pdf(self, pdf_path: str, png_target: str) -> Optional[str]:
        """Write the SVG twin of a PDF next to its PNG, if we can."""
        if not self.vector:
            return None
        svg_target = os.path.splitext(png_target)[0] + ".svg"
        if pdf_to_svg(pdf_path, svg_target):
            return svg_target
        log.debug("No vector conversion for %s; using the raster version.", pdf_path)
        return None

    # -- format helpers ---------------------------------------------------

    def _eps_to_pdf(self, source: str, target: str) -> Optional[str]:
        """Turn EPS/PS into a PDF we can then rasterise and vectorise."""
        pdf_target = os.path.splitext(target)[0] + ".tmp.pdf"

        ghostscript = (find_executable("mgs") or find_executable("gswin64c")
                       or find_executable("gswin32c") or find_executable("gs"))
        if ghostscript:
            command = [ghostscript, "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER",
                       "-dEPSCrop", "-sDEVICE=pdfwrite",
                       "-sOutputFile=" + pdf_target, source]
            try:
                subprocess.run(command, capture_output=True, timeout=120, check=False)
            except (OSError, subprocess.TimeoutExpired) as exc:
                log.warning("Ghostscript failed on %s: %s", source, exc)
            if os.path.isfile(pdf_target):
                return pdf_target

        # Fall back to letting LaTeX embed it via epstopdf.
        if self.runner and self.runner.available:
            body = "\\includegraphics{%s}" % source.replace("\\", "/")
            document = self.runner.standalone_document(
                body, "", packages=("graphicx", "epstopdf"))
            pdf = self.runner.compile_to_pdf(document, key="eps")
            if pdf:
                shutil.copy(pdf, pdf_target)
                return pdf_target
        return None

    def _eps_to_png(self, source: str, target: str) -> bool:
        """Last-resort EPS rasterisation when the PDF route did not work."""
        pdf = self._eps_to_pdf(source, target)
        if pdf is None:
            return False
        ok = pdf_to_png(pdf, target, dpi=self.dpi) is not None
        if os.path.isfile(pdf):
            os.remove(pdf)
        return ok

    def _svg_to_png(self, source: str, target: str) -> bool:
        try:
            from reportlab.graphics import renderPDF
            from svglib.svglib import svg2rlg
        except ImportError:
            log.warning("svglib is not installed; cannot convert %s", source)
            return False
        try:
            drawing = svg2rlg(source)
            if drawing is None:
                return False
            pdf_target = os.path.splitext(target)[0] + ".tmp.pdf"
            renderPDF.drawToFile(drawing, pdf_target)
            ok = pdf_to_png(pdf_target, target, dpi=self.dpi) is not None
            if os.path.isfile(pdf_target):
                os.remove(pdf_target)
            return ok
        except Exception as exc:
            log.warning("Could not convert SVG %s: %s", source, exc)
            return False

    def _pillow_convert(self, source: str, target: str) -> bool:
        try:
            from PIL import Image

            with Image.open(source) as image:
                image.convert("RGBA" if "A" in image.getbands() else "RGB").save(target, "PNG")
            return True
        except Exception as exc:
            log.warning("Could not read %s: %s", source, exc)
            return False

    # -- naming -----------------------------------------------------------

    def _unique_name(self, prefix: str, source: str, extension: str) -> str:
        stem = _SAFE_NAME.sub("_", os.path.splitext(os.path.basename(source))[0])[:40]
        digest = hashlib.sha1(source.encode("utf-8", "replace")).hexdigest()[:6]
        name = "%s_%s_%s%s" % (prefix, stem or "image", digest, extension)
        candidate = os.path.join(self.out_dir, name)
        counter = 2
        while candidate in self._used_names:
            candidate = os.path.join(
                self.out_dir, "%s_%s_%s_%d%s" % (prefix, stem, digest, counter, extension))
            counter += 1
        self._used_names.add(candidate)
        return candidate


def _image_size(path: str) -> Tuple[int, int]:
    try:
        from PIL import Image

        with Image.open(path) as image:
            return image.size
    except Exception:
        return (0, 0)
