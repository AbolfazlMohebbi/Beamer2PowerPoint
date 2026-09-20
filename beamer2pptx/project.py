"""Locating and unpacking the input presentation.

Accepts either a single ``.tex`` file or a ``.zip`` holding a whole project,
and works out which file is the Beamer master document.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import List, Optional

log = logging.getLogger(__name__)

DOCUMENTCLASS_RE = re.compile(r"\documentclass\s*(\[[^\]]*\])?\s*\{([^}]*)\}")
BEGIN_DOC_RE = re.compile(r"\begin\s*\{document\}")


@dataclass
class Project:
    """A prepared input: the master ``.tex`` plus the directory it lives in."""

    main_tex: str
    root: str
    temp_dir: Optional[str] = None
    all_tex: List[str] = field(default_factory=list)

    def cleanup(self, keep: bool = False) -> None:
        if self.temp_dir and not keep and os.path.isdir(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        elif self.temp_dir and keep:
            log.info("Kept extracted project at %s", self.temp_dir)


def _read(path: str) -> str:
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(path, "r", encoding=encoding) as fh:
                return fh.read()
        except UnicodeDecodeError:
            continue
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        return fh.read()


def read_text(path: str) -> str:
    """Read a LaTeX source file, tolerating the usual encoding zoo."""
    return _read(path)


def _score_main_candidate(text: str) -> int:
    """Higher is more likely to be the master document."""
    score = 0
    m = DOCUMENTCLASS_RE.search(text)
    if m:
        score += 10
        cls = m.group(2).strip()
        if cls == "beamer":
            score += 40
        if (m.group(1) or "").find("beamer") >= 0:
            score += 10
    if BEGIN_DOC_RE.search(text):
        score += 20
    score += min(text.count(r"\begin{frame}") + text.count(r"\frame{"), 20)
    return score


def find_main_tex(root: str, explicit: Optional[str] = None) -> str:
    """Pick the master ``.tex`` inside *root*."""
    if explicit:
        candidate = explicit if os.path.isabs(explicit) else os.path.join(root, explicit)
        if not os.path.isfile(candidate):
            raise FileNotFoundError(f"--main file not found: {candidate}")
        return os.path.abspath(candidate)

    candidates = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(("__MACOSX", "."))]
        for name in filenames:
            if name.lower().endswith(".tex"):
                candidates.append(os.path.join(dirpath, name))

    if not candidates:
        raise FileNotFoundError(f"No .tex file found under {root}")
    if len(candidates) == 1:
        return os.path.abspath(candidates[0])

    scored = []
    for path in candidates:
        try:
            score = _score_main_candidate(_read(path))
        except OSError:
            continue
        # Prefer shallow files and conventional names on ties.
        depth = os.path.relpath(path, root).count(os.sep)
        name = os.path.basename(path).lower()
        bonus = 5 if name in ("main.tex", "slides.tex", "presentation.tex", "talk.tex") else 0
        scored.append((score + bonus - depth, path))

    scored.sort(key=lambda item: (-item[0], item[1]))
    best_score, best = scored[0]
    if best_score <= 0:
        raise RuntimeError(
            "Could not identify the master .tex file; pass --main explicitly."
        )
    log.debug("Main document: %s (score %d of %d candidates)", best, best_score, len(scored))
    return os.path.abspath(best)


def open_project(input_path: str, main: Optional[str] = None) -> Project:
    """Prepare *input_path* (a ``.tex`` or ``.zip``) for conversion."""
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        raise FileNotFoundError(input_path)

    if os.path.isdir(input_path):
        root = input_path
        return Project(main_tex=find_main_tex(root, main), root=root)

    if zipfile.is_zipfile(input_path):
        temp_dir = tempfile.mkdtemp(prefix="beamer2pptx-")
        with zipfile.ZipFile(input_path) as zf:
            for member in zf.infolist():
                # Refuse paths that would escape the extraction directory.
                target = os.path.normpath(os.path.join(temp_dir, member.filename))
                if not target.startswith(os.path.normpath(temp_dir) + os.sep):
                    log.warning("Skipping unsafe zip entry %s", member.filename)
                    continue
                zf.extract(member, temp_dir)
        root = temp_dir
        # A zip that wraps everything in one folder: descend into it.
        entries = [e for e in os.listdir(root) if not e.startswith("__MACOSX")]
        if len(entries) == 1 and os.path.isdir(os.path.join(root, entries[0])):
            root = os.path.join(root, entries[0])
        return Project(main_tex=find_main_tex(root, main), root=root, temp_dir=temp_dir)

    if not input_path.lower().endswith((".tex", ".ltx")):
        log.warning("Input %s is neither a .zip nor a .tex; treating it as LaTeX.", input_path)
    return Project(main_tex=input_path, root=os.path.dirname(input_path) or ".")
