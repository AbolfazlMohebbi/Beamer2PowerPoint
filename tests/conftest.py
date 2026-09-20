"""Shared fixtures for the test suite."""

from __future__ import annotations

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

FIXTURES = os.path.join(ROOT, "tests", "fixtures")
TEMPLATE = os.path.join(ROOT, "template.pptx")

HEADER = r"""\documentclass{beamer}
\usepackage{amsmath,amssymb,booktabs,graphicx,tikz}
\title{Test deck}\author{Tester}\date{2026-01-01}
\begin{document}
"""
FOOTER = "\n\\end{document}\n"


@pytest.fixture
def template_path() -> str:
    return TEMPLATE


@pytest.fixture
def kitchen_sink() -> str:
    return os.path.join(FIXTURES, "kitchen_sink", "main.tex")


@pytest.fixture
def write_deck(tmp_path):
    """Write a small Beamer document and return its path."""

    def _write(body: str, name: str = "deck.tex", extra_files=None) -> str:
        path = tmp_path / name
        path.write_text(HEADER + body + FOOTER, encoding="utf-8")
        for filename, content in (extra_files or {}).items():
            target = tmp_path / filename
            target.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                target.write_bytes(content)
            else:
                target.write_text(content, encoding="utf-8")
        return str(path)

    return _write


@pytest.fixture
def parse_body(write_deck):
    """Parse a document body straight into a Deck."""
    from beamer2pptx.parse.document import ParseOptions, parse_document
    from beamer2pptx.preprocess import preprocess
    from beamer2pptx.project import open_project

    def _parse(body: str, options: "ParseOptions | None" = None, **kwargs):
        path = write_deck(body, **kwargs)
        project = open_project(path)
        source = preprocess(project.main_tex, project.root)
        return parse_document(source, options or ParseOptions())

    return _parse
