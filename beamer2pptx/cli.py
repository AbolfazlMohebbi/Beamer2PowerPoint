r"""Command-line entry point.

Orchestrates the pipeline — open, preprocess, parse, prepare assets, build —
and writes a run report listing everything that needed a compromise.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import time
from typing import List, Optional

from . import __version__
from .build.slidebuilder import BuildOptions, SlideBuilder
from .build.template import load_template
from .ir import Deck
from .parse.document import ParseOptions, parse_document
from .parse.inline import InlineContext
from .preprocess import preprocess
from .project import open_project
from .render.figures import FigureStore
from .render.latexrun import LatexRunner
from .render.math_omml import MathConverter
from .render.prepare import AssetPreparer

log = logging.getLogger("beamer2pptx")

DEFAULT_TEMPLATE = "template.pptx"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="beamer2pptx",
        description="Convert a Beamer presentation (.tex or a .zip project) "
                    "into a PowerPoint deck built on your own template.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input", help="Beamer .tex file, project folder, or .zip")
    parser.add_argument("-o", "--output", help="output .pptx (default: alongside the input)")
    parser.add_argument("-t", "--template", default=DEFAULT_TEMPLATE,
                        help="PowerPoint template to build on")
    parser.add_argument("--main", help="master .tex inside a multi-file project")

    parser.add_argument("--figures-dir", default="figures",
                        help="directory for extracted figures, relative to the output")
    parser.add_argument("--dpi", type=int, default=300,
                        help="rasterisation resolution for vector figures")
    parser.add_argument("--equation-dpi", type=int, default=600,
                        help="resolution for equations that fall back to images")
    parser.add_argument("--vector", dest="vector", action="store_true", default=True,
                        help="keep TikZ and PDF figures as scalable SVG")
    parser.add_argument("--no-vector", dest="vector", action="store_false",
                        help="embed every figure as a raster image only")

    parser.add_argument("--overlays", choices=("flatten", "expand"), default="flatten",
                        help="flatten each frame to one slide, or expand overlay steps")
    parser.add_argument("--math", choices=("native", "image"), default="native",
                        help="native PowerPoint equations, or rendered images")
    parser.add_argument("--section-slides", dest="section_slides",
                        action="store_true", default=True,
                        help="emit a divider slide for each \\section")
    parser.add_argument("--no-section-slides", dest="section_slides",
                        action="store_false", help=argparse.SUPPRESS)
    parser.add_argument("--subsection-slides", action="store_true",
                        help="also emit a divider slide for each \\subsection")

    parser.add_argument("--min-font-size", type=float, default=14.0,
                        help="smallest body font the fitter may shrink to")
    parser.add_argument("--alert-color", default="C00000",
                        help="colour for \\alert and alertblock, as RRGGBB")

    parser.add_argument("--layout-title", help="template layout name for title slides")
    parser.add_argument("--layout-content", help="template layout name for content slides")
    parser.add_argument("--layout-section", help="template layout name for section slides")

    parser.add_argument("--latex", help="path to pdflatex (autodetected by default)")
    parser.add_argument("--no-latex", action="store_true",
                        help="never invoke LaTeX; TikZ becomes a placeholder")
    parser.add_argument("--mml2omml", help="path to MML2OMML.XSL")
    parser.add_argument("--no-expand-macros", action="store_true",
                        help="do not expand user \\newcommand definitions")

    parser.add_argument("--cache-dir", help="LaTeX build cache (default: .b2p-cache "
                                            "next to the output)")
    parser.add_argument("--no-cache", action="store_true", help="disable the LaTeX cache")
    parser.add_argument("--keep-temp", action="store_true",
                        help="keep an extracted .zip project on disk")
    parser.add_argument("--report", nargs="?", const="", default=None,
                        help="write a Markdown run report (default: next to the output)")
    parser.add_argument("--dump-ir", help="write the parsed structure as JSON")
    parser.add_argument("-q", "--quiet", action="store_true")
    parser.add_argument("-v", "--verbose", action="count", default=0)
    parser.add_argument("--version", action="version", version="beamer2pptx " + __version__)
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    _configure_logging(args)

    started = time.time()

    template_path = _resolve_template(args.template, args.input)
    if template_path is None:
        log.error("Template not found: %s", args.template)
        return 2

    output_path = _resolve_output(args)
    out_dir = os.path.dirname(os.path.abspath(output_path)) or "."
    figures_dir = (args.figures_dir if os.path.isabs(args.figures_dir)
                   else os.path.join(out_dir, args.figures_dir))
    cache_dir = None if args.no_cache else (args.cache_dir
                                            or os.path.join(out_dir, ".b2p-cache"))

    project = open_project(args.input, args.main)
    log.info("Main document: %s", project.main_tex)

    try:
        source = preprocess(project.main_tex, project.root,
                            expand=not args.no_expand_macros)
        log.info("Read %d file(s), %d macro definition(s)",
                 len(source.files), len(source.macros))

        parse_options = ParseOptions(section_slides=args.section_slides,
                                     subsection_slides=args.subsection_slides,
                                     overlays=args.overlays,
                                     alert_color=args.alert_color)
        ctx = InlineContext(verbatims=source.verbatims,
                            alert_color=args.alert_color)
        deck = parse_document(source, parse_options, ctx)
        log.info("Parsed %d slide(s)", len(deck.slides))

        if args.dump_ir:
            _dump_ir(deck, args.dump_ir)

        runner = LatexRunner(cache_dir=cache_dir, dpi=args.dpi,
                             enabled=not args.no_latex)
        if args.latex:
            runner.engine = args.latex
        if cache_dir:
            os.makedirs(cache_dir, exist_ok=True)

        store = FigureStore(out_dir=figures_dir, root=project.root,
                            search_paths=source.graphicspaths, runner=runner,
                            dpi=args.dpi, vector=args.vector)
        math = MathConverter(xsl_path=args.mml2omml, enabled=args.math == "native")

        if args.math == "native" and not math.available:
            log.warning("Native equations unavailable; falling back to images.")

        preparer = AssetPreparer(store=store, runner=runner, math=math,
                                 preamble=source.preamble, math_mode=args.math,
                                 equation_dpi=args.equation_dpi)
        prepare_report = preparer.prepare(deck)

        binding = load_template(template_path, args.layout_title,
                                args.layout_content, args.layout_section)
        build_options = BuildOptions(min_font_pt=args.min_font_size,
                                     alert_color=args.alert_color)
        builder = SlideBuilder(binding, math, build_options)
        presentation = builder.build(deck)

        os.makedirs(out_dir, exist_ok=True)
        presentation.save(output_path)
    finally:
        project.cleanup(keep=args.keep_temp)

    elapsed = time.time() - started
    report_text = _format_report(deck, prepare_report, builder.report, store,
                                 runner, math, output_path, figures_dir, elapsed)

    if args.report is not None:
        report_path = args.report or os.path.splitext(output_path)[0] + "-report.md"
        with open(report_path, "w", encoding="utf-8") as fh:
            fh.write(report_text)
        log.info("Report written to %s", report_path)

    if not args.quiet:
        print(report_text)

    return 0


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------


def _configure_logging(args) -> None:
    level = logging.WARNING
    if args.verbose == 1:
        level = logging.INFO
    elif args.verbose >= 2:
        level = logging.DEBUG
    if args.quiet:
        level = logging.ERROR
    logging.basicConfig(level=level, format="%(levelname)s %(name)s: %(message)s")


def _resolve_template(template: str, input_path: str) -> Optional[str]:
    candidates = [template]
    if not os.path.isabs(template):
        candidates.append(os.path.join(os.getcwd(), template))
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(input_path)),
                                       template))
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "..", template))
    for candidate in candidates:
        if os.path.isfile(candidate):
            return os.path.abspath(candidate)
    return None


def _resolve_output(args) -> str:
    if args.output:
        return os.path.abspath(args.output)
    base = os.path.abspath(args.input)
    if os.path.isdir(base):
        return os.path.join(base, os.path.basename(base.rstrip(os.sep)) + ".pptx")
    return os.path.splitext(base)[0] + ".pptx"


def _dump_ir(deck: Deck, path: str) -> None:
    import dataclasses
    import json

    def encode(obj):
        if dataclasses.is_dataclass(obj):
            return {k: v for k, v in dataclasses.asdict(obj).items()
                    if k != "omml"}
        if isinstance(obj, tuple):
            return list(obj)
        return str(obj)

    with open(path, "w", encoding="utf-8") as fh:
        json.dump(deck, fh, default=encode, indent=2, ensure_ascii=False)
    log.info("Structure written to %s", path)


def _format_report(deck, prepare_report, build_report, store, runner, math,
                   output_path, figures_dir, elapsed: float) -> str:
    lines: List[str] = []
    add = lines.append

    add("# Conversion report")
    add("")
    add("- **Output**: `%s`" % output_path)
    add("- **Slides**: %d" % build_report.slides)
    add("- **Figures**: %d written to `%s`" % (prepare_report.figures_written,
                                               figures_dir))
    if prepare_report.figures_vector:
        add("- **Vector figures**: %d embedded as scalable SVG "
            "(right-click a picture in PowerPoint and choose *Convert to Shape* "
            "to make it editable)" % build_report.vector_pictures)
    add("- **TikZ compiled**: %d" % prepare_report.tikz_compiled)
    add("- **Tables**: %d" % build_report.tables)
    add("- **Equations**: %d native, %d as images, %d left as text"
        % (build_report.native_equations, build_report.image_equations,
           build_report.text_equations))
    add("- **Time**: %.1fs" % elapsed)
    add("")

    problems = False

    if prepare_report.missing_figures:
        problems = True
        add("## Figures that could not be found")
        add("")
        for name in dict.fromkeys(prepare_report.missing_figures):
            add("- `%s`" % name)
        add("")

    if prepare_report.failed_tikz:
        problems = True
        add("## TikZ pictures that failed to compile")
        add("")
        for reason in dict.fromkeys(prepare_report.failed_tikz):
            add("- %s" % reason)
        add("")

    if store.failures:
        problems = True
        add("## Figure conversion problems")
        add("")
        for failure in dict.fromkeys(store.failures):
            add("- %s" % failure)
        add("")

    if math.failures:
        problems = True
        add("## Formulas that needed a fallback")
        add("")
        for failure in list(dict.fromkeys(math.failures))[:20]:
            add("- `%s`" % failure)
        add("")

    if build_report.overflowed:
        problems = True
        add("## Slides at the font-size floor")
        add("")
        add("These frames hold more than the template comfortably fits; "
            "check them by hand.")
        add("")
        add(", ".join("slide %d" % n for n in build_report.overflowed))
        add("")

    if build_report.shrunk:
        add("## Slides with reduced text")
        add("")
        add(", ".join("slide %d (%.0f%%)" % (n, s * 100)
                      for n, s in build_report.shrunk))
        add("")

    if deck.warnings or build_report.warnings:
        problems = True
        add("## Other notes")
        add("")
        for warning in deck.warnings + build_report.warnings:
            add("- %s" % warning)
        add("")

    if not runner.available and runner.unavailable_reason:
        add("## LaTeX")
        add("")
        add(runner.unavailable_reason)
        add("")

    if not problems:
        add("Everything converted without a fallback.")
        add("")

    return "\n".join(lines)


if __name__ == "__main__":
    sys.exit(main())
