# beamer2pptx

Convert a Beamer presentation — a single `.tex` file or a whole multi-file
project — into a PowerPoint deck built on **your** template.

The output is a real PowerPoint file, not a stack of slide images: text stays
editable in the template's own placeholders, tables become native PowerPoint
tables, and equations become native, editable PowerPoint equations. Every
figure is extracted or rendered into a `figures/` folder next to the output.

```bash
python beamer2pptx.py talk.tex -o talk.pptx --template template.pptx
```

## What it converts

| Beamer | PowerPoint |
|---|---|
| `\begin{frame}{Title}` | one slide on the *Title and Content* layout |
| `\titlepage` / `\maketitle` | the template's title-slide layout, with author, institute and date |
| `\section` | a divider slide (`--no-section-slides` to skip) |
| `itemize` / `enumerate` / `description` | bullets at the template's own outline levels |
| `\textbf`, `\emph`, `\texttt`, `\alert`, `\textcolor` | run formatting and colour |
| `\url`, `\href` | live hyperlinks |
| `$x$`, `\[...\]`, `equation`, `align`, `gather`, `cases`, matrices | **native OMML equations** |
| `tabular`, `tabularx`, `longtable`, `booktabs`, `\multicolumn`, `\multirow` | **native tables** with merges and rules |
| `\includegraphics` (PNG, JPG, PDF, EPS, SVG) | a picture in `figures/`; vector sources stay vector |
| `tikzpicture`, `pgfplots` | compiled with LaTeX, embedded as **scalable SVG** |
| `columns` / `column` | side-by-side regions |
| `block`, `alertblock`, `exampleblock`, `theorem`, `definition`… | a titled, theme-coloured block |
| `verbatim`, `lstlisting`, `\verb` | monospaced, literal text |
| `\pause`, `<2->`, `\only`, `\uncover`, `\alt` | flattened by default, or one slide per step |
| `\note{...}` | speaker notes |
| `\input`, `\include`, `\subfile`, `\graphicspath` | resolved across the project |
| `\newcommand`, `\def`, `\DeclareMathOperator` | expanded before parsing |

Accents (`\'e`), quotes, dashes and the usual symbol macros are translated to
Unicode. A macro the converter does not know still contributes its text, and
is listed in the run report rather than silently dropped.

## Requirements

```bash
pip install -r requirements.txt
```

Two optional pieces make the conversion better, and both are detected
automatically:

- **A LaTeX installation** (MiKTeX or TeX Live) — needed to compile TikZ and
  pgfplots pictures, and to render any equation the native converter cannot
  handle. Without it those become labelled placeholders and the deck still
  builds. On Windows: `winget install MiKTeX.MiKTeX`.
- **Microsoft Office** — its `MML2OMML.XSL` stylesheet is what turns LaTeX
  into native PowerPoint equations. Without it, equations are rendered as
  images instead (which needs LaTeX), or kept as LaTeX source as a last resort.

## Usage

```bash
python beamer2pptx.py INPUT [-o OUTPUT.pptx] [-t TEMPLATE.pptx] [options]
```

`INPUT` may be a `.tex` file, a project directory, or a `.zip`. For a project,
the master document is the one with `\documentclass{beamer}`; use `--main` if
that guess is wrong.

### Options that matter most

| Option | Default | Effect |
|---|---|---|
| `-t, --template` | `template.pptx` | the PowerPoint template to build on |
| `--figures-dir` | `figures` | where extracted figures go, relative to the output |
| `--overlays` | `flatten` | `flatten` = one slide per frame; `expand` = one slide per overlay step |
| `--math` | `native` | `native` = editable equations; `image` = always render pictures |
| `--min-font-size` | `14` | the floor the text fitter may shrink to |
| `--dpi` | `300` | resolution of the raster fallback |
| `--no-vector` | off | rasterise everything instead of embedding SVG |
| `--no-section-slides` | off | do not emit a divider slide per `\section` |
| `--alert-color` | `C00000` | colour for `\alert` and `alertblock` |
| `--report [PATH]` | — | write the run report to a Markdown file |
| `--dump-ir PATH` | — | dump the parsed structure as JSON, for debugging |
| `--no-latex` | off | never invoke LaTeX |
| `--layout-title/-content/-section` | autodetected | force a template layout by name or index |

Run `python beamer2pptx.py --help` for the rest.

### Examples

```bash
# A zipped Overleaf project, with one slide per overlay step
python beamer2pptx.py project.zip -o talk.pptx --overlays expand
```

```bash
# A dense deck: allow smaller text and write a report of what needed fixing
python beamer2pptx.py talk.tex --min-font-size 11 --report
```

## Why TikZ is compiled, and how it stays editable

TikZ is not a diagram format — it is a programming language running inside
TeX. A coordinate may be polar (`(30:2cm)`), relative (`++(1,0)`), an anchor on
another node (`(a.north east)`), an interpolation (`($(a)!0.5!(b)$)`), or an
arbitrary pgfmath expression. A node's size is whatever TeX's typesetter
produces for its contents, and everything anchored to that node moves with it.
Styles are macros with inheritance and loops; pgfplots alone is tens of
thousands of lines of TeX that choose axis ranges, tick positions and legend
layout.

So a hand-written translator would handle literal straight lines and little
else — and its failure mode is not an error but a *subtly wrong diagram*.
Compiling with the author's own preamble gives geometry that is right by
construction, for every package, including ones this converter has never
heard of.

What compiling does **not** have to cost you is resolution or editability.
The compiled PDF is converted to SVG (`pdftocairo`, with `dvisvgm` as a
fallback) and embedded the way PowerPoint stores an inserted SVG: a normal
picture whose blip carries an `asvg:svgBlip` extension, with a PNG alongside
as the fallback for older viewers and thumbnails. The result is that a TikZ
picture:

- stays sharp at any zoom or projector resolution;
- is reported by PowerPoint as a graphic, not a bitmap;
- can be turned into native, recolourable PowerPoint shapes with
  right-click → **Convert to Shape**.

The same applies to `\includegraphics` of a PDF or EPS figure. A source that
is genuinely raster (PNG, JPG) is passed through as a raster — there is no
vector information to preserve. `--no-vector` turns all of this off.

## Using your own template

Nothing is hardcoded to any particular `.pptx`. On startup the converter:

1. opens the template and removes its sample slides, keeping the master,
   layouts and theme;
2. finds a title layout, a content layout and a section layout — by name
   first, then by layout type, then by which placeholders they contain;
3. reads the title and body placeholder geometry **at run time**, following
   layout → master inheritance the way PowerPoint does, so a layout that
   inherits its geometry still gives the right content region;
4. reads the theme's body font and accent colours, and uses them for block
   headings, custom bullet markers and measuring text.

If the automatic choice is wrong, name the layout you want:

```bash
python beamer2pptx.py talk.tex --layout-content "Title and Content"
```

## How content is fitted

Beamer frames are usually denser than a PowerPoint template expects. Each
slide is measured with real font metrics (Pillow, using the theme's own
typeface), wrapping text at the placeholder width. If the content does not
fit, every font size is scaled down in 5 % steps until it does, stopping at
`--min-font-size`. Slides that hit that floor are listed in the run report so
you know exactly which ones to look at by hand.

The builder writes the same line spacing and paragraph spacing it assumed when
measuring, so the estimate and what PowerPoint renders cannot drift apart.

## The run report

Every conversion prints a summary, and `--report` saves it. It lists figures
that could not be found, TikZ pictures that failed to compile, formulas that
needed a fallback, slides that hit the font-size floor, and macros that were
rendered as plain text. Per-slide warnings are also written into that slide's
speaker notes, so problems travel with the deck.

```
# Conversion report

- **Output**: `talk.pptx`
- **Slides**: 13
- **Figures**: 4 written to `figures`
- **Vector figures**: 3 embedded as scalable SVG (right-click a picture in PowerPoint and choose *Convert to Shape* to make it editable)
- **TikZ compiled**: 2
- **Tables**: 1
- **Equations**: 9 native, 0 as images, 0 left as text
- **Time**: 8.4s

Everything converted without a fallback.
```

## Caching

Compiled TikZ pictures and equation images are cached in `.b2p-cache` next to
the output, keyed on a hash of the snippet and the document preamble.
Re-converting a deck after an edit only recompiles what changed — typically
turning an eight-second run into a fraction of a second. `--no-cache` disables
it.

## Project layout

```
beamer2pptx.py            CLI entry point
beamer2pptx/
  project.py              .tex / .zip input, master-document detection
  preprocess.py           comments, verbatim, \input inlining, macro expansion
  ir.py                   the intermediate representation
  parse/
    nodes.py              the LaTeX scanner
    inline.py             styled runs
    blocks.py             frame contents to blocks
    tables.py             tabular to a Table
    document.py           metadata, sections, frames
    overlays.py           <...> specs, slide expansion
  render/
    latexrun.py           cached pdflatex driver, PDF to PNG and SVG
    figures.py            figure resolution and format conversion
    math_omml.py          LaTeX to MathML to OMML
    prepare.py            resolves every asset before layout
  build/
    template.py           template binding and geometry
    layout.py             measurement and fitting
    slidebuilder.py       writes the slides
    oxml.py               raw OOXML helpers
tests/                    pytest suite, with fixture decks
```

Parsing and emission are deliberately separated by the IR in `ir.py`: the
builder never sees LaTeX, so a new Beamer construct needs a parser change and
at most one new block type.

## Tests

```bash
python -m pytest tests/ -q
```

The suite covers the preprocessor, the parser, tables, the OMML conversion,
and full end-to-end conversions that reopen the generated `.pptx` and check
the slide count, the layouts used, native equations, table merges and rules,
embedded images, SVG parts with resolvable relationships, that no shape falls
outside the slide, and that missing figures and a missing LaTeX installation
degrade gracefully instead of failing. Tests needing Office's XSL skip themselves when it is absent.

## Known limitations

- Overlay `expand` mode duplicates slides; it does not generate PowerPoint
  click animations.
- `\only`/`\alt` take their first branch when flattening.
- Beamer themes are not reproduced — the point is to adopt *your* template's
  design instead.
- Bibliographies (`\cite`) render as bracketed keys; no reference list is
  generated.
