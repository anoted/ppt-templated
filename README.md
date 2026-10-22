# ppt-templated

**Beamer-style, fully editable PowerPoint decks from a small YAML file.**

Write what goes on each slide in YAML, pick a template, run one command, and get a normal `.pptx`.
All formatting comes from the PowerPoint template, so anyone can keep editing the result in PowerPoint.

```
talk.yaml  +  template.pptx  ──ppt_templated.py──▶  talk.pptx  (+ slide previews)
 (what)        (how it looks)                       (native, editable)
```

---

## Contents

- [Quick start](#quick-start)
- [What you get](#what-you-get)
- [Templates](#templates) (default: **Beamer Oxford**)
- [Example decks](#example-decks)
- [Documentation](#documentation)
- [YAML at a glance](#yaml-at-a-glance)
- [Working loop](#working-loop)
- [Good to know](#good-to-know)
- [Project layout](#project-layout)
- [Maintaining the docs](#maintaining-the-docs)

---

## Quick start

```bash
pip install -r requirements.txt     # python-pptx, pyyaml, pillow, lxml, matplotlib, numpy, pandoc
python ppt_templated.py examples/example.yaml --preview
```

Python 3.9+. PowerPoint is needed only to *open* decks (and, on Windows, gives the most exact previews).

```bash
python ppt_templated.py talk.yaml                                  # → talk.pptx (or meta.output)
python ppt_templated.py talk.yaml --preview                        # + slide PNGs, contact sheets, layout check
python ppt_templated.py talk.yaml --check                          # validate only, write nothing
python ppt_templated.py talk.yaml -t templates/beamer_modern.pptx  # same content, other look
python ppt_templated.py --inspect "Company Template.potx"          # what a template provides
python build_template.py --spec my_theme.yaml                      # a new template from a theme file
```

From Python: `from ppt_templated import build_deck; build_deck("talk.yaml", "talk.pptx")`.

## What you get

| In the YAML | In PowerPoint | Editable afterwards |
|---|---|---|
| title / section / outline / bullets | Real placeholders from the template layouts | Yes; layouts can be re-applied |
| `block`, `alertblock`, `exampleblock` | Grouped shapes in theme colors; blocks side by side share one height | Yes |
| `chart` | **Native chart** | Yes: right-click → Edit Data |
| `plot` | Academic figure (bar, box, line, scatter, confusion matrix) as a **vector** in the template's fonts and colors | Regenerate from YAML |
| `table` | **Native table** (booktabs / striped / grid / plain), auto widths | Yes |
| `math`, `$inline$` | **Native equation** (needs pandoc) | Yes, in the equation editor |
| `image`, `images` | PNG/JPG as is; **SVG and PDF as vectors** with a PNG fallback sized to the slide; a row of figures at one height | Yes |
| `stats` | A row of key numbers with labels | Yes |
| `flow`, `steps` | **Flowcharts and workflows**: native shapes and glued connectors, laid out automatically | Yes: drag a shape, its arrows follow |
| `placeholder` | Dashed box reserving space for a figure you haven't made yet | Yes: Shape Fill → Picture |
| `layout: references` / `closing` | Numbered references split over as many slides as needed / a thank-you slide | Yes |
| `build`, `animate`, `transition` | **Native animations** and transitions, including **Morph** | Yes |
| `code`, `notes` | Monospace box, speaker notes | Yes |

Every build checks the YAML: unknown keys, bad values and overfull slides are reported with the slide
and a suggested fix. `--preview` renders the slides and lets PowerPoint measure real overlaps.

## Templates

All in `templates/`. Switch with `meta.template` or `-t`.

| Template | Style | Inspired by |
|---|---|---|
| **`beamer_oxford.pptx`** ★ **default** | Academic: navy title bar with a gold stripe, square title box; Constantia / Cambria | University navy and gold |
| `beamer_paper.pptx` | Academic, journal: hairline rules, grey text footline; Georgia / Times New Roman | Journal / booktabs |
| `beamer_classic.pptx` | Academic, plain: blue frame titles, no bars; Cambria | Beamer `default` |
| `beamer_modern.pptx` | Clean sans: white, teal accent bar under titles, left-aligned covers; Calibri Light / Calibri | Modern conference talks |
| `beamer_blue.pptx` | Blue title bar and footline boxes; Calibri | Beamer `Madrid` |
| `beamer_charcoal.pptx` | Dark slate title bar, orange alerts; Calibri | Beamer `metropolis` |
| `beamer_crimson.pptx` | Light grey title bar, dark red titles; Cambria / Calibri | Beamer `Beaver` |

Oxford is used whenever neither `template:` nor `-t` is given (the bundled examples set `beamer_blue`).
Your own or your company's template: [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md).

## Example decks

| File | Topic | Shows |
|---|---|---|
| `examples/example.yaml` | ppt-templated itself | Every feature once |
| `transformer/transformer.yaml` | The Transformer architecture | A math-heavy lecture with figure placeholders |
| `segmentation/segmentation.yaml` | Image segmentation methods | A draft deck: placeholders and sample plots |

Build any of them from anywhere, e.g. `python ppt_templated.py segmentation/segmentation.yaml`.

## Documentation

| Document | Read it to |
|---|---|
| **[YAML_GUIDE.md](YAML_GUIDE.md)** | write a content file: **the complete reference** (every key, value and default, figures, plots, flows, colors, animation, layout rules, recipes, messages) |
| [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) | build a template from a theme file, adapt a company template, or restyle one |
| this README | get an overview |

You should not need to read the Python files to write a deck; if you do, the guide is missing something.

## YAML at a glance

```yaml
meta:
  title: My Talk
  author: Jane Doe
  date: today
  transition: fade

slides:
  - layout: title

  - title: Our method in one picture
    content:
      - columns:
          - width: 0.5
            content:
              - image: assets/method.svg            # SVG/PDF/PNG: converted and sized automatically
          - - bullets:
                - "Accuracy up by ==4.2 points=="
                - - on all three datasets
                - 'Loss: $\mathcal{L} = -\sum_t \log p(y_t)$'
              build: true

  - title: Results
    content:
      - stats:
          - {value: "+4.2", label: accuracy points}
          - {value: "3×", label: faster inference}
      - plot:
          type: bar
          categories: [A, B, C]
          series: {Baseline: [70, 72, 68], Ours: [74, 77, 71]}
          ylabel: Accuracy (%)

  - title: Pipeline
    content:
      - flow: {edges: [Data -> Preprocess -> Train -> Evaluate]}
      - steps: [Collect, Clean, Train, Evaluate]
        highlight: 3

  - layout: references
    items: ['A. Author, "A paper," *Venue*, 2024.']
  - layout: closing
```

Slide layouts: `title`, `section`, `outline`, `frame` (default), `plain`, `references`, `closing`.
Content items: `bullets`, `numbered`, `text`, `block`, `alertblock`, `exampleblock`, `math`, `code`,
`image`, `images`, `placeholder`, `table`, `chart`, `plot`, `stats`, `flow`, `steps`, `columns`, `spacer`.
Inline: `**bold**` · `*italic*` · `` `code` `` · `==alert==` · `[link](url)` · `$math$`.

## Working loop

1. Write the YAML (start from the skeleton in [YAML_GUIDE.md](YAML_GUIDE.md#2-skeleton)).
2. `python ppt_templated.py talk.yaml --preview`
3. Fix every `warning:` line; each names the slide and usually the fix.
4. Look at the contact sheets `.preview/talk/sheet-*.png` (6 numbered slides each); repeat.

Before figures exist, use `placeholder:` boxes (with the final `aspect`) and `plot:` with sample
numbers; the layout is final from the start, and swapping in real figures changes nothing else.

## Good to know

- **Fitting.** Text sizes are estimated (from the template fonts' widths) to lay out and shrink
  content; `--preview` with PowerPoint measures the rendered slides and reports real overlaps.
- **Figures.** SVG and PDF stay vector in PowerPoint 2019/365; other viewers show a 300 dpi PNG
  rendered at the exact slide size (needs Chrome/Edge or cairosvg; PDF needs PyMuPDF). Converted
  files are cached in `.ppt-cache/`.
- **Equations** need pandoc when you *build* the deck. Without it they become Unicode text, which is
  also what LibreOffice and Google Slides show.
- **Slide numbers** are a live field; **footline** text lives in the slide master.
- **Rebuilding overwrites** the output. Keep manual edits in a copy, or move them into the YAML.

## Project layout

```
ppt_templated.py      the generator (CLI: build, --check, --preview, --inspect)
plots.py              draws `plot:` figures (matplotlib)
flows.py              lays out `flow:` / `steps:` diagrams and routes their connectors
media.py              reads figure files: SVG/PDF/raster sizes, conversion, cache
preview.py            renders a .pptx to PNGs + contact sheets; measures the layout via PowerPoint
build_template.py     builds templates from themes (built-in or a theme file)
templates/            seven sample templates (beamer_oxford is the default)
examples/             example.yaml → example_deck.pptx: every feature once
transformer/          The Transformer lecture (YAML + deck)
segmentation/         Image segmentation methods (YAML + deck)
YAML_GUIDE.md         complete content-file reference
TEMPLATE_GUIDE.md     how to make or adapt a template
```

## Maintaining the docs

**[YAML_GUIDE.md](YAML_GUIDE.md) must change in the same commit as `ppt_templated.py`, `plots.py`,
`flows.py`, `media.py` or `preview.py`** whenever a YAML key, value, default, item, layout, shape,
effect or message changes; [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) with `build_template.py`. Update
the "at a glance" part of this README if it is affected.
