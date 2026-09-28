# ppt-templated

**Beamer-style, fully editable PowerPoint decks from a small YAML file.**

Write what goes on each slide in YAML, pick a template, run one command, and get a normal `.pptx`.
All formatting comes from the PowerPoint template, so anyone can keep editing the result in PowerPoint.

```
talk.yaml  +  template.pptx  ──ppt_templated.py──▶  talk.pptx
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
- [Draft-first workflow](#draft-first-workflow)
- [Good to know](#good-to-know)
- [Project layout](#project-layout)
- [Maintaining the docs](#maintaining-the-docs)

---

## Quick start

```bash
pip install -r requirements.txt     # python-pptx, pyyaml, pillow, lxml, matplotlib, numpy
pip install pypandoc_binary         # optional: native, editable equations

python ppt_templated.py examples/example.yaml
```

Python 3.9+. PowerPoint is only needed to *open* the decks, not to build them.

```bash
python ppt_templated.py talk.yaml                                  # → talk.pptx (or meta.output)
python ppt_templated.py talk.yaml -o out/talk.pptx                 # explicit output
python ppt_templated.py talk.yaml -t templates/beamer_crimson.pptx # same content, other look
python ppt_templated.py talk.yaml -t "Company Template.potx"       # your own template (.potx works)
```

From Python: `from ppt_templated import build_deck; build_deck("talk.yaml", "talk.pptx")`.

## What you get

| In the YAML | In PowerPoint | Editable afterwards |
|---|---|---|
| title / section / outline / bullets | Real placeholders from the template layouts | Yes; layouts can be re-applied |
| `block`, `alertblock`, `exampleblock` | Grouped rounded shapes in theme colors | Yes |
| `chart` | **Native chart** | Yes: right-click → Edit Data |
| `plot` | Academic figure (bar, box, line, scatter, confusion matrix) in the template's fonts and colors | Regenerate from YAML |
| `table` | **Native table** (booktabs / striped / grid) | Yes |
| `math`, `$inline$` | **Native equation** (needs pandoc) | Yes, in the equation editor |
| `flow`, `steps` | **Flowcharts and workflow diagrams**: native shapes with connectors glued to them, laid out automatically | Yes: drag a shape and its arrows follow |
| `placeholder` | Dashed box reserving space for a diagram you haven't made yet | Yes: Shape Fill → Picture |
| `build`, `animate` | **Native animations** (Animation Pane) | Yes |
| `transition` | Native slide transitions, including **Morph** | Yes |
| `image`, `code`, `notes` | Picture, monospace box, speaker notes | Yes |

## Templates

All in `templates/`. Switch with `meta.template` in the YAML or `-t` on the command line.

| Template | Style | Inspired by |
|---|---|---|
| **`beamer_oxford.pptx`** ★ **default** | Academic, university: navy title bar with a gold stripe, square title box; Constantia headings, Cambria body | University navy-and-gold decks |
| `beamer_paper.pptx` | Academic, journal: hairline rules instead of bars, grey text footline; Georgia / Times New Roman | Journal / booktabs look |
| `beamer_classic.pptx` | Academic, plain: blue frame titles, no bars; Cambria | Beamer `default` |
| `beamer_blue.pptx` | Blue title bar and footline boxes; Calibri | Beamer `Madrid` |
| `beamer_charcoal.pptx` | Minimal: dark slate title bar, orange alerts; Calibri | Beamer `metropolis` |
| `beamer_crimson.pptx` | Light grey title bar with dark red titles; Cambria / Calibri | Beamer `Beaver` |

**Beamer Oxford is the default.** It is used whenever the YAML has no `template:` key and no `-t` is
given. The bundled example decks set `template:` explicitly (to `beamer_blue`), so they override it;
delete that line to get the default.

Your own or your company's template: see [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md).

## Example decks

| Folder | Topic | Shows |
|---|---|---|
| `examples/example.yaml` | ppt-templated itself | Every feature once: blocks, charts, plots, tables, math, flowcharts, animation, Morph |
| `transformer/transformer.yaml` | The Transformer architecture | A math-heavy lecture with figure placeholders |
| `segmentation/segmentation.yaml` | Image segmentation methods | A draft deck: `placeholder:` diagrams and `plot:` sample graphs |

Build any of them from inside its folder, e.g. `cd segmentation && python ../ppt_templated.py segmentation.yaml`.

## Documentation

| Document | Read it when you want to… |
|---|---|
| **[YAML_GUIDE.md](YAML_GUIDE.md)** | write or edit a content file: **the complete YAML reference** (every key, accepted value and default, chart/plot types, flow shapes, colors, animation, gotchas, recipes, error messages) |
| [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) | restyle a template, adapt your company template, or build one in code |
| this README | get an overview and get started |

You should not need to read `ppt_templated.py`, `plots.py` or `flows.py` to write a deck; if you do, the YAML guide
is missing something, so please add it.

## YAML at a glance

A short taste. The full reference is in [YAML_GUIDE.md](YAML_GUIDE.md).

```yaml
meta:
  title: My Talk
  author: Jane Doe
  date: today
  # template: templates/beamer_oxford.pptx   # the default; set another to change the look
  transition: fade

slides:
  - layout: title
  - layout: outline

  - layout: section
    title: Results

  - title: Main result
    content:
      - columns:
          - width: 0.45
            content:
              - placeholder: {label: Pipeline, note: "input → model → output", aspect: "4:3"}
          - content:
              - bullets:
                  - "Accuracy up by ==4.2 points=="
                  - - on all three datasets
                  - 'Loss: $\mathcal{L} = -\sum_t \log p(y_t)$'
                build: true

  - title: Numbers
    content:
      - plot:
          type: bar
          categories: [A, B, C]
          series: {Baseline: [70, 72, 68], Ours: [74, 77, 71]}
          values: true
          ylabel: Accuracy (%)
      - alertblock: {title: Caveat, text: "Single run; see appendix for variance."}
        animate: fade

  - title: Pipeline
    content:
      - flow:
          nodes:
            - {id: data, text: Raw data, shape: data}
            - {id: prep, text: Preprocess}
            - {id: ok, text: "Quality OK?", shape: decision}
            - {id: model, text: Train model}
          edges: [data -> prep -> ok, "ok -> model: yes", "ok -> prep: no"]
      - steps: [Collect, Clean, Train, Evaluate]
        highlight: 3
```

| Slide `layout` | Use |
|---|---|
| `title` | Title page from `meta` |
| `outline` | Table of contents from all `section` slides |
| `section` | Section divider |
| `frame` (default) | Normal slide with `title` + `content` |
| `plain` | No title bar or footline (e.g. "Thank you") |

Content items: `bullets`, `numbered`, `text`, `block`, `alertblock`, `exampleblock`, `math`, `code`,
`image`, `placeholder`, `table`, `chart`, `plot`, `flow`, `steps`, `columns`, `spacer`.
Inline: `**bold**` · `*italic*` · `` `code` `` · `==alert==` · `[link](url)` · `$math$`.

## Draft-first workflow

A deck can be built before any figure exists:

1. **Diagrams:** draw flowcharts and pipelines right away with `flow:` / `steps:`. For other
   figures use `placeholder:` with a `label`, a `note` describing what to draw, and the final
   `aspect`. The layout is final from the start.
2. **Graphs:** use `plot:` (or `chart:`) with sample numbers and a `caption: "Sample data - replace"`.
3. Build, review the flow, and iterate on the YAML only.
4. When the real figures arrive, change `placeholder:` to `image: {path: ...}` and put the real
   numbers into the plots, then rebuild. Or swap boxes in PowerPoint via **Shape Format → Shape Fill → Picture**.

## Good to know

- **Fitting text.** ppt-templated estimates text size to lay out and shrink content. It doesn't measure
  fonts exactly, so check dense slides; it prints a warning when content doesn't fit.
- **Equations** need pandoc when you *generate* the deck, not to open it. Without pandoc they become
  readable Unicode text (`∫ e^(−x²) dx = √π`). LibreOffice and Google Slides show that text
  fallback; PowerPoint shows the editable equation.
- **Images** must be PNG, JPG, GIF, BMP or TIFF. Export SVG/PDF figures to PNG at 200 dpi or more.
- **Slide numbers** are a live field, so they stay correct when you add or reorder slides.
- **Footline** text (author, title, date) lives in the slide master: View → Slide Master.
- **Regenerating** overwrites the output file. Keep manual edits in a copy, or move them into the YAML.

## Project layout

```
ppt_templated.py      the generator
plots.py              draws `plot:` figures (matplotlib, academic defaults)
flows.py              lays out `flow:` / `steps:` diagrams and routes their connectors
build_template.py     builds the sample templates; shows how to build your own in code
templates/            six sample templates (beamer_oxford is the default)
examples/             example.yaml → example_deck.pptx: every feature once
transformer/          The Transformer lecture (YAML + deck)
segmentation/         Image segmentation methods (YAML + deck; placeholders and sample plots)
YAML_GUIDE.md         complete content-file reference
TEMPLATE_GUIDE.md     how to make or adapt a template
```

## Maintaining the docs

**[YAML_GUIDE.md](YAML_GUIDE.md) must be updated every time `ppt_templated.py`, `plots.py` or `flows.py` changes.**
If a change adds, renames or removes a YAML key, content type, default, chart/plot type, flow shape, color name,
transition or animation effect, update the guide (and the "at a glance" part of this README if it
is affected) in the same commit. The guide is the single source of truth for the YAML format.
