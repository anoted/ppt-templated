# ppt-templated YAML Guide

The complete reference for content files (`*.yaml`): every key, every accepted value, every default.
Write decks from this guide; you should not need to read `ppt_templated.py`, `plots.py`, `flows.py`
or `media.py`. If something is missing here, add it.

> **Maintenance rule:** change this guide in the same commit as any change to `ppt_templated.py`,
> `plots.py`, `flows.py`, `media.py` or `preview.py` that adds, renames or removes a key, value,
> default, item, layout, shape, effect or message. The README only summarizes.

**Reading the tables:** *Default* `–` = off / not shown. `a | b` = the complete list of accepted values.
Unknown keys and values are **reported as warnings** (with a guess such as *did you mean 'color'?*),
never silently ignored.

Related: [README.md](README.md) (overview) · [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) (templates).

---

## Contents

1. [Workflow](#1-workflow) · 2. [Skeleton](#2-skeleton) · 3. [Cheat sheet](#3-cheat-sheet) ·
4. [`meta`](#4-meta) · 5. [Slide layouts](#5-slide-layouts) · 6. [Content items](#6-content-items) ·
7. [Figures: images, SVG, PDF](#7-figures-images-svg-pdf) · 8. [Plots](#8-plots) ·
9. [Flows and steps](#9-flows-and-steps) · 10. [Inline formatting](#10-inline-formatting) ·
11. [Colors](#11-colors) · 12. [Animation, transitions, Morph](#12-animation-transitions-morph) ·
13. [Layout, sizing and good slides](#13-layout-sizing-and-good-slides) · 14. [YAML gotchas](#14-yaml-gotchas) ·
15. [Recipes](#15-recipes) · 16. [Messages](#16-messages)

---

## 1. Workflow

```bash
python ppt_templated.py talk.yaml --preview    # build talk.pptx, render slide images, check the layout
```

| Command | Does |
|---|---|
| `python ppt_templated.py talk.yaml` | Build `talk.pptx` next to the YAML (or `meta.output`). Overwrites it |
| `... --preview [DIR]` | Also render `slide-NN.png` and contact sheets `sheet-N.png` (6 numbered slides per image) into `.preview/<deck>/` (or DIR). With PowerPoint, the rendered layout is measured too (see below) |
| `... --check` | Validate and lay out only; writes nothing (fast) |
| `... --strict` | Exit code 1 if there were warnings |
| `... -o out.pptx` / `-t templates/x.pptx` | Output path / template (beat `meta.output` / `meta.template`) |
| `... --keep-template-slides` | Keep the template's own preview slides |
| `python ppt_templated.py --inspect [TEMPLATE]` | Print a template's layouts, content area, fonts, text sizes, theme colours and footline tokens |
| `python preview.py deck.pptx` | Preview any `.pptx` |
| Python | `from ppt_templated import build_deck; build_deck("talk.yaml", "talk.pptx", template=None, check=False)` |

**The loop:** write the YAML → `--preview` → fix every `warning:` line (each names the slide) → look
at the contact sheets → repeat. Warnings marked `(measured)` come from PowerPoint itself after
rendering: text that runs into another shape, out of its block, or below the content area. The
other size checks are estimates, so always look at the sheets.

- Paths in the YAML (`template`, `output`, image paths) are **relative to the YAML file**; `template`
  is also tried relative to the ppt-templated folder (`templates/beamer_blue.pptx` works anywhere).
- Rendering needs PowerPoint (Windows, exact) or LibreOffice (any OS; close, no measuring).
- `plot:` needs matplotlib + numpy. Equations become native only with pandoc (`pip install pypandoc_binary`);
  without it they are written as Unicode text. SVG fallbacks need Chrome/Edge (or `pip install cairosvg`),
  PDF figures need `pip install pymupdf`.
- Generated files: `.preview/` (images) and `.ppt-cache/` (converted figures) next to the deck. Both are disposable.

## 2. Skeleton

```yaml
meta:
  title: My Talk
  author: Jane Doe
  date: today
  template: templates/beamer_oxford.pptx      # the default; see §4

slides:
  - layout: title
  - layout: section
    title: Introduction
  - title: First frame                         # layout: frame is the default
    content:
      - bullets:
          - Point one
          - - sub-point                        # a nested list = children of the item above
          - Point two
  - layout: references
    items: ['A. Author, "Paper," *Venue*, 2024.']
  - layout: closing
```

Top level: `meta` (mapping) and `slides` (list of mappings). A slide is a mapping; `content` is a list.

## 3. Cheat sheet

**Slides** (`layout:`): `title` · `section` · `outline` · `frame` (default) · `plain` · `references` · `closing`.

**Content items** (one key per list entry; settings go beside the key or inside its mapping):

| Item | Smallest form | Main settings |
|---|---|---|
| `bullets` / `numbered` | `- bullets: [A, B, [b1, b2]]` | `size`, `color`, `build` |
| `text` | `- "A paragraph"` | `size`, `align`, `color`, `bold`, `italic` |
| `block` / `alertblock` / `exampleblock` | `- block: {title: T, text: "..."}` | `bullets`, `color`, `size` |
| `math` | `- math: 'E = mc^2'` | `size` |
| `code` | `- code: \|` + lines | `size` |
| `image` | `- image: assets/fig.svg` | `width`, `height`, `caption`, `alt`, `page`, `vector` |
| `images` | `- images: [a.png, b.svg]` | per figure `caption`; `height`, `gap`, `caption`, `build` |
| `placeholder` | `- placeholder: Diagram` | `note`, `aspect`, `width`, `caption` |
| `table` | `- table: {header: [A, B], rows: [[1, 2]]}` | `style`, `align`, `widths`, `width`, `size`, `caption` |
| `chart` | `- chart: {type: column, categories: [...], series: {...}}` | §6.10 |
| `plot` | `- plot: {type: bar, categories: [...], series: {...}}` | §8 |
| `stats` | `- stats: [{value: "95%", label: accuracy}]` | `size`, `color`, `build` |
| `flow` | `- flow: {edges: [A -> B -> C]}` | §9 |
| `steps` | `- steps: [Plan, Build, Test]` | `highlight`, `shape`, `build` |
| `columns` | `- columns: [[...items...], [...items...]]` | per column `width`, `valign`, `animate` |
| `spacer` | `- spacer: 0.3` | inches |

**Any item:** `animate: fade` (entrance), `name: x` (Morph). **Inline:** `**bold**` `*italic*`
`` `code` `` `==alert==` `[link](url)` `$math$`. **Colors:** `structure` `alert` `example` `muted` `accent1`–`6` `"#RRGGBB"` `[accent1, 0.4]`.

## 4. `meta`

| Key | Default | Meaning |
|---|---|---|
| `title` | `""` | Deck title (title slide, footline `{title}`, file properties) |
| `subtitle`, `author`, `institute` | – | Title slide and footline tokens |
| `date` | today | Any text; `today` → e.g. `October 22, 2026` |
| `short_title`, `short_author`, `short_institute` | the long ones | Footline tokens `{short_title}` etc. |
| `template` | `templates/beamer_oxford.pptx` | `.pptx` or `.potx`; `-t` overrides |
| `output` | `<yaml name>.pptx` | Output path; `-o` overrides |
| `transition` | – | Default transition of every slide (§12) |
| `valign` | `center` | Vertical placement of frame content: `center` \| `middle` \| `top` \| `bottom` |
| `section_style` | `page` | `page` = Section Header slides; `outline` = an outline slide per section, current one bold |
| `outline_title` | `Outline` | Title of those outline slides |
| `layouts` | standard names | Map the keys `title`, `section`, `frame`, `columns`, `title_only`, `plain` to your template's layout names (a name or a list to try in order), see [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) |

**Templates** in `templates/`: `beamer_oxford` (**default**; navy bar, gold stripe, serif) ·
`beamer_paper` (journal look, rules) · `beamer_classic` (Beamer default) · `beamer_modern` (clean sans,
teal accent) · `beamer_blue` (Madrid) · `beamer_charcoal` (metropolis) · `beamer_crimson` (Beaver).

## 5. Slide layouts

Keys on **every** slide: `layout` (default `frame`), `notes` (speaker notes; use `|` for several lines),
`transition` (overrides `meta.transition`).

| `layout` | Keys | Result |
|---|---|---|
| `title` | `title`, `subtitle`, `author`, `institute`, `date` (all default to `meta`) | Title Slide layout. Titles over 2 lines shrink to 28 pt |
| `section` | `title`, `subtitle` | Section Header (or an outline slide with `meta.section_style: outline`) |
| `outline` | `title` (default `Outline`) | Numbered list of all `section` titles |
| `frame` | `title`, `subtitle` (smaller second line), `content`, `valign` | A normal slide. Long titles shrink to fit one line (a warning says when they still wrap) |
| `plain` | `content`, `valign` | No title bar or footline; content in a 0.6 in margin |
| `references` | `items` (list, one entry per reference), `title` (`References`), `size` (62% of body level 1, ≥ 12 pt), `start` (1), `numbered` (true) | Hanging-indent list `[1] ...`. Too many for one slide → continues on more slides titled `References (1/2)`, entries spread evenly. An entry that starts with its own label (`"[12] ..."`) keeps it; numbering continues after it |
| `closing` | `title` (`Thank you!`), `subtitle` (`Questions?`; `""` hides it), `text` (extra lines: contact, link) | End slide on the Blank layout, title in the heading font |

`valign`: `center` (default) \| `middle` \| `top` \| `bottom`.
Item keys written directly on a `frame`/`plain` slide (e.g. `bullets:` without `content:`) are moved
into `content` with a warning.

## 6. Content items

`content` is a list; each entry is one item, identified by its key. Shorthands: a bare string = `text`,
a bare list = `bullets`.

- **Settings may sit inside the item's mapping or beside the item key**; both mean the same:
  `- image: {path: fig.png, width: 0.5}` = `- image: fig.png` + `width: 0.5` on the next line.
- `animate` (entrance, §12), `name` (Morph, §12) and `build` (step-by-step reveal) work on any item
  except `spacer` (`build` only on `bullets`, `numbered`, `images`, `stats`, `flow`, `steps`; `columns`
  takes `animate` only). They may also sit either beside the key or inside the mapping.
- Items stack top to bottom with a 0.25 in gap; see §13 for sizing.

### 6.1 `bullets` / `numbered`

```yaml
- bullets:
    - First point
    - - sub-point a               # nested list = children of "First point"
      - - level 3
    - {text: Second point, items: [child 1, child 2]}   # same as a nested list
  build: true                      # one top-level entry per click (children come with it)
  size: 18
```

| Key | Default | Accepted values |
|---|---|---|
| entries | – | string (empty = blank line), nested list, or `{text, items}` |
| `build` | – | `true` (= `fade`) \| `appear` \| `fade` \| `wipe` \| `fly` \| `fly-left` \| `zoom` |
| `size` | master level sizes | Level-1 size (pt); deeper levels keep the master's proportions |
| `color` | master | Any color (§11) |

`numbered` numbers the top level `1. 2. 3.`; sub-entries keep bullets.

### 6.2 `text`

| Key | Default | Accepted values |
|---|---|---|
| `text` | – | String; `\n` (double quotes) or a `\|` block starts a new paragraph. No bullet |
| `size` | master level 1 | pt |
| `align` | `left` | `left` \| `center` \| `right` \| `justify` |
| `color` | master | Any color |
| `bold`, `italic` | false | `true` \| `false` |

### 6.3 `block`, `alertblock`, `exampleblock`

Beamer blocks: colored header bar and tinted body, one group shape.

```yaml
- alertblock: {title: Caveat, text: "Single run; see the appendix."}
- block:
    title: Definition
    bullets: [one, two]            # bullets instead of text
- exampleblock: "Short form: body text only"
```

| Key | Default | Accepted values |
|---|---|---|
| `title` | empty bar | Header text |
| `text` | – | Body (`\n` = new paragraph) |
| `bullets` | – | Body as a list (forms of §6.1); wins over `text` |
| `color` | accent 1 / 2 / 3 | Any color; the body is an 85% lighter tint |
| `size` | 90% of level 1 | Body size (pt); the header is 5% larger |

**Blocks side by side** (`columns` whose columns each hold the same number of blocks) share one height
per row and line up at the top.

### 6.4 `math` (display equation)

```yaml
- math: '\operatorname{softmax}(z)_i = \frac{e^{z_i}}{\sum_j e^{z_j}}'
  size: 22
```

LaTeX, centered, drawn 5% larger than `size` (default level 1). **Single quotes** (no escaping).
Supports what pandoc turns into OMML: `\frac`, `\sum`, `\begin{cases}`, `pmatrix`, `\mathbb`, `\text`, ...
An equation wider than its column is broken over lines by PowerPoint and may run into the next item;
`--preview` reports it (`measured`). Give it a wider column or a smaller `size`.

### 6.5 `code`

```yaml
- code: |
    def f(x):
        return x ** 2
  size: 14                         # default 70% of level 1
```

Courier New on a light box; indentation kept; no highlighting.

### 6.6 `image`

```yaml
- image: assets/diagram.svg        # short form
- image:
    path: results/fig3.pdf
    page: 1
    width: 0.8
    caption: "Figure 3: ..."
```

| Key | Default | Accepted values |
|---|---|---|
| `path` | required | PNG, JPG, GIF, BMP, TIFF, **SVG**, **PDF**, WebP (anything Pillow reads); see §7 |
| `width` | 0.75 | Fraction of the available width |
| `height` | from the aspect ratio | Preferred height (in); shrinks to fit |
| `caption` | – | Italic, muted, centered below |
| `alt` | file name | Accessibility text and shape name |
| `page` | 1 | Page of a PDF |
| `vector` | true | SVG/PDF: `false` embeds only the PNG rendering |

The aspect ratio always comes from the file; the box is the largest of that ratio that fits.

### 6.7 `images` (a row of figures)

```yaml
- images:
    - {path: assets/input.png, caption: "(a) input"}
    - {path: assets/ours.svg, caption: "(b) ours"}
  caption: "Figure 4: qualitative results"
  build: true                      # one figure (with its caption) per click
```

All figures get **one height** and keep their aspect ratios; the row is centered.

| Key | Default | Accepted values |
|---|---|---|
| entries | required | A path, or `{path, caption, alt, page, vector}` |
| `height` | as tall as fits | Preferred figure height (in) |
| `width` | 1.0 | Fraction of the width the row may use |
| `gap` | 0.25 | Gap between figures (in) |
| `caption` | – | Caption under the row |

### 6.8 `placeholder` (a figure to be made later)

A dashed box that reserves exactly the space of a future figure.

```yaml
- placeholder: Architecture diagram
- placeholder: {label: Encoder, note: "attention -> FFN", aspect: "4:3", width: 0.6}
```

| Key | Default | Accepted values |
|---|---|---|
| `label` | `Figure` | Bold text in the box |
| `note` | – | Smaller second line: what to draw |
| `aspect` | `16:9` | `"W:H"`, `"W/H"` or a number |
| `width`, `height`, `caption`, `alt` | 0.75, from `aspect`, –, `Placeholder - <label>` | As for `image` |

Later: change `placeholder:` to `image:` (same `width`), or in PowerPoint use Shape Format → Shape Fill → Picture.

### 6.9 `table`

```yaml
- table:
    header: [Model, Params, Acc]
    rows:
      - [Base, 65M, "81.2%"]
      - ["**Ours**", 213M, "==86.4%=="]   # inline formatting and $math$ work in cells
    caption: "Table 1: results"
```

| Key | Default | Accepted values |
|---|---|---|
| `header` | – | Header cells |
| `rows` | required | Lists of cells; short rows are padded, `null` = empty. A bare list (`table: [[...]]`) = rows |
| `style` | `booktabs` | `booktabs` (3 rules) \| `striped` (filled header, tinted rows) \| `grid` (all borders) \| `plain` (no rules) |
| `align` | auto | Per column `l` \| `c` \| `r`. Auto: first column left, numbers right, long text (> 30 characters) left, the rest centered |
| `widths` | auto | Relative column widths |
| `width` | fit to text (55–100%) | Fraction of the available width |
| `size` | 78% of level 1 | pt |
| `caption` | – | Below the table |

Auto widths keep short columns on one line and share the rest among long-text columns. Rows grow
for wrapped text; the space is reserved below the table, so it does not run into the next item.

### 6.10 `chart` (native, editable PowerPoint chart)

```yaml
- chart:
    type: column
    categories: [Q1, Q2, Q3, Q4]
    series: {"2025": [1.2, 2.4, 3.1, 2.8], "2026": [2.0, 2.9, 3.5, 4.1]}
    y_title: EUR m
    labels: true
```

| Key | Default | Accepted values |
|---|---|---|
| `type` | `column` | `column` \| `column-stacked` \| `column-stacked100` \| `bar` \| `bar-stacked` \| `bar-stacked100` \| `line` (markers) \| `line-plain` \| `area` \| `area-stacked` \| `pie` \| `doughnut` \| `scatter` \| `scatter-lines` \| `radar` |
| `categories` | – | Labels (not for scatter) |
| `series` | – | `{name: [values]}` or `[{name, values}]`; scatter: `{name: [[x, y], ...]}` (list form: `points`) |
| `stacked` | false | `true` turns `column`, `bar`, `area` into the stacked variant |
| `title` | – | Chart title |
| `legend` | `bottom` if 2+ series or pie, else `none` | `bottom` \| `top` \| `right` \| `left` \| `none` |
| `labels` | false | Data labels |
| `percent` | false | Pie/doughnut labels as % (with `labels: true`) |
| `number_format` | – | Excel format: `'0%'`, `'#,##0'`, `'0.0'` |
| `x_title`, `y_title` | – | Axis titles |
| `y_min`, `y_max` | auto | Value-axis limits |
| `smooth` | false | Smooth lines |
| `gap_width` | 70 | Gap between bar groups (%) |
| `width` | 1.0 | Fraction of the width |
| `height` | 4.2 | Preferred height (in); shrinks to 2.2 |
| `caption` | – | Below the chart |

Series colors are accents 1–6; all text uses the template's body font. Edit data: right-click → Edit Data.
**`chart` or `plot`?** `chart` when the audience edits numbers in PowerPoint; `plot` for box plots,
confusion matrices, error bars/bands, trend lines, log axes, publication styling.

### 6.11 `stats` (key numbers)

```yaml
- stats:
    - {value: "95.4%", label: "PCK@0.15 on a held-out session"}
    - {value: "10×", label: "cheaper than a camera rig", color: alert}
  build: true                      # one number per click
```

| Key | Default | Accepted values |
|---|---|---|
| entries | required | `{value, label, color}` or just the value |
| `size` | 180% of level 1 | Size of the numbers (pt); labels are 80% of level 1, muted |
| `color` | `structure` | Color of the numbers (per entry `color` wins) |

Equal-width cells across the row; 2–4 entries read best.

### 6.12 `columns`

```yaml
- columns:
    - width: 0.45                  # fraction of the width; columns without one share the rest
      valign: top
      content:
        - image: assets/fig.svg
    - - bullets: [One, Two]        # short form: a bare list is the column's content
  animate: fade                    # the whole row appears at once
```

| Key (per column) | Default | Accepted values |
|---|---|---|
| `content` | – | A list of items, as on a slide. A column may also be a bare list or one item |
| `width` | equal share | Fraction of the total width |
| `valign` | `center` | `center` \| `middle` \| `top` |
| `animate` | – | This column appears as one step |

Columns are 0.4 in apart and can hold any item, including nested `columns`.

### 6.13 `flow`, `steps`, `plot`

See [§9](#9-flows-and-steps) and [§8](#8-plots).

### 6.14 `spacer`

`- spacer: 0.3` adds blank vertical space (inches).

## 7. Figures: images, SVG, PDF

Any figure file works; conversion is automatic and sized to the exact box on the slide:

| File | In the deck |
|---|---|
| `.png` `.jpg` `.gif` `.bmp` `.tif` | As is. EXIF rotation applied; photos over 2× the needed 300 dpi and > 1.5 MB are downsampled |
| `.svg` | **Vector** (PowerPoint 2019 / 365 draws the SVG), plus a PNG fallback rendered at the slide size, 300 dpi, for other viewers |
| `.pdf` | One page (`page:`), converted to SVG + PNG the same way (needs PyMuPDF) |
| `.webp` and others | Converted to PNG |
| `.eps`, `.emf`, `.wmf` | Not supported: export as SVG, PDF or PNG |

- Aspect ratio: from the pixels (raster), `width`/`height` or `viewBox` (SVG), or the page (PDF).
- An SVG needs a `viewBox` or a `width` and `height`. Linked images inside it are embedded.
- Fallback PNGs come from headless Chrome/Edge (or cairosvg) and are cached in `.ppt-cache/`.
- `vector: false` embeds only the PNG (for an SVG feature PowerPoint draws badly).
- Raster figures under 100 dpi at their size get a warning; prefer SVG for diagrams.
- **Drawing your own SVGs:** use the template's fonts and colours (`--inspect` prints them), a
  `viewBox` about 100 units per inch of slide, and text ≥ 14 units tall.

## 8. Plots

`plot:` draws a matplotlib figure at its exact slide size in the template's body font and accent
colours, embedded as a vector SVG (with a PNG fallback). Regenerate from YAML to change it.

```yaml
- plot:
    type: bar                      # bar | box | line | scatter | confusion
    categories: [BLEU, ROUGE-L]
    series: {Baseline: [27.3, 41.2], Ours: [28.4, 43.9]}
    errors: {Baseline: [0.4, 0.6], Ours: [0.3, 0.5]}
    values: true
    ylabel: Score
```

### 8.1 Common options

| Key | Default | Accepted values |
|---|---|---|
| `type` | required | `bar` \| `box` \| `line` \| `scatter` \| `confusion` |
| `title`, `xlabel`, `ylabel` | – | Text; `$...$` math (single-quote the string) |
| `legend` | `best` with 2+ named series | `best` \| `top` \| `bottom` \| `right` (outside) \| any matplotlib location (`upper left`, ...) \| `none` |
| `grid` | `y` for bar/box (`x` if horizontal), `both` otherwise | `x` \| `y` \| `both` \| `none` |
| `xlim`, `ylim` | auto | `[lo, hi]` |
| `xscale`, `yscale` | `linear` | `linear` \| `log` \| `symlog` \| `logit` |
| `xformat`, `yformat` | auto | Tick format, e.g. `"{x:.0%}"`, `"{x:,.0f}"` |
| `xticks_rotation` | 0 | Degrees |
| `width` | 1.0 | Fraction of the width |
| `height` | 4.2 | Preferred height (in); shrinks to 2.2 |
| `aspect` | fill the area | Fixed ratio, e.g. `"4:3"` |
| `font` | template body font | Font family (falls back to Cambria, Times New Roman, STIX, DejaVu Serif) |
| `font_size` | 66% of level 1 | pt |
| `math` | `font` | `font` \| `stix` \| `cm` \| `stixsans` \| `dejavuserif` \| `dejavusans` |
| `caption`, `alt` | – | As for images |
| `dpi` | 300 | PNG fallback resolution |
| `vector` | true | `false` = PNG only |
| `colors` | theme | §8.3 |

### 8.2 Per type

**`bar`**: `categories` (required); `series: {name: [one value per category]}` (or `values`/`y`: one
unnamed series); `errors: {name: [...]}` (error bars); `values: true` (value labels; with one series use
`series` + `values: true`); `value_format` (`"{:g}"`, `"{:.1f}"`, `"{:.0%}"`); `horizontal`; `stacked`;
`bar_width` (0.8); `hatch` (patterns for grayscale print).

**`box`**: `groups: {name: [samples]}` or `stats: {name: {q1, q3, med|median, min|whislo, max|whishi, mean, fliers}}`;
`points` (overlay samples), `means`, `notch`, `whis` (1.5), `fliers` (true; false with `points`),
`horizontal`, `box_width` (0.55).

**`line`**: `x` (shared x; strings = categorical); `series: {name: [y]}` or
`{name: {y, x, err, lower, upper, style, marker}}` (`err` symmetric; `lower` + `upper` asymmetric;
`style` `-` \| `--` \| `-.` \| `:`; `marker` `o` `s` `^` `D` `v` ... or `null`); `errors` `band` (default) \| `bar`;
`band_alpha` (0.18); `markers` (true); `dashes` (false); `line_width` (2.0); `marker_size` (5.5).

**`scatter`**: `series: {name: [[x, y], ...]}` or `{name: {x, y, size}}`; `trend` (least-squares line,
straight on log axes); `trend_label` (R² in the legend, true); `markers` (shape per series); `marker_size`; `alpha` (0.85).

**`confusion`**: `matrix` (square; rows = true class, columns = predicted); `labels`;
`normalize` `none` \| `"true"`/`rows`/`recall` \| `pred`/`cols`/`precision` \| `all`; `values` (true);
`value_format` (`"{:.0f}"`, `"{:.2f}"` when normalized); `colorbar` (true); `square` (true).
Axis labels default to *Predicted* / *True*.

### 8.3 Plot colors

Under `colors:`; each value is a theme role (`structure`, `alert`, `example`, `muted`, `accent1`–`6`,
`tx1`, `bg1`, ...), a tint `[accent1, 0.5]`, `none`, or a matplotlib color (`"#1F3A5F"`, `tab:blue`).

| Type | Keys |
|---|---|
| all | `series` (palette; list or one color), `text`, `title`, `label`, `tick`, `axis`, `grid`, `background`, `figure`, `legend_text`, `legend_frame`, `legend_background` |
| `bar` | `edge`, `error`, `value` |
| `box` | `box`, `box_edge`, `whisker`, `cap`, `median`, `mean`, `flier`, `points` |
| `line` | `marker_face`, `marker_edge`, `band`, `error` |
| `scatter` | `edge`, `trend` |
| `confusion` | `low`, `high` (scale ends) or `cmap`, `text_light`, `text_dark`, `cell_edge`, `axis`, `tick` |

## 9. Flows and steps

Native shapes joined by **connectors glued to the shapes** (drag a box in PowerPoint and its arrows
follow), laid out automatically in the template's colors and fonts.

```yaml
- flow:
    direction: right                # right | down | left | up
    nodes:
      - {id: raw, text: "Raw data", shape: data}
      - {id: ok, text: "Quality OK?", shape: decision}
      - {id: model, text: "Train model", color: alert}
    edges:
      - raw -> ok
      - ok -> model: "yes"          # label; quote yes/no
      - ok -> raw: "no"             # a loop back goes around the outside
    groups:
      - {label: Preprocessing, nodes: [raw, ok]}
  build: true                       # one step along the flow per click
- flow: {edges: [Load -> Parse -> Validate]}     # smallest form: ids become the texts
- steps: [Collect, Clean, Train, Evaluate]       # a chevron row
  highlight: 3                                   # "we are here"
```

### 9.1 `flow` settings

| Key | Default | Accepted values |
|---|---|---|
| `nodes` | – | List of nodes (below); needed unless `edges` name them all |
| `edges` | chain of `nodes` | List of edges (below); `[]` = no arrows |
| `groups` | – | Dashed boxes: `{nodes: [ids], label, color, fill}` |
| `direction` | `right` | `right` (`lr`) \| `down` (`td`, `tb`) \| `left` (`rl`) \| `up` (`bt`) |
| `shape` | `box` | Default node shape (§9.3) |
| `color` | `structure` | Default node color |
| `style` | `tinted` | `tinted` \| `solid` \| `outline` \| `plain` |
| `uniform` | true | Same-shape nodes get the same size |
| `wrap` | 18 | Characters per line before a node wraps |
| `gap` | 1.0 | Spacing multiplier |
| `size` | 72% of level 1 | Node text (pt) before fitting; labels 80% of it |
| `line_color`, `line_width` | `[tx1, 0.3]`, 1.25 | Arrow color and width (pt) |
| `width`, `height` | 1.0, natural | Max fraction of the width; max height (in) |
| `caption`, `alt` | – | As for images |

**Nodes:** `Text` · `{id, text, shape, color, style, text_color, bold, step, lane}` · `id: text` ·
`id: {settings}`. `step` / `lane` pin the position (1 = first; `lane: 1.5` centers between two).

**Edges:** `a -> b` · `a -> b -> c` (chain) · `a --> b` (dashed) · `a <-> b` · `a -- b` (no arrow) ·
`a -> b: label` · `a -> b: {label, style, color, width, arrow, exit, enter}` · `{from: a, to: b, ...}`.
`style` `solid` \| `dashed` \| `dotted`; `arrow` `end` \| `start` \| `both` \| `none`; `exit`/`enter`
`top` \| `bottom` \| `left` \| `right`. A node cannot connect to itself; an id containing `->` needs the `{from, to}` form.

### 9.2 `steps` settings

`items` (the list; in the short form the list itself; each a text or `{text, color, style, text_color, bold, shape}`),
`highlight` (number or list: drawn solid and bold), `shape` (`chevron`, or any flow shape → boxes with arrows),
`color`, `style`, `wrap` (14 for chevrons, 18 otherwise), `direction`/`gap`/`uniform` (non-chevron only),
`size`, `width`, `height`, `caption`, `alt`. `build: true` = one step per click.

### 9.3 Node shapes

| `shape` | Aliases | Drawn as |
|---|---|---|
| `box` (default) | `rounded` | Rounded rectangle (process) |
| `rect` | `process`, `rectangle`, `square` | Rectangle |
| `terminal` | `start`, `end`, `pill`, `stadium` | Start / end |
| `decision` | `diamond`, `if` | Diamond |
| `data` | `io`, `input`, `output`, `parallelogram` | Parallelogram |
| `document` / `documents` | `doc` / `docs` | Document(s) |
| `database` | `db`, `cylinder` | Cylinder |
| `storage` | `stored` | Stored data |
| `subprocess` | `predefined` | Predefined process |
| `manual` | `manual_input` | Manual input |
| `preparation` | `hexagon` | Hexagon |
| `delay` | `wait` | Delay |
| `circle` / `ellipse` | `node` / `oval` | Circle / ellipse |
| `text` | `label`, `none` | Text only |

### 9.4 Layout rules

- Each node sits one step after its furthest predecessor (loops are ignored for this); a node keeps
  its predecessor's lane, so **the first path you list is a straight line**.
- Arrows are orthogonal and use real connection sites: straight if aligned; a decision's second
  branch leaves from the side; otherwise a Z bend; loops back go around the outside.
- The diagram is scaled down as one piece to fit (never enlarged), but not below 9 pt text; below
  ~8.5 pt you get a warning: split it, give it its own column, or use `direction: down`.

## 10. Inline formatting

In any text (bullets, text, block titles and bodies, table cells, titles, captions, labels, nodes):

| Syntax | Result |
|---|---|
| `**bold**` · `*italic*` · `` `code` `` | bold · italic · Courier New |
| `==alert==` | alert color (accent 2), like `\alert` |
| `[text](https://...)` | hyperlink |
| `$x^2$` | inline native equation |
| `\$` | a literal dollar sign |

Formats nest: `**bold with ==alert==**`.

## 11. Colors

One vocabulary everywhere (`color`, `text_color`, block, flow, edge and group colors):

| Value | Meaning |
|---|---|
| `structure` · `alert` · `example` | accent 1 · accent 2 · accent 3 |
| `muted` | text color, 45% lighter (grey) |
| `white` / `background` / `bg1` | background color |
| `text` / `tx1`, `tx2`, `bg2`, `accent1`–`accent6` | theme slots |
| `"#FF8800"` | fixed RGB (does not follow the template; quote it) |
| `[accent1, 0.4]` | lighter (0 to 1) or darker (0 to -1) |

Theme-linked colors follow the template, so `-t` restyles the whole deck. Plots also accept matplotlib colors (§8.3).

## 12. Animation, transitions, Morph

**Entrance (`animate`):** `animate: fade` or `animate: {effect: zoom, on: with, duration: 0.8}`.

| Key | Default | Accepted values |
|---|---|---|
| `effect` | `fade` | `appear` \| `fade` \| `wipe` \| `fly` (from below) \| `fly-left` \| `zoom` |
| `on` | `click` | `click` \| `with` (with the previous) \| `after` (after the previous) |
| `duration` | 0.5 | Seconds |

`animate: true` = fade on click. Captions animate with their figure. On `columns` (or one column) the whole row (column) appears as one step.

**Step-by-step (`build`):** `true` or an effect name. `bullets`/`numbered`: one top-level entry per
click (like `\pause`); `images`, `stats`, `steps`: one entry per click; `flow`: one step along the flow per click.

**Transitions (`transition`)** on `meta` or a slide: `fade` or `{type: push, speed: fast}`.
`type`: `fade` \| `push` \| `wipe` \| `cover` \| `dissolve` \| `split` \| `zoom` \| `morph` \| `none`;
`speed`: `slow` \| `med` \| `fast` (morph is always slow).

**Morph:** give an item the same `name:` on two consecutive slides and set `transition: morph` on the
second. PowerPoint 2019+ animates the change; older versions fade. A `flow`/`steps`/`images`/`stats`
item without `build` is one group, so it morphs as a whole.

## 13. Layout, sizing and good slides

**How a frame is built:**
1. One `bullets`/`numbered` item without `size` → the layout's real body placeholder (overflow shrinks it, with a warning).
2. One `columns` item with two columns of one bullet list each, no `width` → *Two Content* placeholders.
3. Anything else → *Title Only* plus native shapes stacked in the content area (the *Title and Content* body rectangle).

**Fitting:** text items take their estimated height; figures (`image`, `images`, `placeholder`,
`plot`, `chart`, `flow`, `steps`) are flexible and shrink first. If the stack still does not fit, all
text shrinks in steps to 65% (warning below 90%); past that, *content does not fit*. Text heights are
estimated from the template fonts' widths, not measured: use `--preview`, which measures.

**Default sizes** (from the template master; `--inspect` prints them): body level 1 (e.g. 22 pt),
blocks 90%, tables 78%, flows 72%, code 70%, captions/charts/plots 66%, references 62%.

**Good slides (check the contact sheets for these):**
- One message per slide; make the title that message.
- At most ~6 bullets of ≤ 2 lines; split instead of shrinking. Nothing below 14 pt except captions and references (≥ 12).
- Figure + explanation: `columns` with the figure at `width: 0.45`–`0.6`.
- Key results: `stats` or a short `table`; trends: `plot`; editable numbers: `chart`.
- A row of blocks for comparisons (they align automatically); `alertblock` for the one takeaway.
- Diagrams: `flow`/`steps` when they are boxes and arrows, an SVG otherwise, `placeholder` until drawn.
- End with `layout: references` and `layout: closing`.
- No warnings, no `(measured)` overlaps, nothing tiny, no half-empty slides.

## 14. YAML gotchas

- **Quote** text that starts with `*`, `-`, `[`, `{`, `&`, `!`, `%`, `@`, `` ` ``, `"` or `'`, or contains
  `: ` or ` #`: `"**Bold** start"`, `"Note: this"`. Inside `{...}` also quote text with `,` `?` `[` `]` `{` `}`.
- **Math in single quotes**: `'\frac{a}{b}'`. In double quotes every backslash must be doubled, and
  `"\t..."` silently becomes a TAB (`"$\times$"` breaks; `'$\times$'` works). In single quotes, `''` is a quote.
- **New paragraph:** `"one\ntwo"` (double quotes) or a `|` block.
- **Numbers and booleans as text:** quote them: `"2025": [...]`, `"yes"`, `"no"`, `"on"`, `"off"`.
- **Nested bullets:** the nested list comes right after its parent.
- **Edges are strings:** `- a -> b` is an edge; `- a -> b: label` is an edge with a label.
- **One item per `- `:** `{text: A, image: b.png}` is an error; write two items.

## 15. Recipes

**Figure left, bullets right**

```yaml
- title: Our method in one picture
  content:
    - columns:
        - width: 0.5
          content: [{image: assets/method.svg}]
        - - bullets: [Encoder reads the window, Decoder predicts 33 joints, Trained end to end]
            build: true
```

**Result slide: table + takeaway**

```yaml
- title: "Main result: 4 points better on all datasets"
  content:
    - table: {header: [Model, Acc], rows: [[Baseline, "81.2%"], ["**Ours**", "==85.4%=="]]}
    - alertblock: {title: Takeaway, text: "Same compute, better accuracy."}
      animate: fade
```

**Comparison row** (equal heights)

```yaml
- columns:
    - - block: {title: Cameras, bullets: [rich detail, privacy issues]}
    - - alertblock: {title: Wearables, bullets: [no video, must be worn]}
    - - exampleblock: {title: "Wi-Fi (ours)", bullets: [device-free, cheap]}
```

**Equation with its explanation**

```yaml
- math: '\mathcal{L} = -\sum_t \log p(y_t \mid y_{<t})'
- text: "Summed over all target positions."
  align: center
  color: muted
```

**Classic flowchart beside its explanation**

```yaml
- columns:
    - width: 0.5
      content:
        - flow:
            direction: down
            nodes:
              - {id: s, text: Start, shape: terminal}
              - {id: q, text: "Converged?", shape: decision}
              - {id: u, text: "Update weights"}
              - {id: e, text: Stop, shape: terminal}
            edges: [s -> q, "q -> u: no", "q -> e: yes", u -> q]
          build: true
    - - bullets: [...]
```

**Roadmap across a talk:** repeat a `steps:` row on several slides with a different `highlight:`.

**Draft before figures exist:** `placeholder:` with `label`, `note` and the final `aspect`; `plot:`
with sample numbers and `caption: "Sample data - replace"`. Swap in the real files later.

## 16. Messages

Errors stop the build and name the slide (`error: slide 7 "Results": ...`); warnings never stop it.

| Message | Fix |
|---|---|
| `unknown key 'x' (did you mean 'y'?) - ignored` | Misspelt or misplaced key; use the suggestion |
| `x: 'v' is not one of ...` | Use one of the listed values |
| `unknown layout 'x' - built as a frame` | Use a layout of §5 |
| `'bullets' written on the slide itself ... moved there` | Put items in the list under `content:` |
| `content item {...} has no content key` | Each list entry needs one item key (§3) |
| `one item has two content keys` | Split into two items, each starting with `- ` |
| `<kind> needs settings, e.g. ...` | Give the item a mapping (or list) as shown |
| `build works on ... only - use animate` | `build` is for lists, rows and diagrams |
| `text shrunk to N% to fit` / `bullets shrunk to N%` | Too much on the slide: cut or split |
| `content does not fit even at 65%` | Split the slide |
| `title wraps even at N pt` | Shorten the title |
| `x.png is only N dpi at this size` | Use a larger image or an SVG |
| `no SVG renderer ...` | Install Chrome/Edge or `pip install cairosvg` |
| `image not found: <path>` | Paths are relative to the YAML |
| `x.svg: the SVG needs a viewBox ...` | Add `viewBox="0 0 W H"` to the `<svg>` tag |
| `PDF figures need PyMuPDF` | `pip install pymupdf`, or export SVG/PNG |
| `flow diagram drawn with N pt text` | Give it more room, split it, or change `direction` |
| `... runs into ... (measured)` | Real overlap found by PowerPoint: shorten, resize, or move to a new slide |
| `... runs N pt below the content area (measured)` | Content spills into the footline: cut or split |
| `text runs N pt out of the bottom of ... (measured)` | A block's text is too long for it |
| `unknown chart type` · `plot type must be one of ...` · `unknown flow shape` | See §6.10, §8, §9.3 |
| `bar series 'A' has 3 values for 4 categories` | Match `categories` |
| `flow edge needs two node ids, like 'a -> b'` · `flow node id(s) used twice` | Fix the edge string / ids |
| `unknown flow node/edge/group key(s)` | The message lists the valid keys |
| `unknown colour 'x'` | §11 |
| `not valid YAML` | See §14; the message gives the line and column |
| `cannot write x.pptx - is it open in PowerPoint?` | Close it, or use `-o` |
| `template not found` | Path relative to the YAML or the ppt-templated folder |
