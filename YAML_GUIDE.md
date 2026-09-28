# ppt-templated YAML Guide

The complete reference for ppt-templated content files (`*.yaml`). Every key a content file can
contain is listed here, with every value it accepts and its default, so you can write a deck
**without reading `ppt_templated.py`, `plots.py` or `flows.py`**.

> **Maintenance rule: update this guide every time `ppt_templated.py`, `plots.py` or `flows.py` changes.**
> If you add, rename or remove a key, a value, a content type, a default, a chart/plot type, a flow
> shape, a transition or an animation effect, change this file in the same commit. This guide is the
> single source of truth for the YAML format; the README only has a summary. Last synced with:
> `ppt_templated.py`, `plots.py` and `flows.py` as of the change that added `flow:` / `steps:` (after commit `7e27310`).

Related docs: [README.md](README.md) (overview, setup) · [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) (making templates).

**How to read the tables:** *Default* `–` means "off / not shown". Values separated by `|` are the
complete list of accepted values. Anything not listed is either an error or (where the table says
so) silently ignored.

---

## Contents

1. [Build command](#1-build-command)
2. [File skeleton](#2-file-skeleton)
3. [`meta`: deck settings](#3-meta-deck-settings) (+ [templates](#templates))
4. [Slide types](#4-slide-types)
5. [How a frame picks its layout](#5-how-a-frame-picks-its-layout)
6. [Content items](#6-content-items): [bullets](#61-bullets--numbered), [text](#62-text), [blocks](#63-block-alertblock-exampleblock), [math](#64-math-display-equation), [code](#65-code), [image](#66-image), [placeholder](#67-placeholder-diagram-to-be-made-later), [table](#68-table), [chart](#69-chart-native-editable-powerpoint-chart), [plot](#610-plot-matplotlib-figure), [columns](#611-columns), [flow](#612-flow-flowcharts-and-workflow-diagrams), [steps](#613-steps-a-row-of-workflow-steps), [spacer](#614-spacer)
7. [Plots in detail](#7-plots-in-detail)
8. [Flows in detail](#8-flows-in-detail): shapes, layout rules, sizing, connectors
9. [Inline formatting](#9-inline-formatting)
10. [Colors](#10-colors)
11. [Animation, transitions, Morph](#11-animation-transitions-morph)
12. [Sizing and fitting](#12-sizing-and-fitting)
13. [YAML gotchas](#13-yaml-gotchas)
14. [Recipes](#14-recipes)
15. [Errors and warnings](#15-errors-and-warnings)

---

## 1. Build command

```bash
python ppt_templated.py talk.yaml                                # output: meta.output, else talk.pptx next to the YAML
python ppt_templated.py talk.yaml -o out.pptx                    # explicit output path
python ppt_templated.py talk.yaml -t templates/beamer_blue.pptx  # template override (beats meta.template)
python ppt_templated.py talk.yaml --keep-template-slides         # keep the template's own preview slides
```

`ppt-templated.py` (with a hyphen) is the same command. From Python:
`from ppt_templated import build_deck; build_deck("talk.yaml", "talk.pptx", template=None, keep_template_slides=False)`.

- Paths inside the YAML (`template`, `image.path`, `output`) are **relative to the YAML file**.
- `template` is also tried relative to the ppt-templated folder, so `templates/beamer_blue.pptx` works from anywhere.
- `plot:` needs `matplotlib` + `numpy`. `flow:` / `steps:` need nothing extra. `math` / `$...$` become native
  equations only if pandoc is installed (`pip install pypandoc_binary`); otherwise they fall back to Unicode text.
- Rebuilding **overwrites** the output file.

## 2. File skeleton

```yaml
meta:
  title: My Talk
  author: Jane Doe
  date: today
  template: ../templates/beamer_blue.pptx

slides:
  - layout: title
  - layout: outline
  - layout: section
    title: Introduction
  - title: First frame            # layout: frame is the default
    content:
      - bullets:
          - Point one
          - - sub-point           # nested list = children of the previous item
          - Point two
  - layout: plain
    content:
      - text: "Thank you!"
        size: 44
        align: center
```

Two top-level keys only: `meta` (mapping) and `slides` (list of mappings). Other top-level keys are
ignored. A slide that is not a mapping is skipped with a warning.

## 3. `meta`: deck settings

| Key | Default | Accepted values / meaning |
|---|---|---|
| `title` | `""` | Deck title (title slide, `{title}` footline token, file properties) |
| `subtitle` | – | Title-slide subtitle, `{subtitle}` token |
| `author` | – | Title slide, `{author}` token, file properties |
| `institute` | – | Title slide, `{institute}` token |
| `date` | today | Any string; `today` → e.g. `September 26, 2026` |
| `short_title`, `short_author`, `short_institute` | the long versions | Footline tokens `{short_title}` etc. |
| `template` | `templates/beamer_oxford.pptx` | Path to a `.pptx` or `.potx`. `-t` overrides it |
| `output` | `<yaml name>.pptx` | Output path (relative to the YAML). `-o` overrides it |
| `transition` | none | Default slide transition for every slide, see [§11](#11-animation-transitions-morph) |
| `valign` | `center` | Vertical placement of frame content: `center` \| `middle` (same) \| `top` |
| `section_style` | `page` | `page` = Section Header slide; `outline` = outline slide with the current section bold and the others muted |
| `outline_title` | `Outline` | Title of the section slides when `section_style: outline` |
| `layouts` | standard names | Map ppt-templated layout keys to your template's layout names (below) |

`layouts` keys: `title`, `section`, `frame`, `columns`, `title_only`, `plain`. Each value is a
layout name or a list of names to try in order:

```yaml
meta:
  layouts: {title: "Cover", frame: "Content", columns: "Comparison",
            title_only: "Heading Only", section: "Divider", plain: "Empty"}
```

Defaults: Title Slide, Section Header, Title and Content, Two Content, Title Only, Blank. If a name
is missing, ppt-templated falls back to the layout type; if even that is missing, the build stops
and lists the template's layouts.

### Templates

Any `.pptx`/`.potx` works (see [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md)). The `templates/` folder has
six samples, for example:

| Template | Look |
|---|---|
| `templates/beamer_oxford.pptx` | **default**: navy title bar, gold rule, serif text |
| `templates/beamer_blue.pptx` | Beamer's default blue |
| `templates/beamer_crimson.pptx` | dark red |

The others (`beamer_charcoal`, `beamer_classic`, `beamer_paper`) are in the same folder.

## 4. Slide types

Every slide is a mapping; `layout` defaults to `frame`, and any unknown `layout` value is treated as
`frame`. Keys valid on **every** slide:

| Key | Meaning |
|---|---|
| `notes` | Speaker notes (string; use `|` for multi-line) |
| `transition` | Overrides `meta.transition` for this slide ([§11](#11-animation-transitions-morph)) |

| `layout` | Template layout | Keys |
|---|---|---|
| `title` | Title Slide | Uses `meta`. `title`, `subtitle`, `author`, `institute`, `date` can be overridden on the slide |
| `outline` | Title and Content | `title` (default `Outline`). Numbered list of every `section` title in the deck (like `\tableofcontents`) |
| `section` | Section Header (or an outline slide, see `meta.section_style`) | `title`, `subtitle` (ignored with `section_style: outline`) |
| `frame` | chosen automatically, see [§5](#5-how-a-frame-picks-its-layout) | `title`, `subtitle` (smaller second line in the title bar), `content` (list of items), `valign` |
| `plain` | Blank (no title bar or footline) | `content`, `valign`. Content fills the slide with a 0.6 in margin |

`valign` on a slide: `center` (default, or `meta.valign`) \| `middle` \| `top`. On the bullet-only
routes of [§5](#5-how-a-frame-picks-its-layout) it sets the placeholder anchor and also accepts `bottom`.

Title slide notes: if the Title Slide layout has text placeholders named **Author**, **Institute**,
**Date**, those fields go there; otherwise they are added as extra lines under the subtitle. Long
titles (more than 2 estimated lines) shrink to 28 pt.

Frame titles that would wrap are shrunk automatically (like Beamer).

## 5. How a frame picks its layout

1. `content` is **one** `bullets`/`numbered` item **without** `size` → the real body placeholder of
   *Title and Content* (bullets inherit master styles; overflow uses PowerPoint's shrink-on-overflow).
2. `content` is **one** `columns` item with exactly **two** columns, each holding one
   `bullets`/`numbered` item, and **no** column `width` → *Two Content* placeholders.
3. Anything else → *Title Only*, with native shapes stacked top to bottom in the content area
   (the rectangle of the *Title and Content* body placeholder).

Tip: add `size:` to a lone bullet list, or a `width:` to a column, to force route 3.

## 6. Content items

`content` is a list. Each item is a mapping identified by **one** of these keys:
`bullets`, `numbered`, `text`, `block`, `alertblock`, `exampleblock`, `math`, `code`, `image`,
`placeholder`, `table`, `chart`, `plot`, `flow`, `steps`, `columns`, `spacer`. Use one per item.

Shorthands: a bare string item is `text`; a bare list item is `bullets`.

```yaml
content:
  - "A plain paragraph"          # = text
  - [One, Two, [two-a, two-b]]   # = bullets
```

**Keys every item accepts** (next to the item key, not inside it):

| Key | Meaning |
|---|---|
| `animate` | Entrance effect, see [§11](#11-animation-transitions-morph). Not on `columns`/`spacer` |
| `name` | Object name for Morph ([§11](#11-animation-transitions-morph)). Not on `columns`/`spacer` |
| `build` | Step-by-step reveal; only on `bullets`, `numbered`, `flow`, `steps` |

Items stack vertically with a 0.25 in gap.

### 6.1 `bullets` / `numbered`

```yaml
- bullets:
    - First point
    - - sub-point a            # nested list = sub-items of "First point"
      - sub-point b
      - - level 3
    - text: Second point       # dict form: text + items
      items: [child 1, child 2]
  build: true                  # reveal one top-level item per click (children come with it)
  size: 18                     # level-1 size in pt; deeper levels scale proportionally
  color: muted                 # text color, see §10
```

| Key | Default | Accepted values |
|---|---|---|
| list entries | – | a string (an empty entry is an empty line), a nested list (children of the previous entry), or `{text: ..., items: [...]}` |
| `build` | – | `true` (= `fade`) or an effect: `appear` \| `fade` \| `wipe` \| `fly` \| `fly-left` \| `zoom`. One click per top-level entry |
| `size` | master level sizes | Level-1 font size (pt); levels 2–4 keep the master's proportions |
| `color` | master color | Any text color ([§10](#10-colors)) |

`numbered` takes the same keys; top-level entries get `1. 2. 3.`, sub-entries keep bullets.

### 6.2 `text`

```yaml
- text: "First paragraph.\nSecond paragraph."   # \n (double-quoted) = new paragraph
  size: 16
  align: center
  color: muted
  bold: true
  italic: true
```

| Key | Default | Accepted values |
|---|---|---|
| `text` | – | String; `\n` starts a new paragraph |
| `size` | master level-1 size | pt |
| `align` | left | `left` \| `center` \| `right` \| `justify` |
| `color` | master color | Any text color ([§10](#10-colors)) |
| `bold`, `italic` | false | `true` \| `false` |

No bullet.

### 6.3 `block`, `alertblock`, `exampleblock`

Beamer blocks: colored header bar + tinted body, grouped into one shape.

```yaml
- block:                     # accent 1 (structure)
    title: Definition
    text: "Body text; **inline formatting** works."
- alertblock:                # accent 2 (alert)
    title: Warning
    bullets: [one, two, [sub]]     # bullets instead of text
- exampleblock:              # accent 3 (example)
    title: Example
    text: "..."
    color: accent4           # optional: override the color
    size: 16                 # optional body size (pt); header is 5% larger
  animate: fade              # note: animate/name sit beside the block key, not inside it
- block: "Short form: just body text, no title"
```

| Key (inside the block) | Default | Accepted values |
|---|---|---|
| `title` | none (empty bar) | Header text |
| `text` | – | Body text (`\n` = new paragraph) |
| `bullets` | – | Body as a bullet list (same forms as §6.1); wins over `text` |
| `color` | `accent1` / `accent2` / `accent3` | Any color ([§10](#10-colors)); the body uses an 85% lighter tint |
| `size` | 90% of master level 1 | Body font size (pt) |

### 6.4 `math` (display equation)

```yaml
- math: '\operatorname{softmax}(z)_i = \frac{e^{z_i}}{\sum_j e^{z_j}}'
  size: 22           # optional (pt), default master level 1; equations are drawn 5% larger
```

LaTeX, centered. Use **single quotes** so backslashes need no escaping (in double quotes write
`\\frac`). Supports `\begin{cases}`, `pmatrix`, `\operatorname`, `\mathbb`, `\text`, etc. (whatever
pandoc converts to OMML). Inline math: `$...$` in any text ([§9](#9-inline-formatting)).

### 6.5 `code`

```yaml
- code: |
    def f(x):
        return x ** 2
  size: 14           # optional, default 70% of master level 1
```

Monospace (Courier New) on a light rounded box. Indentation is preserved; no syntax highlighting.

### 6.6 `image`

```yaml
- image:
    path: assets/fig.png     # relative to the YAML
    width: 0.7
    height: 3.0
    caption: "Figure 1: ..."
    alt: "Alt text"
- image: assets/fig.png      # short form
```

| Key | Default | Accepted values |
|---|---|---|
| `path` | required | PNG, JPG, GIF, BMP, TIFF (not SVG/PDF: export to PNG at 200+ dpi) |
| `width` | 0.75 | Fraction of the available width |
| `height` | from the aspect ratio | Preferred height (in); shrinks to fit |
| `caption` | – | Italic, muted, centered below |
| `alt` | file name | Accessibility description and shape name |

The aspect ratio is read from the file.

### 6.7 `placeholder` (diagram to be made later)

A dashed grey box that reserves exactly the space a figure will take. No file needed.

```yaml
- placeholder: Architecture diagram          # short form: label only, 16:9
- placeholder:
    label: Encoder layer
    note: "attention -> FFN, with residuals"
    aspect: "4:3"
    width: 0.6
```

| Key | Default | Accepted values |
|---|---|---|
| `label` | `Figure` | Bold text in the box |
| `note` | – | Smaller second line: what the diagram should show |
| `aspect` | `16:9` | `"W:H"`, `"W/H"` or a number (width / height) |
| `width` | 0.75 | Fraction of the available width |
| `height` | from `aspect` | Preferred height (in) |
| `caption` | – | As for images |
| `alt` | `Placeholder - <label>` | Alt text |

Replacing it later: change `placeholder:` to `image: {path: ..., width: ...}` (or to a `flow:`) and
rebuild, or in PowerPoint select the box → **Shape Format → Shape Fill → Picture** (stretches to the
box, so match `aspect` to the final figure).

### 6.8 `table`

```yaml
- table:
    header: [Model, Params, Acc]
    rows:
      - [Base, 65M, "81.2%"]
      - ["**Big**", 213M, "==86.4%=="]    # inline formatting and $math$ work in cells
    style: booktabs
    align: [l, r, r]
    widths: [2, 1, 1]
    width: 0.8
    size: 16
```

| Key | Default | Accepted values |
|---|---|---|
| `header` | – | List of header cells (optional) |
| `rows` | – | List of rows (lists of cells). Short rows are padded; empty cells may be `null` |
| `style` | `booktabs` | `booktabs` (top/mid/bottom rules) \| `striped` (accent-1 header with white text, tinted alternate rows) \| `grid` (all borders, tinted header). Any other value draws no rules |
| `align` | auto | Per column: `l` \| `c` \| `r`. Missing entries are `l`. Auto: first column left, numeric columns right, others centered |
| `widths` | fit to text | Relative column widths |
| `width` | fit to text (55–100% of the area) | Table width as a fraction of the area |
| `size` | 78% of master level 1 | Font size (pt) |

Native, editable PowerPoint table.

### 6.9 `chart` (native, editable PowerPoint chart)

```yaml
- chart:
    type: column
    categories: [Q1, Q2, Q3, Q4]
    series:
      "2025": [1.2, 2.4, 3.1, 2.8]
      "2026": [2.0, 2.9, 3.5, 4.1]
    title: Revenue
    y_title: EUR m
    labels: true
    number_format: '#,##0.0'
```

| Key | Default | Accepted values |
|---|---|---|
| `type` | `column` | `column` \| `column-stacked` \| `column-stacked100` \| `bar` \| `bar-stacked` \| `bar-stacked100` \| `line` (with markers) \| `line-plain` \| `area` \| `area-stacked` \| `pie` \| `doughnut` \| `scatter` \| `scatter-lines` \| `radar` |
| `categories` | – | Category labels (not for scatter) |
| `series` | – | `{name: [values]}`, or a list `[{name: A, values: [...]}]`. Scatter: `{name: [[x, y], ...]}` (list form uses `points:`) |
| `stacked` | false | `true` turns `column`, `bar`, `area` into their stacked variant |
| `title` | none | Chart title |
| `legend` | `bottom` if >1 series or pie/doughnut, else `none` | `bottom` \| `top` \| `right` \| `left` \| `none` |
| `labels` | false | Data labels (outside end for bars, above for lines, centered when stacked or doughnut) |
| `percent` | false | Pie/doughnut labels show percentages (needs `labels: true`) |
| `number_format` | – | Excel format for values/axis/labels: `'0%'`, `'#,##0'`, `'0.0'` |
| `y_title`, `x_title` | – | Axis titles |
| `y_min`, `y_max` | auto | Value-axis limits |
| `smooth` | false | Smoothed lines (line / scatter / radar) |
| `gap_width` | 70 | Gap between bar groups (%) |
| `width` | 1.0 | Fraction of the available width |
| `height` | 4.2 | Preferred height (in); shrinks to fit (min 2.2) |

Series colors are accent 1–6 in order. Edit data in PowerPoint via right-click → **Edit Data**.

**`chart` or `plot`?** Use `chart` when the audience may edit the numbers in PowerPoint. Use `plot`
for box plots, confusion matrices, error bars/bands, trend lines, log axes, or publication styling.

### 6.10 `plot` (matplotlib figure)

See [§7](#7-plots-in-detail). Embedded as a picture (300 dpi by default) drawn at its exact slide
size, in the template's body font and accent colors. Regenerate from YAML to change it.

### 6.11 `columns`

```yaml
- columns:
    - width: 0.6              # fraction of the width; columns without width share the rest
      valign: top
      content:
        - bullets: [...]
    - content:
        - placeholder: Diagram
    - [ {text: "short form: a bare list is the column's content"} ]
```

| Key (per column) | Default | Accepted values |
|---|---|---|
| `content` | – | List of items, exactly like a slide's `content` |
| `width` | equal share of the rest | Fraction of the total width |
| `valign` | `center` | `center` \| `middle` \| `top` |

Columns are separated by a 0.4 in gap. Items inside a column stack like frame content. Columns can
hold any item, including blocks, plots, tables, flows and nested `columns`.

### 6.12 `flow` (flowcharts and workflow diagrams)

A diagram of native PowerPoint shapes joined by **connectors glued to the shapes**: drag a shape in
PowerPoint and its arrows follow. Positions are computed for you (see [§8](#8-flows-in-detail));
colors and fonts come from the template.

```yaml
- flow:
    direction: right                # right | down | left | up
    nodes:
      - {id: raw,   text: "Raw data", shape: data}
      - {id: clean, text: "Clean + **normalize**"}
      - {id: ok,    text: "Quality OK?", shape: decision}
      - {id: model, text: "Train model $f_\\theta$", color: alert}
      - {id: store, text: "Model registry", shape: database}
    edges:
      - raw -> clean -> ok          # a chain makes several edges
      - ok -> model: "yes"          # label (quote yes/no, see §13)
      - ok -> clean: "no"           # a loop back is routed around the outside
      - model -> store
    groups:
      - {label: Preprocessing, nodes: [clean, ok]}
    caption: "Figure 2: the pipeline"
  build: true                       # optional: one click per step along the flow
```

The smallest flow is a list of nodes, which are chained in order:

```yaml
- flow: {nodes: [Tokenize, Embed, Attention, FFN, Softmax]}
```

or just edges, whose ids become the node texts:

```yaml
- flow: {edges: [Load -> Parse -> Validate, Validate -> Load]}
```

**Flow settings** (inside `flow:`):

| Key | Default | Accepted values |
|---|---|---|
| `nodes` | – | List of nodes (forms below). Needed unless `edges` names them all |
| `edges` | chain of `nodes` in order | List of edges (forms below). `edges: []` = no arrows. An id not in `nodes` creates a node with that text |
| `groups` | – | List of dashed boxes around nodes: `{label, nodes, color, fill}` (below) |
| `direction` | `right` | `right` (aliases `lr`, `horizontal`) \| `down` (`td`, `tb`, `vertical`) \| `left` (`rl`) \| `up` (`bt`) |
| `shape` | `box` | Default node shape ([§8.1](#81-node-shapes)) |
| `color` | `structure` | Default node color ([§10](#10-colors)) |
| `style` | `tinted` | Default node style: `tinted` (light fill, colored border) \| `solid` (colored fill, white text) \| `outline` (background fill, colored border) \| `plain` (text only) |
| `uniform` | true | Nodes of the same shape get the same size (the largest), like a hand-drawn flowchart. `false` = each node fits its own text |
| `wrap` | 18 | Maximum characters per line in a node before it wraps |
| `gap` | 1.0 | Spacing multiplier (both along and across the flow) |
| `size` | 72% of master level 1 | Node font size (pt) before fitting; labels are 80% of it |
| `line_color` | `[tx1, 0.3]` | Default arrow color |
| `line_width` | 1.25 | Default arrow width (pt) |
| `width` | 1.0 | Maximum width as a fraction of the available width |
| `height` | natural | Maximum preferred height (in) |
| `caption` | – | As for images |
| `alt` | "Flowchart: " + node texts | Alt text of the grouped diagram |

**Nodes**, any of these forms:

```yaml
nodes:
  - Tokenize                                   # text only: the id is the text
  - {id: q, text: "Converged?", shape: decision, color: alert}
  - q2: "Converged?"                            # shorthand  id: text
  - q3: {text: "Converged?", shape: decision}   # shorthand  id: {settings}
```

| Node key | Default | Accepted values |
|---|---|---|
| `id` | the text | Name used in `edges` and `groups` (must be unique) |
| `text` | the id | Node text; inline formatting and `$math$` work; `\n` forces a line break |
| `shape` | flow `shape` | See [§8.1](#81-node-shapes) |
| `color` | flow `color` | Any color ([§10](#10-colors)): fill tint, border, and solid fill |
| `style` | flow `style` (`plain` for `shape: text`) | `tinted` \| `solid` \| `outline` \| `plain` |
| `text_color` | `tx1` (white when `solid`) | Any text color |
| `bold` | false | `true` \| `false` |
| `step` | computed | Position along the flow, 1 = first column (row for `down`) |
| `lane` | computed | Position across the flow, 1 = first row (column for `down`); fractions like `1.5` work |

**Edges**, any of these forms:

```yaml
edges:
  - a -> b                                  # arrow
  - a -> b -> c                             # chain: a -> b and b -> c
  - a --> b                                 # dashed arrow
  - a <-> b                                 # arrows at both ends
  - a -- b                                  # line, no arrow
  - a -> b: "label"                         # with a label
  - a -> b: {label: retry, style: dotted, color: alert, exit: bottom, enter: left}
  - {from: a, to: b, label: retry}          # explicit form
```

| Edge key | Default | Accepted values |
|---|---|---|
| `label` | – | Text next to the arrow (inline formatting works); placed where it hits no shape, line, group border or other label |
| `style` | `solid` (`dashed` for `-->`) | `solid` \| `dashed` \| `dotted` |
| `arrow` | `end` (`both` for `<->`, `none` for `--`) | `end` \| `start` \| `both` \| `none` |
| `color` | flow `line_color` | Any color ([§10](#10-colors)); also colors the label |
| `width` | flow `line_width` | Line width (pt) |
| `exit` | automatic | Side of the source shape: `top` \| `bottom` \| `left` \| `right` |
| `enter` | automatic | Side of the target shape: `top` \| `bottom` \| `left` \| `right` |
| `from`, `to` | – | Node ids (explicit form only) |

A node cannot connect to itself.

**Groups:**

| Group key | Default | Accepted values |
|---|---|---|
| `nodes` | required | List of node ids to enclose |
| `label` | – | Bold small text in the top-left corner |
| `color` | `tx1` (grey) | Any color: dashed border and light fill |
| `fill` | true | `false` = border only |

`build: true` (or an effect name, [§11](#11-animation-transitions-morph)) reveals the diagram one
**step along the flow** per click: all nodes in a column (row for `down`) appear together, an
arrow and its label appear with the later of its two nodes, and a group appears with its first member. Without `build`, the
whole diagram is one group shape, so `animate` and `name` (Morph) act on all of it.

### 6.13 `steps` (a row of workflow steps)

A process row: chevrons by default, or any flow shape joined by arrows.

```yaml
- steps: [Collect, Clean, Train, Evaluate, Deploy]   # settings go next to the list
  highlight: 3                                        # "we are here": step 3 solid and bold
  build: true

- steps:                                              # or: settings inside, list under items
    items: [Plan, Build, {text: Measure, color: alert}, Learn]
    shape: box
```

| Key | Default | Accepted values |
|---|---|---|
| `items` | – | The steps (dict form only; in the short form the list itself). Each is a string or `{text, color, style, text_color, bold, shape}` |
| `shape` | `chevron` | `chevron` (first step flat-backed, chevrons touch) or any flow shape ([§8.1](#81-node-shapes)): then it is a linear `flow` with arrows |
| `highlight` | – | Step number or list of numbers (1 = first): drawn `solid` and bold |
| `color` | `structure` | Any color ([§10](#10-colors)) |
| `style` | `tinted` | `tinted` \| `solid` \| `outline` \| `plain` |
| `wrap` | 14 (chevrons), 18 (other shapes) | Maximum characters per line |
| `direction`, `gap`, `uniform` | as for `flow` | Only when `shape` is not `chevron` |
| `size`, `width`, `height`, `caption`, `alt` | as for `flow` | |

All chevrons have the same width. `build: true` reveals one step per click.

### 6.14 `spacer`

```yaml
- spacer: 0.3        # blank vertical space in inches
```

## 7. Plots in detail

```yaml
- plot:
    type: bar                  # bar | box | line | scatter | confusion
    ...type-specific data...
    ...common options...
    colors: {...}              # optional per-component colors
```

### 7.1 Common options (all types)

| Key | Default | Accepted values |
|---|---|---|
| `type` | required | `bar` \| `box` \| `line` \| `scatter` \| `confusion` |
| `title`, `xlabel`, `ylabel` | – | Text; `$...$` for math (single-quote the string) |
| `legend` | `best` if 2+ named series, else none | `best` \| `top` \| `bottom` \| `right` (these three: outside the axes) \| any matplotlib location (`upper left`, `lower right`, `center`, …) \| `none` \| `false` |
| `grid` | `y` for bar/box (`x` when horizontal), `both` for line/scatter | `x` \| `y` \| `both` \| `none` |
| `xlim`, `ylim` | auto | `[lo, hi]` |
| `xscale`, `yscale` | `linear` | `linear` \| `log` (and other matplotlib scales: `symlog`, `logit`) |
| `xformat`, `yformat` | auto | Tick format: `"{x:.0%}"`, `"{x:,.0f}"`, `"{x:.1f}"` |
| `xticks_rotation` | 0 | Degrees, for long category names |
| `width` | 1.0 | Fraction of the available width |
| `height` | 4.2 | Preferred height (in), shrinks to fit (min 2.2) |
| `aspect` | fill the area | Fixed ratio like `"4:3"` or `"1:1"` |
| `font` | template body font | Font family (falls back to Cambria / Times New Roman / STIX / DejaVu Serif) |
| `font_size` | 66% of master level 1 | Text size (pt) |
| `math` | `font` | Math font: `font` (same as text) \| `stix` \| `cm` \| `stixsans` \| `dejavuserif` \| `dejavusans` |
| `caption`, `alt` | – | As for images |
| `dpi` | 300 | Raster resolution |
| `colors` | theme | See [§7.3](#73-plot-colors) |

### 7.2 Per-type data and options

**`bar`**

```yaml
- plot:
    type: bar
    categories: [BLEU, ROUGE-L]
    series: {Baseline: [27.3, 41.2], Ours: [28.4, 43.9]}
    errors: {Baseline: [0.4, 0.6], Ours: [0.3, 0.5]}
    values: true
```

| Key | Default | Accepted values |
|---|---|---|
| `categories` | required | Category labels |
| `series` | required (or `values`/`y`) | `{name: [one value per category]}` |
| `values` / `y` | – | One-series shorthand: a list instead of `series` (unnamed, no legend) |
| `errors` | – | `{series name: [one error per category]}` → error bars |
| `values: true` | false | Value labels on the bars (see note) |
| `value_format` | `"{:g}"` | Python format for value labels, e.g. `"{:.1f}"`, `"{:.0%}"` |
| `horizontal` | false | Horizontal bars (first category on top) |
| `stacked` | false | Stack the series |
| `bar_width` | 0.8 | Width of a category's bar group (0–1) |
| `hatch` | false | Patterns per series, for grayscale print |

Note: `values` is both "value labels" (`true`) and the one-series data shorthand (a list). To get
labels on a single series, use `series: {Name: [...]}` plus `values: true`.

**`box`**

```yaml
- plot:
    type: box
    groups: {A: [0.61, 0.64, 0.59], B: [0.70, 0.72, 0.74]}
    points: true
    means: true
```

| Key | Default | Accepted values |
|---|---|---|
| `groups` | required (or `stats`) | `{name: [raw samples]}` |
| `stats` | – | Summary statistics instead of samples: `{name: {q1, q3, med` or `median, min` or `whislo, max` or `whishi, mean (optional), fliers (optional list)}}` |
| `points` | false | Overlay the raw samples (jittered; raw samples only) |
| `means` | false | Show mean markers |
| `notch` | false | Notched boxes (raw samples only) |
| `whis` | 1.5 | Whisker length in IQRs |
| `fliers` | true (false when `points: true`) | Show outliers |
| `horizontal` | false | Horizontal boxes |
| `box_width` | 0.55 | Box width (0–1) |

**`line`**

```yaml
- plot:
    type: line
    x: [0, 10, 20, 30]
    series:
      A: [0.1, 0.5, 0.6, 0.7]
      B: {y: [0.1, 0.4, 0.6, 0.7], err: [0.02, 0.03, 0.02, 0.01]}
      C: {y: [...], lower: [...], upper: [...], x: [...], style: "--", marker: s}
```

| Key | Default | Accepted values |
|---|---|---|
| `x` | 0, 1, 2, … | Shared x values; strings make a categorical axis |
| `series` | required (or `y`/`values`) | `{name: [y values]}` or `{name: {y, x, err, lower, upper, style, marker}}` |
| per-series `y` | required | y values |
| per-series `x` | shared `x` | This series' own x values |
| per-series `err` | – | Symmetric error (one per point) |
| per-series `lower`, `upper` | – | Asymmetric bounds (both needed) |
| per-series `style` | `-` (or cycled with `dashes`) | Matplotlib line style: `-` \| `--` \| `-.` \| `:` |
| per-series `marker` | cycled | Matplotlib marker: `o` \| `s` \| `^` \| `D` \| `v` \| `P` \| `X` \| `*` \| … ; `null` = none |
| `errors` | `band` | `band` (shaded) \| `bar` (error bars) |
| `band_alpha` | 0.18 | Band opacity |
| `markers` | true | Different marker per series; `false` = no markers |
| `dashes` | false | Different dash pattern per series |
| `line_width` | 2.0 | pt |
| `marker_size` | 5.5 | pt |

**`scatter`**

```yaml
- plot:
    type: scatter
    series:
      Group A: [[1, 2.1], [2, 3.9], [3, 6.2]]
      Group B: {x: [1, 2, 3], y: [1, 1.5, 2], size: [20, 40, 80]}
    trend: true
```

| Key | Default | Accepted values |
|---|---|---|
| `series` | required (or `values`/`y` as a list of `[x, y]`) | `{name: [[x, y], ...]}` or `{name: {x: [...], y: [...], size: [...]}}` |
| `trend` | false | Least-squares line per series (straight on log axes too) |
| `trend_label` | true | R² in the legend (named series only) |
| `markers` | false | Different marker shape per series |
| `marker_size` | scales with the text | Marker area (pt²) |
| `alpha` | 0.85 | Marker opacity |

**`confusion`**

```yaml
- plot:
    type: confusion
    labels: [cat, dog, bird]
    matrix: [[50, 3, 2], [4, 45, 1], [2, 2, 40]]   # rows = true class, columns = predicted
    normalize: "true"
    value_format: "{:.0%}"
```

| Key | Default | Accepted values |
|---|---|---|
| `matrix` | required | Square list of lists; rows = true class, columns = predicted |
| `labels` | `0, 1, 2, …` | Class names |
| `normalize` | `none` | `none` \| `"true"` / `rows` / `recall` (rows sum to 1) \| `pred` / `cols` / `precision` (columns sum to 1) \| `all` |
| `values` | true | Numbers in the cells |
| `value_format` | `"{:.0f}"`, or `"{:.2f}"` when normalized | Python format |
| `colorbar` | true | Show the color bar |
| `square` | true | Square cells |

Quote `"true"` for `normalize` (bare `true` is a YAML boolean, which also works). Axis labels
default to *Predicted* / *True*; override with `xlabel` / `ylabel`. Legend and grid do not apply.

### 7.3 Plot colors

Under `colors:`. Each value is a theme role (`structure`, `alert`, `example`, `muted`,
`accent1`–`accent6`, `tx1`, `tx2`, `bg1`, `bg2`), a tint `[accent1, 0.5]` (positive = lighter,
negative = darker), `none` / `transparent`, or any matplotlib color (`"#1F3A5F"`, `black`, `tab:blue`).
Keys that color several things take one color or a list (cycled).

| Type | Keys |
|---|---|
| all | `series` (palette, default accent 1–6), `text`, `title`, `label`, `tick`, `axis`, `grid`, `background`, `figure`, `legend_text`, `legend_frame`, `legend_background` |
| `bar` | `edge`, `error`, `value` |
| `box` | `box` (fill), `box_edge`, `whisker`, `cap`, `median`, `mean`, `flier`, `points` |
| `line` | `marker_face`, `marker_edge`, `band`, `error` |
| `scatter` | `edge`, `trend` |
| `confusion` | `low`, `high` (scale ends; default background → accent 1) or `cmap` (matplotlib colormap name), `text_light`, `text_dark`, `cell_edge` (`none` = no gaps), `axis`, `tick` |

```yaml
colors: {box: [accent4, 0.6], box_edge: tx1, median: alert, whisker: muted}
```

## 8. Flows in detail

### 8.1 Node shapes

| `shape` | Aliases | Drawn as | Typical meaning |
|---|---|---|---|
| `box` (default) | `rounded` | Rounded rectangle | Process step |
| `rect` | `process`, `rectangle`, `square` | Rectangle | Process step (classic flowchart) |
| `terminal` | `start`, `end`, `pill`, `stadium` | Flowchart terminator | Start / end |
| `decision` | `diamond`, `if` | Diamond | Yes/no question |
| `data` | `io`, `input`, `output`, `parallelogram` | Parallelogram | Input / output |
| `document` | `doc` | Document (wavy bottom) | A report or file |
| `documents` | `docs` | Multi-document | Several files |
| `database` | `db`, `cylinder` | Cylinder (magnetic disk) | Database / storage |
| `storage` | `stored` | Stored data | Stored data |
| `subprocess` | `predefined` | Predefined process | A step defined elsewhere |
| `manual` | `manual_input` | Manual input | Manual entry |
| `preparation` | `hexagon` | Hexagon | Setup / preparation |
| `delay` | `wait` | Delay (D shape) | Waiting |
| `circle` | `node` | Circle (always round) | Connector / state |
| `ellipse` | `oval` | Ellipse | State |
| `text` | `label`, `none` | No shape, just text | Annotation |

Shape names are case-insensitive. `chevron` is for `steps` only.

### 8.2 How the layout is computed

- **Steps (columns for `right`).** Each node goes one step after its furthest predecessor; arrows that
  close a loop are ignored for this, so loops never push nodes forward.
- **Lanes (rows for `right`).** A node takes the lane of its (median) predecessor, so **the first
  path you list stays on one straight line**; a node that would collide takes the next free lane
  below (to the right for `down`). Sources are placed in the order you list them.
- **Your overrides win:** `step:` and `lane:` on a node fix its position (1 = first); fractional
  lanes (`lane: 1.5`) center a node between two others.
- **Sizes.** Each node is sized from its text (wrapped at `wrap` characters; diamonds and circles
  wrap into a compact block); with `uniform: true` all nodes of one shape share the largest size.
  Gaps grow automatically to make room for edge labels and group boxes.

### 8.3 How arrows are routed

Arrows are orthogonal (horizontal and vertical segments only) and leave and enter shapes at their
**real connection sites**: diamond corners, the slanted sides of `data`, the wavy bottom of
`document`, the top rim of `database`, and so on. Rules, in order of preference:

1. Aligned nodes: a **straight** arrow.
2. A node with several outgoing arrows (a decision): the one to the nearest lane goes straight
   (the first listed on a tie); the others leave from the **side** facing their target and turn once (an L).
3. Nodes in different lanes: a **Z** that bends in the gap just before the target (or right after
   the source, if that path is free).
4. Loops back and arrows whose direct path is blocked go **around the outside** (below/right
   first), with parallel detours spaced apart.
5. `exit:` / `enter:` on an edge force the sides; the path adapts (straight, L, Z, U or S).

Of these candidates, the first that passes clear of every shape it doesn't connect is used. Arrows are real
PowerPoint connectors glued to the shapes: moving a shape in PowerPoint re-routes its arrows, and
**Format → Shape Outline** or right-click → *Connector Types* restyles them.

### 8.4 Sizing

- The diagram is laid out at its node font size (`size`, default 72% of the master level-1 size), then
  **scaled down as one piece** (shapes, gaps, text) to fit the available width (× `width`) and
  height. It is never enlarged, so small diagrams keep readable, consistent text.
- On a crowded slide it is flexible like an image: it shrinks to make room for other items, but
  not below the point where its text would be 9 pt.
- If the text still ends up below ~8.5 pt, ppt-templated warns; split the diagram, give it a
  column of its own, or switch `direction` (`down` suits long chains in a narrow column).

## 9. Inline formatting

Works in any text: bullets, text, block titles and bodies, table cells, frame titles, captions,
placeholder labels, flow nodes, edge and group labels, steps.

| Syntax | Result |
|---|---|
| `**bold**` | bold |
| `*italic*` | italic |
| `` `code` `` | Courier New |
| `==alert==` | alert color (accent 2), like Beamer `\alert` |
| `[text](https://…)` | hyperlink |
| `$x^2$` | inline native equation |
| `\$` | literal dollar sign |

Formats can nest: `**bold with ==alert==**`.

## 10. Colors

One color vocabulary for text colors (`color:` on `bullets`, `text`, `text_color:` on nodes), block
colors, and flow node, edge and group colors. Theme-linked values follow the template:

| Value | Meaning |
|---|---|
| `structure` | accent 1 |
| `alert` | accent 2 |
| `example` | accent 3 |
| `muted` | text color, 45% lighter (grey) |
| `white` | background color |
| `accent1`–`accent6`, `tx1`, `tx2`, `bg1`, `bg2`, `text`, `background` | theme slots (`text` = `tx1`, `background` = `bg1`) |
| `"#FF8800"` | fixed RGB (does not follow the template; quote it) |
| `[accent1, 0.4]` | any of the above, lighter (positive, up to 1) or darker (negative, down to -1) |

Plot colors accept the same names plus matplotlib colors, see [§7.3](#73-plot-colors).

## 11. Animation, transitions, Morph

### Entrance effects (`animate`)

```yaml
- chart: {...}
  animate: fade                                   # short form, on click, 0.5 s
- block: {...}
  animate: {effect: zoom, on: with, duration: 0.8}
```

| Key | Default | Accepted values |
|---|---|---|
| `effect` | `fade` | `appear` \| `fade` \| `wipe` \| `fly` (from bottom) \| `fly-left` \| `zoom`. Unknown names fall back to `fade` |
| `on` | `click` | `click` \| `with` (together with the previous effect) \| `after` (after the previous one) |
| `duration` | 0.5 | Seconds |

`animate: true` = fade on click. Effects play in content order; a caption animates with its figure.

### Step-by-step reveals (`build`)

- `bullets` / `numbered`: `build: true` (fade) or `build: <effect>`: one click per top-level entry;
  sub-entries appear with their parent. Like Beamer `\pause` / `<+->`.
- `flow`: one click per step along the flow ([§6.12](#612-flow-flowcharts-and-workflow-diagrams)).
- `steps`: one click per step.

### Slide transitions (`transition`)

On `meta` (default for all slides) or per slide:

```yaml
transition: fade
transition: {type: push, speed: fast}
```

| Key | Default | Accepted values |
|---|---|---|
| `type` (or the short form) | – | `fade` \| `push` \| `wipe` \| `cover` \| `dissolve` \| `split` \| `zoom` \| `morph` \| `none`. Unknown names fall back to `fade` with a warning |
| `speed` | `med` | `slow` \| `med` \| `fast` (`morph` is always slow) |

### Morph

Give an item the same `name:` on two consecutive slides and set `transition: morph` on the second
slide. PowerPoint 2019 / Microsoft 365 animates between them; older versions show a fade. A `flow`
or `steps` item without `build` is one group, so it morphs as a whole.

```yaml
- title: Before
  content:
    - placeholder: {label: Model}
      name: model
- title: After
  transition: morph
  content:
    - columns:
        - - placeholder: {label: Model, width: 0.9}
            name: model
        - - bullets: [It moved!]
```

## 12. Sizing and fitting

- Content is vertically centered (or top-aligned with `valign: top`) in the content area.
- Text items have a fixed height estimated from their length. Figures (`image`, `placeholder`,
  `plot`, `chart`, `flow`, `steps`) are **flexible**: they shrink to make room (images and
  placeholders to about 1.6 in, charts and plots to 2.2 in, flows until their text is 9 pt).
- If the stack still doesn't fit, all text scales down in steps to 65%. Past that, ppt-templated prints
  `warning: content does not fit on slide '<title>'` → split the slide or remove content.
- Text size is estimated, not measured: always look at dense slides.
- Tables with long text cells: the default fit-to-text width can make cells wrap, and wrapped rows
  are taller than estimated, so the table can run into the next item. Set `width: 1.0` (or a
  column `widths`) and `align: [l, l, ...]` so the cells stay on one line.
- Default text sizes come from the template master (level 1–4 sizes and the title size). Blocks use
  90%, tables 78%, flows 72%, code 70%, and chart/plot text 66% of level 1.

## 13. YAML gotchas

- **Quote** strings that start with `*`, `-`, `[`, `{`, `&`, `!`, `%`, `@`, `` ` ``, `"` or `'`, or contain `: ` or ` #`.
  `"**Bold** start"` needs quotes; `Plain: text` needs quotes.
- **Inside `{...}`** also quote text with `?`, `,`, `[`, `]`, `{` or `}`: `{text: "More data?"}`.
- **Math:** single quotes (`'\frac{a}{b}'`). Inside single quotes write `''` for a literal `'`.
  In double quotes every backslash must be doubled (`"$d_{\\text{model}}$"`).
- **New paragraph in `text`:** `"line one\nline two"` (double quotes) or a `|` block.
- **Numbers as labels:** quote them in `series` keys and categories when they must stay text
  (`"2025": [...]`).
- **Booleans:** `yes`, `no`, `on`, `off`, `true`, `false` are booleans in YAML; quote them when you mean
  text. Edge labels are forgiving: `- a -> b: yes` still shows "yes" (and `no` shows "no"), but
  `on`/`off` would show "yes"/"no", so quote labels.
- **Edges are strings:** `- a -> b` is an edge; `- a -> b: label` is an edge with a label (a
  one-key mapping). Node ids containing `->`, `-->`, `<->` or `--` cannot be used in edge strings;
  use `{from: ..., to: ...}` instead.
- **Nested bullets:** a nested list must come right **after** its parent item.
- **`animate`, `build`, `name`** go on the item (next to `block:`, `chart:`, `flow:` …), not inside it.

## 14. Recipes

**Figure left, bullets right**

```yaml
- title: Architecture
  content:
    - columns:
        - width: 0.45
          content:
            - placeholder: {label: Model, note: "encoder -> decoder", aspect: "4:3", width: 0.95}
        - content:
            - bullets: [One, Two, Three]
              build: true
```

**Equation + explanation + highlighted takeaway**

```yaml
- title: The loss
  content:
    - math: '\mathcal{L} = -\sum_t \log p(y_t \mid y_{<t})'
    - text: "Summed over all target positions."
      align: center
      color: muted
    - alertblock: {title: Note, text: "Teacher forcing makes training parallel."}
      animate: fade
```

**Three blocks side by side**

```yaml
- columns:
    - - block: {title: A, text: "..."}
    - - block: {title: B, text: "..."}
    - - block: {title: C, text: "..."}
```

**Classic top-down flowchart next to its explanation**

```yaml
- title: Training loop
  content:
    - columns:
        - width: 0.5
          content:
            - flow:
                direction: down
                nodes:
                  - {id: s, text: Start, shape: terminal}
                  - {id: b, text: "Next batch", shape: data}
                  - {id: q, text: "Converged?", shape: decision}
                  - {id: u, text: "Update weights"}
                  - {id: e, text: Stop, shape: terminal}
                edges: [s -> b -> q, "q -> u: no", "q -> e: yes", u -> b]
              build: true
        - content:
            - bullets: [...]
```

**Roadmap slide that moves along a talk:** repeat a `steps:` row on several slides with a
different `highlight:` each time.

**Pipeline with the stages boxed**

```yaml
- flow:
    edges: [Load -> Clean -> Features -> Train -> Evaluate]
    groups:
      - {label: Data, nodes: [Load, Clean, Features]}
      - {label: Model, nodes: [Train, Evaluate], color: alert}
```

**Sample graph while real data is pending:** use `plot:` with made-up numbers and a
`caption: "Sample data - replace"`; keep diagrams as `placeholder:` until they are drawn (or draw
them now with `flow:`).

**Closing slide**

```yaml
- layout: plain
  content:
    - text: "Thank you!"
      size: 44
      align: center
      color: structure
    - text: "Questions?"
      align: center
      color: muted
```

## 15. Errors and warnings

| Message | Cause / fix |
|---|---|
| `Error on slide N (<title>): ...` | The build stops at the first bad slide; the message names it |
| `content item needs one of bullets, numbered, ...` | Item key misspelled, or `animate`/`name` used without a content key |
| `cannot understand content item` | Item is a number or other non-string/list/mapping |
| `unknown chart type 'x'` | See the chart types in [§6.9](#69-chart-native-editable-powerpoint-chart) |
| `plot type must be one of ...` | See [§7.1](#71-common-options-all-types) |
| `plot type 'x' needs 'series'` / `'groups'` / ... | Missing data key for that plot type |
| `bar series 'A' has 3 values for 4 categories` | Series length must match `categories` |
| `placeholder aspect must look like 16:9, 4/3 or 1.5` | Bad `aspect` value |
| `image not found: <path>` | Paths are relative to the YAML file |
| `unknown colour 'x'` | See [§10](#10-colors) |
| `flow needs settings, e.g. {nodes: [...], ...}` | `flow:` must be a mapping |
| `flow needs 'nodes' (or 'edges' that name them)` | Empty flow |
| `unknown flow shape 'x'` | See [§8.1](#81-node-shapes) |
| `unknown flow node/edge/group key(s) ...` | Misspelled key; the message lists the valid ones |
| `flow edge needs two node ids, like 'a -> b'` | Edge string without a valid arrow (`->`, `-->`, `<->`, `--`) |
| `flow edge 'a -> a' connects a node to itself` | Self-loops are not supported |
| `flow group '...' names unknown node(s)` | Group `nodes` must be node ids |
| `flow node id(s) used twice` | Give each node a unique `id` |
| `steps needs a list of step texts` | `steps:` needs a non-empty list (or `items:`) |
| `warning: flow diagram on slide '...' is drawn with N pt text` | Too big for its space; split it, give it a column, or change `direction` |
| `warning: content does not fit on slide '...'` | Split the slide ([§12](#12-sizing-and-fitting)) |
| `warning: unknown transition 'x', using fade` | See the transition list in [§11](#11-animation-transitions-morph) |
| `warning: slide N is not a mapping - skipped` | A slide entry is a bare string/list; add keys |
