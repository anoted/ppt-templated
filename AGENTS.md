# AGENTS.md — writing decks with ppt-templated

You write **content** in a YAML file; ppt-templated works out everything else: positions, sizes,
fonts, colours, backgrounds, column widths, which things go side by side, where a long slide is
continued, text fitting, chart styling, light/dark adaptation. Spend your effort on what the slides
say. This page is enough to write a deck; every key is in [YAML_GUIDE.md](YAML_GUIDE.md), and
`--explain` prints the keys of any item without opening it.

## 1. The loop

```bash
python ppt_templated.py --templates                    # pick a look (once)
python ppt_templated.py --new talk.yaml -t oxford      # starter file; never overwrites
python ppt_templated.py --explain                      # every layout and item, one line each
python ppt_templated.py --explain table                # keys, accepted values and an example of one
python ppt_templated.py talk.yaml --check --json       # fast: validate + lay out, JSON report
python ppt_templated.py talk.yaml --preview --json     # build talk.pptx, render, measure the layout
```

Repeat *edit → `--check --json` → fix* until `warnings` is empty, then build once with `--preview`
and fix any `overlap-measured` warning. With `--json`, stdout is one JSON object (human text goes to
stderr):

```json
{"ok": true, "output": "talk.pptx", "dark": false, "slides": 12,
 "style": {"fonts": {"headings": "Constantia", "body": "Cambria", "code": "Consolas"},
           "accent": "#002147", "density": "normal"},
 "warnings": [{"slide": 5, "code": "text-shrunk",
               "message": "text shrunk to 80% to fit - consider less content or two slides",
               "fix": "cut text, or let it continue on a second slide (overflow: split is the default)"}],
 "slide_info": [{"slide": 5, "layout": "frame", "title": "Results", "items": ["image", "bullets"],
                 "fill": 0.97, "text_scale": 0.8, "arranged": "columns", "animation_steps": 0}]}
```

| What you see | Meaning | Do |
|---|---|---|
| `"ok": false` + `error` | The build stopped (bad YAML, unknown template, missing file) | Fix what the message names |
| a `warnings` entry | A problem on that slide; `fix` says what to change | Fix it; never ignore |
| `text_scale` < 0.9 | Text was shrunk to fit | Cut text or split the slide |
| `fill` < 0.35 on a frame | Slide looks empty | Merge with a neighbour, or add the figure/number it needs |
| `"arranged": "columns"` | A figure and its text were put side by side for you | Nothing; `arrange: stack` on the slide undoes it |
| a title ending `(1/2)` | The slide was too long and continues on the next | Nothing, or cut content to get one slide back |

Warning codes (`code` is stable, the `message` wording may change):

| `code` | Means | Do |
|---|---|---|
| `text-shrunk` | Text below 90% of its size | Cut text or move part to another slide |
| `content-overflow` | Does not fit even at 65% | Split the slide by hand, or cut |
| `overlap-measured` | Overlap or spill found in the *rendered* slide | Shorten the text named, or split; it is not an estimate |
| `title-long` | Title wraps even when shrunk | Shorten it |
| `unknown-key` · `bad-value` | A key or value the item does not have | Use the suggestion; `--explain ITEM` lists them |
| `unknown-layout` · `misplaced-item` · `bad-slide` · `bad-option` | The slide is not written the way §4 shows | Follow the message |
| `word-too-wide` | A word is wider than its card or timeline entry | Shorter word or fewer entries per row |
| `diagram-small` | A `flow`/`steps` diagram had to use tiny text | Its own slide, or fewer / shorter nodes |
| `image-low-resolution` · `svg-renderer` | Blurry picture · no SVG converter | Larger image or SVG · install a renderer |
| `font-missing` | The template's open font is not installed here | `python fonts.py --install`, or another template |
| `other` | Anything else | The message says what to do |

Exit codes: 0 = built (add `--strict` to fail on warnings), 1 = error.

## 2. Decisions you do not make

Leave these out; the defaults are chosen for you and adapt to every template:

- **No sizes, widths, heights or positions** (`size`, `width`, `height`, `widths`, `valign`, `spacer`).
  Use them only when a warning tells you to.
- **No columns for a figure and its text.** Write the figure and the text as two items, in the order
  you want them read; they are placed side by side (figure first = figure left) whenever that makes
  the figure larger. Use `columns` only for other arrangements (two figures, three columns).
- **No manual slide splitting.** Content that is too long continues on `Title (2/3)` slides: lists and
  tables with repeated headers and running numbers, cards and timelines evenly, a stack of several
  items at item boundaries. Still aim for one message per slide: the split is a safety net.
- **No colours except roles**: `structure`, `alert`, `example`, `muted`. Never `"#RRGGBB"` on an item:
  a fixed colour does not follow the template and can vanish on a dark one.
- **No fonts, backgrounds, logos per slide, footers**: they live in the template / `meta`.
- **No hand-made data arrays when a file exists**: `data: results.csv` on a `table`, `chart` or `plot`.
- **No sampled points for a formula**: `plot: {type: function, f: "1/(1+exp(-x))", x: [-6, 6]}`.

## 3. Choosing a look (once per deck)

| Occasion | `template:` | Fonts |
|---|---|---|
| Academic talk, lecture (default) | `oxford` · `classic` · `crimson` · `blue` | Constantia/Cambria · Cambria · Cambria/Calibri · Calibri |
| Journal club, reading group, paper talk | `paper` · `palatino` · `stix`\* · `article`\* | Georgia/Times · Palatino · STIX Two · Latin Modern Roman |
| Looks like LaTeX Beamer | `latex`\* · `metropolis`\* · `charcoal` · `berlin` · `berkeley` | Latin Modern Sans · Fira Sans · Calibri · Calibri · Cambria/Calibri |
| Long talk where the audience should see the structure | `berlin` (section strip) · `berkeley` (section sidebar) | — |
| Humanities, social science, storytelling | `garamond` · `palatino` · `parchment` (paper texture) | Garamond · Palatino · Georgia/Cambria |
| Teaching, textbook, workshop | `schoolbook` · `chalkboard` (dark) · `notebook` (graph paper) | Century Schoolbook · Cambria/Calibri · Calibri |
| Life sciences, medicine, engineering | `franklin` · `gill` · `stix`\* | Franklin Gothic · Gill Sans · STIX Two |
| Conference, product, business, light | `modern` · `minimal` · `nordic` · `source`\* | Calibri · Arial · Segoe UI · Source Serif/Sans |
| Software, systems, tooling, demos | `console`\* · `solarized`\* · `terminal`\* (dark) · `graphite` (dark) | IBM Plex Mono/Sans · Source Code Pro/Sans · JetBrains Mono · Calibri |
| Dark room, keynote style | `midnight` · `graphite` · `aurora` · `terminal`\* | Calibri · Calibri · Calibri · JetBrains Mono |

\* open fonts: run `python fonts.py --install` once (or `--install latex` for one template). Until
then the build reports `font-missing` and PowerPoint shows a substitute.

The same content file works with every template; switching is one line (`-t name` to try one).
Three `meta` keys adjust a template without leaving it. Each is one word; use them instead of
per-item settings:

| `meta` key | Values | Effect |
|---|---|---|
| `fonts` | `calibri` `cambria` `arial` `times` `georgia` `palatino` `garamond` `schoolbook` `gill` `franklin` `segoe` · open: `latex` `latex-serif` `stix` `source` `source-sans` `fira` `plex` `plex-mono` `mono` | Headings, body and code font of the whole deck |
| `accent` | `navy` `blue` `teal` `green` `forest` `olive` `gold` `orange` `red` `crimson` `maroon` `magenta` `purple` `indigo` `slate` `black`, or `"#RRGGBB"` | The structure colour; adjusted until it is readable on the template |
| `density` | `compact` · `normal` · `airy` | Body text 12% smaller / larger |

Also in `meta`: `arrange: stack` (never put a figure beside its text) and `overflow: shrink` (shrink
instead of continuing on another slide); both can be set on one slide too.

## 4. What to write for what you have

| You have | Write |
|---|---|
| Points to make | `bullets` (nested list = sub-points), `numbered` for sequences |
| One key message | `layout: statement` with `text:` (sized to fill the slide) |
| A definition / warning / example | `block` / `alertblock` / `exampleblock` |
| 2–4 headline numbers | `stats` |
| 3–8 parallel ideas (features, options) | `cards` (auto grid) |
| Events in order | `timeline` |
| A process / decision logic | `flow` (`edges: [A -> B -> C]`), `steps` for a chevron row |
| A citation-worthy sentence | `quote: {text, by}` |
| Tabular results | `table` (or `data: file.csv`) |
| Numbers the audience may edit | `chart` |
| Publication-style figure | `plot` (bar, box, line, scatter, hist, heatmap, confusion, function) |
| An existing figure | `image: fig.svg` (SVG/PDF stay vector) — `images: [...]` for a row |
| A figure that does not exist yet | `placeholder: {label, note, aspect}` |
| An equation | `math: '...'` (single quotes); inline `$...$` |
| Code | `code: \|` + lines (the template's code font); inline `` `code` `` |
| Figure + explanation | the figure, then `bullets` (or text / a block): arranged for you |
| Sources | `layout: references` with `items:` (paginates itself) |

A slide is a mapping with `title:` and `content:` (a list of items). Other slides are
`layout: title | section | outline | statement | references | closing`; `section` slides also feed
the outline and the section navigation of `berlin` / `berkeley`.

```yaml
meta: {title: My Talk, author: Jane Doe, date: today, template: palatino}
slides:
  - layout: title
  - layout: section
    title: Method
  - title: The model learns from pairs          # the title is the message
    content:
      - image: assets/model.svg                  # figure first: it goes on the left
      - bullets: [Two encoders, One contrastive loss, No labels needed]
  - layout: closing
```

## 5. Slide rules that keep warnings away

1. One message per slide; make the **title** that message.
2. At most ~6 bullets of ≤ 2 lines each. More continues on a second slide by itself, but a tighter
   slide is better.
3. One figure per slide, with its text as a second item. Two figures → `images` or `columns`.
4. A row of 2–3 blocks (in `columns`) for comparisons; `alertblock` for the single takeaway.
5. Draw your own SVGs with a transparent background and the colours `--inspect` prints, so they work
   on dark templates too.
6. Monospace and wide fonts (`terminal`, `console`, `mono`, `schoolbook`) hold fewer words per line:
   write shorter bullets, or set `density: compact`.
7. End with `layout: references` (if you cite) and `layout: closing`.

## 6. YAML pitfalls (the usual causes of `not valid YAML`)

- Quote text with `: ` or ` #`, or starting with `*`, `-`, `[`, `{`, `&`, `!`, `%`, `@`, `` ` ``.
- LaTeX in **single quotes**: `'$\frac{a}{b}$'` (in double quotes `\t` becomes a tab).
- Quote `"yes"`, `"no"`, `"on"`, `"off"` and numbers used as names (`"2025": [...]`).
- One item per `- ` entry; settings go beside the item key or inside its mapping.

## 7. Safety

`plot: {type: python, code: ...}` runs Python from the YAML. It only runs with `--allow-python`;
pass that flag only for files you wrote yourself. Everything else (`function` formulas included)
never executes code. `python fonts.py --install` downloads font files from their projects'
repositories (pinned, checksum-verified) and installs them for the current user; nothing is
downloaded unless you run it.
