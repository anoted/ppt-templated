# AGENTS.md — writing decks with ppt-templated

You write **content** in a YAML file; ppt-templated works out everything else: positions, sizes,
fonts, colours, backgrounds, column widths, text fitting, chart styling, light/dark adaptation.
Spend your effort on what the slides say. This page is the short version; every key is in
[YAML_GUIDE.md](YAML_GUIDE.md).

## 1. The loop

```bash
python ppt_templated.py --templates                    # pick a look (once)
python ppt_templated.py --new talk.yaml -t oxford      # starter file; never overwrites
python ppt_templated.py talk.yaml --check --json       # fast: validate + lay out, JSON report
python ppt_templated.py talk.yaml --preview --json     # build talk.pptx, render, measure the layout
```

Repeat *edit → `--check --json` → fix* until `warnings` is empty, then build once with `--preview`
and fix any `(measured)` warning. With `--json`, stdout is one JSON object (human text goes to stderr):

```json
{"ok": true, "output": "talk.pptx", "dark": false, "slides": 12,
 "warnings": [{"slide": 5, "message": "text shrunk to 80% to fit - consider less content or two slides"}],
 "slide_info": [{"slide": 5, "layout": "frame", "title": "Results", "items": ["table", "text"],
                 "fill": 0.97, "text_scale": 0.8, "animation_steps": 0}]}
```

| What you see | Meaning | Do |
|---|---|---|
| `"ok": false` + `error` | The build stopped (bad YAML, unknown template, missing file) | Fix what the message names |
| a `warnings` entry | A problem on that slide; the message says what to change | Fix it; never ignore |
| `text_scale` < 0.9 | Text was shrunk to fit | Cut text or split the slide |
| `fill` < 0.35 on a frame | Slide looks empty | Merge with a neighbour, or add the figure/number it needs |
| `(measured)` in a message | Real overlap found in the rendered slides | Shorten or split; it is not an estimate |

Exit codes: 0 = built (add `--strict` to fail on warnings), 1 = error.

## 2. Decisions you do not make

Leave these out; the defaults are chosen for you and adapt to every template:

- **No sizes, widths, heights or positions** (`size`, `width`, `height`, `widths`, `valign`, `spacer`).
  Use them only when a warning tells you to.
- **No colours except roles**: `structure`, `alert`, `example`, `muted`. Never `"#RRGGBB"`: a fixed
  colour does not follow the template and can vanish on a dark one.
- **No fonts, backgrounds, logos per slide, footers**: they live in the template / `meta`.
- **No manual slide splitting for long lists or tables**: set `meta.overflow: split` (the starter
  does) and they continue on `Title (2/3)` slides with repeated headers and running numbers.
- **No hand-made data arrays when a file exists**: `data: results.csv` on a `table`, `chart` or `plot`.
- **No sampled points for a formula**: `plot: {type: function, f: "1/(1+exp(-x))", x: [-6, 6]}`.

## 3. Choosing a template (once per deck)

| Occasion | `template:` |
|---|---|
| Academic talk, lecture (default) | `oxford` · `classic` · `crimson` · `blue` |
| Journal / reading-group style | `paper` |
| Conference, product, business, light | `modern` · `minimal` |
| Dark room, keynote style | `midnight` · `graphite` · `aurora` |
| Teaching, workshop | `chalkboard` (dark) · `notebook` (graph paper) |
| Humanities, storytelling | `parchment` (paper texture) |

The same content file works with every template; switching is one line (`-t name` to try one).

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
| Figure + explanation | `columns: [[figure], [bullets]]` |
| Sources | `layout: references` with `items:` (paginates itself) |

## 5. Slide rules that keep warnings away

1. One message per slide; make the **title** that message.
2. At most ~6 bullets of ≤ 2 lines each. More → `overflow: split` or a second slide.
3. Text plus a visual → `columns` (the visual left, the text right).
4. A row of 2–3 blocks for comparisons; `alertblock` for the single takeaway.
5. Draw your own SVGs with a transparent background and the colours `--inspect` prints, so they work
   on dark templates too.
6. End with `layout: references` (if you cite) and `layout: closing`.

## 6. YAML pitfalls (the usual causes of `not valid YAML`)

- Quote text with `: ` or ` #`, or starting with `*`, `-`, `[`, `{`, `&`, `!`, `%`, `@`, `` ` ``.
- LaTeX in **single quotes**: `'$\frac{a}{b}$'` (in double quotes `\t` becomes a tab).
- Quote `"yes"`, `"no"`, `"on"`, `"off"` and numbers used as names (`"2025": [...]`).
- One item per `- ` entry; settings go beside the item key or inside its mapping.

## 7. Safety

`plot: {type: python, code: ...}` runs Python from the YAML. It only runs with `--allow-python`;
pass that flag only for files you wrote yourself. Everything else (`function` formulas included)
never executes code.
