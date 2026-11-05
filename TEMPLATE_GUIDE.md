# Template Guide

How ppt-templated uses a PowerPoint template, and three ways to get one:

- **A. Build one from a theme file** with `build_template.py` (easiest; no PowerPoint needed).
- **B. Adapt an existing template**, e.g. your company's.
- **C. Restyle a template by hand** in PowerPoint.

> **Maintenance rule:** when `build_template.py` gains or changes a theme key, or `ppt_templated.py`
> changes what it reads from a template, update this guide in the same commit.

Related: [YAML_GUIDE.md](YAML_GUIDE.md) (content files) · [README.md](README.md).

---

## 1. What a template provides

ppt-templated writes only slides; everything else comes from the template, so changing a theme
color in PowerPoint later recolors the whole deck, and slides added by hand match the generated ones.

```
Theme            12 colors + heading/body fonts (+ colour map: light or dark)
 └─ Slide master   background (colour, gradient, texture or picture), frame-title bar, footline, text styles
     └─ Layouts      Title Slide, Section Header, Title and Content, ... (placeholder positions)
         └─ Slides     content only (written by ppt-templated)
```

### Layouts (found by name, then by layout type)

| Used for | Layout name | Needs |
|---|---|---|
| `layout: title` | **Title Slide** | Title, Subtitle; optional text placeholders **named** `Author`, `Institute`, `Date` (else those lines go under the subtitle) |
| `layout: section` | **Section Header** | Title; optional one text placeholder (the subtitle) |
| one bullet list, `outline` | **Title and Content** | Title + one content placeholder. **Its rectangle is the content area of every frame** |
| two bullet columns | **Two Content** | Title + two content placeholders |
| every other frame, `references` | **Title Only** | Title |
| `plain`, `closing` | **Blank** | – |

Other names: map them in the YAML (`meta.layouts: {frame: "Content", title_only: "Heading Only", ...}`)
or rename the layouts. If nothing fits, the build stops and lists the template's layouts.
Only the first slide master is used.

### Theme color roles

| Slot | Role | Used for |
|---|---|---|
| Dark 1 (`dk1`) | text | body and table text, `muted` (moved towards the background) |
| Light 1 (`lt1`) | background | slides, text on colored bars |
| Dark 2 / Light 2 (`dk2`, `lt2`) | free | e.g. frame titles (`tx2`) or a light title bar (`bg2`) |
| **Accent 1** | *structure* | title box, bullets, `block`, flow nodes, series 1 |
| **Accent 2** | *alert* | `alertblock`, `==alert==`, series 2 |
| **Accent 3** | *example* | `exampleblock`, series 3 |
| Accent 4–6 | – | series 4–6 |
| Hyperlink | links | `[text](url)` |

Make Accent 1 the main color (dark enough for white text on it), Accent 2 a warning/highlight
color, Accent 3 a positive one.

**Dark templates** use PowerPoint's own mechanism: the slide master's colour map says *Background 1 =
Dark 1, Text 1 = Light 1* (`<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" .../>`). ppt-templated
reads the colour map, so any dark template works, including your own: it detects a dark template
(background darker than text), makes every tint (block bodies, table stripes, flow nodes, code boxes,
plot shading) towards the dark background instead of towards white, keeps chart and plot backgrounds
transparent, and draws text on solid accent fills in the background colour. On a dark template the
accents must therefore be **light enough to carry dark text** (e.g. sky blue, coral, light green).

### Text styles and sizes

Bullet boxes copy the master's body styles (bullet characters, colors, indents, sizes of levels 1–9).
Item sizes derive from the master's level-1 size: blocks 90%, tables 78%, flows 72%, code 70%,
captions/charts/plots 66%, references 62%. Text heights are estimated from the theme fonts'
character widths (known for the common Office fonts; others are estimated conservatively).

### Footline tokens

Plain text boxes (not placeholders) on the master or layouts may contain tokens, replaced on build:
`{title}` `{short_title}` `{subtitle}` `{author}` `{short_author}` `{institute}` `{short_institute}`
`{date}`. Combine freely: `{author} ({short_institute})`. Page numbers: a slide-number field.

### Check any template

```bash
python ppt_templated.py --inspect my_template.pptx                         # what ppt-templated sees
python ppt_templated.py examples/example.yaml -t my_template.pptx --preview  # every feature, rendered
```

`--inspect` prints, for example:

```
template      templates/beamer_oxford.pptx  (13.33 x 7.50 in)
layouts       title = 'Title Slide', section = 'Section Header', frame = 'Title and Content', ...
title slide   placeholders: Title 1, Subtitle 2, Author, Institute, Date
content area  x 0.60  y 1.30  w 12.13  h 5.60 in (every frame's content goes here)
mode          light
fonts         headings Constantia, body Cambria
text sizes    frame title 28 pt; body levels 22 / 20 / 18 / 16 pt
colours       dk1 #1C1C1C  lt1 #FFFFFF  dk2 #002147  lt2 #F0EEE8
              accent1 #002147 (structure)  accent2 #A6192E (alert)  accent3 #3D6B4F (example) ...
              bg1 #FFFFFF (background)  tx1 #1C1C1C (text)  bg2 #F0EEE8  tx2 #002147
footline      tokens: {author} {date} {short_title}
```

Use its fonts and colors when you draw SVG figures, so they match the slides. Then go through the
checklist in §5 on the contact sheets.

`python ppt_templated.py --templates` lists the bundled templates (name, light/dark, look). The list
comes from each file's properties, set by `build_template.py`: *Title* = label, *Subject* = the
one-line description, *Category* = `light` or `dark`. A template without them is still listed, with
its mode worked out from its colours. A bundled template is used by its short name (`template:
midnight` = `templates/beamer_midnight.pptx`); drop your own `.pptx` into `templates/` to use it by
name too.

---

## 2. Route A: build a template from a theme file

```yaml
# forest.yaml
name: forest                       # -> templates/beamer_forest.pptx
label: Beamer Forest
description: Forest green bar with a gold stripe
colors: {text: "1B1B1B", background: "FFFFFF", text2: "1F3D2B", background2: "EEF3EF", accent1: "2C5F2D",
         accent2: "C8553D", accent3: "3E7CB1", accent4: "E0A458", accent5: "7FA99B",
         accent6: "8C8C8C", hlink: "2C5F2D", folHlink: "1F3D2B"}
fonts: [Cambria, Calibri]          # [headings, body]
frame: bar
bar_stripe: accent4
```

A dark one with a texture:

```yaml
# night.yaml
name: night
dark: true                         # light text on a dark background
colors: {text: "E6E6E6", background: "16181D", text2: "FFFFFF", background2: "22262E",
         accent1: "7CC4FA", accent2: "FF8A80", accent3: "9CCC65", accent4: "FFD54F"}
background: {texture: grain, strength: 0.7}
frame: accent
rule: accent1
footline: text
cover: left
```

```bash
python build_template.py --spec forest.yaml            # or -o path/to/out.pptx
python build_template.py --list                        # built-in themes
python build_template.py --theme modern                # rebuild one built-in theme
```

Unknown keys and bad values stop the build with a message naming the valid ones. All keys are
optional except what you want to change; missing colors come from the Oxford theme.

| Key | Default | Values |
|---|---|---|
| `name` | file name | Output `templates/beamer_<name>.pptx` (theme files only; letters, digits, `-`, `_`) |
| `label` | `Beamer <Name>` | Theme name shown in PowerPoint |
| `description` | – | One line shown by `ppt_templated.py --templates` |
| `dark` | false | `true`: light text on a dark background. Give `text`/`background` (light/dark) yourself; they are stored in the slots and colour map of §1 the way PowerPoint's dark themes do it |
| `colors` | Oxford's | The 12 slots of §1 as hex `"RRGGBB"` (quote them in YAML). `dk1`/`lt1`/`dk2`/`lt2` may be written `text`/`background`/`text2`/`background2`, and always mean those roles, also with `dark: true` |
| `fonts` | `[Calibri, Calibri]` | `[headings, body]` |
| `background` | `bg1` | The slide background: a theme colour (`bg2`, `[bg1, 95, 0]`) · a texture name · `{gradient: [bg1, bg2], angle: 90}` (two or more theme colours, angle in degrees) · `{texture: paper, strength: 1, ink: accent1}` · `{image: picture.jpg}` (relative to the theme file, stretched over the slide) |
| `frame` | `bar` | Frame titles: `bar` (filled bar) \| `rule` (hairline under the title) \| `accent` (light hairline with a short accent bar) \| `plain` |
| `bar`, `bar_text` | `accent1`, `bg1` | Bar fill and title color (`frame: bar`) |
| `bar_stripe` | – | Thin stripe under the bar |
| `title_color` | `bar_text` | Frame-title color without a bar |
| `rule` | `tx1` / `accent1` | Hairline color (`frame: rule`); accent-bar color (`frame: accent`, `cover: left`) |
| `footline` | `boxes` | `boxes` (three colored segments) \| `text` (grey text) |
| `foot`, `foot_text` | `accent1` ×3, `bg1` ×3 | Segment fills and text colors (`boxes`): author, title, date |
| `foot_stripe` | – | Stripe above a boxed footline |
| `foot_rule` | – | Hairline above a text footline |
| `cover` | `box` | Title and section slides: `box` (filled title box) \| `rules` (rules above and below) \| `left` (flush left, title on an accent bar) \| `plain` |
| `cover_text` | `bg1` (box), else `accent1` | Cover title color |
| `radius` | 12000 | Corner of the title box (0 = square; 50000 = round ends) |
| `bullets` | `[[►, 70], [•, 100], [–, 100]]` | `[char, size %]` per level, repeated for deeper levels |
| `bullet_color` | `accent1` | |
| `title_bold`, `subtitle_italic` | false | |

**Color values** in these keys are theme references, so they follow the palette: `accent1`–`6`,
`tx1`, `tx2`, `bg1`, `bg2`, or `[name, lumMod %, lumOff %]`: `[accent1, 75, 0]` is 25% darker,
`[tx1, 50, 50]` a mid grey on a light theme (on a dark theme use `[tx1, 60, 0]`). Prefer `tx1`/`bg1`
to `dk1`/`lt1`: they follow `dark: true`.

**Textures** (`background: {texture: ...}`) are drawn by `textures.py` from the theme's own colours
with a fixed random seed (rebuilding gives the same file): `paper` (warm mottled sheet with fibres),
`grid` (graph paper, 1/4 in), `dots` (dot grid), `lines` (ruled paper with a margin line), `chalk`
(blackboard smudges and dust, for dark themes), `grain` (fine film grain, light or dark). `strength`
scales the effect (0.5 subtler, 2 stronger); `ink` is the colour of lines, dots and fibres (default
`accent1` for grid/dots/lines, else `tx1`). Keep `background` (the colour) close to the texture's
average, since tints and text are chosen for it.

**Built-in themes** (`THEMES` in `build_template.py`, same keys):

| Theme | Look | Fonts |
|---|---|---|
| `oxford` (default) | Navy bar with gold stripes, square title box | Constantia / Cambria |
| `paper` | Hairline rules, grey text footline, journal look | Georgia / Times New Roman |
| `classic` | No bars, blue titles (Beamer default) | Cambria |
| `modern` | White, teal accent bar under titles, left-aligned covers | Calibri Light / Calibri |
| `blue` | Blue bar and three-part footline (Madrid) | Calibri |
| `charcoal` | Dark slate bar, orange alerts (metropolis) | Calibri |
| `crimson` | Light grey bar, dark red titles (Beaver) | Cambria / Calibri |
| `minimal` | Black on white, bold titles, one red accent | Arial |
| `parchment` | Paper texture, sepia ink, rules | Georgia / Cambria |
| `notebook` | Graph-paper texture, ink-blue titles, red-pen rule | Calibri Light / Calibri |
| `midnight` (dark) | Deep slate with grain, sky-blue accent bar, left-aligned covers | Calibri Light / Calibri |
| `graphite` (dark) | Graphite bar with an orange stripe (metropolis, dark) | Calibri |
| `aurora` (dark) | Blue-to-violet gradient, cyan structure, rounded title box | Calibri Light / Calibri |
| `chalkboard` (dark) | Blackboard texture, chalk-white text, chalk-yellow titles | Cambria / Calibri |

**Fonts:** use fonts installed wherever the deck is opened (Office fonts: Calibri, Cambria,
Constantia, Georgia, Times New Roman, Arial...). For the **body**, pick one with lining figures
(digits of equal height): Georgia and Constantia have old-style figures that bounce in tables and
axes, so use them for headings only. Cambria matches Cambria Math, the equation font.

**Geometry** (title bar, content area, footline) is set by the constants at the top of
`build_template.py` (`BAR_H`, `TITLE`, `BODY`, `FOOT_H`): 16:9, content area 0.6 in from the sides.
The result is an ordinary template; finish it by hand (Route C) if needed.

## 3. Route B: adapt an existing template

1. `python ppt_templated.py --inspect company.potx` shows which layouts were found.
2. Missing or differently named layouts: rename them in PowerPoint, or map them in the YAML:
   `meta.layouts: {title: "Cover", frame: "Content", columns: "Comparison", title_only: "Heading Only", section: "Divider", plain: "Empty"}`.
3. The **Title and Content** placeholder is the content area: make sure it clears logos and the footer.
4. Optional: on **Title Slide**, add text placeholders named `Author`, `Institute`, `Date`
   (Slide Master → Insert Placeholder → Text, then rename in the Selection Pane).
5. Set Accent 1/2/3 to structure / alert / example colors; optionally add footline token text boxes.
6. `python ppt_templated.py examples/example.yaml -t company.potx --preview` and go through §5.

Any `.pptx`/`.potx` works; even a stock blank template gives a valid deck.

## 4. Route C: restyle in PowerPoint

Open a template (its preview slides are removed on every build, so leave them).

- **Colors / fonts:** Design → Variants ▸ → Colors / Fonts → Customize.
- **Background:** View → Slide Master → select the master → Background Styles → Format Background
  (colour, gradient or picture); the layouts follow it.
- **Master:** View → Slide Master (the top thumbnail is the master, the rest are layouts). Bar,
  stripes and footline are shapes named *Frame Title Bar*, *Frame Title Stripe/Rule/Accent*,
  *Footline Author/Title/Date/Slide Number* (Home → Arrange → Selection Pane). Fill them with
  *theme* colors so they follow the palette. Title Slide and Section Header hide the master's
  shapes and carry their own footline copy.
- **Bullets:** select a level of the master's body placeholder → Home → Bullets ▸ → Bullets and Numbering.
- **Content area:** move/resize the content placeholder of *Title and Content*.
- **Logo:** on the master (all frames) or the Title Slide layout only; keep it clear of the content area.
- Save as `.pptx` or `.potx`.

## 5. Checklist (on the `--preview` contact sheets)

- [ ] Title slide: title, subtitle, author, institute and date in place
- [ ] Footline: name, short title, date, correct slide numbers
- [ ] Frame titles readable on the bar; long titles shrink without overflowing
- [ ] Bullets: symbols and colors at levels 1–3
- [ ] Blocks: Accent 1/2/3 headers, readable tinted bodies
- [ ] Charts and plots: series colors distinct; text in the template font
- [ ] Nothing overlaps a logo or the footline; no `(measured)` warnings
- [ ] In PowerPoint: Home → New Slide gives matching slides; bullets build on click; Morph works

## 6. Pitfalls

- Formatting typed on slides in the template is lost: only the master and layouts count.
- Use **theme colors**, not custom RGB, for anything that should follow the palette.
- Set 16:9 or 4:3 (Design → Slide Size) before designing layouts; changing it later rescales everything.
- Background pictures belong in the master's background fill, not as a picture over the content area.
- A dark template made by hand: set the master's colour map to *Background 1 = Dark 1* (or start from
  `build_template.py` with `dark: true`); a dark picture over a light theme makes every tint wrong.
- Layouts are matched by *layout* name; `Author`/`Institute`/`Date` by *placeholder* name.
