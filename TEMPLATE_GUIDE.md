# Template Making Guide

This guide explains how a ppt-templated template works and how to make your own. There are three routes:

- **A.** Restyle one of the sample Beamer templates in PowerPoint (easiest).
- **B.** Make your company's existing template work with ppt-templated.
- **C.** Generate templates in code with `build_template.py`.

---

## 1. How a template controls the look

PowerPoint formatting has four levels. Each level inherits from the one above it:

```
Theme          colors (12 slots) + heading/body fonts
  └─ Slide Master   background, frame-title bar, footline, text styles (bullets, sizes)
       └─ Layouts        Title Slide, Section Header, Title and Content, ...  (placeholder positions)
            └─ Slides         only the content: text, charts, tables
```

ppt-templated only writes at the **Slides** level and references the other three. That's why:

- changing a theme color later in PowerPoint recolors the whole deck, including blocks and charts
- slides you add by hand in PowerPoint automatically match the generated ones
- **Reset** on the Home tab puts a moved placeholder back where the layout says

**Rule of thumb:** put anything that should look the same on every slide in the master or the layouts, not on slides.

---

## 2. What ppt-templated expects from a template

### Layouts (found by name)

| ppt-templated uses it for | Layout name | Placeholders it needs |
|---|---|---|
| `layout: title` | **Title Slide** | Title, Subtitle; optionally text placeholders named **Author**, **Institute**, **Date** |
| `layout: section` | **Section Header** | Title; optionally one text placeholder (the section subtitle) |
| frames with one bullet list, `outline` | **Title and Content** | Title + one Content placeholder. **Its position defines the content area for every frame.** |
| frames with two bullet columns | **Two Content** | Title + two Content placeholders |
| all other frames (charts, blocks, tables…) | **Title Only** | Title |
| `layout: plain` | **Blank** | none |

These are PowerPoint's standard layout names, so most templates already have them. If yours uses other names, either rename the layouts or map them in the content file:

```yaml
meta:
  layouts: {title: "Cover", frame: "Content", columns: "Comparison", title_only: "Heading Only", section: "Divider", plain: "Empty"}
```

If a name isn't found, ppt-templated falls back to the layout *type* (every PowerPoint layout has one), and it stops with a clear message if nothing fits.

**Optional Title Slide fields:** ppt-templated finds the **Author**, **Institute** and **Date** placeholders by their *name* in the Selection Pane. If they're missing, those lines are added under the subtitle instead.

### Theme color roles

| Theme slot | Beamer role | Used for |
|---|---|---|
| Text/Background Dark 1 (`dk1`) | normal text | body text, table text |
| Text/Background Light 1 (`lt1`) | background | slide background, text on colored bars |
| Text/Background Light 2 (`lt2`) | light fill | the Crimson template's title bar; free for your own fills |
| Hyperlink | links | `[text](url)` |
| **Accent 1** | *structure* | title box, section box, bullets, `block`, chart series 1 |
| **Accent 2** | *alert* | `alertblock`, `==alert==` text, series 2 |
| **Accent 3** | *example* | `exampleblock`, series 3 |
| Accent 4-6 | — | chart series 4-6 |

So when you pick colors, make Accent 1 your main brand color, Accent 2 a warning or highlight color, and Accent 3 a positive or "example" color.

### Text styles (master body levels)

Generated bullet boxes, and bullets inside blocks, **copy the master's bullet styles**: bullet characters, colors, indents and sizes for levels 1-9. Get those right in the master and every list in every deck follows them. ppt-templated also reads the level 1-4 font sizes and the title size from the master.

### Footline tokens

Plain text boxes on the master or layouts (not placeholders) can contain tokens. ppt-templated replaces them in the generated deck:

| Token | Value from `meta` |
|---|---|
| `{title}` `{short_title}` | title / short title |
| `{subtitle}` | subtitle |
| `{author}` `{short_author}` | author |
| `{institute}` `{short_institute}` | institute |
| `{date}` | date (`today` gives today's date) |

You can combine them freely, e.g. `{author} ({short_institute})`. For the page number, use a **slide-number field** (Insert → Slide Number while editing a text box on the master). It updates automatically on every slide.

---

## 3. Route A: Restyle a sample template in PowerPoint

Open `templates/beamer_blue.pptx`. The six preview slides show each layout; ppt-templated deletes them when it builds a deck, so leave them in.

1. **Colors:** **Design → Variants ▸ (the drop-down arrow) → Colors → Customize Colors…**
   Set Accent 1/2/3 using the roles above. Give the scheme a name and save. Every bar, box, bullet and chart follows.
2. **Fonts:** **Design → Variants ▸ → Fonts → Customize Fonts…**
   Set the Heading and Body fonts. Stick to fonts installed on every machine that will open the deck (Calibri, Arial, Cambria, Segoe UI…), or embed them via File → Options → Save → *Embed fonts*.
3. **Open the master:** **View → Slide Master.** The top, larger thumbnail is the master; the smaller ones under it are the layouts.
4. **Frame-title bar and footline** are ordinary shapes on the master named *Frame Title Bar*, *Footline Author*, *Footline Title*, *Footline Date* and *Footline Slide Number* (see **Home → Arrange → Selection Pane**).
   - Recolor a shape via Shape Format → Shape Fill → a *theme* color, so it stays theme-linked.
   - Resize or delete them as you like. To remove the footline, delete the four footline shapes on the master **and** on the Title Slide and Section Header layouts (those two hide master graphics, so they carry their own copy).
   - Change what the footline says by editing the token text, e.g. `{short_author}`.
5. **Bullets:** click the master's body placeholder, select a level's line, then **Home → Bullets ▸ → Bullets and Numbering…** to choose the symbol, color and size. Set font sizes per level the same way.
6. **Logo:** insert it on the master (it appears on every frame) or only on the Title Slide layout. Keep it clear of the content area (the *Title and Content* body placeholder).
7. **Content area:** on the **Title and Content** layout, move or resize the content placeholder. ppt-templated uses exactly that rectangle for every frame, including charts and blocks.
8. **Save** as `.pptx` (or **File → Save As → PowerPoint Template (.potx)**). Point the content file at it: `template: my_template.pptx`.

Layout housekeeping, all in Slide Master view:
- **Rename a layout:** right-click the thumbnail → *Rename Layout*.
- **Add a placeholder:** Slide Master tab → *Insert Placeholder* → *Text*, then name it in the Selection Pane (e.g. `Author`).
- **Hide the master's bar and footline on one layout:** tick *Hide Background Graphics*.

## 4. Route B: Make your company template work

1. Open the company template and check **View → Slide Master** for the six layouts in §2. Most corporate templates have them, sometimes under other names; rename them or use `meta.layouts`.
2. On **Title Slide**, add text placeholders named **Author**, **Institute**, **Date** if you want those fields placed separately.
3. Check that **Title and Content** has a sensibly sized content placeholder, since it defines the content area.
4. Set **Accent 1/2/3** to match the block roles (structure / alert / example).
5. Optional: add token text boxes (`{author}`, `{date}`…) and a slide-number field if you want a Beamer footline.
6. Run a test: `python ppt_templated.py examples/example.yaml -t company.potx -o test.pptx`, then go through the checklist in §6.

ppt-templated works with any `.pptx`/`.potx`. Even a stock blank template produces a valid deck; the steps above just make it look intentional.

## 5. Route C: Build templates in code

`build_template.py` builds the sample templates from scratch. To add your own look, copy an entry in `THEMES`:

```python
"forest": {
    "label": "Beamer Forest",
    "colors": dict(dk1="1B1B1B", lt1="FFFFFF", dk2="1F3D2B", lt2="EEF3EF",
                   accent1="2C5F2D", accent2="C8553D", accent3="3E7CB1",
                   accent4="E0A458", accent5="7FA99B", accent6="8C8C8C",
                   hlink="2C5F2D", folHlink="1F3D2B"),
    "fonts": ("Cambria", "Calibri"),              # (headings, body)
    "bar": "accent1", "bar_text": "bg1",          # frame-title bar fill, title text color
    "foot": [("accent1", 50, 0), ("accent1", 75, 0), "accent1"],   # footline segments
    "foot_text": ["bg1", "bg1", "bg1"],
},
```

Then run `python build_template.py --theme forest`. Colors are theme references; a tuple like `("accent1", 75, 0)` means "accent 1, luminance 75%" (a darker shade).
Optional keys change the structure, not just the colors (the three academic templates use them):

| Key | Values | Example |
|---|---|---|
| `frame` | `"bar"` (default), `"rule"` (hairline under the title), `"plain"` | paper: `rule`, classic: `plain` |
| `title_color` | frame-title color when there is no bar | `"tx1"`, `"accent1"` |
| `bar_stripe` / `foot_stripe` | thin stripe under the bar / above the footline | oxford: `"accent4"` (gold) |
| `footline` | `"boxes"` (default) or `"text"` (grey text only); `foot_rule` adds a hairline above it | paper, classic |
| `cover` | Title/Section slides: `"box"` (default), `"rules"` (booktabs-style rules), `"plain"`; `cover_text` sets the title color, `radius: 0` makes the box square | paper: `rules`, oxford: square box |
| `bullets` | `[(char, size %), ...]` for levels 1-3; `bullet_color` | oxford: `[("■", 60), ("–", 100), ("•", 100)]` |
| `title_bold`, `subtitle_italic` | `True` / `False` | |

These add shapes named *Frame Title Rule*, *Frame Title Stripe*, *Footline Rule*, *Footline Stripe* and *Title Rule Top/Bottom*; restyle or delete them in Slide Master view like the bar and footline.

**Fonts for academic decks.** Choose a body font with *lining* figures (digits all the same height). Georgia and Constantia have old-style figures that bounce up and down in tables and chart axes, and PowerPoint can't switch them off, so use them for headings only. Times New Roman and Cambria are safe body fonts. Cambria also matches Cambria Math, the equation font. Georgia, Constantia, Cambria and Times New Roman all come with Office on Windows and Mac.

The geometry constants at the top (`BAR_H`, `TITLE`, `BODY`, `FOOT_H`) move things in every layout at once. The result is a normal template, so you can finish it by hand with Route A.

---

## 6. Test checklist

Build `examples/example.yaml` with your template and open it in PowerPoint:

- [ ] Title slide: title, subtitle, author, institute and date in the right places
- [ ] Footline shows your name, short title, date and the correct slide numbers
- [ ] Frame titles are readable on the title bar (contrast), and long titles shrink without overflowing
- [ ] Bullets use your symbols and colors at levels 1-3
- [ ] Blocks: header colors are Accent 1/2/3 and body tints are readable
- [ ] Charts: series colors come from the accents and are distinguishable
- [ ] Nothing overlaps a logo or the footline (adjust the content placeholder)
- [ ] **Home → New Slide** in the generated deck gives slides that match
- [ ] Slide Show: bullets build on click; Morph works on the two "Morph transitions" frames (Microsoft 365)

## 7. Pitfalls

- **Don't format slides by hand in the template.** Only the master and layouts matter; ppt-templated deletes existing slides.
- **Use theme colors, not custom RGB,** for anything that should follow the palette ("Theme Colors" row in color pickers).
- **16:9 vs 4:3:** set the size (Design → Slide Size) *before* designing layouts. Changing it later rescales everything.
- **Pictures as backgrounds** belong on the master (Format Background → Picture fill), not as a picture shape covering the content area.
- **Placeholder names vs. layout names:** layouts are matched by *layout* name; Author/Institute/Date are matched by *placeholder* name (Selection Pane).
- **Keep one master.** ppt-templated uses the layouts of the first master only.
