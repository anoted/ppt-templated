#!/usr/bin/env python3
"""
build_template.py - generate Beamer-inspired PowerPoint templates.

    python build_template.py                         # builds all themes into templates/
    python build_template.py --theme blue            # just one
    python build_template.py --spec my_theme.yaml    # a theme from a file (same keys as THEMES)
    python build_template.py --list                  # the built-in themes

Each template is a normal .pptx whose slide master and layouts carry all the formatting:
  * theme colours  (accent1 = structure, accent2 = alert, accent3 = example, accent4-6 = charts),
                   light or dark (`dark: true` maps the colours the way PowerPoint's own dark themes do)
  * theme fonts    (headings / body)
  * a background   (theme colour, gradient, generated texture such as paper or chalk, or a picture)
  * layouts with the standard PowerPoint names ppt_templated.py looks for:
      Title Slide, Section Header, Title and Content, Two Content, Title Only, Blank
  * a Beamer "footline" (author | short title | date + slide number) whose text uses tokens
    such as {author} that ppt_templated.py fills in.

You can also open the result in PowerPoint (View > Slide Master) and restyle it by hand -
see TEMPLATE_GUIDE.md, which lists every theme key.
"""

import argparse
import copy
import difflib
import io
import os
import re
import sys

import yaml
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
EMU = 914400
W, H = 13.333, 7.5

# ------------------------------------------------------------------------------------------
# Colour themes.  Colour specs: "accent1" | ("accent1", lumMod%, lumOff%) - always theme-linked
# ------------------------------------------------------------------------------------------
THEMES = {
    "blue": {  # Beamer "Madrid": blue structure, rounded title box, three-part footline
        "label": "Beamer Blue (Madrid-inspired)",
        "colors": dict(dk1="1F1F24", lt1="FFFFFF", dk2="1A1A5E", lt2="EBEBF7", accent1="3333B3",
                       accent2="C0161B", accent3="1E7B34", accent4="E69F00", accent5="3E9BD6",
                       accent6="7A7A8C", hlink="3333B3", folHlink="6A3DB8"),
        "fonts": ("Calibri", "Calibri"),
        "bar": "accent1", "bar_text": "bg1",
        "foot": [("accent1", 50, 0), ("accent1", 75, 0), "accent1"], "foot_text": ["bg1", "bg1", "bg1"],
    },
    "charcoal": {  # Beamer "metropolis" colours: dark teal structure, orange alerts
        "label": "Beamer Charcoal (Metropolis-inspired)",
        "colors": dict(dk1="23373B", lt1="FFFFFF", dk2="23373B", lt2="EEF1F2", accent1="23373B",
                       accent2="EB811B", accent3="14B03D", accent4="4A8FA3", accent5="A3B6BA",
                       accent6="B8860B", hlink="EB811B", folHlink="B5651D"),
        "fonts": ("Calibri", "Calibri"),
        "bar": "accent1", "bar_text": "bg1",
        "foot": ["accent1", "accent1", "accent1"], "foot_text": ["bg1", ("bg1", 75, 0), "bg1"],
    },
    "crimson": {  # Beamer "Beaver": dark red structure on light grey title bar, serif headings
        "label": "Beamer Crimson (Beaver-inspired)",
        "colors": dict(dk1="222222", lt1="FFFFFF", dk2="5C0000", lt2="EFEFEF", accent1="8B0000",
                       accent2="D4430C", accent3="2E6B30", accent4="C9A227", accent5="5B7C99",
                       accent6="8C8C8C", hlink="8B0000", folHlink="5C0000"),
        "fonts": ("Cambria", "Calibri"),
        "bar": "bg2", "bar_text": "accent1",
        "foot": ["accent1", ("bg2", 90, 0), "bg2"], "foot_text": ["bg1", "tx1", "tx1"],
    },
    # ---- academic looks: serif Office fonts (installed with Office on Windows and Mac) -----
    "paper": {  # journal / preprint: ink on paper, hairline rules instead of colour bars
        "label": "Beamer Paper (journal-style)",
        "colors": dict(dk1="1A1A1A", lt1="FDFCF8", dk2="1F3A5F", lt2="F1EFE8", accent1="1F3A5F",
                       accent2="9E1B1B", accent3="2F6B3A", accent4="B07D2B", accent5="5E7C99",
                       accent6="7F7F7F", hlink="1F3A5F", folHlink="5A3E85"),
        "fonts": ("Georgia", "Times New Roman"),  # body needs lining figures: Georgia's 0-9 bounce in tables
        "frame": "rule", "title_color": "tx1", "rule": "tx1",
        "footline": "text", "foot_rule": ("tx1", 50, 50),
        "cover": "rules", "cover_text": "tx1", "subtitle_italic": True,
        "bullets": [("•", 100), ("–", 100), ("•", 100)],
    },
    "oxford": {  # formal university: navy frame bar with gold stripes, square title box
        "label": "Beamer Oxford (university navy and gold)",
        "colors": dict(dk1="1C1C1C", lt1="FFFFFF", dk2="002147", lt2="F0EEE8", accent1="002147",
                       accent2="A6192E", accent3="3D6B4F", accent4="B8923A", accent5="4F7CAC",
                       accent6="8C8C8C", hlink="002147", folHlink="5B2C6F"),
        "fonts": ("Constantia", "Cambria"),  # Constantia has old-style figures, so not for body/tables
        "bar": "accent1", "bar_text": "bg1", "bar_stripe": "accent4",
        "foot": ["accent1", "accent1", "accent1"], "foot_text": ["bg1", ("bg1", 75, 0), "bg1"],
        "foot_stripe": "accent4",
        "radius": 0, "subtitle_italic": True,
        "bullets": [("■", 60), ("–", 100), ("•", 100)],
    },
    "classic": {  # LaTeX Beamer's default theme: no bar, blue structure, Cambria (= the math font)
        "label": "Beamer Classic (default-theme-inspired)",
        "colors": dict(dk1="000000", lt1="FFFFFF", dk2="1A1A5E", lt2="EBEBF7", accent1="3333B3",
                       accent2="C00000", accent3="1B7A1B", accent4="E69F00", accent5="3E9BD6",
                       accent6="7A7A8C", hlink="3333B3", folHlink="6A3DB8"),
        "fonts": ("Cambria", "Cambria"),
        "frame": "plain", "title_color": "accent1",
        "footline": "text",
        "cover": "plain", "cover_text": "accent1",
    },
    "modern": {  # clean sans-serif: white slides, a light hairline under the title, left-aligned covers
        "label": "Beamer Modern (clean sans-serif)",
        "colors": dict(dk1="1F2933", lt1="FFFFFF", dk2="1B2A41", lt2="F1F4F7", accent1="0B7189",
                       accent2="D1495B", accent3="2A7F62", accent4="E9A03B", accent5="5C6BC0",
                       accent6="8A94A6", hlink="0B7189", folHlink="5C6BC0"),
        "fonts": ("Calibri Light", "Calibri"),
        "frame": "rule", "title_color": "tx2", "rule": ("bg1", 85, 0),
        "footline": "text",
        "cover": "left", "cover_text": "tx2",
        "bullets": [("•", 100), ("–", 100), ("•", 100)],
    },
    # ---- dark themes: light text on a dark background; accents light enough to carry dark text -----
    "midnight": {  # deep slate, sky-blue structure, coral alerts
        "label": "Midnight (dark, modern sans)",
        "dark": True,
        "colors": dict(text="E2E8F0", background="0F172A", text2="F8FAFC", background2="1E293B",
                       accent1="38BDF8", accent2="FB7185", accent3="4ADE80", accent4="FBBF24", accent5="A78BFA",
                       accent6="94A3B8", hlink="7DD3FC", folHlink="C4B5FD"),
        "fonts": ("Calibri Light", "Calibri"),
        "background": {"texture": "grain", "strength": 0.6},
        "frame": "rule", "title_color": "tx2", "rule": ("bg1", 75, 25),
        "footline": "text",
        "cover": "left", "cover_text": "tx2",
        "bullets": [("•", 100), ("–", 100), ("•", 100)],
    },
    "graphite": {  # metropolis on dark: graphite title bar, teal structure, orange alerts
        "label": "Graphite (dark, Metropolis-inspired)",
        "dark": True,
        "colors": dict(text="ECEFF1", background="1F2428", text2="FFFFFF", background2="2B3137",
                       accent1="56B6C2", accent2="F28C28", accent3="98C379", accent4="E5C07B", accent5="C678DD",
                       accent6="8B949E", hlink="F28C28", folHlink="E5C07B"),
        "fonts": ("Calibri", "Calibri"),
        "bar": "bg2", "bar_text": "tx1",
        "footline": "text",
        "cover": "left", "cover_text": "tx2",
        "bullets": [("▸", 100), ("–", 100), ("•", 100)],
    },
    "chalkboard": {  # lecture-hall blackboard: chalk-white text, chalk-yellow titles, pastel accents
        "label": "Chalkboard (dark, blackboard texture)",
        "dark": True,
        "colors": dict(text="F2F2EC", background="26332D", text2="F2D16B", background2="2F3E37",
                       accent1="F2D16B", accent2="F4978E", accent3="A8D8A0", accent4="9CC9F2", accent5="F5B97A",
                       accent6="C9C9C0", hlink="9CC9F2", folHlink="F5B97A"),
        "fonts": ("Cambria", "Calibri"),
        "background": {"texture": "chalk"},
        "frame": "rule", "title_color": "tx2", "rule": ["tx1", 60, 0],
        "footline": "text",
        "cover": "rules", "cover_text": "tx2",
        "bullets": [("•", 100), ("–", 100), ("•", 100)],
    },
    "aurora": {  # dark gradient from night blue to violet, cyan structure
        "label": "Aurora (dark gradient)",
        "dark": True,
        "colors": dict(text="EAEFF5", background="141E30", text2="FFFFFF", background2="2A2350",
                       accent1="64D2FF", accent2="FF6B8B", accent3="7EE081", accent4="FFD166", accent5="B39DDB",
                       accent6="90A4AE", hlink="64D2FF", folHlink="B39DDB"),
        "fonts": ("Calibri Light", "Calibri"),
        "background": {"gradient": ["bg1", "bg2"], "angle": 35},
        "frame": "rule", "title_color": "tx2", "rule": ("bg1", 75, 25),
        "footline": "text",
        "cover": "plain", "cover_text": "tx2",
        "bullets": [("•", 100), ("–", 100), ("•", 100)],
    },
    # ---- paper-like and other light looks ---------------------------------------------------------
    "parchment": {  # old book: textured paper, sepia ink, serif
        "label": "Parchment (paper texture, serif)",
        "colors": dict(text="3B2F2F", background="F4ECD8", text2="5A3E2B", background2="EADFC4",
                       accent1="6B3E26", accent2="A23B2A", accent3="4A6B3A", accent4="B8892E", accent5="3E5C76",
                       accent6="8C7B6B", hlink="6B3E26", folHlink="3E5C76"),
        "fonts": ("Georgia", "Cambria"),  # Cambria has lining figures for tables and axes
        "background": {"texture": "paper"},
        "frame": "rule", "title_color": "accent1", "rule": "accent1",
        "footline": "text", "foot_rule": ["tx1", 50, 50],
        "cover": "rules", "cover_text": "accent1", "subtitle_italic": True,
        "bullets": [("❧", 90), ("–", 100), ("•", 100)],
    },
    "notebook": {  # graph paper with ink-blue structure and red-pen alerts
        "label": "Notebook (graph paper)",
        "colors": dict(text="1F2A44", background="FBFBF7", text2="1F2A44", background2="EEF1F7",
                       accent1="2B59C3", accent2="D7263D", accent3="2E8B57", accent4="F4A259", accent5="8E7DBE",
                       accent6="8A8F98", hlink="2B59C3", folHlink="8E7DBE"),
        "fonts": ("Calibri Light", "Calibri"),
        "background": {"texture": "grid", "strength": 0.8},
        "frame": "rule", "title_color": "accent1", "rule": "accent2",
        "footline": "text",
        "cover": "left", "cover_text": "accent1",
        "bullets": [("•", 100), ("–", 100), ("•", 100)],
    },
    "minimal": {  # Swiss style: black on white, one red accent, bold titles, nothing else
        "label": "Minimal (black and white, one red accent)",
        "colors": dict(text="111111", background="FFFFFF", text2="111111", background2="F2F2F2",
                       accent1="1D1D1F", accent2="E63946", accent3="2A9D8F", accent4="F4A261", accent5="457B9D",
                       accent6="8D99AE", hlink="E63946", folHlink="457B9D"),
        "fonts": ("Arial", "Arial"),
        "frame": "plain", "title_color": "tx1", "title_bold": True,
        "footline": "text",
        "cover": "left", "cover_text": "tx1",
        "bullets": [("–", 100), ("•", 100), ("–", 100)],
    },
}
DESCRIPTIONS = {  # one line each for `--list` and `ppt_templated.py --templates`
    "blue": "Blue title bar and three-part footline (Beamer Madrid)",
    "charcoal": "Dark slate title bar, orange alerts (Beamer metropolis)",
    "crimson": "Light grey title bar, dark red titles (Beamer Beaver)",
    "paper": "Journal look: hairline rules, grey footline, Georgia / Times New Roman",
    "oxford": "Navy title bar with gold stripes, square title box, Constantia / Cambria (default)",
    "classic": "No bars, blue titles, Cambria (Beamer default theme)",
    "modern": "White, hairline under titles, left-aligned covers, Calibri Light / Calibri",
    "midnight": "Dark slate with fine grain, sky-blue structure, coral alerts",
    "graphite": "Dark graphite title bar, teal structure, orange alerts (metropolis, dark)",
    "chalkboard": "Blackboard texture, chalk-white text, chalk-yellow titles",
    "aurora": "Dark blue-to-violet gradient, cyan structure, centred covers",
    "parchment": "Paper texture, sepia ink, Georgia / Cambria",
    "notebook": "Graph-paper background, ink-blue structure, red-pen alerts",
    "minimal": "Black on white, bold titles, a single red accent, Arial",
}
# Style keys (all optional; the defaults give the Madrid look of "blue"). TEMPLATE_GUIDE.md explains them.
DEFAULTS = {"fonts": ("Calibri", "Calibri"), "bar": "accent1", "bar_text": "bg1",
            "foot": ["accent1", "accent1", "accent1"], "foot_text": ["bg1", "bg1", "bg1"]}
CHOICES = {"frame": ("bar", "rule", "plain"), "footline": ("boxes", "text"),
           "cover": ("box", "rules", "left", "plain")}
THEME_KEYS = {"label", "description", "dark", "colors", "fonts", "background", "bar", "bar_text", "bar_stripe",
              "frame", "title_color", "rule", "footline", "foot", "foot_text", "foot_rule", "foot_stripe", "cover",
              "cover_text", "radius", "bullets", "bullet_color", "title_bold", "subtitle_italic"}
SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
         "hlink", "folHlink")
ROLE_NAMES = {"text": "dk1", "background": "lt1", "text2": "dk2", "background2": "lt2"}  # theme-file aliases
COLOR_REFS = {"tx1", "tx2", "bg1", "bg2", *SLOTS}
BACKGROUND_KEYS = {"fill", "gradient", "angle", "texture", "strength", "ink", "image"}

# Geometry (inches) - change here to move things in every layout at once
BAR_H = 0.95
TITLE = (0.45, 0.1, W - 0.9, 0.75)
BODY = (0.6, 1.3, W - 1.2, 5.6)
FOOT_H = 0.32
FOOT_Y = H - FOOT_H
COL_W = (BODY[2] - 0.4) / 2
RULE = 0.014    # hairline (~1 pt)
STRIPE = 0.05   # accent stripe under the bar / above the footline


def is_ref(v) -> bool:
    """A theme colour: "accent1" or [name, lumMod %, lumOff %]."""
    return (isinstance(v, str) and v in COLOR_REFS) or (
        isinstance(v, (list, tuple)) and len(v) == 3 and v[0] in COLOR_REFS
        and all(isinstance(x, (int, float)) and not isinstance(x, bool) for x in v[1:]))


def check_theme(t: dict, name: str, base_dir: str = HERE) -> dict:
    """The theme with defaults filled in; exits with a clear message on an unknown key or a bad value.
    base_dir: folder that a background picture path is relative to."""
    def fail(msg):
        sys.exit(f"error: theme '{name}': {msg}")

    if not isinstance(t, dict):
        fail("a theme must be a mapping of theme keys (see TEMPLATE_GUIDE.md)")
    for k in t:
        if k not in THEME_KEYS:
            hit = difflib.get_close_matches(str(k), THEME_KEYS, n=1)
            fail(f"unknown key '{k}'" + (f" (did you mean '{hit[0]}'?)" if hit else "")
                 + f"; keys: {', '.join(sorted(THEME_KEYS))}")
    t = {**DEFAULTS, "label": f"Beamer {name.title()}", **t}
    if not isinstance(t.get("dark", False), bool):
        fail("dark must be true or false")
    if t.get("frame", "bar") != "bar":  # no bar behind the title: the bar's (white) text color would vanish
        t.setdefault("title_color", "tx1")
    given = {}
    for k, v in (t.get("colors") or {}).items():
        slot = ROLE_NAMES.get(k, k)
        if slot not in SLOTS:
            fail(f"unknown colour slot '{k}'; slots: {', '.join(SLOTS)} (or text, background, text2, background2)")
        given[slot] = v
    if t.get("dark") and not {"dk1", "lt1"} <= set(given):
        fail("a dark theme needs its own text and background colours (colors: {text: ..., background: ...})")
    colors = {**THEMES["oxford"]["colors"], **given}  # missing slots: Oxford's
    for k, v in colors.items():
        v = f"{v:06d}" if isinstance(v, int) else str(v).lstrip("#")  # YAML reads 000000 as a number
        if not re.fullmatch(r"[0-9A-Fa-f]{6}", v):
            fail(f"colour {k}: '{v}' is not a hex colour like 1F3A5F (quote it in YAML)")
        colors[k] = v.upper()
    t["colors"] = colors  # by role: dk1 = text, lt1 = background, dk2 = text 2, lt2 = background 2
    if not isinstance(t["fonts"], (list, tuple)) or len(t["fonts"]) != 2:
        fail("fonts must be [headings font, body font]")
    for k, ok in CHOICES.items():
        if t.get(k) is not None and t[k] not in ok:
            fail(f"{k} must be one of {', '.join(ok)}, got {t[k]!r}")
    for k in ("bar", "bar_text", "bar_stripe", "title_color", "rule", "foot_rule", "foot_stripe", "cover_text",
              "bullet_color"):
        if t.get(k) is not None and not is_ref(t[k]):
            fail(f"{k}: {t[k]!r} is not a theme colour (accent1-6, tx1, tx2, bg1, bg2, "
                 f"or [name, lumMod %, lumOff %])")
    for k in ("foot", "foot_text"):
        if not isinstance(t[k], (list, tuple)) or len(t[k]) != 3 or not all(is_ref(v) for v in t[k]):
            fail(f"{k} needs three theme colours: author, title and date segments")
    t["background"] = _check_background(t.get("background"), fail, base_dir)
    return t


def _check_background(bg, fail, base_dir):
    """background: theme colour | texture name | {fill} | {gradient, angle} | {texture, strength, ink} | {image}."""
    import textures
    if bg is None:
        return {"fill": "bg1"}
    if isinstance(bg, str) and bg in textures.TEXTURES:
        bg = {"texture": bg}
    elif is_ref(bg):
        bg = {"fill": bg}
    elif not isinstance(bg, dict):
        fail(f"background: {bg!r} is not a theme colour, a texture ({', '.join(textures.TEXTURES)}) or a mapping")
    bad = set(bg) - BACKGROUND_KEYS
    if bad:
        fail(f"background: unknown key(s) {', '.join(map(str, bad))}; use {', '.join(sorted(BACKGROUND_KEYS))}")
    kinds = [k for k in ("fill", "gradient", "texture", "image") if k in bg]
    if len(kinds) != 1:
        fail("background needs exactly one of fill, gradient, texture or image")
    bg = dict(bg)
    if "fill" in bg and not is_ref(bg["fill"]):
        fail(f"background fill: {bg['fill']!r} is not a theme colour")
    if "gradient" in bg:
        g = bg["gradient"]
        if not isinstance(g, (list, tuple)) or len(g) < 2 or not all(is_ref(c) for c in g):
            fail("background gradient needs two or more theme colours, e.g. [bg1, bg2]")
        try:
            bg["angle"] = float(bg.get("angle", 90))
        except (TypeError, ValueError):
            fail("background angle must be a number of degrees")
    if "texture" in bg:
        if bg["texture"] not in textures.TEXTURES:
            fail(f"background texture must be one of {', '.join(textures.TEXTURES)}, got {bg['texture']!r}")
        try:
            bg["strength"] = float(bg.get("strength", 1.0))
        except (TypeError, ValueError):
            fail("background strength must be a number (1 = default)")
        if bg.get("ink") is not None and bg["ink"] not in COLOR_REFS:
            fail(f"background ink must be a theme colour name (accent1, tx1, ...), got {bg['ink']!r}")
    if "image" in bg:
        path = str(bg["image"])
        path = path if os.path.isabs(path) else os.path.join(base_dir, path)
        if not os.path.exists(path):
            fail(f"background image not found: {bg['image']} (relative to the theme file)")
        if os.path.splitext(path)[1].lower() not in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"):
            fail("background image must be a PNG, JPEG, GIF, BMP or TIFF picture")
        bg["image"] = path
    return bg


def grey(t, strong=False):
    """Quiet text colour (footline, institute line): mid grey on light themes, light grey on dark ones."""
    if t.get("dark"):
        return ("tx1", 75, 0) if strong else ("tx1", 60, 0)
    return ("tx1", 65, 35) if strong else ("tx1", 50, 50)


def hairline(t):
    """A line just visible on the background."""
    return ("bg1", 75, 25) if t.get("dark") else ("bg1", 85, 0)


def role_hex(t, ref) -> str:
    """'RRGGBB' of a plain theme colour name, as the slides show it (tx1 = text, also on dark themes)."""
    c = t["colors"]
    roles = {"tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}
    if ref in roles:
        return c[roles[ref]]
    if t.get("dark") and ref in ("dk1", "lt1", "dk2", "lt2"):  # raw slots hold the swapped colours
        return c[{"dk1": "lt1", "lt1": "dk1", "dk2": "lt2", "lt2": "dk2"}[ref]]
    return c[ref]


def e(v):
    return int(round(v * EMU))


def clr(spec):
    if isinstance(spec, str):
        return f'<a:schemeClr val="{spec}"/>'
    name, mod, off = spec
    mods = (f'<a:lumMod val="{mod * 1000}"/>' if mod else "") + (f'<a:lumOff val="{off * 1000}"/>' if off else "")
    return f'<a:schemeClr val="{name}">{mods}</a:schemeClr>'


def xfrm(x, y, w, h):
    return f'<a:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'


def shape(sid, name, box, *, fill=None, geom="rect", adj="", text="", size=1100, color="bg1",
          align="ctr", bold=False, field=False, lins=0.1):
    fill_xml = f"<a:solidFill>{clr(fill)}</a:solidFill>" if fill else "<a:noFill/>"
    run = ""
    if field:
        run = (f'<a:fld id="{{7C1D2A8E-3B7F-4C11-9D35-2E4B5A6C7D8E}}" type="slidenum"><a:rPr lang="en-US" '
               f'sz="{size}" b="{int(bold)}"><a:solidFill>{clr(color)}</a:solidFill></a:rPr><a:t>‹#›</a:t></a:fld>')
    elif text:
        run = (f'<a:r><a:rPr lang="en-US" sz="{size}" b="{int(bold)}"><a:solidFill>{clr(color)}</a:solidFill>'
               f'</a:rPr><a:t>{text}</a:t></a:r>')
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(*box)}<a:prstGeom prst="{geom}"><a:avLst>{adj}</a:avLst></a:prstGeom>{fill_xml}'
            f'<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="{e(lins)}" tIns="0" '
            f'rIns="{e(lins)}" bIns="0" rtlCol="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
            f'<a:p><a:pPr algn="{align}"/>{run}<a:endParaRPr lang="en-US" sz="{size}"/></a:p></p:txBody></p:sp>')


def footline(t, start_id):
    """Beamer footline: author | short title | date  n."""
    seg = W / 3
    parts = []
    names = ["Footline Author", "Footline Title", "Footline Date"]
    texts = ["{author}", "{short_title}", "{date}"]
    if t.get("footline", "boxes") == "text":  # quiet grey line of text, optional hairline above
        color, margin = grey(t), BODY[0]
        if t.get("foot_rule"):
            parts.append(shape(start_id + 4, "Footline Rule", (margin, FOOT_Y, W - 2 * margin, RULE),
                               fill=t["foot_rule"]))
        for i, align in enumerate(["l", "ctr", "ctr"]):
            parts.append(shape(start_id + i, names[i], (seg * i, FOOT_Y, seg, FOOT_H), text=texts[i],
                               size=1000, color=color, align=align, lins=margin if i == 0 else 0.1))
        parts.append(shape(start_id + 3, "Footline Slide Number", (W - margin - 0.85, FOOT_Y, 0.85, FOOT_H),
                           field=True, size=1000, color=color, align="r", lins=0))
        return "".join(parts)
    if t.get("foot_stripe"):
        parts.append(shape(start_id + 4, "Footline Stripe", (0, FOOT_Y - STRIPE, W, STRIPE), fill=t["foot_stripe"]))
    for i in range(3):
        parts.append(shape(start_id + i, names[i], (seg * i, FOOT_Y, seg + (0.01 if i < 2 else 0), FOOT_H),
                           fill=t["foot"][i], text=texts[i], size=1100, color=t["foot_text"][i]))
    parts.append(shape(start_id + 3, "Footline Slide Number", (W - 1.0, FOOT_Y, 0.85, FOOT_H),
                       field=True, size=1100, color=t["foot_text"][2], align="r", lins=0))
    return "".join(parts)


def ph(sid, name, ph_attrs, box, *, lst="", body_pr="", prompt="", geom=None, fill=None, levels=1,
       autofit=False):
    geom_xml = f'<a:prstGeom prst="{geom[0]}"><a:avLst>{geom[1]}</a:avLst></a:prstGeom>' if geom else ""
    fill_xml = f"<a:solidFill>{clr(fill)}</a:solidFill>" if fill else ""
    body = f"<a:bodyPr {body_pr}><a:normAutofit/></a:bodyPr>" if autofit else f"<a:bodyPr {body_pr}/>"
    if levels == 1:
        paras = f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{prompt}</a:t></a:r></a:p>'
    else:
        labels = ["Click to edit Master text styles", "Second level", "Third level", "Fourth level", "Fifth level"]
        paras = "".join(f'<a:p><a:pPr lvl="{i}"/><a:r><a:rPr lang="en-US"/><a:t>{labels[i]}</a:t></a:r></a:p>'
                        for i in range(levels))
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr><p:spPr>{xfrm(*box)}{geom_xml}{fill_xml}</p:spPr>'
            f'<p:txBody>{body}<a:lstStyle>{lst}</a:lstStyle>{paras}</p:txBody></p:sp>')


def tree(content, name, bg=""):
    """A <p:cSld> with the given shapes; bg: the fill XML of a background (masters only)."""
    bg_xml = f"<p:bg><p:bgPr>{bg}<a:effectLst/></p:bgPr></p:bg>" if bg else ""
    return (f'<p:cSld {NS} name="{name}">{bg_xml}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>'
            f'<p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            f'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{content}</p:spTree></p:cSld>')


def lvl(n, size, color="tx1", algn=None, extra="", attrs="", rattrs=""):
    a = f' algn="{algn}"' if algn else ""
    return (f'<a:lvl{n}pPr{a}{attrs}>{extra}<a:defRPr sz="{size}"{rattrs}><a:solidFill>{clr(color)}</a:solidFill>'
            f'</a:defRPr></a:lvl{n}pPr>')


# ------------------------------------------------------------------------------------------
def build(t, out_path):
    """Write the template for theme dict t (as returned by check_theme) to out_path."""
    prs = Presentation()  # python-pptx's blank default template: 1 master, 11 layouts
    prs.slide_width, prs.slide_height = e(W), e(H)

    # ---- 1. Theme: colours + fonts ------------------------------------------------------
    master = prs.slide_master
    theme_part = master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    root.set("name", t["label"])
    cs = root.find(f".//{a}clrScheme")
    cs.set("name", t["label"])
    slots = dict(t["colors"])
    if t.get("dark"):  # dark background in dk1, light text in lt1, mapped back below (clrMap)
        for x, y in (("dk1", "lt1"), ("dk2", "lt2")):
            slots[x], slots[y] = t["colors"][y], t["colors"][x]
    for key, hexv in slots.items():
        el = cs.find(f"{a}{key}")
        for ch in list(el):
            el.remove(ch)
        etree.SubElement(el, f"{a}srgbClr").set("val", hexv)
    fs = root.find(f".//{a}fontScheme")
    fs.set("name", t["label"])
    fs.find(f"{a}majorFont/{a}latin").set("typeface", t["fonts"][0])
    fs.find(f"{a}minorFont/{a}latin").set("typeface", t["fonts"][1])
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ---- 2. Slide master: frame-title bar, footline, title + body placeholders ------------
    title_ph = ph(2, "Title Placeholder 1", 'type="title"', TITLE, prompt="Click to edit Master title style",
                  body_pr='vert="horz" lIns="91440" tIns="0" rIns="91440" bIns="0" rtlCol="0" anchor="ctr"', autofit=True)
    body_ph = ph(3, "Text Placeholder 2", 'type="body" idx="1"', BODY, levels=5,
                 body_pr='vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"', autofit=True)
    frame = t.get("frame", "bar")
    deco = ""
    if frame == "bar":
        deco = shape(10, "Frame Title Bar", (0, 0, W, BAR_H), fill=t["bar"])
        if t.get("bar_stripe"):
            deco += shape(16, "Frame Title Stripe", (0, BAR_H, W, STRIPE), fill=t["bar_stripe"])
    elif frame == "rule":
        deco = shape(10, "Frame Title Rule", (BODY[0], BAR_H - 0.03, BODY[2], RULE), fill=t.get("rule", "tx1"))
    master_tree = tree(deco + footline(t, 11) + title_ph + body_ph, "", background_xml(t, master))
    mel = master._element
    mel.replace(mel.find(qn("p:cSld")), parse_xml(master_tree))
    if t.get("dark"):  # PowerPoint's own dark themes: background 1 = dark 1, text 1 = light 1
        cmap = mel.find(qn("p:clrMap"))
        for role, slot in (("bg1", "dk1"), ("tx1", "lt1"), ("bg2", "dk2"), ("tx2", "lt2")):
            cmap.set(role, slot)

    # text styles: Beamer-like triangles / bullets in the structure colour
    def blvl(n, marl, ind, size, char, bu_pct, spc):
        return (f'<a:lvl{n}pPr marL="{e(marl)}" indent="{e(-ind)}" algn="l" defTabSz="914400" rtl="0" '
                f'eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>'
                f'<a:spcBef><a:spcPts val="{spc * 100}"/></a:spcBef><a:buClr>{clr(t.get("bullet_color", "accent1"))}</a:buClr>'
                f'<a:buSzPct val="{bu_pct * 1000}"/><a:buFont typeface="Arial" panose="020B0604020202020204"/>'
                f'<a:buChar char="{char}"/><a:defRPr sz="{size * 100}" kern="1200"><a:solidFill>'
                f'<a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/>'
                f'<a:cs typeface="+mn-cs"/></a:defRPr></a:lvl{n}pPr>')

    body_levels = [(0.34, 0.34, 22, "►", 70, 10), (0.74, 0.3, 20, "•", 100, 4), (1.1, 0.28, 18, "–", 100, 3),
                   (1.45, 0.26, 16, "•", 100, 2), (1.8, 0.26, 16, "–", 100, 2), (2.15, 0.26, 16, "•", 100, 2),
                   (2.5, 0.26, 16, "•", 100, 2), (2.85, 0.26, 16, "•", 100, 2), (3.2, 0.26, 16, "•", 100, 2)]
    if t.get("bullets"):
        bl = t["bullets"]
        body_levels = [(m, i, s, *bl[k % len(bl)], sp) for k, (m, i, s, _, _, sp) in enumerate(body_levels)]
    body_style = "".join(blvl(i + 1, *v) for i, v in enumerate(body_levels))
    title_style = (f'<a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
                   f'<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>'
                   f'<a:defRPr sz="2800" b="{int(t.get("title_bold", False))}" kern="1200">'
                   f'<a:solidFill>{clr(t.get("title_color", t.get("bar_text")))}</a:solidFill>'
                   f'<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>')
    txs = mel.find(qn("p:txStyles"))
    txs.replace(txs.find(qn("p:titleStyle")), parse_xml(f"<p:titleStyle {NS}>{title_style}</p:titleStyle>"))
    txs.replace(txs.find(qn("p:bodyStyle")), parse_xml(f"<p:bodyStyle {NS}>{body_style}</p:bodyStyle>"))

    # ---- 3. Layouts -------------------------------------------------------------------------
    layouts = list(prs.slide_layouts)
    keep = {0: "Title Slide", 1: "Title and Content", 2: "Section Header", 3: "Two Content",
            5: "Title Only", 6: "Blank"}
    for i, lay in enumerate(layouts):
        if i not in keep:
            prs.slide_layouts.remove(lay)

    no_bullet = '<a:buNone/>'
    radius = f'<a:gd name="adj" fmla="val {t.get("radius", 12000)}"/>'
    cover = t.get("cover", "box")
    cover_text = t.get("cover_text", "bg1" if cover == "box" else "accent1")
    box = dict(geom=("roundRect", radius), fill="accent1") if cover == "box" else {}
    sub_attrs = ' i="1"' if t.get("subtitle_italic") else ""
    left = cover == "left"
    al, ins, no_bu = ("l" if left else "ctr"), (0 if left else 0.3), (no_bullet, ' marL="0" indent="0"')
    if left:  # flush left, title above the subtitle; no lines or bars
        tx, tw, sx, sw = 0.9, W - 1.8, 0.9, W - 1.8
        T = dict(title=(tx, 1.4, tw, 1.6), size=4000, anchor="b", sub=3.32, author=4.1, inst=4.6, date=5.05)
        S = dict(title=(sx, 2.2, sw, 1.45), size=3600, anchor="b", sub=3.97)
    else:
        tx, tw, sx, sw = (W - 9.8) / 2, 9.8, 2.9, W - 5.8
        T = dict(title=(tx, 1.45, tw, 1.45), size=3600, anchor="ctr", sub=3.0, author=3.95, inst=4.5, date=5.2)
        S = dict(title=(sx, 2.85, sw, 1.15), size=3200, anchor="ctr", sub=4.15)
    pad = f'lIns="{e(ins)}" ' if left else ""

    def rules(top, bottom, x, w):
        """booktabs-style: heavier rule above the title, lighter rule below it."""
        if cover != "rules":
            return ""
        return (shape(30, "Title Rule Top", (x, top, w, RULE * 2), fill=cover_text)
                + shape(31, "Title Rule Bottom", (x, bottom, w, RULE), fill=cover_text))

    title_slide = (
        ph(2, "Title 1", 'type="ctrTitle"', T["title"], **box,
           prompt="Presentation Title", lst=lvl(1, T["size"], cover_text, al),
           body_pr=f'lIns="{e(ins)}" rIns="{e(ins)}" anchor="{T["anchor"]}"', autofit=True)
        + rules(1.35, T["title"][1] + T["title"][3] if left else 2.95, tx, tw)
        + ph(3, "Subtitle 2", 'type="subTitle" idx="1"', (tx, T["sub"], tw, 0.6), prompt="Subtitle",
             lst=lvl(1, 2200, "accent1", al, *no_bu, sub_attrs), body_pr=f'{pad}anchor="t"')
        + ph(4, "Author", 'type="body" sz="quarter" idx="13"', (tx, T["author"], tw, 0.5), prompt="Author Name",
             lst=lvl(1, 2200, "tx1", al, *no_bu), body_pr=f'{pad}anchor="ctr"')
        + ph(5, "Institute", 'type="body" sz="quarter" idx="14"', (tx, T["inst"], tw, 0.5), prompt="Institute",
             lst=lvl(1, 1800, grey(t, True), al, *no_bu), body_pr=f'{pad}anchor="ctr"')
        + ph(6, "Date", 'type="body" sz="quarter" idx="15"', (tx, T["date"], tw, 0.45), prompt="Date",
             lst=lvl(1, 1800, grey(t, True) if left else "tx1", al, *no_bu), body_pr=f'{pad}anchor="ctr"')
        + footline(t, 20))
    section = (
        ph(2, "Title 1", 'type="title"', S["title"], **box,
           prompt="Section Title", lst=lvl(1, S["size"], cover_text, al),
           body_pr=f'lIns="{e(ins)}" rIns="{e(ins)}" anchor="{S["anchor"]}"', autofit=True)
        + rules(2.75, S["title"][1] + S["title"][3] if left else 4.04, sx, sw)
        + ph(3, "Section Subtitle", 'type="body" idx="1"', (sx, S["sub"], sw, 0.6), prompt="Optional subtitle",
             lst=lvl(1, 2000, "accent1", al, *no_bu, sub_attrs), body_pr=f'{pad}anchor="t"')
        + footline(t, 20))
    frame_title = ph(2, "Title 1", 'type="title"', TITLE, prompt="Frame Title")
    content = frame_title + ph(3, "Content Placeholder 2", 'idx="1"', BODY, levels=3, body_pr='anchor="ctr"')
    two = (frame_title
           + ph(3, "Content Placeholder Left", 'sz="half" idx="1"', (BODY[0], BODY[1], COL_W, BODY[3]), levels=2,
                body_pr='anchor="ctr"')
           + ph(4, "Content Placeholder Right", 'sz="half" idx="2"', (BODY[0] + COL_W + 0.4, BODY[1], COL_W, BODY[3]),
                levels=2, body_pr='anchor="ctr"'))
    specs = {
        "Title Slide": (title_slide, False), "Section Header": (section, False),
        "Title and Content": (content, True), "Two Content": (two, True),
        "Title Only": (frame_title, True), "Blank": ("", False),
    }
    for lay in prs.slide_layouts:
        body, show_master = specs[lay.name]
        xml = tree(body, lay.name)
        lel = lay._element
        lel.replace(lel.find(qn("p:cSld")), parse_xml(xml))
        if show_master:
            lel.attrib.pop("showMasterSp", None)
        else:
            lel.set("showMasterSp", "0")
    set_catalog(prs, t)
    prs.save(out_path)
    return out_path


def set_catalog(prs, t):
    """What `ppt_templated.py --templates` lists: name, light/dark, one-line description."""
    cp = prs.core_properties
    cp.title, cp.author = t["label"], "ppt-templated"
    cp.subject = str(t.get("description") or "")
    cp.category = "dark" if t.get("dark") else "light"
    bg = t["background"]
    cp.keywords = "template " + next((f"{k} {bg[k]}" for k in ("texture", "gradient", "image") if k in bg), "plain")


def background_xml(t, master) -> str:
    """Fill XML of the master background; textures and pictures are added to the master as an image."""
    bg = t["background"]
    if "fill" in bg:
        return f"<a:solidFill>{clr(bg['fill'])}</a:solidFill>"
    if "gradient" in bg:
        n = len(bg["gradient"]) - 1
        stops = "".join(f'<a:gs pos="{round(i * 100000 / n)}">{clr(c)}</a:gs>' for i, c in enumerate(bg["gradient"]))
        return (f'<a:gradFill rotWithShape="1"><a:gsLst>{stops}</a:gsLst>'
                f'<a:lin ang="{int(round(bg["angle"] * 60000)) % 21600000}" scaled="0"/></a:gradFill>')
    if "texture" in bg:
        import textures
        default_ink = "accent1" if bg["texture"] in ("grid", "dots", "lines") else "tx1"
        data = textures.texture(bg["texture"], role_hex(t, "bg1"), role_hex(t, bg.get("ink") or default_ink),
                                size=(round(W * textures.PX_PER_IN), round(H * textures.PX_PER_IN)),
                                strength=bg["strength"])
        image = io.BytesIO(data)
    else:
        with open(bg["image"], "rb") as fh:
            image = io.BytesIO(fh.read())
    _, rid = master.part.get_or_add_image_part(image)
    return (f'<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{rid}"/><a:srcRect/>'
            f'<a:stretch><a:fillRect/></a:stretch></a:blipFill>')


DEMO = {
    "meta": {"title": "{title}", "subtitle": "Template preview - ppt-templated removes these slides",
             "author": "{author}", "institute": "{institute}", "date": "{date}"},
    "slides": [
        {"layout": "title", "title": "Beamer-style PowerPoint Template",
         "subtitle": "Layout preview - these slides are removed when ppt-templated builds a deck",
         "author": "Author Name", "institute": "Institute / Department", "date": "Date"},
        {"layout": "section", "title": "Section Header layout", "subtitle": "Used for layout: section"},
        {"title": "Title and Content layout",
         "content": [{"bullets": ["Frame title at the top; body text in the master bullet style",
                                  ["Second level", ["Third level"]],
                                  "Footline text comes from tokens such as {author} in the Slide Master",
                                  "Edit colours in Design > Variants > Colors, fonts in Variants > Fonts"]}]},
        {"title": "Two Content layout",
         "content": [{"columns": [[{"bullets": ["Left column", ["Detail"]]}], [{"bullets": ["Right column", ["Detail"]]}]]}]},
        {"title": "Title Only layout - generated blocks",
         "content": [{"columns": [
             [{"block": {"title": "Block", "text": "Copy these groups onto your own slides."}}],
             [{"alertblock": {"title": "Alert block", "text": "Uses accent 2 of the theme."}}],
             [{"exampleblock": {"title": "Example block", "text": "Uses accent 3 of the theme."}}]]}]},
        {"layout": "plain", "content": [{"text": "Blank layout (Beamer [plain] frame)", "align": "center"}]},
    ],
}


def load_spec(path):
    """(name, theme, folder) from a YAML/JSON theme file with the THEMES keys plus an optional `name`."""
    try:
        with open(path, encoding="utf-8") as fh:
            spec = yaml.safe_load(fh)
    except FileNotFoundError:
        sys.exit(f"error: theme file not found: {path}")
    except yaml.YAMLError as exc:
        sys.exit(f"error: {path} is not valid YAML\n{exc}")
    if not isinstance(spec, dict):
        sys.exit(f"error: {path} must be a mapping of theme keys (see TEMPLATE_GUIDE.md)")
    name = str(spec.pop("name", os.path.splitext(os.path.basename(path))[0]))
    if not re.fullmatch(r"[A-Za-z0-9_-]+", name):
        sys.exit(f"error: theme name {name!r} may only use letters, digits, - and _ (it becomes the file name)")
    return name, spec, os.path.dirname(os.path.abspath(path))


def make(name, theme, out, base_dir=HERE, preview=True):
    """Check a theme, write its template to `out` and (by default) add the layout preview slides."""
    t = check_theme({"description": DESCRIPTIONS.get(name, ""), **theme}, name, base_dir)
    build(t, out)
    if preview:
        import ppt_templated  # preview slides are made by the generator itself
        b = ppt_templated.DeckBuilder(copy.deepcopy(DEMO), out, HERE, substitute_tokens=False)
        prs = b.build()
        set_catalog(prs, t)  # the demo build wrote its own title into the file properties
        prs.save(out)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--theme", choices=list(THEMES), help="build only this built-in theme")
    ap.add_argument("--spec", metavar="FILE", help="build the theme described in a YAML/JSON file")
    ap.add_argument("-o", "--output", help="output path for --spec / --theme (default: <outdir>/beamer_<name>.pptx)")
    ap.add_argument("--outdir", default=os.path.join(HERE, "templates"))
    ap.add_argument("--list", action="store_true", help="list the built-in themes and exit")
    ap.add_argument("--no-preview", action="store_true", help="don't add the layout preview slides")
    a = ap.parse_args()
    if a.list:
        for key, t in THEMES.items():
            print(f"{key:<11} {'dark ' if t.get('dark') else 'light'}  {DESCRIPTIONS.get(key, t['label'])}")
        return
    jobs = ([load_spec(a.spec)] if a.spec else
            [(k, THEMES[k], HERE) for k in ([a.theme] if a.theme else THEMES)])
    os.makedirs(a.outdir, exist_ok=True)
    for name, theme, base in jobs:
        out = a.output if a.output and len(jobs) == 1 else os.path.join(a.outdir, f"beamer_{name}.pptx")
        make(name, theme, out, base, not a.no_preview)
        print("wrote", out)


if __name__ == "__main__":
    main()
