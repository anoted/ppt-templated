#!/usr/bin/env python3
"""
ppt_templated.py - build an editable, Beamer-style PowerPoint deck from a YAML content file.

    python ppt_templated.py talk.yaml                          # -> talk.pptx
    python ppt_templated.py talk.yaml -o out.pptx -t templates/beamer_blue.pptx

The formatting lives in the TEMPLATE (slide master, layouts, theme colours and fonts).
The CONTENT FILE only says what goes on each slide. Everything in the output is a native,
editable PowerPoint object: real placeholders, native charts (right-click > Edit Data),
native tables, native equations (when pandoc is installed) and native animations.

Requirements:  pip install python-pptx pyyaml pillow
Optional:      pandoc (or `pip install pypandoc_binary`) for native, editable equations.
               matplotlib + numpy for `plot:` figures (see plots.py).
Layout of `flow:` / `steps:` diagrams lives in flows.py (pure Python, no extra packages).

See YAML_GUIDE.md for the complete content-file reference and TEMPLATE_GUIDE.md for making templates.
When you change any YAML key, default, item/chart type, colour, transition or effect here,
update YAML_GUIDE.md in the same commit - it is the single source of truth for the format.
"""

from __future__ import annotations

import argparse
import copy
import datetime as _dt
import io
import itertools
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field

import yaml
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

# --------------------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------------------
EMU = 914400  # EMU per inch
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NSDECL = f'xmlns:a="{NS_A}" xmlns:p="{NS_P}"'

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(HERE, "templates", "beamer_oxford.pptx")

# Layout keys used by the script -> layout names searched in the template (first match wins),
# then the OOXML layout type as a fallback. Standard PowerPoint names, so most templates work.
DEFAULT_LAYOUTS = {
    "title": (["Title Slide"], "title"),
    "section": (["Section Header"], "secHead"),
    "frame": (["Title and Content"], "obj"),
    "columns": (["Two Content"], "twoObj"),
    "title_only": (["Title Only"], "titleOnly"),
    "plain": (["Blank"], "blank"),
}

THEME = {
    "accent1": MSO_THEME_COLOR.ACCENT_1, "accent2": MSO_THEME_COLOR.ACCENT_2,
    "accent3": MSO_THEME_COLOR.ACCENT_3, "accent4": MSO_THEME_COLOR.ACCENT_4,
    "accent5": MSO_THEME_COLOR.ACCENT_5, "accent6": MSO_THEME_COLOR.ACCENT_6,
    "tx1": MSO_THEME_COLOR.TEXT_1, "tx2": MSO_THEME_COLOR.TEXT_2,
    "bg1": MSO_THEME_COLOR.BACKGROUND_1, "bg2": MSO_THEME_COLOR.BACKGROUND_2,
    "text": MSO_THEME_COLOR.TEXT_1, "background": MSO_THEME_COLOR.BACKGROUND_1,
}
# Beamer-like semantic colours, all theme-linked so they follow the template.
SEMANTIC = {
    "structure": ("accent1", 0.0),
    "alert": ("accent2", 0.0),
    "example": ("accent3", 0.0),
    "muted": ("tx1", 0.45),
    "white": ("bg1", 0.0),
}
BLOCK_COLORS = {"block": "accent1", "alertblock": "accent2", "exampleblock": "accent3"}

ITEM_TYPES = ("bullets", "numbered", "text", "block", "alertblock", "exampleblock", "chart",
              "table", "image", "placeholder", "plot", "flow", "steps", "math", "code", "columns", "spacer")

MARK = "\u2063"  # invisible separator used to mark math runs until they are converted
FLOW_MIN_PT = 9   # flow/steps diagrams shrink to fit, but their text not below this size (pt)
GAP = 0.25        # vertical gap between stacked items (inches)
COL_GAP = 0.4     # gap between columns (inches)


def inch(v: float) -> Emu:
    return Emu(int(round(v * EMU)))


def to_in(emu) -> float:
    return (emu or 0) / EMU


@dataclass
class Box:
    x: float
    y: float
    w: float
    h: float


@dataclass
class Effect:
    spid: int
    effect: str = "fade"
    para: tuple | None = None  # (start, end) paragraph range, or None for whole shape
    dur: int = 500
    trigger: str = "click"      # click | with | after


# --------------------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------------------
def warn(msg: str) -> None:
    print(f"  warning: {msg}", file=sys.stderr)


def strip_markup(text: str) -> str:
    """Approximate visible text (for size estimates)."""
    t = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", str(text))
    t = re.sub(r"\$([^$]+)\$", lambda m: "x" * max(1, int(len(m.group(1)) * 0.55)), t)
    return re.sub(r"\*\*|==|`|(?<!\w)\*|\*(?!\w)", "", t)


def parse_aspect(v) -> float:
    """'16:9', '4/3' or 1.5 -> width / height."""
    m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*(?:[:/]\s*(\d+(?:\.\d+)?))?\s*", str(v))
    if not m or float(m[2] or 1) == 0:
        raise ValueError(f"placeholder aspect must look like 16:9, 4/3 or 1.5, got {v!r}")
    return float(m[1]) / float(m[2] or 1)


def est_lines(text: str, width_in: float, size_pt: float, char_w: float = 0.5) -> int:
    per_line = max(6.0, width_in * 72.0 / (size_pt * char_w))
    return sum(max(1, math.ceil(len(seg) / per_line)) for seg in strip_markup(text).split("\n"))


def set_color(font, spec) -> None:
    """spec: any colour spec of resolve_color(), e.g. 'accent1' | 'alert' | ['tx1', 0.4] | '#RRGGBB'."""
    if spec is None:
        return
    paint(font.color, spec)


def resolve_color(spec):
    """Colour spec -> (kind, value, brightness); kind 'theme' (slot name) or 'rgb' (hex).
    spec: structure | alert | example | muted | white | accent1-6 | tx1 | tx2 | bg1 | bg2 | '#RRGGBB'
    | [spec, brightness] (positive = lighter, negative = darker)."""
    bright = 0.0
    if isinstance(spec, (list, tuple)) and len(spec) == 2:
        spec, bright = spec[0], float(spec[1])
    spec = str(spec)
    if spec in SEMANTIC:
        spec, base = SEMANTIC[spec]
        bright = base + (1 - base) * bright if bright >= 0 else bright
    if spec.startswith("#"):
        return "rgb", spec[1:], bright
    if spec not in THEME:
        raise ValueError(f"unknown colour {spec!r}: use structure, alert, example, muted, white, "
                         f"accent1-6, tx1, tx2, bg1, bg2, '#RRGGBB' or [colour, brightness]")
    return "theme", spec, bright


def paint(color_format, spec, tint: float = 0.0) -> None:
    """Set a fill/line ColorFormat from a colour spec, optionally lightened further by `tint` (0-1)."""
    kind, val, bright = resolve_color(spec)
    if kind == "rgb":
        color_format.rgb = RGBColor.from_string(val)
    else:
        color_format.theme_color = THEME[val]
    if tint:
        bright = bright + (1 - bright) * tint if bright >= 0 else tint
    if bright:
        color_format.brightness = max(-1.0, min(1.0, bright))


def fill_theme(fill, name: str, brightness: float = 0.0) -> None:
    fill.solid()
    fill.fore_color.theme_color = THEME[name]
    if brightness:
        fill.fore_color.brightness = brightness


def drop_style(shape) -> None:
    """Remove the theme-referencing <p:style> python-pptx adds to autoshapes (white text etc.)."""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


BU_TAGS = {qn("a:buNone"), qn("a:buAutoNum"), qn("a:buChar"), qn("a:buBlip")}
PPR_ORDER = ["a:lnSpc", "a:spcBef", "a:spcAft", "a:buClrTx", "a:buClr", "a:buSzTx", "a:buSzPct",
             "a:buSzPts", "a:buFontTx", "a:buFont", "a:buNone", "a:buAutoNum", "a:buChar",
             "a:buBlip", "a:tabLst", "a:defRPr", "a:extLst"]


def pPr_insert(pPr, el) -> None:
    """Insert a child into <a:pPr> respecting the schema order."""
    rank = {qn(t): i for i, t in enumerate(PPR_ORDER)}
    r = rank.get(el.tag, 99)
    for i, child in enumerate(pPr):
        if rank.get(child.tag, 99) > r:
            pPr.insert(i, el)
            return
    pPr.append(el)


def set_bullet(paragraph, kind: str | None) -> None:
    """kind: None (inherit) | 'none' | 'number'."""
    if kind is None:
        return
    pPr = paragraph._p.get_or_add_pPr()
    for ch in list(pPr):
        if ch.tag in BU_TAGS:
            pPr.remove(ch)
    if kind == "none":
        pPr.set("marL", "0")
        pPr.set("indent", "0")
        pPr_insert(pPr, etree.SubElement(etree.Element("x"), qn("a:buNone")))
    elif kind == "number":
        el = etree.SubElement(etree.Element("x"), qn("a:buAutoNum"))
        el.set("type", "arabicPeriod")
        pPr_insert(pPr, el)


def copy_body_style(prs, text_frame) -> None:
    """Give a free text box the master's bullet/level styling so it looks like a placeholder."""
    body_style = prs.slide_master._element.find(qn("p:txStyles") + "/" + qn("p:bodyStyle"))
    txBody = text_frame._txBody
    old = txBody.find(qn("a:lstStyle"))
    new = etree.SubElement(etree.Element("x"), qn("a:lstStyle"))
    if body_style is not None:
        for child in body_style:
            if child.tag != qn("a:extLst"):
                new.append(copy.deepcopy(child))
    if old is not None:
        txBody.replace(old, new)
    else:
        txBody.insert(1, new)


def flatten_bullets(items, level=0):
    """YAML bullets -> [(level, text)]. A nested list = sub-items of the previous item.
    Dict form {text: ..., items: [...]} also works."""
    out = []
    for it in items or []:
        if isinstance(it, list):
            out += flatten_bullets(it, level + 1)
        elif isinstance(it, dict):
            out.append((level, str(it.get("text", ""))))
            out += flatten_bullets(it.get("items"), level + 1)
        else:
            out.append((level, "" if it is None else str(it)))
    return out


# --------------------------------------------------------------------------------------
# LaTeX -> Unicode fallback (used when pandoc is missing, and as the mc:Fallback text)
# --------------------------------------------------------------------------------------
_TEX_SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε", "varepsilon": "ε",
    "zeta": "ζ", "eta": "η", "theta": "θ", "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ",
    "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ", "phi": "φ",
    "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω", "Gamma": "Γ", "Delta": "Δ",
    "Theta": "Θ", "Lambda": "Λ", "Pi": "Π", "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
    "sum": "∑", "prod": "∏", "int": "∫", "oint": "∮", "infty": "∞", "partial": "∂", "nabla": "∇",
    "cdot": "·", "times": "×", "div": "÷", "pm": "±", "mp": "∓", "leq": "≤", "le": "≤",
    "geq": "≥", "ge": "≥", "neq": "≠", "ne": "≠", "approx": "≈", "equiv": "≡", "sim": "∼",
    "propto": "∝", "to": "→", "rightarrow": "→", "leftarrow": "←", "Rightarrow": "⇒",
    "Leftrightarrow": "⇔", "mapsto": "↦", "in": "∈", "notin": "∉", "subset": "⊂",
    "subseteq": "⊆", "cup": "∪", "cap": "∩", "forall": "∀", "exists": "∃", "emptyset": "∅",
    "ldots": "…", "cdots": "⋯", "dots": "…", "langle": "⟨", "rangle": "⟩", "sqrt": "√",
    "mathbb{R}": "ℝ", "mathbb{N}": "ℕ", "mathbb{Z}": "ℤ", "mathbb{Q}": "ℚ", "mathbb{C}": "ℂ",
    "mathbb{E}": "𝔼", "ell": "ℓ", "hbar": "ℏ", "degree": "°", "quad": "  ", "qquad": "    ",
}
_SUP = str.maketrans("0123456789+-=()niT", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱᵀ")
_SUB = str.maketrans("0123456789+-=()aeijknmx", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑᵢⱼₖₙₘₓ")


def latex_to_unicode(tex: str) -> str:
    """Readable plain-text approximation of LaTeX (for the fallback / no-pandoc case)."""
    s = tex.replace("\\mid", "|").replace("\\lvert", "|").replace("\\rvert", "|")
    for k in sorted(_TEX_SYMBOLS, key=len, reverse=True):
        s = re.sub(r"\\" + re.escape(k) + r"(?![A-Za-z])", _TEX_SYMBOLS[k], s)
    s = re.sub(r"\\(left|right|big|Big|bigg|Bigg)\b", "", s)
    s = re.sub(r"\\(mathrm|mathbf|mathit|text|operatorname|mathcal|boldsymbol)\{([^{}]*)\}", r"\2", s)

    def script(m, table, chars, mark):
        body = m.group(1) if m.group(1) is not None else m.group(2)
        if all(c in chars for c in body):
            return body.translate(table)
        return mark + (body if len(body) == 1 else "(" + body + ")")

    for _ in range(4):  # innermost groups first
        s = re.sub(r"√\{([^{}]*)\}", lambda m: "√" + (m.group(1) if len(m.group(1)) == 1 else "(" + m.group(1) + ")"), s)
        s = re.sub(r"\\[dt]?frac\{([^{}]*)\}\{([^{}]*)\}", r"(\1)/(\2)", s)
        s = re.sub(r"\^(?:\{([^{}]*)\}|([^{\s(]))", lambda m: script(m, _SUP, "0123456789+-=()niT", "^"), s)
        s = re.sub(r"_(?:\{([^{}]*)\}|([^{\s(]))", lambda m: script(m, _SUB, "0123456789+-=()aeijknmx", "_"), s)
    s = s.replace("\\,", " ").replace("\\;", " ").replace("\\!", "").replace("\\ ", " ")
    s = s.replace("{", "").replace("}", "").replace("\\", "")
    s = re.sub(r"(?<=[\s(=])-|^-", "−", s)
    return re.sub(r"\s+", " ", s).strip()


class MathEngine:
    """Converts LaTeX to OMML (native PowerPoint equations) with one batched pandoc call."""

    def __init__(self):
        self.pandoc = shutil.which("pandoc")
        if not self.pandoc:
            try:
                import pypandoc  # type: ignore
                self.pandoc = pypandoc.get_pandoc_path()
            except Exception:
                self.pandoc = None

    def convert(self, formulas):
        """formulas: list of (latex, display). Returns list of lxml elements (or None)."""
        if not formulas:
            return []
        if not self.pandoc:
            warn("pandoc not found - equations are written as plain text. "
                 "Install pandoc (or `pip install pypandoc_binary`) for native equations.")
            return [None] * len(formulas)
        md = "\n\n---\n\n".join((f"$${t}$$" if d else f"${t}$") for t, d in formulas)
        with tempfile.TemporaryDirectory() as td:
            src, dst = os.path.join(td, "m.md"), os.path.join(td, "m.pptx")
            with open(src, "w", encoding="utf-8") as fh:
                fh.write(md)
            try:
                subprocess.run([self.pandoc, src, "-o", dst], check=True, capture_output=True, timeout=120)
            except Exception as exc:  # noqa: BLE001
                warn(f"pandoc failed ({exc}); equations written as plain text.")
                return [None] * len(formulas)
            out = []
            with zipfile.ZipFile(dst) as z:
                for i in range(1, len(formulas) + 1):
                    try:
                        root = etree.fromstring(z.read(f"ppt/slides/slide{i}.xml"))
                        m = root.find(f".//{{{NS_A14}}}m")
                        out.append(copy.deepcopy(m[0]) if m is not None and len(m) else None)
                    except KeyError:
                        out.append(None)
        for (tex, _), el in zip(formulas, out):
            if el is None:
                warn(f"could not convert equation: {tex}")
        return out


# --------------------------------------------------------------------------------------
# Animation / transition XML
# --------------------------------------------------------------------------------------
PRESETS = {  # name -> (presetID, presetSubtype)
    "appear": (1, 0), "fade": (10, 0), "wipe": (22, 8), "fly": (2, 4), "fly-left": (2, 8),
    "zoom": (53, 16),
}


def _tgt(e: Effect) -> str:
    tx = f'<p:txEl><p:pRg st="{e.para[0]}" end="{e.para[1]}"/></p:txEl>' if e.para else ""
    return f'<p:tgtEl><p:spTgt spid="{e.spid}">{tx}</p:spTgt></p:tgtEl>'


def _anim(nid, e, attr, v0, v1, fv0=False):
    val0 = f'<p:fltVal val="{v0}"/>' if fv0 else f'<p:strVal val="{v0}"/>'
    return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            f'<p:cTn id="{next(nid)}" dur="{e.dur}" fill="hold"/>{_tgt(e)}'
            f'<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:tavLst><p:tav tm="0"><p:val>{val0}</p:val></p:tav>'
            f'<p:tav tm="100000"><p:val><p:strVal val="{v1}"/></p:val></p:tav></p:tavLst></p:anim>')


def _effect_xml(e: Effect, node: str, nid) -> str:
    name = e.effect if e.effect in PRESETS else "fade"
    pid, sub = PRESETS[name]
    ctn = next(nid)
    beh = (f'<p:set><p:cBhvr><p:cTn id="{next(nid)}" dur="1" fill="hold"><p:stCondLst>'
           f'<p:cond delay="0"/></p:stCondLst></p:cTn>{_tgt(e)}<p:attrNameLst>'
           f'<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
           f'<p:to><p:strVal val="visible"/></p:to></p:set>')
    if name == "fade":
        beh += (f'<p:animEffect transition="in" filter="fade"><p:cBhvr>'
                f'<p:cTn id="{next(nid)}" dur="{e.dur}"/>{_tgt(e)}</p:cBhvr></p:animEffect>')
    elif name == "wipe":
        beh += (f'<p:animEffect transition="in" filter="wipe(right)"><p:cBhvr>'
                f'<p:cTn id="{next(nid)}" dur="{e.dur}"/>{_tgt(e)}</p:cBhvr></p:animEffect>')
    elif name == "fly":
        beh += _anim(nid, e, "ppt_x", "#ppt_x", "#ppt_x") + _anim(nid, e, "ppt_y", "1+#ppt_h/2", "#ppt_y")
    elif name == "fly-left":
        beh += _anim(nid, e, "ppt_x", "0-#ppt_w/2", "#ppt_x") + _anim(nid, e, "ppt_y", "#ppt_y", "#ppt_y")
    elif name == "zoom":
        beh += (_anim(nid, e, "ppt_w", "0", "#ppt_w", fv0=True) + _anim(nid, e, "ppt_h", "0", "#ppt_h", fv0=True)
                + f'<p:animEffect transition="in" filter="fade"><p:cBhvr>'
                  f'<p:cTn id="{next(nid)}" dur="{e.dur}"/>{_tgt(e)}</p:cBhvr></p:animEffect>')
    return (f'<p:par><p:cTn id="{ctn}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" '
            f'fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{beh}</p:childTnLst></p:cTn></p:par>')


def timing_xml(steps, builds) -> str:
    """steps: [[group, ...], ...]; group = [Effect, ...]. builds: {spid: 'para'|'shape'|'graphic'}."""
    nid = itertools.count(3)
    pars = []
    for step in steps:
        sid = next(nid)
        groups, delay = [], 0
        for gi, group in enumerate(step):
            gid = next(nid)
            effs = []
            for ei, e in enumerate(group):
                node = "withEffect" if ei > 0 else ("clickEffect" if gi == 0 else "afterEffect")
                effs.append(_effect_xml(e, node, nid))
            groups.append(f'<p:par><p:cTn id="{gid}" fill="hold"><p:stCondLst><p:cond delay="{delay}"/>'
                          f'</p:stCondLst><p:childTnLst>{"".join(effs)}</p:childTnLst></p:cTn></p:par>')
            delay += max(e.dur for e in group)
        pars.append(f'<p:par><p:cTn id="{sid}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
                    f'</p:stCondLst><p:childTnLst>{"".join(groups)}</p:childTnLst></p:cTn></p:par>')
    bld = []
    for spid, kind in builds.items():
        if kind == "para":
            bld.append(f'<p:bldP spid="{spid}" grpId="0" build="p"/>')
        elif kind == "shape":
            bld.append(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>')
        elif kind == "graphic":
            bld.append(f'<p:bldGraphic spid="{spid}" grpId="0"><p:bldAsOne/></p:bldGraphic>')
    bld_xml = f'<p:bldLst>{"".join(bld)}</p:bldLst>' if bld else ""
    return (f'<p:timing {NSDECL}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
            f'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" '
            f'dur="indefinite" nodeType="mainSeq"><p:childTnLst>{"".join(pars)}</p:childTnLst></p:cTn>'
            f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
            f'</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
            f'</p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
            f'{bld_xml}</p:timing>')


TRANSITIONS = {
    "fade": "<p:fade/>", "push": '<p:push dir="u"/>', "wipe": '<p:wipe dir="r"/>',
    "cover": '<p:cover dir="l"/>', "dissolve": "<p:dissolve/>",
    "split": '<p:split orient="vert" dir="out"/>', "zoom": '<p:zoom/>',
}


def transition_xml(spec) -> str | None:
    if not spec or spec == "none":
        return None
    if isinstance(spec, dict):
        name, speed = spec.get("type", "fade"), spec.get("speed", "med")
    else:
        name, speed = str(spec), "med"
    if name == "morph":
        return (f'<mc:AlternateContent xmlns:mc="{NS_MC}" {NSDECL}><mc:Choice '
                f'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
                f'<p:transition spd="slow"><p159:morph option="byObject"/></p:transition></mc:Choice>'
                f'<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>'
                f'</mc:AlternateContent>')
    if name not in TRANSITIONS:
        warn(f"unknown transition '{name}', using fade")
        name = "fade"
    return f'<p:transition {NSDECL} spd="{speed}">{TRANSITIONS[name]}</p:transition>'


# --------------------------------------------------------------------------------------
# Template helpers
# --------------------------------------------------------------------------------------
def open_template(path: str):
    with open(path, "rb") as fh:
        data = fh.read()
    if path.lower().endswith((".potx", ".potm")):  # python-pptx only opens presentations
        zin = zipfile.ZipFile(io.BytesIO(data))
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                blob = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    blob = blob.replace(b"presentationml.template.main+xml",
                                        b"presentationml.presentation.main+xml")
                zout.writestr(item, blob)
        data = buf.getvalue()
    return Presentation(io.BytesIO(data))


def remove_all_slides(prs) -> None:
    lst = prs.slides._sldIdLst
    for sldId in list(lst):
        prs.part.drop_rel(sldId.rId)
        lst.remove(sldId)


def iter_shapes_deep(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type is not None and hasattr(sh, "shapes"):
            yield from iter_shapes_deep(sh.shapes)


def replace_tokens(prs, tokens: dict) -> None:
    """Replace {author}, {date}, ... in master and layout shapes (the footline)."""
    pat = re.compile(r"\{(" + "|".join(map(re.escape, tokens)) + r")\}")
    containers = [prs.slide_master] + list(prs.slide_layouts)
    for cont in containers:
        for sh in iter_shapes_deep(cont.shapes):
            if not sh.has_text_frame or sh.is_placeholder:
                continue
            for p in sh.text_frame.paragraphs:
                if not pat.search("".join(r.text for r in p.runs)):
                    continue
                runs = p.runs
                if not any(pat.search(r.text) for r in runs):  # token split over runs: merge
                    runs[0].text = "".join(r.text for r in runs)
                    for r in runs[1:]:
                        r._r.getparent().remove(r._r)
                    runs = p.runs
                for r in runs:
                    r.text = pat.sub(lambda m: str(tokens[m.group(1)]), r.text)


def style_sizes(prs):
    """Font sizes (pt) of body levels 1-4 and of the title, from the slide master."""
    ts = prs.slide_master._element.find(qn("p:txStyles"))
    sizes = []
    for lvl in range(1, 5):
        el = ts.find(f"{qn('p:bodyStyle')}/{qn(f'a:lvl{lvl}pPr')}/{qn('a:defRPr')}") if ts is not None else None
        sizes.append(int(el.get("sz")) / 100 if el is not None and el.get("sz") else [24, 20, 18, 16][lvl - 1])
    t = ts.find(f"{qn('p:titleStyle')}/{qn('a:lvl1pPr')}/{qn('a:defRPr')}") if ts is not None else None
    title = int(t.get("sz")) / 100 if t is not None and t.get("sz") else 32
    return sizes, title


# --------------------------------------------------------------------------------------
# The deck builder
# --------------------------------------------------------------------------------------
class DeckBuilder:
    def __init__(self, content: dict, template: str | None = None, base_dir: str = ".",
                 substitute_tokens: bool = True, keep_template_slides: bool = False):
        self.meta = content.get("meta", {}) or {}
        self.slides_spec = content.get("slides", []) or []
        self.base_dir = base_dir
        tpl = template or self.meta.get("template") or DEFAULT_TEMPLATE
        if not os.path.isabs(tpl) and not os.path.exists(tpl):
            cand = os.path.join(base_dir, tpl)
            tpl = cand if os.path.exists(cand) else os.path.join(HERE, tpl)
        self.template_path = tpl
        self.prs = open_template(tpl)
        if not keep_template_slides:
            remove_all_slides(self.prs)
        self.layouts = self._resolve_layouts(self.meta.get("layouts", {}))
        self.level_sizes, self.title_size = style_sizes(self.prs)
        self.W, self.H = to_in(self.prs.slide_width), to_in(self.prs.slide_height)
        self.area = self._content_area()
        self.math_jobs = []       # (latex, display, run-properties element)
        self.valign = self.meta.get("valign", "center")
        self.sections = [s.get("title", "") for s in self.slides_spec
                         if isinstance(s, dict) and s.get("layout") == "section"]
        self.section_no = 0
        self.substitute_tokens = substitute_tokens
        self._flow_cache = {}

    # ---- template introspection -------------------------------------------------------
    def _resolve_layouts(self, overrides):
        by_name = {l.name.strip().lower(): l for l in self.prs.slide_layouts}
        by_type = {}
        for l in self.prs.slide_layouts:
            by_type.setdefault(l._element.get("type"), l)
        out = {}
        for key, (names, typ) in DEFAULT_LAYOUTS.items():
            want = overrides.get(key)
            cands = ([want] if isinstance(want, str) else list(want or [])) + names
            lay = next((by_name[n.lower()] for n in cands if n and n.lower() in by_name), None)
            out[key] = lay or by_type.get(typ)
        if out["title_only"] is None:
            out["title_only"] = out["frame"]
        if out["plain"] is None:
            out["plain"] = out["title_only"]
        missing = [k for k, v in out.items() if v is None]
        if missing:
            raise SystemExit(f"Template {self.template_path} has no layout for: {', '.join(missing)}. "
                             f"Available: {[l.name for l in self.prs.slide_layouts]}. "
                             f"Map them with meta.layouts in the content file.")
        return out

    def _content_area(self) -> Box:
        for ph in self.layouts["frame"].placeholders:
            if ph.placeholder_format.idx == 1 or ph.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
                if ph.width:
                    return Box(to_in(ph.left), to_in(ph.top), to_in(ph.width), to_in(ph.height))
        return Box(0.6, 1.3, self.W - 1.2, self.H - 2.2)

    @staticmethod
    def _layout_ph_idx(layout, name):
        for ph in layout.placeholders:
            if ph.name.strip().lower() == name.lower():
                return ph.placeholder_format.idx
        return None

    # ---- public ------------------------------------------------------------------------
    def build(self):
        today = _dt.date.today().strftime("%B %d, %Y").replace(" 0", " ")
        m = self.meta
        date = m.get("date", today)
        if str(date).lower() == "today":
            date = today
        self.meta["date"] = str(date)
        if self.substitute_tokens:
            replace_tokens(self.prs, {
                "title": m.get("title", ""), "short_title": m.get("short_title", m.get("title", "")),
                "subtitle": m.get("subtitle", ""), "author": m.get("author", ""),
                "short_author": m.get("short_author", m.get("author", "")),
                "institute": m.get("institute", ""), "short_institute": m.get("short_institute", m.get("institute", "")),
                "date": self.meta["date"],
            })
        cp = self.prs.core_properties
        cp.title, cp.author = str(m.get("title", "")), str(m.get("author", ""))

        slides = []
        for i, spec in enumerate(self.slides_spec, 1):
            if not isinstance(spec, dict):
                warn(f"slide {i} is not a mapping - skipped")
                continue
            try:
                slides.append((self._build_slide(spec), spec))
            except Exception as exc:  # keep going, report clearly
                raise SystemExit(f"Error on slide {i} ({spec.get('title', spec.get('layout', ''))}): {exc}")
        self._finalize_math([s for s, _ in slides])
        return self.prs

    # ---- slides ------------------------------------------------------------------------
    def _build_slide(self, spec):
        kind = spec.get("layout", "frame")
        self.steps, self.builds = [], {}
        if kind == "title":
            slide = self._title_slide(spec)
        elif kind == "section":
            self.section_no += 1
            if self.meta.get("section_style") == "outline":
                slide = self._outline_slide(spec, current=self.section_no, title=self.meta.get("outline_title", "Outline"))
            else:
                slide = self._section_slide(spec)
        elif kind == "outline":
            slide = self._outline_slide(spec, current=None, title=spec.get("title", "Outline"))
        elif kind == "plain":
            slide = self.prs.slides.add_slide(self.layouts["plain"])
            self._remove_empty_placeholders(slide, keep=set())
            m = 0.6
            self._render_items(slide, spec.get("content", []), Box(m, m, self.W - 2 * m, self.H - 2 * m),
                               spec.get("valign", "center"))
        else:
            slide = self._frame_slide(spec)
        if spec.get("notes"):
            slide.notes_slide.notes_text_frame.text = str(spec["notes"]).strip()
        self._apply_timing(slide, spec.get("transition", self.meta.get("transition")))
        return slide

    def _set_title(self, slide, text, subtitle=None, fit_width=None):
        ph = slide.shapes.title
        if ph is None:
            return
        tf = ph.text_frame
        tf.clear()
        size = self.title_size
        width = fit_width or to_in(ph.width) or self.W - 1
        lines = est_lines(text, width - 0.2, size)
        if lines > 1:  # beamer-like: shrink long frame titles instead of overflowing the bar
            size = max(size * 0.72, size / lines ** 0.5)
        p = tf.paragraphs[0]
        self._add_runs(p, str(text), {"size": size if size != self.title_size else None})
        if subtitle:
            p2 = tf.add_paragraph()
            self._add_runs(p2, str(subtitle), {"size": round(size * 0.62, 1), "bold": False})

    def _remove_empty_placeholders(self, slide, keep):
        for ph in list(slide.placeholders):
            if ph.placeholder_format.idx not in keep:
                ph._element.getparent().remove(ph._element)

    def _title_slide(self, spec):
        lay = self.layouts["title"]
        slide = self.prs.slides.add_slide(lay)
        m = {**self.meta, **{k: v for k, v in spec.items() if k != "layout"}}
        keep = set()
        title_ph = next((p for p in slide.placeholders if p.placeholder_format.type in (PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE)), None)
        if title_ph is not None:
            title_ph.text_frame.clear()
            size = None
            lines = est_lines(str(m.get("title", "")), to_in(title_ph.width) - 0.5, 36)
            if lines > 2:
                size = 28
            self._add_runs(title_ph.text_frame.paragraphs[0], str(m.get("title", "")), {"size": size})
            keep.add(title_ph.placeholder_format.idx)
        extra = []
        for field_name in ("author", "institute", "date"):
            idx = self._layout_ph_idx(lay, field_name)
            val = m.get(field_name)
            if not val:
                continue
            if idx is not None:
                ph = slide.placeholders[idx]
                self._fill_text(ph.text_frame, str(val))
                keep.add(idx)
            else:
                extra.append(str(val))
        sub_ph = next((p for p in slide.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE), None)
        sub_lines = ([str(m["subtitle"])] if m.get("subtitle") else []) + extra
        if sub_ph is not None and sub_lines:
            self._fill_text(sub_ph.text_frame, "\n".join(sub_lines))
            keep.add(sub_ph.placeholder_format.idx)
        self._remove_empty_placeholders(slide, keep)
        return slide

    def _section_slide(self, spec):
        slide = self.prs.slides.add_slide(self.layouts["section"])
        keep = set()
        if slide.shapes.title is not None:
            slide.shapes.title.text_frame.clear()
            self._add_runs(slide.shapes.title.text_frame.paragraphs[0], str(spec.get("title", "")), {})
            keep.add(slide.shapes.title.placeholder_format.idx)
        sub = spec.get("subtitle")
        if sub:
            body = next((p for p in slide.placeholders if p.placeholder_format.type not in
                         (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)), None)
            if body is not None:
                self._fill_text(body.text_frame, str(sub))
                keep.add(body.placeholder_format.idx)
        self._remove_empty_placeholders(slide, keep)
        return slide

    def _outline_slide(self, spec, current, title):
        slide = self.prs.slides.add_slide(self.layouts["frame"])
        self._set_title(slide, title)
        body = slide.placeholders[1] if 1 in [p.placeholder_format.idx for p in slide.placeholders] else None
        if body is None:
            return slide
        tf = body.text_frame
        tf.clear()
        for i, sec in enumerate(self.sections):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            set_bullet(p, "number")
            style = {}
            if current is not None:
                style = {"bold": True} if i + 1 == current else {"color": "muted"}
            self._add_runs(p, str(sec), style)
        self._remove_empty_placeholders(slide, {0, 1, body.placeholder_format.idx,
                                               slide.shapes.title.placeholder_format.idx if slide.shapes.title else 0})
        return slide

    def _frame_slide(self, spec):
        items = [self._norm_item(i) for i in (spec.get("content") or [])]
        valign = spec.get("valign", self.valign)
        # 1) a single bullet list -> the real body placeholder of "Title and Content"
        if len(items) == 1 and items[0][0] in ("bullets", "numbered") and not items[0][1].get("size"):
            slide = self.prs.slides.add_slide(self.layouts["frame"])
            self._set_title(slide, spec.get("title", ""), spec.get("subtitle"))
            body = next((p for p in slide.placeholders if p.placeholder_format.idx != slide.shapes.title.placeholder_format.idx), None) if slide.shapes.title else None
            if body is not None:
                self._fill_placeholder_bullets(body, items[0], spec.get("valign"))
                self._remove_empty_placeholders(slide, {slide.shapes.title.placeholder_format.idx, body.placeholder_format.idx})
                return slide
        # 2) two columns that are each one bullet list -> "Two Content" placeholders
        if (len(items) == 1 and items[0][0] == "columns" and self.layouts["columns"] is not None):
            cols = self._norm_columns(items[0][1]["columns"])
            simple = (len(cols) == 2 and all(len(c["content"]) == 1 and c["content"][0][0] in ("bullets", "numbered")
                                             and not c.get("width") for c in cols))
            if simple:
                slide = self.prs.slides.add_slide(self.layouts["columns"])
                self._set_title(slide, spec.get("title", ""), spec.get("subtitle"))
                tidx = slide.shapes.title.placeholder_format.idx
                bodies = sorted([p for p in slide.placeholders if p.placeholder_format.idx != tidx], key=lambda p: p.left)
                if len(bodies) >= 2:
                    for ph, col in zip(bodies, cols):
                        self._fill_placeholder_bullets(ph, col["content"][0], spec.get("valign"))
                    self._remove_empty_placeholders(slide, {tidx, bodies[0].placeholder_format.idx, bodies[1].placeholder_format.idx})
                    return slide
                self._delete_slide(slide)
        # 3) everything else -> "Title Only" + generated native shapes in the content area
        slide = self.prs.slides.add_slide(self.layouts["title_only"])
        self._set_title(slide, spec.get("title", ""), spec.get("subtitle"))
        keep = {slide.shapes.title.placeholder_format.idx} if slide.shapes.title else set()
        self._remove_empty_placeholders(slide, keep)
        self._render_items(slide, items, self.area, valign)
        return slide

    def _delete_slide(self, slide):
        lst = self.prs.slides._sldIdLst
        for sldId in list(lst):
            if self.prs.part.related_part(sldId.rId) is slide.part:
                self.prs.part.drop_rel(sldId.rId)
                lst.remove(sldId)

    # ---- items: normalisation & measuring ---------------------------------------------
    def _norm_item(self, item):
        if isinstance(item, tuple):
            return item
        if isinstance(item, str):
            return ("text", {"text": item})
        if isinstance(item, list):
            return ("bullets", {"bullets": item})
        if not isinstance(item, dict):
            raise ValueError(f"cannot understand content item: {item!r}")
        for t in ITEM_TYPES:
            if t in item:
                return (t, item)
        raise ValueError(f"content item needs one of {', '.join(ITEM_TYPES)}: {item!r}")

    def _norm_columns(self, cols):
        out = []
        for c in cols:
            if isinstance(c, list):
                c = {"content": c}
            out.append({**c, "content": [self._norm_item(i) for i in (c.get("content") or [])]})
        return out

    def _sz(self, level, scale, opts):
        base = opts.get("size") or self.level_sizes[min(level, 3)]
        if opts.get("size") and level:
            base = opts["size"] * self.level_sizes[min(level, 3)] / self.level_sizes[0]
        return round(base * scale, 1)

    def _text_height(self, paras, width, scale, opts, indent=True):
        h = 0.0
        for lvl, text in paras:
            size = self._sz(lvl, scale, opts)
            avail = width - (0.36 * (lvl + 1) if indent else 0) - 0.05
            h += est_lines(text, avail, size) * size * 1.2 / 72 + (8 if lvl == 0 else 4) / 72
        return h + 0.12

    def _measure(self, kind, it, w, scale):
        """-> (height, flexible, min_height)."""
        if kind in ("bullets", "numbered"):
            return self._text_height(flatten_bullets(it[kind]), w, scale, it), False, 0
        if kind == "text":
            paras = [(0, s) for s in str(it["text"]).split("\n")]
            return self._text_height(paras, w, scale, it, indent=False), False, 0
        if kind in BLOCK_COLORS:
            b = it[kind] if isinstance(it[kind], dict) else {"text": str(it[kind])}
            size = (b.get("size") or self.level_sizes[0] * 0.9) * scale
            head = size * 1.1 * 1.35 / 72 + 0.12
            if b.get("bullets"):
                body = self._text_height(flatten_bullets(b["bullets"]), w - 0.3, scale, {"size": size / scale})
            else:
                paras = [(0, s) for s in str(b.get("text", "")).split("\n")]
                body = self._text_height(paras, w - 0.3, scale, {"size": size / scale}, indent=False)
            return head + body + 0.12, False, 0
        if kind == "chart":
            c = it["chart"]
            pref = float(c.get("height", 4.2))
            return pref, True, min(pref, 2.2)
        if kind == "table":
            t = it["table"]
            size = (t.get("size") or self.level_sizes[0] * 0.78) * scale
            rows = len(t.get("rows", [])) + (1 if t.get("header") else 0)
            return rows * (size * 2.05 / 72), False, 0
        if kind == "plot":
            p, ratio = self._figure(kind, it)
            if ratio is None:
                cap = 0.4 if p.get("caption") else 0
                pref = float(p.get("height", 4.2)) + cap
                return pref, True, min(pref, 2.2 + cap)
        if kind in ("image", "placeholder", "plot"):
            im, ratio = self._figure(kind, it)
            width = w * float(im.get("width", 1.0 if kind == "plot" else 0.75))
            cap = 0.4 if im.get("caption") else 0
            pref = float(im.get("height", width / ratio)) + cap
            return pref, True, min(pref, 1.6 + cap)
        if kind in ("flow", "steps"):
            lay, spec = self._flow(kind, it)
            f = self._flow_font(spec, scale)
            cap = 0.4 if spec.get("caption") else 0
            k = min(1.0, w * float(spec.get("width", 1.0)) / (lay.w * f / 72))
            pref = lay.h * f / 72 * k
            if spec.get("height"):
                pref = min(pref, float(spec["height"]))
            return pref + cap, True, min(pref, lay.h * FLOW_MIN_PT / 72) + cap
        if kind == "math":
            tex = str(it["math"])
            tall = any(k in tex for k in ("\\frac", "\\sum", "\\int", "\\prod", "\\begin", "\\sqrt", "\\over"))
            return ((1.05 if tall else 0.7) * scale * self.level_sizes[0] / 22), False, 0
        if kind == "code":
            n = len(str(it["code"]).rstrip("\n").split("\n"))
            size = (it.get("size") or self.level_sizes[0] * 0.7) * scale
            return n * size * 1.22 / 72 + 0.3, False, 0
        if kind == "spacer":
            return float(it["spacer"]), False, 0
        if kind == "columns":
            cols = self._norm_columns(it["columns"])
            widths = self._col_widths(cols, w)
            pref, mins, flex = 0.0, 0.0, False
            for c, cw in zip(cols, widths):
                ms = [self._measure(k, i, cw, scale) for k, i in c["content"]]
                gaps = GAP * max(0, len(ms) - 1)
                pref = max(pref, sum(m[0] for m in ms) + gaps)
                mins = max(mins, sum((m[2] if m[1] else m[0]) for m in ms) + gaps)
                flex = flex or any(m[1] for m in ms)
            return pref, flex, mins
        raise ValueError(kind)

    def _col_widths(self, cols, w):
        avail = w - COL_GAP * (len(cols) - 1)
        given = [c.get("width") for c in cols]
        fixed = sum(float(g) for g in given if g)
        free = max(0.0, 1.0 - fixed) / max(1, sum(1 for g in given if not g))
        return [avail * (float(g) if g else free) for g in given]

    def _figure(self, kind, it):
        """(spec, width/height ratio) of an image or placeholder item; both accept a short form."""
        v = it[kind]
        if kind == "image":
            spec = v if isinstance(v, dict) else {"path": v}
            return spec, self._image_ratio(spec)
        if kind == "plot":
            if not isinstance(v, dict):
                raise ValueError("plot needs settings, e.g. {type: bar, categories: [...], series: {...}}")
            return v, (parse_aspect(v["aspect"]) if v.get("aspect") else None)
        spec = v if isinstance(v, dict) else {"label": v}
        return spec, parse_aspect(spec.get("aspect", "16:9"))

    def _image_ratio(self, im):
        path = self._path(im["path"])
        try:
            from PIL import Image
            with Image.open(path) as img:
                return img.width / img.height
        except Exception:
            return 16 / 9

    def _path(self, p):
        return p if os.path.isabs(p) else os.path.join(self.base_dir, p)

    def _plan(self, items, box, scale):
        """Heights for a vertical stack. Returns (heights, total)."""
        ms = [self._measure(k, i, box.w, scale) for k, i in items]
        gaps = GAP * max(0, len(items) - 1)
        fixed = sum(m[0] for m in ms if not m[1])
        flex_pref = sum(m[0] for m in ms if m[1])
        free = box.h - fixed - gaps
        heights = []
        for m in ms:
            if not m[1]:
                heights.append(m[0])
            elif flex_pref <= free or flex_pref == 0:
                heights.append(m[0])
            else:
                heights.append(max(m[2], m[0] * free / flex_pref))
        return heights, sum(heights) + gaps

    # ---- items: rendering ---------------------------------------------------------------
    def _render_items(self, slide, items, box, valign="center", scale=None):
        items = [self._norm_item(i) for i in items]
        if not items:
            return
        if scale is None:
            for scale in (1.0, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65):
                heights, total = self._plan(items, box, scale)
                if total <= box.h + 0.02:
                    break
            else:
                title = slide.shapes.title.text_frame.text if slide.shapes.title else "(untitled)"
                warn(f"content does not fit on slide '{title}' - consider splitting it")
        heights, total = self._plan(items, box, scale)
        y = box.y + (max(0.0, box.h - total) / 2 if valign in ("center", "middle") else 0)
        for (kind, it), h in zip(items, heights):
            self._render(slide, kind, it, Box(box.x, y, box.w, h), scale)
            y += h + GAP

    def _render(self, slide, kind, it, b, scale):
        self._last_main = None
        if kind in ("bullets", "numbered"):
            shp = slide.shapes.add_textbox(inch(b.x), inch(b.y), inch(b.w), inch(b.h))
            self._prep_tf(shp.text_frame)
            copy_body_style(self.prs, shp.text_frame)
            groups = self._write_bullets(shp.text_frame, flatten_bullets(it[kind]), scale, it,
                                         numbered=(kind == "numbered"))
            self._animate(it, shp, "sp_text", groups)
        elif kind == "text":
            shp = slide.shapes.add_textbox(inch(b.x), inch(b.y), inch(b.w), inch(b.h))
            self._prep_tf(shp.text_frame)
            copy_body_style(self.prs, shp.text_frame)
            self._fill_text(shp.text_frame, str(it["text"]), size=self._sz(0, scale, it), style=it,
                            align=it.get("align"))
            self._animate(it, shp, "sp_text")
        elif kind in BLOCK_COLORS:
            self._render_block(slide, kind, it, b, scale)
        elif kind == "chart":
            self._render_chart(slide, it, b, scale)
        elif kind == "table":
            self._render_table(slide, it, b, scale)
        elif kind == "image":
            self._render_image(slide, it, b, scale)
        elif kind == "placeholder":
            self._render_placeholder(slide, it, b, scale)
        elif kind == "plot":
            self._render_plot(slide, it, b, scale)
        elif kind in ("flow", "steps"):
            self._render_flow(slide, kind, it, b, scale)
        elif kind == "math":
            shp = slide.shapes.add_textbox(inch(b.x), inch(b.y), inch(b.w), inch(b.h))
            self._prep_tf(shp.text_frame, anchor=MSO_ANCHOR.MIDDLE)
            p = shp.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            self._math_run(p, str(it["math"]), {"size": self._sz(0, scale, it) * 1.05}, display=True)
            self._animate(it, shp, "sp_text")
        elif kind == "code":
            self._render_code(slide, it, b, scale)
        elif kind == "columns":
            cols = self._norm_columns(it["columns"])
            widths = self._col_widths(cols, b.w)
            x = b.x
            for c, cw in zip(cols, widths):
                self._render_items(slide, c["content"], Box(x, b.y, cw, b.h), c.get("valign", "center"), scale)
                x += cw + COL_GAP
        if it.get("name") and kind not in ("columns", "spacer") and self._last_main is not None:
            self._last_main.name = "!!" + str(it["name"])  # "!!" prefix = Morph matches objects by name

    def _prep_tf(self, tf, anchor=MSO_ANCHOR.TOP):
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = inch(0.02)
        tf.margin_top = tf.margin_bottom = inch(0.04)

    def _write_bullets(self, tf, paras, scale, opts, numbered=False, explicit_size=True):
        """Writes paragraphs; returns [(first_para, last_para), ...] per top-level item."""
        groups = []
        tf.clear()
        for i, (lvl, text) in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.level = min(lvl, 8)
            if numbered and lvl == 0:
                set_bullet(p, "number")
            style = {"size": self._sz(lvl, scale, opts) if explicit_size else None}
            if opts.get("color"):
                style["color"] = opts["color"]
            self._add_runs(p, text, style)
            if lvl == 0 or not groups:
                groups.append([i, i])
            else:
                groups[-1][1] = i
        return [tuple(g) for g in groups]

    def _fill_placeholder_bullets(self, ph, item, valign=None):
        kind, it = item
        paras = flatten_bullets(it[kind])
        tf = ph.text_frame
        groups = self._write_bullets(tf, paras, 1.0, it, numbered=(kind == "numbered"), explicit_size=False)
        need = self._text_height(paras, to_in(ph.width), 1.0, it)
        have = to_in(ph.height)
        if need > have:  # same thing PowerPoint's "shrink text on overflow" stores in the file
            scale = max(0.55, have / need)
            bp = tf._txBody.find(qn("a:bodyPr"))
            for ch in list(bp):
                if ch.tag in (qn("a:normAutofit"), qn("a:spAutoFit"), qn("a:noAutofit")):
                    bp.remove(ch)
            af = etree.Element(qn("a:normAutofit"))
            af.set("fontScale", str(int(scale * 100000)))
            af.set("lnSpcReduction", "10000" if scale < 0.9 else "0")
            warp = bp.find(qn("a:prstTxWarp"))
            bp.insert(1 if warp is not None else 0, af)
        if valign:
            tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "center": MSO_ANCHOR.MIDDLE,
                                  "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[valign]
        self._animate(it, ph, "sp_text", groups)

    def _fill_text(self, tf, text, size=None, style=None, align=None):
        style = dict(style or {})
        tf.clear()
        for i, line in enumerate(str(text).split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            if style.get("_nobullet", True):
                set_bullet(p, "none")
            if align:
                p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT,
                               "justify": PP_ALIGN.JUSTIFY}[align]
            self._add_runs(p, line, {"size": size, "color": style.get("color"),
                                     "bold": style.get("bold"), "italic": style.get("italic")})

    # inline formatting: **bold**  *italic*  `code`  ==alert==  [link](url)  $math$
    INLINE = re.compile(r"\*\*(?P<b>.+?)\*\*|==(?P<alert>.+?)==|(?<![\w*])\*(?P<i>[^*\s][^*]*?)\*(?![\w*])"
                        r"|`(?P<code>[^`]+)`|\[(?P<lt>[^\]]+)\]\((?P<href>[^)\s]+)\)"
                        r"|(?<![\\$])\$(?P<math>[^\s$](?:[^$]*[^\s$\\])?)\$(?!\d)")

    def _add_runs(self, p, text, style):
        text = str(text)
        pos = 0
        for m in self.INLINE.finditer(text):
            if m.start() > pos:
                self._run(p, text[pos:m.start()], style)
            g = m.groupdict()
            if g["b"] is not None:
                self._add_runs(p, g["b"], {**style, "bold": True})
            elif g["alert"] is not None:
                self._add_runs(p, g["alert"], {**style, "color": "alert"})
            elif g["i"] is not None:
                self._add_runs(p, g["i"], {**style, "italic": True})
            elif g["code"] is not None:
                self._run(p, g["code"], {**style, "font": "Courier New"})
            elif g["lt"] is not None:
                self._run(p, g["lt"], {**style, "href": g["href"]})
            elif g["math"] is not None:
                self._math_run(p, g["math"], style, display=False)
            pos = m.end()
        if pos < len(text) or not text:
            self._run(p, text[pos:], style)

    def _run(self, p, text, style):
        text = text.replace("\\$", "$")
        r = p.add_run()
        r.text = text
        f = r.font
        if style.get("size"):
            f.size = Pt(style["size"])
        if style.get("bold") is not None:
            f.bold = style["bold"]
        if style.get("italic"):
            f.italic = True
        if style.get("font"):
            f.name = style["font"]
        if style.get("color"):
            set_color(f, style["color"])
        if style.get("href"):
            r.hyperlink.address = style["href"]
        return r

    def _math_run(self, p, tex, style, display):
        n = len(self.math_jobs)
        r = self._run(p, f"{MARK}{n}{MARK}", {**style, "font": "Cambria Math"})
        self.math_jobs.append((tex, display))
        return r

    # ---- blocks -------------------------------------------------------------------------
    def _render_block(self, slide, kind, it, b, scale):
        spec = it[kind] if isinstance(it[kind], dict) else {"text": str(it[kind])}
        color = spec.get("color", BLOCK_COLORS[kind])
        size = round((spec.get("size") or self.level_sizes[0] * 0.9) * scale, 1)
        head_h = size * 1.1 * 1.35 / 72 + 0.12
        radius = 0.09
        body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, inch(b.x), inch(b.y), inch(b.w), inch(b.h))
        drop_style(body)
        body.adjustments[0] = min(0.5, radius / min(b.w, b.h))
        body.fill.solid()
        paint(body.fill.fore_color, color, 0.85)
        body.line.fill.background()
        tf = body.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = tf.margin_right = inch(0.15)
        tf.margin_top, tf.margin_bottom = inch(head_h + 0.08), inch(0.08)
        copy_body_style(self.prs, tf)
        if spec.get("bullets"):
            self._write_bullets(tf, flatten_bullets(spec["bullets"]), 1.0, {"size": size})
            for p in tf.paragraphs:
                p.alignment = PP_ALIGN.LEFT
        else:
            self._fill_text(tf, str(spec.get("text", "")), size=size, style={"color": "tx1"})
            for p in tf.paragraphs:
                p.alignment = PP_ALIGN.LEFT
        head = slide.shapes.add_shape(MSO_SHAPE.ROUND_2_SAME_RECTANGLE, inch(b.x), inch(b.y), inch(b.w), inch(head_h))
        drop_style(head)
        head.adjustments[0] = min(0.5, radius / min(b.w, head_h))
        head.adjustments[1] = 0.0
        head.fill.solid()
        paint(head.fill.fore_color, color)
        head.line.fill.background()
        htf = head.text_frame
        htf.word_wrap = True
        htf.vertical_anchor = MSO_ANCHOR.MIDDLE
        htf.margin_left = htf.margin_right = inch(0.15)
        htf.margin_top = htf.margin_bottom = inch(0.02)
        hp = htf.paragraphs[0]
        hp.alignment = PP_ALIGN.LEFT
        self._add_runs(hp, str(spec.get("title", "")), {"size": size * 1.05, "color": "white", "bold": True})
        grp = slide.shapes.add_group_shape([body, head])
        grp.name = f"{kind.capitalize()}: {strip_markup(spec.get('title', ''))[:40]}"
        self._animate(it, grp, "group")

    # ---- code ---------------------------------------------------------------------------
    def _render_code(self, slide, it, b, scale):
        size = round((it.get("size") or self.level_sizes[0] * 0.7) * scale, 1)
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, inch(b.x), inch(b.y), inch(b.w), inch(b.h))
        drop_style(shp)
        shp.adjustments[0] = min(0.5, 0.06 / min(b.w, b.h))
        fill_theme(shp.fill, "tx1", 0.95)
        shp.line.fill.background()
        shp.name = "Code"
        tf = shp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = inch(0.2)
        tf.margin_top = tf.margin_bottom = inch(0.1)
        for i, line in enumerate(str(it["code"]).rstrip("\n").split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = line if line else " "
            r._r.find(qn("a:t")).set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.font.name, r.font.size = "Courier New", Pt(size)
            set_color(r.font, "tx1")
        self._animate(it, shp, "sp_text")

    # ---- charts -------------------------------------------------------------------------
    CHART_TYPES = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED, "column-stacked": XL_CHART_TYPE.COLUMN_STACKED,
        "column-stacked100": XL_CHART_TYPE.COLUMN_STACKED_100, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "bar-stacked": XL_CHART_TYPE.BAR_STACKED, "bar-stacked100": XL_CHART_TYPE.BAR_STACKED_100,
        "line": XL_CHART_TYPE.LINE_MARKERS, "line-plain": XL_CHART_TYPE.LINE,
        "area": XL_CHART_TYPE.AREA, "area-stacked": XL_CHART_TYPE.AREA_STACKED,
        "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT,
        "scatter": XL_CHART_TYPE.XY_SCATTER, "scatter-lines": XL_CHART_TYPE.XY_SCATTER_LINES,
        "radar": XL_CHART_TYPE.RADAR_MARKERS,
    }

    def _render_chart(self, slide, it, b, scale):
        c = it["chart"]
        ctype = str(c.get("type", "column"))
        if c.get("stacked") and ctype in ("column", "bar", "area"):
            ctype += "-stacked"
        if ctype not in self.CHART_TYPES:
            raise ValueError(f"unknown chart type '{ctype}'. Use one of: {', '.join(self.CHART_TYPES)}")
        xl = self.CHART_TYPES[ctype]
        series = c.get("series", {})
        if isinstance(series, list):  # [{name:, values:}]
            series = {s.get("name", f"Series {i+1}"): s.get("values", s.get("points")) for i, s in enumerate(series)}
        if ctype.startswith("scatter"):
            data = XyChartData()
            for name, pts in series.items():
                s = data.add_series(str(name))
                for x, y in pts:
                    s.add_data_point(x, y)
        else:
            data = CategoryChartData()
            data.categories = [str(x) for x in c.get("categories", [])]
            for name, vals in series.items():
                data.add_series(str(name), vals, number_format=c.get("number_format"))
        w = b.w * float(c.get("width", 1.0))
        x = b.x + (b.w - w) / 2
        gf = slide.shapes.add_chart(xl, inch(x), inch(b.y), inch(w), inch(b.h), data)
        gf.name = f"Chart: {c.get('title', ctype)}"
        chart = gf.chart
        fs = round(self.level_sizes[0] * 0.66 * scale, 1)
        chart.font.size = Pt(fs)
        set_color(chart.font, ("tx1", 0.25))
        if c.get("title"):
            chart.has_title = True
            chart.chart_title.text_frame.text = str(c["title"])
            tr = chart.chart_title.text_frame.paragraphs[0].runs[0].font
            tr.size, tr.bold = Pt(fs * 1.15), True
            set_color(tr, "structure")
        else:
            chart.has_title = False
        pie = ctype in ("pie", "doughnut")
        n_series = len(series)
        legend = c.get("legend", "bottom" if (n_series > 1 or pie) else "none")
        if legend and legend != "none":
            chart.has_legend = True
            chart.legend.position = {"bottom": XL_LEGEND_POSITION.BOTTOM, "right": XL_LEGEND_POSITION.RIGHT,
                                     "top": XL_LEGEND_POSITION.TOP, "left": XL_LEGEND_POSITION.LEFT}[legend]
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(fs)
        else:
            chart.has_legend = False
        plot = chart.plots[0]
        palette = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
        if pie:
            s = plot.series[0]
            for j in range(len(c.get("categories", []))):
                pt = s.points[j]
                fill_theme(pt.format.fill, palette[j % 6], 0.0 if j < 6 else 0.4)
                pt.format.line.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
        else:
            for i, s in enumerate(plot.series):
                col = palette[i % 6]
                if ctype.startswith(("line", "scatter", "radar")):
                    s.format.line.color.theme_color = THEME[col]
                    s.format.line.width = Pt(2.5)
                    s.smooth = bool(c.get("smooth", False))
                    if ctype.startswith("scatter") and ctype != "scatter-lines":
                        s.format.line.fill.background()
                    try:
                        s.marker.style = XL_MARKER_STYLE.CIRCLE
                        s.marker.size = 7
                        fill_theme(s.marker.format.fill, col)
                        s.marker.format.line.color.theme_color = THEME[col]
                    except Exception:
                        pass
                else:
                    fill_theme(s.format.fill, col)
            if ctype.startswith(("column", "bar")):
                plot.gap_width = int(c.get("gap_width", 70))
            va = chart.value_axis
            va.has_major_gridlines = True
            va.major_gridlines.format.line.color.theme_color = MSO_THEME_COLOR.TEXT_1
            va.major_gridlines.format.line.color.brightness = 0.85
            va.major_gridlines.format.line.width = Pt(0.75)
            va.format.line.fill.background()
            va.tick_labels.font.size = Pt(fs)
            if c.get("y_min") is not None:
                va.minimum_scale = float(c["y_min"])
            if c.get("y_max") is not None:
                va.maximum_scale = float(c["y_max"])
            if c.get("number_format"):
                va.tick_labels.number_format = c["number_format"]
                va.tick_labels.number_format_is_linked = False
            ca = chart.category_axis
            ca.tick_labels.font.size = Pt(fs)
            ca.format.line.color.theme_color = MSO_THEME_COLOR.TEXT_1
            ca.format.line.color.brightness = 0.6
            ca.has_major_gridlines = False
            for axis, key in ((va, "y_title"), (ca, "x_title")):
                if c.get(key):
                    axis.has_title = True
                    axis.axis_title.text_frame.text = str(c[key])
                    af = axis.axis_title.text_frame.paragraphs[0].runs[0].font
                    af.size, af.bold = Pt(fs), False
        if c.get("labels"):
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.font.size = Pt(fs * 0.95)
            if c.get("number_format"):
                dl.number_format, dl.number_format_is_linked = c["number_format"], False
            if pie:
                dl.position = XL_LABEL_POSITION.OUTSIDE_END if ctype == "pie" else XL_LABEL_POSITION.CENTER
                if c.get("percent"):
                    dl.show_percentage, dl.show_value = True, False
            elif "stacked" in ctype:
                dl.position = XL_LABEL_POSITION.CENTER  # outside-end is invalid for stacked charts
            elif ctype.startswith(("column", "bar")):
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
            elif ctype.startswith("line"):
                dl.position = XL_LABEL_POSITION.ABOVE
        self._animate(it, gf, "graphic")

    # ---- tables -------------------------------------------------------------------------
    def _render_table(self, slide, it, b, scale):
        t = it["table"]
        header = t.get("header")
        rows = [list(map(lambda v: "" if v is None else str(v), r)) for r in t.get("rows", [])]
        all_rows = ([list(map(str, header))] if header else []) + rows
        ncols = max(len(r) for r in all_rows)
        all_rows = [r + [""] * (ncols - len(r)) for r in all_rows]
        size = round((t.get("size") or self.level_sizes[0] * 0.78) * scale, 1)
        need = [max(len(strip_markup(r[j])) for r in all_rows) * size * 0.55 / 72 + 0.35 for j in range(ncols)]
        if t.get("widths"):
            need = [float(x) for x in t["widths"]]
        width = b.w * float(t["width"]) if t.get("width") else min(b.w, max(sum(need) * 1.25, b.w * 0.55))
        colw = [width * n / sum(need) for n in need]
        x = b.x + (b.w - width) / 2
        gf = slide.shapes.add_table(len(all_rows), ncols, inch(x), inch(b.y), inch(width), inch(b.h))
        gf.name = "Table"
        tbl = gf.table
        style = t.get("style", "booktabs")
        tblPr = tbl._tbl.tblPr
        tblPr.set("firstRow", "1" if header else "0")
        tblPr.set("bandRow", "0")
        sid = tblPr.find(qn("a:tableStyleId"))
        if sid is None:
            sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # "No Style, No Grid": we draw our own rules
        for j, cw in enumerate(colw):
            tbl.columns[j].width = inch(cw)
        rh = b.h / len(all_rows)
        for i in range(len(all_rows)):
            tbl.rows[i].height = inch(rh)
        align = t.get("align")
        numeric = re.compile(r"^[\s$€£+\-−]*[\d.,]+\s*[%kKmMbB×x]*\s*$")
        last = len(all_rows) - 1
        for i, row in enumerate(all_rows):
            is_head = bool(header) and i == 0
            for j, val in enumerate(row):
                cell = tbl.cell(i, j)
                cell.margin_left = cell.margin_right = inch(0.1)
                cell.margin_top = cell.margin_bottom = inch(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                if align:
                    a = align[j] if j < len(align) else "l"
                else:
                    # test the visible text, so **12M** or ==86.2%== still count as numbers
                    a = "r" if (numeric.match(strip_markup(val)) and not is_head and j > 0) else ("l" if j == 0 else "c")
                    if is_head and j > 0 and all(numeric.match(strip_markup(r[j])) for r in rows if r[j]):
                        a = "r"
                p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[a]
                st = {"size": size, "bold": True if is_head else None}
                if style == "striped" and is_head:
                    st["color"] = "white"
                else:
                    st["color"] = "tx1"
                self._add_runs(p, val, st)
                borders, fill = {}, None
                if style == "booktabs":
                    if i == 0:
                        borders["T"] = 1.75
                    if is_head:
                        borders["B"] = 0.9
                    if i == last:
                        borders["B"] = 1.75
                elif style == "grid":
                    borders = {"T": 0.75, "B": 0.75, "L": 0.75, "R": 0.75}
                    fill = ("accent1", 0.85) if is_head else None
                elif style == "striped":
                    fill = ("accent1", 0.0) if is_head else (("accent1", 0.9) if i % 2 == 0 else None)
                self._cell_format(cell, borders, fill)
        self._animate(it, gf, "table")

    def _cell_format(self, cell, borders, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        for ch in list(tcPr):
            tcPr.remove(ch)
        for side in ("L", "R", "T", "B"):
            ln = etree.SubElement(tcPr, qn(f"a:ln{side}"))
            if side in borders:
                ln.set("w", str(int(borders[side] * 12700)))
                ln.set("cap", "flat")
                ln.set("cmpd", "sng")
                sf = etree.SubElement(ln, qn("a:solidFill"))
                sc = etree.SubElement(sf, qn("a:schemeClr"))
                sc.set("val", "tx1")
                etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
            else:
                ln.set("w", "0")
                etree.SubElement(ln, qn("a:noFill"))
        if fill:
            sf = etree.SubElement(tcPr, qn("a:solidFill"))
            sc = etree.SubElement(sf, qn("a:schemeClr"))
            sc.set("val", fill[0])
            if fill[1]:
                etree.SubElement(sc, qn("a:lumMod")).set("val", str(int((1 - fill[1]) * 100000)))
                etree.SubElement(sc, qn("a:lumOff")).set("val", str(int(fill[1] * 100000)))
        else:
            etree.SubElement(tcPr, qn("a:noFill"))

    # ---- images -------------------------------------------------------------------------
    def _render_image(self, slide, it, b, scale):
        im, ratio = self._figure("image", it)
        path = self._path(im["path"])
        if not os.path.exists(path):
            raise FileNotFoundError(f"image not found: {path}")
        x, y, w, h = self._figure_box(im, ratio, b)
        pic = slide.shapes.add_picture(path, inch(x), inch(y), inch(w), inch(h))
        pic.name = im.get("alt", os.path.basename(path))
        if im.get("alt"):
            pic._element.nvPicPr.cNvPr.set("descr", str(im["alt"]))
        self._animate(it, pic, "pic")
        self._caption(slide, it, im, b, y + h, scale)

    def _render_placeholder(self, slide, it, b, scale):
        """A native dashed box that reserves a figure's space; swap it later via Shape Fill > Picture."""
        ph, ratio = self._figure("placeholder", it)
        x, y, w, h = self._figure_box(ph, ratio, b)
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, inch(x), inch(y), inch(w), inch(h))
        drop_style(shp)
        fill_theme(shp.fill, "tx1", 0.93)
        shp.line.width = Pt(1.25)
        shp.line.dash_style = MSO_LINE.DASH
        shp.line.color.theme_color = THEME["tx1"]
        shp.line.color.brightness = 0.55
        label = str(ph.get("label") or "Figure")
        shp.name = f"Placeholder: {strip_markup(label)[:40]}"
        shp._element.nvSpPr.cNvPr.set("descr", str(ph.get("alt") or f"Placeholder - {strip_markup(label)}"))
        tf = shp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = inch(0.15)
        size = round(self.level_sizes[0] * 0.8 * scale, 1)
        self._fill_text(tf, label, size=size, style={"color": "muted", "bold": True}, align="center")
        if ph.get("note"):
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            self._add_runs(p, str(ph["note"]), {"size": round(size * 0.75, 1), "color": "muted"})
        self._animate(it, shp, "sp_text")
        self._caption(slide, it, ph, b, y + h, scale)

    def _render_plot(self, slide, it, b, scale):
        """matplotlib figure drawn at its final size in the template's font and colours (see plots.py)."""
        import plots  # only decks with plots need matplotlib

        p, ratio = self._figure("plot", it)
        cap_h = 0.4 if p.get("caption") else 0
        if ratio is None:  # fill the area; confusion matrices stay roughly square
            ratio = (b.w * float(p.get("width", 1.0))) / max(b.h - cap_h, 0.1)
            if p.get("type") == "confusion":
                ratio = min(ratio, 1.2)
        x, y, w, h = self._figure_box({"width": 1.0, **p}, ratio, b)
        colors, fonts = self._theme()
        png = plots.render(p, colors, fonts["minor"], w, h, round(self.level_sizes[0] * 0.66 * scale, 1))
        pic = slide.shapes.add_picture(io.BytesIO(png), inch(x), inch(y), inch(w), inch(h))
        label = p.get("title") or p.get("ylabel") or p.get("type")
        pic.name = f"Plot: {strip_markup(str(label))[:40]}"
        pic._element.nvPicPr.cNvPr.set("descr", str(p.get("alt") or f"{p.get('type')} plot: {label}"))
        self._animate(it, pic, "pic")
        self._caption(slide, it, p, b, y + h, scale)

    # ---- flows: flowcharts, workflows, step rows (layout in flows.py) ---------------------
    def _flow(self, kind, it):
        """(layout in em, settings) of a flow or steps item, computed once per item."""
        import flows

        hit = self._flow_cache.get(id(it))
        if hit and hit[0] is it:
            return hit[1], hit[2]
        v = it[kind]
        if kind == "flow":
            if not isinstance(v, dict):
                raise ValueError("flow needs settings, e.g. {nodes: [...], edges: ['a -> b']}")
            spec = v
        elif isinstance(v, dict):
            spec = v
        else:  # steps: [A, B, C] with its settings next to it, like bullets
            spec = {**{k: x for k, x in it.items() if k not in ("steps", "animate", "build", "name")}, "items": v}
        lay = flows.layout(kind, spec, strip_markup)
        self._flow_cache[id(it)] = (it, lay, spec)
        return lay, spec

    def _flow_font(self, spec, scale):
        return float(spec.get("size") or self.level_sizes[0] * 0.72) * scale

    def _render_flow(self, slide, kind, it, b, scale):
        """Native shapes and connectors glued to their connection sites: drag a shape, its arrows follow."""
        import flows

        lay, spec = self._flow(kind, it)
        f = self._flow_font(spec, scale)
        cap = 0.4 if spec.get("caption") else 0
        avail_h = max(b.h - cap, 0.1)
        k = min(1.0, b.w * float(spec.get("width", 1.0)) / (lay.w * f / 72), avail_h / (lay.h * f / 72))
        fs = f * k
        if fs < FLOW_MIN_PT - 0.5:
            title = slide.shapes.title.text_frame.text if slide.shapes.title else "(untitled)"
            warn(f"{kind} diagram on slide '{title}' is drawn with {fs:.0f} pt text - give it more room, "
                 f"split it, or change its direction")
        em = fs / 72  # inches per em
        x0 = b.x + (b.w - lay.w * em) / 2
        y0 = b.y + (avail_h - lay.h * em) / 2
        X, Y, L = (lambda v: inch(x0 + v * em)), (lambda v: inch(y0 + v * em)), (lambda v: inch(v * em))
        line_w = Pt(1.25 if k > 0.7 else 1.0)
        groups, nodes, arrows, labels = [], [], [], []  # (shape, build step)

        for g in lay.groups:
            shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, X(g.x), Y(g.y), L(g.w), L(g.h))
            drop_style(shp)
            shp.adjustments[0] = min(0.5, 0.5 / min(g.w, g.h))
            neutral = g.color in ("tx1", "text")
            if g.fill:
                shp.fill.solid()
                paint(shp.fill.fore_color, g.color, 0.94)
            else:
                shp.fill.background()
            shp.line.width = Pt(1)
            shp.line.dash_style = MSO_LINE.DASH
            paint(shp.line.color, g.color, 0.45 if neutral else 0.0)
            tf = shp.text_frame
            tf.word_wrap, tf.auto_size, tf.vertical_anchor = True, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.TOP
            tf.margin_left = tf.margin_right = L(flows.GPAD * 0.6)
            tf.margin_top, tf.margin_bottom = L(0.12), L(0.05)
            self._fill_text(tf, g.label, size=round(fs * flows.LABEL, 1),
                            style={"color": "muted" if neutral else g.color, "bold": True}, align="left")
            shp.name = f"Group: {strip_markup(g.label)[:40]}"
            groups.append((shp, g.rank))

        by_id = {}
        for n in lay.nodes:
            shp = slide.shapes.add_shape(getattr(MSO_SHAPE, flows.SHAPES[n.shape][0]), X(n.x), Y(n.y), L(n.w), L(n.h))
            drop_style(shp)
            if n.shape == "box":
                shp.adjustments[0] = min(0.5, 0.3 / min(n.w, n.h))
            elif n.adj is not None:
                shp.adjustments[0] = n.adj
            if n.style == "plain":
                shp.fill.background()
                shp.line.fill.background()
            else:
                shp.fill.solid()
                if n.style == "outline":
                    paint(shp.fill.fore_color, "bg1")
                else:
                    paint(shp.fill.fore_color, n.color, 0.85 if n.style == "tinted" else 0.0)
                shp.line.width = line_w
                paint(shp.line.color, n.color)
            tf = shp.text_frame
            tf.word_wrap, tf.auto_size, tf.vertical_anchor = True, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.MIDDLE
            tf.margin_left = tf.margin_right = L(n.pad[0])
            tf.margin_top = tf.margin_bottom = L(n.pad[1])
            color = n.text_color or ("white" if n.style == "solid" else "tx1")
            self._fill_text(tf, n.text, size=round(fs, 1), style={"color": color, "bold": True if n.bold else None},
                            align="center")
            shp.name = f"{'Step' if kind == 'steps' else 'Node'}: {strip_markup(n.text)[:40]}"
            by_id[n.id] = shp
            nodes.append((shp, n.rank))

        line_color = spec.get("line_color", ["tx1", 0.3])
        for e in lay.edges:
            prst, off, ext, rot, fh, fv, adj = flows.connector_geometry([(int(X(x)), int(Y(y))) for x, y in e.pts])
            cx = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 1, 1)
            drop_style(cx)
            el = cx._element
            xfrm = el.spPr.find(qn("a:xfrm"))
            for key, on in (("flipH", fh), ("flipV", fv)):
                if on:
                    xfrm.set(key, "1")
                elif key in xfrm.attrib:
                    del xfrm.attrib[key]
            if rot:
                xfrm.set("rot", str(rot * 60000))
            xfrm.find(qn("a:off")).set("x", str(off[0]))
            xfrm.find(qn("a:off")).set("y", str(off[1]))
            xfrm.find(qn("a:ext")).set("cx", str(ext[0]))
            xfrm.find(qn("a:ext")).set("cy", str(ext[1]))
            geom = el.spPr.find(qn("a:prstGeom"))
            geom.set("prst", prst)
            av = geom.find(qn("a:avLst"))
            if av is None:
                av = etree.SubElement(geom, qn("a:avLst"))
            for ch in list(av):
                av.remove(ch)
            for i, val in enumerate(adj, 1):
                gd = etree.SubElement(av, qn("a:gd"))
                gd.set("name", f"adj{i}")
                gd.set("fmla", f"val {val}")
            glue = el.find(qn("p:nvCxnSpPr")).find(qn("p:cNvCxnSpPr"))
            for tag, nid, idx in (("a:stCxn", e.a, e.sites[0]), ("a:endCxn", e.b, e.sites[1])):
                c = etree.SubElement(glue, qn(tag))
                c.set("id", str(by_id[nid].shape_id))
                c.set("idx", str(idx))
            cx.line.width = Pt(e.width or float(spec.get("line_width", 1.25)))
            paint(cx.line.color, e.color or line_color)
            if e.dash != "solid":
                cx.line.dash_style = MSO_LINE.DASH if e.dash == "dashed" else MSO_LINE.ROUND_DOT
            ln = el.spPr.find(qn("a:ln"))
            for tag, on in (("a:headEnd", e.arrow in ("start", "both")), ("a:tailEnd", e.arrow in ("end", "both"))):
                if on:
                    end = etree.SubElement(ln, qn(tag))
                    end.set("type", "triangle")
                    end.set("w", "med")
                    end.set("len", "med")
            cx.name = f"Arrow: {e.a} -> {e.b}"
            nodes[0][0]._element.addprevious(el)  # connectors sit under the shapes they join
            arrows.append((cx, e.rank))
            if e.label_box:
                bx, by, bw, bh = e.label_box
                tb = slide.shapes.add_textbox(X(bx), Y(by), L(bw), L(bh))
                tf = tb.text_frame
                tf.word_wrap, tf.auto_size, tf.vertical_anchor = False, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.MIDDLE
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                self._fill_text(tf, e.label, size=round(fs * flows.LABEL, 1), style={"color": e.color or line_color},
                                align="center")
                tb.name = f"Label: {strip_markup(e.label)[:40]}"
                labels.append((tb, e.rank))

        parts = groups + arrows + nodes + labels  # z-order
        alt = spec.get("alt") or f"{'Steps' if kind == 'steps' else 'Flowchart'}: " + ", ".join(
            strip_markup(n.text) for n in lay.nodes)[:250]
        if it.get("build"):  # one click per step along the flow; arrows arrive with their target
            effect = it["build"] if isinstance(it["build"], str) else "fade"
            for r in sorted({rank for _, rank in parts}):
                effs = [Effect(shp.shape_id, effect) for shp, rank in parts if rank == r]
                self.steps.append([effs])
            for shp, _ in groups + nodes + labels:
                self.builds[shp.shape_id] = "shape"
            self._last_main = nodes[0][0]
        else:
            grp = slide.shapes.add_group_shape([shp for shp, _ in parts])
            grp.name = f"{'Steps' if kind == 'steps' else 'Flow'}: {strip_markup(lay.nodes[0].text)[:40]}"
            grp._element.nvGrpSpPr.cNvPr.set("descr", str(alt))
            self._animate(it, grp, "group")
        self._caption(slide, it, spec, b, y0 + lay.h * em, scale)

    def _theme(self):
        """Theme colours {dk1: 'RRGGBB', accent1: ...} and fonts {major, minor} of the template."""
        if getattr(self, "_theme_cache", None) is None:
            root = etree.fromstring(self.prs.slide_master.part.part_related_by(RT.THEME).blob)
            colors = {}
            for el in root.find(f".//{qn('a:clrScheme')}"):
                c = el[0]
                colors[etree.QName(el).localname] = c.get("val") if c.tag == qn("a:srgbClr") else c.get("lastClr")
            fonts = {k: root.find(f".//{qn(f'a:{k}Font')}/{qn('a:latin')}").get("typeface") for k in ("major", "minor")}
            self._theme_cache = ({k: v for k, v in colors.items() if v}, fonts)
        return self._theme_cache

    @staticmethod
    def _figure_box(spec, ratio, b):
        """Largest box of the given ratio within spec.width of the area, centred above any caption."""
        cap_h = 0.4 if spec.get("caption") else 0
        max_w, max_h = b.w * float(spec.get("width", 0.75)), b.h - cap_h
        w = min(max_w, max_h * ratio)
        h = w / ratio
        return b.x + (b.w - w) / 2, b.y + (max_h - h) / 2, w, h

    def _caption(self, slide, it, spec, b, top, scale):
        if not spec.get("caption"):
            return
        tb = slide.shapes.add_textbox(inch(b.x), inch(top + 0.05), inch(b.w), inch(0.4))
        self._prep_tf(tb.text_frame)
        self._fill_text(tb.text_frame, str(spec["caption"]), size=round(self.level_sizes[0] * 0.66 * scale, 1),
                        style={"color": "muted", "italic": True}, align="center")
        tb.name = "Caption"
        if it.get("animate"):
            self._animate({"animate": {"effect": self._anim_spec(it)[0], "on": "with"}}, tb, "sp_text")

    # ---- animation bookkeeping -----------------------------------------------------------
    @staticmethod
    def _anim_spec(it):
        a = it.get("animate")
        if isinstance(a, dict):
            return a.get("effect", "fade"), a.get("on", "click"), int(float(a.get("duration", 0.5)) * 1000)
        return (str(a) if a not in (True, None) else "fade"), "click", 500

    def _animate(self, it, shape, kind, groups=None):
        if getattr(self, "_last_main", None) is None:
            self._last_main = shape
        build = it.get("build")
        spid = shape.shape_id
        if build and groups:
            effect = build if isinstance(build, str) else "fade"
            for st, en in groups:
                step = [[Effect(spid, effect, (st, st))]]
                for k in range(st + 1, en + 1):  # sub-items appear with their parent
                    step[0].append(Effect(spid, effect, (k, k), trigger="with"))
                self.steps.append(step)
            self.builds[spid] = "para"
            return
        if not it.get("animate"):
            return
        effect, on, dur = self._anim_spec(it)
        e = Effect(spid, effect, None, dur, on)
        if on == "with" and self.steps:
            self.steps[-1][-1].append(e)
        elif on == "after" and self.steps:
            self.steps[-1].append([e])
        else:
            self.steps.append([[e]])
        self.builds[spid] = {"sp_text": "shape", "graphic": "graphic"}.get(kind, "none")

    def _apply_timing(self, slide, transition):
        sld = slide._element
        for tag in (qn("p:transition"), qn("p:timing"), f"{{{NS_MC}}}AlternateContent"):
            for el in sld.findall(tag):
                sld.remove(el)
        anchor = sld.find(qn("p:clrMapOvr"))
        if anchor is None:
            anchor = sld.find(qn("p:cSld"))
        idx = list(sld).index(anchor) + 1
        tx = transition_xml(transition)
        if tx:
            sld.insert(idx, parse_xml(tx))
            idx += 1
        if self.steps:
            sld.insert(idx, parse_xml(timing_xml(self.steps, self.builds)))

    # ---- equations -----------------------------------------------------------------------
    def _finalize_math(self, slides):
        if not self.math_jobs:
            return
        omml = MathEngine().convert(self.math_jobs)
        mark_re = re.compile(re.escape(MARK) + r"(\d+)" + re.escape(MARK))
        for slide in slides:
            tree = slide.shapes._spTree
            # text shapes and table frames (math in table cells lives in a p:graphicFrame)
            targets = [sp for sp in tree.iter(qn("p:sp"), qn("p:graphicFrame"))
                       if any(mark_re.search(t.text or "") for t in sp.iter(qn("a:t")))]
            for sp in targets:
                fallback = copy.deepcopy(sp)
                for t in fallback.iter(qn("a:t")):
                    if mark_re.search(t.text or ""):
                        rpr = t.getparent().find(qn("a:rPr"))
                        latin = rpr.find(qn("a:latin")) if rpr is not None else None
                        if latin is not None:  # plain-text fallback reads better in the body font
                            rpr.remove(latin)
                            rpr.set("i", "1")
                    t.text = mark_re.sub(lambda m: latex_to_unicode(self.math_jobs[int(m.group(1))][0]), t.text or "")
                has_native = False
                choice = copy.deepcopy(sp)
                for r in list(choice.iter(qn("a:r"))):
                    t = r.find(qn("a:t"))
                    m = mark_re.fullmatch(t.text or "") if t is not None else None
                    if not m:
                        continue
                    n = int(m.group(1))
                    el = omml[n]
                    if el is None:
                        t.text = latex_to_unicode(self.math_jobs[n][0])
                        continue
                    has_native = True
                    wrapper = etree.Element(f"{{{NS_A14}}}m", nsmap={"a14": NS_A14})
                    math_el = copy.deepcopy(el)
                    self._style_math(math_el, r.find(qn("a:rPr")))
                    wrapper.append(math_el)
                    r.getparent().replace(r, wrapper)
                if not has_native:
                    sp.getparent().replace(sp, fallback)
                    continue
                ac = etree.Element(f"{{{NS_MC}}}AlternateContent", nsmap={"mc": NS_MC})
                ch = etree.SubElement(ac, f"{{{NS_MC}}}Choice", nsmap={"a14": NS_A14})
                ch.set("Requires", "a14")
                ch.append(choice)
                fb = etree.SubElement(ac, f"{{{NS_MC}}}Fallback")
                fb.append(fallback)
                sp.getparent().replace(sp, ac)

    @staticmethod
    def _style_math(math_el, rpr):
        """Give every math run the size/colour of the text it replaces (as PowerPoint does)."""
        for mr in math_el.iter(f"{{{NS_M}}}r"):
            mt = mr.find(f"{{{NS_M}}}t")
            sty = mr.find(f"{{{NS_M}}}rPr/{{{NS_M}}}sty")
            new = etree.Element(qn("a:rPr"))
            new.set("lang", "en-US")
            if rpr is not None:
                for k in ("sz", "b"):
                    if rpr.get(k):
                        new.set(k, rpr.get(k))
            upright = sty is not None and sty.get(f"{{{NS_M}}}val") == "p"
            text = mt.text if mt is not None else ""
            new.set("i", "0" if upright or not any(ch.isalpha() for ch in text or "") else "1")
            if rpr is not None and rpr.find(qn("a:solidFill")) is not None:
                new.append(copy.deepcopy(rpr.find(qn("a:solidFill"))))
            latin = etree.SubElement(new, qn("a:latin"))
            latin.set("typeface", "Cambria Math")
            pos = 1 if (len(mr) and mr[0].tag == f"{{{NS_M}}}rPr") else 0
            mr.insert(pos, new)


# --------------------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------------------
def build_deck(content_path: str, output: str | None = None, template: str | None = None,
               keep_template_slides: bool = False) -> str:
    with open(content_path, encoding="utf-8") as fh:
        content = yaml.safe_load(fh) or {}
    base = os.path.dirname(os.path.abspath(content_path))
    builder = DeckBuilder(content, template, base, keep_template_slides=keep_template_slides)
    prs = builder.build()
    out = output or (content.get("meta", {}) or {}).get("output") or os.path.splitext(content_path)[0] + ".pptx"
    if not os.path.isabs(out) and output is None and (content.get("meta", {}) or {}).get("output"):
        out = os.path.join(base, out)
    prs.save(out)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build an editable Beamer-style .pptx from a YAML content file.")
    ap.add_argument("content", help="YAML content file")
    ap.add_argument("-o", "--output", help="output .pptx (default: next to the content file)")
    ap.add_argument("-t", "--template", help="template .pptx/.potx (default: meta.template or templates/beamer_oxford.pptx)")
    ap.add_argument("--keep-template-slides", action="store_true", help="keep slides already in the template")
    a = ap.parse_args(argv)
    out = build_deck(a.content, a.output, a.template, a.keep_template_slides)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
