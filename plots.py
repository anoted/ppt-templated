"""
plots.py - academic-style figures for ppt-templated: bar, box, line, scatter, histogram, heatmap,
confusion matrix, formulas (`function`) and your own matplotlib code (`python`).

ppt-templated calls render() for every `plot:` item. The figure is drawn at the exact size it gets on
the slide, in the template's body font and theme colours, and embedded straight into the deck as
an SVG (vector, drawn by PowerPoint 2019+) with a 300 dpi PNG fallback (nothing is written to disk).
Use `chart:` instead when you need a native, editable PowerPoint chart; use `plot:` for box plots,
confusion matrices, error bands, trend lines and print-quality styling.

Colour specs, accepted by every colour setting:
    accent1 ... accent6, tx1, bg1, tx2, bg2    theme slots of the template
    structure | alert | example | muted        Beamer roles (accent1 | accent2 | accent3 | grey text)
    [accent1, 0.6]                             moved 60% towards the slide background (negative = away)
    "#1F3A5F", black, tab:blue                 anything matplotlib understands
Settings that colour several things (series, boxes, ...) take one colour or a list (cycled).

`function` evaluates formulas such as "1 / (1 + exp(-x))" with a small evaluator that only knows
arithmetic, comparisons and the maths functions in FUNCS - no Python. `python` runs Python code from
the YAML, so it only runs when the build was started with --allow-python.

All plot settings are documented in YAML_GUIDE.md (`plot`). When you add or change a setting
or plot type here, update that guide in the same commit.
"""

from __future__ import annotations

import ast
import io
import logging
import math
import os

import matplotlib
import numpy as np
from matplotlib import font_manager
from matplotlib.cbook import boxplot_stats
from matplotlib.colors import LinearSegmentedColormap, to_hex, to_rgb
from matplotlib.figure import Figure
from matplotlib.ticker import StrMethodFormatter

PLOT_TYPES = ("bar", "box", "line", "scatter", "hist", "heatmap", "confusion", "function", "python")
COMMON_KEYS = {"type", "title", "xlabel", "ylabel", "legend", "grid", "xlim", "ylim", "xscale", "yscale",
               "xformat", "yformat", "xticks_rotation", "width", "height", "aspect", "font", "font_size", "math",
               "caption", "alt", "dpi", "colors", "vector", "data"}
TYPE_KEYS = {
    "bar": {"categories", "series", "values", "y", "x", "errors", "value_format", "horizontal", "stacked",
            "bar_width", "hatch"},
    "box": {"groups", "stats", "points", "means", "notch", "whis", "fliers", "horizontal", "box_width", "y"},
    "line": {"x", "series", "y", "values", "errors", "band_alpha", "markers", "dashes", "line_width", "marker_size"},
    "scatter": {"series", "values", "y", "x", "trend", "trend_label", "markers", "marker_size", "alpha"},
    "hist": {"series", "values", "y", "bins", "density", "stacked", "alpha", "cumulative", "range"},
    "heatmap": {"matrix", "xlabels", "ylabels", "labels", "values", "value_format", "colorbar", "square", "vmin",
                "vmax", "center", "x", "y"},
    "confusion": {"labels", "matrix", "normalize", "values", "value_format", "colorbar", "square", "x", "y"},
    "function": {"f", "functions", "x", "samples", "params", "dashes", "line_width"},
    "python": {"code", "script"},
}
HATCHES = ["", "//", "..", "xx", "\\\\", "oo", "++", "--"]
MARKERS = ["o", "s", "^", "D", "v", "P", "X", "*"]
DASHES = ["-", "--", "-.", ":"]
ALIASES = {"structure": "accent1", "alert": "accent2", "example": "accent3", "text": "tx1", "background": "bg1",
           "white": "bg1"}
SLOT_OF = {"tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}  # when the theme has no colour map
SERIF_FALLBACK = ["Cambria", "Caladea", "Times New Roman", "Liberation Serif", "STIXGeneral", "DejaVu Serif"]
SANS_FALLBACK = ["Calibri", "Carlito", "Arial", "Liberation Sans", "Helvetica", "DejaVu Sans"]
# metric-compatible open fonts (installed with LibreOffice): same widths, so labels fit the same way
SAME_METRICS = {"calibri": "Carlito", "calibri light": "Carlito", "cambria": "Caladea", "arial": "Liberation Sans",
                "helvetica": "Liberation Sans", "times new roman": "Liberation Serif", "courier new": "Liberation Mono",
                "georgia": "Gelasio", "segoe ui": "Selawik", "aptos": "Carlito"}
SANS = {"calibri", "calibri light", "carlito", "arial", "helvetica", "segoe ui", "aptos", "verdana", "tahoma",
        "trebuchet ms", "liberation sans", "inter", "open sans", "roboto", "lato", "source sans pro", "gill sans",
        "franklin gothic", "century gothic", "dejavu sans"}


# ------------------------------------------------------------------------------------------
# colours and fonts
# ------------------------------------------------------------------------------------------
def tint(color, amount: float, background="#FFFFFF") -> str:
    """amount > 0 blends towards the background, < 0 away from it (towards black on a light background,
    towards white on a dark one), like the tints of the slides themselves."""
    rgb, bg = np.array(to_rgb(color)), np.array(to_rgb(background))
    if amount >= 0:
        rgb = rgb + (bg - rgb) * amount
    else:
        far = np.zeros(3) if bg.mean() >= 0.5 else np.ones(3)
        rgb = rgb + (far - rgb) * -amount
    return to_hex(np.clip(rgb, 0, 1))


class Palette:
    """Resolves colour specs against the template's theme colours (tx1/bg1 as the slides use them)."""

    def __init__(self, theme_colors: dict):
        self.theme = {k: "#" + v for k, v in theme_colors.items()}
        for role, slot in SLOT_OF.items():
            if role not in self.theme and slot in self.theme:
                self.theme[role] = self.theme[slot]
        self.bg = self.theme.get("bg1", "#FFFFFF")

    @staticmethod
    def _is_tint(spec) -> bool:
        return (isinstance(spec, (list, tuple)) and len(spec) == 2 and isinstance(spec[0], str)
                and isinstance(spec[1], (int, float)))

    def __call__(self, spec, default=None):
        spec = default if spec is None else spec
        if spec is None:
            return None
        if self._is_tint(spec):
            return tint(self(spec[0]), float(spec[1]), self.bg)
        if isinstance(spec, str):
            if spec in ("none", "transparent"):
                return "none"
            if spec == "muted":
                return tint(self("tx1"), 0.45, self.bg)
            key = ALIASES.get(spec, spec)
            if key in self.theme:
                return self.theme[key]
        return spec

    def cycle(self, spec, n: int, default) -> list:
        """n colours from one colour or a list of colours."""
        spec = default if spec is None else spec
        cols = [self(x) for x in spec] if isinstance(spec, (list, tuple)) and not self._is_tint(spec) else [self(spec)]
        return [cols[i % len(cols)] for i in range(n)]


def _lum(color) -> float:
    r, g, b = to_rgb(color)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _installed(name) -> bool:
    try:
        font_manager.findfont(font_manager.FontProperties(family=name), fallback_to_default=False)
        return True
    except ValueError:
        return False


def pick_font(*candidates) -> str:
    """First installed font of the candidates, else a metric-compatible stand-in (Carlito for Calibri, ...),
    else a fallback of the same kind (sans for a sans template font, serif otherwise)."""
    logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)  # "findfont: ..." for every miss
    names = [str(c) for c in candidates if c]
    for name in names:
        if _installed(name):
            return name
        alt = SAME_METRICS.get(name.lower())
        if alt and _installed(alt):
            return alt
    sans = bool(names) and names[-1].lower() in SANS
    for name in (SANS_FALLBACK if sans else SERIF_FALLBACK):
        if _installed(name):
            return name
    return "DejaVu Sans" if sans else "DejaVu Serif"


def _rc(font: str, size: float, math: str) -> dict:
    rc = {
        "font.family": [font], "font.size": size,
        "axes.titlesize": size * 1.1, "axes.labelsize": size, "legend.fontsize": size * 0.9,
        "xtick.labelsize": size * 0.9, "ytick.labelsize": size * 0.9,
        "axes.spines.top": False, "axes.spines.right": False, "axes.linewidth": 0.8,
        "axes.axisbelow": True, "axes.titlepad": 8, "axes.labelpad": 5, "axes.titleweight": "normal",
        "xtick.direction": "out", "ytick.direction": "out", "xtick.major.size": 3.5, "ytick.major.size": 3.5,
        "xtick.major.width": 0.8, "ytick.major.width": 0.8, "xtick.minor.visible": False,
        "legend.frameon": False, "legend.handlelength": 1.6,
        "lines.linewidth": 2.0, "lines.markersize": 5.5, "patch.linewidth": 0.8,
        "savefig.transparent": True, "figure.dpi": 100,
    }
    if math == "font":  # $...$ in labels uses the body font, missing glyphs from STIX
        rc.update({"mathtext.fontset": "custom", "mathtext.rm": font, "mathtext.it": f"{font}:italic",
                   "mathtext.bf": f"{font}:bold", "mathtext.fallback": "stix"})
    else:  # stix | cm | stixsans | dejavuserif | dejavusans
        rc["mathtext.fontset"] = math
    return rc


# ------------------------------------------------------------------------------------------
# entry point
# ------------------------------------------------------------------------------------------
def keys(kind=None) -> set:
    """Settings a `plot:` of this type accepts (all plot settings for an unknown type)."""
    return COMMON_KEYS | TYPE_KEYS.get(kind, set().union(*TYPE_KEYS.values()))


def render(spec: dict, theme_colors: dict, theme_font: str | None, width_in: float, height_in: float,
           font_size: float, vector: bool = False, allow_python: bool = False, base_dir: str = "."):
    """Draw one `plot:` item at width_in x height_in inches. Returns PNG bytes, or (PNG, SVG) with vector.
    allow_python: run the code of `type: python` plots; base_dir: folder their `script:` is relative to."""
    spec = dict(spec)  # drawing fills in defaults (e.g. confusion axis labels); keep the content untouched
    kind = spec.get("type")
    if kind not in PLOT_TYPES:
        raise ValueError(f"plot type must be one of {', '.join(PLOT_TYPES)}, got {kind!r}")
    if kind == "python":
        if not allow_python:
            raise ValueError("plot type 'python' runs Python code from the YAML: build with --allow-python "
                             "(build_deck(..., allow_python=True)) if you trust this file")
        spec.setdefault("legend", "none")  # the code makes its own legend if it wants one
        spec.setdefault("grid", "none")
        spec["_base_dir"] = base_dir
    P = Palette(theme_colors)
    c = dict(spec.get("colors") or {})
    font = pick_font(spec.get("font"), theme_font)
    size = float(spec.get("font_size") or font_size)
    with matplotlib.rc_context(_rc(font, size, spec.get("math", "font"))):
        fig = Figure(figsize=(width_in, height_in), layout="constrained")
        fig.set_facecolor(P(c.get("figure"), "none"))
        ax = fig.add_subplot()
        n_labeled = DRAW[kind](fig, ax, spec, c, P, size)
        _style(fig, ax, spec, c, P, kind, n_labeled)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=int(spec.get("dpi", 300)), metadata={"Software": None})
        if not vector:
            return buf.getvalue()
        # text as outlines: PowerPoint draws exactly what matplotlib drew, whatever fonts are installed
        with matplotlib.rc_context({"svg.fonttype": "path", "svg.hashsalt": "ppt-templated"}):
            svg = io.BytesIO()
            fig.savefig(svg, format="svg", metadata={"Date": None})
    return buf.getvalue(), svg.getvalue()


def _need(s: dict, key: str, kind: str):
    if s.get(key) is None:
        raise ValueError(f"plot type '{kind}' needs '{key}'")
    return s[key]


def _series(s: dict, kind: str) -> dict:
    """series: {name: data}; a bare `values`/`y` list becomes one unnamed series."""
    if s.get("series") is not None:
        ser = s["series"]
        if not isinstance(ser, dict):
            raise ValueError(f"plot '{kind}': series must be a mapping of name -> data")
        return {str(k): v for k, v in ser.items()}
    for key in ("values", "y"):
        if s.get(key) is not None:
            return {"": s[key]}
    raise ValueError(f"plot type '{kind}' needs 'series'")


# ------------------------------------------------------------------------------------------
# bar
# ------------------------------------------------------------------------------------------
def _bar(fig, ax, s, c, P, size):
    cats = [str(x) for x in _need(s, "categories", "bar")]
    series = _series(s, "bar")
    names, n, k = list(series), len(series), len(cats)
    horiz, stacked = bool(s.get("horizontal")), bool(s.get("stacked"))
    width = float(s.get("bar_width", 0.8))
    fills = P.cycle(c.get("series"), n, [f"accent{i}" for i in range(1, 7)])
    hatch = bool(s.get("hatch"))
    edge = P(c.get("edge"), "bg1" if hatch else "none")  # hatches are drawn in the edge colour
    errs = s.get("errors") or {}
    ecolor = P(c.get("error"), "tx1")
    vcolor = P(c.get("value"), "tx1")
    fmt = s.get("value_format", "{:g}")
    x, base = np.arange(k), np.zeros(k)
    for i, name in enumerate(names):
        vals = np.asarray(series[name], float)
        if len(vals) != k:
            raise ValueError(f"bar series '{name}' has {len(vals)} values for {k} categories")
        if stacked:
            pos, w, start = x, width, base.copy()
        else:
            w = width / n
            pos, start = x - width / 2 + w * (i + 0.5), None
        err = errs.get(name)
        kw = dict(color=fills[i], edgecolor=edge, linewidth=0 if edge == "none" else 0.8, label=name or None,
                  hatch=HATCHES[i % len(HATCHES)] if hatch else None,
                  error_kw=dict(ecolor=ecolor, elinewidth=1, capsize=3, capthick=1))
        if horiz:
            bars = ax.barh(pos, vals, height=w, left=start, xerr=err, **kw)
        else:
            bars = ax.bar(pos, vals, width=w, bottom=start, yerr=err, **kw)
        if stacked:
            base += vals
        if s.get("values"):
            ax.bar_label(bars, fmt=fmt, padding=2, color=vcolor, fontsize=size * 0.8,
                         label_type="center" if stacked else "edge")
    if horiz:
        ax.set_yticks(x, cats)
        ax.invert_yaxis()  # first category on top, as read
    else:
        ax.set_xticks(x, cats)
    return sum(1 for nm in names if nm)


# ------------------------------------------------------------------------------------------
# box
# ------------------------------------------------------------------------------------------
def _box(fig, ax, s, c, P, size):
    horiz = bool(s.get("horizontal"))
    if s.get("stats") is not None:  # summary statistics only: {name: {q1, med, q3, whislo, whishi, ...}}
        stats = []
        for name, d in s["stats"].items():
            d = dict(d)
            st = {"label": str(name), "med": d.get("med", d.get("median")), "q1": d["q1"], "q3": d["q3"],
                  "whislo": d.get("whislo", d.get("min")), "whishi": d.get("whishi", d.get("max")),
                  "fliers": d.get("fliers", [])}
            if d.get("mean") is not None:
                st["mean"] = d["mean"]
            stats.append(st)
        data = None
    else:
        groups = _need(s, "groups", "box")
        data = [np.asarray(v, float) for v in groups.values()]
        stats = boxplot_stats(data, whis=float(s.get("whis", 1.5)), labels=[str(k) for k in groups])
    n = len(stats)
    palette = P.cycle(c.get("series"), n, [f"accent{i}" for i in range(1, 7)])
    fills = P.cycle(c.get("box"), n, [[p, 0.55] for p in palette])
    edges = P.cycle(c.get("box_edge"), n, palette)
    whisk = P.cycle(c.get("whisker"), n, palette)
    caps = P.cycle(c.get("cap"), n, palette)
    meds = P.cycle(c.get("median"), n, [[p, -0.35] for p in palette])
    means = P.cycle(c.get("mean"), n, "tx1")
    fliers = P.cycle(c.get("flier"), n, palette)
    notch = bool(s.get("notch")) and data is not None
    kw = dict(positions=range(1, n + 1), widths=float(s.get("box_width", 0.55)), patch_artist=True,
              showmeans=bool(s.get("means")), shownotches=notch,
              showfliers=s.get("fliers", not (s.get("points") and data is not None)),  # points already show them
              medianprops=dict(linewidth=2), whiskerprops=dict(linewidth=1.2), capprops=dict(linewidth=1.2),
              flierprops=dict(marker="o", markersize=4, markerfacecolor="none"),
              meanprops=dict(marker="D", markersize=5))
    try:
        art = ax.bxp(stats, orientation="horizontal" if horiz else "vertical", **kw)
    except TypeError:  # matplotlib < 3.10
        art = ax.bxp(stats, vert=not horiz, **kw)
    for i, b in enumerate(art["boxes"]):
        b.set_facecolor(fills[i])
        b.set_edgecolor(edges[i])
        b.set_linewidth(1.2)
    for i in range(n):
        for w in art["whiskers"][2 * i:2 * i + 2]:
            w.set_color(whisk[i])
        for cp in art["caps"][2 * i:2 * i + 2]:
            cp.set_color(caps[i])
        art["medians"][i].set_color(meds[i])
        if art["fliers"]:
            art["fliers"][i].set_markeredgecolor(fliers[i])
        if art["means"]:
            art["means"][i].set_markerfacecolor(means[i])
            art["means"][i].set_markeredgecolor(means[i])
    if s.get("points") and data is not None:  # raw observations as a jittered strip
        pts = P.cycle(c.get("points"), n, [[p, -0.3] for p in palette])
        rng = np.random.default_rng(0)  # fixed seed: rebuilding gives the same figure
        for i, vals in enumerate(data):
            jit = (i + 1) + rng.uniform(-0.12, 0.12, len(vals))
            xy = (vals, jit) if horiz else (jit, vals)
            ax.scatter(*xy, s=size * 0.9, color=pts[i], alpha=0.55, linewidths=0, zorder=3)
    if horiz:
        ax.invert_yaxis()
    return 0


# ------------------------------------------------------------------------------------------
# line
# ------------------------------------------------------------------------------------------
def _line(fig, ax, s, c, P, size):
    series = _series(s, "line")
    n = len(series)
    colors = P.cycle(c.get("series"), n, [f"accent{i}" for i in range(1, 7)])
    xs_all = s.get("x")
    labels = None
    if xs_all is not None and any(isinstance(v, str) for v in xs_all):  # categorical x
        labels, xs_all = [str(v) for v in xs_all], list(range(len(xs_all)))
    lw, ms = float(s.get("line_width", 2.0)), float(s.get("marker_size", 5.5))
    for i, (name, v) in enumerate(series.items()):
        d = v if isinstance(v, dict) else {"y": v}
        y = np.asarray(_need(d, "y", "line"), float)
        x = np.asarray(d.get("x", xs_all if xs_all is not None else range(len(y))), float)
        if len(x) != len(y):
            raise ValueError(f"line series '{name}': {len(x)} x values but {len(y)} y values")
        col = colors[i]
        marker = d.get("marker", MARKERS[i % len(MARKERS)] if s.get("markers", True) else None)
        style = d.get("style", DASHES[i % len(DASHES)] if s.get("dashes") else "-")
        ax.plot(x, y, color=col, linestyle=style, marker=marker, linewidth=lw, markersize=ms,
                markerfacecolor=P(c.get("marker_face"), col), markeredgecolor=P(c.get("marker_edge"), col),
                label=name or None, zorder=3)
        if d.get("err") is not None or d.get("lower") is not None:
            err = np.asarray(d["err"], float) if d.get("err") is not None else None
            lo = y - err if err is not None else np.asarray(d["lower"], float)
            hi = y + err if err is not None else np.asarray(_need(d, "upper", "line"), float)
            if s.get("errors", "band") == "band":
                ax.fill_between(x, lo, hi, color=P(c.get("band"), col), alpha=float(s.get("band_alpha", 0.18)),
                                linewidth=0, zorder=2)
            else:
                ax.errorbar(x, y, yerr=[y - lo, hi - y], fmt="none", ecolor=P(c.get("error"), col),
                            elinewidth=1, capsize=3, capthick=1, zorder=2)
    if labels:
        ax.set_xticks(xs_all, labels)
    return sum(1 for nm in series if nm)


# ------------------------------------------------------------------------------------------
# scatter
# ------------------------------------------------------------------------------------------
def _scatter(fig, ax, s, c, P, size):
    series = _series(s, "scatter")
    n = len(series)
    colors = P.cycle(c.get("series"), n, [f"accent{i}" for i in range(1, 7)])
    edge = P(c.get("edge"), "bg1")
    area = float(s.get("marker_size", (size * 0.55) ** 2))  # marker area in pt^2, scaled to the text
    alpha = float(s.get("alpha", 0.85))
    for i, (name, v) in enumerate(series.items()):
        if isinstance(v, dict):
            x, y = np.asarray(_need(v, "x", "scatter"), float), np.asarray(_need(v, "y", "scatter"), float)
            sizes = np.asarray(v["size"], float) if v.get("size") is not None else area
        else:
            pts = np.asarray(v, float).reshape(-1, 2)
            x, y, sizes = pts[:, 0], pts[:, 1], area
        col = colors[i]
        label = name or None
        if s.get("trend") and len(x) > 1:
            # least squares in the axes' own space, so the trend is a straight line on log axes too
            logx, logy = s.get("xscale") == "log", s.get("yscale") == "log"
            tx, ty = (np.log10(x) if logx else x), (np.log10(y) if logy else y)
            k, b0 = np.polyfit(tx, ty, 1)
            ss = ((ty - ty.mean()) ** 2).sum()
            r2 = 1 - ((ty - (k * tx + b0)) ** 2).sum() / ss if ss else 1.0
            if s.get("trend_label", True) and label:
                label = f"{label} ($R^2$ = {r2:.2f})"
            xx = np.linspace(tx.min(), tx.max(), 50)
            yy = k * xx + b0
            ax.plot(10 ** xx if logx else xx, 10 ** yy if logy else yy, color=P(c.get("trend"), [col, -0.25]),
                    linestyle="--", linewidth=1.5, zorder=4)
        ax.scatter(x, y, s=sizes, color=col, alpha=alpha, edgecolors=edge, linewidths=0.6, label=label,
                   marker=MARKERS[i % len(MARKERS)] if s.get("markers") else "o", zorder=3)
    return sum(1 for nm in series if nm)


# ------------------------------------------------------------------------------------------
# confusion matrix
# ------------------------------------------------------------------------------------------
def _confusion(fig, ax, s, c, P, size):
    m = np.asarray(_need(s, "matrix", "confusion"), float)
    if m.ndim != 2 or m.shape[0] != m.shape[1]:
        raise ValueError(f"confusion matrix must be square, got shape {m.shape}")
    k = m.shape[0]
    labels = [str(x) for x in (s.get("labels") or range(k))]
    norm = str(s.get("normalize", "none")).lower()
    with np.errstate(invalid="ignore", divide="ignore"):
        if norm in ("true", "rows", "recall"):
            vals = m / m.sum(axis=1, keepdims=True)
        elif norm in ("pred", "cols", "precision"):
            vals = m / m.sum(axis=0, keepdims=True)
        elif norm == "all":
            vals = m / m.sum()
        else:
            vals = m
    vals = np.nan_to_num(vals)
    if c.get("cmap"):
        cmap = matplotlib.colormaps[c["cmap"]]
    else:
        cmap = LinearSegmentedColormap.from_list("deck", [P(c.get("low"), "bg1"), P(c.get("high"), "accent1")])
    vmax = 1.0 if norm != "none" else (vals.max() or 1.0)
    im = ax.imshow(vals, cmap=cmap, vmin=0, vmax=vmax, aspect="equal" if s.get("square", True) else "auto")
    if s.get("values", True):
        fmt = s.get("value_format", "{:.0f}" if norm == "none" else "{:.2f}")
        pair = sorted([P("bg1"), P("tx1")], key=_lum)  # on a dark template the text is the light one
        light, dark = P(c.get("text_light"), pair[1]), P(c.get("text_dark"), pair[0])
        for r in range(k):
            for q in range(k):
                ax.text(q, r, fmt.format(vals[r, q]), ha="center", va="center", fontsize=size * 0.9,
                        color=light if _lum(cmap(vals[r, q] / vmax)) < 0.5 else dark)
    ax.set_xticks(range(k), labels)
    ax.set_yticks(range(k), labels)
    edge = P(c.get("cell_edge"), "bg1")
    if edge != "none":  # gaps between cells
        ax.set_xticks(np.arange(k + 1) - 0.5, minor=True)
        ax.set_yticks(np.arange(k + 1) - 0.5, minor=True)
        ax.grid(which="minor", color=edge, linewidth=2)
        ax.tick_params(which="minor", length=0)
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.tick_params(length=0)
    if s.get("colorbar", True):
        cb = fig.colorbar(im, ax=ax, fraction=0.05, pad=0.03)
        cb.outline.set_visible(False)
        cb.ax.tick_params(colors=P(c.get("axis"), ["tx1", 0.35]), labelcolor=P(c.get("tick"), c.get("text", "tx1")),
                          labelsize=size * 0.8, length=2)
    s.setdefault("xlabel", "Predicted")
    s.setdefault("ylabel", "True")
    return 0


# ------------------------------------------------------------------------------------------
# histogram, heatmap
# ------------------------------------------------------------------------------------------
def _hist(fig, ax, s, c, P, size):
    series = {k: np.asarray(v, float) for k, v in _series(s, "hist").items()}
    n = len(series)
    colors = P.cycle(c.get("series"), n, [f"accent{i}" for i in range(1, 7)])
    stacked = bool(s.get("stacked"))
    alpha = float(s.get("alpha", 0.9 if n == 1 or stacked else 0.6))
    bins = s.get("bins", "auto")
    data = [v[~np.isnan(v)] for v in series.values()]
    if not all(len(d) for d in data):
        raise ValueError("plot 'hist': every series needs at least one number")
    kw = dict(bins=bins, density=bool(s.get("density")), cumulative=bool(s.get("cumulative")),
              range=tuple(s["range"]) if s.get("range") else None, edgecolor=P(c.get("edge"), "bg1"), linewidth=0.6)
    if stacked:
        ax.hist(data, stacked=True, color=colors, alpha=alpha, label=[k or None for k in series], **kw)
    else:
        for i, (name, d) in enumerate(zip(series, data)):
            ax.hist(d, color=colors[i], alpha=alpha, label=name or None, **kw)
    return sum(1 for nm in series if nm)


def _cell_texts(ax, vals, cmap, norm, fmt, size, c, P):
    """Values in the cells, light on dark cells and dark on light ones (any template)."""
    pair = sorted([P("bg1"), P("tx1")], key=_lum)
    light, dark = P(c.get("text_light"), pair[1]), P(c.get("text_dark"), pair[0])
    for r in range(vals.shape[0]):
        for q in range(vals.shape[1]):
            ax.text(q, r, fmt.format(vals[r, q]), ha="center", va="center", fontsize=size * 0.85,
                    color=light if _lum(cmap(norm(vals[r, q]))) < 0.5 else dark)


def _heatmap(fig, ax, s, c, P, size):
    from matplotlib.colors import Normalize, TwoSlopeNorm
    m = np.asarray(_need(s, "matrix", "heatmap"), float)
    if m.ndim != 2:
        raise ValueError(f"heatmap matrix must be a table of numbers (rows of equal length), got shape {m.shape}")
    rows, cols = m.shape
    xl = s.get("xlabels") or s.get("labels")
    yl = s.get("ylabels") or s.get("labels")
    for name, lab, k in (("xlabels", xl, cols), ("ylabels", yl, rows)):
        if lab is not None and len(lab) != k:
            raise ValueError(f"heatmap {name} has {len(lab)} labels for {k} {'columns' if k == cols else 'rows'}")
    vmin = float(s["vmin"]) if s.get("vmin") is not None else float(np.nanmin(m))
    vmax = float(s["vmax"]) if s.get("vmax") is not None else float(np.nanmax(m))
    if vmax <= vmin:
        vmax = vmin + 1.0
    if s.get("center") is not None:  # diverging: alert colour below the centre, structure colour above
        ctr = float(s["center"])
        cmap = LinearSegmentedColormap.from_list("deck", [P(c.get("low"), "accent2"), P(c.get("mid"), "bg1"),
                                                          P(c.get("high"), "accent1")])
        norm = TwoSlopeNorm(vcenter=ctr, vmin=min(vmin, ctr - 1e-9), vmax=max(vmax, ctr + 1e-9))
    else:
        cmap = (matplotlib.colormaps[c["cmap"]] if c.get("cmap") else
                LinearSegmentedColormap.from_list("deck", [P(c.get("low"), "bg1"), P(c.get("high"), "accent1")]))
        norm = Normalize(vmin=vmin, vmax=vmax)
    im = ax.imshow(m, cmap=cmap, norm=norm, aspect="equal" if s.get("square") else "auto")
    if s.get("values", rows * cols <= 144):
        _cell_texts(ax, m, cmap, norm, s.get("value_format", "{:.2g}"), size, c, P)
    ax.set_xticks(range(cols), [str(x) for x in xl] if xl is not None else range(cols))
    ax.set_yticks(range(rows), [str(x) for x in yl] if yl is not None else range(rows))
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.tick_params(length=0)
    if s.get("colorbar", True):
        cb = fig.colorbar(im, ax=ax, fraction=0.05, pad=0.03)
        cb.outline.set_visible(False)
        cb.ax.tick_params(colors=P(c.get("axis"), ["tx1", 0.35]), labelcolor=P(c.get("tick"), c.get("text", "tx1")),
                          labelsize=size * 0.8, length=2)
    return 0


# ------------------------------------------------------------------------------------------
# function: formulas, evaluated safely (arithmetic and maths functions only)
# ------------------------------------------------------------------------------------------
def _vec(f):
    return np.vectorize(f, otypes=[float])


FUNCS = {
    "sin": np.sin, "cos": np.cos, "tan": np.tan, "arcsin": np.arcsin, "asin": np.arcsin, "arccos": np.arccos,
    "acos": np.arccos, "arctan": np.arctan, "atan": np.arctan, "arctan2": np.arctan2, "atan2": np.arctan2,
    "sinh": np.sinh, "cosh": np.cosh, "tanh": np.tanh, "exp": np.exp, "log": np.log, "ln": np.log,
    "log2": np.log2, "log10": np.log10, "log1p": np.log1p, "sqrt": np.sqrt, "abs": np.abs, "sign": np.sign,
    "floor": np.floor, "ceil": np.ceil, "round": np.round, "max": np.maximum, "min": np.minimum,
    "maximum": np.maximum, "minimum": np.minimum, "where": np.where, "clip": np.clip, "mod": np.mod,
    "sigmoid": lambda z: 1 / (1 + np.exp(-z)), "relu": lambda z: np.maximum(z, 0),
    "softplus": lambda z: np.logaddexp(0, z), "heaviside": lambda z: np.heaviside(z, 0.5),
    "erf": _vec(math.erf), "gamma": _vec(math.gamma),
    "normal": lambda z, mu=0.0, sd=1.0: np.exp(-0.5 * ((z - mu) / sd) ** 2) / (sd * np.sqrt(2 * np.pi)),
}
CONSTS = {"pi": np.pi, "e": np.e, "inf": np.inf, "tau": 2 * np.pi}
_BIN = {ast.Add: np.add, ast.Sub: np.subtract, ast.Mult: np.multiply, ast.Div: np.divide, ast.Pow: np.power,
        ast.Mod: np.mod, ast.FloorDiv: np.floor_divide}
_CMP = {ast.Lt: np.less, ast.LtE: np.less_equal, ast.Gt: np.greater, ast.GtE: np.greater_equal,
        ast.Eq: np.equal, ast.NotEq: np.not_equal}


def formula(expr: str, env: dict):
    """Value of a formula like "x^2 + 2*x" or "where(x > 0, x, 0)" for the names in env (numbers/arrays).
    Raises ValueError for anything that is not arithmetic, a comparison or a function of FUNCS."""
    text = str(expr)
    if len(text) > 500:
        raise ValueError("formula is longer than 500 characters")
    try:  # ^ means power, as in maths (and binds as tightly as **)
        tree = ast.parse(text.replace("^", "**"), mode="eval")
    except SyntaxError as exc:
        raise ValueError(f"cannot read formula {text!r}: {exc.msg}") from None

    def ev(n):
        if isinstance(n, ast.Expression):
            return ev(n.body)
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)) and not isinstance(n.value, bool):
            return np.float64(n.value)
        if isinstance(n, ast.Name):
            if n.id in env:
                return env[n.id]
            if n.id in CONSTS:
                return CONSTS[n.id]
            raise ValueError(f"unknown name '{n.id}' in formula {text!r}; use x, pi, e, params or a function")
        if isinstance(n, ast.BinOp) and type(n.op) in _BIN:
            return _BIN[type(n.op)](ev(n.left), ev(n.right))
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, (ast.USub, ast.UAdd)):
            v = ev(n.operand)
            return -v if isinstance(n.op, ast.USub) else v
        if isinstance(n, ast.Compare):
            left, out = ev(n.left), True
            for op, right in zip(n.ops, n.comparators):
                if type(op) not in _CMP:
                    raise ValueError(f"comparison not allowed in formula {text!r}")
                r = ev(right)
                out = np.logical_and(out, _CMP[type(op)](left, r))
                left = r
            return out
        if isinstance(n, ast.BoolOp):
            vals = [ev(v) for v in n.values]
            fn = np.logical_and if isinstance(n.op, ast.And) else np.logical_or
            out = vals[0]
            for v in vals[1:]:
                out = fn(out, v)
            return out
        if isinstance(n, ast.IfExp):  # "x if x > 0 else 0"
            return np.where(ev(n.test), ev(n.body), ev(n.orelse))
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and not n.keywords:
            if n.func.id not in FUNCS:
                raise ValueError(f"unknown function '{n.func.id}' in formula {text!r}; use one of "
                                 f"{', '.join(sorted(FUNCS))}")
            return FUNCS[n.func.id](*[ev(a) for a in n.args])
        raise ValueError(f"not allowed in a formula: {ast.dump(n)[:60]} (in {text!r})")

    with np.errstate(all="ignore"):
        return ev(tree)


def _function(fig, ax, s, c, P, size):
    funcs = s.get("functions")
    if funcs is None and s.get("f") is not None:
        funcs = {"": s["f"]}
    if not isinstance(funcs, dict) or not funcs:
        raise ValueError("plot type 'function' needs f: 'formula in x' or functions: {label: 'formula', ...}")
    dom = s.get("x", [-5, 5])
    if not (isinstance(dom, (list, tuple)) and len(dom) == 2):
        raise ValueError("plot 'function': x must be the range [from, to], e.g. x: [-6, 6]")
    lo, hi = float(dom[0]), float(dom[1])
    if not hi > lo:
        raise ValueError("plot 'function': x must be [from, to] with from < to")
    n = max(10, min(int(s.get("samples", 400)), 5000))
    log = s.get("xscale") == "log"
    if log and lo <= 0:
        raise ValueError("plot 'function': a log x axis needs x values above 0")
    x = np.geomspace(lo, hi, n) if log else np.linspace(lo, hi, n)
    params = s.get("params") or {}
    if not isinstance(params, dict):
        raise ValueError("plot 'function': params must be a mapping, e.g. params: {k: 2}")
    env = {str(k): np.float64(float(v)) for k, v in params.items()}
    env["x"] = x
    colors = P.cycle(c.get("series"), len(funcs), [f"accent{i}" for i in range(1, 7)])
    lw = float(s.get("line_width", 2.0))
    for i, (name, f) in enumerate(funcs.items()):
        d = f if isinstance(f, dict) else {"f": f}
        if d.get("f") is None:
            raise ValueError(f"plot 'function': '{name}' needs a formula, e.g. {name}: 'x^2'")
        y = np.broadcast_to(np.asarray(formula(d["f"], env), float), x.shape).copy()
        y[~np.isfinite(y)] = np.nan
        style = d.get("style", DASHES[i % len(DASHES)] if s.get("dashes") else "-")
        ax.plot(x, y, color=P(d.get("color"), colors[i]), linestyle=style, linewidth=lw,
                label=str(name) or None, zorder=3)
    return sum(1 for nm in funcs if str(nm))


# ------------------------------------------------------------------------------------------
# python: your own drawing code, in the template's fonts, colours and the slide's exact size
# ------------------------------------------------------------------------------------------
def _python(fig, ax, s, c, P, size):
    code, script = s.get("code"), s.get("script")
    if bool(code) == bool(script):
        raise ValueError("plot type 'python' needs code: | (inline) or script: file.py - one of the two")
    name = "<plot code>"
    if script:
        name = str(script)
        path = name if os.path.isabs(name) else os.path.join(s.get("_base_dir", "."), name)
        if not os.path.exists(path):
            raise ValueError(f"plot script not found: {script} (relative to the YAML)")
        with open(path, encoding="utf-8") as fh:
            code = fh.read()
    ns = {"fig": fig, "ax": ax, "np": np, "data": s.get("data"), "size": size,
          "colors": P.cycle(c.get("series"), 6, [f"accent{i}" for i in range(1, 7)]),
          "color": P, "tint": lambda col, amount: tint(P(col), amount, P.bg), "__name__": "__plot__"}
    exec(compile(str(code), name, "exec"), ns)  # noqa: S102 - only with --allow-python
    for extra in fig.axes:  # axes the code added get the same quiet styling
        if extra is not ax and not hasattr(extra, "_colorbar"):
            for sp in extra.spines.values():
                sp.set_color(P(c.get("axis"), ["tx1", 0.35]))
            extra.tick_params(color=P(c.get("axis"), ["tx1", 0.35]), labelcolor=P(c.get("tick"), "tx1"))
    return len(ax.get_legend_handles_labels()[1])


DRAW = {"bar": _bar, "box": _box, "line": _line, "scatter": _scatter, "hist": _hist, "heatmap": _heatmap,
        "confusion": _confusion, "function": _function, "python": _python}


# ------------------------------------------------------------------------------------------
# shared styling: axes, labels, grid, legend
# ------------------------------------------------------------------------------------------
def _style(fig, ax, s, c, P, kind, n_labeled):
    text = P(c.get("text"), "tx1")
    axis = P(c.get("axis"), ["tx1", 0.35])
    ax.set_facecolor(P(c.get("background"), "none"))
    for sp in ax.spines.values():
        sp.set_color(axis)
    ax.tick_params(color=axis, labelcolor=P(c.get("tick"), text))
    if s.get("title"):
        ax.set_title(str(s["title"]), color=P(c.get("title"), text))
    for key, setter, lab in (("xlabel", ax.set_xlabel, ax.xaxis.label), ("ylabel", ax.set_ylabel, ax.yaxis.label)):
        if s.get(key):
            setter(str(s[key]))
            lab.set_color(P(c.get("label"), text))
    for key, setter in (("xlim", ax.set_xlim), ("ylim", ax.set_ylim)):
        if s.get(key) is not None:
            setter(*s[key])
    for key, setter in (("xscale", ax.set_xscale), ("yscale", ax.set_yscale)):
        if s.get(key):
            setter(s[key])
    for key, axis_ in (("xformat", ax.xaxis), ("yformat", ax.yaxis)):
        if s.get(key):  # e.g. "{x:.0%}" or "{x:,.0f}"
            axis_.set_major_formatter(StrMethodFormatter(str(s[key])))
    if s.get("xticks_rotation"):
        for t in ax.get_xticklabels():
            t.set_rotation(float(s["xticks_rotation"]))
            t.set_ha("right")
            t.set_rotation_mode("anchor")
    if kind not in ("confusion", "heatmap"):
        horiz = bool(s.get("horizontal"))
        default_grid = ("x" if horiz else "y") if kind in ("bar", "box", "hist") else "both"
        grid = str(s.get("grid", default_grid)).lower()
        if grid in ("x", "y", "both"):
            ax.grid(True, axis=grid, color=P(c.get("grid"), ["tx1", 0.85]), linewidth=0.6)
        else:
            ax.grid(False)
    _legend(fig, ax, s, c, P, n_labeled, text)


def _legend(fig, ax, s, c, P, n_labeled, text):
    pos = s.get("legend")
    if pos is None:
        pos = "best" if n_labeled >= 2 else "none"
    if pos in (False, "none") or n_labeled == 0:
        return
    handles, labels = ax.get_legend_handles_labels()
    kw = dict(labelcolor=P(c.get("legend_text"), text))
    if c.get("legend_frame") or c.get("legend_background"):
        kw.update(frameon=True, edgecolor=P(c.get("legend_frame"), "none"),
                  facecolor=P(c.get("legend_background"), "none"), framealpha=1)
    outside = {"top": "outside upper center", "bottom": "outside lower center", "right": "outside right upper"}
    if pos in outside:
        leg = fig.legend(handles, labels, loc=outside[pos], ncols=len(labels) if pos != "right" else 1, **kw)
    else:  # best, upper left, lower right, ...
        leg = ax.legend(handles, labels, loc=pos, **kw)
    leg.set_zorder(5)
