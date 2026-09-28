"""
plots.py - academic-style figures for ppt-templated: bar, box, line, scatter and confusion matrix.

ppt-templated calls render() for every `plot:` item. The figure is drawn at the exact size it gets on
the slide, in the template's body font and theme colours, and embedded straight into the deck as
a 300 dpi PNG (nothing is written to disk). Use `chart:` instead when you need a native,
editable PowerPoint chart; use `plot:` for box plots, confusion matrices, error bands, trend
lines and print-quality styling.

Colour specs, accepted by every colour setting:
    accent1 ... accent6, tx1, bg1, tx2, bg2    theme slots of the template
    structure | alert | example | muted        Beamer roles (accent1 | accent2 | accent3 | grey text)
    [accent1, 0.6]                             lightened 60% towards white (negative = darker)
    "#1F3A5F", black, tab:blue                 anything matplotlib understands
Settings that colour several things (series, boxes, ...) take one colour or a list (cycled).

All plot settings are documented in YAML_GUIDE.md (section 7). When you add or change a setting
or plot type here, update that guide in the same commit.
"""

from __future__ import annotations

import io

import matplotlib
import numpy as np
from matplotlib import font_manager
from matplotlib.cbook import boxplot_stats
from matplotlib.colors import LinearSegmentedColormap, to_hex, to_rgb
from matplotlib.figure import Figure
from matplotlib.ticker import StrMethodFormatter

PLOT_TYPES = ("bar", "box", "line", "scatter", "confusion")
HATCHES = ["", "//", "..", "xx", "\\\\", "oo", "++", "--"]
MARKERS = ["o", "s", "^", "D", "v", "P", "X", "*"]
DASHES = ["-", "--", "-.", ":"]
ALIASES = {"structure": "accent1", "alert": "accent2", "example": "accent3",
           "tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}
SERIF_FALLBACK = ["Cambria", "Times New Roman", "STIXGeneral", "DejaVu Serif"]


# ------------------------------------------------------------------------------------------
# colours and fonts
# ------------------------------------------------------------------------------------------
def tint(color, amount: float) -> str:
    """amount > 0 blends towards white, < 0 towards black (like PowerPoint's brightness)."""
    rgb = np.array(to_rgb(color))
    rgb = rgb + (1 - rgb) * amount if amount >= 0 else rgb * (1 + amount)
    return to_hex(np.clip(rgb, 0, 1))


class Palette:
    """Resolves colour specs against the template's theme colours."""

    def __init__(self, theme_colors: dict):
        self.theme = {k: "#" + v for k, v in theme_colors.items()}

    @staticmethod
    def _is_tint(spec) -> bool:
        return (isinstance(spec, (list, tuple)) and len(spec) == 2 and isinstance(spec[0], str)
                and isinstance(spec[1], (int, float)))

    def __call__(self, spec, default=None):
        spec = default if spec is None else spec
        if spec is None:
            return None
        if self._is_tint(spec):
            return tint(self(spec[0]), float(spec[1]))
        if isinstance(spec, str):
            if spec in ("none", "transparent"):
                return "none"
            if spec == "muted":
                return tint(self("tx1"), 0.45)
            key = ALIASES.get(spec, spec)
            if key in self.theme:
                return self.theme[key]
        return spec

    def cycle(self, spec, n: int, default) -> list:
        """n colours from one colour or a list of colours."""
        spec = default if spec is None else spec
        cols = [self(x) for x in spec] if isinstance(spec, (list, tuple)) and not self._is_tint(spec) else [self(spec)]
        return [cols[i % len(cols)] for i in range(n)]


def pick_font(*candidates) -> str:
    """First installed font; the fallbacks keep a serif look when the template font is missing."""
    for name in [*candidates, *SERIF_FALLBACK]:
        if not name:
            continue
        try:
            font_manager.findfont(font_manager.FontProperties(family=name), fallback_to_default=False)
            return name
        except ValueError:
            continue
    return "DejaVu Serif"


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
def render(spec: dict, theme_colors: dict, theme_font: str | None, width_in: float, height_in: float,
           font_size: float) -> bytes:
    """Draw one `plot:` item at width_in x height_in inches and return PNG bytes."""
    spec = dict(spec)  # drawing fills in defaults (e.g. confusion axis labels); keep the content untouched
    kind = spec.get("type")
    if kind not in PLOT_TYPES:
        raise ValueError(f"plot type must be one of {', '.join(PLOT_TYPES)}, got {kind!r}")
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
    return buf.getvalue()


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
        light, dark = P(c.get("text_light"), "bg1"), P(c.get("text_dark"), "tx1")
        for r in range(k):
            for q in range(k):
                rgb = cmap(vals[r, q] / vmax)[:3]
                lum = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]
                ax.text(q, r, fmt.format(vals[r, q]), ha="center", va="center", fontsize=size * 0.9,
                        color=light if lum < 0.5 else dark)
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


DRAW = {"bar": _bar, "box": _box, "line": _line, "scatter": _scatter, "confusion": _confusion}


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
    if kind != "confusion":
        horiz = bool(s.get("horizontal"))
        default_grid = ("x" if horiz else "y") if kind in ("bar", "box") else "both"
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
