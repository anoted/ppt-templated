"""
flows.py - layout engine for the `flow:` (flowchart / workflow) and `steps:` items of ppt-templated.

Pure geometry, no PowerPoint objects. It parses the YAML spec, sizes every node from its text,
places the nodes in steps along the flow direction and in lanes across it, routes every edge as an
orthogonal path between real connection sites of the two shapes, and turns each path into the
preset connector PowerPoint itself uses (straight or bent, with rotation, flips and adjust values).
ppt_templated draws the result as native shapes plus connectors glued to them, so dragging a shape
in PowerPoint keeps its arrows attached.

All lengths are in em (multiples of the node font size), so a diagram scales as one piece.
All settings are documented in YAML_GUIDE.md (section 6.12 / 6.13). When you add or change a
setting, shape or rule here, update that guide in the same commit.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

CHAR = 0.55        # average character width (em); a little generous, so text never overflows its shape
CHAR_BOLD = 0.6
LINE = 1.2         # line height (em)
LABEL = 0.8        # edge and group label size, relative to the node text
STUB = 0.9         # shortest straight piece where a connector leaves or enters a shape (em)
CLEAR = 1.0        # distance between the outermost shape and a connector routed around it (em)
STAGGER = 0.7      # distance between two connectors routed around on the same side (em)
MARGIN = 0.3       # a connector must pass at least this far from shapes it does not connect (em)
GPAD = 0.6         # group box padding around its members (em)
EPS = 1e-6

# Connection sites, measured in PowerPoint (Shape.ConnectionSiteCount + ConnectorFormat.BeginConnect):
# side -> (OOXML site idx, x fraction, y fraction of the shape box).
S4 = {"t": (0, .5, 0), "l": (1, 0, .5), "b": (2, .5, 1), "r": (3, 1, .5)}
S8 = {"t": (0, .5, 0), "l": (2, 0, .5), "b": (4, .5, 1), "r": (6, 1, .5)}

# name -> (MSO_SHAPE member, text area as a fraction of the shape (w, h), text padding (em x, y), sites)
SHAPES = {
    "box":         ("ROUNDED_RECTANGLE", (1, 1), (.6, .35), S4),
    "rect":        ("RECTANGLE", (1, 1), (.6, .35), S4),
    "terminal":    ("FLOWCHART_TERMINATOR", (.9, .7), (.45, .2), S4),
    "decision":    ("FLOWCHART_DECISION", (.5, .5), (.15, .05), S4),
    "data":        ("FLOWCHART_DATA", (.6, 1), (.3, .3),
                    {"t": (1, .5, 0), "l": (2, .1, .5), "b": (4, .5, 1), "r": (5, .9, .5)}),
    "document":    ("FLOWCHART_DOCUMENT", (1, .8), (.5, .25), {**S4, "b": (2, .5, .934)}),
    "documents":   ("FLOWCHART_MULTIDOCUMENT", (.83, .63), (.4, .2),
                    {**S4, "t": (0, .569, 0), "b": (2, .43, .962)}),
    "database":    ("FLOWCHART_MAGNETIC_DISK", (1, .5), (.5, .1),
                    {"t": (1, .5, 0), "l": (2, 0, .5), "b": (3, .5, 1), "r": (4, 1, .5)}),
    "storage":     ("FLOWCHART_STORED_DATA", (.667, 1), (.3, .3), {**S4, "r": (3, .833, .5)}),
    "subprocess":  ("FLOWCHART_PREDEFINED_PROCESS", (.75, 1), (.35, .35), S4),
    "manual":      ("FLOWCHART_MANUAL_INPUT", (1, .8), (.5, .25), {**S4, "t": (0, .5, .1)}),
    "preparation": ("FLOWCHART_PREPARATION", (.6, 1), (.3, .3), S4),
    "delay":       ("FLOWCHART_DELAY", (.85, .7), (.4, .15), S4),
    "circle":      ("FLOWCHART_CONNECTOR", (.7, .7), (.1, .05), S8),
    "ellipse":     ("OVAL", (.7, .7), (.1, .05), S8),
    "text":        ("RECTANGLE", (1, 1), (.1, .05), S4),
    # steps: only
    "chevron":     ("CHEVRON", (1, 1), (.45, .35), S4),
    "pentagon":    ("PENTAGON", (1, 1), (.45, .35), S4),
}
SHAPE_ALIASES = {
    "rounded": "box", "process": "rect", "rectangle": "rect", "square": "rect",
    "start": "terminal", "end": "terminal", "pill": "terminal", "stadium": "terminal",
    "diamond": "decision", "if": "decision", "io": "data", "input": "data", "output": "data",
    "parallelogram": "data", "doc": "document", "docs": "documents", "db": "database",
    "cylinder": "database", "stored": "storage", "predefined": "subprocess", "manual_input": "manual",
    "hexagon": "preparation", "wait": "delay", "oval": "ellipse", "node": "circle",
    "label": "text", "none": "text",
}
FLOW_SHAPES = [s for s in SHAPES if s not in ("chevron", "pentagon")]
STYLES = ("tinted", "solid", "outline", "plain")
DIRECTIONS = {"right": "right", "lr": "right", "horizontal": "right", "down": "down", "td": "down",
              "tb": "down", "vertical": "down", "left": "left", "rl": "left", "up": "up", "bt": "up"}
PORTS = {"top": "t", "bottom": "b", "left": "l", "right": "r", "t": "t", "b": "b", "l": "l", "r": "r"}
DASHES = ("solid", "dashed", "dotted")
ARROWS = ("end", "start", "both", "none")

# abstract side of a node -> physical side, per direction. u runs along the flow, v across it.
PHYS = {
    "right": {"fwd": "r", "back": "l", "neg": "t", "pos": "b"},
    "down":  {"fwd": "b", "back": "t", "neg": "l", "pos": "r"},
    "left":  {"fwd": "l", "back": "r", "neg": "t", "pos": "b"},
    "up":    {"fwd": "t", "back": "b", "neg": "l", "pos": "r"},
}
DIRS = {"fwd": (1.0, 0.0), "back": (-1.0, 0.0), "neg": (0.0, -1.0), "pos": (0.0, 1.0)}
OPPOSITE = {"fwd": "back", "back": "fwd", "neg": "pos", "pos": "neg"}

NODE_KEYS = {"id", "text", "shape", "color", "style", "text_color", "bold", "step", "lane"}
EDGE_KEYS = {"from", "to", "label", "style", "color", "arrow", "exit", "enter", "width"}
GROUP_KEYS = {"label", "nodes", "color", "fill"}
STEP_KEYS = {"text", "color", "style", "text_color", "bold", "shape"}
EDGE_RE = re.compile(r"\s*(<->|-->|->|--)\s*")
EDGE_OPS = {"->": ("end", "solid"), "-->": ("end", "dashed"), "<->": ("both", "solid"), "--": ("none", "solid")}


@dataclass
class Node:
    id: str
    text: str
    shape: str = "box"
    color: object = "structure"
    style: str = "tinted"
    text_color: object = None
    bold: bool = False
    step: int | None = None
    lane: float | None = None
    index: int = 0
    w: float = 0.0      # physical size and top-left corner (em)
    h: float = 0.0
    x: float = 0.0
    y: float = 0.0
    pad: tuple = (0.0, 0.0)
    adj: float | None = None   # shape adjustment (chevron point depth)
    rank: int = 0       # step along the flow (0-based) - also the build order
    pos: float = 0.0    # lane across the flow
    u: float = 0.0      # centre in flow coordinates
    v: float = 0.0


@dataclass
class Edge:
    a: str
    b: str
    label: str | None = None
    dash: str = "solid"
    color: object = None
    arrow: str = "end"
    exit: str | None = None
    enter: str | None = None
    width: float | None = None
    pts: list = field(default_factory=list)   # physical polyline (em)
    sites: tuple = (0, 0)                     # connection-site idx on a and b
    label_box: tuple | None = None            # x, y, w, h (em)
    rank: int = 0


@dataclass
class Group:
    label: str
    members: list
    color: object = "tx1"
    fill: bool = True
    x: float = 0.0
    y: float = 0.0
    w: float = 0.0
    h: float = 0.0
    rank: int = 0


@dataclass
class Layout:
    w: float
    h: float
    nodes: list
    edges: list
    groups: list


# ------------------------------------------------------------------------------------------
# entry point
# ------------------------------------------------------------------------------------------
def layout(kind: str, spec: dict, visible) -> Layout:
    """kind: 'flow' or 'steps'. visible(text) -> the text as displayed (markup removed)."""
    if kind == "steps":
        return _steps(spec, visible)
    return _flow(spec, visible)


def shape_name(name) -> str:
    key = str(name).strip().lower()
    key = SHAPE_ALIASES.get(key, key)
    if key not in SHAPES:
        raise ValueError(f"unknown flow shape '{name}'. Use one of: {', '.join(FLOW_SHAPES)}")
    return key


def _label(v):
    if v is None:
        return None
    if v is True:   # YAML reads an unquoted yes / on / true as a boolean
        return "yes"
    if v is False:
        return "no"
    return str(v)


def _style(v) -> str:
    s = str(v).lower()
    if s not in STYLES:
        raise ValueError(f"flow style must be one of {', '.join(STYLES)}, got {v!r}")
    return s


def _block(text: str, limit: float, cw: float):
    """Greedy word wrap at `limit` em -> (width, height) of the text block in em."""
    lines = []
    for para in text.split("\n"):
        cur = ""
        for word in para.split():
            cand = f"{cur} {word}" if cur else word
            if cur and len(cand) * cw > limit:
                lines.append(cur)
                cur = word
            else:
                cur = cand
        lines.append(cur)
    return max(max(len(l) for l in lines) * cw, cw), len(lines) * LINE


def _size(n: Node, wrap: float, visible) -> None:
    _, (fx, fy), (px, py), _ = SHAPES[n.shape]
    text = visible(n.text)
    cw = CHAR_BOLD if n.bold else CHAR
    limit = wrap
    if n.shape in ("decision", "circle", "ellipse"):  # wrap into a compact block, not one long line
        aspect = 2.2 if n.shape == "decision" else 1.6
        longest = max((len(w) for w in text.split()), default=1) * cw
        limit = min(wrap, max(longest, math.sqrt(len(text) * cw * LINE * aspect)))
    tw, th = _block(text, limit, cw)
    n.pad = (px, py)
    n.w = (tw + 2 * px) / fx
    n.h = (th + 2 * py) / fy
    if n.shape not in ("text", "circle", "ellipse", "decision"):
        n.w = max(n.w, 3.5)
    if n.shape == "circle":
        n.w = n.h = max(n.w, n.h)


# ------------------------------------------------------------------------------------------
# parsing
# ------------------------------------------------------------------------------------------
def _parse_nodes(items, d: dict) -> list:
    if not isinstance(items, list):
        raise ValueError("flow 'nodes' must be a list")
    out = []
    for i, it in enumerate(items):
        if isinstance(it, dict) and "id" not in it and "text" not in it and len(it) == 1:
            (k, v), = it.items()  # shorthand:  - raw: Raw CSI   or   - raw: {text: ..., shape: data}
            it = {"id": k, **(v if isinstance(v, dict) else {"text": v})}
        elif not isinstance(it, dict):
            it = {"id": it, "text": it}
        bad = set(it) - NODE_KEYS
        if bad:
            raise ValueError(f"unknown flow node key(s) {', '.join(map(str, bad))}; use {', '.join(sorted(NODE_KEYS))}")
        nid = str(it.get("id", it.get("text", f"node{i + 1}")))
        shape = shape_name(it.get("shape", d["shape"]))
        style = _style(it.get("style", "plain" if shape == "text" else d["style"]))
        out.append(Node(nid, "" if it.get("text", nid) is None else str(it.get("text", nid)), shape,
                        it.get("color", d["color"]), style, it.get("text_color"), bool(it.get("bold", False)),
                        int(it["step"]) if it.get("step") is not None else None,
                        float(it["lane"]) if it.get("lane") is not None else None, index=i))
    ids = [n.id for n in out]
    dup = {x for x in ids if ids.count(x) > 1}
    if dup:
        raise ValueError(f"flow node id(s) used twice: {', '.join(sorted(dup))}")
    return out


def _parse_edges(items, nodes: list, byid: dict, d: dict) -> list:
    if not isinstance(items, list):
        raise ValueError("flow 'edges' must be a list")

    def ref(name):
        name = str(name).strip()
        if name not in byid:  # an edge may introduce a node: its id is also its text
            n = Node(name, name, shape_name(d["shape"]), d["color"], _style(d["style"]), index=len(nodes))
            nodes.append(n)
            byid[name] = n
        return name

    edges = []
    for it in items:
        if isinstance(it, dict) and ("from" in it or "to" in it):
            chain, ops, opts = [it.get("from"), it.get("to")], [None], it
        else:
            if isinstance(it, dict) and len(it) == 1:
                (expr, val), = it.items()  # - a -> b: label   or   - a -> b: {label: ..., style: dashed}
                opts = dict(val) if isinstance(val, dict) else {"label": val}
            elif isinstance(it, str):
                expr, opts = it, {}
            else:
                raise ValueError(f"cannot understand flow edge {it!r}; write 'a -> b' or {{from: a, to: b}}")
            parts = EDGE_RE.split(str(expr))
            chain, ops = parts[0::2], parts[1::2]
        if len(chain) < 2 or any(c is None or not str(c).strip() for c in chain):
            raise ValueError(f"flow edge needs two node ids, like 'a -> b', got {it!r}")
        bad = set(opts) - EDGE_KEYS
        if bad:
            raise ValueError(f"unknown flow edge key(s) {', '.join(map(str, bad))}; use {', '.join(sorted(EDGE_KEYS))}")
        for a, b, op in zip(chain, chain[1:], ops):
            arrow, dash = EDGE_OPS.get(op, ("end", "solid"))
            e = Edge(ref(a), ref(b), _label(opts.get("label")), str(opts.get("style", dash)).lower(),
                     opts.get("color"), str(opts.get("arrow", arrow)).lower(),
                     opts.get("exit"), opts.get("enter"),
                     float(opts["width"]) if opts.get("width") is not None else None)
            if e.dash not in DASHES:
                raise ValueError(f"flow edge style must be one of {', '.join(DASHES)}, got {e.dash!r}")
            if e.arrow not in ARROWS:
                raise ValueError(f"flow edge arrow must be one of {', '.join(ARROWS)}, got {e.arrow!r}")
            for key in ("exit", "enter"):
                val = getattr(e, key)
                if val is not None:
                    if str(val).lower() not in PORTS:
                        raise ValueError(f"flow edge {key} must be top, bottom, left or right, got {val!r}")
                    setattr(e, key, PORTS[str(val).lower()])
            if e.a == e.b:
                raise ValueError(f"flow edge '{e.a} -> {e.b}' connects a node to itself; that is not supported")
            edges.append(e)
    return edges


def _parse_groups(items, byid: dict) -> list:
    out = []
    for g in items or []:
        if not isinstance(g, dict) or not g.get("nodes"):
            raise ValueError(f"flow group needs 'nodes': [ids], got {g!r}")
        bad = set(g) - GROUP_KEYS
        if bad:
            raise ValueError(f"unknown flow group key(s) {', '.join(map(str, bad))}; use {', '.join(sorted(GROUP_KEYS))}")
        members = [str(x) for x in g["nodes"]]
        missing = [m for m in members if m not in byid]
        if missing:
            raise ValueError(f"flow group '{g.get('label', '')}' names unknown node(s): {', '.join(missing)}")
        out.append(Group(str(g.get("label") or ""), members, g.get("color", "tx1"), bool(g.get("fill", True))))
    return out


# ------------------------------------------------------------------------------------------
# flow layout
# ------------------------------------------------------------------------------------------
def _flow(spec: dict, visible) -> Layout:
    direction = DIRECTIONS.get(str(spec.get("direction", "right")).lower())
    if direction is None:
        raise ValueError(f"flow direction must be right, down, left or up, got {spec.get('direction')!r}")
    horiz = direction in ("right", "left")
    d = {"shape": spec.get("shape", "box"), "color": spec.get("color", "structure"), "style": spec.get("style", "tinted")}
    nodes = _parse_nodes(spec.get("nodes") or [], d)
    byid = {n.id: n for n in nodes}
    if "edges" in spec:
        edges = _parse_edges(spec.get("edges") or [], nodes, byid, d)
    else:  # no edges given: a linear chain in node order
        edges = [Edge(a.id, b.id) for a, b in zip(nodes, nodes[1:])]
    if not nodes:
        raise ValueError("flow needs 'nodes' (or 'edges' that name them)")
    groups = _parse_groups(spec.get("groups"), byid)
    wrap = float(spec.get("wrap", 18)) * CHAR
    gap = float(spec.get("gap", 1.0))
    for n in nodes:
        _size(n, wrap, visible)
    if spec.get("uniform", True):  # same-shape nodes get the same size, like a hand-made flowchart
        for shape in {n.shape for n in nodes if n.shape != "text"}:
            same = [n for n in nodes if n.shape == shape]
            w, h = max(n.w for n in same), max(n.h for n in same)
            for n in same:
                n.w, n.h = w, h

    # 1) steps: longest path from the sources, ignoring edges that close a cycle
    out = {n.id: [] for n in nodes}
    for e in edges:
        out[e.a].append(e)
    back, state = set(), {}

    def dfs(nid):
        state[nid] = 1
        for e in out[nid]:
            s = state.get(e.b)
            if s == 1:
                back.add(id(e))
            elif s is None:
                dfs(e.b)
        state[nid] = 2

    for n in nodes:
        if n.id not in state:
            dfs(n.id)
    preds = {n.id: [] for n in nodes}
    for e in edges:
        if id(e) not in back:
            preds[e.b].append(e.a)
    done = {}

    def rank(nid):
        if nid not in done:
            done[nid] = -1  # guard
            n = byid[nid]
            done[nid] = (n.step - 1) if n.step is not None else max((rank(p) + 1 for p in preds[nid]), default=0)
        return done[nid]

    for n in nodes:
        n.rank = max(0, rank(n.id))
    ranks = sorted({n.rank for n in nodes})
    by_rank = {r: [n for n in nodes if n.rank == r] for r in ranks}

    # 2) lanes: a node sits in the lane of its (median) predecessor; the first path stays straight
    for r in ranks:
        taken = []
        for n in by_rank[r]:  # pinned lanes first
            if n.lane is not None:
                n.pos = n.lane - 1
                taken.append(n.pos)
        items = []
        for n in by_rank[r]:
            if n.lane is None:
                ps = sorted(byid[p].pos for p in preds[n.id] if byid[p].rank < r)
                items.append((n, ps[(len(ps) - 1) // 2] if ps else None))
        items.sort(key=lambda t: (math.inf if t[1] is None else t[1], t[0].index))
        for n, want in items:  # wanted lane, else the next free one below it
            pos = 0.0 if want is None else want
            while True:
                hit = [t for t in taken if abs(t - pos) < 1 - EPS]
                if not hit:
                    break
                pos = max(hit) + 1
            n.pos = pos
            taken.append(pos)
    low = min(n.pos for n in nodes)
    for n in nodes:
        n.pos -= low

    # 3) spacing: gaps grow to fit edge labels and group boxes
    M = (lambda n: n.w) if horiz else (lambda n: n.h)   # extent along the flow
    C = (lambda n: n.h) if horiz else (lambda n: n.w)   # extent across it
    labels = {}
    for e in edges:
        if e.label:
            lw, lh = _block(visible(e.label), 16 * CHAR, CHAR)
            labels[id(e)] = (lw * LABEL, lh * LABEL)
    gap_main = {r: 2.4 * gap for r in ranks}
    gap_cross = 1.3 * gap
    for e in edges:
        if id(e) in labels and byid[e.b].rank > byid[e.a].rank:
            lw, lh = labels[id(e)]
            ra = byid[e.a].rank
            gap_main[ra] = max(gap_main[ra], (lw if horiz else lh) + 1.4)
    if labels:
        gap_cross = max(gap_cross, max((lh if horiz else lw) for lw, lh in labels.values()) + 0.8)
    glabel = LABEL * LINE + 0.2
    if groups:
        need = 2 * GPAD + glabel + 0.6
        gap_main = {r: max(g, need) for r, g in gap_main.items()}
        gap_cross = max(gap_cross, need)
    col = {r: max(M(n) for n in by_rank[r]) for r in ranks}
    start, u = {}, 0.0
    for r in ranks:
        start[r] = u
        u += col[r] + gap_main[r]
    pitch = max(C(n) for n in nodes) + gap_cross
    for n in nodes:
        n.u = start[n.rank] + col[n.rank] / 2
        n.v = n.pos * pitch

    def rect(n, pad=0.0):
        return (n.u - M(n) / 2 - pad, n.v - C(n) / 2 - pad, n.u + M(n) / 2 + pad, n.v + C(n) / 2 + pad)

    # group boxes (flow coordinates); the label strip goes on the physical top edge
    top_side = {"right": 1, "left": 1, "down": 0, "up": 2}[direction]  # 1: v-, 0: u-, 2: u+
    grects = []
    for g in groups:
        rs = [rect(byid[m], GPAD) for m in g.members]
        u0, v0 = min(r[0] for r in rs), min(r[1] for r in rs)
        u1, v1 = max(r[2] for r in rs), max(r[3] for r in rs)
        if top_side == 1:
            v0 -= glabel
        elif top_side == 0:
            u0 -= glabel
        else:
            u1 += glabel
        grects.append((u0, v0, u1, v1))
        g.rank = min(byid[m].rank for m in g.members)

    # 4) routing
    def port(n, side):
        idx, fx, fy = SHAPES[n.shape][3][PHYS[direction][side]]
        fu, fv = {"right": (fx, fy), "down": (fy, fx), "left": (1 - fx, fy), "up": (1 - fy, fx)}[direction]
        r = rect(n)
        return (r[0] + fu * (r[2] - r[0]), r[1] + fv * (r[3] - r[1])), DIRS[side], idx

    def clear(pts, skip):
        for p, q in zip(pts, pts[1:]):
            u0, u1 = sorted((p[0], q[0]))
            v0, v1 = sorted((p[1], q[1]))
            for n in nodes:
                if n.id in skip:
                    continue
                r = rect(n, MARGIN)
                if u0 < r[2] and u1 > r[0] and v0 < r[3] and v1 > r[1]:
                    return False
        return True

    used = {"pos": [], "neg": []}

    def outer(side, ua, ub):
        """Coordinate (across the flow) of a path routed around everything between ua and ub."""
        lo, hi = min(ua, ub), max(ua, ub)
        rs = [rect(n) for n in nodes] + grects
        rs = [r for r in rs if r[0] < hi + EPS and r[2] > lo - EPS]
        sgn = 1 if side == "pos" else -1
        lvl = (max(r[3] for r in rs) + CLEAR) if side == "pos" else (min(r[1] for r in rs) - CLEAR)
        moved = True
        while moved:
            moved = False
            for a, b, l in used[side]:
                if a < hi and b > lo and abs(l - lvl) < STAGGER - EPS:
                    lvl = l + sgn * STAGGER
                    moved = True
        return lvl

    inv = {v: k for k, v in PHYS[direction].items()}
    fwd_out = {n.id: [e for e in out[n.id] if byid[e.b].rank > n.rank and not (e.exit or e.enter)] for n in nodes}
    primary = {nid: min(es, key=lambda e: (abs(byid[e.b].pos - byid[nid].pos), edges.index(e)))
               for nid, es in fwd_out.items() if es}
    order = sorted(edges, key=lambda e: (byid[e.b].rank <= byid[e.a].rank, byid[e.a].rank, edges.index(e)))
    for e in order:
        A, B = byid[e.a], byid[e.b]
        ra, rb = A.rank, B.rank
        cands = []

        def add(pri, sa, sb, bend=None, level=None, around=None):
            (s, ds, ia), (t, dt, ib) = port(A, sa), port(B, sb)
            pts = _simplify(route(s, ds, t, (-dt[0], -dt[1]), bend, level))
            cost = sum(abs(p[0] - q[0]) + abs(p[1] - q[1]) for p, q in zip(pts, pts[1:])) + 1.5 * (len(pts) - 2)
            cands.append((pri, cost, pts, ia, ib, around, level))

        if e.exit or e.enter:
            sa = inv[e.exit] if e.exit else ("fwd" if rb > ra else "pos")
            sb = inv[e.enter] if e.enter else ("back" if rb > ra else "pos")
            add(0, sa, sb)
        elif rb > ra:
            (s, _, _), (t, _, _) = port(A, "fwd"), port(B, "back")
            if abs(s[1] - t[1]) < EPS:
                add(0, "fwd", "back")
            if primary.get(A.id) is not e:  # a branch leaves from the side, like a decision's "no"
                side = "pos" if B.v > A.v else "neg"
                s2, ds2, _ = port(A, side)
                if (t[1] - s2[1]) * ds2[1] > STUB:
                    add(1, side, "back")
            add(2, "fwd", "back", bend=start[rb] - gap_main[ranks[ranks.index(rb) - 1]] / 2)
            add(2.5, "fwd", "back", bend=start[ra] + col[ra] + gap_main[ra] / 2)
            for side in ("pos", "neg"):
                add(3, side, side, level=outer(side, A.u, B.u), around=side)
        elif rb == ra:
            lo, hi = sorted((A.pos, B.pos))
            side = "pos" if B.pos > A.pos else "neg"
            if not any(lo < n.pos < hi for n in by_rank[ra]):
                add(0, side, OPPOSITE[side])
            add(3, "fwd", "fwd", level=start[ra] + col[ra] + gap_main[ra] / 2)
            first = ranks.index(ra) == 0
            add(3.5, "back", "back", level=start[ra] - (2.4 * gap if first else gap_main[ranks[ranks.index(ra) - 1]]) / 2)
        else:  # back edge: around the outside, below/right first
            for side in ("pos", "neg"):
                add(3, side, side, level=outer(side, A.u, B.u), around=side)
                add(4, "fwd", "back", level=outer(side, B.u - M(B) / 2 - STUB, A.u + M(A) / 2 + STUB), around=side)
        ok = [c for c in cands if clear(c[2], {A.id, B.id})] or cands
        pri, _, pts, ia, ib, around, level = min(ok, key=lambda c: (c[0], c[1]))
        if around:
            us = [p[0] for p in pts]
            lvl = max(p[1] for p in pts) if around == "pos" else min(p[1] for p in pts)
            used[around].append((min(us), max(us), lvl))
        e.pts, e.sites, e.rank = pts, (ia, ib), max(ra, rb)

    # 5) flow coordinates -> physical (x right, y down), then labels and the bounding box
    def P(p):
        u_, v_ = p
        return {"right": (u_, v_), "down": (v_, u_), "left": (-u_, v_), "up": (v_, -u_)}[direction]

    def Prect(r):
        (x0, y0), (x1, y1) = P((r[0], r[1])), P((r[2], r[3]))
        return min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0)

    for n in nodes:
        cx, cy = P((n.u, n.v))
        n.x, n.y = cx - n.w / 2, cy - n.h / 2
    for g, r in zip(groups, grects):
        g.x, g.y, g.w, g.h = Prect(r)
    for e in edges:
        e.pts = [P(p) for p in e.pts]
    _place_labels(nodes, edges, groups, labels)
    return _finish(nodes, edges, groups)


def _spot(p, q, lw, lh, side, start):
    """Label box beside segment p-q: near p (start) or centred; side -1 = above / left, +1 = below / right."""
    if abs(p[1] - q[1]) < EPS:  # horizontal
        if start:
            x = p[0] + 0.3 if q[0] > p[0] else p[0] - 0.3 - lw
        else:
            x = (p[0] + q[0]) / 2 - lw / 2
        return x, (p[1] - 0.12 - lh if side < 0 else p[1] + 0.12), lw, lh
    if start:
        y = p[1] + 0.25 if q[1] > p[1] else p[1] - 0.25 - lh
    else:
        y = (p[1] + q[1]) / 2 - lh / 2
    return (p[0] - 0.2 - lw if side < 0 else p[0] + 0.2), y, lw, lh


def _place_labels(nodes, edges, groups, labels):
    """Put each edge label where it touches no shape, group border, line or other label. Preferred
    spot: beside the first segment, away from the next turn (centred on a straight edge)."""
    segs = [(p, q) for e in edges for p, q in zip(e.pts, e.pts[1:])]
    boxes = [(n.x, n.y, n.w, n.h) for n in nodes]
    frames = [(g.x, g.y, g.w, g.h) for g in groups]
    placed = []

    def overlap(a, b, pad=0.05):
        return a[0] < b[0] + b[2] + pad and b[0] < a[0] + a[2] + pad and a[1] < b[1] + b[3] + pad and b[1] < a[1] + a[3] + pad

    def hits(bx):
        n = sum(overlap(bx, b) for b in boxes) + sum(overlap(bx, b) for b in placed)
        for f in frames:  # a label may sit inside or outside a group box, not on its border
            inside = f[0] <= bx[0] and f[1] <= bx[1] and bx[0] + bx[2] <= f[0] + f[2] and bx[1] + bx[3] <= f[1] + f[3]
            n += overlap(bx, f, 0.0) and not inside
        for p, q in segs:
            seg = (min(p[0], q[0]), min(p[1], q[1]), abs(q[0] - p[0]), abs(q[1] - p[1]))
            n += overlap(bx, seg, 0.02)
        return n

    for e in edges:
        if id(e) not in labels:
            continue
        lw, lh = labels[id(e)]
        pts = e.pts
        p, q = pts[0], pts[1]
        horiz = abs(p[1] - q[1]) < EPS
        if len(pts) > 2:  # away from the next turn
            side = -math.copysign(1, (pts[2][1] - q[1]) if horiz else (pts[2][0] - q[0]))
        else:
            side = -1 if horiz else 1
        straight = len(pts) == 2
        spots = [_spot(p, q, lw, lh, side, not straight), _spot(p, q, lw, lh, -side, not straight)]
        for a, b in sorted(zip(pts, pts[1:]), key=lambda s: -(abs(s[1][0] - s[0][0]) + abs(s[1][1] - s[0][1]))):
            spots += [_spot(a, b, lw, lh, -1, False), _spot(a, b, lw, lh, 1, False)]
        scored = [(hits(s), i, s) for i, s in enumerate(spots)]
        e.label_box = min(scored)[2]
        placed.append(e.label_box)


def _finish(nodes, edges, groups) -> Layout:
    """Shift everything to start at (0, 0) and measure the diagram."""
    boxes = [(n.x, n.y, n.w, n.h) for n in nodes] + [(g.x, g.y, g.w, g.h) for g in groups]
    boxes += [e.label_box for e in edges if e.label_box]
    boxes += [(p[0] - .25, p[1] - .25, .5, .5) for e in edges for p in e.pts]  # arrowheads
    x0 = min(b[0] for b in boxes) - 0.1
    y0 = min(b[1] for b in boxes) - 0.1
    x1 = max(b[0] + b[2] for b in boxes) + 0.1
    y1 = max(b[1] + b[3] for b in boxes) + 0.1
    for n in nodes:
        n.x, n.y = n.x - x0, n.y - y0
    for g in groups:
        g.x, g.y = g.x - x0, g.y - y0
    for e in edges:
        e.pts = [(x - x0, y - y0) for x, y in e.pts]
        if e.label_box:
            bx, by, bw, bh = e.label_box
            e.label_box = (bx - x0, by - y0, bw, bh)
    return Layout(x1 - x0, y1 - y0, nodes, edges, groups)


# ------------------------------------------------------------------------------------------
# steps: a row of chevrons (or a linear flow of any other shape)
# ------------------------------------------------------------------------------------------
def _steps(spec: dict, visible) -> Layout:
    items = spec.get("items")
    if not isinstance(items, list) or not items:
        raise ValueError("steps needs a list of step texts, e.g. steps: [Collect, Clean, Train]")
    hl = spec.get("highlight")
    hl = set() if hl is None else ({int(hl)} if isinstance(hl, (int, float)) else {int(x) for x in hl})
    style = _style(spec.get("style", "tinted"))
    color = spec.get("color", "structure")
    shape = str(spec.get("shape", "chevron")).lower()
    nodes = []
    for i, it in enumerate(items):
        it = it if isinstance(it, dict) else {"text": it}
        bad = set(it) - STEP_KEYS
        if bad:
            raise ValueError(f"unknown steps item key(s) {', '.join(map(str, bad))}; use {', '.join(sorted(STEP_KEYS))}")
        st = "solid" if (i + 1) in hl else _style(it.get("style", style))
        nodes.append({"id": f"step{i + 1}", "text": "" if it.get("text") is None else str(it.get("text")),
                      "color": it.get("color", color), "style": st, "text_color": it.get("text_color"),
                      "bold": bool(it.get("bold", (i + 1) in hl)), "shape": it.get("shape", shape)})
    if shape != "chevron":  # any flow shape: a linear flow with arrows
        flow = {k: v for k, v in spec.items() if k in ("direction", "wrap", "gap", "uniform")}
        return _flow({**flow, "nodes": nodes, "shape": shape}, visible)
    wrap = float(spec.get("wrap", 14)) * CHAR
    out = []
    for i, nd in enumerate(nodes):
        n = Node(nd["id"], nd["text"], "pentagon" if i == 0 else "chevron", nd["color"], nd["style"],
                 nd["text_color"], nd["bold"], index=i, rank=i)
        cw = CHAR_BOLD if n.bold else CHAR
        n.w, n.h = _block(visible(n.text), wrap, cw)  # text block for now
        out.append(n)
    px, py = SHAPES["chevron"][2]
    h = max(n.h for n in out) + 2 * py
    tip = 0.35 * h
    w = max(max(n.w for n in out) + 2 * px + 2 * tip, 1.2 * h)
    step = w - tip + 0.3
    for i, n in enumerate(out):
        n.x, n.y, n.w, n.h, n.pad, n.adj = i * step, 0.0, w, h, (px, py), 0.35
    return Layout(step * (len(out) - 1) + w, h, out, [], [])


# ------------------------------------------------------------------------------------------
# routing
# ------------------------------------------------------------------------------------------
def _dot(p, d):
    return p[0] * d[0] + p[1] * d[1]


def _mv(p, d, k):
    return (p[0] + k * d[0], p[1] + k * d[1])


def route(s, ds, e, de, bend=None, level=None, stub=STUB):
    """Orthogonal path from s (leaving in direction ds) to e (arriving while moving in direction de).
    bend: coordinate on the ds axis of the middle segment of a Z. level: coordinate on the ds axis
    of the far segment of a U, or across ds for an S that has to double back."""
    sgn = ds[0] + ds[1]
    q = (abs(ds[1]), abs(ds[0]))  # the other axis
    rel = (e[0] - s[0], e[1] - s[1])
    along = _dot(rel, ds)
    if ds == de:
        if abs(_dot(rel, q)) < EPS and along > EPS:
            return [s, e]  # straight
        if along > 0.2:  # Z
            t = along / 2 if bend is None else bend * sgn - _dot(s, ds)
            t = min(max(t, 0.1), along - 0.1)
            return [s, _mv(s, ds, t), _mv(e, ds, t - along), e]
        lv = (_dot(s, q) + _dot(e, q)) / 2 if level is None else level  # S: double back around
        a, m = _mv(s, ds, stub), _mv(e, ds, -stub)
        return [s, a, _mv(a, q, lv - _dot(a, q)), _mv(m, q, lv - _dot(m, q)), m, e]
    if ds == (-de[0], -de[1]):  # U
        far = max(_dot(s, ds), _dot(e, ds)) + stub if level is None else level * sgn
        return [s, _mv(s, ds, far - _dot(s, ds)), _mv(e, ds, far - _dot(e, ds)), e]
    corner = _mv(s, ds, along)  # perpendicular: L if the corner lies ahead of both ends
    if along > stub / 2 and _dot((e[0] - corner[0], e[1] - corner[1]), de) > stub / 2:
        return [s, corner, e]
    a, m = _mv(s, ds, stub), _mv(e, de, -stub)
    b = _mv(a, de, _dot((m[0] - a[0], m[1] - a[1]), de))
    return [s, a, b, m, e]


def _simplify(pts):
    """Drop repeated points and merge straight runs."""
    out = []
    for p in pts:
        if out and abs(p[0] - out[-1][0]) < EPS and abs(p[1] - out[-1][1]) < EPS:
            continue
        out.append(p)
    i = 1
    while i < len(out) - 1:
        a, b, c = out[i - 1], out[i], out[i + 1]
        if (abs(a[0] - b[0]) < EPS and abs(b[0] - c[0]) < EPS) or (abs(a[1] - b[1]) < EPS and abs(b[1] - c[1]) < EPS):
            del out[i]
        else:
            i += 1
    return out


# ------------------------------------------------------------------------------------------
# polyline -> PowerPoint connector
# ------------------------------------------------------------------------------------------
CONNECTORS = {1: "straightConnector1", 2: "bentConnector2", 3: "bentConnector3", 4: "bentConnector4",
              5: "bentConnector5"}
MIN_EXT = 12700  # 1 pt: a bent connector needs some width to express a U that returns to its own line


def connector_geometry(pts):
    """Orthogonal polyline in integer EMU -> (preset, (x, y), (cx, cy), rot_deg, flipH, flipV, [adj...]).

    The preset paths, in the unrotated, unflipped box (w, h), all start at (0, 0) and end at (w, h):
      straightConnector1  (0,0) (w,h)
      bentConnector2      (0,0) (w,0) (w,h)
      bentConnector3      (0,0) (x1,0) (x1,h) (w,h)                     x1 = adj1 * w
      bentConnector4      (0,0) (x1,0) (x1,y2) (w,y2) (w,h)             y2 = adj2 * h
      bentConnector5      (0,0) (x1,0) (x1,y2) (x3,y2) (x3,h) (w,h)     x3 = adj3 * w
    Find the rotation and flips that map one of them onto the path.
    """
    n = len(pts) - 1
    if n not in CONNECTORS:
        raise ValueError(f"connector path with {n} segments")
    S, E = pts[0], pts[-1]
    tol = 3
    for rot in (0, 90, 270, 180):
        c, s = round(math.cos(math.radians(rot))), round(math.sin(math.radians(rot)))
        for fh in (0, 1):
            for fv in (0, 1):
                loc = []
                for P in pts:
                    dx, dy = P[0] - S[0], P[1] - S[1]
                    x, y = c * dx + s * dy, -s * dx + c * dy  # rotate back
                    loc.append((-x if fh else x, -y if fv else y))
                W, H = loc[-1]
                if W < -tol or H < -tol:
                    continue
                W, H = max(W, 0), max(H, 0)
                ok, adj = _match(n, loc, W, H, tol)
                if not ok:
                    continue
                if n > 1 and n != 2:  # the adjust values divide by the box size
                    if adj and W < MIN_EXT and any(k in (1, 3) for k, _ in adj):
                        W = MIN_EXT
                    if adj and H < MIN_EXT and any(k == 2 for k, _ in adj):
                        H = MIN_EXT
                vals = []
                for k, v in adj:
                    size = W if k in (1, 3) else H
                    vals.append(int(round(v / size * 100000)) if size else 50000)
                cx, cy = (S[0] + E[0]) / 2, (S[1] + E[1]) / 2
                return (CONNECTORS[n], (int(round(cx - W / 2)), int(round(cy - H / 2))), (int(W), int(H)),
                        rot, fh, fv, vals)
    raise ValueError("no connector matches this path")


def _match(n, loc, W, H, tol):
    """Does the local path follow preset n? Returns (ok, [(adj number, local coordinate)])."""
    near = lambda a, b: abs(a - b) <= tol
    if n == 1:
        return True, []
    p = loc
    if n == 2:
        return near(p[1][1], 0) and near(p[1][0], W), []
    if n == 3:
        ok = near(p[1][1], 0) and near(p[2][0], p[1][0]) and near(p[2][1], H)
        return ok, [(1, p[1][0])]
    if n == 4:
        ok = (near(p[1][1], 0) and near(p[2][0], p[1][0]) and near(p[3][1], p[2][1]) and near(p[3][0], W))
        return ok, [(1, p[1][0]), (2, p[2][1])]
    ok = (near(p[1][1], 0) and near(p[2][0], p[1][0]) and near(p[3][1], p[2][1]) and near(p[4][0], p[3][0])
          and near(p[4][1], H))
    return ok, [(1, p[1][0]), (2, p[2][1]), (3, p[3][0])]
