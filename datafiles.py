"""
datafiles.py - `data:` for tables, charts and plots: numbers come from a CSV/TSV/JSON file instead of the YAML.

    - table: {data: results.csv}                          # header row + rows, as written in the file
    - table: {data: results.csv, select: [Model, Acc]}    # only these columns, in this order
    - chart: {type: line, data: sales.csv}                # first column = categories, the others = series
    - plot:  {type: line, data: log.csv, x: epoch, y: [train, val]}

apply() fills in the keys the item would otherwise need (header/rows, categories/series, x/series, groups,
matrix/labels) and leaves keys written in the YAML alone, so the YAML can still override a title or a
label. File paths are relative to the YAML. All of this is documented in YAML_GUIDE.md (Data files).
"""

from __future__ import annotations

import csv
import difflib
import json
import os
import re

EXTS = (".csv", ".tsv", ".txt", ".json")


def load(path: str) -> tuple[list[str], list[list[str]]]:
    """(header, rows) of a CSV/TSV (delimiter detected) or JSON file: a list of records, a list of rows
    whose first row is the header, or {header: [...], rows: [[...]]}. Cells are strings ('' = empty)."""
    name = os.path.basename(path)
    if not os.path.exists(path):
        raise ValueError(f"data file not found: {path} (paths are relative to the YAML file)")
    ext = os.path.splitext(path)[1].lower()
    if ext not in EXTS:
        raise ValueError(f"{name}: data files must be {', '.join(EXTS)}")
    with open(path, encoding="utf-8-sig", newline="") as fh:
        text = fh.read()
    if ext == ".json":
        try:
            obj = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(f"{name}: not valid JSON ({exc})") from None
        if isinstance(obj, dict) and "rows" in obj:
            header, rows = list(obj.get("header") or []), list(obj["rows"])
        elif isinstance(obj, list) and obj and all(isinstance(r, dict) for r in obj):
            header = []
            for r in obj:
                header += [k for k in r if k not in header]
            rows = [[r.get(k) for k in header] for r in obj]
        elif isinstance(obj, list) and obj and all(isinstance(r, list) for r in obj):
            header, rows = obj[0], obj[1:]
        else:
            raise ValueError(f"{name}: JSON must be a list of records, a list of rows, or {{header, rows}}")
    else:
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|") if ext != ".tsv" else "excel-tab"
        except csv.Error:
            dialect = "excel"
        table = [r for r in csv.reader(text.splitlines(), dialect) if any(c.strip() for c in r)]
        if not table:
            raise ValueError(f"{name}: the file is empty")
        header, rows = table[0], table[1:]
    cell = lambda v: "" if v is None else str(v).strip()  # noqa: E731
    header = [cell(h) for h in header]
    rows = [[cell(v) for v in r] for r in rows if isinstance(r, (list, tuple))]
    width = max([len(header)] + [len(r) for r in rows])
    header += [f"column {i + 1}" for i in range(len(header), width)]
    rows = [r + [""] * (width - len(r)) for r in rows]
    if not rows:
        raise ValueError(f"{name}: no data rows below the header")
    return header, rows


def number(v):
    """'1,234.5' -> 1234.5, '12%' -> 12.0, '-3' -> -3.0; '' and text -> None."""
    s = str(v).strip().replace("\u2212", "-").replace("\u00a0", "")
    if s.endswith("%"):
        s = s[:-1].strip()
    if re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+(\.\d+)?", s):  # thousands separators
        s = s.replace(",", "")
    if not s or s.lower().lstrip("+-") in ("nan", "inf", "infinity"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _col(header, name, file, what):
    if isinstance(name, int) and not isinstance(name, bool):
        if not 0 <= name < len(header):
            raise ValueError(f"{file}: {what} column number {name} does not exist (0 to {len(header) - 1})")
        return name
    name = str(name)
    if name in header:
        return header.index(name)
    low = [h.lower() for h in header]
    if name.lower() in low:
        return low.index(name.lower())
    hit = difflib.get_close_matches(name, header, n=1)
    raise ValueError(f"{file}: no column '{name}'" + (f" (did you mean '{hit[0]}'?)" if hit else "")
                     + f"; columns: {', '.join(header)}")


def _numbers(header, rows, j, file, keep_empty=False):
    out = []
    for i, r in enumerate(rows):
        v = number(r[j])
        if v is None and r[j] != "":
            raise ValueError(f"{file}: row {i + 2}, column '{header[j]}': '{r[j]}' is not a number")
        if v is not None or keep_empty:
            out.append(v)
    return out


def _is_numeric(rows, j) -> bool:
    vals = [r[j] for r in rows if r[j] != ""]
    return bool(vals) and all(number(v) is not None for v in vals)


def apply(kind: str, spec: dict, base_dir: str) -> dict:
    """The item settings with the data file's contents filled in (keys given in the YAML win)."""
    path = str(spec["data"])
    path = path if os.path.isabs(path) else os.path.join(base_dir, path)
    header, rows = load(path)
    file = os.path.basename(path)
    out = dict(spec)
    out.pop("data")

    if kind == "table":
        cols = [_col(header, c, file, "select") for c in (out.pop("select", None) or range(len(header)))]
        out.setdefault("header", [header[j] for j in cols])
        out.setdefault("rows", [[r[j] for j in cols] for r in rows])
        return out

    ptype = str(out.get("type", "column" if kind == "chart" else "")).lower()
    if kind == "plot" and ptype == "function":
        raise ValueError("plot type 'function' draws formulas; it does not read data files")
    if kind == "plot" and ptype == "python":  # the code gets the whole file
        out["data"] = {"header": header, "rows": rows,
                       "columns": {h: [r[j] for r in rows] for j, h in enumerate(header)}}
        return out
    xname = out.get("x") if isinstance(out.get("x"), (str, int)) and not isinstance(out.get("x"), bool) else None
    xj = _col(header, xname, file, "x") if xname is not None else 0
    ynames = out.get("y")
    if isinstance(ynames, (str, int)) and not isinstance(ynames, bool):
        ynames = [ynames]
    if isinstance(ynames, list) and all(isinstance(v, (str, int)) and not isinstance(v, bool) for v in ynames):
        yj = [_col(header, n, file, "y") for n in ynames]
    else:
        ynames = None
        yj = [j for j in range(len(header)) if j != xj and _is_numeric(rows, j)]
        if kind == "plot" and ptype in ("box", "hist") and _is_numeric(rows, xj):
            yj = [j for j in range(len(header)) if _is_numeric(rows, j)]  # every numeric column is a sample
    if not yj:
        hint = ""
        for j in (j for j in range(len(header)) if j != xj):  # point at the cell that spoils a column
            bad = next(((i, r[j]) for i, r in enumerate(rows) if r[j] != "" and number(r[j]) is None), None)
            if bad and sum(number(r[j]) is not None for r in rows) >= len(rows) / 2:
                hint = f"; column '{header[j]}' has text in row {bad[0] + 2}: '{bad[1]}'"
                break
        raise ValueError(f"{file}: no numeric columns to draw (columns: {', '.join(header)}){hint}")
    for key in ("x", "y"):  # column names are used up here; drawing must not see them as data
        if (key == "x" and xname is not None) or (key == "y" and ynames is not None):
            out.pop(key, None)
    labels = [r[xj] for r in rows]
    series = {header[j]: _numbers(header, rows, j, file, keep_empty=True) for j in yj}

    if kind == "chart":
        if ptype.startswith("scatter"):
            xs = _numbers(header, rows, xj, file, keep_empty=True)
            out.setdefault("series", {n: [[x, y] for x, y in zip(xs, v) if x is not None and y is not None]
                                      for n, v in series.items()})
        else:
            out.setdefault("categories", labels)
            out.setdefault("series", series)
            if ptype not in ("pie", "doughnut", "radar"):
                out.setdefault("x_title", header[xj])
        return out

    # plots
    if ptype == "bar":
        out.setdefault("categories", labels)
        out.setdefault("series", {n: [0 if v is None else v for v in vals] for n, vals in series.items()})
    elif ptype == "line":
        xs = _numbers(header, rows, xj, file, keep_empty=True) if _is_numeric(rows, xj) else labels
        out.setdefault("x", xs)
        out.setdefault("series", series)
        out.setdefault("xlabel", header[xj])
    elif ptype == "scatter":
        xs = _numbers(header, rows, xj, file, keep_empty=True)
        out.setdefault("series", {n: [[x, y] for x, y in zip(xs, v) if x is not None and y is not None]
                                  for n, v in series.items()})
        out.setdefault("xlabel", header[xj])
        if len(series) == 1:
            out.setdefault("ylabel", next(iter(series)))
    elif ptype == "box":
        out.setdefault("groups", {n: [v for v in vals if v is not None] for n, vals in series.items()})
    elif ptype == "hist":
        out.setdefault("series", {n: [v for v in vals if v is not None] for n, vals in series.items()})
    elif ptype in ("heatmap", "confusion"):
        out.setdefault("matrix", [[0 if s[i] is None else s[i] for s in series.values()] for i in range(len(rows))])
        if ptype == "confusion":
            out.setdefault("labels", list(series))
        else:
            out.setdefault("xlabels", list(series))
            out.setdefault("ylabels", labels)
    else:
        raise ValueError(f"{file}: plot type {ptype!r} cannot read a data file")
    return out
