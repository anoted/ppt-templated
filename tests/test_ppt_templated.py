"""Tests for ppt-templated.   python -m unittest discover tests   (or: python -m pytest tests)

They build real decks into a temporary folder and read them back with python-pptx. No PowerPoint or
LibreOffice is needed (previews are not rendered here)."""

from __future__ import annotations

import contextlib
import io
import json
import os
import sys
import tempfile
import textwrap
import unittest

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)

import yaml  # noqa: E402
from pptx import Presentation  # noqa: E402

import build_template  # noqa: E402
import datafiles  # noqa: E402
import flows  # noqa: E402
import ppt_templated as pt  # noqa: E402

try:
    import plots  # noqa: E402
except ImportError:  # matplotlib missing
    plots = None


def build(content: dict, files: dict | None = None, **kw):
    """Build a deck from a content dict in a temp folder -> (builder, Presentation, folder)."""
    tmp = tempfile.mkdtemp(prefix="ppt_test_")
    for name, text in (files or {}).items():
        with open(os.path.join(tmp, name), "w", encoding="utf-8") as fh:
            fh.write(textwrap.dedent(text))
    path = os.path.join(tmp, "deck.yaml")
    with open(path, "w", encoding="utf-8") as fh:
        yaml.safe_dump(content, fh, allow_unicode=True)
    with contextlib.redirect_stderr(io.StringIO()):
        builder, prs, out = pt._build(path, **kw)
    if out:
        prs = Presentation(out)
    return builder, prs, tmp


def titles(prs):
    return [s.shapes.title.text_frame.text if s.shapes.title is not None else "" for s in prs.slides]


class Templates(unittest.TestCase):
    def test_every_bundled_template_builds_the_feature_example(self):
        for name in pt.template_names():
            with self.subTest(template=name), contextlib.redirect_stderr(io.StringIO()):
                builder, prs, _ = pt._build(os.path.join(HERE, "examples", "example.yaml"), template=name, check=True)
                self.assertGreater(len(prs.slides), 20)
                self.assertEqual([w for w in builder.warnings if "measured" not in w], [], name)

    def test_names_resolve_and_mode_is_detected(self):
        self.assertTrue(pt.find_template("midnight").endswith("beamer_midnight.pptx"))
        self.assertTrue(pt.find_template("beamer_paper").endswith("beamer_paper.pptx"))
        self.assertEqual(pt.find_template(None), pt.DEFAULT_TEMPLATE)
        with self.assertRaises(SystemExit) as cm:
            pt.find_template("midnite")
        self.assertIn("midnight", str(cm.exception))  # the suggestion
        modes = {n: m for n, m, _ in pt.describe_templates()}
        self.assertEqual(modes["midnight"], "dark")
        self.assertEqual(modes["chalkboard"], "dark")
        self.assertEqual(modes["oxford"], "light")
        self.assertEqual(modes["parchment"], "light")

    def test_dark_template_maps_text_and_background(self):
        b = pt.DeckBuilder({"meta": {"template": "midnight"}}, None, HERE)
        self.assertTrue(b.dark)
        colors = b._theme()[0]
        self.assertEqual(colors["bg1"], "0F172A")
        self.assertEqual(colors["tx1"], "E2E8F0")
        self.assertFalse(pt.DeckBuilder({"meta": {}}, None, HERE).dark)


class Colours(unittest.TestCase):
    def test_tints_move_towards_the_background(self):
        class CF:  # stands in for a python-pptx ColorFormat
            theme_color = rgb = brightness = None
        light, dark = CF(), CF()
        pt.paint(light, ["accent1", 0.6])
        pt.paint(dark, ["accent1", 0.6], dark=True)
        self.assertAlmostEqual(light.brightness, 0.6)
        self.assertAlmostEqual(dark.brightness, -0.6)
        muted = CF()
        pt.paint(muted, "muted", dark=True)
        self.assertLess(muted.brightness, 0)

    def test_bad_colours_are_reported(self):
        with self.assertRaises(ValueError):
            pt.resolve_color("accent9")
        with self.assertRaises(ValueError):
            pt.resolve_color(["accent1", "lots"])
        self.assertEqual(pt.resolve_color("#1f3a5f")[:2], ("rgb", "1F3A5F"))


class Items(unittest.TestCase):
    def test_new_items_and_statement(self):
        builder, prs, _ = build({"meta": {"title": "T"}, "slides": [
            {"layout": "statement", "title": "Kicker", "text": "One big message", "subtitle": "below"},
            {"title": "Q", "content": [{"quote": {"text": "Simple is good.", "by": "Someone"}}]},
            {"title": "TL", "content": [{"timeline": [{"2019": "Idea"}, "2021: Prototype",
                                                      {"date": 2023, "text": "Launch"}], "highlight": 2}]},
            {"title": "C", "content": [{"cards": [{"title": "A", "text": "x"}, {"title": "B", "bullets": ["y"]},
                                                  "C"], "build": True}]},
        ]})
        self.assertEqual(builder.warnings, [])
        names = [sh.name for s in prs.slides for sh in s.shapes]
        self.assertIn("Quote", names)
        self.assertIn("Timeline", names)
        texts = " ".join(sh.text_frame.text for sh in prs.slides[0].shapes if sh.has_text_frame)
        self.assertIn("One big message", texts)
        self.assertIn("Kicker", texts)
        info = {i["slide"]: i for i in builder.slide_info}
        self.assertEqual(info[1]["layout"], "statement")
        self.assertEqual(info[4]["animation_steps"], 3)  # one card per click

    def test_unknown_keys_get_suggestions(self):
        builder, _, _ = build({"slides": [{"title": "x", "content": [{"cards": ["a"], "colour": "alert"}]}]},
                              check=True)
        self.assertTrue(any("did you mean 'color'" in w for w in builder.warnings), builder.warnings)


class Overflow(unittest.TestCase):
    def test_long_numbered_list_continues_with_numbering(self):
        items = [f"point {i}" for i in range(1, 21)]
        builder, prs, _ = build({"meta": {"overflow": "split"},
                                 "slides": [{"title": "Long", "content": [{"numbered": items}]}]})
        t = titles(prs)
        self.assertGreater(len(t), 1)
        self.assertTrue(all(x.startswith("Long (") for x in t), t)
        xml = prs.slides[1].shapes._spTree.xml
        self.assertIn('startAt="', xml)  # numbering goes on from the previous slide
        self.assertEqual(builder.warnings, [])

    def test_long_table_repeats_header(self):
        rows = [[f"r{i}", i, i * 2] for i in range(40)]
        _, prs, _ = build({"slides": [{"title": "T", "overflow": "split",
                                       "content": [{"table": {"header": ["a", "b", "c"], "rows": rows}}]}]})
        self.assertGreater(len(prs.slides), 1)
        for s in prs.slides:
            tbl = next(sh for sh in s.shapes if sh.has_table).table
            self.assertEqual(tbl.cell(0, 0).text, "a")

    def test_short_content_is_not_split(self):
        _, prs, _ = build({"meta": {"overflow": "split"}, "slides": [{"title": "S", "content": [["a", "b"]]}]})
        self.assertEqual(titles(prs), ["S"])


class Data(unittest.TestCase):
    CSV = """\
        Model,Params,Accuracy
        Base,12M,"81.4"
        Ours,20M,86.2
        """

    def test_table_chart_from_csv(self):
        builder, prs, _ = build({"slides": [
            {"title": "T", "content": [{"table": {"data": "r.csv", "select": ["Model", "Accuracy"]}}]},
            {"title": "C", "content": [{"chart": {"type": "column", "data": "r.csv"}}]},
        ]}, files={"r.csv": self.CSV})
        tbl = next(sh for sh in prs.slides[0].shapes if sh.has_table).table
        self.assertEqual([tbl.cell(0, j).text for j in range(2)], ["Model", "Accuracy"])
        self.assertEqual(tbl.cell(2, 1).text, "86.2")
        chart = next(sh for sh in prs.slides[1].shapes if sh.has_chart).chart
        self.assertEqual([s.name for s in chart.plots[0].series], ["Accuracy"])  # the numeric column

    def test_errors_name_the_column(self):
        with self.assertRaises(SystemExit) as cm:
            build({"slides": [{"title": "T", "content": [{"table": {"data": "r.csv", "select": ["Acuracy"]}}]}]},
                  files={"r.csv": self.CSV}, check=True)
        self.assertIn("did you mean 'Accuracy'", str(cm.exception))

    def test_number_parsing(self):
        self.assertEqual(datafiles.number("1,234.5"), 1234.5)
        self.assertEqual(datafiles.number("12%"), 12.0)
        self.assertIsNone(datafiles.number("n/a"))
        self.assertIsNone(datafiles.number("nan"))


@unittest.skipIf(plots is None, "matplotlib is not installed")
class Plots(unittest.TestCase):
    def test_formula_evaluator(self):
        import numpy as np
        x = np.array([-1.0, 0.0, 2.0])
        self.assertEqual(list(plots.formula("x^2 + 1", {"x": x})), [2.0, 1.0, 5.0])
        self.assertEqual(list(plots.formula("where(x > 0, x, 0)", {"x": x})), [0.0, 0.0, 2.0])
        self.assertEqual(list(plots.formula("x if x > 0 else -x", {"x": x})), [1.0, 0.0, 2.0])
        for bad in ("__import__('os')", "x.real", "open('f')", "[1, 2]", "lambda: 1", "(1).__class__"):
            with self.assertRaises(ValueError, msg=bad):
                plots.formula(bad, {"x": x})

    def test_new_plot_types_build(self):
        builder, prs, _ = build({"meta": {"template": "notebook"}, "slides": [
            {"title": "F", "content": [{"plot": {"type": "function", "functions": {"a": "sin(x)", "b": "x^2/10"},
                                                 "x": [-3, 3]}}]},
            {"title": "H", "content": [{"plot": {"type": "hist", "values": [1, 2, 2, 3, 3, 3]}}]},
            {"title": "M", "content": [{"plot": {"type": "heatmap", "matrix": [[1, -1], [0.5, 0]], "center": 0}}]},
        ]})
        self.assertEqual(builder.warnings, [])
        self.assertEqual(len(prs.slides), 3)

    def test_python_plots_need_permission(self):
        spec = {"slides": [{"title": "P", "content": [{"plot": {"type": "python", "code": "ax.plot([1, 2])"}}]}]}
        with self.assertRaises(SystemExit) as cm:
            build(spec, check=True)
        self.assertIn("--allow-python", str(cm.exception))
        builder, prs, _ = build(spec, allow_python=True)
        self.assertEqual(builder.warnings, [])


class Logo(unittest.TestCase):
    def test_logo_narrows_the_title_but_keeps_its_place(self):
        svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 100"><rect width="300" height="100"/></svg>'
        _, prs, _ = build({"meta": {"logo": "logo.svg"}, "slides": [
            {"title": "A frame title", "content": ["text"]}]}, files={"logo.svg": svg})
        slide = prs.slides[0]
        logo = next(sh for sh in slide.shapes if sh.name == "Logo")
        title = slide.shapes.title
        layout_title = slide.slide_layout.placeholders[0]
        self.assertEqual(title.top, layout_title.top)
        self.assertLessEqual(title.left + title.width, logo.left)


class CommandLine(unittest.TestCase):
    def run_main(self, *argv):
        out, err = io.StringIO(), io.StringIO()
        code = 0
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                pt.main(list(argv))
            except SystemExit as exc:
                code = exc.code
        return code, out.getvalue(), err.getvalue()

    def test_json_report(self):
        code, out, _ = self.run_main(os.path.join(HERE, "examples", "example.yaml"), "--check", "--json")
        self.assertEqual(code, 0)
        rep = json.loads(out)
        self.assertTrue(rep["ok"])
        self.assertEqual(rep["slides"], len(rep["slide_info"]))
        self.assertIn("fill", rep["slide_info"][3])

    def test_json_error(self):
        tmp = tempfile.mkdtemp()
        path = os.path.join(tmp, "x.yaml")
        with open(path, "w") as fh:
            fh.write("meta: {template: nope}\nslides: []\n")
        code, out, _ = self.run_main(path, "--json")
        self.assertEqual(code, 1)
        self.assertFalse(json.loads(out)["ok"])

    def test_new_starter_builds_cleanly(self):
        tmp = tempfile.mkdtemp()
        path = os.path.join(tmp, "talk.yaml")
        self.assertEqual(self.run_main("--new", path, "-t", "parchment")[0], 0)
        code, _, err = self.run_main(path, "--check", "--strict")
        self.assertEqual(code, 0, err)
        self.assertNotEqual(self.run_main("--new", path)[0], 0)  # never overwrites

    def test_templates_list(self):
        code, out, _ = self.run_main("--templates", "--json")
        self.assertEqual(code, 0)
        self.assertGreaterEqual(len(json.loads(out)), 14)


def _have_libreoffice():
    import importlib.util
    import shutil
    if importlib.util.find_spec("pymupdf") is None and importlib.util.find_spec("fitz") is None:
        return False
    return bool(shutil.which("soffice") or shutil.which("libreoffice"))


@unittest.skipUnless(_have_libreoffice(), "needs LibreOffice and PyMuPDF")
class Measuring(unittest.TestCase):
    def test_overflow_is_measured_in_the_rendered_slides(self):
        long = " ".join(["WWWWWWWW"] * 16)
        content = {"meta": {"template": "blue"}, "slides": [
            {"title": "Overflow", "content": [{"columns": [[{"block": {"title": "A", "text": long, "size": 26}}],
                                                           [{"block": {"title": "B", "text": "short"}}]]}]},
            {"title": "Fine", "content": [["one", "two"]]}]}
        tmp = tempfile.mkdtemp(prefix="ppt_test_")
        path = os.path.join(tmp, "deck.yaml")
        with open(path, "w", encoding="utf-8") as fh:
            yaml.safe_dump(content, fh)
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            pt.main([path, "--preview", "--json"])
        rep = json.loads(out.getvalue())
        measured = [w for w in rep["warnings"] if "(measured)" in w["message"]]
        self.assertTrue(any(w["slide"] == 1 and "Block: A" in w["message"] for w in measured), rep["warnings"])
        self.assertFalse(any(w["slide"] == 2 for w in measured), rep["warnings"])


class ThemeFiles(unittest.TestCase):
    def test_validation(self):
        bad = [
            ({"dark": True}, "needs its own text and background"),
            ({"background": "marble"}, "not a theme colour, a texture"),
            ({"background": {"gradient": ["bg1"]}}, "two or more"),
            ({"colors": {"accent1": "12345"}}, "not a hex colour"),
            ({"frame": "box"}, "frame must be one of"),
        ]
        for theme, msg in bad:
            with self.subTest(theme=theme), self.assertRaises(SystemExit) as cm:
                build_template.check_theme(theme, "t")
            self.assertIn(msg, str(cm.exception))

    def test_dark_textured_theme_builds(self):
        out = os.path.join(tempfile.mkdtemp(), "t.pptx")
        theme = {"dark": True, "colors": {"text": "EEEEEE", "background": "202020"},
                 "background": {"texture": "chalk", "strength": 0.5}}
        build_template.make("t", theme, out, preview=False)
        prs = Presentation(out)
        cmap = prs.slide_master._element.find(pt.qn("p:clrMap"))
        self.assertEqual(cmap.get("bg1"), "dk1")
        self.assertIn("blipFill", prs.slide_master._element.xml)
        self.assertEqual(prs.core_properties.category, "dark")


class Flows(unittest.TestCase):
    def test_wider_fonts_get_wider_boxes(self):
        spec = {"edges": ["Measure -> Learn"]}
        narrow = flows.layout("flow", dict(spec), lambda t: t)
        wide = flows.layout("flow", dict(spec), lambda t: t, char_scale=1.15)
        self.assertGreater(wide.nodes[0].w, narrow.nodes[0].w)


if __name__ == "__main__":
    unittest.main()
