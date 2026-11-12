"""Tests for what ppt-templated decides by itself: arranging a figure beside its text, continuing long
content on more slides, fonts / accent / text size from meta, section navigation, the warning codes of
the JSON report and `--explain`. Also: the docs name every key, item and template the code knows.

    python -m unittest discover tests
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import sys
import unittest

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "tests"))

import build_template  # noqa: E402
import fonts  # noqa: E402
import metrics  # noqa: E402
import ppt_templated as pt  # noqa: E402
from test_ppt_templated import build, titles  # noqa: E402

LONG = ("A sentence that is long enough to take two lines on a slide, so that a handful of them "
        "cannot share one slide with anything else.")
FIGURE = {"placeholder": {"label": "Diagram", "aspect": "4:3"}}
POINTS = {"bullets": ["First point", "Second point", "Third point"]}


def info(builder):
    return [i for i in builder.slide_info if i["layout"] == "frame"]


class Arrange(unittest.TestCase):
    def test_figure_and_text_go_side_by_side(self):
        b, prs, _ = build({"slides": [{"title": "T", "content": [FIGURE, POINTS]}]}, check=True)
        self.assertEqual(info(b)[0]["arranged"], "columns")
        self.assertEqual(b.warnings, [])

    def test_the_figure_stays_on_the_side_it_was_written(self):
        for content, fig_first in (([FIGURE, POINTS], True), ([POINTS, FIGURE], False)):
            b, prs, _ = build({"slides": [{"title": "T", "content": content}]})
            shapes = {s.name: s.left for s in prs.slides[0].shapes}
            fig = next(v for k, v in shapes.items() if "Diagram" in k or "Placeholder" in k)
            text = next(s.left for s in prs.slides[0].shapes
                        if s.has_text_frame and "First point" in s.text_frame.text)
            self.assertEqual(fig < text, fig_first)

    def test_stack_and_wide_figures_are_left_alone(self):
        b, _, _ = build({"slides": [{"title": "T", "arrange": "stack", "content": [FIGURE, POINTS]}]}, check=True)
        self.assertIsNone(info(b)[0]["arranged"])
        wide = {"placeholder": {"label": "Pipeline", "aspect": "5:1"}}
        b, _, _ = build({"slides": [{"title": "T", "content": [wide, "One remark."]}]}, check=True)
        self.assertIsNone(info(b)[0]["arranged"])

    def test_explicit_sizes_and_columns_are_respected(self):
        sized = {"placeholder": {"label": "Diagram", "aspect": "4:3", "width": 0.4}}
        b, _, _ = build({"slides": [{"title": "T", "content": [sized, POINTS]}]}, check=True)
        self.assertIsNone(info(b)[0]["arranged"])
        b, _, _ = build({"slides": [{"title": "T", "content": [{"columns": [[FIGURE], [POINTS]]}]}]}, check=True)
        self.assertIsNone(info(b)[0]["arranged"])


class Split(unittest.TestCase):
    def test_splitting_is_the_default_and_can_be_turned_off(self):
        slide = {"title": "Long", "content": [{"bullets": [LONG] * 14}]}
        b, prs, _ = build({"slides": [slide]}, check=True)
        self.assertGreater(len(prs.slides), 1)
        self.assertEqual(b.warnings, [])
        self.assertTrue(titles(prs)[0].endswith(f"(1/{len(prs.slides)})"))
        b, prs, _ = build({"meta": {"overflow": "shrink"}, "slides": [slide]}, check=True)
        self.assertEqual(len(prs.slides), 1)
        self.assertTrue(b.warnings)

    def test_a_stack_of_items_is_divided_between_slides(self):
        content = [{"text": LONG}, {"block": {"title": "Definition", "text": LONG}},
                   {"table": {"header": ["A", "B"], "rows": [[str(i), LONG[:40]] for i in range(5)]}},
                   {"alertblock": {"title": "Caveat", "text": LONG}}, {"bullets": [LONG] * 4}]
        b, prs, _ = build({"slides": [{"title": "Mixed", "content": content}]}, check=True)
        self.assertGreater(len(prs.slides), 1)
        self.assertEqual(b.warnings, [])
        kinds = [k for i in info(b) for k in i["items"]]
        self.assertEqual(kinds, ["text", "block", "table", "alertblock", "bullets"])  # nothing lost, same order
        self.assertTrue(all(i["text_scale"] >= 0.9 for i in info(b)))

    def test_cards_continue_evenly(self):
        cards = [{"title": f"Card {i}", "text": LONG} for i in range(12)]
        b, prs, _ = build({"slides": [{"title": "Cards", "content": [{"cards": cards}]}]})
        self.assertEqual(len(prs.slides), 2)
        self.assertEqual(titles(prs), ["Cards (1/2)", "Cards (2/2)"])
        self.assertEqual(b.warnings, [])

    def test_what_fits_is_not_split(self):
        b, prs, _ = build({"slides": [{"title": "T", "content": [FIGURE, POINTS, {"text": "A remark."}]}]},
                          check=True)
        self.assertEqual(len(prs.slides), 1)


class Style(unittest.TestCase):
    def test_font_sets(self):
        b, _, _ = build({"meta": {"fonts": "palatino"}, "slides": [{"title": "T", "content": [POINTS]}]}, check=True)
        self.assertEqual(b.style["fonts"]["body"], "Palatino Linotype")
        self.assertEqual(b.fb.name, "Palatino Linotype")
        b, prs, _ = build({"meta": {"fonts": {"headings": "Georgia", "code": "Consolas"}},
                           "slides": [{"title": "T", "content": [{"code": "x = 1"}]}]})
        self.assertEqual(b.style["fonts"]["headings"], "Georgia")
        self.assertIn('typeface="Consolas"', prs.slides[0]._element.xml)
        with self.assertRaises(SystemExit) as cm:
            build({"meta": {"fonts": "palatin"}, "slides": []}, check=True)
        self.assertIn("did you mean 'palatino'", str(cm.exception))

    def test_every_font_set_has_measured_widths(self):
        for name, fams in pt.FONT_SETS.items():
            for fam in fams:
                self.assertIn(fam.lower(), metrics.FONTS, f"{name}: {fam}")

    def test_accent_is_made_readable(self):
        for template, accent in (("oxford", "#FFE600"), ("midnight", "#101040"), ("modern", "teal")):
            b, _, _ = build({"meta": {"template": template, "accent": accent}, "slides": [{"title": "T"}]},
                            check=True)
            colors = b._theme()[0]
            self.assertGreaterEqual(pt.contrast(colors["accent1"], colors["bg1"]), 4.4, template)
        b, _, _ = build({"meta": {"accent": "#1F6FEB"}, "slides": [{"title": "T"}]}, check=True)
        self.assertEqual(b.style["accent"], "#1F6FEB")   # already readable: untouched
        with self.assertRaises(SystemExit):
            build({"meta": {"accent": "blurple"}, "slides": []}, check=True)

    def test_density_scales_body_text(self):
        sizes = {}
        for d in ("compact", "normal", "airy"):
            b, _, _ = build({"meta": {"density": d}, "slides": [{"title": "T"}]}, check=True)
            sizes[d] = b.level_sizes[0]
        self.assertLess(sizes["compact"], sizes["normal"])
        self.assertGreater(sizes["airy"], sizes["normal"])

    def test_code_font_comes_from_the_template(self):
        b, prs, _ = build({"meta": {"template": "terminal"}, "slides": [{"title": "T", "content": [{"code": "ls"}]}]})
        self.assertEqual(b.mono, "JetBrains Mono")
        self.assertIn('typeface="JetBrains Mono"', prs.slides[0]._element.xml)


class Navigation(unittest.TestCase):
    DECK = {"slides": [{"layout": "title"}, {"layout": "section", "title": "Intro"},
                       {"title": "One", "content": [POINTS]}, {"layout": "section", "title": "Results"},
                       {"title": "Two", "content": [POINTS]}]}

    def nav(self, slide):
        return next((s for s in slide.shapes if s.name == "Section Navigation"), None)

    def test_sections_are_shown_on_frames(self):
        for template in ("berlin", "berkeley"):
            with self.subTest(template=template):
                b, prs, _ = build({"meta": {"template": template}, **self.DECK})
                self.assertEqual(b.warnings, [])
                self.assertNotIn("{sections}", prs.slide_master._element.xml)
                self.assertIsNone(self.nav(prs.slides[1]))            # not on section slides
                for slide, current in ((prs.slides[2], "Intro"), (prs.slides[4], "Results")):
                    shape = self.nav(slide)
                    self.assertIn("Intro", shape.text_frame.text)
                    self.assertIn("Results", shape.text_frame.text)
                    bold = [r.text for p in shape.text_frame.paragraphs for r in p.runs if r.font.bold]
                    self.assertEqual(bold, [current])

    def test_many_sections_still_fit(self):
        deck = {"meta": {"template": "berlin"}, "slides": [{"layout": "title"}]}
        for i in range(14):
            deck["slides"] += [{"layout": "section", "title": f"A rather long section name {i}"},
                               {"title": f"Slide {i}", "content": [POINTS]}]
        b, prs, _ = build(deck)
        text = self.nav(prs.slides[-1]).text_frame.text
        self.assertIn("section name 13", text)
        self.assertIn("…", text)

    def test_other_templates_have_none(self):
        _, prs, _ = build(self.DECK)
        self.assertIsNone(self.nav(prs.slides[2]))

    def test_sidebar_narrows_the_content_area(self):
        wide = pt.DeckBuilder({"meta": {"template": "oxford"}}).area
        side = pt.DeckBuilder({"meta": {"template": "berkeley"}}).area
        self.assertAlmostEqual(wide.w - side.w, build_template.SIDE_W, places=2)
        self.assertAlmostEqual(side.x - wide.x, build_template.SIDE_W, places=2)


class Report(unittest.TestCase):
    def test_warnings_carry_a_code_and_a_fix(self):
        b, _, _ = build({"meta": {"overflow": "shrink"}, "slides": [
            {"title": "T", "content": [{"bullets": [LONG] * 14}, {"tabel": {"rows": [[1]]}}]}]}, check=True) \
            if False else build({"meta": {"overflow": "shrink"}, "slides": [
                {"title": "T", "colour": "red", "content": [{"bullets": [LONG] * 14}]}]}, check=True)
        codes = {p["code"] for p in b.problems}
        self.assertEqual(codes, {"unknown-key", "text-shrunk"})
        self.assertTrue(all(p["fix"] for p in b.problems))

    def test_every_code_is_documented(self):
        with open(os.path.join(HERE, "AGENTS.md"), encoding="utf-8") as fh:
            text = fh.read()
        for _, code, _ in pt.WARNING_CODES:
            self.assertIn(f"`{code}`", text, code)

    def test_missing_open_font_is_reported_once(self):
        b, _, _ = build({"meta": {"template": "terminal"}, "slides": [
            {"title": "T", "content": [{"code": "ls"}]}, {"title": "U", "content": [{"code": "cd"}]}]}, check=True)
        hits = [p for p in b.problems if p["code"] == "font-missing"]
        self.assertEqual(len(hits), 0 if fonts.installed("JetBrains Mono") is not False else 1)


class Explain(unittest.TestCase):
    def test_everything_can_be_explained(self):
        self.assertEqual(set(pt.ITEM_DOCS), set(pt.ITEM_TYPES))
        self.assertEqual(set(pt.LAYOUT_DOCS), set(pt.LAYOUTS))
        self.assertEqual(set(pt.META_DOCS), set(pt.META_KEYS))
        for name in [*pt.ITEM_TYPES, *pt.LAYOUTS, "meta", None]:
            self.assertTrue(pt.explain_text(name))

    def test_examples_build_without_warnings(self):
        import yaml
        for name, (_, example) in pt.ITEM_DOCS.items():
            if name in ("image", "images", "columns"):   # need figure files
                continue
            with self.subTest(item=name):
                b, _, _ = build({"slides": [{"title": "T", "content": yaml.safe_load(example)}]}, check=True)
                self.assertEqual(b.warnings, [])

    def test_keys_come_from_the_validator(self):
        d = pt.explain("table")
        self.assertEqual(set(d["keys"]), pt.SETTINGS["table"])
        self.assertEqual(d["values"]["style"], list(pt.CHOICES[("table", "style")]))
        with self.assertRaises(SystemExit) as cm:
            pt.explain("tabel")
        self.assertIn("did you mean 'table'", str(cm.exception))

    def test_command_line(self):
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            pt.main(["--explain", "cards", "--json"])
        self.assertEqual(json.loads(out.getvalue())["name"], "cards")


class Fonts(unittest.TestCase):
    def test_every_template_font_has_measured_widths(self):
        for name, path in pt.template_names().items():
            for fam in pt.fonts_of(path):
                self.assertIn(fam.lower(), metrics.FONTS, f"{name}: {fam}")

    def test_open_fonts_are_pinned_and_measured(self):
        for fam, spec in fonts.FAMILIES.items():
            self.assertIn(fam.lower(), metrics.FONTS, fam)
            self.assertTrue(spec["base"].startswith("https://"))
            for f in spec["files"]:
                self.assertRegex(fonts.SHA256.get(f, ""), r"^[0-9a-f]{64}$", f)

    def test_monospace_is_measured_as_monospace(self):
        f = metrics.font("JetBrains Mono")
        self.assertAlmostEqual(f.char("i"), f.char("m"))
        self.assertGreater(metrics.text_lines("x" * 60, 8, 20, "JetBrains Mono"),
                           metrics.text_lines("x" * 60, 8, 20, "Garamond"))

    def test_installed_matches_faces_of_a_family(self):
        fonts._installed_cache = {"palatino linotype bold", "fira sans", "cambria"}
        try:
            self.assertTrue(fonts.installed("Palatino Linotype"))
            self.assertTrue(fonts.installed("Fira Sans"))
            self.assertFalse(fonts.installed("Fira"))
            self.assertEqual(fonts.missing(["Cambria", "LM Roman 10", "LM Roman 10"]), ["LM Roman 10"])
        finally:
            fonts._installed_cache = None


class Docs(unittest.TestCase):
    """The guides are what an agent reads instead of the code: they must name everything."""

    @classmethod
    def setUpClass(cls):
        cls.text = {}
        for name in ("AGENTS.md", "YAML_GUIDE.md", "TEMPLATE_GUIDE.md", "README.md"):
            with open(os.path.join(HERE, name), encoding="utf-8") as fh:
                cls.text[name] = fh.read()

    def test_guides_name_every_template(self):
        for name in pt.template_names():
            for doc in ("AGENTS.md", "README.md"):
                self.assertIn(f"`{name}`", self.text[doc], f"{name} in {doc}")
        self.assertEqual(set(pt.template_names()), set(build_template.THEMES))

    def test_guides_name_every_item_layout_and_meta_key(self):
        for name in [*pt.ITEM_TYPES, *pt.LAYOUTS, *pt.META_KEYS]:
            self.assertIn(f"`{name}", self.text["YAML_GUIDE.md"], name)
        for name in ("fonts", "accent", "density", "arrange", "overflow"):
            self.assertIn(f"`{name}", self.text["AGENTS.md"], name)

    def test_guides_name_every_font_set_and_theme_key(self):
        for name in pt.FONT_SETS:
            self.assertIn(f"`{name}`", self.text["YAML_GUIDE.md"], name)
        for key in build_template.THEME_KEYS:
            self.assertIn(f"`{key}", self.text["TEMPLATE_GUIDE.md"], key)


if __name__ == "__main__":
    unittest.main()
