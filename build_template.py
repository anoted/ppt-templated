#!/usr/bin/env python3
"""
build_template.py - generate Beamer-inspired PowerPoint templates.

    python build_template.py                 # builds all themes into templates/
    python build_template.py --theme blue    # just one

Each template is a normal .pptx whose slide master and layouts carry all the formatting:
  * theme colours  (accent1 = structure, accent2 = alert, accent3 = example, accent4-6 = charts)
  * theme fonts    (headings / body)
  * layouts with the standard PowerPoint names ppt_templated.py looks for:
      Title Slide, Section Header, Title and Content, Two Content, Title Only, Blank
  * a Beamer "footline" (author | short title | date + slide number) whose text uses tokens
    such as {author} that ppt_templated.py fills in.

You can also open the result in PowerPoint (View > Slide Master) and restyle it by hand -
see TEMPLATE_GUIDE.md. Add your own colour theme by copying an entry in THEMES.
"""

import argparse
import copy
import os
import sys

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
}
# Optional style keys (defaults give the Madrid look above):
#   frame      "bar" (filled title bar) | "rule" (hairline under the title) | "plain"
#   title_color / bar_text   frame-title colour;  bar_stripe   thin stripe under the bar
#   rule       colour of the title hairline (frame: rule)
#   footline   "boxes" (coloured segments, uses foot/foot_text) | "text" (grey text only)
#   foot_rule  hairline above a text footline;  foot_stripe  stripe above a boxed footline
#   cover      "box" (filled title box) | "rules" (rules above/below the title) | "plain"
#   cover_text title colour on Title/Section slides;  radius  title-box corner (0 = square)
#   bullets    [(char, size %), ...] for levels 1-3 (repeated below);  bullet_color
#   title_bold, subtitle_italic

# Geometry (inches) - change here to move things in every layout at once
BAR_H = 0.95
TITLE = (0.45, 0.1, W - 0.9, 0.75)
BODY = (0.6, 1.3, W - 1.2, 5.6)
FOOT_H = 0.32
FOOT_Y = H - FOOT_H
COL_W = (BODY[2] - 0.4) / 2
RULE = 0.014    # hairline (~1 pt)
STRIPE = 0.05   # accent stripe under the bar / above the footline


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
        color, margin = ("tx1", 50, 50), BODY[0]
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


def tree(content, name, bg=True):
    bg_xml = ('<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
              if bg else "")
    return (f'<p:cSld {NS} name="{name}">{bg_xml}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>'
            f'<p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            f'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{content}</p:spTree></p:cSld>')


def lvl(n, size, color="tx1", algn=None, extra="", attrs="", rattrs=""):
    a = f' algn="{algn}"' if algn else ""
    return (f'<a:lvl{n}pPr{a}{attrs}>{extra}<a:defRPr sz="{size}"{rattrs}><a:solidFill>{clr(color)}</a:solidFill>'
            f'</a:defRPr></a:lvl{n}pPr>')


# ------------------------------------------------------------------------------------------
def build(theme_key, out_path):
    t = THEMES[theme_key]
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
    for key, hexv in t["colors"].items():
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
    master_tree = tree(deco + footline(t, 11) + title_ph + body_ph, "")
    mel = master._element
    mel.replace(mel.find(qn("p:cSld")), parse_xml(master_tree))

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
    tbox_w = 9.8
    tbox_x = (W - tbox_w) / 2
    radius = f'<a:gd name="adj" fmla="val {t.get("radius", 12000)}"/>'
    cover = t.get("cover", "box")
    cover_text = t.get("cover_text", "bg1" if cover == "box" else "accent1")
    box = dict(geom=("roundRect", radius), fill="accent1") if cover == "box" else {}
    sub_attrs = ' i="1"' if t.get("subtitle_italic") else ""

    def rules(top, bottom, x, w):
        """booktabs-style: heavier rule above the title, lighter rule below it."""
        if cover != "rules":
            return ""
        return (shape(30, "Title Rule Top", (x, top, w, RULE * 2), fill=cover_text)
                + shape(31, "Title Rule Bottom", (x, bottom, w, RULE), fill=cover_text))

    title_slide = (
        ph(2, "Title 1", 'type="ctrTitle"', (tbox_x, 1.45, tbox_w, 1.45), **box,
           prompt="Presentation Title", lst=lvl(1, 3600, cover_text, "ctr"),
           body_pr=f'lIns="{e(0.3)}" rIns="{e(0.3)}" anchor="ctr"', autofit=True)
        + rules(1.35, 2.95, tbox_x, tbox_w)
        + ph(3, "Subtitle 2", 'type="subTitle" idx="1"', (tbox_x, 3.0, tbox_w, 0.6), prompt="Subtitle",
             lst=lvl(1, 2200, "accent1", "ctr", no_bullet, ' marL="0" indent="0"', sub_attrs), body_pr='anchor="t"')
        + ph(4, "Author", 'type="body" sz="quarter" idx="13"', (tbox_x, 3.95, tbox_w, 0.5), prompt="Author Name",
             lst=lvl(1, 2200, "tx1", "ctr", no_bullet, ' marL="0" indent="0"'), body_pr='anchor="ctr"')
        + ph(5, "Institute", 'type="body" sz="quarter" idx="14"', (tbox_x, 4.5, tbox_w, 0.5), prompt="Institute",
             lst=lvl(1, 1800, ("tx1", 65, 35), "ctr", no_bullet, ' marL="0" indent="0"'), body_pr='anchor="ctr"')
        + ph(6, "Date", 'type="body" sz="quarter" idx="15"', (tbox_x, 5.2, tbox_w, 0.45), prompt="Date",
             lst=lvl(1, 1800, "tx1", "ctr", no_bullet, ' marL="0" indent="0"'), body_pr='anchor="ctr"')
        + footline(t, 20))
    section = (
        ph(2, "Title 1", 'type="title"', (2.9, 2.85, W - 5.8, 1.15), **box,
           prompt="Section Title", lst=lvl(1, 3200, cover_text, "ctr"),
           body_pr=f'lIns="{e(0.3)}" rIns="{e(0.3)}" anchor="ctr"', autofit=True)
        + rules(2.75, 4.04, 2.9, W - 5.8)
        + ph(3, "Section Subtitle", 'type="body" idx="1"', (2.9, 4.15, W - 5.8, 0.6), prompt="Optional subtitle",
             lst=lvl(1, 2000, "accent1", "ctr", no_bullet, ' marL="0" indent="0"', sub_attrs), body_pr='anchor="t"')
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
        xml = tree(body, lay.name, bg=False)
        lel = lay._element
        lel.replace(lel.find(qn("p:cSld")), parse_xml(xml))
        if show_master:
            lel.attrib.pop("showMasterSp", None)
        else:
            lel.set("showMasterSp", "0")
    prs.core_properties.title = t["label"]
    prs.core_properties.author = "ppt-templated"
    prs.save(out_path)
    return out_path


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


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--theme", choices=list(THEMES), help="build only this theme")
    ap.add_argument("--outdir", default=os.path.join(HERE, "templates"))
    ap.add_argument("--no-preview", action="store_true", help="don't add the layout preview slides")
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    for key in ([a.theme] if a.theme else THEMES):
        out = os.path.join(a.outdir, f"beamer_{key}.pptx")
        build(key, out)
        if not a.no_preview:
            import ppt_templated  # preview slides are made by the generator itself
            b = ppt_templated.DeckBuilder(copy.deepcopy(DEMO), out, HERE, substitute_tokens=False)
            b.build().save(out)
        print("wrote", out)


if __name__ == "__main__":
    main()
