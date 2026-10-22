"""
media.py - figure files for ppt_templated: aspect ratio, format conversion and caching.

load() reads any figure an `image:` item names; the deck builder then asks for the bytes to embed
at the exact size the figure gets on the slide, so nothing has to be converted by hand:

    .png .jpg .gif .bmp .tif    embedded as they are (EXIF rotation applied, huge photos downsampled)
    .webp and other formats     converted to PNG (anything Pillow reads)
    .svg                        embedded as a vector, plus a PNG fallback drawn at the slide size
    .pdf                        one page (default 1) converted to SVG + PNG with PyMuPDF

PowerPoint 2019 / Microsoft 365 draws the SVG; older versions, LibreOffice and Google Slides show the
PNG. SVG fallbacks are rendered with headless Chrome/Edge (or cairosvg) and cached in .ppt-cache/
next to the YAML, keyed by file content and pixel size, so rebuilding is fast.
"""

from __future__ import annotations

import base64
import hashlib
import io
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
from dataclasses import dataclass

from lxml import etree
from PIL import Image, ImageOps

DPI = 300                     # resolution of rendered PNGs at their size on the slide
MAX_PX = 6000                 # longest side of a rendered PNG
RASTER = {"PNG", "JPEG", "GIF", "BMP", "TIFF"}   # Pillow formats PowerPoint embeds directly
UNITS = {"": 1.0, "px": 1.0, "pt": 96 / 72, "pc": 16.0, "in": 96.0, "cm": 96 / 2.54, "mm": 96 / 25.4,
         "em": 16.0, "ex": 8.0}  # CSS px per unit
SVG_NS = "http://www.w3.org/2000/svg"
XLINK = "http://www.w3.org/1999/xlink"
BROWSERS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "chrome", "msedge", "microsoft-edge",
]


class MediaError(ValueError):
    pass


@dataclass
class Media:
    path: str
    kind: str                  # "raster" | "svg"
    ratio: float               # width / height as displayed
    data: bytes                # raster: bytes to embed; svg: normalised SVG
    px: tuple = (0, 0)         # raster: pixel size
    fmt: str = "PNG"           # raster: Pillow format of `data`
    page: int = 0              # svg converted from this PDF page (1-based), else 0


def load(path: str, page: int = 1) -> Media:
    """Read a figure file. Raises MediaError with a fix-it message for anything unusable."""
    name = os.path.basename(path)
    if not os.path.exists(path):
        raise MediaError(f"image not found: {path} (paths are relative to the YAML file)")
    ext = os.path.splitext(path)[1].lower()
    if ext == ".svg" or ext == ".svgz":
        with open(path, "rb") as fh:
            svg, ratio = normalize_svg(fh.read(), os.path.dirname(os.path.abspath(path)), name)
        return Media(path, "svg", ratio, svg)
    if ext == ".pdf":
        return _pdf(path, int(page or 1))
    if ext in (".eps", ".ps", ".emf", ".wmf"):
        raise MediaError(f"{name}: {ext} is not supported; export the figure as SVG, PDF or PNG")
    try:
        with Image.open(path) as im:
            fmt = im.format
            turned = im.getexif().get(0x0112, 1) not in (1, None)  # photos keep their rotation in EXIF
            if turned or fmt not in RASTER:
                img = ImageOps.exif_transpose(im) if turned else im
                fmt = "JPEG" if fmt in ("JPEG", "MPO") else "PNG"
                if img.mode not in ("RGB", "RGBA", "L", "LA", "P") or fmt == "JPEG":
                    img = img.convert("RGBA" if fmt == "PNG" and "A" in img.mode else "RGB")
                out = io.BytesIO()
                img.save(out, fmt)
                data, size = out.getvalue(), img.size
            else:
                with open(path, "rb") as fh:
                    data = fh.read()
                size = im.size
    except Exception as exc:  # noqa: BLE001
        raise MediaError(f"{name}: cannot read this image ({exc})") from None
    return Media(path, "raster", size[0] / size[1], data, size, fmt)


def raster_bytes(m: Media, w_in: float) -> bytes:
    """Bytes of a raster figure shown w_in inches wide; photos far above DPI are downsampled."""
    need = w_in * DPI
    if m.fmt == "GIF" or m.px[0] <= 2 * need or len(m.data) < 1_500_000:
        return m.data
    with Image.open(io.BytesIO(m.data)) as im:
        small = im.resize((max(1, round(need)), max(1, round(need / m.ratio))), Image.LANCZOS)
        out = io.BytesIO()
        if m.fmt == "JPEG":
            small.convert("RGB").save(out, "JPEG", quality=90, optimize=True)
        else:
            small.save(out, "PNG", optimize=True)
    return out.getvalue()


def png_fallback(m: Media, w_in: float, h_in: float, cache_dir: str | None) -> bytes | None:
    """PNG of a vector figure at exactly its slide size (DPI), from the cache when possible."""
    scale = min(DPI, MAX_PX / max(w_in, h_in))
    w_px, h_px = max(1, round(w_in * scale)), max(1, round(h_in * scale))
    key = hashlib.sha1(m.data + f"|{w_px}x{h_px}".encode()).hexdigest()[:20]
    cached = os.path.join(cache_dir, f"{key}.png") if cache_dir else None
    if cached and os.path.exists(cached):
        with open(cached, "rb") as fh:
            return fh.read()
    png = _pdf_png(m, w_px, h_px) if m.page else (_browser_png(m.data, w_px, h_px) or _cairo_png(m.data, w_px, h_px))
    if png and cached:
        os.makedirs(cache_dir, exist_ok=True)
        with open(cached, "wb") as fh:
            fh.write(png)
    return png


def blank_png() -> bytes:
    out = io.BytesIO()
    Image.new("RGBA", (1, 1), (0, 0, 0, 0)).save(out, "PNG")
    return out.getvalue()


# ------------------------------------------------------------------------------------------
# SVG
# ------------------------------------------------------------------------------------------
def _length(v) -> float | None:
    """SVG length -> CSS px; None for missing, percentages and unknown units."""
    m = re.fullmatch(r"\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*([a-zA-Z%]*)\s*", str(v or ""))
    if not m or m[2].lower() not in UNITS:
        return None
    px = float(m[1]) * UNITS[m[2].lower()]
    return px if px > 0 else None


def normalize_svg(data: bytes, base_dir: str, name: str) -> tuple[bytes, float]:
    """-> (SVG with a viewBox, a width/height in px and linked images inlined, width / height).
    The displayed aspect ratio is width:height when both are absolute, else the viewBox's."""
    if data[:2] == b"\x1f\x8b":  # .svgz
        import gzip
        data = gzip.decompress(data)
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
    try:
        root = etree.fromstring(data, parser)
    except etree.XMLSyntaxError as exc:
        raise MediaError(f"{name}: not a valid SVG file ({exc})") from None
    if etree.QName(root).localname != "svg":
        raise MediaError(f"{name}: the root element is <{etree.QName(root).localname}>, not <svg>")
    vb = [float(x) for x in re.split(r"[\s,]+", (root.get("viewBox") or "").strip()) if x]
    w, h = _length(root.get("width")), _length(root.get("height"))
    if len(vb) == 4 and vb[2] > 0 and vb[3] > 0:
        if w is None and h is None:
            w, h = vb[2], vb[3]
        elif w is None:
            w = h * vb[2] / vb[3]
        elif h is None:
            h = w * vb[3] / vb[2]
    elif w and h:
        root.set("viewBox", f"0 0 {w:g} {h:g}")  # lets the drawing scale to any size
    else:
        raise MediaError(f'{name}: the SVG needs a viewBox (e.g. viewBox="0 0 800 450") or a width and height')
    root.set("width", f"{w:g}")
    root.set("height", f"{h:g}")
    for el in root.iter(f"{{{SVG_NS}}}image"):  # linked pictures would be lost inside the deck: inline them
        for attr in ("href", f"{{{XLINK}}}href"):
            ref = el.get(attr)
            if not ref or re.match(r"(?i)(data:|https?:|#)", ref):
                continue
            p = urllib.parse.unquote(ref[7:] if ref.startswith("file://") else ref)
            p = p if os.path.isabs(p) else os.path.join(base_dir, p)
            if os.path.exists(p):
                with open(p, "rb") as fh:
                    b64 = base64.b64encode(fh.read()).decode()
                el.set(attr, f"data:{mimetypes.guess_type(p)[0] or 'image/png'};base64,{b64}")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8"), w / h


def _sized(svg: bytes, w_px: int, h_px: int) -> bytes:
    root = etree.fromstring(svg, etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True))
    root.set("width", str(w_px))
    root.set("height", str(h_px))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def find_browser() -> str | None:
    for c in BROWSERS:
        p = c if os.path.isabs(c) else shutil.which(c)
        if p and os.path.exists(p):
            return p
    return None


def _browser_png(svg: bytes, w_px: int, h_px: int) -> bytes | None:
    exe = find_browser()
    if not exe:
        return None
    work = tempfile.mkdtemp(prefix="ppt_svg_")
    try:
        src, out = os.path.join(work, "figure.svg"), os.path.join(work, "figure.png")
        with open(src, "wb") as fh:
            fh.write(_sized(svg, w_px, h_px))
        # a private profile keeps headless Chrome from handing the job to a running browser
        cmd = [exe, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
               "--no-default-browser-check", f"--user-data-dir={os.path.join(work, 'profile')}",
               "--force-device-scale-factor=1", "--default-background-color=00000000",
               f"--window-size={w_px},{h_px}", f"--screenshot={out}", "file:///" + src.replace("\\", "/").lstrip("/")]
        if sys.platform.startswith("linux") and hasattr(os, "geteuid") and os.geteuid() == 0:
            cmd.insert(1, "--no-sandbox")
        subprocess.run(cmd, capture_output=True, timeout=120)
        if not os.path.exists(out):
            return None
        with Image.open(out) as im:
            if im.size != (w_px, h_px):
                im = im.resize((w_px, h_px), Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "PNG", optimize=True)
            return buf.getvalue()
    except Exception:  # noqa: BLE001 - any failure means "try the next renderer"
        return None
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _cairo_png(svg: bytes, w_px: int, h_px: int) -> bytes | None:
    try:
        import cairosvg  # type: ignore
        return cairosvg.svg2png(bytestring=svg, output_width=w_px, output_height=h_px)
    except Exception:  # noqa: BLE001 - not installed or no cairo library
        return None


# ------------------------------------------------------------------------------------------
# PDF (PyMuPDF)
# ------------------------------------------------------------------------------------------
def _pymupdf():
    try:
        import pymupdf  # type: ignore
    except ImportError:
        try:
            import fitz as pymupdf  # type: ignore  # older PyMuPDF
        except ImportError:
            raise MediaError("PDF figures need PyMuPDF: pip install pymupdf "
                             "(or export the figure as SVG or PNG)") from None
    return pymupdf


def _pdf(path: str, page: int) -> Media:
    pymupdf = _pymupdf()
    name = os.path.basename(path)
    with pymupdf.open(path) as doc:
        if not 1 <= page <= doc.page_count:
            raise MediaError(f"{name} has {doc.page_count} page(s); page {page} does not exist")
        svg = doc[page - 1].get_svg_image(text_as_path=True).encode()
    data, ratio = normalize_svg(svg, os.path.dirname(os.path.abspath(path)), name)
    return Media(path, "svg", ratio, data, page=page)


def _pdf_png(m: Media, w_px: int, h_px: int) -> bytes | None:
    pymupdf = _pymupdf()
    with pymupdf.open(m.path) as doc:
        pg = doc[m.page - 1]
        pix = pg.get_pixmap(matrix=pymupdf.Matrix(w_px / pg.rect.width, h_px / pg.rect.height), alpha=True)
        return pix.tobytes("png")
