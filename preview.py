#!/usr/bin/env python3
"""
preview.py - render a .pptx to PNG images, so a deck can be checked without opening PowerPoint.

    python preview.py talk.pptx                    # -> .preview/talk/slide-01.png ... and sheet-1.png ...
    python ppt_templated.py talk.yaml --preview    # build, then the same

Renderer: PowerPoint through COM on Windows (exact), else LibreOffice (any OS; close, but equations
and some effects differ). A contact sheet puts 6 numbered slides on one image for a quick look.
With PowerPoint, the rendered layout is also measured: text that runs into another shape, out of its
block or below the content area is reported (ppt_templated's sizes are estimates; these are not).
"""

from __future__ import annotations

import argparse
import base64
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

WIDTH = 1280          # slide image width (px)
PER_SHEET, COLS, TILE = 6, 2, 800

POWERSHELL = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$app = New-Object -ComObject PowerPoint.Application
$deck = $app.Presentations.Open('__DECK__', -1, 0, 0)
try {
  $h = [int][Math]::Round(__WIDTH__ * $deck.PageSetup.SlideHeight / $deck.PageSetup.SlideWidth)
  $bottom = __BOTTOM__
  if ($bottom -le 0) { $bottom = $deck.PageSetup.SlideHeight }
  foreach ($s in $deck.Slides) {
    $s.Export((Join-Path '__OUT__' ('slide-{0:D2}.png' -f $s.SlideIndex)), 'PNG', __WIDTH__, $h)
    # measured extents (pt) of everything but the title: rendered text included, connectors left out
    $boxes = New-Object System.Collections.Generic.List[object]
    foreach ($sh in $s.Shapes) {
      if ($sh.Type -eq 14) { $pt = $sh.PlaceholderFormat.Type; if ($pt -eq 1 -or $pt -eq 3) { continue } }
      $parts = if ($sh.Type -eq 6) { @($sh.GroupItems) } else { @($sh) }
      $l = 1e9; $t = 1e9; $r = -1e9; $b = -1e9; $label = $sh.Name
      foreach ($p in $parts) {
        if ($p.Connector) { continue }
        $pt = $p.Top; $pb = $p.Top + $p.Height
        if ($p.HasTextFrame -and $p.TextFrame2.HasText) {
          $tr = $p.TextFrame2.TextRange
          $tb = $tr.BoundTop + $tr.BoundHeight
          if (-not $p.Fill.Visible -and -not $p.Line.Visible) { $pt = $tr.BoundTop; $pb = $tb }  # plain text: its ink
          if ($sh.Type -eq 6 -and $tb -gt $sh.Top + $sh.Height + 2) {
            Write-Output ("overflow`t{0}`t{1}`t{2}" -f $s.SlideIndex, $sh.Name, [int]($tb - $sh.Top - $sh.Height)) }
          $pb = [Math]::Max($pb, $tb)
          if ($label -eq $sh.Name) { $label = ($tr.Text -replace '\s+', ' ').Trim() }
        }
        $l = [Math]::Min($l, $p.Left); $t = [Math]::Min($t, $pt); $r = [Math]::Max($r, $p.Left + $p.Width)
        $b = [Math]::Max($b, $pb)
      }
      if ($r -lt $l) { continue }
      $boxes.Add(@($label, $l, $t, $r, $b, ($sh.Name -match '^(Node|Step|Group|Label|Arrow): ')))
      if ($b -gt $bottom + 3) { Write-Output ("bottom`t{0}`t{1}`t{2}" -f $s.SlideIndex, $label, [int]($b - $bottom)) }
    }
    for ($i = 0; $i -lt $boxes.Count; $i++) { for ($j = $i + 1; $j -lt $boxes.Count; $j++) {
      $a = $boxes[$i]; $c = $boxes[$j]
      if ($a[5] -and $c[5]) { continue }  # parts of one diagram (chevrons interlock, groups hold nodes)
      $ox = [Math]::Min($a[3], $c[3]) - [Math]::Max($a[1], $c[1])
      $oy = [Math]::Min($a[4], $c[4]) - [Math]::Max($a[2], $c[2])
      if ($ox -gt 2 -and $oy -gt 2) { Write-Output ("overlap`t{0}`t{1}`t{2}`t{3}" -f $s.SlideIndex, $a[0], $c[0], [int]$oy) }
    } }
  }
} finally {
  $deck.Close()
  if ($app.Presentations.Count -eq 0) { $app.Quit() }
}
"""
SOFFICE = ["soffice", "libreoffice", r"C:\Program Files\LibreOffice\program\soffice.exe",
           r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
           "/Applications/LibreOffice.app/Contents/MacOS/soffice"]


def preview(pptx: str, outdir: str | None = None, width: int = WIDTH, bottom: float | None = None):
    """Render every slide and the contact sheets. -> (slide PNGs, sheet PNGs, renderer name, layout
    problems). bottom: lowest edge content may reach, in inches (default: the slide's bottom edge)."""
    pptx = os.path.abspath(pptx)
    outdir = os.path.abspath(outdir or os.path.join(os.path.dirname(pptx), ".preview", Path(pptx).stem))
    os.makedirs(outdir, exist_ok=True)
    for old in glob.glob(os.path.join(outdir, "slide-*.png")) + glob.glob(os.path.join(outdir, "sheet-*.png")):
        os.remove(old)
    errors = []
    for name, fn in (("PowerPoint", _powerpoint), ("LibreOffice", _libreoffice)):
        try:
            found = fn(pptx, outdir, width, bottom)
            slides = sorted(glob.glob(os.path.join(outdir, "slide-*.png")), key=_number)
            if found is not None and slides:
                return slides, sheets(slides, outdir), name, found
        except Exception as exc:  # noqa: BLE001 - try the next renderer
            errors.append(f"{name}: {exc}")
    raise RuntimeError("cannot render slides: needs PowerPoint (Windows) or LibreOffice"
                       + ("".join(f"\n  {e}" for e in errors) if errors else ""))


def _problems(lines):
    """PowerShell check lines -> readable layout problems."""
    out = []

    def short(s):  # an equation's text is a jumble of math symbols: name it instead
        if sum(ord(c) > 0x2000 for c in s) > len(s) / 5:
            return "an equation"
        return f'"{s[:45]}..."' if len(s) > 48 else f'"{s}"'

    for line in lines:
        f = line.split("\t")
        if f[0] == "overlap" and len(f) == 5:
            out.append(f"slide {f[1]}: {short(f[2])} runs into {short(f[3])} ({f[4]} pt)")
        elif f[0] == "overflow" and len(f) == 4:
            out.append(f"slide {f[1]}: text runs {f[3]} pt out of the bottom of {short(f[2])}")
        elif f[0] == "bottom" and len(f) == 4:
            out.append(f"slide {f[1]}: {short(f[2])} runs {f[3]} pt below the content area")
    return out


def sheets(slides, outdir, per=PER_SHEET, cols=COLS, tile=TILE):
    """Contact sheets: `per` slides per image, `cols` across, numbered in the top-right corner."""
    out = []
    try:
        font = ImageFont.load_default(size=20)
    except TypeError:  # Pillow < 10.1
        font = ImageFont.load_default()
    for k in range(0, len(slides), per):
        chunk = slides[k:k + per]
        ims = [Image.open(p).convert("RGB") for p in chunk]
        th = round(tile * ims[0].height / ims[0].width)
        rows = -(-len(ims) // cols)
        pad = 8
        sheet = Image.new("RGB", (cols * tile + (cols + 1) * pad, rows * th + (rows + 1) * pad), "#6E6E6E")
        draw = ImageDraw.Draw(sheet)
        for i, (p, im) in enumerate(zip(chunk, ims)):
            x, y = pad + (i % cols) * (tile + pad), pad + (i // cols) * (th + pad)
            sheet.paste(im.resize((tile, th), Image.LANCZOS), (x, y))
            label = str(_number(p))
            box = draw.textbbox((0, 0), label, font=font)
            x1 = x + tile - box[2] - 12
            draw.rectangle((x1, y, x + tile, y + box[3] + 8), fill="#C00000")
            draw.text((x1 + 6, y + 3), label, fill="white", font=font)
        path = os.path.join(outdir, f"sheet-{k // per + 1}.png")
        sheet.save(path)
        out.append(path)
    return out


def _number(path) -> int:
    m = re.search(r"(\d+)\.png$", path)
    return int(m.group(1)) if m else 0


def _powerpoint(pptx, outdir, width, bottom):
    """Export through PowerPoint and measure the layout -> problems, or None without PowerPoint."""
    if sys.platform != "win32":
        return None
    shell = shutil.which("powershell") or shutil.which("pwsh")
    if not shell:
        return None
    q = lambda s: s.replace("'", "''")  # noqa: E731 - PowerShell single-quote escaping
    script = (POWERSHELL.replace("__DECK__", q(pptx)).replace("__OUT__", q(outdir))
              .replace("__WIDTH__", str(int(width))).replace("__BOTTOM__", f"{(bottom or 0) * 72:.1f}"))
    r = subprocess.run([shell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                        "-EncodedCommand", base64.b64encode(script.encode("utf-16-le")).decode()],
                       capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900)
    if r.returncode != 0:
        msg = (r.stderr or r.stdout).strip().splitlines()
        raise RuntimeError(msg[0] if msg else f"PowerShell exit code {r.returncode}")
    return _problems(r.stdout.splitlines())


def _libreoffice(pptx, outdir, width, bottom):
    """Export through LibreOffice -> [] (no layout measurements), or None without LibreOffice."""
    exe = next((p for p in (c if os.path.isabs(c) else shutil.which(c) for c in SOFFICE) if p and os.path.exists(p)), None)
    if not exe:
        return None
    work = tempfile.mkdtemp(prefix="ppt_preview_")
    try:
        # a private profile: a running LibreOffice would otherwise swallow the job
        subprocess.run([exe, f"-env:UserInstallation={Path(work, 'profile').as_uri()}", "--headless",
                        "--convert-to", "pdf", "--outdir", work, pptx], capture_output=True, timeout=900)
        pdf = os.path.join(work, Path(pptx).stem + ".pdf")
        if not os.path.exists(pdf):
            raise RuntimeError("LibreOffice did not produce a PDF")
        try:
            import pymupdf  # type: ignore
        except ImportError:
            try:
                import fitz as pymupdf  # type: ignore
            except ImportError:
                pymupdf = None
        if pymupdf is not None:
            with pymupdf.open(pdf) as doc:
                for i, page in enumerate(doc, 1):
                    z = width / page.rect.width
                    page.get_pixmap(matrix=pymupdf.Matrix(z, z)).save(os.path.join(outdir, f"slide-{i:02d}.png"))
            return []
        if shutil.which("pdftoppm"):
            subprocess.run(["pdftoppm", "-png", "-scale-to-x", str(width), "-scale-to-y", "-1", pdf,
                            os.path.join(work, "p")], check=True, capture_output=True, timeout=900)
            for p in glob.glob(os.path.join(work, "p-*.png")):
                shutil.move(p, os.path.join(outdir, f"slide-{_number(p):02d}.png"))
            return []
        raise RuntimeError("install PyMuPDF (pip install pymupdf) or poppler's pdftoppm to turn the PDF into PNGs")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Render a .pptx to slide PNGs and contact sheets.")
    ap.add_argument("pptx")
    ap.add_argument("-o", "--outdir", help="default: .preview/<deck name>/ next to the deck")
    ap.add_argument("--width", type=int, default=WIDTH, help=f"slide image width in px (default {WIDTH})")
    a = ap.parse_args(argv)
    try:
        slides, sheet_paths, engine, problems = preview(a.pptx, a.outdir, a.width)
    except RuntimeError as exc:
        sys.exit(f"error: {exc}")
    for p in problems:
        print(f"  layout: {p}", file=sys.stderr)
    print(f"preview ({engine}): {len(slides)} slides in {os.path.dirname(slides[0])}")
    for p in sheet_paths:
        print(f"  {p}")


if __name__ == "__main__":
    main()
