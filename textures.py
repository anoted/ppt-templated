"""
textures.py - background pictures for templates, drawn from numbers (no image files to find or license).

build_template.py calls texture() for a theme with `background: {texture: ...}`. Every texture is made
from the theme's own colours and a fixed random seed, so rebuilding a template gives the same bytes.

    paper   warm sheet: soft mottling, fine grain, faint fibres, darker edges
    grid    graph paper: thin lines every 1/4 in, a stronger line every inch
    dots    dot grid every 1/4 in (bullet-journal paper)
    lines   ruled paper: a line every 1/3 in and a margin line on the left
    chalk   blackboard: smudges, chalk dust, faint wipe marks (for dark themes)
    grain   plain colour with a fine film grain (light or dark)

The result is a JPEG of the whole slide (1920 x 1080 for 16:9), stretched over the master background.
"""

from __future__ import annotations

import io

import numpy as np
from PIL import Image, ImageFilter

TEXTURES = ("paper", "grid", "dots", "lines", "chalk", "grain")
PX_PER_IN = 144  # 1920 px over a 13.33 in slide


def _rgb(hex_rgb: str) -> np.ndarray:
    h = hex_rgb.lstrip("#")
    return np.array([int(h[i:i + 2], 16) for i in (0, 2, 4)], float)


def _smooth(rng, w, h, cell, amp):
    """Low-frequency noise: random values on a coarse grid, enlarged with bicubic smoothing."""
    gw, gh = max(2, w // cell + 2), max(2, h // cell + 2)
    small = Image.fromarray(((rng.random((gh, gw)) * 255).astype(np.uint8)), "L")
    big = np.asarray(small.resize((gw * cell, gh * cell), Image.BICUBIC), float)[:h, :w] / 255 - 0.5
    return big * 2 * amp


def _vignette(w, h, amount):
    y, x = np.mgrid[0:h, 0:w]
    d = np.sqrt(((x - w / 2) / (w / 2)) ** 2 + ((y - h / 2) / (h / 2)) ** 2) / np.sqrt(2)
    return -amount * np.clip(d - 0.35, 0, None) ** 1.6 / 0.65 ** 1.6


def texture(kind: str, base: str, ink: str, size=(1920, 1080), strength: float = 1.0, seed: int = 7) -> bytes:
    """JPEG bytes of a background texture. base: background colour 'RRGGBB'; ink: colour of lines, dots
    and fibres; strength scales the effect (1 = default, 0.5 = subtler, 2 = stronger)."""
    if kind not in TEXTURES:
        raise ValueError(f"texture must be one of {', '.join(TEXTURES)}, got {kind!r}")
    w, h = size
    rng = np.random.default_rng(seed)
    b, k = _rgb(base), _rgb(ink)
    dark = b.mean() < 128
    s = float(strength)
    light = np.zeros((h, w))          # brightness offset, in units of 255
    mix = np.zeros((h, w))            # 0..1 share of the ink colour

    if kind in ("paper", "chalk", "grain"):
        light += rng.normal(0, 2.2 if kind == "grain" else 1.6, (h, w)) * s
    if kind == "paper":
        light += (_smooth(rng, w, h, 240, 5.0) + _smooth(rng, w, h, 60, 2.2)) * s
        light += _vignette(w, h, 12.0) * s
        fib = Image.new("L", (w, h), 0)
        px = fib.load()
        for _ in range(int(900 * s)):  # short, slightly curved fibres
            x, y = rng.uniform(0, w), rng.uniform(0, h)
            ang, ln = rng.uniform(0, np.pi), rng.uniform(8, 30)
            bend = rng.normal(0, 0.03)
            for t in range(int(ln)):
                a = ang + bend * t
                xx, yy = int(x + np.cos(a) * t), int(y + np.sin(a) * t)
                if 0 <= xx < w and 0 <= yy < h:
                    px[xx, yy] = 110
        mix += np.asarray(fib.filter(ImageFilter.GaussianBlur(0.6)), float) / 255 * 0.16 * s
    elif kind == "chalk":
        light += (_smooth(rng, w, h, 320, 7.0) + _smooth(rng, w, h, 90, 3.0)) * s
        streaks = _smooth(rng, w, h * 1, 400, 1.0)
        streaks = np.repeat(streaks.mean(axis=1, keepdims=True), w, axis=1) * 5.0  # horizontal wipe marks
        light += streaks * s
        dust = (rng.random((h, w)) > 0.9994).astype(float)
        dust = np.asarray(Image.fromarray((dust * 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(0.8)),
                          float) / 255
        light += dust * 80 * s
        light += _vignette(w, h, 10.0) * s
    elif kind in ("grid", "dots", "lines"):
        light += rng.normal(0, 0.8, (h, w)) * s
        step = PX_PER_IN / 4 if kind != "lines" else PX_PER_IN / 3
        y, x = np.mgrid[0:h, 0:w]
        if kind == "grid":
            fx, fy = (x % step), (y % step)
            thin = ((fx < 1.2) | (fy < 1.2)).astype(float) * 0.22
            big = (((x % (step * 4)) < 1.6) | ((y % (step * 4)) < 1.6)).astype(float) * 0.40
            mix += np.maximum(thin, big) * s
        elif kind == "dots":
            dx, dy = (x + step / 2) % step - step / 2, (y + step / 2) % step - step / 2
            r = np.sqrt(dx ** 2 + dy ** 2)
            mix += np.clip(2.2 - r, 0, 1) * 0.55 * s
        else:
            mix += ((y % step) < 1.4).astype(float) * 0.30 * s
            margin = int(PX_PER_IN * 1.0)
            red = np.array([205.0, 70.0, 70.0]) if not dark else np.array([230.0, 120.0, 120.0])
            band = (np.abs(x - margin) < 1.2).astype(float) * 0.45 * s
            img = b + (k - b) * np.clip(mix, 0, 1)[..., None] + light[..., None]
            img = img + (red - img) * band[..., None]
            return _jpeg(img)
    img = b + (k - b) * np.clip(mix, 0, 1)[..., None] + light[..., None]
    return _jpeg(img)


def _jpeg(img: np.ndarray) -> bytes:
    out = io.BytesIO()
    Image.fromarray(np.clip(img, 0, 255).astype(np.uint8), "RGB").save(out, "JPEG", quality=88, optimize=True,
                                                                         subsampling=0)
    return out.getvalue()
