#!/usr/bin/env python3
"""
fonts.py - the open fonts some templates use: what they are, whether they are installed, installing them.

    python fonts.py                      # every font the bundled templates use: installed or not
    python fonts.py --install            # download and install all open fonts (for your user only)
    python fonts.py --install latex      # only the fonts of one template (or name a family: "Fira Sans")
    python fonts.py --load               # Windows: usable until you sign out; nothing is installed
    python fonts.py --download           # only fetch the files into fonts/ (no install)

Most templates use fonts that come with Microsoft Office. A few use open fonts (SIL Open Font License or
GUST Font License: free to use, share and embed) that give looks Office cannot: LaTeX's own typefaces,
the real Metropolis theme, console and code-editor decks. PowerPoint shows a font only when it is
installed on the machine that opens the deck, so:

  * install them once here (`--install`), and
  * to share a deck, either send the PDF or tick File > Options > Save > "Embed fonts in the file".

Text is laid out from width tables stored in metrics.py, so a deck is built the same whether or not the
font is installed; only the look in PowerPoint depends on it. Files are downloaded from the projects'
own repositories at pinned versions and checked against the SHA-256 sums below. Nothing is downloaded
unless you run --install or --download.
"""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "fonts")

_GF = "https://raw.githubusercontent.com/google/fonts/c30b221beb0c1da59d563feb91ec8656a1fa972a/ofl/"
_LM = "https://mirrors.ctan.org/fonts/lm/fonts/opentype/public/lm/"
_PLEX = "https://raw.githubusercontent.com/IBM/plex/1da12f02587b630c07e92692d21492d722f53614/packages/"
_OFL = "SIL Open Font License 1.1"

# family name (as PowerPoint lists it) -> kind, licence, home page, base URL, files (regular first, then bold)
FAMILIES = {
    "LM Roman 10": dict(kind="serif", licence="GUST Font License", base=_LM,
                        home="https://www.gust.org.pl/projects/e-foundry/latin-modern",
                        note="Latin Modern Roman: the typeface of LaTeX documents (Computer Modern)",
                        files=["lmroman10-regular.otf", "lmroman10-bold.otf", "lmroman10-italic.otf",
                               "lmroman10-bolditalic.otf"]),
    "LM Sans 10": dict(kind="sans", licence="GUST Font License", base=_LM,
                       home="https://www.gust.org.pl/projects/e-foundry/latin-modern",
                       note="Latin Modern Sans: the typeface of LaTeX Beamer slides",
                       files=["lmsans10-regular.otf", "lmsans10-bold.otf", "lmsans10-oblique.otf",
                              "lmsans10-boldoblique.otf"]),
    "LM Mono 10": dict(kind="mono", licence="GUST Font License", base=_LM,
                       home="https://www.gust.org.pl/projects/e-foundry/latin-modern",
                       note="Latin Modern Mono: LaTeX's typewriter face (\\texttt, verbatim)",
                       files=["lmmono10-regular.otf", "lmmono10-italic.otf"]),
    "STIX Two Text": dict(kind="serif", licence=_OFL, home="https://www.stixfonts.org",
                          base="https://raw.githubusercontent.com/stipub/stixfonts/v2.13b171/fonts/static_ttf/",
                          note="the Times-like face of science publishers (AIP, APS, IEEE, Elsevier)",
                          files=["STIXTwoText-Regular.ttf", "STIXTwoText-Bold.ttf", "STIXTwoText-Italic.ttf",
                                 "STIXTwoText-BoldItalic.ttf"]),
    "Source Serif 4": dict(kind="serif", licence=_OFL, home="https://github.com/adobe-fonts/source-serif",
                           base="https://raw.githubusercontent.com/adobe-fonts/source-serif/4.005R/TTF/",
                           note="Adobe's text serif: sturdy, reads well on screens",
                           files=["SourceSerif4-Regular.ttf", "SourceSerif4-Bold.ttf", "SourceSerif4-It.ttf",
                                  "SourceSerif4-BoldIt.ttf"]),
    "Source Sans 3": dict(kind="sans", licence=_OFL, home="https://github.com/adobe-fonts/source-sans",
                          base="https://raw.githubusercontent.com/adobe-fonts/source-sans/3.052R/TTF/",
                          note="Adobe's interface sans, the companion of Source Serif",
                          files=["SourceSans3-Regular.ttf", "SourceSans3-Bold.ttf", "SourceSans3-It.ttf",
                                 "SourceSans3-BoldIt.ttf"]),
    "Source Code Pro": dict(kind="mono", licence=_OFL, home="https://github.com/adobe-fonts/source-code-pro",
                            base="https://raw.githubusercontent.com/adobe-fonts/source-code-pro/"
                                 "2.042R-u/1.062R-i/1.026R-vf/TTF/",
                            note="Adobe's coding face",
                            files=["SourceCodePro-Regular.ttf", "SourceCodePro-Bold.ttf", "SourceCodePro-It.ttf",
                                   "SourceCodePro-BoldIt.ttf"]),
    "Fira Sans": dict(kind="sans", licence=_OFL, home="https://github.com/mozilla/Fira", base=_GF + "firasans/",
                      note="the typeface of Beamer's Metropolis theme",
                      files=["FiraSans-Regular.ttf", "FiraSans-Bold.ttf", "FiraSans-Italic.ttf",
                             "FiraSans-BoldItalic.ttf"]),
    "Fira Mono": dict(kind="mono", licence=_OFL, home="https://github.com/mozilla/Fira", base=_GF + "firamono/",
                      note="Fira's coding face (no italic)",
                      files=["FiraMono-Regular.ttf", "FiraMono-Bold.ttf"]),
    "IBM Plex Sans": dict(kind="sans", licence=_OFL, home="https://github.com/IBM/plex",
                          base=_PLEX + "plex-sans/fonts/complete/ttf/",
                          note="IBM's engineered sans",
                          files=["IBMPlexSans-Regular.ttf", "IBMPlexSans-Bold.ttf", "IBMPlexSans-Italic.ttf",
                                 "IBMPlexSans-BoldItalic.ttf"]),
    "IBM Plex Mono": dict(kind="mono", licence=_OFL, home="https://github.com/IBM/plex",
                          base=_PLEX + "plex-mono/fonts/complete/ttf/",
                          note="IBM's typewriter-flavoured coding face",
                          files=["IBMPlexMono-Regular.ttf", "IBMPlexMono-Bold.ttf", "IBMPlexMono-Italic.ttf",
                                 "IBMPlexMono-BoldItalic.ttf"]),
    "JetBrains Mono": dict(kind="mono", licence=_OFL, home="https://github.com/JetBrains/JetBrainsMono",
                           base="https://raw.githubusercontent.com/JetBrains/JetBrainsMono/v2.304/fonts/ttf/",
                           note="a tall, clear coding face for terminals and editors",
                           files=["JetBrainsMono-Regular.ttf", "JetBrainsMono-Bold.ttf", "JetBrainsMono-Italic.ttf",
                                  "JetBrainsMono-BoldItalic.ttf"]),
}

SHA256 = {  # of every file above; `python fonts.py --sums` prints this table from a fresh download
    "lmroman10-regular.otf": "1aa18cfefa58132c52ce5de70db1fd1154201c19cd2b2cdaffba4906a33e6852",
    "lmroman10-bold.otf": "102fe06c430a8b681b2bf6876b7cd967ae4d47b4b6b41d915eb7913b726d9fb1",
    "lmroman10-italic.otf": "c1fce25075567bb8dbf2151658c3b442690041db17a2d49fc9e55905ea5b7169",
    "lmroman10-bolditalic.otf": "c37a28eed7a6e03f792b98b5e5f637b2fcda378bb4855f99284f1a88fe35f124",
    "lmsans10-regular.otf": "d431b786b9b603662718e79cfe9b441f47a8b0b3e854dde89d5acb3ed7cfd682",
    "lmsans10-bold.otf": "a597b710326c1a8a2c7238d808e5d38711638a72a32383478db4829d63afd687",
    "lmsans10-oblique.otf": "1b38b36efc21f9d3f34edf6a13f9d71f59c41f824e790f5d3ee5ba5b4936f377",
    "lmsans10-boldoblique.otf": "e45673d84754a94f2f053bd4c15950b4b320144fbb622c630a26fd7b68cfabc5",
    "lmmono10-regular.otf": "22deb6d3be3ffcb40b33f0e010afb256c677d97c2a2bdee1f5740cdc97751558",
    "lmmono10-italic.otf": "14e6328a5383d2b927c47051093ce0aff60deea07e6df71b8b99ee8d9c2ff8f2",
    "STIXTwoText-Regular.ttf": "0566a1be411c0b2463fad8563b33f1c299442b8882c7fce9ef95e079f43233da",
    "STIXTwoText-Bold.ttf": "89adbb9907c40465e061f084598dbebd5a802e30c6e2d38612205b9c87281cae",
    "STIXTwoText-Italic.ttf": "5a0ed1c0c1d3d4f83f2d00bcd4a0a991a7b0755beacee0a57b26762327a84c22",
    "STIXTwoText-BoldItalic.ttf": "f11eced04bcd8919ad92ccc0251878be9a46fa45fa76c514e3abfbf605a3160f",
    "SourceSerif4-Regular.ttf": "e5a4ee6a3d87bb9024796be390c6771e2a0eb1883dae25effaf57ca01668e24b",
    "SourceSerif4-Bold.ttf": "7cf4f4e1ad74f45058d5bc61716b82560442fbdcd9d3654d2dea96bf6c683d86",
    "SourceSerif4-It.ttf": "9d2950a8f1da66e21502c35d646a1d2148e79f9ea43fd2158cf02f5232e7f430",
    "SourceSerif4-BoldIt.ttf": "7b215b37f8873f5579f3f8d2ded3ca7c588e2f435cd996605ebfc5befe2cd5eb",
    "SourceSans3-Regular.ttf": "4644c81b86ec9caaa76b634889968ed3c4f4f52f054855933acc7c2b21e53b0f",
    "SourceSans3-Bold.ttf": "9214b9d95e4231c609802815c2646c98174e2102d0d37f88978a7f8e71006e6a",
    "SourceSans3-It.ttf": "192afd78f0f54a3c69eaf02d43f4d9a821e9d6110e41d3d25d61a7385cd580e4",
    "SourceSans3-BoldIt.ttf": "7978291fc1bf314db887e0366853b33c5cf2e964c7b95cfb9ce403a6ec46a842",
    "SourceCodePro-Regular.ttf": "74bd80d3e42a08517cd7e1108ba3d86f2da29ac0f3065be95e0357956ab9db37",
    "SourceCodePro-Bold.ttf": "b2095e0d657e6d28dc32444a9dacabab0c9241d0bf39d96371756cc9bdbc3a5f",
    "SourceCodePro-It.ttf": "9c9e0f4d016210a3c5bdfba5262637c5b26ddff4ccc382ebbc781de5961d0042",
    "SourceCodePro-BoldIt.ttf": "1b49d9304012bf8db9e5dd4104183d5c122c445d0570a2259125f71977595b90",
    "FiraSans-Regular.ttf": "c29556a2719bf613ef3d5e070e40d903a8965d9c081beca1375dc1e6e0f93c23",
    "FiraSans-Bold.ttf": "a4d8e149ecdd4874a0726eb0af894488b3b31c423d6b0017c8f415ed1b795b45",
    "FiraSans-Italic.ttf": "a4204f9a8029c1c55daa5385ac95936f74abea3e232770f1fb2862a19d1dfa13",
    "FiraSans-BoldItalic.ttf": "1b96c16db9819bcaa9d588068bb397bb6e977c4874965d3d50d789a1943757eb",
    "FiraMono-Regular.ttf": "2e00b0cf3106a3d792f35711b7722740b76c02ee7300ea8c60f940584b7a8ac9",
    "FiraMono-Bold.ttf": "61f0ca3ae72a83deb807040f314e1b8b4e40b08213df243db9fd430095cab305",
    "IBMPlexSans-Regular.ttf": "975dcda37d80f038dcd143c22e33ca2d97a0cc5a929aace1c749153b0fe1afa5",
    "IBMPlexSans-Bold.ttf": "9e6c74a889a700d707613d24548fe4ffa6bc59559a0689d2cf9e133bdcdafb2f",
    "IBMPlexSans-Italic.ttf": "a9c6ef9942c49e49d11e11a6dacc0b3a087978757e9b22a06b8ac22a6400fb15",
    "IBMPlexSans-BoldItalic.ttf": "0e3142ba2ef31fe5c02f0c6c36424f251609cd6b73880076e21c2e81931ba2b9",
    "IBMPlexMono-Regular.ttf": "fe11304a5fe956d5744e9b6a246cc83d90425245e75a62230044966ca96a7f50",
    "IBMPlexMono-Bold.ttf": "ca403c56931baef307d20ba64b69acb71abcad61f75e66414661d57484b690ec",
    "IBMPlexMono-Italic.ttf": "8ebe04c8c6cc82f0be19896ddc61d9935cdd0f027b0173c1945b8d247d7dfc2a",
    "IBMPlexMono-BoldItalic.ttf": "0e45a5a540992163229d2a29662553f313fab391757ca2ab3dc8f4e0d9be0979",
    "JetBrainsMono-Regular.ttf": "a0bf60ef0f83c5ed4d7a75d45838548b1f6873372dfac88f71804491898d138f",
    "JetBrainsMono-Bold.ttf": "5590990c82e097397517f275f430af4546e1c45cff408bde4255dad142479dcb",
    "JetBrainsMono-Italic.ttf": "9d0a1f7a708e6af183f1193b7e81d40da294f5c67682c085d8401c60aac8ded4",
    "JetBrainsMono-BoldItalic.ttf": "4039d5ce0ed225bf9c8b2c8c6436290ae2f356b7e90d70fa666227238324aa3b",
}


# ---- which fonts does this machine have? -------------------------------------------------------
_STYLES = {"regular", "bold", "italic", "oblique", "light", "medium", "semibold", "demibold", "demi", "black",
           "heavy", "thin", "book", "extralight", "semilight", "extrabold", "condensed", "cond", "it", "roman"}
_installed_cache = None


def _system_families():
    """Lower-case names of the installed fonts, or None when this system cannot be asked."""
    global _installed_cache
    if _installed_cache is not None:
        return _installed_cache or None
    names = set()
    if sys.platform == "win32":
        names |= _gdi_families()      # what programs see right now: installed fonts and those of --load
        import winreg
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
                    for i in range(winreg.QueryInfoKey(key)[1]):
                        value = winreg.EnumValue(key, i)[0]
                        value = value.rsplit(" (", 1)[0]          # "Cambria & Cambria Math (TrueType)"
                        names.update(part.strip().lower() for part in value.split(" & "))
            except OSError:
                pass
    else:
        exe = shutil.which("fc-list")
        if exe:
            try:
                out = subprocess.run([exe, ":", "family"], capture_output=True, text=True, timeout=20).stdout
                for line in out.splitlines():
                    names.update(part.strip().lower() for part in line.split(","))
            except (OSError, subprocess.SubprocessError):
                pass
    _installed_cache = names
    return names or None


def _gdi_families() -> set:
    """Family names Windows offers to programs at this moment (lower case); empty if it cannot be asked."""
    try:
        import ctypes
        from ctypes import wintypes

        class LOGFONTW(ctypes.Structure):
            _fields_ = [("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG), ("lfEscapement", wintypes.LONG),
                        ("lfOrientation", wintypes.LONG), ("lfWeight", wintypes.LONG), ("lfItalic", wintypes.BYTE),
                        ("lfUnderline", wintypes.BYTE), ("lfStrikeOut", wintypes.BYTE), ("lfCharSet", wintypes.BYTE),
                        ("lfOutPrecision", wintypes.BYTE), ("lfClipPrecision", wintypes.BYTE),
                        ("lfQuality", wintypes.BYTE), ("lfPitchAndFamily", wintypes.BYTE),
                        ("lfFaceName", wintypes.WCHAR * 32)]

        proc = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.POINTER(LOGFONTW), ctypes.c_void_p, wintypes.DWORD,
                                  wintypes.LPARAM)
        names = set()

        def each(lf, _metrics, _kind, _param):
            names.add(lf.contents.lfFaceName.lstrip("@").lower())
            return 1

        user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
        user32.GetDC.restype, user32.GetDC.argtypes = wintypes.HDC, [wintypes.HWND]
        user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
        gdi32.EnumFontFamiliesExW.argtypes = [wintypes.HDC, ctypes.POINTER(LOGFONTW), proc, wintypes.LPARAM,
                                              wintypes.DWORD]
        hdc = user32.GetDC(None)
        lf = LOGFONTW()
        lf.lfCharSet = 1  # DEFAULT_CHARSET: every font
        gdi32.EnumFontFamiliesExW(hdc, ctypes.byref(lf), proc(each), 0, 0)
        user32.ReleaseDC(None, hdc)
        return names
    except Exception:  # noqa: BLE001 - fall back to the registry
        return set()


def installed(family: str):
    """True / False, or None when the system's fonts cannot be listed (then nothing is claimed).
    The registry lists faces ("Palatino Linotype Bold"), so a face that starts with the family counts."""
    names = _system_families()
    if names is None:
        return None
    fam = str(family).strip().lower()
    return any(n == fam or (n.startswith(fam + " ") and set(n[len(fam):].split()) <= _STYLES) for n in names)


def missing(families) -> list:
    """The families (unique, in order) that are definitely not installed."""
    out = []
    for f in families:
        if f and f not in out and installed(f) is False:
            out.append(f)
    return out


def hint(family: str) -> str:
    """What to do about a font that is not installed."""
    if family in FAMILIES:
        return f'install it with: python fonts.py --install "{family}"'
    return "it comes with Microsoft Office; install it, or pick a template whose fonts you have"


# ---- download and install ----------------------------------------------------------------------
def download(family: str, say=print) -> list:
    """The font files of an open family in fonts/ (fetched when missing, checked against SHA256)."""
    spec = FAMILIES[family]
    os.makedirs(CACHE, exist_ok=True)
    paths = []
    for name in spec["files"]:
        path = os.path.join(CACHE, name)
        if not os.path.exists(path):
            say(f"  downloading {name}")
            req = urllib.request.Request(spec["base"] + name, headers={"User-Agent": "ppt-templated"})
            try:
                with urllib.request.urlopen(req, timeout=60) as r:
                    data = r.read()
            except OSError as exc:
                raise SystemExit(f"error: cannot download {spec['base'] + name}: {exc}\n"
                                 f"get the font from {spec['home']} and install it by hand") from None
            want = SHA256.get(name)
            if want and hashlib.sha256(data).hexdigest() != want:
                raise SystemExit(f"error: {name} from {spec['base']} does not match its recorded checksum - "
                                 f"not installed. Get the font from {spec['home']} instead.")
            with open(path, "wb") as fh:
                fh.write(data)
        paths.append(path)
    return paths


def load(family: str, say=print) -> None:
    """Windows: make an open family usable by every program until you sign out. Nothing is installed:
    no file is copied and no setting changed. Good for building and previewing on a machine you do
    not want to change."""
    global _installed_cache
    if sys.platform != "win32":
        raise SystemExit("error: --load works on Windows only; use --install (it installs for your user only)")
    import ctypes
    for path in download(family, say):
        if not ctypes.windll.gdi32.AddFontResourceW(path):
            raise SystemExit(f"error: Windows could not load {path}")
    ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001D, 0, 0, 0x0002, 1000, None)  # WM_FONTCHANGE to all
    _installed_cache = None
    say(f"  loaded {family} (until you sign out)")


def user_font_dir() -> str:
    if sys.platform == "win32":
        return os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "Microsoft", "Windows", "Fonts")
    if sys.platform == "darwin":
        return os.path.expanduser("~/Library/Fonts")
    return os.path.expanduser("~/.local/share/fonts")


def _face_name(path: str) -> str:
    """Name of a font file as Windows lists it ("Fira Sans Bold", "LM Roman 10 Regular")."""
    try:
        from fontTools.ttLib import TTFont
        names = TTFont(path, lazy=True)["name"]
        family, style = names.getName(1, 3, 1, 0x409), names.getName(2, 3, 1, 0x409)
        return f"{family.toUnicode()} {style.toUnicode()}"
    except Exception:  # noqa: BLE001 - fontTools is optional here
        stem = os.path.splitext(os.path.basename(path))[0]
        family = next((f for f, spec in FAMILIES.items() if os.path.basename(path) in spec["files"]), stem)
        return f"{family} {stem.rsplit('-', 1)[-1]}"


def install(family: str, say=print) -> None:
    """Install an open family for the current user (no administrator rights needed)."""
    global _installed_cache
    paths = download(family, say)
    target = user_font_dir()
    os.makedirs(target, exist_ok=True)
    for path in paths:
        dest = os.path.join(target, os.path.basename(path))
        if not os.path.exists(dest):
            shutil.copyfile(path, dest)
        if sys.platform == "win32":
            import ctypes
            import winreg
            kind = "OpenType" if dest.lower().endswith(".otf") else "TrueType"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                                  r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
                winreg.SetValueEx(key, f"{_face_name(dest)} ({kind})", 0, winreg.REG_SZ, dest)
            ctypes.windll.gdi32.AddFontResourceW(dest)   # usable now, without signing out
    if sys.platform not in ("win32", "darwin") and shutil.which("fc-cache"):
        subprocess.run(["fc-cache", "-f", target], capture_output=True)
    _installed_cache = None
    say(f"  installed {family} in {target}")


def template_fonts() -> dict:
    """{template name: [fonts it uses]} for the bundled templates."""
    import ppt_templated
    out = {}
    for name, path in ppt_templated.template_names().items():
        try:
            out[name] = ppt_templated.fonts_of(path)
        except Exception:  # noqa: BLE001 - a broken template should not hide the others
            out[name] = []
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--install", nargs="*", metavar="NAME",
                    help="install open fonts: all, or those of the named templates / families")
    ap.add_argument("--load", nargs="*", metavar="NAME",
                    help="Windows: make open fonts usable until you sign out, without installing anything")
    ap.add_argument("--download", action="store_true", help="fetch the open fonts into fonts/ without installing")
    ap.add_argument("--sums", action="store_true", help=argparse.SUPPRESS)
    a = ap.parse_args(argv)
    if a.sums:  # maintenance: the SHA256 table of the files now in fonts/
        for fam in FAMILIES:
            for p in download(fam):
                with open(p, "rb") as fh:
                    print(f'    "{os.path.basename(p)}": "{hashlib.sha256(fh.read()).hexdigest()}",')
        return
    if a.download:
        for fam in FAMILIES:
            download(fam)
        print(f"open fonts are in {CACHE}")
        return
    if a.install is not None or a.load is not None:
        by_template = template_fonts()
        lower = {f.lower(): f for f in FAMILIES}
        wanted = []
        for name in (a.install if a.install is not None else a.load) or list(FAMILIES):
            if name.lower() in lower:
                wanted.append(lower[name.lower()])
            elif name in by_template:
                wanted += [f for f in by_template[name] if f in FAMILIES]
            else:
                sys.exit(f"error: '{name}' is neither an open font family ({', '.join(FAMILIES)}) "
                         f"nor a bundled template")
        for fam in dict.fromkeys(wanted):
            if installed(fam):
                print(f"  {fam}: already there")
            elif a.install is not None:
                install(fam)
            else:
                load(fam)
        print("done - restart PowerPoint if it is open")
        return
    state = {True: "installed", False: "MISSING", None: "unknown"}
    print("open fonts (python fonts.py --install):")
    for fam, spec in FAMILIES.items():
        print(f"  {fam:<16} {state[installed(fam)]:<10} {spec['kind']:<6} {spec['note']}")
    print("\ntemplates:")
    for name, fams in template_fonts().items():
        gone = missing(fams)
        print(f"  {name:<11} {', '.join(fams):<46} " + ("ok" if not gone else "missing: " + ", ".join(gone)))


if __name__ == "__main__":
    main()
