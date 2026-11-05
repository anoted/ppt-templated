"""
metrics.py - how many lines a text takes in a box of a given width, from real glyph widths.

Every box ppt-templated draws (cards, timeline labels, table rows, titles, ...) gets its height from
text_lines(): it wraps the text word by word, like PowerPoint does, with the advance widths of the
template's font. The widths are stored below, so a deck comes out the same on every machine.

They come from fonts with the same metrics as the Office fonts the templates use (Carlito = Calibri,
Caladea = Cambria, Liberation Sans = Arial, Liberation Serif = Times New Roman). Fonts without such a
twin (Georgia, Constantia, Segoe UI, ...) use the closest one, scaled to their average width.
To regenerate the table (needs fontTools and those four font families): python metrics.py, then
paste its output over WIDTHS.
"""

from __future__ import annotations

import math
import re

CHARSET = (' !"#$%&\'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~'
           '\xa0¡¢£¤¥¦§¨©ª«¬\xad®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ'
           '–—‘’‚“”„•…‰‹›€™−≤≥≈≠→←↑↓↔⇒∞√∑∫∂∇∈αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ✓✗★')
WIDTHS = {  # 1/1000 em per CHARSET character, regular then bold; 0 = not in the font
    "carlito": (
        '226 326 401 498 507 715 682 221 303 303 498 498 250 306 252 386 507 507 507 507 507 507 507 507 '
        '507 507 268 268 498 498 498 463 894 579 544 533 615 488 459 631 623 252 319 520 420 855 646 662 '
        '517 673 543 459 487 642 567 890 519 487 468 307 386 307 498 498 291 479 525 423 525 498 305 471 '
        '525 229 239 455 229 799 525 527 525 525 349 391 335 525 452 715 433 453 395 314 460 314 498 226 '
        '326 498 507 498 507 498 498 393 834 402 512 498 0 507 394 339 498 336 334 292 550 586 252 307 '
        '246 422 512 636 671 675 463 579 579 579 579 579 579 763 533 488 488 488 488 252 252 252 252 625 '
        '646 662 662 662 662 662 498 664 642 642 642 642 487 517 527 479 479 479 479 479 479 773 423 498 '
        '498 498 498 229 229 229 229 525 525 527 527 527 527 527 498 529 525 525 525 525 453 525 453 498 '
        '905 250 250 250 418 418 418 498 690 1038 339 339 507 705 498 498 498 498 498 905 905 905 905 '
        '1288 0 853 498 541 366 533 0 0 567 531 446 523 456 348 537 532 274 455 463 550 449 376 553 509 '
        '532 387 542 651 426 708 696 416 564 662 573 492 623 458 759 750 664 0 0 0 ',
        '226 326 438 498 507 729 705 233 312 312 498 498 258 306 267 430 507 507 507 507 507 507 507 507 '
        '507 507 276 276 498 498 498 463 898 606 561 529 630 488 459 637 631 267 331 547 423 874 659 676 '
        '532 686 563 473 495 653 591 906 551 520 478 325 430 325 498 498 300 494 537 418 537 503 316 474 '
        '537 246 255 480 246 813 537 538 537 537 355 399 347 537 473 745 459 474 397 344 475 344 498 226 '
        '326 498 507 498 507 498 498 415 834 416 539 498 0 507 390 342 498 338 336 301 563 598 268 303 '
        '252 435 539 658 691 702 463 606 606 606 606 606 606 775 529 488 488 488 488 267 267 267 267 639 '
        '659 676 676 676 676 676 498 681 653 653 653 653 520 532 555 494 494 494 494 494 494 775 418 503 '
        '503 503 503 246 246 246 246 537 537 538 538 538 538 538 498 544 537 537 537 537 474 537 474 498 '
        '905 258 258 258 435 435 435 498 711 1062 344 344 507 720 498 498 498 498 498 905 905 905 905 '
        '1310 0 838 498 550 375 537 0 0 591 545 467 533 458 364 544 545 286 480 486 563 471 385 579 530 '
        '544 393 548 699 451 738 718 410 578 676 594 492 631 470 801 808 681 0 0 0 ',
    ),
    "caladea": (
        '220 286 393 619 506 890 687 237 382 382 427 554 205 329 206 490 517 362 472 461 492 475 470 470 '
        '485 478 264 264 554 554 554 422 885 599 579 519 625 540 509 539 696 343 304 630 501 888 693 599 '
        '533 599 613 453 557 687 598 889 571 557 507 350 490 350 554 371 151 470 520 421 520 441 313 459 '
        '554 277 255 525 271 813 558 480 520 501 396 392 325 546 470 720 483 482 428 387 316 387 712 220 '
        '286 441 528 544 609 316 500 261 851 417 491 554 332 851 278 375 554 407 407 151 544 588 106 285 '
        '407 428 491 829 867 866 422 599 599 599 599 599 599 866 519 540 540 540 540 343 343 343 343 662 '
        '693 599 599 599 599 599 554 653 687 687 687 687 557 574 605 470 470 470 470 470 470 708 421 441 '
        '441 441 441 277 277 277 277 530 558 480 480 480 480 480 554 490 546 546 546 546 482 547 482 500 '
        '1000 221 221 205 375 375 360 443 752 1263 304 304 628 679 554 554 554 554 554 783 783 294 294 '
        '883 0 851 657 708 425 545 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 '
        '0 0 0 ',
        '220 335 422 618 543 976 740 251 408 408 453 592 232 311 230 505 533 458 507 508 553 520 538 494 '
        '556 534 280 280 592 592 592 452 921 653 634 551 664 559 534 586 719 358 340 680 541 845 686 647 '
        '593 647 640 473 587 695 632 959 616 587 554 368 505 368 592 371 184 525 568 450 570 513 335 508 '
        '600 308 297 589 311 891 604 530 573 559 454 444 365 609 519 779 525 521 464 393 320 393 592 220 '
        '335 469 556 587 604 320 533 294 851 420 531 592 337 851 311 378 592 437 437 183 605 588 140 285 '
        '437 436 531 897 904 970 452 653 653 653 653 653 653 880 551 559 559 559 559 358 358 358 358 709 '
        '686 647 647 647 647 647 592 695 695 695 695 695 587 617 677 525 525 525 525 525 525 794 450 513 '
        '513 513 513 308 308 308 311 572 604 530 530 530 530 530 592 569 609 609 609 609 521 591 521 500 '
        '1000 235 235 232 398 398 395 443 772 1380 306 306 640 684 592 592 592 592 592 783 783 294 294 '
        '883 0 869 766 731 425 591 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 '
        '0 0 0 ',
    ),
    "sans": (
        '278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 '
        '556 556 278 278 584 584 584 556 1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 '
        '778 667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 333 556 556 500 556 556 278 '
        '556 556 222 222 500 222 833 556 556 556 556 333 500 278 556 500 722 500 500 500 334 260 334 584 '
        '278 333 556 556 556 556 260 556 333 737 370 556 584 333 737 552 400 549 333 333 333 576 537 333 '
        '333 333 365 556 834 834 834 611 667 667 667 667 667 667 1000 722 667 667 667 667 278 278 278 '
        '278 722 722 778 778 778 778 778 584 778 722 722 722 722 667 667 611 556 556 556 556 556 556 889 '
        '500 556 556 556 556 278 278 278 278 556 556 556 556 556 556 556 549 611 556 556 556 556 500 556 '
        '500 556 1000 222 222 222 333 333 333 350 1000 1000 333 333 556 1000 584 549 549 549 549 1000 '
        '1000 500 500 1000 0 713 549 713 274 494 0 0 578 575 500 557 446 441 556 556 222 500 500 576 500 '
        '448 690 569 617 395 547 648 525 713 781 551 668 778 668 650 722 618 798 835 748 0 0 0 ',
        '278 333 474 556 556 889 722 238 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 '
        '556 556 333 333 584 584 584 611 975 722 722 722 722 667 611 778 722 278 556 722 611 833 722 778 '
        '667 778 722 667 611 722 667 944 667 667 611 333 278 333 584 556 333 556 611 556 611 556 333 611 '
        '611 278 278 556 278 889 611 611 611 611 389 556 333 611 556 778 556 556 500 389 280 389 584 278 '
        '333 556 556 556 556 280 556 333 737 370 556 584 333 737 552 400 549 333 333 333 576 556 333 333 '
        '333 365 556 834 834 834 611 722 722 722 722 722 722 1000 722 667 667 667 667 278 278 278 278 '
        '722 722 778 778 778 778 778 584 778 722 722 722 722 667 667 611 556 556 556 556 556 556 889 556 '
        '556 556 556 556 278 278 278 278 611 611 611 611 611 611 611 549 611 611 611 611 611 556 611 556 '
        '556 1000 278 278 278 500 500 500 350 1000 1000 333 333 556 1000 584 549 549 549 549 1000 1000 '
        '500 500 1000 0 713 549 713 274 494 0 0 615 610 556 606 475 460 611 541 278 558 556 612 556 445 '
        '766 619 684 446 582 715 576 753 845 601 719 778 667 644 722 600 821 809 802 0 0 0 ',
    ),
    "serif": (
        '250 333 408 500 500 833 778 180 333 333 500 564 250 333 250 278 500 500 500 500 500 500 500 500 '
        '500 500 278 278 564 564 564 444 921 722 667 667 722 611 556 722 722 333 389 722 611 889 722 722 '
        '556 722 667 556 611 722 722 944 722 722 611 333 278 333 469 500 333 444 500 444 500 444 333 500 '
        '500 278 278 500 278 778 500 500 500 500 333 389 278 500 500 722 500 500 444 480 200 480 541 250 '
        '333 500 500 500 500 200 500 333 760 276 500 564 333 760 500 400 549 300 300 333 576 453 333 333 '
        '300 310 500 750 750 750 444 722 722 722 722 722 722 889 667 611 611 611 611 333 333 333 333 722 '
        '722 722 722 722 722 722 564 722 722 722 722 722 722 556 500 444 444 444 444 444 444 667 444 444 '
        '444 444 444 278 278 278 278 500 500 500 500 500 500 500 549 500 500 500 500 500 500 500 500 500 '
        '1000 333 333 333 444 444 444 350 1000 1000 333 333 500 980 564 549 549 549 549 1000 1000 500 '
        '500 1000 0 713 549 713 274 494 0 0 524 509 442 471 420 414 523 479 269 504 485 536 452 446 505 '
        '499 539 402 495 577 444 626 658 578 643 722 725 643 722 582 731 738 743 0 0 0 ',
        '250 333 555 500 500 1000 833 278 333 333 500 570 250 333 250 278 500 500 500 500 500 500 500 '
        '500 500 500 333 333 570 570 570 500 930 722 667 722 722 667 611 778 778 389 500 778 667 944 722 '
        '778 611 778 722 556 667 722 722 1000 722 722 667 333 278 333 581 500 333 500 556 444 556 444 '
        '333 500 556 278 333 556 278 833 556 500 556 556 444 389 333 556 500 722 500 500 444 394 220 394 '
        '520 250 333 500 500 500 500 220 500 333 747 300 500 570 333 747 500 400 549 300 300 333 576 540 '
        '333 333 300 330 500 750 750 750 500 722 722 722 722 722 722 1000 722 667 667 667 667 389 389 '
        '389 389 722 722 778 778 778 778 778 570 778 722 722 722 722 722 611 556 500 500 500 500 500 500 '
        '722 444 444 444 444 444 278 278 278 278 500 556 500 500 500 500 500 549 500 556 556 556 556 500 '
        '556 500 500 1000 333 333 333 500 500 500 350 1000 1000 333 333 500 1000 570 549 549 549 549 '
        '1000 1000 500 500 1000 0 713 549 713 274 494 0 0 558 527 468 519 427 413 568 522 310 555 491 '
        '567 447 446 548 535 544 461 519 624 462 688 731 636 627 778 717 677 778 654 829 780 801 0 0 0 ',
    ),
}

LINE = {"carlito": 1.22}  # line height (em) of single spacing; others 1.2
# template font -> (stored table, width scale)
# Fonts without a twin are mapped to the closest stored one, scaled to their average width and rounded up.
FONTS = {"calibri": ("carlito", 1.0), "calibri light": ("carlito", 1.0), "carlito": ("carlito", 1.0),
         "segoe ui": ("sans", 1.0), "aptos": ("sans", 1.0),
         "cambria": ("caladea", 1.0), "cambria math": ("caladea", 1.0), "caladea": ("caladea", 1.0),
         "constantia": ("caladea", 1.05), "georgia": ("serif", 1.12),
         "arial": ("sans", 1.0), "helvetica": ("sans", 1.0), "liberation sans": ("sans", 1.0),
         "times new roman": ("serif", 1.0), "times": ("serif", 1.0), "liberation serif": ("serif", 1.0)}
UNKNOWN = ("sans", 1.04)  # a font we know nothing about: a little wider than Arial, so text still fits
MONO = 0.6  # Courier New, for `code`
SAMPLE = "The quick brown fox jumps over the lazy dog; results improved by 12% in 2025."
MARGIN = 0.98  # wrap a little early: renderers differ slightly in kerning and hinting


class Font:
    """Advance widths (em) of one template font, regular and bold."""

    def __init__(self, name: str):
        self.name = str(name)
        key, self.scale = FONTS.get(self.name.lower(), UNKNOWN)
        self.line = LINE.get(key, 1.2)
        regular, bold = (part.split() for part in WIDTHS[key])
        self._w = {}
        for weight, values in ((False, regular), (True, bold)):
            table = {c: int(v) / 1000 * self.scale for c, v in zip(CHARSET, values) if int(v)}
            letters = [table[c] for c in "abcdefghijklmnopqrstuvwxyz"]
            table[""] = sum(letters) / len(letters)  # characters the table does not have
            self._w[weight] = table
        self.avg = self.width(SAMPLE) / len(SAMPLE)  # average character width of running text (em)

    def char(self, c: str, bold: bool = False) -> float:
        t = self._w[bold]
        w = t.get(c)
        if w is not None:
            return w
        if c in "\t":
            return t[" "] * 4
        o = ord(c)
        if 0x2E80 <= o <= 0xFFEF:  # CJK and full-width forms
            return 1.0 * self.scale
        if 0x0300 <= o <= 0x036F or c in "​‌‍⁣﻿\xad":  # combining / invisible
            return 0.0
        return t[""]

    def width(self, text: str, bold: bool = False) -> float:
        """Width of plain text in em."""
        return sum(self.char(c, bold) for c in str(text))

    def __repr__(self):
        return f"Font({self.name!r})"


_CACHE: dict = {}


def font(name) -> Font:
    """The Font for a template font name (cached); unknown names get a safe, slightly wide stand-in."""
    if isinstance(name, Font):
        return name
    key = str(name or "").lower()
    if key not in _CACHE:
        _CACHE[key] = Font(name or "Calibri")
    return _CACHE[key]


# inline markup of the YAML text: **bold**, `code`, [link](url), ==alert==, $math$, *italic*
_TOKEN = re.compile(r"\*\*(?P<b>.+?)\*\*|`(?P<code>[^`]+)`|\[(?P<lt>[^\]]+)\]\([^)\s]+\)|==(?P<alert>.+?)=="
                    r"|(?<![\\$])\$(?P<math>[^\s$](?:[^$]*[^\s$\\])?)\$(?!\d)")
_ITALIC = re.compile(r"(?<![\w*])\*([^*\s][^*]*?)\*(?![\w*])")


def _plain(text):
    return _ITALIC.sub(r"\1", text)


def _glyphs(text: str, f: Font, bold: bool) -> list:
    """[(character, width in em)] of one line of YAML text, markup removed."""
    out, pos = [], 0

    def add(s, b=bold, mono=False):
        for c in s:
            out.append((c, MONO if mono and c != " " else f.char(c, b)))

    for m in _TOKEN.finditer(text):
        add(_plain(text[pos:m.start()]))
        g = m.groupdict()
        if g["b"] is not None:
            for c, w in _glyphs(g["b"], f, True):
                out.append((c, w))
        elif g["code"] is not None:
            add(g["code"], mono=True)
        elif g["lt"] is not None:
            add(_plain(g["lt"]))
        elif g["alert"] is not None:
            out.extend(_glyphs(g["alert"], f, bold))
        else:  # math: about half its source length, set in a wider math font
            add("x" * max(1, round(len(g["math"]) * 0.55)))
        pos = m.end()
    add(_plain(text[pos:]))
    return out


BREAK_AFTER = "-‐–—/"  # a line may also end after these, as in PowerPoint


def line_count(glyphs, limit: float) -> int:
    """Lines needed by [(char, width)] when a line holds `limit` em: greedy word wrap; spaces at the end
    of a line take no room; a word longer than a line is broken anywhere."""
    words, word, space = [], 0.0, 0.0
    for c, w in glyphs:
        if c == " " or c == "\t":
            space += w
            continue
        if space:
            words.append((word, space))
            word, space = 0.0, 0.0
        word += w
        if c in BREAK_AFTER:
            words.append((word, 0.0))
            word = 0.0
    words.append((word, space))
    lines, x = 1, 0.0
    for wd, sp in words:
        if x > 0 and x + wd > limit:
            lines, x = lines + 1, 0.0
        if x == 0 and wd > limit:  # one long word: broken across lines
            extra = math.ceil(wd / limit) - 1
            lines += extra
            wd -= extra * limit
        x += wd + sp
    return lines


def text_lines(text, width_in: float, size_pt: float, font_name="Calibri", bold: bool = False) -> int:
    """Lines `text` takes in a box `width_in` inches wide at `size_pt` (each newline starts a paragraph)."""
    f = font(font_name)
    limit = max(width_in, 0.05) * 72.0 / max(size_pt, 1.0) * MARGIN
    return sum(line_count(_glyphs(seg, f, bold), limit) for seg in str(text).split("\n"))


def text_width(text, size_pt: float, font_name="Calibri", bold: bool = False) -> float:
    """Width in inches of the longest line of `text` set on one line each."""
    f = font(font_name)
    return max((sum(w for _, w in _glyphs(seg, f, bold)) for seg in str(text).split("\n")), default=0.0) \
        * size_pt / 72.0


def _generate():  # pragma: no cover - maintenance only
    from fontTools.ttLib import TTFont
    root = "/usr/share/fonts/truetype/"
    files = {"carlito": "crosextra/Carlito", "caladea": "crosextra/Caladea",
             "sans": "liberation/LiberationSans", "serif": "liberation/LiberationSerif"}
    for key, path in files.items():
        rows = []
        for style in ("Regular", "Bold"):
            t = TTFont(f"{root}{path}-{style}.ttf")
            cmap, hmtx, upm = t.getBestCmap(), t["hmtx"], t["head"].unitsPerEm
            rows.append(" ".join(str(round(hmtx[cmap[ord(c)]][0] * 1000 / upm)) if ord(c) in cmap else "0"
                                 for c in CHARSET))
        print(f'"{key}": {tuple(rows)!r},')


if __name__ == "__main__":  # pragma: no cover
    _generate()
