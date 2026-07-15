"""
Local Windows font resolution for the PDF editor's export.

Two jobs:
  1. Map a (family, bold, italic) style to an installed Windows TrueType file.
  2. Produce an embeddable font subset with the U+00A0 (no-break space) cmap
     entry removed, so that text written by PyMuPDF extracts/copies as normal
     U+0020 spaces rather than nbsp. (PyMuPDF reverse-maps the space glyph via
     the font cmap; many Windows fonts map both 0x20 and 0xA0 to the same
     glyph and 0xA0 wins, corrupting extracted text. Dropping 0xA0 fixes it.)

The editor UI itself just uses the installed fonts by family name (no files
served) - the cmap fix changes only the space codepoint mapping, never metrics
or shapes, so on-screen and exported text stay visually identical.
"""

from __future__ import annotations

import io
import os
import threading

import fitz
from fontTools.ttLib import TTFont, TTLibError

FONTS_DIR = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "Fonts")

# family -> {(bold, italic): filename}. Regular is the fallback for any missing
# style. Cambria's regular ships as a .ttc collection (index handled below).
_FAMILY_FILES = {
    "Arial": {(0, 0): "arial.ttf", (1, 0): "arialbd.ttf", (0, 1): "ariali.ttf", (1, 1): "arialbi.ttf"},
    "Calibri": {(0, 0): "calibri.ttf", (1, 0): "calibrib.ttf", (0, 1): "calibrii.ttf", (1, 1): "calibriz.ttf"},
    "Georgia": {(0, 0): "georgia.ttf", (1, 0): "georgiab.ttf", (0, 1): "georgiai.ttf", (1, 1): "georgiaz.ttf"},
    "Times New Roman": {(0, 0): "times.ttf", (1, 0): "timesbd.ttf", (0, 1): "timesi.ttf", (1, 1): "timesbi.ttf"},
    "Cambria": {(0, 0): "cambria.ttc", (1, 0): "cambriab.ttf", (0, 1): "cambriai.ttf", (1, 1): "cambriaz.ttf"},
    "Verdana": {(0, 0): "verdana.ttf", (1, 0): "verdanab.ttf", (0, 1): "verdanai.ttf", (1, 1): "verdanaz.ttf"},
    "Tahoma": {(0, 0): "tahoma.ttf", (1, 0): "tahomabd.ttf"},
    "Trebuchet MS": {(0, 0): "trebuc.ttf", (1, 0): "trebucbd.ttf", (0, 1): "trebucit.ttf", (1, 1): "trebucbi.ttf"},
    "Segoe UI": {(0, 0): "segoeui.ttf", (1, 0): "segoeuib.ttf", (0, 1): "segoeuii.ttf", (1, 1): "segoeuiz.ttf"},
    "Consolas": {(0, 0): "consola.ttf", (1, 0): "consolab.ttf", (0, 1): "consolai.ttf", (1, 1): "consolaz.ttf"},
    "Courier New": {(0, 0): "cour.ttf", (1, 0): "courbd.ttf", (0, 1): "couri.ttf", (1, 1): "courbi.ttf"},
    "Segoe UI Symbol": {(0, 0): "seguisym.ttf"},
}

_DEFAULT_FAMILY = "Arial"

_cache_lock = threading.Lock()
_bytes_cache: dict = {}   # (family, bold, italic) -> fixed font bytes


def known_family(family: str) -> str:
    return family if family in _FAMILY_FILES else _DEFAULT_FAMILY


def resolve_file(family: str, bold: bool, italic: bool):
    """Return (path, ttc_index) for the closest available style, or (None, 0)."""
    fam = known_family(family)
    styles = _FAMILY_FILES[fam]
    # try exact, then drop italic, then drop bold, then regular
    for key in ((int(bold), int(italic)), (int(bold), 0), (0, int(italic)), (0, 0)):
        fn = styles.get(key)
        if fn:
            path = os.path.join(FONTS_DIR, fn)
            if os.path.exists(path):
                return path, 0
    return None, 0


def _prefer_ascii(tt: TTFont) -> bytes:
    """PyMuPDF reverse-maps each drawn glyph to a codepoint via the font cmap.
    Many Windows fonts map several codepoints to one glyph (e.g. U+0020 & U+00A0
    -> space glyph; U+002D & U+00AD -> hyphen glyph), and the non-ASCII one can
    win, corrupting extracted/copied text (nbsp for space, soft-hyphen for
    hyphen). For every glyph reachable from an ASCII codepoint (0x20-0x7E),
    drop any non-ASCII codepoint that points at the same glyph, so ASCII always
    wins on extraction. Visual shaping is unaffected."""
    for tbl in tt.get("cmap").tables:
        try:
            cmap = tbl.cmap
        except Exception:
            continue
        ascii_glyphs = {g for cp, g in cmap.items() if 0x20 <= cp <= 0x7E}
        for cp in [c for c in list(cmap.keys())
                   if c > 0x7E and cmap[c] in ascii_glyphs]:
            del cmap[cp]
    buf = io.BytesIO()
    tt.save(buf)
    return buf.getvalue()


def fixed_font_bytes(family: str, bold: bool, italic: bool):
    """cmap-fixed embeddable bytes for the closest available style, cached."""
    key = (known_family(family), bool(bold), bool(italic))
    with _cache_lock:
        if key in _bytes_cache:
            return _bytes_cache[key]
    path, idx = resolve_file(family, bold, italic)
    if path is None:
        return None
    try:
        if path.lower().endswith(".ttc"):
            tt = TTFont(path, fontNumber=idx)
        else:
            tt = TTFont(path)
        data = _prefer_ascii(tt)
    except (TTLibError, Exception):
        # last resort: hand the raw file to fitz (extraction may show nbsp,
        # but the glyphs are correct - better than a missing font)
        try:
            with open(path, "rb") as f:
                data = f.read()
        except Exception:
            return None
    with _cache_lock:
        _bytes_cache[key] = data
    return data


def _css_family(family: str) -> str:
    return known_family(family)


def build_export_fonts(families):
    """
    Given an iterable of mapped family names used on a page, return
    (css_text, fitz.Archive) suitable for page.insert_htmlbox: for each family,
    up to four @font-face rules (regular/bold/italic/bold-italic) pointing at
    cmap-fixed font files inside the archive. HTML sets font-family to the
    family name; <b>/<i> pick the right face.
    """
    arch = fitz.Archive()
    seen_files = set()
    css_parts = []
    for fam in sorted(set(known_family(f) for f in families)):
        for bold in (False, True):
            for italic in (False, True):
                data = fixed_font_bytes(fam, bold, italic)
                if data is None:
                    continue
                arcname = f"{fam}-{int(bold)}{int(italic)}.ttf".replace(" ", "_")
                if arcname not in seen_files:
                    arch.add(data, arcname)
                    seen_files.add(arcname)
                weight = "bold" if bold else "normal"
                style = "italic" if italic else "normal"
                css_parts.append(
                    f'@font-face{{font-family:"{fam}";font-weight:{weight};'
                    f'font-style:{style};src:url({arcname});}}')
    return "\n".join(css_parts), arch
