"""
Shared PDF extraction layer.

This is the ONE common layer every conversion target builds on: it turns a PDF
page into logical, positioned content - lines clustered into paragraph blocks,
document-wide font mapping, table-cell fill sampling, geometry helpers - and
nothing more. It has no opinion about output.

Each target then applies its OWN placement policy over this layer:
  * app/converter.py       - positioned PPTX shapes + native tables
  * app/edit_model.py      - absolutely-positioned editable HTML blocks
  * app/pdf_edit_export.py - real text drawn back into a PDF

Keeping placement OUT of here is deliberate: PPTX wants fixed positioned
blocks, Word/HTML want flowing text, and a PDF redraw wants baseline-true
coordinates. Sharing extraction while separating placement is what stops a
change made for one target from silently shifting another's output. Do not add
target-specific placement logic to this module.
"""

from __future__ import annotations

import os
import re

import fitz
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# --------------------------------------------------------------------------- #
# Shared constants
# --------------------------------------------------------------------------- #
EMU_PER_PT = 12700          # 1 point = 1/72 inch = 12700 EMU
SAMPLE_DPI = 150            # render resolution for sampling table cell fills

# Paragraph clustering.
PARA_GAP_FACTOR = 1.45      # lines within this * line-height belong together
PARA_BREAK_FACTOR = 1.6     # a gap > this * leading starts a new paragraph
SIZE_TOLERANCE = 1.2        # pt; lines whose size differs by more split apart
LEFT_TOLERANCE = 0.16       # fraction of block width; left edges within this align

# PyMuPDF span flag bits (the "serif" bit is unreliable on many PDFs, so font
# classification is name-first and only falls back to flags for mono).
FLAG_ITALIC = 1 << 1
FLAG_MONO = 1 << 3
FLAG_BOLD = 1 << 4

# --------------------------------------------------------------------------- #
# Font mapping (document-wide, consistent, metric-compatible)
# --------------------------------------------------------------------------- #
METRIC_MAP = [
    ("carlito", "Calibri"), ("calibri", "Calibri"),
    ("arimo", "Arial"), ("arialmt", "Arial"), ("arial", "Arial"),
    ("helvetica", "Arial"), ("liberationsans", "Arial"), ("notosans", "Arial"),
    ("segoeui", "Segoe UI"), ("verdana", "Verdana"), ("tahoma", "Tahoma"),
    ("trebuchet", "Trebuchet MS"),
    ("tinos", "Times New Roman"), ("timesnewroman", "Times New Roman"),
    ("times", "Times New Roman"), ("liberationserif", "Times New Roman"),
    ("georgia", "Georgia"), ("cambria", "Cambria"),
    ("notoserif", "Georgia"), ("ptserif", "Georgia"), ("garamond", "Georgia"),
    ("minion", "Georgia"),
    ("cousine", "Consolas"), ("consolas", "Consolas"),
    ("couriernew", "Courier New"), ("courier", "Courier New"),
    ("liberationmono", "Consolas"),
]
FALLBACK_SANS = "Arial"
FALLBACK_SERIF = "Georgia"
FALLBACK_MONO = "Consolas"
FALLBACK_SYMBOL = "Segoe UI Symbol"

# Width measurement for the SUBSTITUTED font, so a reflowed paragraph can be
# tracked back to the source's line widths. The browser engine measures with
# canvas; here the real Windows font file is measured through PyMuPDF, which is
# already a dependency. A family whose file is missing measures 0 and gets no
# tracking at all - the same safe degradation as a browser lacking the family.
_WIN_FONTS = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "Fonts")
_FONT_FILES = {
    # family: (regular, bold, italic, bold-italic)
    "arial": ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    "calibri": ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "cambria": ("cambria.ttc", "cambriab.ttf", "cambriai.ttf", "cambriaz.ttf"),
    "consolas": ("consola.ttf", "consolab.ttf", "consolai.ttf", "consolaz.ttf"),
    "courier new": ("cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"),
    "georgia": ("georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf"),
    "segoe ui": ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"),
    "segoe ui symbol": ("seguisym.ttf",) * 4,
    "tahoma": ("tahoma.ttf", "tahomabd.ttf", "tahoma.ttf", "tahomabd.ttf"),
    "times new roman": ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "trebuchet ms": ("trebuc.ttf", "trebucbd.ttf", "trebucit.ttf", "trebucbi.ttf"),
    "verdana": ("verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf"),
}
_font_cache: dict = {}


def _measure_font(family: str, bold: bool, italic: bool):
    key = ((family or "").lower(), bool(bold), bool(italic))
    if key in _font_cache:
        return _font_cache[key]
    files = _FONT_FILES.get(key[0])
    font = None
    if files:
        path = os.path.join(_WIN_FONTS, files[(2 if italic else 0) + (1 if bold else 0)])
        try:
            font = fitz.Font(fontfile=path)
        except Exception:
            font = None
    _font_cache[key] = font
    return font


def _text_width_pt(text: str, family: str, size: float, bold: bool, italic: bool) -> float:
    if not text or size <= 0:
        return 0.0
    font = _measure_font(family, bold, italic)
    if font is None:
        return 0.0
    try:
        return font.text_length(text, size)
    except Exception:
        return 0.0


_SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")
_STYLE_WORDS = re.compile(
    r"[-,]?\s*(bold|italic|oblique|regular|light|medium|semibold|demibold|"
    r"black|book|condensed|narrow|roman|mt|ps|psmt)\b", re.IGNORECASE)


def _clean_font_name(raw: str) -> str:
    name = _SUBSET_PREFIX.sub("", raw or "")
    name = _STYLE_WORDS.sub("", name)
    name = name.replace("-", " ").replace(",", " ")
    return re.sub(r"\s+", " ", name).strip()


class FontMapper:
    """Maps each distinct PDF font to one Windows font, consistently per document."""

    def __init__(self):
        self.cache: dict[str, str] = {}
        self.substitutions: dict[str, str] = {}   # "NotoSerif -> Georgia"

    def map(self, raw_name: str, flags: int) -> str:
        if raw_name in self.cache:
            return self.cache[raw_name]
        cleaned = _clean_font_name(raw_name)
        low = cleaned.lower()
        compact = re.sub(r"[^a-z]", "", low)

        target = None
        for key, val in METRIC_MAP:
            if key in compact:
                target = val
                break
        if target is None:
            if any(k in low for k in ("mono", "consol", "courier")) or (flags & FLAG_MONO):
                target = FALLBACK_MONO
            elif "sans" in low:
                target = FALLBACK_SANS
            elif any(k in low for k in ("serif", "times", "roman", "georgia",
                                        "cambria", "minion", "garamond")):
                target = FALLBACK_SERIF
            elif any(k in low for k in ("symbol", "wingding", "dingbat", "webding")):
                target = FALLBACK_SYMBOL
            else:
                target = FALLBACK_SANS

        self.cache[raw_name] = target
        base = cleaned or (raw_name or "unknown")
        if base.lower() != target.lower():
            self.substitutions[base] = target
        return target


# --------------------------------------------------------------------------- #
# Geometry / colour helpers
# --------------------------------------------------------------------------- #
def _int_color_to_rgb(color_int: int) -> RGBColor:
    color_int = int(color_int) & 0xFFFFFF
    return RGBColor((color_int >> 16) & 0xFF, (color_int >> 8) & 0xFF, color_int & 0xFF)


def _center(bbox):
    return ((bbox[0] + bbox[2]) / 2.0, (bbox[1] + bbox[3]) / 2.0)


def _point_in(bbox, x, y, tol=2.0):
    return (bbox[0] - tol) <= x <= (bbox[2] + tol) and (bbox[1] - tol) <= y <= (bbox[3] + tol)


def _sample_fill(pix: fitz.Pixmap, bbox, z: float) -> tuple:
    """Median background colour of a cell, sampled away from its text."""
    x0, y0, x1, y1 = bbox
    samples = []
    for fx, fy in ((0.10, 0.22), (0.10, 0.7), (0.9, 0.22), (0.9, 0.7), (0.5, 0.12)):
        sx = int((x0 + (x1 - x0) * fx) * z)
        sy = int((y0 + (y1 - y0) * fy) * z)
        sx = min(max(sx, 0), pix.width - 1)
        sy = min(max(sy, 0), pix.height - 1)
        samples.append(pix.pixel(sx, sy))
    r = sorted(s[0] for s in samples)[len(samples) // 2]
    g = sorted(s[1] for s in samples)[len(samples) // 2]
    b = sorted(s[2] for s in samples)[len(samples) // 2]
    return (r, g, b)


# --------------------------------------------------------------------------- #
# Line model + clustering into logical blocks
# --------------------------------------------------------------------------- #
def _collect_lines(text_dict) -> list:
    """Flatten the page into a list of text lines, each with its spans + bbox."""
    lines = []
    for block in text_dict.get("blocks", []):
        if block.get("type", 0) != 0:
            continue
        for ln in block.get("lines", []):
            spans = [s for s in ln.get("spans", []) if s.get("text", "") != ""]
            if not spans or not any(s["text"].strip() for s in spans):
                continue
            bbox = ln["bbox"]
            sizes = [s.get("size", 10.0) for s in spans]
            lines.append({
                "bbox": bbox,
                "spans": spans,
                "size": max(sizes),
                "x0": bbox[0], "y0": bbox[1], "x1": bbox[2], "y1": bbox[3],
            })
    # Overprinted text (the same run drawn twice for weight or shadow) must
    # extract once. The browser engine applies the same rule.
    seen, out = set(), []
    for ln in lines:
        key = ("".join(s["text"] for s in ln["spans"]),
               round(ln["x0"], 1), round(ln["y0"], 1), round(ln["size"], 1))
        if key in seen:
            continue
        seen.add(key)
        out.append(ln)
    return out


def _line_alignment(lines, region_x0, region_x1):
    """Infer paragraph alignment from how lines sit within their column."""
    if len(lines) < 2:
        return None
    width = max(region_x1 - region_x0, 1.0)
    lefts = [ln["x0"] - region_x0 for ln in lines]
    rights = [region_x1 - ln["x1"] for ln in lines]
    centers = [((ln["x0"] + ln["x1"]) / 2) - ((region_x0 + region_x1) / 2) for ln in lines]

    def spread(v):
        return (max(v) - min(v)) / width

    if spread(centers) < 0.05 and spread(lefts) > 0.08 and spread(rights) > 0.08:
        return PP_ALIGN.CENTER
    if spread(rights) < 0.05 and spread(lefts) > 0.08:
        return PP_ALIGN.RIGHT
    return PP_ALIGN.LEFT


BULLET_LIKE = set("•◦‣·∙▪▫◾--")


def _is_marker(line) -> bool:
    """A standalone bullet/marker glyph line (e.g. •, ‣, or a symbol-font PUA bullet).

    A lone glyph from a dingbat font counts whatever its code point: the whole
    font is ornaments, and a detached ZapfDingbats "H" is a hollow circle.
    """
    txt = "".join(s.get("text", "") for s in line["spans"]).strip()
    if not txt or len(txt) > 2:
        return False
    if (line["x1"] - line["x0"]) > 14:
        return False
    if all(_font_is_dingbat(s.get("font")) for s in line["spans"] if (s.get("text") or "").strip()):
        return True
    return all((c in BULLET_LIKE) or (0xF000 <= ord(c) <= 0xF0FF) or c in "-*" for c in txt)


_MARKER_START = re.compile(r"^\s*(\S)\s")

# A bullet is not a character, it is a character in a dingbat font. The W3C
# WCAG working draft sets every list marker as a 7pt ZapfDingbats "H", which
# draws a hollow circle; the outer level uses "G". Nothing about the code point
# says "bullet" - it is an ASCII letter - so every code-point test missed it,
# the table of contents clustered into ONE paragraph, and 25 entries reflowed
# into a block of prose. It was the corpus's worst page.
#
# A whole dingbat font is ornaments, so any lone glyph from one leads a list
# item. Symbol fonts are NOT: they carry Greek and maths, and arXiv papers open
# lines with them, so a Symbol glyph must still look like a bullet (Bank of
# England's markers are U+F0B7 in SymbolMT).
DINGBAT_FONTS = ("zapfdingbats", "dingbat", "wingding", "webding")


def _font_is_dingbat(name) -> bool:
    n = (name or "").lower()
    return any(d in n for d in DINGBAT_FONTS)


def _bullet_shaped(glyph) -> bool:
    return (glyph in BULLET_LIKE) or glyph == "*" or (0xF000 <= ord(glyph) <= 0xF0FF)


def _leading_marker_span(line) -> int:
    """Index of a leading marker span, or -1.

    The marker must be a lone glyph with real text after it, or it is a drop
    cap or a maths run, not a list item.
    """
    spans = line.get("spans") or []
    if len(spans) < 2:
        return -1
    glyph = (spans[0].get("text") or "").strip()
    if len(glyph) != 1:
        return -1
    if not any((s.get("text") or "").strip() for s in spans[1:]):
        return -1
    if _font_is_dingbat(spans[0].get("font")):
        return 0
    return 0 if _bullet_shaped(glyph) else -1


def _starts_with_marker(line) -> bool:
    """A list whose markers were never separate glyph runs.

    Bank of England reports draw each item as one line already reading
    " Andrew Bailey, Chair": there is no marker line for _attach_markers to
    find, so nothing flagged the item as a list item, and nine evenly-leaded
    bullets clustered into ONE paragraph and reflowed onto two lines. A line
    that opens with a bullet glyph and a space is a list item however the
    marker got there.
    """
    txt = "".join(s.get("text", "") for s in line["spans"])
    m = _MARKER_START.match(txt)
    if not m:
        return False
    c = m.group(1)
    return (c in BULLET_LIKE) or c == "*" or (0xF000 <= ord(c) <= 0xF0FF)


def _attach_markers(lines) -> list:
    """Merge bullet-marker lines into the text line they introduce, turning them
    into a real "• " prefix so a list becomes one text box of bulleted paragraphs
    rather than a swarm of tiny fragments."""
    markers = [l for l in lines if _is_marker(l)]
    texts = [l for l in lines if not _is_marker(l)]
    for t in texts:
        mi = _leading_marker_span(t)
        if mi >= 0:
            # Normalise the marker to a real bullet at the text's own size.
            # Left alone, a 7pt ZapfDingbats "H" maps to a substituted font and
            # ships as a tiny letter H beside every list item.
            after = t["spans"][mi + 1:]
            body = next((s for s in after if (s.get("text") or "").strip()), after[0])
            t["spans"] = [{"text": "• ", "font": body.get("font", ""),
                           "size": t["size"],
                           "flags": int(body.get("flags", 0)) & ~FLAG_BOLD,
                           "color": body.get("color", 0)}] + after
            t["bullet"] = True
        elif _starts_with_marker(t):
            t["bullet"] = True
    for m in markers:
        mcy = (m["y0"] + m["y1"]) / 2
        best, best_d = None, 1e9
        for t in texts:
            if t["x0"] <= m["x0"] + 1 or t.get("bullet"):
                continue
            tcy = (t["y0"] + t["y1"]) / 2
            if abs(tcy - mcy) <= max(t["y1"] - t["y0"], t["size"]) * 0.9:
                d = abs(tcy - mcy) + abs(t["x0"] - m["x1"]) * 0.01
                if d < best_d:
                    best_d, best = d, t
        if best is not None:
            first = best["spans"][0]
            bullet = {"text": "• ", "font": first.get("font", ""),
                      "size": best["size"], "flags": int(first.get("flags", 0)) & ~FLAG_BOLD,
                      "color": first.get("color", 0)}
            best["spans"] = [bullet] + best["spans"]
            best["bullet"] = True
    return texts


def _cluster_lines(lines) -> list:
    """Group lines into logical blocks sharing column, size and leading."""
    if not lines:
        return []
    lines = sorted(lines, key=lambda l: (round(l["y0"], 1), l["x0"]))
    clusters = []
    current = [lines[0]]

    for ln in lines[1:]:
        prev = current[-1]
        line_h = max(prev["y1"] - prev["y0"], prev["size"], 1.0)
        gap = ln["y0"] - prev["y1"]
        block_w = max(max(c["x1"] for c in current) - min(c["x0"] for c in current), 1.0)
        left_close = abs(ln["x0"] - prev["x0"]) <= LEFT_TOLERANCE * block_w + 4
        overlap = min(ln["x1"], prev["x1"]) - max(ln["x0"], prev["x0"])
        x_overlap = overlap > 0.3 * min(ln["x1"] - ln["x0"], prev["x1"] - prev["x0"])
        size_close = abs(ln["size"] - prev["size"]) <= SIZE_TOLERANCE
        vertical_close = -0.4 * line_h <= gap <= PARA_GAP_FACTOR * line_h

        if size_close and vertical_close and (left_close or x_overlap):
            current.append(ln)
        else:
            clusters.append(current)
            current = [ln]
    clusters.append(current)
    return clusters


def _split_paragraphs(cluster) -> list:
    """Within a block, split into paragraphs on larger-than-leading vertical gaps."""
    if len(cluster) == 1:
        return [cluster]
    deltas = [cluster[i + 1]["y0"] - cluster[i]["y0"] for i in range(len(cluster) - 1)]
    leading = sorted(deltas)[len(deltas) // 2] if deltas else cluster[0]["size"]
    paras, cur = [], [cluster[0]]
    for i in range(1, len(cluster)):
        step = cluster[i]["y0"] - cluster[i - 1]["y0"]
        if cluster[i].get("bullet") or step > PARA_BREAK_FACTOR * leading:
            paras.append(cur)
            cur = [cluster[i]]
        else:
            cur.append(cluster[i])
    paras.append(cur)
    return paras


# --------------------------------------------------------------------------- #
# Glyph-outline paths: text drawn as filled vector curves (statement class)
# --------------------------------------------------------------------------- #
# Statement generators paint visible words as filled glyph-outline paths with
# an invisible text layer on top for selectability. Consequences if outlines
# are treated as ordinary art: the hybrid background keeps the painted word
# while the invisible text re-emits editable (ghost doubling), and run colour
# comes from the invisible layer (usually black) instead of the ink.
# A fill is a glyph-outline blob iff it contains curves (glyphs always do;
# rules/underlines/cell fills never do), it is not huge, and its bbox is
# mostly covered by the union of extracted text-line bboxes.

def _glyph_pad(line) -> float:
    return max(2.0, 0.3 * (line.get("size") or (line["y1"] - line["y0"]) or 10))


def _is_glyph_outline(rect, has_curve, lines, page_area) -> bool:
    if not has_curve:
        return False
    fw, fh = rect[2] - rect[0], rect[3] - rect[1]
    fa = fw * fh
    if fa <= 0 or (page_area and fa > 0.25 * page_area):
        return False
    covered = 0.0
    for ln in lines:
        pad = _glyph_pad(ln)
        ix = min(rect[2], ln["x1"] + pad) - max(rect[0], ln["x0"] - pad)
        iy = min(rect[3], ln["y1"] + pad) - max(rect[1], ln["y0"] - pad)
        if ix > 0 and iy > 0:
            covered += ix * iy
        if covered >= 0.70 * fa:
            return True
    return False


def _glyph_outline_drawings(drawings, lines, page_area) -> list:
    """The subset of page.get_drawings() that is glyph-outline ink over
    extracted text lines: [(rect, fill_rgb01 or None)]."""
    out = []
    for d in drawings:
        if d.get("fill") is None:
            continue
        r = d.get("rect")
        if r is None:
            continue
        has_curve = any(it and it[0] == "c" for it in d.get("items", []))
        rect = (r.x0, r.y0, r.x1, r.y1)
        if _is_glyph_outline(rect, has_curve, lines, page_area):
            out.append((rect, d.get("fill")))
    return out


def _inherit_glyph_colors(lines, drawings, page_w, page_h) -> None:
    """Spans re-emitted from an invisible text layer inherit the colour of the
    glyph-outline fills actually painted over them. Mutates spans in place."""
    page_area = (page_w or 0) * (page_h or 0)
    glyphs = _glyph_outline_drawings(drawings, lines, page_area)
    if not glyphs:
        return
    for ln in lines:
        pad = _glyph_pad(ln)
        over = [(rect, fill) for rect, fill in glyphs
                if min(rect[2], ln["x1"] + pad) > max(rect[0], ln["x0"] - pad)
                and min(rect[3], ln["y1"] + pad) > max(rect[1], ln["y0"] - pad)]
        if not over:
            continue
        total_chars = max(1, sum(len(s["text"]) for s in ln["spans"]))
        x = ln["x0"]
        for sp in ln["spans"]:
            w = max(1.0, ln["x1"] - ln["x0"]) * (len(sp["text"]) / total_chars)
            sx0, sx1 = x, x + w
            x = sx1
            best, best_cover = None, 0.0
            for rect, fill in over:
                ov = min(rect[2], sx1) - max(rect[0], sx0)
                if ov > best_cover:
                    best_cover, best = ov, fill
            if best is not None and best_cover >= 0.35 * (sx1 - sx0):
                r, g, b = (int(round(v * 255)) for v in best[:3])
                sp["color"] = (r << 16) | (g << 8) | b


# --------------------------------------------------------------------------- #
# Column-alignment inference for UNRULED tables
# --------------------------------------------------------------------------- #
# Statement ledgers frequently ship with no ruling lines, so the lines-strategy
# detector never sees them. The remaining signal is alignment: consecutive
# baseline rows whose spans cluster at shared x-rails (left OR right edge, so
# right-aligned money columns count), usually with a header row on top.
# Deliberately conservative so prose can never tabulate: a run only starts at
# a row with >= _MIN_COLS spans, needs >= _MIN_DATA_ROWS + header rows and
# >= _MIN_COLS supported columns, any stray span breaks the run, and column
# x-extents may not overlap. Mirrors assets/js/engine/tables.js.

_COL_TOL = 3.5
_MIN_DATA_ROWS = 3
_MIN_COLS = 3
_ROW_PITCH_FACTOR = 2.6


class InferredTableRow:
    def __init__(self, cells):
        self.cells = cells


class InferredTable:
    """fitz.table.Table-shaped result of column inference (unruled source)."""

    def __init__(self, bbox, rows):
        self.bbox = bbox
        self.rows = rows
        self.row_count = len(rows)
        self.col_count = len(rows[0].cells) if rows else 0
        self.inferred = True


# A money value: a currency symbol on digits (a statement's IN/OUT/BALANCE).
# Strict on purpose - a plain number is not enough - so it only fires on
# financial columns, never on the aligned prose of an instruction form.
_MONEY_RE = re.compile(r"^[£$€]\s?-?[\d,]+(\.\d{1,2})?$")


def _seg_is_money(seg) -> bool:
    return bool(_MONEY_RE.match("".join(s["text"] for s in seg["spans"]).strip()))


def _group_rows(lines):
    rows = []
    for ln in sorted(lines, key=lambda l: (l["y0"], l["x0"])):
        cy = (ln["y0"] + ln["y1"]) / 2
        r = rows[-1] if rows else None
        if r is not None and abs(cy - r["cy"]) <= 0.5 * max(ln["size"] or 8, r["size"] or 8):
            r["segs"].append(ln)
            r["cy"] = (r["cy"] * (len(r["segs"]) - 1) + cy) / len(r["segs"])
            r["size"] = max(r["size"], ln["size"] or 0)
        else:
            rows.append({"cy": cy, "size": ln["size"] or 8, "segs": [ln]})
    for r in rows:
        r["segs"].sort(key=lambda s: s["x0"])
        r["y0"] = min(s["y0"] for s in r["segs"])
        r["y1"] = max(s["y1"] for s in r["segs"])
    return rows


def _infer_aligned_tables(lines, existing_bboxes=()) -> list:
    loose = [ln for ln in lines
             if any(s["text"].strip() for s in ln["spans"])
             and not any(_point_in(b, *_center(ln["bbox"])) for b in existing_bboxes)]
    rows = _group_rows(loose)
    tables = []
    i = 0
    while i < len(rows):
        start = rows[i]
        if len(start["segs"]) < _MIN_COLS:
            i += 1
            continue
        clusters = [{"x0m": s["x0"], "x1m": s["x1"], "minX0": s["x0"], "maxX1": s["x1"],
                     "n": 1, "money": 1 if _seg_is_money(s) else 0}
                    for s in start["segs"]]
        run = [start]
        j = i + 1
        while j < len(rows):
            row, prev = rows[j], run[-1]
            if row["cy"] - prev["cy"] > _ROW_PITCH_FACTOR * max(prev["size"], row["size"], 8):
                break
            # Classify each span: land on a rail, fold into ONE existing column
            # (a wrapped continuation), open a clean new column, or straddle two
            # columns (a hard break).
            matched, anchored, plan, ok = 0, 0, [], True
            for seg in row["segs"]:
                c = next((c for c in clusters
                          if abs(seg["x0"] - c["x0m"]) <= _COL_TOL
                          or abs(seg["x1"] - c["x1m"]) <= _COL_TOL), None)
                if c is not None:
                    matched += 1
                    anchored += 1
                    plan.append((c, seg))
                    continue
                over = [c for c in clusters
                        if min(seg["x1"], c["maxX1"]) - max(seg["x0"], c["minX0"]) > 1]
                if not over:
                    plan.append((None, seg))
                elif (len(over) == 1 and seg["x0"] >= over[0]["minX0"] - _COL_TOL
                      and seg["x1"] <= over[0]["maxX1"] + _COL_TOL):
                    anchored += 1
                    plan.append((over[0], seg))
                else:
                    ok = False
                    break
            if not ok:
                break
            new_cols = sum(1 for c, _ in plan if c is None)
            # A data row lands >=2 spans on rails. Otherwise the only extension
            # is a WRAPPED CONTINUATION - every span folded into an existing
            # column, opening none - and only inside an established money ledger
            # (a column with >=2 currency values). Instruction forms have no
            # money column, so their aligned prose never folds and the strict
            # run-breaker holds where a looser one wrecked the corpus.
            money_ledger = any(c["money"] >= 2 for c in clusters)
            is_data = matched >= 2
            is_cont = (money_ledger and anchored >= 1
                       and anchored == len(row["segs"]) and new_cols == 0)
            if not is_data and not is_cont:
                break
            for c, seg in plan:
                if c is not None:
                    c["x0m"] = (c["x0m"] * c["n"] + seg["x0"]) / (c["n"] + 1)
                    c["x1m"] = (c["x1m"] * c["n"] + seg["x1"]) / (c["n"] + 1)
                    c["minX0"] = min(c["minX0"], seg["x0"])
                    c["maxX1"] = max(c["maxX1"], seg["x1"])
                    c["n"] += 1
                    if _seg_is_money(seg):
                        c["money"] += 1
                else:
                    clusters.append({"x0m": seg["x0"], "x1m": seg["x1"],
                                     "minX0": seg["x0"], "maxX1": seg["x1"],
                                     "n": 1, "money": 1 if _seg_is_money(seg) else 0})
            if is_data:
                run.append(row)
            else:
                # fold the continuation into the PREVIOUS row: extend its extent
                # and absorb its spans, so the description becomes a 2-line cell
                # rather than a separate ghost row (which overspilled/fragmented).
                # Recompute the centre from the extended extent, else the next
                # row sits ~2 line-heights below this row's FIRST baseline and
                # trips the pitch gate, re-fragmenting the ledger.
                prev["segs"].extend(row["segs"])
                prev["y1"] = max(prev["y1"], row["y1"])
                prev["cy"] = (prev["y0"] + prev["y1"]) / 2
            j += 1
        # A column needs support in >=25% of rows (min 3), so prose cannot
        # tabulate. EXCEPT a pure-money column: when every value in it is a
        # currency amount it is real even with a single entry - a statement
        # period with one payment IN still has an IN column, and dropping it
        # collapses IN into OUT. Currency-gated, so instruction forms (no
        # currency anywhere) are unaffected and keep the strict rule.
        thr = max(3, -(-len(run) // 4))
        supported = [c for c in clusters
                     if c["n"] >= thr or (c["money"] >= 1 and c["money"] == c["n"])]
        dense = sum(1 for r in run if len(r["segs"]) >= _MIN_COLS)
        if len(run) >= _MIN_DATA_ROWS + 1 and len(supported) >= _MIN_COLS and dense >= _MIN_DATA_ROWS:
            tables.append(_build_inferred_table(run, supported))
            i = j
        else:
            i += 1
    return tables


def _build_inferred_table(run, cols) -> InferredTable:
    cols = sorted(cols, key=lambda c: c["minX0"])
    col_bounds = [cols[0]["minX0"] - 2]
    for k in range(len(cols) - 1):
        col_bounds.append((cols[k]["maxX1"] + cols[k + 1]["minX0"]) / 2)
    col_bounds.append(cols[-1]["maxX1"] + 2)
    row_bounds = [run[0]["y0"] - 2]
    for k in range(len(run) - 1):
        row_bounds.append((run[k]["y1"] + run[k + 1]["y0"]) / 2)
    row_bounds.append(run[-1]["y1"] + 2)
    rows = []
    for ri in range(len(run)):
        cells = [(col_bounds[ci], row_bounds[ri], col_bounds[ci + 1], row_bounds[ri + 1])
                 for ci in range(len(cols))]
        rows.append(InferredTableRow(cells))
    bbox = (col_bounds[0], row_bounds[0], col_bounds[-1], row_bounds[-1])
    return InferredTable(bbox, rows)
