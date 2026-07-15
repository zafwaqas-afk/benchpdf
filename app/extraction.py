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
    return lines


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
    """A standalone bullet/marker glyph line (e.g. •, ‣, or a symbol-font PUA bullet)."""
    txt = "".join(s.get("text", "") for s in line["spans"]).strip()
    if not txt or len(txt) > 2:
        return False
    if (line["x1"] - line["x0"]) > 14:
        return False
    return all((c in BULLET_LIKE) or (0xF000 <= ord(c) <= 0xF0FF) or c in "-*" for c in txt)


def _attach_markers(lines) -> list:
    """Merge bullet-marker lines into the text line they introduce, turning them
    into a real "• " prefix so a list becomes one text box of bulleted paragraphs
    rather than a swarm of tiny fragments."""
    markers = [l for l in lines if _is_marker(l)]
    texts = [l for l in lines if not _is_marker(l)]
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
