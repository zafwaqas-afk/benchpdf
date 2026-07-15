"""
Editable page-model extraction for in-app PDF editing.

This is the SAME extraction pipeline as the PPTX converter - the same
line-clustering into logical paragraphs, the same table detection, the same
document-wide font mapping - with a different output target: instead of PPTX
shapes, it emits a JSON model of absolutely-positioned, editable HTML blocks
plus a text-free background image per page.

Everything is built lazily per page (open() is cheap; a page's blocks and its
background are produced on first request and cached) so a 50-page document
opens instantly and stays responsive as pages scroll into view.
"""

from __future__ import annotations

import html as _html
import threading

import fitz

from app import extraction as C   # shared extraction layer (not the PPTX placement)

EDITOR_BG_DPI = 150   # background render resolution for on-screen editing


def _hex(color_int: int) -> str:
    c = int(color_int) & 0xFFFFFF
    return f"#{c:06x}"


def _center(b):
    return ((b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0)


def _span_style(span):
    flags = int(span.get("flags", 0))
    return (bool(flags & C.FLAG_BOLD), bool(flags & C.FLAG_ITALIC),
            _hex(span.get("color", 0)), round(float(span.get("size", 10.0)), 1))


def _dominant(cluster, mapper):
    """Most common (family, size, color, bold, italic) across a cluster's spans,
    weighted by text length, plus the alignment."""
    from collections import Counter
    fam_c, size_c, col_c = Counter(), Counter(), Counter()
    bold_c, ital_c = Counter(), Counter()
    for ln in cluster:
        for sp in ln["spans"]:
            w = max(len(sp.get("text", "")), 1)
            b, i, col, sz = _span_style(sp)
            fam = mapper.map(sp.get("font", ""), int(sp.get("flags", 0)))
            fam_c[fam] += w
            size_c[sz] += w
            col_c[col] += w
            bold_c[b] += w
            ital_c[i] += w
    family = fam_c.most_common(1)[0][0] if fam_c else "Arial"
    size = size_c.most_common(1)[0][0] if size_c else 10.0
    color = col_c.most_common(1)[0][0] if col_c else "#000000"
    bold = bold_c.get(True, 0) > bold_c.get(False, 0)
    italic = ital_c.get(True, 0) > ital_c.get(False, 0)
    return family, size, color, bold, italic


def _runs_html(cluster, mapper, dom_color):
    """Build editable HTML for a cluster: one <div> per logical paragraph,
    inline <b>/<i>/<span color> runs preserved; wrapped lines joined by a space
    so the browser (and the PDF exporter) re-wrap naturally."""
    paras = C._split_paragraphs(cluster)
    out = []
    for para in paras:
        buf = []
        prev_text = ""
        for li, ln in enumerate(para):
            if li > 0 and prev_text and not prev_text[-1].isspace():
                buf.append(" ")
            for sp in ln["spans"]:
                text = sp.get("text", "")
                if text == "":
                    continue
                b, i, col, _sz = _span_style(sp)
                esc = _html.escape(text)
                if col and col != dom_color:
                    esc = f'<span style="color:{col}">{esc}</span>'
                if i:
                    esc = f"<i>{esc}</i>"
                if b:
                    esc = f"<b>{esc}</b>"
                buf.append(esc)
                prev_text = text
        content = "".join(buf).strip()
        if content:
            out.append(f"<div>{content}</div>")
    return "".join(out) or "<div><br></div>"


def _cell_html(lines, mapper, dom_color):
    if not lines:
        return ""
    lines = sorted(lines, key=lambda l: (round(l["y0"], 1), l["x0"]))
    return _runs_html(lines, mapper, dom_color)


class EditSession:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.mapper = C.FontMapper()
        self._lock = threading.Lock()
        self._page_cache: dict = {}
        self._bg_cache: dict = {}

    # -- summary (cheap) ----------------------------------------------------
    def summary(self):
        pages = []
        for i in range(self.doc.page_count):
            r = self.doc[i].rect
            pages.append({"index": i, "width": round(r.width, 2), "height": round(r.height, 2)})
        return {"page_count": self.doc.page_count, "pages": pages}

    # -- one page's editable model (lazy, cached) ---------------------------
    def page_model(self, index: int):
        with self._lock:
            if index in self._page_cache:
                return self._page_cache[index]
        page = self.doc[index]
        rect = page.rect
        text_dict = page.get_text("dict")
        try:
            found = page.find_tables(strategy="lines")
            tables = [t for t in found.tables if t.row_count >= 1 and t.col_count >= 1]
        except Exception:
            tables = []
        table_bboxes = [t.bbox for t in tables]

        all_lines = C._collect_lines(text_dict)
        loose = [ln for ln in all_lines
                 if not any(C._point_in(tb, *_center(ln["bbox"])) for tb in table_bboxes)]
        loose = C._attach_markers(loose)

        blocks = []

        # sample pixmap for table fills
        pix = None
        if tables:
            z = C.SAMPLE_DPI / 72.0
            pix = page.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)

        # --- tables ---
        for ti, t in enumerate(tables):
            cells = []
            for ri in range(t.row_count):
                for ci in range(t.col_count):
                    cb = t.rows[ri].cells[ci]
                    if cb is None:
                        continue
                    clines = [ln for ln in all_lines if C._point_in(cb, *_center(ln["bbox"]))]
                    fam, size, color, bold, italic = _dominant(clines, self.mapper) if clines \
                        else ("Arial", 9.0, "#000000", False, False)
                    fill = C._sample_fill(pix, cb, C.SAMPLE_DPI / 72.0) if pix else (255, 255, 255)
                    align = C._line_alignment(clines, cb[0], cb[2]) if len(clines) > 1 else None
                    cells.append({
                        "r": ri, "c": ci,
                        "bbox": [round(v, 2) for v in cb],
                        "fill": "#%02x%02x%02x" % fill,
                        "font": fam, "size": size, "color": color,
                        "bold": bold, "italic": italic,
                        "align": _align_name(align),
                        "html": _cell_html(clines, self.mapper, color),
                    })
            blocks.append({
                "id": f"t{ti}", "type": "table",
                "bbox": [round(v, 2) for v in t.bbox],
                "rows": t.row_count, "cols": t.col_count, "cells": cells,
            })

        # --- loose text blocks ---
        for bi, cluster in enumerate(C._cluster_lines(loose)):
            x0 = min(c["x0"] for c in cluster)
            y0 = min(c["y0"] for c in cluster)
            x1 = max(c["x1"] for c in cluster)
            y1 = max(c["y1"] for c in cluster)
            fam, size, color, bold, italic = _dominant(cluster, self.mapper)
            align = C._line_alignment(cluster, x0, x1)
            blocks.append({
                "id": f"b{bi}", "type": "text",
                "bbox": [round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)],
                "font": fam, "size": size, "color": color,
                "bold": bold, "italic": italic,
                "align": _align_name(align),
                "html": _runs_html(cluster, self.mapper, color),
            })

        model = {
            "index": index,
            "width": round(rect.width, 2), "height": round(rect.height, 2),
            "blocks": blocks,
            "has_text": bool(page.get_text("text").strip()),
            "substituted_fonts": [f"{k} → {v}" for k, v in sorted(self.mapper.substitutions.items())],
        }
        with self._lock:
            self._page_cache[index] = model
        return model

    # -- text-free background (lazy, cached) --------------------------------
    def background_png(self, index: int, dpi: int = EDITOR_BG_DPI) -> bytes:
        key = (index, dpi)
        with self._lock:
            if key in self._bg_cache:
                return self._bg_cache[key]
        tmp = fitz.open()
        tmp.insert_pdf(self.doc, from_page=index, to_page=index)
        pg = tmp[0]
        try:
            pg.add_redact_annot(pg.rect)
            pg.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                                graphics=fitz.PDF_REDACT_LINE_ART_NONE)
        except Exception:
            pass
        pix = pg.get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0), alpha=False)
        data = pix.tobytes("png")
        tmp.close()
        with self._lock:
            self._bg_cache[key] = data
        return data

    def close(self):
        try:
            self.doc.close()
        except Exception:
            pass


def _align_name(align):
    from pptx.enum.text import PP_ALIGN
    if align == PP_ALIGN.CENTER:
        return "center"
    if align == PP_ALIGN.RIGHT:
        return "right"
    return "left"
