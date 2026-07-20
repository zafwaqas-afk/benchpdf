"""
PDF -> editable PPTX conversion engine.

This module owns ONE thing: the PPTX *placement policy*. All PDF extraction
(line clustering into logical paragraphs, document-wide font mapping, table-cell
fill sampling, geometry) lives in app/extraction.py and is shared with the
editor. Here we take those extracted blocks and lay them down as positioned
PowerPoint shapes and native tables. Keep placement here; keep extraction there.

One PDF page => one slide. The engine works structurally, not by screenshotting:

  * TABLES   are detected from the PDF's vector ruling lines and rebuilt as NATIVE
             PowerPoint tables (rows/columns/merges, per-cell fill sampled from the
             PDF, per-cell text with real fonts). Table regions are excluded from
             any background render so nothing is drawn twice.
  * TEXT     spans are clustered into LOGICAL BLOCKS (paragraphs / headings) that
             become native, editable text boxes with word-wrap on. Wrapped lines are
             joined with a single space -- never a separator character.
  * IMAGES   embedded rasters are placed as pictures at their true position/size.
  * HYBRID   only genuinely non-tabular vector art falls back to an image render
             with editable text on top.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import Callable, Optional

import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# ---- shared extraction layer (do not re-implement these here) ----
from app.extraction import (
    EMU_PER_PT, SAMPLE_DPI, FLAG_ITALIC, FLAG_BOLD,
    FontMapper, _int_color_to_rgb, _center, _point_in, _sample_fill,
    _collect_lines, _line_alignment, _attach_markers, _cluster_lines,
    _split_paragraphs,
)

# --------------------------------------------------------------------------- #
# PPTX-placement-only constants
# --------------------------------------------------------------------------- #
HYBRID_DPI = 200            # render resolution for hybrid backgrounds
NONTABLE_ART_COVERAGE = 0.03   # >3% of page area in non-table drawings -> hybrid
# "No Style, Table Grid" - thin borders, no theme fills (we set fills per cell).
TABLE_GRID_STYLE = "{5940675A-B579-460E-94D1-54222C63F5DA}"


# --------------------------------------------------------------------------- #
# Report data structures
# --------------------------------------------------------------------------- #
@dataclass
class PageReport:
    page_number: int
    mode: str
    text_boxes: int = 0
    tables: int = 0
    images: int = 0
    vector_paths: int = 0
    substituted_fonts: list = field(default_factory=list)
    note: str = ""


@dataclass
class ConversionReport:
    source_pdf: str
    output_pptx: str
    page_count: int = 0
    pages: list = field(default_factory=list)
    scanned_warning: bool = False
    warnings: list = field(default_factory=list)

    @property
    def all_substituted_fonts(self) -> list:
        seen = {}
        for p in self.pages:
            for s in p.substituted_fonts:
                seen[s] = True
        return list(seen.keys())


# --------------------------------------------------------------------------- #
# PPTX placement: text blocks
# --------------------------------------------------------------------------- #
def _emit_paragraph(text_frame, para_lines, first, alignment, fonts: FontMapper):
    """Write one paragraph's lines as a single wrapped paragraph (space-joined)."""
    p = text_frame.paragraphs[0] if first else text_frame.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    prev_text = ""
    for li, ln in enumerate(para_lines):
        if li > 0 and prev_text and not prev_text[-1].isspace():
            # join wrapped lines with a single space - never a separator character
            first_span = ln["spans"][0]
            sep = p.add_run()
            sep.text = " "
            _style_run(sep, first_span, fonts)
        for sp in ln["spans"]:
            run = p.add_run()
            run.text = sp["text"]
            _style_run(run, sp, fonts)
            prev_text = sp["text"]
    return p


def _style_run(run, span, fonts: FontMapper):
    f = run.font
    f.name = fonts.map(span.get("font", ""), int(span.get("flags", 0)))
    size = float(span.get("size", 10.0))
    if size > 0:
        f.size = Pt(size)   # preserve the exact point size
    flags = int(span.get("flags", 0))
    f.bold = bool(flags & FLAG_BOLD)
    f.italic = bool(flags & FLAG_ITALIC)
    try:
        f.color.rgb = _int_color_to_rgb(span.get("color", 0))
    except Exception:
        pass


def _add_text_block(slide, cluster, scale, off_x, off_y, fonts: FontMapper):
    x0 = min(c["x0"] for c in cluster)
    y0 = min(c["y0"] for c in cluster)
    x1 = max(c["x1"] for c in cluster)
    y1 = max(c["y1"] for c in cluster)

    left = Emu(int((off_x + x0 * scale) * EMU_PER_PT))
    top = Emu(int((off_y + y0 * scale) * EMU_PER_PT))
    width = Emu(max(int((x1 - x0) * scale * EMU_PER_PT), EMU_PER_PT))
    height = Emu(max(int((y1 - y0) * scale * EMU_PER_PT), EMU_PER_PT // 2))

    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE          # never let PPT change the font size
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    alignment = _line_alignment(cluster, x0, x1)
    for pi, para in enumerate(_split_paragraphs(cluster)):
        _emit_paragraph(tf, para, pi == 0, alignment, fonts)
    return tb


# --------------------------------------------------------------------------- #
# PPTX placement: native tables
# --------------------------------------------------------------------------- #
def _set_table_grid_style(table):
    tblPr = table._tbl.tblPr
    if tblPr is None:
        return
    for child in list(tblPr):
        if child.tag == qn("a:tableStyleId"):
            tblPr.remove(child)
    el = tblPr.makeelement(qn("a:tableStyleId"), {})
    el.text = TABLE_GRID_STYLE
    tblPr.append(el)


def _lines_in(bbox, lines):
    out = []
    for ln in lines:
        cx, cy = _center(ln["bbox"])
        if _point_in(bbox, cx, cy):
            out.append(ln)
    return out


def _fill_cell(cell, para_lines, fonts: FontMapper, fill_rgb, is_dark):
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(*fill_rgb)
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(1)
    cell.margin_bottom = Pt(1)
    cell.vertical_anchor = MSO_ANCHOR.TOP

    tf = cell.text_frame
    tf.word_wrap = True
    if not para_lines:
        # An empty cell still carries a paragraph, and without an explicit size
        # PowerPoint reserves line height for its 18pt default, silently
        # inflating every sparse row until dense tables overflow the slide.
        # (Found by the phase-3 render comparison against the browser engine.)
        tf.paragraphs[0].font.size = Pt(6)
        return
    para_lines = sorted(para_lines, key=lambda l: (round(l["y0"], 1), l["x0"]))
    paras = _split_paragraphs(para_lines)
    for pi, para in enumerate(paras):
        _emit_paragraph(tf, para, pi == 0, None, fonts)


def _add_table(slide, table, page_lines, pix, z, scale, off_x, off_y, fonts: FontMapper):
    nrows, ncols = table.row_count, table.col_count
    if nrows < 1 or ncols < 1:
        return False
    bx0, by0, bx1, by1 = table.bbox

    left = Emu(int((off_x + bx0 * scale) * EMU_PER_PT))
    top = Emu(int((off_y + by0 * scale) * EMU_PER_PT))
    width = Emu(int((bx1 - bx0) * scale * EMU_PER_PT))
    height = Emu(int((by1 - by0) * scale * EMU_PER_PT))

    gf = slide.shapes.add_table(nrows, ncols, left, top, width, height)
    table_obj = gf.table
    table_obj.first_row = False
    table_obj.horz_banding = False
    _set_table_grid_style(table_obj)

    row0 = table.rows[0].cells
    for ci in range(ncols):
        cb = row0[ci]
        if cb is not None:
            table_obj.columns[ci].width = Emu(max(int((cb[2] - cb[0]) * scale * EMU_PER_PT), EMU_PER_PT))
    for ri in range(nrows):
        rcells = [c for c in table.rows[ri].cells if c is not None]
        if rcells:
            rh = max(c[3] for c in rcells) - min(c[1] for c in rcells)
            table_obj.rows[ri].height = Emu(max(int(rh * scale * EMU_PER_PT), EMU_PER_PT // 3))

    for ri in range(nrows):
        for ci in range(ncols):
            cb = table.rows[ri].cells[ci]
            cell = table_obj.cell(ri, ci)
            if cb is None:
                cell.fill.background()
                cell.text_frame.paragraphs[0].font.size = Pt(6)
                continue
            fill = _sample_fill(pix, cb, z)
            is_dark = (fill[0] + fill[1] + fill[2]) < 384
            cell_lines = _lines_in(cb, page_lines)
            _fill_cell(cell, cell_lines, fonts, fill, is_dark)
    return True


# --------------------------------------------------------------------------- #
# Hybrid background (only for genuine non-table vector art)
# --------------------------------------------------------------------------- #
def _nontable_art_ratio(page, drawings, table_bboxes) -> float:
    page_area = abs(page.rect.width * page.rect.height) or 1.0
    area = 0.0
    for d in drawings:
        r = d.get("rect")
        if r is None:
            continue
        w, h = r.width, r.height
        if w <= 2 or h <= 2:
            continue
        if w > 0.92 * page.rect.width and h > 0.92 * page.rect.height:
            continue
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        if any(_point_in(tb, cx, cy) for tb in table_bboxes):
            continue
        area += w * h
    return area / page_area


def _render_hybrid_bg(src_doc, page_index, table_bboxes) -> bytes:
    zoom = HYBRID_DPI / 72.0
    tmp = fitz.open()
    tmp.insert_pdf(src_doc, from_page=page_index, to_page=page_index)
    pg = tmp[0]
    try:
        pg.add_redact_annot(pg.rect, text=None)
        pg.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                            graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    except Exception:
        tmp.close()
        tmp = fitz.open()
        tmp.insert_pdf(src_doc, from_page=page_index, to_page=page_index)
        pg = tmp[0]
    for tb in table_bboxes:
        try:
            pg.draw_rect(fitz.Rect(*tb), color=(1, 1, 1), fill=(1, 1, 1))
        except Exception:
            pass
    pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    data = pix.tobytes("png")
    tmp.close()
    return data


def _lock_picture(picture):
    try:
        pic = picture._element
        cNvPicPr = pic.find(qn("p:nvPicPr")).find(qn("p:cNvPicPr"))
        locks = cNvPicPr.find(qn("a:picLocks"))
        if locks is None:
            locks = cNvPicPr.makeelement(qn("a:picLocks"), {})
            cNvPicPr.append(locks)
        for attr in ("noMove", "noResize", "noSelect", "noChangeAspect"):
            locks.set(attr, "1")
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Main entry point
# --------------------------------------------------------------------------- #
def convert_pdf_to_pptx(
    pdf_path: str,
    pptx_path: str,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> ConversionReport:
    report = ConversionReport(source_pdf=pdf_path, output_pptx=pptx_path)
    fonts = FontMapper()

    doc = fitz.open(pdf_path)
    report.page_count = doc.page_count
    if doc.page_count == 0:
        report.warnings.append("The PDF has no pages.")
        doc.close()
        return report

    first = doc[0]
    slide_w_pt, slide_h_pt = first.rect.width, first.rect.height
    prs = Presentation()
    prs.slide_width = Emu(int(slide_w_pt * EMU_PER_PT))
    prs.slide_height = Emu(int(slide_h_pt * EMU_PER_PT))
    blank = prs.slide_layouts[6]

    pages_with_text = 0

    for i in range(doc.page_count):
        page = doc[i]
        if progress_callback:
            progress_callback(i, doc.page_count, f"Converting page {i + 1} of {doc.page_count}")
        slide = prs.slides.add_slide(blank)

        pw, ph = page.rect.width, page.rect.height
        scale = min(slide_w_pt / pw, slide_h_pt / ph) if pw and ph else 1.0
        off_x = (slide_w_pt - pw * scale) / 2.0
        off_y = (slide_h_pt - ph * scale) / 2.0

        text_dict = page.get_text("dict")
        raw_text = page.get_text("text").strip()
        try:
            drawings = page.get_drawings()
        except Exception:
            drawings = []

        # ---- scanned / image-only ----
        if not raw_text:
            pix = page.get_pixmap(matrix=fitz.Matrix(HYBRID_DPI / 72.0, HYBRID_DPI / 72.0), alpha=False)
            bg = slide.shapes.add_picture(
                io.BytesIO(pix.tobytes("png")),
                Emu(int(off_x * EMU_PER_PT)), Emu(int(off_y * EMU_PER_PT)),
                Emu(int(pw * scale * EMU_PER_PT)), Emu(int(ph * scale * EMU_PER_PT)))
            _lock_picture(bg)
            report.pages.append(PageReport(page_number=i + 1, mode="image-only",
                                           vector_paths=len(drawings),
                                           note="No extractable text (scanned/image-only page)."))
            report.scanned_warning = True
            continue

        pages_with_text += 1

        # ---- tables ----
        try:
            found = page.find_tables(strategy="lines")
            tables = [t for t in found.tables if t.row_count >= 1 and t.col_count >= 1]
        except Exception:
            tables = []
        table_bboxes = [t.bbox for t in tables]

        all_lines = _collect_lines(text_dict)
        loose_lines = [ln for ln in all_lines
                       if not any(_point_in(tb, *_center(ln["bbox"])) for tb in table_bboxes)]

        # ---- hybrid decision (non-table vector art only) ----
        art_ratio = _nontable_art_ratio(page, drawings, table_bboxes)
        use_hybrid = art_ratio > NONTABLE_ART_COVERAGE
        pr = PageReport(page_number=i + 1, mode="hybrid" if use_hybrid else "native",
                        vector_paths=len(drawings))

        if use_hybrid:
            pr.note = f"non-table vector art ({art_ratio*100:.0f}% of page)"
            png = _render_hybrid_bg(doc, i, table_bboxes)
            bg = slide.shapes.add_picture(
                io.BytesIO(png),
                Emu(int(off_x * EMU_PER_PT)), Emu(int(off_y * EMU_PER_PT)),
                Emu(int(pw * scale * EMU_PER_PT)), Emu(int(ph * scale * EMU_PER_PT)))
            _lock_picture(bg)
        else:
            for img in page.get_images(full=True):
                xref = img[0]
                try:
                    rects = page.get_image_rects(xref)
                except Exception:
                    rects = []
                for rect in rects:
                    cx, cy = (rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2
                    if any(_point_in(tb, cx, cy) for tb in table_bboxes):
                        continue
                    try:
                        base = doc.extract_image(xref)
                        slide.shapes.add_picture(
                            io.BytesIO(base["image"]),
                            Emu(int((off_x + rect.x0 * scale) * EMU_PER_PT)),
                            Emu(int((off_y + rect.y0 * scale) * EMU_PER_PT)),
                            Emu(int(rect.width * scale * EMU_PER_PT)),
                            Emu(int(rect.height * scale * EMU_PER_PT)))
                        pr.images += 1
                    except Exception:
                        pass

        # ---- native tables ----
        if tables:
            pix = page.get_pixmap(matrix=fitz.Matrix(SAMPLE_DPI / 72.0, SAMPLE_DPI / 72.0), alpha=False)
            z = SAMPLE_DPI / 72.0
            for t in tables:
                if _add_table(slide, t, all_lines, pix, z, scale, off_x, off_y, fonts):
                    pr.tables += 1

        # ---- loose text as logical blocks ----
        loose_lines = _attach_markers(loose_lines)
        for cluster in _cluster_lines(loose_lines):
            _add_text_block(slide, cluster, scale, off_x, off_y, fonts)
            pr.text_boxes += 1

        pr.substituted_fonts = [f"{k} -> {v}" for k, v in sorted(fonts.substitutions.items())]
        report.pages.append(pr)

    if progress_callback:
        progress_callback(doc.page_count, doc.page_count, "Saving presentation")

    if report.scanned_warning and pages_with_text == 0:
        report.warnings.append(
            "This PDF appears to be scanned or image-only (no selectable text). Pages were "
            "placed as images so nothing is lost, but there is no editable text to extract. "
            "OCR is not part of this tool.")
    elif report.scanned_warning:
        report.warnings.append("Some pages had no extractable text and were placed as images.")

    all_subs = [f"{k} -> {v}" for k, v in sorted(fonts.substitutions.items())]
    for p in report.pages:
        if p.mode != "image-only":
            p.substituted_fonts = all_subs

    prs.save(pptx_path)
    doc.close()
    return report
