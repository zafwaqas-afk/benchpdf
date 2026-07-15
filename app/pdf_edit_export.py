"""
Export an edited PDF.

Given the original PDF and an edits payload (only the pages the user actually
touched, each carrying the final state of ALL its blocks), this rebuilds each
edited page as:

    text-free raster background  +  real, editable text (insert_htmlbox) drawn
    with cmap-fixed, embedded, mapped fonts

and leaves every untouched page byte-identical to the source by writing an
incremental save over a copy of the original. Text stays REAL text on edited
pages (only the non-text graphics of those pages are rasterised into the
background) - nothing silently turns editable text into an image.
"""

from __future__ import annotations

import os
import shutil

import fitz

from app import fonts_local

EXPORT_BG_DPI = 300


class EditExportError(Exception):
    pass


def _families_in(blocks) -> set:
    fams = set()
    for b in blocks:
        if b.get("type") == "table":
            for c in b.get("cells", []):
                fams.add(c.get("font", "Arial"))
        else:
            fams.add(b.get("font", "Arial"))
    return fams


def _text_free_background(doc, index, dpi) -> bytes:
    tmp = fitz.open()
    tmp.insert_pdf(doc, from_page=index, to_page=index)
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
    return data


def _wrap_html(block, page_w, page_h, is_cell=False) -> tuple:
    """Return (rect, html, block_css) for a text block or a single table cell.
    All styling goes through the css parameter (never inline style=) so nested
    quotes can't break MuPDF's parser and drop the @font-face / font-family."""
    x0, y0, x1, y1 = block["bbox"]
    size = float(block.get("size", 10.0))
    align = block.get("align", "left")

    # Width headroom: a mapped font's metrics differ slightly from the original,
    # so a line that fit tightly in the source can spuriously wrap. Table cells
    # must stay within their column, but a loose single-line block (heading,
    # label) should keep its one line - extend its width so it can't wrap.
    x1_eff = x1
    if not is_cell:
        single_line = (y1 - y0) <= size * 1.7
        if single_line and align == "left":
            x1_eff = page_w - 6
        elif single_line:
            x1_eff = x1 + 16
        else:
            x1_eff = min(page_w - 6, x1 + 4)

    # generous bottom headroom so insert_htmlbox never shrinks the font to fit;
    # text is top-anchored at y0, so extending the bottom doesn't move it.
    rect = fitz.Rect(x0, y0, x1_eff, min(page_h, y1 + max((y1 - y0), size) * 6 + 40))
    color = block.get("color", "#000000")
    align = block.get("align", "left")
    family = fonts_local.known_family(block.get("font", "Arial"))
    inner = block.get("html", "") or ""
    block_css = (
        f'#blk{{font-family:"{family}";font-size:{size}pt;color:{color};'
        f'text-align:{align};line-height:1.15;margin:0;padding:0;}}'
        f'#blk div{{margin:0;padding:0;}}'
    )
    html = f'<div id="blk">{inner}</div>'
    return rect, html, block_css


def export_edited_pdf(source_pdf: str, edits: dict, output_path: str) -> dict:
    pages = (edits or {}).get("pages", {})
    if not pages:
        # no edits: a straight copy is trivially byte-identical
        shutil.copyfile(source_pdf, output_path)
        return {"edited_pages": [], "note": "no edits"}

    shutil.copyfile(source_pdf, output_path)
    doc = fitz.open(output_path)
    edited_indices = []

    try:
        for key in sorted(pages.keys(), key=lambda k: int(k)):
            index = int(key)
            if index < 0 or index >= doc.page_count:
                continue
            blocks = pages[key].get("blocks", [])
            page = doc[index]
            page_h = page.rect.height

            # 1) text-free background from the ORIGINAL page content
            bg = _text_free_background(doc, index, EXPORT_BG_DPI)

            # 2) clear the page entirely (text + line-art + images)
            page.add_redact_annot(page.rect)
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE,
                                  graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED,
                                  text=fitz.PDF_REDACT_TEXT_REMOVE)

            # 3) draw the text-free background
            page.insert_image(page.rect, stream=bg)

            # 4) fonts used on this page
            font_css, arch = fonts_local.build_export_fonts(_families_in(blocks))

            # 5) draw every block's real text
            page_w = page.rect.width
            for b in blocks:
                if b.get("type") == "table":
                    for cell in b.get("cells", []):
                        if not (cell.get("html") or "").strip():
                            continue
                        rect, html, block_css = _wrap_html(cell, page_w, page_h, is_cell=True)
                        _draw(page, rect, html, font_css + block_css, arch)
                else:
                    if not (b.get("html") or "").strip():
                        continue
                    rect, html, block_css = _wrap_html(b, page_w, page_h, is_cell=False)
                    _draw(page, rect, html, font_css + block_css, arch)

            edited_indices.append(index)

        doc.save(output_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
    finally:
        doc.close()

    return {"edited_pages": edited_indices}


def _draw(page, rect, html, font_css, arch):
    try:
        # scale_low=1 forbids shrinking the font to fit (preserves point size).
        page.insert_htmlbox(rect, html, css=font_css, archive=arch, scale_low=1)
    except Exception as exc:
        raise EditExportError(f"Failed to render a text block: {exc}") from exc
