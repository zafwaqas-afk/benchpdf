"""
PDF -> images (per-page PNG/JPG at selectable DPI) and PDF -> plain text.
Both are pure PyMuPDF; no network, no Office COM.
"""

from __future__ import annotations

import os
import zipfile

import fitz  # PyMuPDF


class PdfMiscError(Exception):
    pass


def pdf_to_images(src: str, out_dir: str, fmt: str = "png", dpi: int = 150) -> list[str]:
    fmt = fmt.lower()
    if fmt not in ("png", "jpg", "jpeg"):
        raise PdfMiscError(f"Unsupported image format: {fmt}")
    ext = "jpg" if fmt in ("jpg", "jpeg") else "png"
    dpi = max(72, min(int(dpi), 600))

    try:
        doc = fitz.open(src)
    except Exception as exc:
        raise PdfMiscError("Couldn't open this PDF; it may be corrupt or password-protected.") from exc

    if doc.page_count == 0:
        doc.close()
        raise PdfMiscError("This PDF has no pages.")

    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    made = []
    for i in range(doc.page_count):
        page = doc[i]
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        out_path = os.path.join(out_dir, f"{stem}_page{i + 1:03d}.{ext}")
        if ext == "jpg":
            pix.save(out_path, jpg_quality=90)
        else:
            pix.save(out_path)
        made.append(out_path)
    doc.close()
    return made


def zip_files(paths: list[str], zip_path: str) -> str:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, arcname=os.path.basename(p))
    return zip_path


def pdf_to_text(src: str, dst: str) -> str:
    try:
        doc = fitz.open(src)
    except Exception as exc:
        raise PdfMiscError("Couldn't open this PDF; it may be corrupt or password-protected.") from exc

    if doc.page_count == 0:
        doc.close()
        raise PdfMiscError("This PDF has no pages.")

    parts = []
    for i in range(doc.page_count):
        parts.append(doc[i].get_text("text"))
    doc.close()

    text = ("\n\f\n").join(parts).strip() + "\n"
    with open(dst, "w", encoding="utf-8") as f:
        f.write(text)
    return dst
