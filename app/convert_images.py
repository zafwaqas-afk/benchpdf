"""
Images -> PDF. Entirely local (Pillow + img2pdf). Supports JPG, PNG, and HEIC
(via pillow-heif). Multiple images merge into ONE PDF, one page per image, in
the order given.
"""

from __future__ import annotations

import io
import os

import img2pdf
from PIL import Image

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    _HEIC_OK = True
except Exception:
    _HEIC_OK = False

SUPPORTED_EXT = {".jpg", ".jpeg", ".png"} | ({".heic", ".heif"} if _HEIC_OK else set())


class ImageConvertError(Exception):
    pass


def _prepare_bytes(path: str) -> bytes:
    """Return image bytes img2pdf can embed directly (JPEG/PNG passthrough,
    everything else - HEIC included - is decoded and re-encoded to PNG)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".jpg", ".jpeg"):
        with open(path, "rb") as f:
            data = f.read()
        try:
            img2pdf.get_imgmetadata(io.BytesIO(data))
            return data
        except Exception:
            pass  # fall through to re-encode a non-standard JPEG
    if ext == ".png":
        with open(path, "rb") as f:
            data = f.read()
        try:
            img2pdf.get_imgmetadata(io.BytesIO(data))
            return data
        except Exception:
            pass

    im = Image.open(path)
    if im.mode in ("RGBA", "P", "LA"):
        im = im.convert("RGB")
    elif im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


def images_to_pdf(paths: list[str], dst: str) -> str:
    if not paths:
        raise ImageConvertError("No images were provided.")
    blobs = []
    for p in paths:
        ext = os.path.splitext(p)[1].lower()
        if ext not in SUPPORTED_EXT:
            raise ImageConvertError(f"Unsupported image type: {os.path.basename(p)}")
        try:
            blobs.append(_prepare_bytes(p))
        except Exception as exc:
            raise ImageConvertError(f"Couldn't read {os.path.basename(p)}; it may be corrupt.") from exc

    try:
        pdf_bytes = img2pdf.convert(blobs)
    except Exception as exc:
        raise ImageConvertError(f"Couldn't build the PDF ({exc}).") from exc

    with open(dst, "wb") as f:
        f.write(pdf_bytes)
    return dst
