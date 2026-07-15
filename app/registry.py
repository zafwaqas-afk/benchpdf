"""
Conversion matrix: the single source of truth for what this hub can convert,
used by both the server (dispatch + validation) and the UI (the grid + the
auto-detecting drop zone's target picker).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Optional

EXT_TO_TYPE = {
    ".pdf": "pdf",
    ".docx": "word", ".doc": "word",
    ".xlsx": "excel", ".xls": "excel",
    ".pptx": "powerpoint", ".ppt": "powerpoint",
    ".jpg": "image", ".jpeg": "image", ".png": "image",
    ".heic": "image", ".heif": "image",
}

TYPE_LABELS = {
    "pdf": "PDF", "word": "Word", "excel": "Excel",
    "powerpoint": "PowerPoint", "image": "Image", "url": "Web page",
}


def detect_type(filename: str) -> Optional[str]:
    import os
    ext = os.path.splitext(filename)[1].lower()
    return EXT_TO_TYPE.get(ext)


@dataclass
class Target:
    id: str
    label: str                 # short label for the grid cell, e.g. "Word → PDF"
    group: str                 # "to_pdf" | "from_pdf"
    source_type: str           # key into EXT_TO_TYPE values, or "url" or "image_multi"
    output_ext: str            # extension of the produced file (without dot), or "zip"
    accepts: tuple             # file extensions this target's input accepts
    multi: bool = False        # True: multiple dropped files become ONE job (merge)
    enabled: bool = True
    note: str = ""             # shown on disabled cells, e.g. "Coming soon"
    description: str = ""
    fn: Optional[Callable] = None   # (input_paths_or_url, output_path, **params) -> extra dict
    params: tuple = ()               # names of extra UI params this target accepts
    action: str = "convert"          # "convert" (batch) or "edit" (opens the editor)
    requires: Optional[str] = None   # Office app needed: "word"|"excel"|"powerpoint"


def _word_to_pdf(inputs, out, **kw):
    from app.office_com import word_to_pdf
    word_to_pdf(inputs[0], out)
    return {}


def _excel_to_pdf(inputs, out, **kw):
    from app.office_com import excel_to_pdf
    excel_to_pdf(inputs[0], out)
    return {}


def _ppt_to_pdf(inputs, out, **kw):
    from app.office_com import ppt_to_pdf
    ppt_to_pdf(inputs[0], out)
    return {}


def _images_to_pdf(inputs, out, **kw):
    from app.convert_images import images_to_pdf
    images_to_pdf(inputs, out)
    return {}


def _url_to_pdf(inputs, out, **kw):
    from app.convert_web import url_to_pdf
    url_to_pdf(inputs[0], out)
    return {}


def _pdf_to_pptx(inputs, out, **kw):
    from app.converter import convert_pdf_to_pptx
    report = convert_pdf_to_pptx(inputs[0], out, progress_callback=kw.get("progress_callback"))
    return {"page_report": _report_to_dict(report)}


def _report_to_dict(report):
    return {
        "page_count": report.page_count,
        "warnings": report.warnings,
        "scanned_warning": report.scanned_warning,
        "substituted_fonts": report.all_substituted_fonts,
        "pages": [
            {"page": p.page_number, "mode": p.mode, "text_boxes": p.text_boxes,
             "tables": p.tables, "images": p.images, "note": p.note}
            for p in report.pages
        ],
    }


def _pdf_to_docx(inputs, out, **kw):
    from app.office_com import pdf_to_docx
    pdf_to_docx(inputs[0], out)
    return {}


def _pdf_to_images(inputs, out, **kw):
    import os
    from app.convert_pdf_misc import pdf_to_images, zip_files
    tmp_dir = out + "_pages"
    fmt = kw.get("format", "png")
    dpi = int(kw.get("dpi", 150))
    paths = pdf_to_images(inputs[0], tmp_dir, fmt=fmt, dpi=dpi)
    zip_files(paths, out)
    return {"page_count": len(paths)}


def _pdf_to_text(inputs, out, **kw):
    from app.convert_pdf_misc import pdf_to_text
    pdf_to_text(inputs[0], out)
    return {}


TARGETS: dict[str, Target] = {}


def _register(t: Target):
    TARGETS[t.id] = t


# ---- TO PDF ---------------------------------------------------------------
_register(Target(
    id="word_to_pdf", label="Word → PDF", group="to_pdf", source_type="word",
    output_ext="pdf", accepts=(".docx", ".doc"), requires="word",
    description="Convert a Word document to PDF via Microsoft Word.",
    fn=_word_to_pdf))

_register(Target(
    id="excel_to_pdf", label="Excel → PDF", group="to_pdf", source_type="excel",
    output_ext="pdf", accepts=(".xlsx",), requires="excel",
    description="Convert a spreadsheet to PDF via Microsoft Excel.",
    fn=_excel_to_pdf))

_register(Target(
    id="ppt_to_pdf", label="PowerPoint → PDF", group="to_pdf", source_type="powerpoint",
    output_ext="pdf", accepts=(".pptx", ".ppt"), requires="powerpoint",
    description="Convert a presentation to PDF via Microsoft PowerPoint.",
    fn=_ppt_to_pdf))

_register(Target(
    id="images_to_pdf", label="Images → PDF", group="to_pdf", source_type="image",
    output_ext="pdf", accepts=(".jpg", ".jpeg", ".png", ".heic", ".heif"), multi=True,
    description="Merge one or more photos into a single PDF, in order.",
    fn=_images_to_pdf))

_register(Target(
    id="url_to_pdf", label="Web page → PDF", group="to_pdf", source_type="url",
    output_ext="pdf", accepts=(), multi=False,
    description="Print a web page to PDF using a local headless browser.",
    fn=_url_to_pdf))

# ---- FROM PDF ---------------------------------------------------------------
_register(Target(
    id="pdf_edit", label="PDF → Edit & export", group="from_pdf", source_type="pdf",
    output_ext="pdf", accepts=(".pdf",), action="edit",
    description="Edit the text directly in the browser, then export a finished PDF."))

_register(Target(
    id="pdf_to_pptx", label="PDF → PowerPoint", group="from_pdf", source_type="pdf",
    output_ext="pptx", accepts=(".pdf",),
    description="Editable slides. Tables, text, and images stay editable.",
    fn=_pdf_to_pptx))

_register(Target(
    id="pdf_to_docx", label="PDF → Word", group="from_pdf", source_type="pdf",
    output_ext="docx", accepts=(".pdf",), requires="word",
    description="Uses Word's own PDF import to rebuild an editable document.",
    fn=_pdf_to_docx))

_register(Target(
    id="pdf_to_images", label="PDF → Images", group="from_pdf", source_type="pdf",
    output_ext="zip", accepts=(".pdf",),
    description="One PNG or JPG per page, at your chosen resolution.",
    fn=_pdf_to_images, params=("format", "dpi")))

_register(Target(
    id="pdf_to_text", label="PDF → Text", group="from_pdf", source_type="pdf",
    output_ext="txt", accepts=(".pdf",),
    description="Plain text, one form-feed between pages.",
    fn=_pdf_to_text))

_register(Target(
    id="pdf_to_excel", label="PDF → Excel", group="from_pdf", source_type="pdf",
    output_ext="xlsx", accepts=(".pdf",), enabled=False, note="Coming soon",
    description="Reliable table extraction to spreadsheets is still in progress."))


def targets_for_type(detected_type: str) -> list[Target]:
    return [t for t in TARGETS.values() if t.source_type == detected_type]


OFFICE_MISSING_NOTE = "Needs Microsoft Office, not found on this PC"


def matrix_json(office: Optional[dict] = None) -> list[dict]:
    """Serialise the matrix. If `office` (a {'word':bool,...} availability map)
    is given, Office-dependent targets are greyed out with an honest note on
    machines where the required app isn't installed."""
    out = []
    for t in TARGETS.values():
        enabled, note = t.enabled, t.note
        if enabled and t.requires and office is not None and not office.get(t.requires, False):
            enabled, note = False, OFFICE_MISSING_NOTE
        out.append(
            {"id": t.id, "label": t.label, "group": t.group, "source_type": t.source_type,
             "output_ext": t.output_ext, "accepts": list(t.accepts), "multi": t.multi,
             "enabled": enabled, "note": note, "description": t.description,
             "params": list(t.params), "action": t.action, "requires": t.requires})
    return out


def is_available(target: "Target", office: Optional[dict]) -> bool:
    if not target.enabled or target.fn is None:
        return False
    if target.requires and office is not None and not office.get(target.requires, False):
        return False
    return True
