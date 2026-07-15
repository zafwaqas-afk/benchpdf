"""
Local file-conversion hub. Runs a Flask server bound to 127.0.0.1 (localhost
only). No file bytes ever leave the machine: Office formats are converted by
driving locally-installed Microsoft Office via COM automation, never a web
API; images and PDFs are processed with local libraries; the one exception
that touches the network is the "web page -> PDF" target, which - like
opening the page in a browser - fetches the URL the user gave it, nothing
else. There is no telemetry and no CDN assets; all CSS/JS/fonts are served
from the local ./static folder.
"""

import atexit
import os
import sys
import threading
import uuid
import webbrowser
import zipfile

from flask import (
    Flask, render_template, request, jsonify, send_file, abort,
)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import registry                                # noqa: E402
from app.office_com import shutdown_worker              # noqa: E402
from app.convert_images import ImageConvertError         # noqa: E402
from app.convert_web import WebConvertError               # noqa: E402
from app.convert_pdf_misc import PdfMiscError             # noqa: E402
from app.office_com import OfficeError                    # noqa: E402
from app.edit_model import EditSession                    # noqa: E402
from app.pdf_edit_export import export_edited_pdf, EditExportError  # noqa: E402

BASE = os.path.dirname(os.path.abspath(__file__))

# Frozen-aware resource + working directories. When PyInstaller-frozen, the
# bundled templates/static live under sys._MEIPASS/app/..., and the writable
# work area must NOT be inside the (possibly read-only) bundle — use a per-user
# temp dir instead.
if getattr(sys, "frozen", False):
    _RES = os.path.join(sys._MEIPASS, "app")           # noqa: SLF001
    TEMPLATES = os.path.join(_RES, "templates")
    STATIC = os.path.join(_RES, "static")
    WORK = os.path.join(os.environ.get("LOCALAPPDATA", os.environ.get("TEMP", ".")),
                        "BenchPDF", "work")
else:
    TEMPLATES = os.path.join(BASE, "templates")
    STATIC = os.path.join(BASE, "static")
    WORK = os.environ.get("BENCHPDF_WORK", os.path.abspath(os.path.join(BASE, "..", "_work")))
os.makedirs(WORK, exist_ok=True)

app = Flask(__name__, static_folder=STATIC, template_folder=TEMPLATES)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB upload cap

BATCHES: dict = {}
BATCHES_LOCK = threading.Lock()

EDITS: dict = {}          # edit_id -> {session, dir, source, output}
EDITS_LOCK = threading.Lock()

KNOWN_ERRORS = (ImageConvertError, WebConvertError, PdfMiscError, OfficeError)


# --------------------------------------------------------------------------- #
# Batch model
# --------------------------------------------------------------------------- #
def _new_batch_dir(batch_id: str) -> str:
    d = os.path.join(WORK, batch_id)
    os.makedirs(d, exist_ok=True)
    return d


def _set_item(batch_id, item_id, **kw):
    with BATCHES_LOCK:
        b = BATCHES.get(batch_id)
        if not b:
            return
        for it in b["items"]:
            if it["id"] == item_id:
                it.update(kw)
                break


def _run_batch(batch_id: str, target: registry.Target, jobs: list, params: dict):
    """jobs: list of (item_id, input_paths_or_url, output_path, display_name)."""
    for item_id, inputs, out_path, display_name in jobs:
        _set_item(batch_id, item_id, status="converting", message="Converting…")

        def progress_cb(cur, total, msg, _item_id=item_id):
            _set_item(batch_id, _item_id, current=cur, total=total, message=msg)

        try:
            extra = target.fn(inputs, out_path, progress_callback=progress_cb, **params)
            size = os.path.getsize(out_path) if os.path.exists(out_path) else 0
            _set_item(batch_id, item_id, status="done", message="Done", output=out_path,
                      size=size, extra=extra or {})
        except KNOWN_ERRORS as exc:
            _set_item(batch_id, item_id, status="error", error=str(exc), message="Failed")
        except Exception as exc:  # unexpected - still a clean, human-readable failure
            _set_item(batch_id, item_id, status="error",
                      error=f"Unexpected error: {exc}", message="Failed")

    with BATCHES_LOCK:
        b = BATCHES.get(batch_id)
        if b:
            b["status"] = "done"


def _output_name(display_name: str, output_ext: str) -> str:
    stem = os.path.splitext(display_name)[0]
    return f"{stem}.{output_ext}"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/matrix")
def api_matrix():
    from app.office_com import office_availability
    office = office_availability()
    return jsonify({"targets": registry.matrix_json(office), "office": office})


@app.route("/api/about")
def api_about():
    from app.version import APP_NAME, VERSION, SOURCE_URL
    from app import licenses
    from app.office_com import office_availability
    return jsonify({
        "name": APP_NAME, "version": VERSION, "source_url": SOURCE_URL,
        "license": "GNU AGPL-3.0",
        "privacy": ("BenchPDF runs entirely on this computer. Office files are converted "
                    "by the copy of Word, Excel, and PowerPoint installed here; PDFs and "
                    "images are handled by local libraries. Your documents are never sent "
                    "to a server. The only feature that reaches the internet is Web page "
                    "to PDF, which fetches the address you type, and nothing else. There "
                    "are no accounts and no telemetry."),
        "office": office_availability(),
        "licenses": licenses.summary_lines(),
    })


@app.route("/api/detect", methods=["POST"])
def api_detect():
    data = request.get_json(silent=True) or {}
    names = data.get("filenames", [])
    types = {n: registry.detect_type(n) for n in names}
    return jsonify({"types": types})


@app.route("/api/convert", methods=["POST"])
def api_convert():
    from app.office_com import office_availability
    target_id = request.form.get("target", "")
    target = registry.TARGETS.get(target_id)
    if target is None or not target.enabled or target.fn is None:
        return jsonify({"error": "That conversion isn't available."}), 400
    if target.requires and not office_availability().get(target.requires, False):
        return jsonify({"error": "This conversion needs Microsoft " + target.requires.capitalize()
                                 + ", which isn't installed on this PC."}), 400

    batch_id = uuid.uuid4().hex[:12]
    batch_dir = _new_batch_dir(batch_id)

    params = {}
    if "dpi" in request.form:
        params["dpi"] = request.form.get("dpi")
    if "format" in request.form:
        params["format"] = request.form.get("format")

    jobs = []
    items = []

    if target.source_type == "url":
        url = (request.form.get("url") or "").strip()
        if not url:
            return jsonify({"error": "Enter a web address to convert."}), 400
        item_id = uuid.uuid4().hex[:8]
        out_name = _output_name(url.replace("://", "_").replace("/", "_")[:60] or "page", target.output_ext)
        out_path = os.path.join(batch_dir, out_name)
        items.append({"id": item_id, "name": url, "target": target.id, "status": "pending",
                      "message": "Queued", "current": 0, "total": 0, "output": None,
                      "size": 0, "error": None, "extra": {}})
        jobs.append((item_id, [url], out_path, url))

    else:
        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "No file was uploaded."}), 400
        for f in files:
            if not f.filename:
                return jsonify({"error": "No file was uploaded."}), 400
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in target.accepts:
                return jsonify({"error": f"{f.filename} isn't a supported input for this conversion."}), 400

        saved_paths = []
        for f in files:
            safe_name = os.path.basename(f.filename)
            p = os.path.join(batch_dir, safe_name)
            if os.path.exists(p):
                stem, ext = os.path.splitext(safe_name)
                p = os.path.join(batch_dir, f"{stem}_{uuid.uuid4().hex[:6]}{ext}")
            f.save(p)
            saved_paths.append(p)

        if target.multi and len(saved_paths) > 1:
            item_id = uuid.uuid4().hex[:8]
            display = f"{len(saved_paths)} images merged"
            out_path = os.path.join(batch_dir, _output_name("merged", target.output_ext))
            items.append({"id": item_id, "name": display, "target": target.id, "status": "pending",
                          "message": "Queued", "current": 0, "total": 0, "output": None,
                          "size": 0, "error": None, "extra": {}})
            jobs.append((item_id, saved_paths, out_path, display))
        else:
            for p in saved_paths:
                item_id = uuid.uuid4().hex[:8]
                name = os.path.basename(p)
                out_path = os.path.join(batch_dir, _output_name(name, target.output_ext))
                items.append({"id": item_id, "name": name, "target": target.id, "status": "pending",
                              "message": "Queued", "current": 0, "total": 0, "output": None,
                              "size": 0, "error": None, "extra": {}})
                jobs.append((item_id, [p], out_path, name))

    with BATCHES_LOCK:
        BATCHES[batch_id] = {"status": "running", "target": target.id, "items": items}

    t = threading.Thread(target=_run_batch, args=(batch_id, target, jobs, params), daemon=True)
    t.start()
    return jsonify({"batch_id": batch_id})


@app.route("/api/batch/<batch_id>")
def api_batch(batch_id):
    with BATCHES_LOCK:
        b = BATCHES.get(batch_id)
        if not b:
            return jsonify({"error": "unknown batch"}), 404
        # strip absolute output paths from the wire payload
        items = []
        for it in b["items"]:
            safe = {k: v for k, v in it.items() if k != "output"}
            safe["has_output"] = bool(it.get("output"))
            items.append(safe)
        return jsonify({"status": b["status"], "target": b["target"], "items": items})


def _find_item(batch_id, item_id):
    with BATCHES_LOCK:
        b = BATCHES.get(batch_id)
        if not b:
            return None
        for it in b["items"]:
            if it["id"] == item_id:
                return it
    return None


@app.route("/api/batch/<batch_id>/download/<item_id>")
def api_download(batch_id, item_id):
    it = _find_item(batch_id, item_id)
    if not it or it.get("status") != "done" or not it.get("output"):
        abort(404)
    return send_file(it["output"], as_attachment=True,
                     download_name=os.path.basename(it["output"]))


@app.route("/api/batch/<batch_id>/download_all")
def api_download_all(batch_id):
    with BATCHES_LOCK:
        b = BATCHES.get(batch_id)
        if not b:
            abort(404)
        outputs = [it["output"] for it in b["items"] if it.get("status") == "done" and it.get("output")]
    if not outputs:
        abort(404)
    zip_path = os.path.join(WORK, batch_id, "all_outputs.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in outputs:
            z.write(p, arcname=os.path.basename(p))
    return send_file(zip_path, as_attachment=True, download_name="converted_files.zip")


@app.route("/api/batch/<batch_id>/open/<item_id>", methods=["POST"])
def api_open(batch_id, item_id):
    it = _find_item(batch_id, item_id)
    if not it or not it.get("output") or not os.path.exists(it["output"]):
        return jsonify({"error": "file not ready"}), 404
    try:
        os.startfile(it["output"])  # noqa: E1101 (Windows only)
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/batch/<batch_id>/open_folder", methods=["POST"])
def api_open_folder(batch_id):
    d = os.path.join(WORK, batch_id)
    if not os.path.isdir(d):
        return jsonify({"error": "not ready"}), 404
    try:
        os.startfile(d)
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# --------------------------------------------------------------------------- #
# In-app PDF editing
# --------------------------------------------------------------------------- #
def _edit(edit_id):
    with EDITS_LOCK:
        return EDITS.get(edit_id)


@app.route("/api/edit/open", methods=["POST"])
def api_edit_open():
    f = request.files.get("file")
    if f is None or not f.filename:
        return jsonify({"error": "No file was uploaded."}), 400
    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Editing is available for PDF files only."}), 400

    edit_id = uuid.uuid4().hex[:12]
    edit_dir = os.path.join(WORK, "edit_" + edit_id)
    os.makedirs(edit_dir, exist_ok=True)
    src = os.path.join(edit_dir, os.path.basename(f.filename))
    f.save(src)
    try:
        session = EditSession(src)
    except Exception as exc:
        return jsonify({"error": f"Couldn't open this PDF for editing ({exc})."}), 400

    with EDITS_LOCK:
        EDITS[edit_id] = {"session": session, "dir": edit_dir, "source": src,
                          "name": os.path.basename(f.filename), "output": None}
    summary = session.summary()
    summary["edit_id"] = edit_id
    summary["name"] = os.path.basename(f.filename)
    return jsonify(summary)


@app.route("/api/edit/<edit_id>/page/<int:index>")
def api_edit_page(edit_id, index):
    e = _edit(edit_id)
    if not e:
        return jsonify({"error": "This editing session has expired. Re-open the PDF."}), 404
    try:
        return jsonify(e["session"].page_model(index))
    except Exception as exc:
        return jsonify({"error": f"Couldn't read page {index + 1} ({exc})."}), 500


@app.route("/api/edit/<edit_id>/bg/<int:index>")
def api_edit_bg(edit_id, index):
    e = _edit(edit_id)
    if not e:
        abort(404)
    try:
        png = e["session"].background_png(index)
    except Exception:
        abort(404)
    from flask import Response
    return Response(png, mimetype="image/png",
                    headers={"Cache-Control": "no-store"})


@app.route("/api/edit/<edit_id>/export", methods=["POST"])
def api_edit_export(edit_id):
    e = _edit(edit_id)
    if not e:
        return jsonify({"error": "This editing session has expired. Re-open the PDF."}), 404
    edits = request.get_json(silent=True) or {}
    stem = os.path.splitext(e["name"])[0]
    out = os.path.join(e["dir"], f"{stem}-edited.pdf")
    try:
        result = export_edited_pdf(e["source"], edits, out)
    except (EditExportError, Exception) as exc:
        return jsonify({"error": f"Couldn't export the edited PDF ({exc})."}), 500
    with EDITS_LOCK:
        e["output"] = out
    return jsonify({"ok": True, "edited_pages": result.get("edited_pages", [])})


@app.route("/api/edit/<edit_id>/pptx", methods=["POST"])
def api_edit_pptx(edit_id):
    """Download-as-PPTX from the edit screen: apply any edits, then run the
    existing PDF->PPTX engine on the edited PDF so the deck reflects the edits."""
    e = _edit(edit_id)
    if not e:
        return jsonify({"error": "This editing session has expired. Re-open the PDF."}), 404
    edits = request.get_json(silent=True) or {}
    stem = os.path.splitext(e["name"])[0]
    edited_pdf = os.path.join(e["dir"], f"{stem}-edited.pdf")
    pptx_out = os.path.join(e["dir"], f"{stem}-edited.pptx")
    try:
        export_edited_pdf(e["source"], edits, edited_pdf)
        from app.converter import convert_pdf_to_pptx
        convert_pdf_to_pptx(edited_pdf, pptx_out)
    except Exception as exc:
        return jsonify({"error": f"Couldn't build the PowerPoint ({exc})."}), 500
    with EDITS_LOCK:
        e["pptx"] = pptx_out
    return jsonify({"ok": True})


@app.route("/api/edit/<edit_id>/download")
def api_edit_download(edit_id):
    e = _edit(edit_id)
    if not e or not e.get("output") or not os.path.exists(e["output"]):
        abort(404)
    return send_file(e["output"], as_attachment=True,
                     download_name=os.path.basename(e["output"]))


@app.route("/api/edit/<edit_id>/download_pptx")
def api_edit_download_pptx(edit_id):
    e = _edit(edit_id)
    if not e or not e.get("pptx") or not os.path.exists(e["pptx"]):
        abort(404)
    return send_file(e["pptx"], as_attachment=True,
                     download_name=os.path.basename(e["pptx"]))


@app.route("/api/edit/<edit_id>/open", methods=["POST"])
def api_edit_open_file(edit_id):
    e = _edit(edit_id)
    if not e or not e.get("output") or not os.path.exists(e["output"]):
        return jsonify({"error": "Export the PDF first."}), 404
    try:
        os.startfile(e["output"])  # noqa: E1101
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


def run_server(port=5000):
    """Run the Flask server (blocking). Used by the packaged entrypoint, which
    manages the browser + tray itself."""
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False, threaded=True)


def main():
    port = int(os.environ.get("PORT", "5000"))
    url = f"http://127.0.0.1:{port}/"
    atexit.register(shutdown_worker)
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    print(f"BenchPDF - running locally at {url}")
    print("Close this window to stop the tool.")
    try:
        run_server(port)
    finally:
        shutdown_worker()


if __name__ == "__main__":
    main()
