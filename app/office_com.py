"""
Single COM session manager for Microsoft Office automation.

Design: one dedicated background worker thread owns every Office COM object.
COM's apartment-threading model requires objects to be created and used on the
same thread, so routing every call through one thread avoids marshaling
entirely and gives us the "one session, reused instance" behaviour for free.

Safety net: every call is submitted with a timeout. If Office wedges (e.g. a
password prompt that blocks even with Visible=False, or a corrupt file that
hangs a filter), the watchdog force-kills the underlying Office process and
starts a fresh worker thread, so the app never shows a frozen spinner and
always recovers for the next job. If Office's process dies mid-call for any
reason (crashed, killed externally), the same recovery path applies.

All processing is local; nothing here makes a network call.
"""

from __future__ import annotations

import os
import queue
import threading
import time
from typing import Callable, Optional

import psutil

# --------------------------------------------------------------------------- #
# Office presence detection (no Office launched; just a registry probe)
# --------------------------------------------------------------------------- #
_APP_PROGIDS = {"word": "Word.Application", "excel": "Excel.Application",
                "powerpoint": "PowerPoint.Application"}
_office_cache: Optional[dict] = None
_office_lock = threading.Lock()


def _progid_registered(progid: str) -> bool:
    """True if a COM ProgID (e.g. 'Word.Application') is registered on this PC.
    Reads the registry only, so it never launches Office and never blocks."""
    try:
        import winreg
    except Exception:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"):
            return True
    except OSError:
        return False


def office_availability(force: bool = False) -> dict:
    """Return {'word': bool, 'excel': bool, 'powerpoint': bool}, cached.
    Used to grey out Office-dependent conversions on machines without Office
    so the app never tries (and never crashes) when Office isn't installed."""
    global _office_cache
    with _office_lock:
        if _office_cache is None or force:
            _office_cache = {k: _progid_registered(p) for k, p in _APP_PROGIDS.items()}
        return dict(_office_cache)


CALL_TIMEOUT = 30       # seconds; Office->PDF exports normally take 1-15s
REFLOW_TIMEOUT = 90     # seconds; Word's PDF-import reflow is inherently slower
PROC_NAMES = {
    "Word.Application": "WINWORD.EXE",
    "Excel.Application": "EXCEL.EXE",
    "PowerPoint.Application": "POWERPNT.EXE",
}

WD_ALERTS_NONE = 0
XL_ALERTS_NONE = False  # Excel's DisplayAlerts is a plain bool
PP_ALERTS_NONE = 2      # ppAlertsNone
WD_EXPORT_PDF = 17      # wdExportFormatPDF (Document.ExportAsFixedFormat)
WD_FORMAT_DOCX = 12     # wdFormatXMLDocument (Document.SaveAs2 FileFormat)
XL_EXPORT_PDF = 0       # xlTypePDF (Workbook.ExportAsFixedFormat Type)
PP_EXPORT_PDF = 32      # ppSaveAsPDF


class OfficeError(Exception):
    """Human-readable, user-facing conversion failure."""


class _Job:
    __slots__ = ("fn", "args", "kwargs", "result", "error", "done")

    def __init__(self, fn, args, kwargs):
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.result = None
        self.error = None
        self.done = threading.Event()


class OfficeComWorker:
    """Owns one Office COM session on one dedicated thread."""

    def __init__(self):
        self._lock = threading.Lock()
        self._queue: "queue.Queue[Optional[_Job]]" = queue.Queue()
        self._apps: dict[str, object] = {}
        self._pids: dict[str, int] = {}
        self._thread = None
        self._alive = False
        self._start_thread()

    # -- thread lifecycle ---------------------------------------------------
    def _start_thread(self):
        self._queue = queue.Queue()
        t = threading.Thread(target=self._run, name="office-com-worker", daemon=True)
        self._thread = t
        self._alive = True
        t.start()

    def _run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            while True:
                job = self._queue.get()
                if job is None:
                    break
                try:
                    job.result = job.fn(self, *job.args, **job.kwargs)
                except Exception as exc:  # noqa: BLE001 - report anything back to caller
                    job.error = exc
                finally:
                    job.done.set()
        finally:
            self._quit_all_unsafe()
            pythoncom.CoUninitialize()

    def _restart_after_wedge(self):
        """Called from the CALLER thread when a job timed out. Force-kills any
        tracked Office processes (which unblocks the wedged worker thread's
        pending call, eventually) and starts a brand-new worker so the app
        keeps working immediately, without waiting for the old thread."""
        with self._lock:
            for name, pid in list(self._pids.items()):
                self._kill_pid(pid)
            self._pids.clear()
            self._apps.clear()
            self._alive = False
            self._start_thread()

    @staticmethod
    def _kill_pid(pid: int):
        try:
            p = psutil.Process(pid)
            p.kill()
            p.wait(timeout=5)
        except Exception:
            pass

    def _quit_all_unsafe(self):
        """Runs on the worker thread only (normal shutdown path)."""
        for name, app in list(self._apps.items()):
            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            try:
                app.Quit()
            except Exception:
                pass
        self._apps.clear()
        for name, pid in list(self._pids.items()):
            self._kill_pid(pid)
        self._pids.clear()

    def shutdown(self):
        """Graceful shutdown: ask the worker thread to Quit() every app."""
        if not self._alive:
            return
        self._alive = False
        try:
            self._queue.put(None)
            if self._thread:
                self._thread.join(timeout=10)
        except Exception:
            pass
        # belt-and-braces: kill anything still tracked (thread may be wedged)
        for name, pid in list(self._pids.items()):
            self._kill_pid(pid)
        self._pids.clear()

    # -- submitting work ------------------------------------------------------
    def submit(self, fn: Callable, *args, timeout: float = CALL_TIMEOUT, **kwargs):
        """Run fn(worker, *args, **kwargs) on the worker thread; block up to
        `timeout` seconds. Raises OfficeError on timeout or underlying failure."""
        job = _Job(fn, args, kwargs)
        self._queue.put(job)
        finished = job.done.wait(timeout)
        if not finished:
            self._restart_after_wedge()
            raise OfficeError(
                "Office didn't respond in time. The file may be password-protected, "
                "corrupt, or waiting on a dialog that needs manual input. "
                "The automatic conversion was stopped. Try opening the file in "
                "Office yourself first to confirm it opens cleanly."
            )
        if job.error is not None:
            raise self._friendly(job.error)
        return job.result

    def _friendly(self, exc: Exception) -> OfficeError:
        if isinstance(exc, OfficeError):
            return exc
        msg = str(exc)
        low = msg.lower()
        if "password" in low or "0x800a03ec" in low and "protect" in low:
            return OfficeError(
                "This file is password-protected. Remove the password in its "
                "original app, then convert again.")
        if any(code in low for code in ("800706ba", "800706be", "80010108",
                                        "8001010a", "the rpc server is unavailable")):
            # underlying Office process died / connection severed
            return OfficeError(
                "Office closed unexpectedly while converting this file. "
                "It may be corrupt. Try opening it directly in Office to check, "
                "then convert again.")
        return OfficeError(f"Office couldn't convert this file ({msg.splitlines()[0][:180]}).")

    # -- app acquisition (runs ON the worker thread, inside a submitted fn) --
    def _get_app(self, prog_id: str):
        app = self._apps.get(prog_id)
        if app is not None:
            try:
                _ = app.Visible  # cheap liveness probe
                return app
            except Exception:
                self._apps.pop(prog_id, None)
                self._pids.pop(prog_id, None)

        import win32com.client
        before = {p.pid for p in psutil.process_iter(["pid", "name"])
                  if p.info["name"] == PROC_NAMES.get(prog_id)}
        app = win32com.client.DispatchEx(prog_id)
        if prog_id == "PowerPoint.Application":
            # PowerPoint's Application object refuses Visible=False outright
            # ("Hiding the application window is not allowed"); invisibility
            # is controlled per-document via Presentations.Open(WithWindow=False).
            try:
                app.DisplayAlerts = PP_ALERTS_NONE
            except Exception:
                pass
        else:
            app.Visible = False
            try:
                app.DisplayAlerts = WD_ALERTS_NONE if prog_id == "Word.Application" else False
            except Exception:
                pass

        after = {p.pid for p in psutil.process_iter(["pid", "name"])
                 if p.info["name"] == PROC_NAMES.get(prog_id)}
        new_pids = after - before
        if new_pids:
            self._pids[prog_id] = next(iter(new_pids))

        self._apps[prog_id] = app
        return app


_worker: Optional[OfficeComWorker] = None
_worker_lock = threading.Lock()


def get_worker() -> OfficeComWorker:
    global _worker
    with _worker_lock:
        if _worker is None:
            _worker = OfficeComWorker()
        return _worker


def shutdown_worker():
    global _worker
    with _worker_lock:
        if _worker is not None:
            _worker.shutdown()
            _worker = None


# --------------------------------------------------------------------------- #
# Conversion functions - each body runs ON the worker thread via submit().
# --------------------------------------------------------------------------- #
def _impl_word_to_pdf(worker: OfficeComWorker, src: str, dst: str):
    app = worker._get_app("Word.Application")
    doc = app.Documents.Open(os.path.abspath(src), ReadOnly=True,
                             AddToRecentFiles=False, ConfirmConversions=False)
    try:
        doc.ExportAsFixedFormat(os.path.abspath(dst), 17)  # wdExportFormatPDF
    finally:
        doc.Close(False)
    return dst


def _impl_excel_to_pdf(worker: OfficeComWorker, src: str, dst: str):
    app = worker._get_app("Excel.Application")
    wb = app.Workbooks.Open(os.path.abspath(src), ReadOnly=True,
                            UpdateLinks=0, AddToMru=False)
    try:
        wb.ExportAsFixedFormat(XL_EXPORT_PDF, os.path.abspath(dst))
    finally:
        wb.Close(False)
    return dst


def _impl_ppt_to_pdf(worker: OfficeComWorker, src: str, dst: str):
    app = worker._get_app("PowerPoint.Application")
    deck = app.Presentations.Open(os.path.abspath(src), ReadOnly=True,
                                  Untitled=False, WithWindow=False)
    try:
        deck.SaveAs(os.path.abspath(dst), PP_EXPORT_PDF)
    finally:
        deck.Close()
    return dst


def _impl_pdf_to_docx(worker: OfficeComWorker, src: str, dst: str):
    app = worker._get_app("Word.Application")
    # Word's built-in PDF reflow import: opening a .pdf converts it to a Word
    # document. ConfirmConversions=False suppresses the "this may take a
    # while" prompt that would otherwise block with Visible=False.
    doc = app.Documents.Open(os.path.abspath(src), ConfirmConversions=False,
                             ReadOnly=False, AddToRecentFiles=False)
    try:
        doc.SaveAs2(os.path.abspath(dst), FileFormat=WD_FORMAT_DOCX)
    finally:
        doc.Close(False)
    return dst


def word_to_pdf(src: str, dst: str) -> str:
    return get_worker().submit(_impl_word_to_pdf, src, dst)


def excel_to_pdf(src: str, dst: str) -> str:
    return get_worker().submit(_impl_excel_to_pdf, src, dst)


def ppt_to_pdf(src: str, dst: str) -> str:
    return get_worker().submit(_impl_ppt_to_pdf, src, dst)


def pdf_to_docx(src: str, dst: str) -> str:
    return get_worker().submit(_impl_pdf_to_docx, src, dst, timeout=REFLOW_TIMEOUT)
