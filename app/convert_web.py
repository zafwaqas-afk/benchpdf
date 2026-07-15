"""
Web page (URL) -> PDF, via headless Microsoft Edge (falls back to Chrome)
using --print-to-pdf. This launches a local browser process on this machine
to render the page and print it - no external conversion API is used, and
nothing about the conversion is sent anywhere except the ordinary HTTP(S)
request the browser itself makes to fetch the page the user asked for.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from urllib.parse import urlparse

# Substrings the browser itself prints when it renders its OWN "can't reach
# this page" screen instead of the requested site - that page still "prints"
# successfully, so we must recognise it and report a clean failure instead of
# silently handing back a PDF of a browser error screen.
_BROWSER_ERROR_MARKERS = (
    "dns_probe_finished", "err_connection", "err_name_not_resolved",
    "err_internet_disconnected", "err_ssl", "can't reach this page",
    "this site can't be reached", "err_cert", "err_address_unreachable",
    "err_empty_response", "err_timed_out",
)

EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
PRINT_TIMEOUT = 45  # seconds


class WebConvertError(Exception):
    pass


def _find_browser() -> str:
    for c in EDGE_CANDIDATES + CHROME_CANDIDATES:
        if os.path.exists(c):
            return c
    raise WebConvertError(
        "No installed browser (Edge or Chrome) was found to render the page.")


def _normalize_url(raw: str) -> str:
    raw = (raw or "").strip()
    if not raw:
        raise WebConvertError("Enter a web address to convert.")
    parsed = urlparse(raw if "://" in raw else f"https://{raw}")
    if parsed.scheme not in ("http", "https"):
        raise WebConvertError("Only http:// and https:// web addresses are supported.")
    if not parsed.netloc:
        raise WebConvertError("That doesn't look like a valid web address.")
    return parsed.geturl()


def url_to_pdf(url: str, dst: str, timeout: int = PRINT_TIMEOUT) -> str:
    url = _normalize_url(url)
    browser = _find_browser()
    dst = os.path.abspath(dst)

    with tempfile.TemporaryDirectory(prefix="pdf2pptx_edge_") as profile_dir:
        cmd = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            f"--user-data-dir={profile_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-extensions",
            "--print-to-pdf-no-header",
            f"--print-to-pdf={dst}",
            "--virtual-time-budget=15000",
            url,
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, timeout=timeout, text=True)
        except subprocess.TimeoutExpired as exc:
            raise WebConvertError(
                "The page took too long to load and print. Check the address "
                "and your network connection, then try again.") from exc

        if not os.path.exists(dst) or os.path.getsize(dst) == 0:
            detail = (proc.stderr or proc.stdout or "").strip().splitlines()
            hint = detail[-1][:200] if detail else "no output was produced"
            raise WebConvertError(
                f"Couldn't load or print that page ({hint}). Check the address and try again.")

        if _looks_like_browser_error_page(dst):
            os.remove(dst)
            raise WebConvertError(
                "That address couldn't be reached. Check the URL and your network "
                "connection, then try again.")
        return dst


def _looks_like_browser_error_page(pdf_path: str) -> bool:
    """The browser 'prints' its own can't-reach-this-page screen just like any
    other page, producing a valid-looking PDF - detect that case so it isn't
    handed back as a successful conversion."""
    try:
        import fitz
        doc = fitz.open(pdf_path)
        text = doc[0].get_text("text").lower() if doc.page_count else ""
        doc.close()
    except Exception:
        return False
    return any(marker in text for marker in _BROWSER_ERROR_MARKERS)
