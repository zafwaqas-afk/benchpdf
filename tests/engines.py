"""Conversion engines under test.

The fidelity suite runs the SAME fixtures and the SAME assertions against every
engine that ships. This module is the registry: each engine is a small adapter
exposing convert(src_pdf, out_pptx).

Why this exists
---------------
A browser PDF->PPTX converter shipped on the marketing site without ever being
run through this suite. It failed every invariant the Python engine holds: no
background layer, one text box per line (116 on a page the Python engine renders
as 6), zero native tables where the source had 17, and every font collapsed to
Arial. The suite could not have caught it, because the suite only knew how to
call Python.

An engine that is not in this registry may not be linked from the site.
"""

import base64
import functools
import http.server
import os
import socketserver
import sys
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))

# The site repo holds the vendored browser libraries the JS engine runs against.
# Measure against the real files at the real versions, or the measurement is fiction.
SITE_DIR = os.environ.get("BENCHPDF_SITE_DIR") or os.path.join(
    os.path.dirname(os.path.dirname(HERE)), "benchpdf-site")
JS_ENGINE_DIR = os.path.join(HERE, "js_engine")


class EngineUnavailable(Exception):
    """Raised when an engine cannot run here (missing dependency, not missing quality)."""


# --------------------------------------------------------------------------- #
class PythonEngine:
    name = "python"
    ships = True          # currently linked from the site
    label = "desktop app"

    def available(self):
        return True

    def convert(self, src_pdf, out_pptx):
        from app.converter import convert_pdf_to_pptx
        convert_pdf_to_pptx(src_pdf, out_pptx)
        return out_pptx


# --------------------------------------------------------------------------- #
class _Server(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


class BrowserEngine:
    """The quarantined JS engine, driven in headless Chromium.

    Serves the site's vendored libraries and tests/js_engine/ together, then
    calls the engine through a harness page. This is not wired to any shipped
    page: it exists so the suite can keep a number on the JS engine's quality,
    and so a future port has to clear the same bar before it can be linked.
    """

    name = "browser"
    ships = True          # relinked 2026-07-20 after the port reached parity
    label = "browser (site drop zone)"

    def available(self):
        if not os.path.isdir(os.path.join(SITE_DIR, "assets", "vendor")):
            return False
        try:
            import playwright.sync_api  # noqa: F401
        except ImportError:
            return False
        return os.path.exists(os.path.join(SITE_DIR, "assets", "js", "engine", "engine.js"))

    def _handler(self):
        site, engine = SITE_DIR, JS_ENGINE_DIR

        class H(http.server.SimpleHTTPRequestHandler):
            def translate_path(self, path):
                p = path.split("?", 1)[0].split("#", 1)[0]
                if p in ("/", "/harness.html"):
                    return os.path.join(engine, "harness.html")
                return os.path.join(site, p.lstrip("/").replace("/", os.sep))

            def log_message(self, *a):
                pass

        return H

    def convert(self, src_pdf, out_pptx):
        if not self.available():
            raise EngineUnavailable(
                "needs playwright and the benchpdf-site checkout "
                "(set BENCHPDF_SITE_DIR)")
        from playwright.sync_api import sync_playwright

        srv = _Server(("127.0.0.1", 0), self._handler())
        port = srv.server_address[1]
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        try:
            with sync_playwright() as p:
                b = p.chromium.launch()
                pg = b.new_page()
                errors = []
                pg.on("pageerror", lambda e: errors.append(str(e)))
                pg.goto(f"http://127.0.0.1:{port}/harness.html", wait_until="load")
                pg.wait_for_function("window.harnessReady === true", timeout=30000)
                data = list(open(src_pdf, "rb").read())
                b64 = pg.evaluate("bytes => window.convertToPptx(bytes)", data)
                b.close()
            if not b64:
                raise EngineUnavailable("harness returned nothing: " + "; ".join(errors))
            with open(out_pptx, "wb") as f:
                f.write(base64.b64decode(b64))
            return out_pptx
        finally:
            srv.shutdown()
            srv.server_close()


ENGINES = {e.name: e for e in (PythonEngine(), BrowserEngine())}
