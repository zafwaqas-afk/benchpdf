"""
BenchPDF desktop entrypoint.

Runs the local Flask server in a background thread, opens the default browser
to it, and shows a system-tray icon (Open / About / Diagnostics / Quit) so the
app has no console window yet a clear way to close it.

Hidden diagnostics: `BenchPDF.exe --diagnostics` runs the fidelity regression
suite against the bundled synthetic fixtures, prints to the parent console (a
GUI exe is re-attached to it), writes a log, and exits with the suite's code.
"""

from __future__ import annotations

import os
import socket
import sys
import threading
import time
import webbrowser


def _resource_base() -> str:
    if getattr(sys, "frozen", False):
        return sys._MEIPASS  # noqa: SLF001
    return os.path.dirname(os.path.abspath(__file__))


def _find_free_port(preferred=5000) -> int:
    for port in (preferred, 5001, 5002, 0):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(("127.0.0.1", port))
            p = s.getsockname()[1]
            s.close()
            return p
        except OSError:
            continue
    return preferred


# --------------------------------------------------------------------------- #
# Diagnostics (hidden): run the regression suite
# --------------------------------------------------------------------------- #
def run_diagnostics() -> int:
    # A windowed (noconsole) exe is detached from the console; re-attach so the
    # user sees output when they run `BenchPDF.exe --diagnostics` from a terminal.
    attached = False
    if getattr(sys, "frozen", False) and os.name == "nt":
        try:
            import ctypes
            if ctypes.windll.kernel32.AttachConsole(-1):
                attached = True
                sys.stdout = open("CONOUT$", "w", encoding="utf-8", buffering=1)
                sys.stderr = sys.stdout
        except Exception:
            pass

    import tempfile
    log_dir = os.path.join(os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), "BenchPDF")
    os.makedirs(log_dir, exist_ok=True)
    os.environ["BENCHPDF_DIAG_WORK"] = tempfile.mkdtemp(prefix="benchpdf_diag_")
    log_path = os.path.join(log_dir, "diagnostics.log")

    base = _resource_base()
    sys.path.insert(0, base)
    sys.path.insert(0, os.path.join(base, "tests"))

    import io as _io
    buf = _io.StringIO()

    class _Tee:
        def __init__(self, *streams): self.streams = [s for s in streams if s]
        def write(self, d):
            for s in self.streams:
                try: s.write(d)
                except Exception: pass
        def flush(self):
            for s in self.streams:
                try: s.flush()
                except Exception: pass

    real_out = sys.stdout
    sys.stdout = _Tee(real_out, buf)
    code = 0
    try:
        import fidelity_suite
        try:
            fidelity_suite.main()
        except SystemExit as e:
            code = int(e.code or 0)
    except Exception as exc:
        print("DIAGNOSTICS ERROR:", exc)
        code = 2
    finally:
        sys.stdout = real_out
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(buf.getvalue())
        except Exception:
            pass

    # If double-clicked (no console, not suppressed), surface a summary box so a
    # non-technical user still sees the result. Skipped when run from a terminal
    # or when BENCHPDF_DIAG_NO_MSGBOX is set (e.g. automated verification).
    if (getattr(sys, "frozen", False) and os.name == "nt"
            and not attached and not os.environ.get("BENCHPDF_DIAG_NO_MSGBOX")):
        try:
            import ctypes
            msg = "Fidelity suite: GREEN" if code == 0 else "Fidelity suite: RED (see log)"
            ctypes.windll.user32.MessageBoxW(0, f"{msg}\n\nLog: {log_path}", "BenchPDF diagnostics", 0)
        except Exception:
            pass
    return code


# --------------------------------------------------------------------------- #
# Normal launch: server thread + browser + tray
# --------------------------------------------------------------------------- #
def _make_icon_image():
    from PIL import Image, ImageDraw
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([10, 6, 54, 58], radius=6, outline=(44, 75, 224, 255), width=4)
    d.line([32, 8, 32, 56], fill=(44, 75, 224, 255), width=4)
    return img


def main() -> int:
    if "--diagnostics" in sys.argv or "--self-test" in sys.argv:
        return run_diagnostics()

    base = _resource_base()
    if base not in sys.path:
        sys.path.insert(0, base)

    from app import server
    from app.office_com import shutdown_worker

    port = _find_free_port(5000)
    url = f"http://127.0.0.1:{port}/"

    srv = threading.Thread(target=server.run_server, kwargs={"port": port}, daemon=True)
    srv.start()

    # wait until the server answers, then open the browser
    def _open_when_ready():
        for _ in range(100):
            try:
                socket.create_connection(("127.0.0.1", port), timeout=0.3).close()
                break
            except OSError:
                time.sleep(0.1)
        webbrowser.open(url)
    threading.Thread(target=_open_when_ready, daemon=True).start()

    def _quit(icon=None, item=None):
        try:
            shutdown_worker()
        except Exception:
            pass
        if icon:
            icon.stop()
        os._exit(0)

    try:
        import pystray
        menu = pystray.Menu(
            pystray.MenuItem("Open BenchPDF", lambda i, it: webbrowser.open(url), default=True),
            pystray.MenuItem("About", lambda i, it: webbrowser.open(url)),
            pystray.MenuItem("Quit", _quit),
        )
        icon = pystray.Icon("BenchPDF", _make_icon_image(), "BenchPDF", menu)
        icon.run()  # blocks on the main thread until Quit
    except Exception:
        # no tray available: keep the server alive in the foreground
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            _quit()
    return 0


if __name__ == "__main__":
    sys.exit(main())
