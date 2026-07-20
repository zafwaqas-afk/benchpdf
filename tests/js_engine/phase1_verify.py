"""Phase-1 parity check: JS extraction vs the Python engine's extraction.

Bar (from the port brief):
  * extracted line coordinates match the Python engine's within 1pt
  * background renders contain no visible text

Python truth comes from the same _collect_lines the desktop engine feeds its
clustering, so agreement here means the ported clustering sees the same input
the Python clustering does.
"""

import base64
import glob
import http.server
import io
import json
import os
import socketserver
import sys
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", ".."))

import fitz
from app.extraction import _collect_lines

SITE = os.environ.get("BENCHPDF_SITE_DIR") or os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "benchpdf-site")
FIX = os.path.join(HERE, "..", "fixtures")
COORD_TOL = 1.0     # pt


class H(http.server.SimpleHTTPRequestHandler):
    def translate_path(self, path):
        p = path.split("?", 1)[0]
        if p.startswith("/fixtures/"):
            return os.path.join(FIX, p[len("/fixtures/"):])
        if p in ("/", "/harness"):
            return os.path.join(HERE, "extract_harness.html")
        return os.path.join(SITE, p.lstrip("/").replace("/", os.sep))

    def log_message(self, *a):
        pass


def python_lines(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []
    for i in range(doc.page_count):
        lines = _collect_lines(doc[i].get_text("dict"))
        pages.append([{
            "bbox": [round(v, 2) for v in ln["bbox"]],
            "text": "".join(s["text"] for s in ln["spans"]),
            "size": round(ln["size"], 2),
        } for ln in lines])
    doc.close()
    return pages


def match_lines(py, js):
    """Greedy nearest-match by text; report worst coordinate delta."""
    unmatched_py, worst, matched = [], 0.0, 0
    js_pool = list(js)
    for pl in py:
        best, best_d = None, None
        for jl in js_pool:
            if jl["text"].replace(" ", "") != pl["text"].replace(" ", ""):
                continue
            d = max(abs(a - b) for a, b in zip(pl["bbox"], jl["bbox"]))
            if best_d is None or d < best_d:
                best, best_d = jl, d
        if best is None:
            unmatched_py.append(pl["text"][:40])
            continue
        js_pool.remove(best)
        if best_d > worst:
            worst = best_d
            deltas = [round(b - a, 2) for a, b in zip(pl["bbox"], best["bbox"])]
            match_lines.worst_detail = (pl["text"][:30], deltas, pl["bbox"], best["bbox"])
        matched += 1
    return matched, unmatched_py, [j["text"][:40] for j in js_pool], worst


def main():
    from playwright.sync_api import sync_playwright

    srv = socketserver.ThreadingTCPServer(("127.0.0.1", 0), H)
    srv.daemon_threads = True
    port = srv.server_address[1]
    threading.Thread(target=srv.serve_forever, daemon=True).start()

    ok = True
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page()
        pg.goto(f"http://127.0.0.1:{port}/harness")
        pg.wait_for_function("window.harnessReady === true", timeout=30000)

        for src in sorted(glob.glob(os.path.join(FIX, "*.pdf"))):
            name = os.path.splitext(os.path.basename(src))[0]
            py = python_lines(src)
            js = pg.evaluate("u => window.dumpExtraction(u)", "/fixtures/" + os.path.basename(src))
            print(f"\n== {name} ==")
            for i, (pp, jp) in enumerate(zip(py, js)):
                matched, only_py, only_js, worst = match_lines(pp, jp["lines"])
                line_ok = not only_py and not only_js and worst <= COORD_TOL
                ok = ok and line_ok
                print(f"  page {i+1}: py_lines={len(pp):3} js_lines={len(jp['lines']):3} "
                      f"matched={matched:3} worst_delta={worst:5.2f}pt "
                      f"segs={jp['nSegments']:4} [{'PASS' if line_ok else 'FAIL'}]")
                for t in only_py[:4]:
                    print(f"      only-python: {t!r}")
                for t in only_js[:4]:
                    print(f"      only-js:     {t!r}")
                if not line_ok and hasattr(match_lines, "worst_detail"):
                    t, d, pb, jb = match_lines.worst_detail
                    print(f"      worst: {t!r} js-py={d}")
                    print(f"        py={pb} js={jb}")

        # ---- background cleanliness: no glyphs where the text sits ----
        print("\n== background text suppression ==")
        from PIL import Image
        bg_checks = [("tables_charts.pdf", 1), ("slide_deck.pdf", 1),
                     ("form_xobject_statement.pdf", 1)]
        for fixture, pageno in [(f, p) for f, p in bg_checks
                                if os.path.exists(os.path.join(FIX, f))]:
            res = pg.evaluate("([u,n]) => window.dumpBackground(u,n,150)",
                              ["/fixtures/" + fixture, pageno])
            bg = Image.open(io.BytesIO(base64.b64decode(res["bg"].split(",")[1]))).convert("L")
            full = Image.open(io.BytesIO(base64.b64decode(res["full"].split(",")[1]))).convert("L")

            doc = fitz.open(os.path.join(FIX, fixture))
            lines = _collect_lines(doc[pageno - 1].get_text("dict"))
            doc.close()
            z = 150 / 72.0

            def edge_energy(img, bb):
                x0, y0, x1, y1 = (max(int(v * z), 0) for v in bb)
                x1 = min(x1, img.width); y1 = min(y1, img.height)
                if x1 - x0 < 2 or y1 - y0 < 2:
                    return 0
                px = img.crop((x0, y0, x1, y1)).load()
                w, h = x1 - x0, y1 - y0
                e = 0
                for yy in range(h):
                    for xx in range(w - 1):
                        e += abs(px[xx + 1, yy] - px[xx, yy])
                return e / max(w * h, 1)

            worst_ratio = 0.0
            for ln in lines[:40]:
                fe = edge_energy(full, ln["bbox"])
                be = edge_energy(bg, ln["bbox"])
                if fe > 4:            # only boxes where text visibly renders
                    worst_ratio = max(worst_ratio, be / fe)
            clean = worst_ratio < 0.15
            ok = ok and clean
            print(f"  {fixture} p{pageno}: worst bg/full edge-energy ratio in text boxes "
                  f"= {worst_ratio:.3f}  [{'PASS' if clean else 'FAIL'}]")
        b.close()
    srv.shutdown()
    print("\nPHASE 1:", "PASS" if ok else "FAIL")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
