"""Phase-2 parity: JS placement structures vs the Python engine's.

Two comparisons per fixture page:
  * tables: JS detectTables vs fitz find_tables(strategy="lines"):
    same count, bboxes within 2pt, same row/col counts
  * clusters: JS clusterLines over loose lines vs the Python engine's
    _cluster_lines over ITS loose lines: same count, bboxes within 2pt

Agreement here means the writer receives the same shapes the desktop engine
places, which is what the golden-layout suite ultimately asserts.
"""

import glob
import http.server
import os
import socketserver
import sys
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", ".."))

import fitz
from app.extraction import (_collect_lines, _attach_markers, _cluster_lines,
                            _center, _point_in, _infer_aligned_tables)

SITE = os.environ.get("BENCHPDF_SITE_DIR") or os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "benchpdf-site")
FIX = os.path.join(HERE, "..", "fixtures")
BBOX_TOL = 2.0


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


def python_placement(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []
    for i in range(doc.page_count):
        page = doc[i]
        lines = _collect_lines(page.get_text("dict"))
        try:
            found = page.find_tables(strategy="lines")
            tables = [t for t in found.tables if t.row_count >= 1 and t.col_count >= 1]
        except Exception:
            tables = []
        # mirror the engine pipeline: text-less grids are demoted to graphics,
        # unruled tables are recovered by column-alignment inference
        tables = [t for t in tables
                  if any(_point_in(t.bbox, *_center(ln["bbox"])) for ln in lines)]
        tables = tables + _infer_aligned_tables(lines, [t.bbox for t in tables])
        tinfo = [{"bbox": [round(v, 1) for v in t.bbox],
                  "rows": t.row_count, "cols": t.col_count} for t in tables]
        tb = [t.bbox for t in tables]
        loose = [ln for ln in lines if not any(_point_in(b, *_center(ln["bbox"])) for b in tb)]
        loose = _attach_markers(loose)
        clusters = _cluster_lines(loose)
        cinfo = [{"bbox": [round(min(c["x0"] for c in cl), 1), round(min(c["y0"] for c in cl), 1),
                           round(max(c["x1"] for c in cl), 1), round(max(c["y1"] for c in cl), 1)],
                  "nLines": len(cl),
                  "text": " ".join("".join(s["text"] for s in ln["spans"]) for ln in cl)[:40]}
                 for cl in clusters]
        pages.append({"tables": tinfo, "clusters": cinfo})
    doc.close()
    return pages


def cmp_boxes(pys, jss, tol):
    """Greedy nearest bbox matching; returns (matched, py_only, js_only, worst)."""
    js_pool = list(jss)
    worst, matched, py_only = 0.0, 0, []
    for p in pys:
        best, best_d = None, None
        for j in js_pool:
            d = max(abs(a - b) for a, b in zip(p["bbox"], j["bbox"]))
            if best_d is None or d < best_d:
                best, best_d = j, d
        if best is None or best_d > tol:
            py_only.append(p)
            continue
        js_pool.remove(best)
        worst = max(worst, best_d)
        matched += 1
    return matched, py_only, js_pool, worst


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
            py = python_placement(src)
            js = pg.evaluate("u => window.dumpPlacement(u)", "/fixtures/" + os.path.basename(src))
            print(f"\n== {name} ==")
            for i, (pp, jp) in enumerate(zip(py, js)):
                tm, tpo, tjo, tworst = cmp_boxes(pp["tables"], jp["tables"], BBOX_TOL)
                grid_ok = all(
                    any(pt["rows"] == jt["rows"] and pt["cols"] == jt["cols"]
                        and max(abs(a - b) for a, b in zip(pt["bbox"], jt["bbox"])) <= BBOX_TOL
                        for jt in jp["tables"])
                    for pt in pp["tables"])
                cm, cpo, cjo, cworst = cmp_boxes(pp["clusters"], jp["clusters"], BBOX_TOL)
                page_ok = (not tpo and not tjo and grid_ok and not cpo and not cjo)
                ok = ok and page_ok
                print(f"  page {i+1}: tables py={len(pp['tables'])} js={len(jp['tables'])} "
                      f"(worst {tworst:4.1f}pt, grids {'ok' if grid_ok else 'MISMATCH'}) | "
                      f"clusters py={len(pp['clusters'])} js={len(jp['clusters'])} "
                      f"(worst {cworst:4.1f}pt) [{'PASS' if page_ok else 'FAIL'}]")
                for t in tpo[:3]:
                    print(f"      table only-py: {t}")
                for t in tjo[:3]:
                    print(f"      table only-js: {t}")
                for c in cpo[:3]:
                    print(f"      cluster only-py: {c['text']!r} {c['bbox']}")
                for c in cjo[:3]:
                    print(f"      cluster only-js: {c['text']!r} {c['bbox']}")
        b.close()
    srv.shutdown()
    print("\nPHASE 2 (structures):", "PASS" if ok else "FAIL")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
