"""Phase-3 golden-output comparison.

For each fixture, convert with BOTH engines, export every slide to PNG through
PowerPoint itself (the one renderer whose opinion matters), and compare page
images. Near-identical is quantified: mean absolute pixel difference on
downscaled greyscale pages, plus a structural check that no page pair diverges
wildly (which is what a missing table or background would do).
"""

import glob
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", ".."))

from tests.engines import ENGINES
from PIL import Image

FIX = os.path.join(HERE, "..", "fixtures")
WORK = os.path.join(HERE, "..", "..", "_work", "phase3")
MEAN_TOL = 12.0      # mean |diff| per pixel, 0-255 greyscale
WORST_TOL = 26.0     # worst single page


def export_pngs(pptx_path, out_dir):
    import win32com.client
    os.makedirs(out_dir, exist_ok=True)
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(os.path.abspath(pptx_path),
                                      ReadOnly=True, Untitled=False, WithWindow=False)
        try:
            pres.SaveAs(os.path.abspath(out_dir), 18)   # ppSaveAsPNG
        finally:
            pres.Close()
    finally:
        app.Quit()
    seen = {}
    for f in glob.glob(os.path.join(out_dir, "*.png")) + glob.glob(os.path.join(out_dir, "*.PNG")):
        seen[os.path.normcase(f)] = f
    def slide_no(path):
        import re
        m = re.search(r"(\d+)", os.path.basename(path))
        return int(m.group(1)) if m else 0
    return sorted(seen.values(), key=slide_no)


def page_diff(a_path, b_path):
    a = Image.open(a_path).convert("L").resize((400, 300))
    b = Image.open(b_path).convert("L").resize((400, 300))
    pa, pb = a.load(), b.load()
    total = 0
    for y in range(300):
        for x in range(400):
            total += abs(pa[x, y] - pb[x, y])
    return total / (400 * 300)


def main():
    if os.path.isdir(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK, exist_ok=True)

    py_eng, js_eng = ENGINES["python"], ENGINES["browser"]
    ok = True
    print(f"{'fixture':26} {'pages':>5} {'mean diff':>10} {'worst page':>11}  verdict")
    print("-" * 66)
    for src in sorted(glob.glob(os.path.join(FIX, "*.pdf"))):
        name = os.path.splitext(os.path.basename(src))[0]
        py_out = os.path.join(WORK, name + ".python.pptx")
        js_out = os.path.join(WORK, name + ".browser.pptx")
        py_eng.convert(src, py_out)
        js_eng.convert(src, js_out)
        py_pngs = export_pngs(py_out, os.path.join(WORK, name + "_py"))
        js_pngs = export_pngs(js_out, os.path.join(WORK, name + "_js"))
        if len(py_pngs) != len(js_pngs):
            print(f"{name:26} PAGE COUNT MISMATCH {len(py_pngs)} vs {len(js_pngs)}  [FAIL]")
            ok = False
            continue
        diffs = [page_diff(a, b) for a, b in zip(py_pngs, js_pngs)]
        mean = sum(diffs) / len(diffs)
        worst = max(diffs)
        good = mean <= MEAN_TOL and worst <= WORST_TOL
        ok = ok and good
        print(f"{name:26} {len(diffs):>5} {mean:>10.2f} {worst:>11.2f}  [{'PASS' if good else 'FAIL'}]")
    print("\nPHASE 3 (golden render):", "PASS" if ok else "FAIL")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
