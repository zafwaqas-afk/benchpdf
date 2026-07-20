"""Corpus fidelity run: convert every corpus PDF with the browser engine,
render both sides, score structural similarity per page, report.

Scores land in corpus_scores.json. The first run writes corpus_baseline.json;
later runs FAIL if the corpus median regresses by more than 0.01 against the
baseline. The ten worst documents are the fidelity backlog, printed each run.
"""

import glob
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", ".."))
DOCS = os.path.join(HERE, "docs")
WORK = os.path.join(HERE, "..", "..", "_work", "corpus")
SCORES = os.path.join(HERE, "corpus_scores.json")
BASELINE = os.path.join(HERE, "corpus_baseline.json")
SZ = (320, 240)


def grey(img):
    from PIL import Image
    return img.convert("L").resize(SZ)


def ssim(a, b):
    pa, pb = a.load(), b.load()
    n = SZ[0] * SZ[1]
    sa = sb = saa = sbb = sab = 0.0
    for y in range(SZ[1]):
        for x in range(SZ[0]):
            va, vb = pa[x, y], pb[x, y]
            sa += va; sb += vb; saa += va * va; sbb += vb * vb; sab += va * vb
    ma, mb = sa / n, sb / n
    va_, vb_ = saa / n - ma * ma, sbb / n - mb * mb
    cov = sab / n - ma * mb
    c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
    return ((2 * ma * mb + c1) * (2 * cov + c2)) / ((ma * ma + mb * mb + c1) * (va_ + vb_ + c2))


def main():
    import fitz
    from PIL import Image
    import win32com.client
    from tests.engines import ENGINES

    eng = ENGINES["browser"]
    os.makedirs(WORK, exist_ok=True)
    pdfs = sorted(glob.glob(os.path.join(DOCS, "*.pdf")))
    print(f"corpus run: {len(pdfs)} documents")

    # 1) convert everything first, one browser per doc
    for i, src in enumerate(pdfs):
        name = os.path.splitext(os.path.basename(src))[0]
        out = os.path.join(WORK, name + ".pptx")
        if os.path.exists(out):
            continue
        try:
            eng.convert(src, out)
        except Exception as e:
            print(f"  CONVERT FAIL {name}: {type(e).__name__}")
        if (i + 1) % 20 == 0:
            print(f"  converted {i + 1}/{len(pdfs)}")

    # 2) render all PPTX through ONE PowerPoint instance
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for src in pdfs:
            name = os.path.splitext(os.path.basename(src))[0]
            pptx = os.path.join(WORK, name + ".pptx")
            outdir = os.path.join(WORK, name + "_png")
            if not os.path.exists(pptx) or os.path.isdir(outdir):
                continue
            try:
                pres = app.Presentations.Open(os.path.abspath(pptx), ReadOnly=True,
                                              Untitled=False, WithWindow=False)
                try:
                    pres.SaveAs(os.path.abspath(outdir), 18)
                finally:
                    pres.Close()
            except Exception as e:
                print(f"  RENDER FAIL {name}: {type(e).__name__}")
    finally:
        app.Quit()

    # 3) score: SSIM of PPTX page renders vs source page renders
    scores = {}
    for src in pdfs:
        name = os.path.splitext(os.path.basename(src))[0]
        outdir = os.path.join(WORK, name + "_png")
        if not os.path.isdir(outdir):
            scores[name] = {"pages": [], "mean": 0.0, "error": "no output"}
            continue
        pngs = sorted(glob.glob(os.path.join(outdir, "*.PNG")) +
                      glob.glob(os.path.join(outdir, "*.png")))
        seen = {os.path.normcase(f): f for f in pngs}
        pngs = sorted(seen.values(), key=lambda f: int("".join(c for c in os.path.basename(f) if c.isdigit()) or 0))
        doc = fitz.open(src)
        page_scores = []
        for i in range(min(doc.page_count, len(pngs))):
            pix = doc[i].get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
            src_img = grey(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
            out_img = grey(Image.open(pngs[i]))
            page_scores.append(round(ssim(src_img, out_img), 4))
        doc.close()
        scores[name] = {"pages": page_scores,
                        "mean": round(sum(page_scores) / len(page_scores), 4) if page_scores else 0.0}

    means = sorted(s["mean"] for s in scores.values())
    median = means[len(means) // 2] if means else 0.0
    json.dump({"median": median, "docs": scores}, open(SCORES, "w"), indent=1)

    print(f"\ncorpus median SSIM: {median:.4f} over {len(means)} documents")
    worst = sorted(scores.items(), key=lambda kv: kv[1]["mean"])[:10]
    print("10 worst (the fidelity backlog):")
    for name, s in worst:
        print(f"  {s['mean']:.4f}  {name}  {s.get('error', '')}")

    if not os.path.exists(BASELINE):
        json.dump({"median": median}, open(BASELINE, "w"))
        print(f"baseline written: median {median:.4f}")
        return 0
    base = json.load(open(BASELINE))["median"]
    if median < base - 0.01:
        print(f"REGRESSION: median {median:.4f} vs baseline {base:.4f}")
        return 1
    print(f"vs baseline {base:.4f}: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
