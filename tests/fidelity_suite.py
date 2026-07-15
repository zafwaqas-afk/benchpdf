"""
Fidelity regression suite for the conversion engine.

Runs with one command (or run-tests.bat) and fails if any placement invariant
regresses on the committed fixtures. Covers the text-PLACEMENT paths, where
alignment/overlap regressions live:

  * PDF -> PPTX          (positioned blocks + native tables)
  * PDF -> edited PDF    (editor export: positioned blocks, byte-identical
                          unedited pages)

plus lightweight structural smoke checks for the other from-PDF paths
(PDF -> images, PDF -> text). Office-COM paths are covered by test_com_crash /
verify_outputs and are intentionally not re-run here (Office layout has no
text-placement fidelity for us to assert, and the suite must run without Office).

Invariants asserted per PPTX slide:
  1. Paragraph-level blocks — no sub-0.3in text box that continues an adjacent
     paragraph (fragmentation).
  2. No fabricated ' / ' line-join artifacts.
  3. Every table the PDF detects becomes a NATIVE PowerPoint table.
  4. Zero text-box insets (all four margins 0) so positioning is baseline-true.
  5. Consistent font mapping — each source font maps to exactly one target.
  6. No two elements overlap by >10% of the smaller element's area.
  7. No element extends past the page boundary.
  8. Text-box positions stay within tolerance of the committed golden layout.

Usage:
  python tests/fidelity_suite.py                 # run the suite
  python tests/fidelity_suite.py --update-golden # re-bless golden layout
"""

import glob
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))

import fitz
from pptx import Presentation
from pptx.util import Emu

from app.converter import convert_pdf_to_pptx
from app.edit_model import EditSession
from app.pdf_edit_export import export_edited_pdf
from app.convert_pdf_misc import pdf_to_images, pdf_to_text

FIX = os.path.join(HERE, "fixtures")
GOLD = os.path.join(HERE, "fidelity", "golden.json")
# Writable work area; overridable so a frozen/packaged build can point it at a
# per-user temp dir instead of the (read-only) app bundle.
WORK = os.environ.get("BENCHPDF_DIAG_WORK") or os.path.join(HERE, "..", "_work", "fidelity")
os.makedirs(WORK, exist_ok=True)

FRAG_MAX_H_IN = 0.30          # boxes shorter than this are fragmentation suspects
OVERLAP_FRAC = 0.10           # >10% of smaller element's area = overlap
BOUND_EPS_IN = 0.06           # allowed slop past a page edge
POS_TOL = 0.02                # golden position drift tolerance (fraction of page)


class Result:
    def __init__(self):
        self.checks = []      # (ok, label)
        self.skips = []

    def check(self, ok, label):
        self.checks.append((bool(ok), label))
        return ok

    def skip(self, label):
        self.skips.append(label)

    @property
    def passed(self):
        return all(ok for ok, _ in self.checks)


def _emu_in(v):
    return v / 914400.0


def slide_elements(slide):
    """Return [(kind, x, y, w, h, text)] for text boxes and native tables."""
    els = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            els.append(("text", sh.left, sh.top, sh.width, sh.height,
                        " ".join(sh.text_frame.text.split())[:24]))
        elif sh.has_table:
            els.append(("table", sh.left, sh.top, sh.width, sh.height, "TABLE"))
    return els


def overlap_frac(a, b):
    ix = max(0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))
    iy = max(0, min(a[2] + a[4], b[2] + b[4]) - max(a[2], b[2]))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    return inter / max(1, min(a[3] * a[4], b[3] * b[4]))


def check_text_insets(prs, res, name):
    bad = 0
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                tf = sh.text_frame
                if any(m not in (0, None) and _emu_in(m or 0) > 0.001
                       for m in (tf.margin_left, tf.margin_right, tf.margin_top, tf.margin_bottom)):
                    bad += 1
    res.check(bad == 0, f"[{name}] text-box insets are zero (offenders: {bad})")


def _first_run_size(sh):
    try:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size is not None:
                    return r.font.size.pt
    except Exception:
        pass
    return None


def check_fragments(prs, res, name):
    """A real fragment is a wrapped line of a paragraph broken into its own box:
    short, sharing the previous box's left edge and font size, a tiny gap below
    it, and the upper text ending mid-sentence (no terminal punctuation). Short
    headings differ (larger/different font size than the body beneath them) and
    complete lines end in punctuation, so neither is flagged."""
    suspects = 0
    for slide in prs.slides:
        tb = []
        for sh in slide.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                x, y, w, h = sh.left, sh.top, sh.width, sh.height
                tb.append((x, y, w, h, " ".join(sh.text_frame.text.split()),
                           _first_run_size(sh)))
        for a in tb:
            if _emu_in(a[3]) >= FRAG_MAX_H_IN:
                continue
            a_ends_clean = a[4][-1:] in ".!?:;)’\"" if a[4] else True
            if a_ends_clean:
                continue
            for b in tb:
                if b is a:
                    continue
                same_left = abs(_emu_in(a[0] - b[0])) < 0.05
                gap = _emu_in(b[1] - (a[1] + a[3]))
                below = -0.05 <= gap <= 0.06
                same_size = (a[5] is not None and b[5] is not None
                             and abs(a[5] - b[5]) <= 0.6)
                if same_left and below and same_size:
                    suspects += 1
                    break
    res.check(suspects == 0, f"[{name}] no fragmented paragraphs (suspect boxes: {suspects})")


def check_slashes(prs, src_pdf, res, name):
    doc = fitz.open(src_pdf)
    src = "\n".join(doc[i].get_text("text") for i in range(doc.page_count))
    doc.close()
    out = "\n".join(sh.text_frame.text for slide in prs.slides for sh in slide.shapes
                    if sh.has_text_frame)
    res.check(out.count(" / ") <= src.count(" / "),
              f"[{name}] no fabricated ' / ' joins (out={out.count(' / ')}, src={src.count(' / ')})")


def check_tables(report, res, name):
    mism = [p.page_number for p in report.pages
            if p.mode != "image-only" and p.tables < _pdf_table_count(report.source_pdf, p.page_number - 1)]
    res.check(not mism, f"[{name}] every PDF table is native (short pages: {mism})")


def _pdf_table_count(src, index):
    try:
        doc = fitz.open(src)
        n = len([t for t in doc[index].find_tables(strategy="lines").tables
                 if t.row_count >= 1 and t.col_count >= 1])
        doc.close()
        return n
    except Exception:
        return 0


def check_fonts(report, res, name):
    mapping = {}
    consistent = True
    for entry in report.all_substituted_fonts:
        if " -> " in entry:
            a, b = entry.split(" -> ", 1)
            if a in mapping and mapping[a] != b:
                consistent = False
            mapping[a] = b
    res.check(consistent, f"[{name}] font mapping consistent (each source -> one target)")


def check_overlap_bounds(prs, res, name):
    SW, SH = prs.slide_width, prs.slide_height
    overlaps, oob = 0, 0
    for si, slide in enumerate(prs.slides):
        els = slide_elements(slide)
        for i in range(len(els)):
            e = els[i]
            if (_emu_in(e[1]) < -BOUND_EPS_IN or _emu_in(e[2]) < -BOUND_EPS_IN
                    or _emu_in(e[1] + e[3] - SW) > BOUND_EPS_IN
                    or _emu_in(e[2] + e[4] - SH) > BOUND_EPS_IN):
                oob += 1
            for j in range(i + 1, len(els)):
                if overlap_frac(e, els[j]) > OVERLAP_FRAC:
                    overlaps += 1
    res.check(overlaps == 0, f"[{name}] no element overlaps >10% (found: {overlaps})")
    res.check(oob == 0, f"[{name}] no element past the page boundary (found: {oob})")


def layout_signature(prs):
    """Normalised text-box positions per slide for golden comparison."""
    SW, SH = prs.slide_width, prs.slide_height
    sig = []
    for slide in prs.slides:
        boxes = [(e[1] / SW, e[2] / SH, e[3] / SW, e[4] / SH)
                 for e in slide_elements(slide) if e[0] == "text"]
        boxes.sort()
        sig.append(boxes)
    return sig


def check_golden(prs, golden, res, name):
    if name not in golden:
        res.skip(f"[{name}] no golden layout yet")
        return
    cur = layout_signature(prs)
    ref = golden[name]
    if len(cur) != len(ref):
        res.check(False, f"[{name}] slide count matches golden ({len(cur)} vs {len(ref)})")
        return
    worst = 0.0
    countmis = 0
    for cs, rs in zip(cur, ref):
        if len(cs) != len(rs):
            countmis += 1
            continue
        for cb, rb in zip(cs, rs):
            worst = max(worst, abs(cb[0] - rb[0]), abs(cb[1] - rb[1]))
    res.check(countmis == 0, f"[{name}] golden box counts per slide match (mismatched slides: {countmis})")
    res.check(worst <= POS_TOL, f"[{name}] text-box drift within {POS_TOL*100:.0f}% (worst: {worst*100:.2f}%)")


# --------------------------------------------------------------------------- #
def run_pptx(update_golden=False):
    print("\n== PDF -> PPTX ==")
    res = Result()
    golden = {}
    if os.path.exists(GOLD):
        golden = json.load(open(GOLD))
    new_golden = dict(golden)
    for src in sorted(glob.glob(os.path.join(FIX, "*.pdf"))):
        name = os.path.splitext(os.path.basename(src))[0]
        out = os.path.join(WORK, name + ".pptx")
        report = convert_pdf_to_pptx(src, out)
        report.source_pdf = src
        prs = Presentation(out)
        check_fragments(prs, res, name)
        check_slashes(prs, src, res, name)
        check_tables(report, res, name)
        check_text_insets(prs, res, name)
        check_fonts(report, res, name)
        check_overlap_bounds(prs, res, name)
        if update_golden:
            new_golden[name] = layout_signature(prs)
        else:
            check_golden(prs, golden, res, name)
    if update_golden:
        os.makedirs(os.path.dirname(GOLD), exist_ok=True)
        json.dump(new_golden, open(GOLD, "w"), indent=1)
        print("  golden layout written to", os.path.relpath(GOLD, HERE))
    return res


def run_editor():
    print("\n== PDF -> edited PDF (editor export) ==")
    res = Result()
    for src in sorted(glob.glob(os.path.join(FIX, "*.pdf"))):
        name = os.path.splitext(os.path.basename(src))[0]
        # 1) no-edit export must be byte-identical everywhere
        out0 = os.path.join(WORK, name + "-noedit.pdf")
        export_edited_pdf(src, {"pages": {}}, out0)
        a = open(src, "rb").read()
        res.check(open(out0, "rb").read()[:len(a)] == a,
                  f"[{name}] no-edit export keeps original bytes")

        # 2) one edited page: edit persists, other pages byte-identical, clean text
        s = EditSession(src)
        model = s.page_model(0)
        blocks = model["blocks"]
        tgt = next((b for b in blocks if b["type"] == "text" and len(b["html"]) > 20), None)
        s.close()
        if tgt is None:
            res.skip(f"[{name}] page 1 has no text block to edit")
            continue
        end = tgt["html"].index("</div>") + 6 if "</div>" in tgt["html"] else len(tgt["html"])
        tgt["html"] = "<div>FIDELITY-EDIT sentinel value.</div>" + tgt["html"][end:]
        out1 = os.path.join(WORK, name + "-edited.pdf")
        export_edited_pdf(src, {"pages": {"0": {"blocks": blocks}}}, out1)
        o1, o2 = fitz.open(src), fitz.open(out1)
        ident = all(o1[i].read_contents() == o2[i].read_contents()
                    for i in range(o1.page_count) if i != 0)
        res.check(ident, f"[{name}] editor: unedited pages byte-identical")
        t = o2[0].get_text("text")
        res.check("FIDELITY-EDIT sentinel value" in " ".join(t.split()),
                  f"[{name}] editor: edited text present (normal spaces)")
        res.check(chr(0xA0) not in t and chr(0xAD) not in t,
                  f"[{name}] editor: no nbsp/soft-hyphen artifacts")
        o1.close(); o2.close()
    return res


def run_other_paths():
    """Structural smoke for the remaining conversion-hub paths that don't have
    text-placement fidelity to assert. Office-COM paths (Word/Excel/PPTX -> PDF,
    PDF -> Word) and the URL path are exercised by their own tests
    (tests/verify_outputs.py, tests/verify_hub.py, tests/test_com_crash.py) and
    are intentionally not re-run here so the suite stays fast and needs no Office
    or network. Those paths don't import the extraction/placement code changed
    in this area, so they can't be affected by a placement regression."""
    print("\n== other hub paths (structural) ==")
    res = Result()
    src = os.path.join(FIX, "text_report.pdf")

    imgs = pdf_to_images(src, os.path.join(WORK, "imgs"), fmt="png", dpi=120)
    doc = fitz.open(src); npages = doc.page_count; doc.close()
    res.check(len(imgs) == npages, f"PDF->images: one image per page ({len(imgs)}/{npages})")

    txt = os.path.join(WORK, "text_report.txt")
    pdf_to_text(src, txt)
    body = open(txt, encoding="utf-8").read()
    res.check("Quarterly Operations Review" in body, "PDF->text: extracted expected content")

    # images -> merged PDF (the TO-PDF image path; Pillow/img2pdf, no Office)
    imgdir = os.path.join(HERE, "input")
    pics = [os.path.join(imgdir, f) for f in ("photo_a.jpg", "photo_b.png")
            if os.path.exists(os.path.join(imgdir, f))]
    if len(pics) >= 2:
        from app.convert_images import images_to_pdf
        merged = os.path.join(WORK, "merged.pdf")
        images_to_pdf(pics, merged)
        d = fitz.open(merged); ok = d.page_count == len(pics); d.close()
        res.check(ok, f"images->PDF: merged {len(pics)} images into {len(pics)}-page PDF")
    else:
        res.skip("images->PDF (sample images not present)")

    res.skip("Office-COM paths (Word/Excel/PPTX->PDF, PDF->Word) — see verify_outputs.py / test_com_crash.py")
    res.skip("Web page->PDF — see verify_hub.py (needs network)")
    return res


def main():
    update = "--update-golden" in sys.argv
    results = [run_pptx(update_golden=update)]
    if not update:
        results.append(run_editor())
        results.append(run_other_paths())

    print("\n" + "=" * 62)
    all_ok = True
    for res in results:
        for ok, label in res.checks:
            print(f"  [{'PASS' if ok else 'FAIL'}] {label}")
            all_ok = all_ok and ok
        for s in res.skips:
            print(f"  [SKIP] {s}")
    total = sum(len(r.checks) for r in results)
    passed = sum(1 for r in results for ok, _ in r.checks if ok)
    print("=" * 62)
    if update:
        print("Golden layout updated. Re-run without --update-golden to verify.")
        return
    print(f"{passed}/{total} checks passed")
    print("FIDELITY SUITE: " + ("GREEN — all invariants hold" if all_ok else "RED — regression detected"))
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
