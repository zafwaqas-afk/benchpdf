"""Assemble the fidelity corpus: 100+ diverse PDFs, public or synthetic only.

90 synthetic documents across six classes (statements, invoices, proposals,
two-column reports, slide exports, forms), deterministic by seed, plus a small
set of public PDFs (arXiv papers, US government forms) trimmed to two pages.
Nothing private ever enters this directory.
"""

import os
import random
import urllib.request

import fitz

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(HERE, "docs")

PUBLIC = [
    ("arxiv_1706_03762.pdf", "https://arxiv.org/pdf/1706.03762v7"),
    ("arxiv_1810_04805.pdf", "https://arxiv.org/pdf/1810.04805v2"),
    ("arxiv_2005_14165.pdf", "https://arxiv.org/pdf/2005.14165v4"),
    ("arxiv_1512_03385.pdf", "https://arxiv.org/pdf/1512.03385v1"),
    ("arxiv_1409_0473.pdf", "https://arxiv.org/pdf/1409.0473v7"),
    ("arxiv_2010_11929.pdf", "https://arxiv.org/pdf/2010.11929v2"),
    ("irs_f1040.pdf", "https://www.irs.gov/pub/irs-pdf/f1040.pdf"),
    ("irs_fw9.pdf", "https://www.irs.gov/pub/irs-pdf/fw9.pdf"),
    ("irs_fw4.pdf", "https://www.irs.gov/pub/irs-pdf/fw4.pdf"),
    ("irs_f4506t.pdf", "https://www.irs.gov/pub/irs-pdf/f4506t.pdf"),
]


def synth_statement(page, rng):
    W = page.rect.width
    page.insert_text((40, 60), "EXAMPLE BANK %d" % rng.randrange(9), fontname="hebo", fontsize=15)
    cols = [40, 110, 320, 400, 480, W - 40]
    top, rh, n = 100, rng.choice([16, 18, 20]), rng.randrange(20, 30)
    for r in range(n + 2):
        page.draw_line(fitz.Point(40, top + r * rh), fitz.Point(W - 40, top + r * rh), width=0.7)
    for x in cols:
        page.draw_line(fitz.Point(x, top), fitz.Point(x, top + rh * (n + 1)), width=0.7)
    bal = rng.uniform(500, 9000)
    for r in range(n):
        y = top + (r + 1) * rh + 12
        amt = rng.uniform(2, 300); bal -= amt
        for ci, t in enumerate(["%02d JAN" % (r % 28 + 1), "PAYEE %03d" % rng.randrange(400),
                                "%.2f" % amt, "", "%.2f" % bal]):
            if t: page.insert_text((cols[ci] + 4, y), t, fontname="helv", fontsize=8)


def synth_invoice(page, rng):
    page.insert_text((40, 70), "INVOICE %05d" % rng.randrange(99999), fontname="hebo", fontsize=22)
    page.insert_text((40, 100), "Fabricated Supplier Ltd", fontname="helv", fontsize=10)
    top, rh, n = 160, 24, rng.randrange(6, 14)
    for r in range(n + 2):
        page.draw_line(fitz.Point(40, top + r * rh), fitz.Point(555, top + r * rh), width=0.8)
    for x in (40, 300, 380, 460, 555):
        page.draw_line(fitz.Point(x, top), fitz.Point(x, top + rh * (n + 1)), width=0.8)
    total = 0
    for r in range(n):
        y = top + (r + 1) * rh + 15
        q, up = rng.randrange(1, 9), rng.uniform(5, 400)
        total += q * up
        for ci, t in enumerate([f"Fabricated item {r + 1}", str(q), "%.2f" % up, "%.2f" % (q * up)]):
            page.insert_text(((40, 300, 380, 460)[ci] + 6, y), t, fontname="helv", fontsize=9)
    page.insert_text((380, top + rh * (n + 1) + 24), "Total  %.2f" % total, fontname="hebo", fontsize=11)


def synth_proposal(page, rng):
    for _ in range(2 if rng.random() < 0.5 else 1):
        page.insert_text((60, 90), "PROPOSAL %d" % rng.randrange(90), fontname="hebo", fontsize=26)
    y = 150
    for pnum in range(rng.randrange(2, 4)):
        for li in range(rng.randrange(2, 5)):
            page.insert_text((60, y), f"Fabricated paragraph {pnum} line {li} with several words of body copy.",
                             fontname="helv", fontsize=11)
            y += 11 * rng.choice([1.4, 1.7, 2.0])
        y += 14
    for ci, cx in enumerate((60, 318)):
        page.draw_rect(fitz.Rect(cx, y, cx + 217, y + 90), width=1)
        page.insert_text((cx + 12, y + 24), f"Card {ci + 1}", fontname="hebo", fontsize=12)
        page.insert_text((cx + 12, y + 44), "One line of fabricated detail.", fontname="helv", fontsize=9)


def synth_twocol(page, rng):
    page.insert_text((50, 70), "Two Column Report %d" % rng.randrange(50), fontname="hebo", fontsize=16)
    for cx in (50, 310):
        y = 110
        for i in range(rng.randrange(18, 28)):
            page.insert_text((cx, y), f"Column text line {i} with fabricated words here.",
                             fontname="tiro" if cx > 200 else "helv", fontsize=9)
            y += 13


def synth_slide(page, rng):
    page.draw_rect(fitz.Rect(0, 0, page.rect.width, page.rect.height),
                   fill=(0.96, 0.94, 0.90), width=0)
    page.insert_text((60, 120), "Slide Title %d" % rng.randrange(30), fontname="hebo", fontsize=34)
    y = 190
    for i in range(rng.randrange(3, 6)):
        page.insert_text((80, y), chr(0x2022) + f" Fabricated bullet point number {i + 1}",
                         fontname="helv", fontsize=16)
        y += 34
    page.draw_rect(fitz.Rect(420, 300, 700, 480), fill=(0.75, 0.78, 0.82), width=0)


def synth_form(page, rng):
    page.insert_text((40, 60), "FABRICATED FORM %02d" % rng.randrange(40), fontname="hebo", fontsize=14)
    y = 100
    for i in range(rng.randrange(8, 14)):
        page.insert_text((40, y), f"Field {i + 1}:", fontname="helv", fontsize=9)
        page.draw_line(fitz.Point(120, y + 2), fitz.Point(400, y + 2), width=0.6)
        if rng.random() < 0.4:
            page.draw_rect(fitz.Rect(420, y - 8, 432, y + 4), width=0.8)
            page.insert_text((440, y), "Tick", fontname="helv", fontsize=8)
        y += 26


CLASSES = [("statement", synth_statement, (595, 842)), ("invoice", synth_invoice, (595, 842)),
           ("proposal", synth_proposal, (595, 842)), ("twocol", synth_twocol, (595, 842)),
           ("slide", synth_slide, (792, 612)), ("form", synth_form, (595, 842))]


def main():
    os.makedirs(DOCS, exist_ok=True)
    for name, fn, (w, h) in CLASSES:
        for seed in range(15):
            path = os.path.join(DOCS, f"{name}_{seed:02d}.pdf")
            if os.path.exists(path):
                continue
            rng = random.Random(hash((name, seed)) & 0xffffffff)
            d = fitz.open()
            for _ in range(rng.randrange(1, 3)):
                fn(d.new_page(width=w, height=h), rng)
            d.save(path); d.close()
    for fname, url in PUBLIC:
        path = os.path.join(DOCS, fname)
        if os.path.exists(path):
            continue
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "BenchPDF-corpus/1.0"})
            data = urllib.request.urlopen(req, timeout=60).read()
            d = fitz.open(stream=data, filetype="pdf")
            out = fitz.open()
            out.insert_pdf(d, from_page=0, to_page=min(1, d.page_count - 1))
            out.save(path); out.close(); d.close()
        except Exception as e:
            print("skip", fname, type(e).__name__)
    n = len([f for f in os.listdir(DOCS) if f.endswith(".pdf")])
    print(f"corpus: {n} documents in {DOCS}")


if __name__ == "__main__":
    main()
