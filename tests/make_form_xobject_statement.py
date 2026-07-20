"""Generate the SYNTHETIC form-XObject statement fixture.

Guards the bug class found in a real statement conversion on 2026-07-20: text
living inside (nested) Form XObjects with scaling matrices. The browser
engine's operator walk was not composing the full transform stack, so footer
small print extracted as zero-height boxes piled at the page origin, with
matrix-corrupted font sizes and every span falling back to Arial.

The forms deliberately use different base fonts from the page body: grafting
the same base-14 font into the page merges font dictionaries into a degenerate
one whose widths PDF readers resolve differently, which is a separate quirk
this fixture is not about.

Constructs exercised, all on one page:
  * a dense ruled transaction table (normal content, must keep working)
  * a logo image (embedded raster)
  * footer small print INSIDE a Form XObject, placed scaled at page bottom
  * a second, NESTED Form XObject (form within form) with its own matrix

Synthetic only: every value is fabricated here. Never commit a real statement.
"""

import os
import random

import fitz

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "fixtures", "form_xobject_statement.pdf")


def footer_form_doc():
    """A little document whose page becomes the footer Form XObject.

    Includes space-only and trailing-space text ops: the constructs that, in
    the original bug, made an un-composed operator walk synthesize phantom
    zero-height spans at garbage positions once a scaling form matrix was in
    play."""
    d = fitz.open()
    p = d.new_page(width=400, height=60)
    p.insert_text((10, 18),
                  "Example Bank plc is authorised by the Fabricated Regulation Authority",
                  fontname="tiro", fontsize=9)
    p.insert_text((305, 18), "   ", fontname="tiro", fontsize=9)   # space-only op
    p.insert_text((10, 4), "Terms apply ", fontname="tiro", fontsize=9)  # trailing space
    p.insert_text((10, 32),
                  "and regulated under fabricated registration number 000000.",
                  fontname="tiro", fontsize=9)
    p.insert_text((10, 50),
                  "Registered office: 1 Example Street, Example City EX1 1EX.",
                  fontname="tiro", fontsize=9)
    return d


def nested_form_doc(inner):
    """A document that itself shows the footer form, scaled: form in form."""
    d = fitz.open()
    p = d.new_page(width=300, height=50)
    p.insert_text((6, 14), "Sheet 1 of 1", fontname="cobo", fontsize=10)
    p.insert_text((70, 14), "  ", fontname="cour", fontsize=10)    # space-only op in nested form
    # inner form drawn scaled into the right side of this form
    p.show_pdf_page(fitz.Rect(80, 5, 295, 45), inner, 0)
    return d


def main():
    rng = random.Random(20260721)
    doc = fitz.open()
    W, H = 595, 842
    page = doc.new_page(width=W, height=H)
    left, right = 40, W - 40

    # logo: a generated raster, no external file
    logo = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 120, 40))
    logo.clear_with(230)
    for x in range(120):
        for y in range(40):
            if (x // 10 + y // 10) % 2 == 0:
                logo.set_pixel(x, y, (40, 40, 40))
    page.insert_image(fitz.Rect(left, 40, left + 90, 70), pixmap=logo)
    page.insert_text((left + 110, 62), "EXAMPLE BANK", fontname="hebo", fontsize=15)

    # a HEADER form at the top of the page: on unfixed code the operator walk
    # desyncs here and corrupts everything below it on the page
    hdr = fitz.open()
    hp = hdr.new_page(width=200, height=24)
    hp.insert_text((4, 16), "Statement of account ", fontname="tiro", fontsize=11)
    hp.insert_text((120, 16), "  ", fontname="tiro", fontsize=11)
    page.show_pdf_page(fitz.Rect(right - 180, 42, right, 66), hdr, 0)
    hdr.close()

    # dense ruled table
    cols = [left, 110, 330, 405, 480, right]
    top = 100
    row_h = 18
    n_rows = 30
    bottom = top + row_h * (n_rows + 1)
    for r in range(n_rows + 2):
        yy = top + r * row_h
        page.draw_line(fitz.Point(left, yy), fitz.Point(right, yy), width=0.7)
    for x in cols:
        page.draw_line(fitz.Point(x, top), fitz.Point(x, bottom), width=0.7)
    page.draw_rect(fitz.Rect(left, top, right, top + row_h), fill=(0.92, 0.92, 0.92), width=0)
    for ci, htext in enumerate(["Date", "Description", "Money out", "Money in", "Balance"]):
        page.insert_text((cols[ci] + 4, top + 13), htext, fontname="hebo", fontsize=8)
    balance = 2500.00
    for r in range(n_rows):
        yy = top + (r + 1) * row_h + 13
        amt = round(rng.uniform(2.5, 180.0), 2)
        balance -= amt
        page.insert_text((cols[0] + 4, yy), "%02d JUN" % (r % 28 + 1), fontname="helv", fontsize=8)
        page.insert_text((cols[1] + 4, yy), "FABRICATED PAYEE %02d" % (r % 12), fontname="helv", fontsize=8)
        page.insert_text((cols[2] + 4, yy), "%.2f" % amt, fontname="helv", fontsize=8)
        page.insert_text((cols[4] + 4, yy), "%.2f" % balance, fontname="helv", fontsize=8)

    # footer small print: nested Form XObjects, scaled, anchored at page bottom
    inner = footer_form_doc()
    nested = nested_form_doc(inner)
    # outer placement scales the nested form down: matrix-scaled small print
    page.show_pdf_page(fitz.Rect(left, H - 70, left + 420, H - 20), nested, 0)

    doc.save(OUT)
    for d in (inner, nested, doc):
        d.close()
    print("wrote", OUT)


if __name__ == "__main__":
    main()
