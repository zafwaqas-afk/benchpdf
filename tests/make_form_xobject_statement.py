"""Generate the SYNTHETIC statement fixture for the statement bug classes.

Guards the fault classes found in real statement conversions (2026-07-20/21):

  (a) header/footer text inside (nested) Form XObjects with scaling matrices:
      an un-composed operator walk extracts zero-height boxes piled at the
      origin. ALSO: a line whose visible ink is filled glyph-OUTLINE paths
      with an invisible text layer on top (the Starling pattern: Type3
      charprocs that are bare d1 metrics). An engine that treats the outlines
      as vector art paints the word into the background raster AND re-emits
      the invisible text editable: ghost-doubled text. The outline colour is
      the word's real colour; the invisible layer says black.
  (b) an UNRULED transaction ledger: no ruling lines at all, so a lines-only
      table detector ships every transaction as loose text. Column alignment
      is the only signal.
  (c) a decorative 8-square brand mark with no text: a grid detector that
      promotes it emits an empty 1x8 native table.
  (d) a narrow multi-word column header ("END OF DAY" / "ACCOUNT BALANCE"):
      sized to source width, a substituted font wraps it mid-word.
  (e) grey and brand-coloured text runs: a converter that flattens every run
      to #000000 passes every geometry check while destroying the palette.

Synthetic only: every value is fabricated here. Never commit a real statement.
"""

import os
import random

import fitz

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "fixtures", "form_xobject_statement.pdf")

BRAND_BLUE = (0.10, 0.28, 0.60)      # heading colour
BRAND_PURPLE = (0.196, 0.118, 0.216)  # the outline-text colour (#321E37-ish)
GREY = (0.45, 0.45, 0.45)


# --------------------------------------------------------------------------- #
# Text drawn as filled glyph-outline paths (fontTools) + invisible text layer
# --------------------------------------------------------------------------- #
def _glyph_commands(text, fontsize):
    """[(advance_x, [contour commands])] per glyph, y-up glyph space in pt."""
    from fontTools.ttLib import TTFont
    from fontTools.pens.basePen import BasePen

    class Pen(BasePen):
        def __init__(self, glyphset):
            super().__init__(glyphset)
            self.cmds = []

        def _moveTo(self, p):
            self.cmds.append(("m", p))

        def _lineTo(self, p):
            self.cmds.append(("l", p))

        def _curveToOne(self, c1, c2, p):
            self.cmds.append(("c", c1, c2, p))

        def _qCurveToOne(self, cp, p):
            p0 = self._getCurrentPoint()
            c1 = (p0[0] + 2 / 3 * (cp[0] - p0[0]), p0[1] + 2 / 3 * (cp[1] - p0[1]))
            c2 = (p[0] + 2 / 3 * (cp[0] - p[0]), p[1] + 2 / 3 * (cp[1] - p[1]))
            self.cmds.append(("c", c1, c2, p))

        def _closePath(self):
            self.cmds.append(("h",))

    f = TTFont(os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arialbd.ttf"))
    upm = f["head"].unitsPerEm
    scale = fontsize / upm
    glyphset = f.getGlyphSet()
    cmap = f.getBestCmap()
    hmtx = f["hmtx"]
    out = []
    for ch in text:
        gname = cmap.get(ord(ch))
        if gname is None:
            out.append((fontsize * 0.5, []))
            continue
        pen = Pen(glyphset)
        glyphset[gname].draw(pen)
        adv = hmtx[gname][0] * scale
        cmds = []
        for c in pen.cmds:
            cmds.append((c[0],) + tuple((x * scale, y * scale) for x, y in
                                        [p for p in c[1:]]))
        out.append((adv, cmds))
    return out


def draw_outline_text(page, origin, text, fontsize, color):
    """Paint text as filled glyph-outline paths with an INVISIBLE (render
    mode 3) text layer on top - the statement-generator pattern. The visible
    colour lives only on the paths; the invisible layer is black."""
    x, base_y = origin
    shape = page.new_shape()
    for adv, cmds in _glyph_commands(text, fontsize):
        cur = None
        for c in cmds:
            if c[0] == "m":
                cur = fitz.Point(x + c[1][0], base_y - c[1][1])
                start = cur
            elif c[0] == "l":
                p = fitz.Point(x + c[1][0], base_y - c[1][1])
                shape.draw_line(cur, p)
                cur = p
            elif c[0] == "c":
                p1 = fitz.Point(x + c[1][0], base_y - c[1][1])
                p2 = fitz.Point(x + c[2][0], base_y - c[2][1])
                p3 = fitz.Point(x + c[3][0], base_y - c[3][1])
                shape.draw_bezier(cur, p1, p2, p3)
                cur = p3
            elif c[0] == "h" and cur is not None and cur != start:
                shape.draw_line(cur, start)
                cur = start
        x += adv
    shape.finish(color=None, fill=color, closePath=True)
    shape.commit()
    # invisible selectable layer, deliberately BLACK: the fault this guards is
    # an engine trusting the invisible layer's colour instead of the ink's
    page.insert_text(origin, text, fontsize=fontsize, fontname="hebo",
                     render_mode=3, color=(0, 0, 0))


# --------------------------------------------------------------------------- #
def footer_form_doc():
    """Footer small print that becomes a Form XObject, with space-only and
    trailing-space ops, plus a GREY line (colour must survive the form)."""
    d = fitz.open()
    p = d.new_page(width=400, height=60)
    p.insert_text((10, 18),
                  "Example Bank plc is authorised by the Fabricated Regulation Authority",
                  fontname="tiro", fontsize=9)
    p.insert_text((305, 18), "   ", fontname="tiro", fontsize=9)   # space-only op
    p.insert_text((10, 4), "Terms apply ", fontname="tiro", fontsize=9)  # trailing space
    p.insert_text((10, 32),
                  "and regulated under fabricated registration number 000000.",
                  fontname="tiro", fontsize=9, color=GREY)
    p.insert_text((10, 50),
                  "Registered office: 1 Example Street, Example City EX1 1EX.",
                  fontname="tiro", fontsize=9)
    return d


def nested_form_doc(inner):
    """A document that itself shows the footer form, scaled: form in form."""
    d = fitz.open()
    p = d.new_page(width=300, height=50)
    p.insert_text((6, 14), "Sheet 1 of 2", fontname="cobo", fontsize=10)
    p.insert_text((70, 14), "  ", fontname="cour", fontsize=10)    # space-only op in nested form
    p.show_pdf_page(fitz.Rect(80, 5, 295, 45), inner, 0)
    return d


def outline_form_doc():
    """A form whose text is glyph-outline paths + invisible overlay: the
    exact construct behind the ghost-doubling and colour-flattening faults."""
    d = fitz.open()
    p = d.new_page(width=260, height=30)
    draw_outline_text(p, (4, 20), "STATEMENT OF FEES", 14, BRAND_PURPLE)
    return d


def page_one(doc):
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
    # (e) brand-coloured heading
    page.insert_text((left + 110, 62), "EXAMPLE BANK", fontname="hebo",
                     fontsize=15, color=BRAND_BLUE)

    # (c) decorative 8-square brand mark: contiguous stroked squares, NO text.
    # A grid detector that promotes this emits an empty 1x8 native table.
    sq, sy = 9, 74
    for i in range(8):
        page.draw_rect(fitz.Rect(left + 110 + i * sq, sy, left + 110 + (i + 1) * sq, sy + sq),
                       width=0.8)

    # (a) header form at top right: on unfixed code the operator walk desyncs
    # here and corrupts everything below it on the page
    hdr = fitz.open()
    hp = hdr.new_page(width=200, height=24)
    hp.insert_text((4, 16), "Statement of account ", fontname="tiro", fontsize=11)
    hp.insert_text((120, 16), "  ", fontname="tiro", fontsize=11)
    page.show_pdf_page(fitz.Rect(right - 180, 42, right, 66), hdr, 0)
    hdr.close()

    # (d) narrow multi-word column header: two stacked short lines. Sized to
    # source width, a substituted font breaks "BALANCE" mid-word. Placed in
    # its own y-band so the two lines cluster into one block.
    page.insert_text((right - 105, 104), "END OF DAY", fontname="hebo", fontsize=8)
    page.insert_text((right - 105, 114), "ACCOUNT BALANCE", fontname="hebo", fontsize=8)

    # (a) glyph-outline text + invisible overlay, inside a Form XObject
    of = outline_form_doc()
    page.show_pdf_page(fitz.Rect(left, 78, left + 234, 105), of, 0)
    of.close()

    # dense ruled table (normal content, must keep working)
    rng = random.Random(20260721)
    cols = [left, 110, 330, 405, 480, right]
    top = 128
    row_h = 18
    n_rows = 28
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

    # (e) grey small print under the table
    page.insert_text((left, bottom + 20),
                     "Interest is calculated daily on fabricated balances. Data for testing only.",
                     fontname="helv", fontsize=7, color=GREY)

    # genuine non-table vector art: a brand band. Forces the hybrid background
    # in BOTH engines so the ghost-doubling invariant exercises each engine's
    # background text-strip (python's stays native otherwise and skips it).
    page.draw_rect(fitz.Rect(left, 690, right, 726), fill=(0.85, 0.90, 0.97), width=0)
    page.draw_rect(fitz.Rect(left, 690, left + 6, 726), fill=BRAND_BLUE, width=0)

    # (a) footer small print: nested Form XObjects, scaled, at page bottom
    inner = footer_form_doc()
    nested = nested_form_doc(inner)
    page.show_pdf_page(fitz.Rect(left, H - 70, left + 420, H - 20), nested, 0)
    inner.close()
    nested.close()


def page_two(doc):
    """(b) the UNRULED transaction ledger: aligned columns, zero ruling lines."""
    W, H = 595, 842
    page = doc.new_page(width=W, height=H)
    left, right = 40, W - 40
    rng = random.Random(20260722)

    page.insert_text((left, 60), "TRANSACTION LEDGER (CONTINUED)", fontname="hebo",
                     fontsize=13, color=BRAND_BLUE)
    page.insert_text((left, 78), "Sheet 2 of 2 - fabricated data", fontname="helv",
                     fontsize=9, color=GREY)

    cols = [left, 110, 330, 405, 480]
    colr = [105, 320, 395, 470, right]     # right rails for the money columns
    top = 110
    row_h = 16
    for ci, htext in enumerate(["Date", "Description", "Money out", "Money in", "Balance"]):
        page.insert_text((cols[ci], top), htext, fontname="hebo", fontsize=8)
    balance = 1234.56
    n_rows = 16
    for r in range(n_rows):
        yy = top + (r + 1) * row_h
        out_amt = in_amt = None
        if rng.random() < 0.75:
            out_amt = round(rng.uniform(3.5, 220.0), 2)
            balance -= out_amt
        else:
            in_amt = round(rng.uniform(100.0, 1400.0), 2)
            balance += in_amt
        page.insert_text((cols[0], yy), "%02d JUL" % (r % 28 + 1), fontname="helv", fontsize=8)
        page.insert_text((cols[1], yy), "UNRULED PAYEE %02d" % (r % 9), fontname="helv", fontsize=8)
        if out_amt is not None:
            s = "%.2f" % out_amt
            page.insert_text((colr[2] - fitz.get_text_length(s, "helv", 8), yy), s,
                             fontname="helv", fontsize=8)
        if in_amt is not None:
            s = "%.2f" % in_amt
            page.insert_text((colr[3] - fitz.get_text_length(s, "helv", 8), yy), s,
                             fontname="helv", fontsize=8)
        s = "%.2f" % balance
        page.insert_text((colr[4] - fitz.get_text_length(s, "helv", 8), yy), s,
                         fontname="helv", fontsize=8)

    # prose control: aligned-column inference must never tabulate a paragraph
    y = top + (n_rows + 3) * row_h
    for i, t in enumerate([
            "This closing paragraph is ordinary prose and must never be promoted",
            "to a table by column inference. It clusters into one logical block",
            "of wrapped body text under the ledger, nothing more."]):
        page.insert_text((left, y + i * 12), t, fontname="helv", fontsize=9)


def main():
    doc = fitz.open()
    page_one(doc)
    page_two(doc)
    doc.save(OUT)
    doc.close()
    print("wrote", OUT)


if __name__ == "__main__":
    main()
