"""Generate the SYNTHETIC styled-proposal fixture.

Guards four defect classes found in a real business-proposal conversion on
2026-07-20, none of which the earlier fixtures exercised:

  1. OVERPRINTED TITLE: the same title drawn twice (a poor man's bold/shadow).
     MuPDF dedupes identical glyphs at one position; the browser engine must
     too, or the title renders twice and garbled.
  2. LETTER-SPACED CAPS: "PREPARED FOR MODA" tracked out per glyph with no
     space glyphs at all; word gaps exist only as larger advances (0.45em vs
     0.15em letter gaps) and must come back as spaces.
  3. TIGHT NUMBERED MARKERS: "01" as two digits drawn with negative tracking
     (the second starts left of where the first ends); they must stay one
     horizontal run, not stack vertically.
  4. GENEROUS LEADING PARAGRAPHS: body copy at 1.7x leading; block counts
     must match the desktop engine's (golden), not split mid-sentence.

Plus two-column feature cards, the layout the defects appeared in.
Synthetic only: every word here is fabricated.
"""

import os

import fitz

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "fixtures", "styled_proposal.pdf")


def main():
    doc = fitz.open()
    W, H = 595, 842
    page = doc.new_page(width=W, height=H)
    left = 60

    # 1) overprinted title: same string, same position, drawn twice
    for _ in range(2):
        page.insert_text((left, 90), "BUSINESS PROPOSAL", fontname="hebo", fontsize=28)

    # 2) caps line with NO space glyphs: one TJ op, word gaps exist only as
    # kern advances of 0.45em. This is the construct that produced
    # "PREPAREDFORMODA" in the field.
    page.insert_text((left, 130), " ", fontname="helv", fontsize=1)  # register /helv
    tj = b"q BT /helv 14 Tf 1 0 0 1 60 712 Tm [(PREPARED) -450 (FOR) -450 (MODA)] TJ ET Q "
    xref = page.get_contents()[0]
    doc.update_stream(xref, doc.xref_stream(xref) + b" " + tj)

    # 3) numbered circle markers with negative tracking on the digits
    for i, ny in enumerate((200, 260)):
        page.draw_circle(fitz.Point(left + 14, ny), 16, width=1.2)
        d1, d2 = ("0", str(i + 1))
        w1 = fitz.get_text_length(d1, fontname="hebo", fontsize=16)
        page.insert_text((left + 5, ny + 6), d1, fontname="hebo", fontsize=16)
        page.insert_text((left + 5 + w1 * 0.62, ny + 6), d2, fontname="hebo", fontsize=16)
        page.insert_text((left + 44, ny + 5),
                         ["Fabricated milestone one, stated plainly.",
                          "Fabricated milestone two, stated plainly."][i],
                         fontname="helv", fontsize=11)

    # 4) paragraphs at 1.7x leading
    body = [
        "This proposal sets out a fabricated plan in several sentences that",
        "wrap across lines with generous leading, so the paragraph must stay",
        "one block in the converted deck rather than splitting mid-sentence.",
    ]
    y = 330
    for ln in body:
        page.insert_text((left, y), ln, fontname="helv", fontsize=11)
        y += 11 * 1.7
    y += 11 * 0.9   # paragraph break
    for ln in ["A second fabricated paragraph follows the first after a larger",
               "gap and must arrive as its own paragraph, not merge upward."]:
        page.insert_text((left, y), ln, fontname="helv", fontsize=11)
        y += 11 * 1.7

    # two-column feature cards
    for ci, (cx, head) in enumerate([(left, "Fabricated scope"), (318, "Fabricated terms")]):
        r = fitz.Rect(cx, 520, cx + 217, 620)
        page.draw_rect(r, width=1.0)
        page.insert_text((cx + 14, 548), head, fontname="hebo", fontsize=12)
        page.insert_text((cx + 14, 570), "One line of fabricated detail.", fontname="helv", fontsize=10)
        page.insert_text((cx + 14, 588), "A second line of detail.", fontname="helv", fontsize=10)

    doc.save(OUT)
    doc.close()
    print("wrote", OUT)


if __name__ == "__main__":
    main()
