"""Generate the SYNTHETIC bordered-callout fixture.

Guards one defect class, found on page 3 of a real government guidance note on
2026-07-22 (real/govuk_r43_notes, then the corpus's worst page at 0.1414):

  A BORDERED CALLOUT IS NOT A TABLE. The page draws a rule around its right
  column and one more rule above the page number, which reads to any
  line-based table detector as a ruled 2x2: one cell holding a whole
  1,600-character column of prose, one cell empty, one cell holding the page
  number, one cell absent. Shipped as a native table, the column of prose
  reflows inside a single cell and the page collapses.

  The test that catches it is on cell CONTENT, not on how the grid was
  detected: a table earns its cells by using them.

One column of guidance prose inside the rule, page number bottom right. The
real page is two-column; this fixture is not, because the engines still differ
on multi-column clustering and this fixture exists to guard ONE class.
Synthetic only: every word here is fabricated.
"""

import os

import fitz

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "fixtures", "callout_notes.pdf")

LEFT_PARAS = [
    "These notes describe the allowances claimable by a resident of the "
    "Northern Territory who holds a registered interest in a licensed "
    "workshop.",
    "Do not use these notes if the workshop is unregistered. The notes will "
    "help you claim relief on tooling, claim relief on materials, and record "
    "the balance carried forward at the end of the period.",
    "To make a claim you need to complete and sign form W12. The form asks "
    "for details of every registered interest, the allowances taken in the "
    "period, and the deductions carried across from the preceding period.",
    "If there is not enough space on the form, list the remaining items on a "
    "separate sheet, put the total on the form itself, and send the sheet "
    "with your claim. Your reference number is printed in the top right "
    "corner of the first page of the form.",
]

RIGHT_PARAS = [
    "Enquiries",
    "We may ask you to send evidence of the allowances taken, such as "
    "tooling receipts or a workshop register extract.",
    "Relief on registered tooling",
    "A registered holder is liable to duty on the value of tooling brought "
    "into the workshop during the period. Relief is normally given at the "
    "standard rate, and is given at the higher rate where the tooling was "
    "acquired on or after 6 April 2013.",
    "If you are not a registered holder you are not liable to duty on "
    "tooling, and no relief arises. Duty on registered tooling is normally "
    "settled with no deduction taken. You do not need to enter details of "
    "relief on tooling acquired from 6 April 2013 onwards.",
]


def _flow(page, paras, x, y, width, size, leading, bold_first_line=()):
    """Lay out wrapped paragraphs, returning the y after the last line."""
    font = fitz.Font("helv")
    bold = fitz.Font("hebo")
    for para in paras:
        heading = para in bold_first_line
        f = bold if heading else font
        fn = "hebo" if heading else "helv"
        words, line = para.split(), ""
        for w in words:
            probe = (line + " " + w).strip()
            if f.text_length(probe, size) > width and line:
                page.insert_text((x, y), line, fontname=fn, fontsize=size)
                y += leading
                line = w
            else:
                line = probe
        if line:
            page.insert_text((x, y), line, fontname=fn, fontsize=size)
            y += leading
        y += leading * 0.45
    return y


def main():
    doc = fitz.open()
    W, H = 558, 540
    page = doc.new_page(width=W, height=H)
    size, leading = 9, 12.4

    y = _flow(page, LEFT_PARAS, 300, 40, 224, size, leading)
    _flow(page, RIGHT_PARAS, 300, y + leading, 224, size, leading,
          bold_first_line={"Enquiries", "Relief on registered tooling"})

    # The furniture: a rule around the whole column of prose, and one more rule
    # above the page number. A line-based detector reads this as a 2x2 grid.
    box = fitz.Rect(286, 24, 534, 532)
    page.draw_rect(box, color=(0, 0, 0), width=0.7)
    page.draw_line(fitz.Point(286, 515.6), fitz.Point(534, 515.6),
                   color=(0, 0, 0), width=0.7)
    page.draw_line(fitz.Point(504, 515.6), fitz.Point(504, 532),
                   color=(0, 0, 0), width=0.7)
    page.insert_text((515, 528), "3", fontname="helv", fontsize=9)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT, deflate=True)
    doc.close()
    print("wrote", OUT)


if __name__ == "__main__":
    main()
