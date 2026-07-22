"""Generate the SYNTHETIC dingbat-marker list fixture.

Guards one defect class, found on page 3 of the W3C WCAG 2.0 working draft on
2026-07-22, then the real corpus's worst page at 0.1727:

  A BULLET IS NOT A CHARACTER, IT IS A CHARACTER IN A DINGBAT FONT. The page
  sets every list marker as a 7pt ZapfDingbats "H", which draws a hollow
  circle, and the outer level as "G", a filled one. Nothing about the code
  point says "bullet" - both are ASCII letters - so every code-point test
  missed them, nothing flagged the lines as list items, and 25 evenly-leaded
  entries clustered into ONE paragraph that reflowed into a block of prose.

  Two consequences, both asserted: the list must stay one paragraph per item,
  and the marker must not ship as a tiny letter H in a substituted font.

A two-level list with wrapped items, drawn the way a 2001 browser print did.
Synthetic only: every word here is fabricated.
"""

import os

import fitz

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "fixtures", "dingbat_toc.pdf")

# (level, text) - level 0 is a filled marker, level 1 a hollow one
ITEMS = [
    (0, "Section 1 - Handling. Make the workshop as easy to enter and leave "
        "as is practical for a registered holder"),
    (1, "Item 1.1 Keep the entry register beside the door."),
    (1, "Item 1.2 Record the tooling brought in on each visit, including "
        "tooling carried by a visiting holder."),
    (1, "Item 1.3 Do not leave the register open when the workshop is "
        "unattended."),
    (0, "Section 2 - Storage. Keep tooling where a later holder can find it"),
    (1, "Item 2.1 Label each rack with its own reference."),
    (1, "Item 2.2 Store duplicate tooling together, and record the duplicate "
        "reference beside the original in the register."),
    (1, "Item 2.3 Report a missing reference before the end of the period."),
    (0, "Section 3 - Returns"),
    (1, "Item 3.1 Return the register at the end of the period."),
    (1, "Item 3.2 Keep a copy of every entry you return."),
]

# ZapfDingbats: "G" draws a filled circle, "H" a hollow one. fitz calls the
# base-14 font "zadb".
MARKERS = {0: "G", 1: "H"}
INDENT = {0: 94.5, 1: 133.8}
TEXT_X = {0: 115.0, 1: 154.0}


def main():
    doc = fitz.open()
    W, H = 612, 792
    page = doc.new_page(width=W, height=H)
    font = fitz.Font("helv")
    size, leading = 14, 18.7

    page.insert_text((18, 12), "Workshop Handling Guidelines 2.0",
                     fontname="helv", fontsize=9)

    zadb, helv = fitz.Font("zadb"), fitz.Font("helv")
    y = 46
    for level, text in ITEMS:
        # Wrap first, so the marker can be laid down with the opening line.
        lines, line = [], ""
        for w in text.split():
            probe = (line + " " + w).strip()
            if font.text_length(probe, size) > (W - TEXT_X[level] - 40) and line:
                lines.append(line)
                line = w
            else:
                line = probe
        if line:
            lines.append(line)

        # Both real shapes, so both detection paths are guarded: an outer
        # marker drawn as its own text object (a detached marker line), an
        # inner one drawn contiguously with its text (one line, three spans:
        # dingbat glyph, spaces, text).
        if level == 0:
            page.insert_text((INDENT[level], y), MARKERS[level],
                             fontname="zadb", fontsize=7)
            page.insert_text((TEXT_X[level], y), lines[0],
                             fontname="helv", fontsize=size)
        else:
            # drawn contiguously, so MuPDF reads one line of three spans;
            # insert_text (not TextWriter, which re-embeds "zadb" as a serif
            # text font and loses the ZapfDingbats name the detector reads)
            page.insert_text((INDENT[level], y), MARKERS[level],
                             fontname="zadb", fontsize=7)
            page.insert_text((INDENT[level] + zadb.text_length(MARKERS[level], 7), y),
                             "     " + lines[0], fontname="helv", fontsize=size)
        y += leading
        for cont in lines[1:]:
            page.insert_text((TEXT_X[level] + 16, y), cont,
                             fontname="helv", fontsize=size)
            y += leading

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT, deflate=True)
    doc.close()
    print("wrote", OUT)


if __name__ == "__main__":
    main()
