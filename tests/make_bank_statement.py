"""Generate the SYNTHETIC bank-statement fixture.

Dense ruled transaction tables with running balances, across two pages: the
document class users convert most and the one that punishes a weak table
detector. Every name, account and amount is fabricated here, deterministically,
so the fixture is reproducible and nothing real is ever committed.
"""

import os
import random

import fitz

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures", "bank_statement.pdf")

PAYEES = ["GREENGROCER LTD", "NORTHWIND ENERGY", "CITY TRANSIT", "ACME PAYROLL",
          "BLUE CAFE 0042", "STREAMFLIX", "TELCO MOBILE", "CORNER PHARMACY",
          "BOOKS & MAPS CO", "RAILWAYS ONLINE", "HOME INSURANCE PLC", "GYM CLUB 12"]


def main():
    rng = random.Random(20260720)
    doc = fitz.open()
    W, H = 595, 842            # A4 portrait
    left, right = 40, W - 40
    cols = [left, 110, 330, 405, 480, right]   # date | description | out | in | balance
    headers = ["Date", "Description", "Money out", "Money in", "Balance"]

    balance = 4180.55
    day, month = 1, 5

    for pageno in range(2):
        page = doc.new_page(width=W, height=H)
        y = 60
        page.insert_text((left, y), "EXAMPLE BANK", fontname="hebo", fontsize=16)
        page.insert_text((right - 150, y), "Statement of account", fontname="helv", fontsize=9)
        y += 26
        page.insert_text((left, y), "A N Example - Account 00-00-00 12345678 - Sheet %d of 2" % (pageno + 1),
                         fontname="helv", fontsize=9)
        y += 24

        row_h = 20
        n_rows = 32
        top = y
        bottom = top + row_h * (n_rows + 1)

        # ruling lines: full grid, the lines-strategy detector's bread and butter
        for r in range(n_rows + 2):
            yy = top + r * row_h
            page.draw_line(fitz.Point(left, yy), fitz.Point(right, yy), width=0.7)
        for x in cols:
            page.draw_line(fitz.Point(x, top), fitz.Point(x, bottom), width=0.7)

        # header row, shaded
        page.draw_rect(fitz.Rect(left, top, right, top + row_h), fill=(0.92, 0.92, 0.92), width=0)
        for ci, htext in enumerate(headers):
            page.insert_text((cols[ci] + 4, top + 14), htext, fontname="hebo", fontsize=8)

        for r in range(n_rows):
            yy = top + (r + 1) * row_h + 14
            out_amt = in_amt = None
            if rng.random() < 0.82:
                out_amt = round(rng.uniform(3.5, 220.0), 2)
                balance -= out_amt
            else:
                in_amt = round(rng.uniform(100.0, 2400.0), 2)
                balance += in_amt
            date = "%02d MAY" % day
            if rng.random() < 0.55:
                day = min(day + 1, 31)
            payee = PAYEES[rng.randrange(len(PAYEES))]
            page.insert_text((cols[0] + 4, yy), date, fontname="helv", fontsize=8)
            page.insert_text((cols[1] + 4, yy), payee, fontname="helv", fontsize=8)
            if out_amt is not None:
                page.insert_text((cols[2] + 4, yy), "%.2f" % out_amt, fontname="helv", fontsize=8)
            if in_amt is not None:
                page.insert_text((cols[3] + 4, yy), "%.2f" % in_amt, fontname="helv", fontsize=8)
            page.insert_text((cols[4] + 4, yy), "%.2f" % balance, fontname="helv", fontsize=8)

        page.insert_text((left, bottom + 24),
                         "Interest rate 1.20%% AER variable. Fabricated data for testing only.",
                         fontname="helv", fontsize=7)

    doc.save(OUT)
    doc.close()
    print("wrote", OUT)


if __name__ == "__main__":
    main()
