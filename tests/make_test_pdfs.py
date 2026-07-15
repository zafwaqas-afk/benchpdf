"""
Generate three representative test PDFs entirely offline with reportlab:

  1. text_report.pdf   - text-heavy multi-page report (native mode expected)
  2. slide_deck.pdf     - slide-style pages with raster images (native mode)
  3. tables_charts.pdf  - tables with ruling lines + a vector bar chart (hybrid)
"""

import os

from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

from PIL import Image as PILImage, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "input")
os.makedirs(OUT, exist_ok=True)


# --------------------------------------------------------------------------- #
def make_text_report(path):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CenterTitle", parent=styles["Title"], alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="RightNote", parent=styles["Normal"], alignment=TA_RIGHT,
                              fontName="Times-Italic", fontSize=10))
    doc = SimpleDocTemplate(path, pagesize=letter,
                            leftMargin=1 * inch, rightMargin=1 * inch,
                            topMargin=1 * inch, bottomMargin=1 * inch)
    story = []
    story.append(Paragraph("Quarterly Operations Review", styles["CenterTitle"]))
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph("Prepared by the Operations Team", styles["RightNote"]))
    story.append(Spacer(1, 0.3 * inch))

    body = (
        "The third quarter demonstrated resilient performance across every "
        "major business unit. Revenue expanded while operating costs held "
        "steady, producing a meaningful improvement in overall margin. "
        "Customer retention reached its highest level in two years, driven by "
        "investments in onboarding and support."
    )
    heading = ParagraphStyle(name="H", parent=styles["Heading2"], fontName="Helvetica-Bold")
    for section in ("Executive Summary", "Financial Highlights", "Operational Metrics",
                    "Regional Performance", "Outlook and Priorities"):
        story.append(Paragraph(section, heading))
        for _ in range(3):
            story.append(Paragraph(body, styles["BodyText"]))
            story.append(Spacer(1, 0.12 * inch))
        story.append(Spacer(1, 0.15 * inch))

    story.append(Paragraph(
        "In closing, the organisation enters the final quarter with strong "
        "momentum, a healthy pipeline, and a disciplined approach to spending.",
        styles["BodyText"]))
    doc.build(story)
    print("wrote", path)


# --------------------------------------------------------------------------- #
def _gradient_image(w, h, c1, c2, name):
    img = PILImage.new("RGB", (w, h), c1)
    draw = ImageDraw.Draw(img)
    for y in range(h):
        t = y / max(h - 1, 1)
        r = int(c1[0] + (c2[0] - c1[0]) * t)
        g = int(c1[1] + (c2[1] - c1[1]) * t)
        b = int(c1[2] + (c2[2] - c1[2]) * t)
        draw.line([(0, y), (w, y)], fill=(r, g, b))
    draw.ellipse([w * 0.2, h * 0.2, w * 0.8, h * 0.8], outline=(255, 255, 255), width=6)
    p = os.path.join(OUT, name)
    img.save(p)
    return p


def make_slide_deck(path):
    """Slide-style landscape pages: a big title + a raster image per page."""
    page = landscape(letter)
    c = canvas.Canvas(path, pagesize=page)
    W, H = page

    imgs = [
        _gradient_image(600, 400, (30, 90, 160), (120, 200, 240), "_img1.png"),
        _gradient_image(600, 400, (150, 40, 60), (240, 170, 120), "_img2.png"),
    ]
    slides = [
        ("Product Vision", "Building tools people love to use every day.", imgs[0]),
        ("Market Opportunity", "A large and growing segment ready for change.", imgs[1]),
        ("Roadmap 2026", "Three themes: speed, trust, and delight.", imgs[0]),
    ]
    for title, subtitle, img in slides:
        c.setFillColorRGB(1, 1, 1)
        c.rect(0, 0, W, H, fill=1, stroke=0)
        c.setFillColorRGB(0.1, 0.1, 0.2)
        c.setFont("Helvetica-Bold", 40)
        c.drawString(0.8 * inch, H - 1.3 * inch, title)
        c.setFillColorRGB(0.3, 0.3, 0.35)
        c.setFont("Helvetica", 20)
        c.drawString(0.8 * inch, H - 1.9 * inch, subtitle)
        c.drawImage(img, W - 5.2 * inch, 1.0 * inch, width=4.2 * inch, height=2.8 * inch)
        c.showPage()
    c.save()
    print("wrote", path)


# --------------------------------------------------------------------------- #
def make_tables_charts(path):
    doc = SimpleDocTemplate(path, pagesize=letter,
                            leftMargin=0.9 * inch, rightMargin=0.9 * inch,
                            topMargin=1 * inch, bottomMargin=1 * inch)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph("Regional Sales Table", styles["Title"]))
    story.append(Spacer(1, 0.2 * inch))

    data = [["Region", "Q1", "Q2", "Q3", "Q4", "Total"]]
    rows = [
        ("North", 120, 140, 160, 180),
        ("South", 90, 110, 100, 130),
        ("East", 200, 210, 230, 250),
        ("West", 75, 80, 95, 110),
    ]
    for name, *vals in rows:
        data.append([name] + [str(v) for v in vals] + [str(sum(vals))])
    tbl = Table(data, colWidths=[1.1 * inch] * 6)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.75, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#eef3f8")]),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.5 * inch))

    # a vector bar chart drawn as flowable via a small drawing
    from reportlab.graphics.shapes import Drawing, Rect, String, Line
    d = Drawing(400, 220)
    d.add(Rect(0, 0, 400, 220, fillColor=colors.HexColor("#f7f7f7"), strokeColor=None))
    d.add(Line(40, 30, 40, 200, strokeColor=colors.black))
    d.add(Line(40, 30, 380, 30, strokeColor=colors.black))
    bars = [90, 130, 70, 160, 110]
    palette = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd"]
    for idx, (val, col) in enumerate(zip(bars, palette)):
        x = 60 + idx * 62
        d.add(Rect(x, 30, 40, val, fillColor=colors.HexColor(col), strokeColor=None))
        d.add(String(x + 4, 32 + val, str(val), fontSize=9))
    d.add(String(150, 205, "Vector Bar Chart", fontSize=12, fillColor=colors.black))
    story.append(d)

    doc.build(story)
    print("wrote", path)


if __name__ == "__main__":
    make_text_report(os.path.join(OUT, "text_report.pdf"))
    make_slide_deck(os.path.join(OUT, "slide_deck.pdf"))
    make_tables_charts(os.path.join(OUT, "tables_charts.pdf"))
    print("done")
