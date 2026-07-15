"""Extra test PDFs for failure-path testing: a scanned/image-only PDF and a
50+ page text PDF."""
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "input")
os.makedirs(OUT, exist_ok=True)


def make_scanned(path):
    """Image-only PDF: a picture of text, with NO selectable text layer."""
    W, H = int(8.5 * 150), int(11 * 150)  # 150 dpi letter
    img = Image.new("RGB", (W, H), (247, 245, 239))
    d = ImageDraw.Draw(img)
    lines = [
        "MEMORANDUM",
        "",
        "This page exists only as a scanned image.",
        "There is no selectable text layer here,",
        "so the converter should detect it and warn",
        "rather than produce an empty slide.",
        "",
        "— End of scan —",
    ]
    y = 220
    for ln in lines:
        d.text((180, y), ln, fill=(30, 28, 22))
        y += 60
    d.rectangle([120, 150, W - 120, y + 60], outline=(150, 145, 130), width=2)
    tmp = os.path.join(OUT, "_scan.png")
    img.save(tmp, dpi=(150, 150))

    c = canvas.Canvas(path, pagesize=letter)
    c.drawImage(tmp, 0, 0, width=8.5 * inch, height=11 * inch)
    c.showPage()
    c.save()
    os.remove(tmp)
    print("wrote", path)


def make_big(path, pages=52):
    doc = SimpleDocTemplate(path, pagesize=letter,
                            leftMargin=1 * inch, rightMargin=1 * inch,
                            topMargin=1 * inch, bottomMargin=1 * inch)
    styles = getSampleStyleSheet()
    body = ("This is a long multi-page document used to check that the tool stays "
            "responsive and the progress bar advances smoothly across many pages. "
            "Each page carries several paragraphs of ordinary running text.")
    story = []
    for i in range(1, pages + 1):
        story.append(Paragraph(f"Section {i}", styles["Heading2"]))
        for _ in range(4):
            story.append(Paragraph(body, styles["BodyText"]))
            story.append(Spacer(1, 0.1 * inch))
        if i < pages:
            story.append(PageBreak())
    doc.build(story)
    print("wrote", path, "pages≈", pages)


if __name__ == "__main__":
    make_scanned(os.path.join(OUT, "scanned.pdf"))
    make_big(os.path.join(OUT, "big50.pdf"))
    print("done")
