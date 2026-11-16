"""One script, three formats: a .docx, an .xlsx and a .pptx with pictures.

    python examples/showcase.py OUTDIR

Needs Pillow only to *draw* the sample images; oox itself doesn't use it.
"""

import io
import sys
from pathlib import Path

from PIL import Image, ImageDraw

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument
from oox.ns import w
from oox.spreadsheet import Formula
from oox.units import Cm, Inches

SALES = [("North", 120, 9.5), ("South", 95, 11.0), ("East", 143, 8.75), ("West", 77, 12.25)]


def bar_chart_png() -> bytes:
    img = Image.new("RGB", (640, 360), "white")
    d = ImageDraw.Draw(img)
    colors = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000"]
    for i, (region, units, _) in enumerate(SALES):
        x0 = 60 + i * 140
        d.rectangle([x0, 330 - units * 2, x0 + 90, 330], fill=colors[i])
        d.text((x0 + 25, 335), region, fill="black")
    d.line([40, 330, 620, 330], fill="black", width=2)
    buf = io.BytesIO()
    img.save(buf, "PNG", dpi=(144, 144))
    return buf.getvalue()


def photo_jpeg() -> bytes:
    img = Image.new("RGB", (400, 300))
    px = img.load()
    for yy in range(300):
        for xx in range(400):
            px[xx, yy] = (30 + xx // 3, 80 + yy // 3, 160)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85, dpi=(72, 72))
    return buf.getvalue()


def make_docx(path: Path, chart: bytes, photo: bytes) -> None:
    doc = WordprocessingDocument.new()
    doc.body.add(
        w.P(w.PPr(w.Jc(val="center")),
            w.R(w.RPr(w.B(), w.Sz(val=40)), w.T("Regional sales"))),
        w.P(w.R(w.T("Units sold per region this quarter:"))),
    )
    doc.add_picture(chart, width=Cm(14), descr="Bar chart of units per region")
    para = w.P(w.R(w.T("A photo inline with text: ")))
    doc.body.add(para)
    doc.add_picture(photo, height=Cm(2), paragraph=para)
    doc.save(path)


def make_xlsx(path: Path, chart: bytes) -> None:
    book = SpreadsheetDocument.new("Sales")
    book.append_row("Sales", ["Region", "Units", "Price", "Revenue"])
    for i, (region, units, price) in enumerate(SALES, start=2):
        book.append_row("Sales", [region, units, price, Formula(f"B{i}*C{i}")])
    n = len(SALES) + 1
    book.append_row("Sales", ["Total", Formula(f"SUM(B2:B{n})"), None, Formula(f"SUM(D2:D{n})")])
    book.add_picture("Sales", chart, anchor="F2", width=Cm(12))
    book.save(path)


def make_pptx(path: Path, chart: bytes, photo: bytes) -> None:
    deck = PresentationDocument.new()
    s1 = deck.add_slide()
    deck.add_textbox(s1, Inches(0.7), Inches(0.5), Inches(12), Inches(1), "Regional sales",
                     size=40, bold=True, color="1F3864")
    deck.add_textbox(s1, Inches(0.7), Inches(1.8), Inches(5), Inches(3),
                     "East leads on units\nWest has the best price\nTotal: 435 units", size=22)
    deck.add_picture(s1, chart, Inches(6.3), Inches(1.7), width=Inches(6.3))
    s2 = deck.add_slide()
    deck.add_picture(s2, photo, Inches(3.67), Inches(1), width=Inches(6))
    deck.add_textbox(s2, Inches(0.7), Inches(5.8), Inches(12), Inches(1), "Thank you",
                     size=32, align="center")
    deck.save(path)


def main(outdir: str) -> None:
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    chart, photo = bar_chart_png(), photo_jpeg()
    make_docx(out / "showcase.docx", chart, photo)
    make_xlsx(out / "showcase.xlsx", chart)
    make_pptx(out / "showcase.pptx", chart, photo)
    print(f"wrote showcase.docx/.xlsx/.pptx to {out}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
