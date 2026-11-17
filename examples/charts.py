"""Charts in all three formats: a workbook with four charts from a cell range,
a Word report with an inline chart and a slide with two charts.

    python examples/charts.py outdir
"""

import sys
from pathlib import Path

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument
from oox.units import Cm, Inches


def main(outdir: str) -> None:
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)

    # Excel: series from cells
    book = SpreadsheetDocument.new("Sales")
    book.append_row("Sales", ["Region", "Q1", "Q2", "Q3"], style={"bold": True})
    for row in [("North", 120, 135, 150), ("South", 95, 110, 99), ("West", 80, 84.5, 91),
                ("East", 60, 72, 88)]:
        book.append_row("Sales", list(row))
    book.add_chart("Sales", "column", "A1:D5", anchor="F2", title="Quarterly sales",
                   x_title="Region", y_title="EUR k", legend="bottom", labels="value")
    book.add_chart("Sales", "pie", "A1:B5", anchor="F18", title="Q1 share", labels="percent")
    book.add_chart("Sales", "line", "A1:D5", anchor="N2", series_in="rows", title="Trend")
    book.add_chart("Sales", "area", "A1:D5", anchor="N18", grouping="percent",
                   title="Share of quarter", number_format="0%")
    assert not book.validate(target="Office2007")
    book.save(out / "charts.xlsx")

    # Word: data from Python, kept in an embedded workbook
    doc = WordprocessingDocument.new()
    doc.add_heading("Sales report", 1)
    doc.add_paragraph("Sales by quarter:")
    doc.add_chart("column", ["Q1", "Q2", "Q3"], {"North": [120, 135, 150],
                                                  "South": [95, 110, 99]},
                  title="Sales", legend="bottom", labels="value")
    doc.add_chart("doughnut", ["Web", "Shop", "Phone"], {"Channel": [55, 30, 15]},
                  title="Channels", labels="percent", width=Cm(9), height=Cm(7))
    doc.save(out / "charts.docx")

    # PowerPoint
    deck = PresentationDocument.new()
    slide = deck.add_slide("Blank")
    deck.add_text_box(slide, "Traffic", x=Inches(0.5), y=Inches(0.3), width=Inches(9),
                      height=Inches(0.8), size=28, bold=True)
    deck.add_chart(slide, "line", ["Jan", "Feb", "Mar", "Apr"],
                   {"Visits": [1200, 1450, 1390, 1600], "Orders": [310, 365, 330, 410]},
                   x=Inches(0.5), y=Inches(1.3), width=Inches(6), height=Inches(4.5),
                   y_title="count")
    deck.add_chart(slide, "scatter", [1, 2, 3, 4, 5], {"Measured": [2.1, 3.9, 6.2, 7.8, 10.1]},
                   x=Inches(6.8), y=Inches(1.3), width=Inches(5.5), height=Inches(4.5),
                   lines=True, smooth=True, legend=None, title="Fit")
    deck.save(out / "charts.pptx")

    print([c.kind for c in book.charts("Sales")], [c.kind for c in doc.charts()],
          [c.kind for c in deck.charts(slide)])


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "charts-out")
