"""Open files made elsewhere, read them, change them, save them.

    python examples/edit_existing.py OUTDIR [workbook.xlsx] [deck.pptx]

Pass your own Excel workbook and PowerPoint deck (any producer: Excel,
PowerPoint, LibreOffice, openpyxl, python-pptx ...).  Without them, the script
first writes two stand-in inputs so it runs anywhere.
"""

import sys
import zipfile
from pathlib import Path

from oox import PresentationDocument, SpreadsheetDocument, open_document
from oox.spreadsheet import Formula


def make_stand_ins(out: Path) -> tuple[Path, Path]:
    book = SpreadsheetDocument.new("Stock")
    book.append_row("Stock", ["Item", "Qty", "Unit price"])
    for item, qty, price in [("Pens", 120, 0.8), ("Paper", 40, 4.5), ("Ink", 15, 12.0)]:
        book.append_row("Stock", [item, qty, price])
    book.save(out / "input.xlsx")
    PresentationDocument.new().save(out / "input.pptx")
    return out / "input.xlsx", out / "input.pptx"


def edit_workbook(src: Path, dst: Path) -> None:
    book = open_document(src)                      # picks SpreadsheetDocument
    sheet = book.sheet_names[0]
    rows = book.values(sheet)                      # dense list of rows from A1
    print(f"{src.name}: sheet {sheet!r}, {len(rows)} rows")
    for row in rows:
        print("   ", row)

    last = len(rows)
    book.set_value(sheet, f"A{last + 2}", "Total qty")
    book.set_value(sheet, f"B{last + 2}", Formula(f"SUM(B2:B{last})"))
    book.set_value(sheet, "D1", "Checked")         # a new cell in an existing row
    book.save(dst)
    print(f"  -> {dst.name}: total formula in B{last + 2}")


def edit_deck(src: Path, dst: Path) -> None:
    deck = open_document(src)                      # picks PresentationDocument
    print(f"{src.name}: {len(deck.slides)} slides, layouts {deck.layout_names}")
    layout = "Title and Content" if "Title and Content" in deck.layout_names else None
    slide = deck.add_slide(layout)
    deck.set_placeholder_text(slide, "title", "Added by oox")
    if layout:
        deck.set_placeholder_text(slide, "body", ["Read the deck", (1, "kept every part"),
                                                  "Added this slide"])
    deck.save(dst)

    before, after = zipfile.ZipFile(src), zipfile.ZipFile(dst)
    changed = [n for n in before.namelist() if before.read(n) != after.read(n)]
    print(f"  -> {dst.name}: {len(changed)} of {len(before.namelist())} zip entries changed: {changed}")


def main(argv: list[str]) -> None:
    out = Path(argv[1] if len(argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    if len(argv) > 3:
        xlsx, pptx = Path(argv[2]), Path(argv[3])
    else:
        xlsx, pptx = make_stand_ins(out)
    edit_workbook(xlsx, out / "edited.xlsx")
    edit_deck(pptx, out / "edited.pptx")


if __name__ == "__main__":
    main(sys.argv)
