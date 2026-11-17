"""An Excel workbook with a table, conditional formatting, comments, links,
print setup and a hidden lookup sheet.

    python examples/excel_dashboard.py dashboard.xlsx
"""

import sys

from oox import SpreadsheetDocument
from oox.spreadsheet import Formula


def main(path: str) -> None:
    book = SpreadsheetDocument.new("Sales")
    book.add_sheet("Lookup")

    book.append_row("Sales", ["Region", "Target", "Actual", "Trend"])
    rows = [("North", 100, 120, 9), ("South", 100, 30, 1), ("West", 80, 84, 5),
            ("East", 90, 60, 10)]
    for row in rows:
        book.append_row("Sales", list(row))

    # an Excel table with a totals row; structured references work in formulas
    book.add_table("Sales", "A1:D5", name="Sales2026",
                   totals={"Region": "Total", "Target": "sum", "Actual": "sum"})
    book.set_value("Sales", "F1", "Hit rate")
    book.set_value("Sales", "G1", Formula("SUM(Sales2026[Actual])/SUM(Sales2026[Target])", 294 / 370),
                   style={"number_format": "0.0%"})

    # conditional formatting
    book.conditional_format("Sales", "C2:C5", "formula", formula="$C2<$B2",
                            style={"fill": "FFC7CE", "color": "9C0006"})
    book.conditional_format("Sales", "D2:D5", "icon_set", icons="3Arrows")

    # notes and links
    book.add_comment("Sales", "C3", "Storm damage closed two stores.", author="Ada")
    book.add_hyperlink("Sales", "F3", "https://ecma-international.org/", text="About OOXML")
    book.add_hyperlink("Sales", "F4", "#Lookup!A1", text="Region codes")

    # a lookup sheet users don't need to see
    book.append_row("Lookup", ["Code", "Region"])
    for code, region in (("N", "North"), ("S", "South"), ("W", "West"), ("E", "East")):
        book.append_row("Lookup", [code, region])
    book.set_sheet_options("Lookup", hidden=True)

    # print setup, tab colour, protection
    book.set_page_setup("Sales", orientation="landscape", paper="A4", fit_width=1,
                    repeat_rows="1:1", footer={"right": "Page {page} of {pages}"})
    book.set_sheet_options("Sales", tab_color="2F5496")
    book.autofit_columns("Sales")
    book.protect_sheet("Sales", allow=("sort", "autoFilter"))

    book.share_strings()          # store text as Excel does
    assert not book.validate(target="Office2007")
    book.save(path)
    print(book.tables("Sales")[0].ref, len(book.conditional_formats("Sales")),
          book.comments("Sales"))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "dashboard.xlsx")
