"""A formatted Excel report: styles, number formats, dates, column widths,
frozen header, filter, a named range and a drop-down.

    python examples/excel_report.py report.xlsx
"""

import datetime as dt
import sys

from oox import SpreadsheetDocument
from oox.spreadsheet import Formula


def main(path: str) -> None:
    book = SpreadsheetDocument.new("Orders")
    header = book.cell_style(bold=True, color="FFFFFF", fill="2F5496", align="center",
                             border="thin")
    money = book.cell_style(number_format="#,##0.00;[Red]-#,##0.00", border="thin")
    plain = book.cell_style(border="thin")
    dates = book.cell_style(number_format="d mmm yyyy", border="thin")

    book.append_row("Orders", ["Date", "Customer", "Amount", "Status"], style=header)
    orders = [(dt.date(2026, 11, 2), "Acme", 1250.0, "Paid"),
              (dt.date(2026, 11, 9), "Globex", 980.5, "Open"),
              (dt.date(2026, 11, 13), "Initech", -120.0, "Refund")]
    for i, (when, customer, amount, status) in enumerate(orders, start=2):
        book.set_value("Orders", f"A{i}", when, style=dates)
        book.set_value("Orders", f"B{i}", customer, style=plain)
        book.set_value("Orders", f"C{i}", amount, style=money)
        book.set_value("Orders", f"D{i}", status, style=plain)
    total = len(orders) + 2
    book.set_value("Orders", f"B{total}", "Total", style={"bold": True, "align": "right"})
    book.set_value("Orders", f"C{total}", Formula(f"SUM(C2:C{total - 1})"),
                   style={"bold": True, "number_format": "#,##0.00", "border": {"top": "double"}})

    book.set_column_width("Orders", "A", 14)
    book.set_column_width("Orders", "B", 22)
    book.set_column_width("Orders", "C:D", 12)
    book.freeze_panes("Orders", "A2")
    book.auto_filter("Orders", f"A1:D{total - 1}")
    book.define_name("Amounts", "Orders", f"C2:C{total - 1}")
    book.add_validation("Orders", f"D2:D{total - 1}", values=["Open", "Paid", "Refund"])

    assert not book.validate(target="Office2007")
    book.save(path)
    print([book.display_text("Orders", f"C{i}") for i in range(2, total)])


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "report.xlsx")
