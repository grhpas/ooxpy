"""Excel: cell styles, number formats and display text, dates, sheet layout."""

from __future__ import annotations

import datetime as dt
import io

import pytest

from oox import SpreadsheetDocument, open_document
from oox.xlsheet import stored_to_width, width_to_stored
from oox.xlstyles import format_value, is_date_format


def reopen(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return open_document(buf)


@pytest.mark.parametrize("value,code,text", [
    (1234.5, "#,##0.00", "1,234.50"), (1234.5, "0", "1235"), (2.5, "0", "3"), (0.125, "0.00", "0.13"), (-1234.5, "#,##0.00", "-1,234.50"),
    (-98.25, "#,##0.00;[Red](#,##0.00)", "(98.25)"), (0, '0;-0;"zero"', "zero"),
    (0.125, "0.0%", "12.5%"), (123456, "0.00E+00", "1.23E+05"), (7, "000", "007"),
    (12, "0.##", "12."), (0.5, "#.00", ".50"), (1234567, '#,##0,"K"', "1,235K"),
    (1234.5, '"€"#,##0.00', "€1,234.50"), (1234.5, "[$€-407] #,##0.00", "€ 1,234.50"),
    (1 / 3, "General", "0.333333333"), (1250.5, "General", "1250.5"), (1e12, "General", "1E+12"),
    (True, "General", "TRUE"), ("abc", '"<"@">"', "<abc>"), ("abc", "0.00", "abc"),
    (dt.datetime(2026, 11, 14, 15, 7, 9), "yyyy-mm-dd hh:mm:ss", "2026-11-14 15:07:09"),
    (dt.datetime(2026, 11, 14, 15, 7, 9), "d mmmm yyyy", "14 November 2026"),
    (dt.datetime(2026, 11, 14, 15, 7, 9), "dddd", "Saturday"),
    (dt.datetime(2026, 11, 14, 15, 7, 9), "h:mm AM/PM", "3:07 PM"),
    (dt.datetime(2026, 11, 14, 15, 7, 9), "m/d/yy h:mm", "11/14/26 15:07"),
    (dt.datetime(2026, 11, 14), "yyyy\\-mm\\-dd", "2026-11-14"),   # LibreOffice's escapes
    (46340.630, "hh:mm", "15:07"),
])
def test_format_value(value, code, text):
    assert format_value(value, code) == text


def test_date_format_detection():
    assert is_date_format("yyyy-mm-dd") and is_date_format("[h]:mm") and is_date_format("x", 14)
    assert not is_date_format("0.00") and not is_date_format("[Red]0.00")
    assert not is_date_format('"days" 0') and not is_date_format("General")


def test_styles_are_shared_and_described():
    book = SpreadsheetDocument.new()
    a = book.cell_style(bold=True, fill="#d9e2f3", border="thin", align="center")
    assert book.cell_style(align="center", border="thin", fill="D9E2F3", bold=True) == a
    b = book.cell_style(number_format="#,##0.00", border={"bottom": "double"},
                        border_color="FF0000", font="Arial", size=12, color="1F3864")
    book.set_value("Sheet1", "A1", "h", style=a)
    book.set_value("Sheet1", "B1", 1.5, style=b)
    back = reopen(book)
    fa, fb = back.cell_format("Sheet1", "A1"), back.cell_format("Sheet1", "B1")
    assert (fa["bold"], fa["fill"], fa["border"], fa["align"]) == (True, "D9E2F3", "thin", "center")
    assert (fb["number_format"], fb["font"], fb["size"], fb["color"]) == \
        ("#,##0.00", "Arial", 12.0, "1F3864")
    assert fb["border"] == {"bottom": "double"}
    assert back.display_text("Sheet1", "B1") == "1.50"
    assert not book.validate(target="Office2007")


def test_style_errors():
    book = SpreadsheetDocument.new()
    with pytest.raises(TypeError):
        book.cell_style(boldd=True)
    with pytest.raises(ValueError):
        book.cell_style(fill="red")
    with pytest.raises(ValueError):
        book.cell_style(border="wavy")
    with pytest.raises(ValueError):
        book.cell_style(align="middle")


def test_unstyled_workbooks_get_no_stylesheet():
    book = SpreadsheetDocument.new()
    book.set_value("Sheet1", "A1", 1)
    _ = book.cell_format("Sheet1", "A1"), book.display_text("Sheet1", "A1")
    assert "/xl/styles.xml" not in reopen(book).package.parts


def test_dates_times_round_trip():
    book = SpreadsheetDocument.new()
    when = dt.datetime(2026, 11, 14, 15, 7, 9)
    book.set_value("Sheet1", "A1", when.date())
    book.set_value("Sheet1", "A2", when)
    book.set_value("Sheet1", "A3", when.time())
    book.append_row("Sheet1", [when.date(), "x"])
    back = reopen(book)
    assert back.value("Sheet1", "A1") == dt.datetime(2026, 11, 14)
    assert back.value("Sheet1", "A2") == when
    assert back.value("Sheet1", "A3") == dt.time(15, 7, 9)
    assert back.value("Sheet1", "A4") == dt.datetime(2026, 11, 14)
    assert isinstance(back.value("Sheet1", "A2", dates=False), float)
    assert [back.display_text("Sheet1", r) for r in ("A1", "A2", "A3")] == \
        ["2026-11-14", "2026-11-14 15:07", "15:07:09"]
    assert back.values("Sheet1")[3] == [dt.datetime(2026, 11, 14), "x"]


def test_style_range_and_kept_styles():
    book = SpreadsheetDocument.new()
    book.set_value("Sheet1", "B2", 5)
    book.style_range("Sheet1", "A1:C2", bold=True)
    assert all(book.cell_format("Sheet1", r)["bold"] for r in ("A1", "C1", "B2", "C2"))
    book.set_value("Sheet1", "B2", 6)  # style kept when not given
    assert book.cell_format("Sheet1", "B2")["bold"]
    assert book.value("Sheet1", "B2") == 6


def test_existing_stylesheet_is_extended_not_duplicated():
    from tools.officecheck.generate import _excel_shaped_workbook

    book = open_document(io.BytesIO(_excel_shaped_workbook()))
    book.set_value("Data", "C1", "When", style={"bold": True})
    c1 = book.cell_element("Data", "C1")
    assert c1.get("s") == "1"  # Excel's own bold xf
    book.set_value("Data", "C2", dt.date(2026, 11, 14))
    assert book.display_text("Data", "C2") == "2026-11-14"
    assert book.display_text("Data", "A1") == "Name"


# -- layout -------------------------------------------------------------------------------------
def test_column_widths_as_excel_shows_them():
    assert width_to_stored(20) == 20.7109375
    assert stored_to_width(20.7109375) == 20.0
    book = SpreadsheetDocument.new()
    book.set_column_width("Sheet1", "A", 20)
    book.set_column_width("Sheet1", "B:D", 12)
    book.set_column_width("Sheet1", "C", 8)  # splits the B:D run
    back = reopen(book)
    assert [back.column_width("Sheet1", c) for c in "ABCDE"] == [20.0, 12.0, 8.0, 12.0, None]
    cols = back.sheet("Sheet1").first("x:cols")
    assert [(c.get("min"), c.get("max")) for c in cols] == [("1", "1"), ("2", "2"), ("3", "3"),
                                                            ("4", "4")]
    book.set_column_width("Sheet1", "A:D", None)
    assert book.sheet("Sheet1").first("x:cols") is None
    with pytest.raises(ValueError):
        book.set_column_width("Sheet1", "A1", 3)
    assert not book.validate(target="Office2007")


def test_rows_merges_freeze_filter_names_validation():
    book = SpreadsheetDocument.new("Sales")
    book.append_row("Sales", ["Region", "Q1"])
    book.set_row_height("Sales", 1, 24)
    book.set_row_height("Sales", 5, 30)  # creates the row
    book.merge_cells("Sales", "A6:C6")
    with pytest.raises(ValueError):
        book.merge_cells("Sales", "B6:D7")
    book.freeze_panes("Sales", "B2")
    book.auto_filter("Sales", "A1:B4")
    with pytest.raises(ValueError):
        book.define_name("Q1", "Sales", "B2:B4")  # "Q1" is a cell reference, as Excel says
    book.define_name("FirstQuarter", "Sales", "B2:B4")
    book.define_name("Local", "Sales", "A1", local=True)
    book.add_validation("Sales", "C2:C4", values=["Yes", "No"])
    book.add_validation("Sales", "D2", whole=(1, 10), error="1 to 10")
    back = reopen(book)
    assert (back.row_height("Sales", 1), back.row_height("Sales", 5)) == (24.0, 30.0)
    assert back.merged_cells("Sales") == ["A6:C6"]
    assert back.frozen_panes("Sales") == "B2"
    assert back.auto_filter_range("Sales") == "A1:B4"
    assert back.defined_names() == {"Sales!_xlnm._FilterDatabase": "Sales!$A$1:$B$4",
                                    "FirstQuarter": "Sales!$B$2:$B$4",
                                    "Sales!Local": "Sales!$A$1"}
    assert back.validations("Sales") == [
        {"range": "C2:C4", "type": "list", "values": ["Yes", "No"]},
        {"range": "D2", "type": "whole", "min": "1", "max": "10"}]
    assert not book.validate(target="Office2007")
    book.freeze_panes("Sales", None)
    book.auto_filter("Sales", None)
    book.unmerge_cells("Sales", "A6:C6")
    assert (book.frozen_panes("Sales"), book.auto_filter_range("Sales"),
            book.merged_cells("Sales")) == (None, None, [])
    assert "Sales!_xlnm._FilterDatabase" not in book.defined_names()


def test_name_and_validation_errors():
    book = SpreadsheetDocument.new()
    with pytest.raises(ValueError):
        book.define_name("A1", "Sheet1", "A1")
    with pytest.raises(ValueError):
        book.add_validation("Sheet1", "A1", values=["a,b"])
    with pytest.raises(ValueError):
        book.add_validation("Sheet1", "A1")


def test_sheet_names_with_spaces_are_quoted():
    book = SpreadsheetDocument.new("My Sheet")
    book.define_name("N", "My Sheet", "A1:B2")
    assert book.defined_names()["N"] == "'My Sheet'!$A$1:$B$2"


def test_only_locale_invariant_codes_become_builtin_ids():
    """Excel draws built-in ids 14-22 and 37-40 in the viewer's regional
    variant, so those codes must be written out (Greek Excel, 2026-11-14)."""
    from oox.xlstyles import StyleSheet

    book = SpreadsheetDocument.new()
    st = StyleSheet(book)
    assert st.number_format_id("0.00") == 2 and st.number_format_id("0%") == 9
    assert st.number_format_id("@") == 49
    for code in ("#,##0.00;[Red](#,##0.00)", "mm-dd-yy", "h:mm AM/PM", "d-mmm-yy"):
        assert st.number_format_id(code) >= 164, code
    # reading a file that uses built-in ids still works
    assert st.number_format_code(14) == "mm-dd-yy" and st.number_format_code(40) == \
        "#,##0.00;[Red](#,##0.00)"


def test_autofit_columns():
    book = SpreadsheetDocument.new()
    book.set_value("Sheet1", "A1", "short")
    book.set_value("Sheet1", "B1", dt.datetime(2026, 11, 14, 15, 7, 9),
                   style={"number_format": "yyyy-mm-dd hh:mm:ss"})
    book.set_value("Sheet1", "C1", "a very long title that is merged")
    book.merge_cells("Sheet1", "C1:E1")
    book.set_value("Sheet1", "C2", "x")
    widths = book.autofit_columns("Sheet1")
    assert widths == {"A": 8.43, "B": 21.0, "C": 8.43}
    assert book.column_width("Sheet1", "B") == 21.0
    assert book.autofit_columns("Sheet1", "B") == {"B": 21.0}
