import datetime as dt
import io
import zipfile

import pytest

from oox import SpreadsheetDocument
from oox.ns import x
from oox.spreadsheet import (
    CellError,
    Formula,
    cell,
    cell_ref,
    column_letter,
    excel_serial,
    from_excel_serial,
    split_ref,
)

openpyxl = pytest.importorskip("openpyxl")


def test_a1_helpers():
    assert [column_letter(n) for n in (1, 26, 27, 52, 703, 16384)] == \
        ["A", "Z", "AA", "AZ", "AAA", "XFD"]
    assert split_ref("$AB$12") == (28, 12) and cell_ref(28, 12) == "AB12"
    with pytest.raises(ValueError):
        split_ref("12A")
    assert excel_serial(dt.date(2024, 1, 1)) == 45292
    assert excel_serial(dt.datetime(2024, 1, 1, 12)) == 45292.5


def test_cell_value_mapping():
    assert cell("A1", 3).first("x:v").text == "3"
    assert cell("A1", True).get("t") == "b"
    assert cell("A1", "hi").get("t") == "inlineStr"
    assert cell("A1", Formula("SUM(B1:B3)")).first("x:f").text == "SUM(B1:B3)"
    with pytest.raises(TypeError, match="excel_serial"):
        cell("A1", dt.date.today())


def test_workbook_roundtrip_through_openpyxl(tmp_path):
    book = SpreadsheetDocument.new()
    book.append_row("Sheet1", ["Item", "Qty", "Price", "Total"])
    book.append_row("Sheet1", ["Pens", 3, 1.5, Formula("B2*C2")])
    notes = book.add_sheet("Notes")
    book.append_row(notes, ["  padded  ", False, None, -7])
    path = tmp_path / "b.xlsx"
    book.save(path)

    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["Sheet1", "Notes"]
    assert [[c.value for c in r] for r in wb["Sheet1"].iter_rows()] == \
        [["Item", "Qty", "Price", "Total"], ["Pens", 3, 1.5, "=B2*C2"]]
    assert [c.value for c in next(wb["Notes"].iter_rows())] == ["  padded  ", False, None, -7]
    for part in book.package.parts.values():
        assert part.element.validate() == [], part.partname


def test_sheet_name_rules():
    book = SpreadsheetDocument.new()
    for bad in ("", "a" * 32, "x/y", "Sheet1", "SHEET1"):
        with pytest.raises(ValueError):
            book.add_sheet(bad)


def test_ambiguous_cell_is_checked_once_attached():
    book = SpreadsheetDocument.new()
    # x:c exists in calcChain (t: bool) and worksheets (t: enum)
    with pytest.raises(Exception, match="not one of"):
        book.sheet_data("Sheet1").add(x.Row(x.C(t="maybe"), r=1))


def test_reopen_and_extend(tmp_path):
    path = tmp_path / "b.xlsx"
    book = SpreadsheetDocument.new()
    book.append_row("Sheet1", [1, 2])
    book.save(path)
    again = SpreadsheetDocument.open(path)
    again.append_row("Sheet1", [3, 4])
    again.save(path)
    rows = [[c.value for c in r] for r in openpyxl.load_workbook(path).active.iter_rows()]
    assert rows == [[1, 2], [3, 4]]


# -- reading files written by other producers ------------------------------------

_NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' \
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def _excel_like_xlsx() -> bytes:
    """A workbook shaped like Excel's output: shared strings (with rich text,
    phonetic runs and _xHHHH_ escapes), cached formula results, an error."""
    files = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
        "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": f'<styleSheet {_NS}><fonts count="1"><font/></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>'
        '<cellXfs count="4"><xf/><xf/><xf/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>',
        "xl/workbook.xml": f'<workbook {_NS}><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/sharedStrings.xml": f'<sst {_NS} count="3" uniqueCount="3">'
        '<si><t>plain</t></si>'
        '<si><r><t xml:space="preserve">rich </t></r><r><rPr><b/></rPr><t>bold</t></r><rPh sb="0" eb="1"><t>IGNORED</t></rPh></si>'
        '<si><t>a_x000D_b</t></si></sst>',
        "xl/worksheets/sheet1.xml": f'<worksheet {_NS}><sheetData>'
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s" s="3"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>'
        '<row r="3"><c r="A3"><f>1+1</f><v>2</v></c><c r="B3" t="e"><f>1/0</f><v>#DIV/0!</v></c>'
        '<c r="C3" t="str"><f>"x"&amp;"y"</f><v>xy</v></c><c r="D3" t="b"><v>0</v></c><c r="E3"><v>45292.5</v></c></row>'
        '</sheetData></worksheet>',
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, text in files.items():
            z.writestr(name, text)
    return buf.getvalue()


def test_reads_excel_style_workbook():
    book = SpreadsheetDocument.open(io.BytesIO(_excel_like_xlsx()))
    assert book.cells("Data") == {"A1": "plain", "B1": "rich bold", "C1": "a\rb",
                                  "A3": 2, "B3": CellError("#DIV/0!"), "C3": "xy",
                                  "D3": False, "E3": 45292.5}
    assert book.value("Data", "B3") == "#DIV/0!" and isinstance(book.value("Data", "B3"), CellError)
    assert book.value("Data", "A2") is None and book.value("Data", "Z9") is None
    assert book.value("Data", "A3", formulas=True) == Formula("1+1")
    assert from_excel_serial(book.value("Data", "E3")) == dt.datetime(2024, 1, 1, 12)
    assert book.values("Data")[1] == [None] * 5  # missing row 2 is filled in


def test_set_value_inserts_in_order_and_keeps_style(tmp_path):
    book = SpreadsheetDocument.open(io.BytesIO(_excel_like_xlsx()))
    book.set_value("Data", "B1", 99)  # replaces a styled shared-string cell
    book.set_value("Data", "A2", "new row between 1 and 3")
    book.set_value("Data", "B3", 7)
    assert book.cell_element("Data", "B1").get("s") == "3"
    refs = [c.get("r") for c in book.sheet_data("Data").iter("{*}c")]
    assert refs == ["A1", "B1", "C1", "A2", "A3", "B3", "C3", "D3", "E3"]
    path = tmp_path / "e.xlsx"
    book.save(path)
    ws = openpyxl.load_workbook(path).active
    assert (ws["A2"].value, ws["B3"].value, ws["A1"].value) == ("new row between 1 and 3", 7, "plain")
    # B1 kept style 3, a date format in this fixture: the independent reader
    # shows our 99 as a date -- proof the style survived
    assert ws["B1"].is_date and ws["B1"].value == from_excel_serial(99)
