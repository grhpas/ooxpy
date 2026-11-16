"""Excel step 2: conditional formatting, tables, comments, hyperlinks, shared
strings, print setup, sheet options and protection."""

from __future__ import annotations

import io
import zipfile

import pytest

from oox import SpreadsheetDocument, open_document
from oox.spreadsheet import Formula
from oox.xlprint import _hf_code, _hf_text, excel_password_hash


def reopen(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return open_document(buf)


def data_book():
    book = SpreadsheetDocument.new("S")
    book.append_row("S", ["Region", "Sales", "Margin"])
    for row in [("North", 120, 14), ("South", 30, 21), ("North", 80, 35)]:
        book.append_row("S", list(row))
    return book


# -- conditional formatting -----------------------------------------------------------------
def test_every_rule_kind_validates_and_reads_back():
    book = data_book()
    red = {"fill": "FFC7CE", "color": "9C0006"}
    book.conditional_format("S", "B2:B4", "cell", operator=">=", value=100, style=red)
    book.conditional_format("S", "B2:B4", "cell", operator="not between", values=(10, 50),
                            style={"bold": True})
    book.conditional_format("S", "B2:B4", "cell", operator="=", value="North", style=red)
    book.conditional_format("S", "C2:C4", "color_scale", colors=["FFFFFF", "63BE7B"])
    book.conditional_format("S", "C2:C4", "data_bar")
    book.conditional_format("S", "C2:C4", "icon_set", icons="5Rating", reverse=True)
    book.conditional_format("S", "B2:B4", "bottom", n=10, percent=True, style={"italic": True})
    book.conditional_format("S", "B2:B4", "above_average", style={"fill": "C6EFCE"})
    book.conditional_format("S", "A2:A4", "unique", style={"color": "0070C0"})
    book.conditional_format("S", "A2:A4", "text", text='Sou"th', how="begins", style=red)
    book.conditional_format("S", "A2:C4", "blanks", style={"fill": "EEEEEE"})
    book.conditional_format("S", "A2:C4 E2:E4", "formula", formula="=$C2>$B2",
                            style={"border": "thin", "border_color": "FF0000"}, stop=True)
    assert not book.validate(target="Office2007")
    rules = reopen(book).conditional_formats("S")
    assert [r["priority"] for r in rules] == list(range(1, 13))
    by_type = {r["type"]: r for r in rules}
    assert rules[0]["operator"] == "greaterThanOrEqual" and rules[0]["style"] == \
        {"color": "9C0006", "fill": "FFC7CE"}
    assert rules[1]["formulas"] == ["10", "50"] and rules[2]["formulas"] == ['"North"']
    assert by_type["colorScale"]["colors"] == ["FFFFFF", "63BE7B"]
    assert by_type["iconSet"]["icons"] == "5Rating"
    assert by_type["beginsWith"]["formulas"] == ['LEFT(A2,6)="Sou""th"']
    assert by_type["expression"]["range"] == "A2:C4 E2:E4"
    assert by_type["expression"]["formulas"] == ["$C2>$B2"]


def test_highlight_styles_are_shared_dxfs():
    book = data_book()
    for ref in ("B2:B4", "C2:C4"):
        book.conditional_format("S", ref, "cell", operator=">", value=1, style={"bold": True})
    styles = book.related_parts(book.main_part, "WorkbookStylesPart")[0].element
    assert len(styles.first("x:dxfs")) == 1


def test_rule_errors():
    book = data_book()
    with pytest.raises(ValueError):
        book.conditional_format("S", "B2:B4", "sparkle")
    with pytest.raises(ValueError):
        book.conditional_format("S", "B2:B4", "cell", operator="~", value=1, style={"bold": True})
    with pytest.raises(ValueError):
        book.conditional_format("S", "B2:B4", "top")  # needs a style
    with pytest.raises(ValueError):
        book.conditional_format("S", "B2:B4", "color_scale", colors=["FF0000"])
    with pytest.raises(TypeError):
        book.conditional_format("S", "B2:B4", "duplicates", style={"size": 12})


# -- tables -----------------------------------------------------------------------------------
def test_table_with_totals_and_structured_reference():
    book = data_book()
    book.set_value("S", "D1", None)
    book.add_table("S", "A1:D4", name="Sales", totals={"Region": "Total", "Sales": "sum",
                                                      "Margin": "average", "Column4": "count"})
    book.set_value("S", "F1", Formula("SUM(Sales[Sales])", 230))
    back = reopen(book)
    (t,) = back.tables("S")
    assert t == {"name": "Sales", "ref": "A1:D5", "columns": ["Region", "Sales", "Margin",
                 "Column4"], "style": "TableStyleMedium2", "totals": True, "header": True}
    assert back.value("S", "D1") == "Column4"  # header written into the cell
    assert [back.value("S", r) for r in ("A5", "B5", "C5", "D5")] == [
        "Total", 230, pytest.approx(23.3333333), 0]
    assert back.value("S", "B5", formulas=True) == Formula("SUBTOTAL(109,Sales[Sales])")
    assert back.value("S", "F1") == 230
    assert not book.validate(target="Office2007")


def test_table_names_and_overlaps():
    book = data_book()
    book.add_table("S", "A1:C4")
    assert book.tables("S")[0]["name"] == "Table1"
    with pytest.raises(ValueError):
        book.add_table("S", "B3:D6")
    book.set_value("S", "F1", "h")
    book.set_value("S", "F2", 1)
    with pytest.raises(ValueError):
        book.add_table("S", "F1:F2", name="table1")  # names are case-insensitive
    with pytest.raises(ValueError):
        book.add_table("S", "F1:F2", name="A1")
    with pytest.raises(ValueError):
        book.add_table("S", "F1:F1")
    with pytest.raises(ValueError):
        book.add_table("S", "F1:F2", totals={"Nope": "sum"})
    book.add_table("S", "F1:F2", name="Other", style=None)
    assert [t["style"] for t in book.tables("S")] == ["TableStyleMedium2", None]


def test_duplicate_headers_get_unique_names():
    book = SpreadsheetDocument.new("S")
    book.append_row("S", ["x", "X", ""])
    book.append_row("S", [1, 2, 3])
    book.add_table("S", "A1:C2")
    assert book.tables("S")[0]["columns"] == ["x", "Column2", "Column3"]


# -- comments and hyperlinks ---------------------------------------------------------------------
def test_comments_on_two_sheets_with_replacement():
    book = data_book()
    book.add_sheet("T")
    book.add_comment("S", "B2", "first", author="Ada")
    book.add_comment("S", "A1", "visible\nnote", visible=True)
    book.add_comment("S", "B2", "second", author="Ada")
    book.add_comment("T", "C3", "other sheet", author="Bob")
    back = reopen(book)
    assert back.comments("S") == {"A1": ("oox", "visible\nnote"), "B2": ("Ada", "second")}
    assert back.comments("T") == {"C3": ("Bob", "other sheet")}
    z = zipfile.ZipFile(io.BytesIO(_bytes(book)))
    vml = z.read("xl/drawings/vmldrawing1.vml").decode()
    assert vml.count("<v:shape ") == 2 and vml.count("<x:Visible/>") == 1
    vml2 = z.read("xl/drawings/vmldrawing2.vml").decode()
    assert 'data="2"' in vml2 and "_x0000_s2049" in vml2  # its own shape-id block
    assert not book.validate(target="Office2007")


def _bytes(book) -> bytes:
    buf = io.BytesIO()
    book.save(buf)
    return buf.getvalue()


def test_comment_added_to_excel_style_vml_keeps_its_text():
    book = data_book()
    book.add_comment("S", "B2", "x")
    part = book.related_parts(book.part_of(book.sheet("S")), "VmlDrawingPart")[0]
    part._blob = part._blob.replace(b"</xml>", b"<!-- Excel's own <br> stays --></xml>")
    book.add_comment("S", "C3", "y")
    assert b"<br> stays" in part._blob and part._blob.count(b"<v:shape ") == 2


def test_hyperlinks_external_and_internal():
    book = data_book()
    book.add_sheet("T")
    book.add_hyperlink("S", "A6", "https://example.com/", text="Example", tooltip="tip")
    book.add_hyperlink("S", "A7", "#T!A1", text="Go")
    book.add_hyperlink("S", "A6", "https://example.org/", text="Replaced")
    back = reopen(book)
    assert back.hyperlinks("S") == {"A6": "https://example.org/", "A7": "#T!A1"}
    assert back.value("S", "A6") == "Replaced"
    assert back.cell_format("S", "A6")["underline"] is True
    assert not book.validate(target="Office2007")


# -- shared strings --------------------------------------------------------------------------------
def test_share_strings_dedupes_and_keeps_values():
    book = data_book()
    book.add_sheet("T")
    book.append_row("T", ["North", "Other"])
    assert book.share_strings() == 8  # 3 headers + 3 regions on S, 2 on T
    assert book.share_strings() == 0
    sst = book.related_parts(book.main_part, "SharedStringTablePart")[0].element
    assert sst.get("count") == "8" and sst.get("uniqueCount") == "6"
    back = reopen(book)
    assert back.values("S")[1] == ["North", 120, 14] and back.values("T") == [["North", "Other"]]
    assert not book.validate(target="Office2007")


# -- print setup, sheet options, protection -------------------------------------------------------
def test_header_footer_codes():
    code = _hf_code({"left": "R&D", "center": "{sheet}", "right": "Page {page} of {pages}"})
    assert code == "&LR&&D&C&A&RPage &P of &N"
    assert _hf_text(code) == {"left": "R&D", "center": "{sheet}",
                              "right": "Page {page} of {pages}"}
    with pytest.raises(ValueError):
        _hf_code({"middle": "x"})


def test_page_setup_round_trip():
    book = data_book()
    book.page_setup("S", orientation="landscape", paper="A4", fit_width=1, scale=90,
                    print_area="A1:C4", repeat_rows="1:1", repeat_cols="A:A",
                    header="Report", footer={"right": "Page {page}"}, gridlines=True,
                    margins={"left": 0.5}, center_horizontally=True)
    info = reopen(book).page_setup_info("S")
    assert (info["orientation"], info["paper"], info["fit_width"], info["scale"]) == \
        ("landscape", "a4", 1, 90)
    assert info["margins"]["left"] == 0.5 and info["margins"]["top"] == 0.75
    assert info["print_area"] == "S!$A$1:$C$4"
    assert info["print_titles"] == "S!$A:$A,S!$1:$1"
    assert info["header"] == {"center": "Report"} and info["footer"] == {"right": "Page {page}"}
    assert info["gridlines"] is True
    assert not book.validate(target="Office2007")
    with pytest.raises(ValueError):
        book.page_setup("S", orientation="sideways")
    with pytest.raises(ValueError):
        book.page_setup("S", paper="A9")


def test_sheet_options_hidden_and_protection():
    assert excel_password_hash("test") == "CBEB"
    assert excel_password_hash("secret") == "DAA7"  # LibreOffice wrote the same
    book = data_book()
    book.add_sheet("Lookup")
    book.sheet_options("S", tab_color="#2f5496", zoom=120, gridlines=False)
    book.sheet_options("Lookup", hidden="very")
    with pytest.raises(ValueError):
        book.sheet_options("S", hidden=True)  # the last visible sheet
    book.protect_sheet("S", password="secret", allow=("sort",))
    back = reopen(book)
    assert back.sheet_options_info("S") == {"tab_color": "2F5496", "hidden": False,
                                            "protected": True, "zoom": 120, "gridlines": False}
    assert back.sheet_options_info("Lookup")["hidden"] == "very"
    prot = back.sheet("S").first("x:sheetProtection")
    assert prot.get("password") == "DAA7" and prot.get("sort") in ("0", "false")
    back.unprotect_sheet("S")
    back.sheet_options("Lookup", hidden=False)
    back.sheet_options("S", tab_color="")
    assert back.sheet_options_info("S")["protected"] is False
    assert back.sheet_options_info("Lookup")["hidden"] is False
    assert back.sheet_options_info("S")["tab_color"] is None
    assert not book.validate(target="Office2007")


def test_hiding_the_active_sheet_moves_the_active_tab():
    book = SpreadsheetDocument.new("A")
    book.add_sheet("B")
    book.sheet_options("A", hidden=True)
    view = book.workbook.first("x:bookViews").first("x:workbookView")
    assert view.get("activeTab") == "1"
