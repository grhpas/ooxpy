"""save(target=...): refuse or strip markup an Office version can't read."""

import io

import pytest

from oox import SpreadsheetDocument, open_document
from oox.errors import TargetError
from oox.ns import x


def _book_with_2010_markup():
    book = SpreadsheetDocument.new("Data")
    book.workbook.add(x.WorkbookPr(dateCompatibility=True, date1904=False))
    book.append_row("Data", ["hello"])
    return book


def test_save_refuses_markup_the_target_cannot_read():
    book = _book_with_2010_markup()
    with pytest.raises(TargetError) as err:
        book.save(io.BytesIO(), target="Office2007")
    assert list(err.value.findings) == ["/xl/workbook.xml"]
    assert "dateCompatibility needs Office2010" in err.value.findings["/xl/workbook.xml"][0]
    assert "strip=True" in str(err.value)
    book.save(io.BytesIO(), target="Office2010")   # fine for 2010 and later


def test_save_strips_what_the_target_cannot_read():
    book = _book_with_2010_markup()
    buf = io.BytesIO()
    book.save(buf, target="Office2007", strip=True)
    assert "dateCompatibility" in book.stripped["/xl/workbook.xml"][0]
    buf.seek(0)
    again = open_document(buf)
    pr = again.workbook.first("x:workbookPr")
    assert pr is not None and pr.get("dateCompatibility") is None
    assert pr.get("date1904") == "0"                    # Office 2007 markup stays
    assert again.validate(target="Office2007") == {}
    assert again.value("Data", "A1") == "hello"


def test_clean_documents_save_unchanged_with_a_target():
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", [1, 2])
    a, b = io.BytesIO(), io.BytesIO()
    book.save(a)
    book.save(b, target="Office2007")
    assert a.getvalue() == b.getvalue() and book.stripped == {}
    with pytest.raises(ValueError, match="target must be one of"):
        book.save(io.BytesIO(), target="Office97")


def test_word_2010_renames_go_back_to_their_2007_names():
    from oox import WordprocessingDocument
    from oox.ns import w

    doc = WordprocessingDocument.new()
    doc.body.add(w.P(w.PPr(w.Ind(start=720, end=360)), w.R(w.T("indented"))))
    tbl_borders = w.TcBorders(w.Start(val="single", sz=4), w.End(val="single", sz=4))
    with pytest.raises(TargetError):
        doc.save(io.BytesIO(), target="Office2007")
    buf = io.BytesIO()
    doc.save(buf, target="Office2007", strip=True)
    assert any("renamed to w:left" in m for m in doc.stripped["/word/document.xml"])
    buf.seek(0)
    again = open_document(buf)
    ind = next(again.document.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind"))
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    assert (ind.get(W + "left"), ind.get(W + "right"), ind.get(W + "start")) == ("720", "360", None)
    # element renames (border/margin children)
    changes = tbl_borders.strip_for("Office2007")
    assert [c.tag.split("}")[1] for c in tbl_borders] == ["left", "right"]
    assert all("renamed to <w:" in m for m in changes)


def test_open_enums_accept_any_value():
    """ISO/IEC 29500-1 §17.8.3.2: characterSet is ST_String (IANA names); the
    SDK's closed list misses windows-1252, which LibreOffice writes."""
    from oox.ns import w

    cs = w.Charset(characterSet="windows-1252")
    assert cs.validate() == []
    assert w.Jc(val="center").validate() == []
    from oox.errors import InvalidAttributeValue

    with pytest.raises(InvalidAttributeValue):
        w.Jc(val="middle")                  # real enums stay closed


def test_iso_values_word_2007_accepts():
    """word-94 (Office 2007, 2026-11-15): lvlJc/table jc "start" and
    cantSplit/tblHeader "true" open without a prompt -> accepted."""
    from oox.ns import w

    assert w.LvlJc(val="start").validate() == []
    assert w.CantSplit(val="true").validate() == []
    assert w.TblHeader(val="1").validate() == []
    tbl_pr = w.TblPr(w.Jc(val="end"))
    assert tbl_pr.validate() == []
    from oox.errors import InvalidAttributeValue

    with pytest.raises(InvalidAttributeValue):
        w.LvlJc(val="both")               # still not a level justification
