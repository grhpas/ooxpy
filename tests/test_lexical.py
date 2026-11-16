"""Attribute value shapes (from SDK validators) and Office-version targeting."""

import pytest

from oox import InvalidAttributeValue, SpreadsheetDocument, WordprocessingDocument
from oox.lexical import check
from oox.ns import a, w, w14, x


@pytest.mark.parametrize("make,good", [
    (lambda v: w.Spacing(after=v), [0, 240, "1.5in", "12pt"]),
    (lambda v: w.Color(val=v), ["auto", "1F3864", "ff0000"]),
    (lambda v: w.P(rsidR=v), ["00AB12CD"]),
    (lambda v: a.Off(x=v, y=0), [0, -914400]),
])
def test_valid_values_pass(make, good):
    for v in good:
        make(v)


@pytest.mark.parametrize("make,bad,expected", [
    (lambda v: w.Spacing(after=v), "lots", "a measure like 1.5in"),
    (lambda v: w.Spacing(after=v), -5, "integer >= 0"),
    (lambda v: w.Color(val=v), "red", "6 hex digits"),
    (lambda v: w.Color(val=v), "1F38", "6 hex digits"),
    (lambda v: w.P(rsidR=v), "XYZ", "8 hex digits"),
])
def test_invalid_values_are_explained(make, bad, expected):
    with pytest.raises(InvalidAttributeValue, match=expected):
        make(bad)


def test_check_primitives():
    num = (("num", False, 0, 10, None, None),)
    assert check(num, "5") is None and check(num, "11") and check(num, "x")
    union = (("union", (("num", False, None, None, None, None), ("str", "auto", None, None, None))),)
    assert check(union, "auto") is None and check(union, "-3") is None and check(union, "a")


def test_validate_reports_bad_values_in_loaded_trees():
    p = w.P()
    color = w.Color(val="000000")
    p.append(w.R(w.RPr(color)))
    color.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "blue")  # raw lxml
    assert any("6 hex digits" in m for m in p.validate())


def test_target_flags_newer_markup_in_old_namespaces():
    book = SpreadsheetDocument.new()
    book.workbook.add(x.WorkbookPr(dateCompatibility=True))  # added in Office 2010
    report = book.validate(target="Office2007")
    assert report == {"/xl/workbook.xml": [
        "<x:workbookPr>: attribute dateCompatibility needs Office2010 (not understood by Office2007)"]}
    assert book.validate(target="Office2010") == {}
    assert book.validate() == {}


def test_target_accepts_ignorable_namespaces():
    doc = WordprocessingDocument.new()
    doc.body.add(w.P(w.R(w.RPr(w14.Ligatures(val="standard")), w.T("x")), paraId="00AB12CD"))
    assert doc.validate(target="Office2007") == {}  # w14 is declared mc:Ignorable on save
    with pytest.raises(ValueError, match="target must be one of"):
        doc.document.validate(target="Office2003")
