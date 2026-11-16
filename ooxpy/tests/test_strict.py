"""Strict conformance: read into Transitional trees, write back as Strict."""

import io
import zipfile

import pytest

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument, open_document
from oox import strict as S
from oox.ns import w


def test_uri_mapping_round_trips():
    t = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    assert S.ns_to_strict(t) == "http://purl.oclc.org/ooxml/wordprocessingml/main"
    assert S.ns_to_transitional(S.ns_to_strict(t)) == t
    ext = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    assert S.ns_to_strict(ext) == "http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
    assert S.ns_to_transitional(S.ns_to_strict(ext)) == ext
    mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    assert S.ns_to_strict(mc) == mc  # unchanged in Strict
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    assert S.reltype_to_strict(rel) == "http://purl.oclc.org/ooxml/officeDocument/relationships/image"
    assert S.reltype_to_transitional(S.reltype_to_strict(rel)) == rel


def _strict_docx() -> bytes:
    doc = WordprocessingDocument.new()
    doc.body.add(w.P(w.R(w.T("Strict text"))))
    doc.conformance = "strict"
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def test_writing_strict():
    z = zipfile.ZipFile(io.BytesIO(_strict_docx()))
    xml, rels = z.read("word/document.xml"), z.read("_rels/.rels")
    assert b"http://purl.oclc.org/ooxml/wordprocessingml/main" in xml
    assert b"schemas.openxmlformats.org/wordprocessingml" not in xml
    assert b'conformance="strict"' in xml
    assert b"http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" in rels


def test_reading_strict_gives_transitional_trees():
    doc = open_document(io.BytesIO(_strict_docx()))
    assert doc.conformance == "strict"
    assert doc.body.schema_key == "w:CT_Body/w:body"  # grammar applies
    assert doc.getElementsByType(w.T)[0].text == "Strict text"
    assert doc.validate() == {}


@pytest.mark.parametrize("make", [WordprocessingDocument.new, SpreadsheetDocument.new,
                                  PresentationDocument.new])
def test_strict_round_trip_is_byte_identical_and_editable(make):
    doc = make()
    doc.conformance = "strict"
    buf = io.BytesIO()
    doc.save(buf)
    data = buf.getvalue()
    again = open_document(io.BytesIO(data))
    again.validate()  # parse everything
    out = io.BytesIO()
    again.save(out)
    a, b = zipfile.ZipFile(io.BytesIO(data)), zipfile.ZipFile(io.BytesIO(out.getvalue()))
    assert all(a.read(n) == b.read(n) for n in a.namelist())


def test_edit_keeps_strict_and_convert_back():
    doc = open_document(io.BytesIO(_strict_docx()))
    doc.body.add(w.P(w.R(w.T("added"))))
    out = io.BytesIO()
    doc.save(out)
    xml = zipfile.ZipFile(io.BytesIO(out.getvalue())).read("word/document.xml")
    assert b"purl.oclc.org/ooxml/wordprocessingml/main" in xml and b"added" in xml
    doc.conformance = "transitional"
    out = io.BytesIO()
    doc.save(out)
    z = zipfile.ZipFile(io.BytesIO(out.getvalue()))
    assert b"purl.oclc.org" not in z.read("word/document.xml")
    assert b"purl.oclc.org" not in z.read("_rels/.rels")
    assert b'conformance="strict"' not in z.read("word/document.xml")
    with pytest.raises(ValueError):
        doc.conformance = "loose"
