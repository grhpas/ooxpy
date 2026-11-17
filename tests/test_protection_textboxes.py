"""Document protection (Word, Excel workbook) and Word text boxes."""

from __future__ import annotations

import base64
import io

import pytest

from oox import SpreadsheetDocument, WordprocessingDocument, open_document
from oox.grammar import URI
from oox.protection import agile_password_hash, word_legacy_key, word_password_hash
from oox.records import Protection, TextBox, WorkbookProtection
from oox.units import Cm

SALT = bytes(range(16))


def _reopen(doc):
    buf = io.BytesIO()
    doc.save(buf)
    return open_document(io.BytesIO(buf.getvalue()))


# -- the hash ----------------------------------------------------------------------
def test_legacy_key_matches_the_standards_example():
    assert word_legacy_key("Example") == 0x64CEED7E  # ECMA-376 Part 4, 14.8.1.1
    assert word_legacy_key("") == 0


def test_hash_matches_one_word_2019_wrote():
    salt = base64.b64decode("pH1TDVHSfGBxkd3Q88UNhQ==")
    assert word_password_hash("johnjohn", salt, "SHA-512") == (
        "pVjR9ktO9vlxijXcMPlH+4PLwD4Xwy1aqbNQOFmWaSpvBjipNh//T8S3nBhq6HRoRVfWL6s/+NdUCPTxUr0vZw==")


def test_legacy_key_uses_15_characters_and_the_high_byte_of_wide_characters():
    assert word_legacy_key("a" * 15) == word_legacy_key("a" * 20)
    assert word_legacy_key("Ā") == word_legacy_key("\x01")  # low byte 0: the high byte


def test_agile_hash_is_salted_and_iterated():
    a = agile_password_hash("secret", SALT, "SHA-512", 10)
    assert a != agile_password_hash("secret", SALT[::-1], "SHA-512", 10)
    assert a != agile_password_hash("secret", SALT, "SHA-512", 11)
    assert len(base64.b64decode(a)) == 64


# -- Word ---------------------------------------------------------------------------
@pytest.mark.parametrize("edit", ["readOnly", "comments", "trackedChanges", "forms"])
def test_protect_round_trip(edit):
    doc = WordprocessingDocument.new()
    doc.protect(edit, password="secret", salt=SALT, spin_count=1000)
    back = _reopen(doc)
    assert back.protection() == Protection(edit=edit, enforced=True, password=True,
                                           algorithm="SHA-1", formatting=False)
    assert back.check_protection_password("secret")
    assert not back.check_protection_password("Secret")
    assert back.validate(target="Office2007") == {}
    back.unprotect()
    assert back.protection() is None and back.check_protection_password("anything")


def test_protect_writes_word_2007_attributes():
    doc = WordprocessingDocument.new()
    doc.protect("readOnly", password="secret", salt=SALT)
    el = doc.settings.find(f"{{{URI['w']}}}documentProtection")
    attrs = {k.split("}")[1]: v for k, v in el.attrib.items()}
    assert attrs["cryptAlgorithmSid"] == "4" and attrs["cryptProviderType"] == "rsaFull"
    assert attrs["salt"] == base64.b64encode(SALT).decode()
    assert "algorithmName" not in attrs and "hashValue" not in attrs  # Word 2010's set
    # salts are random unless given
    other = WordprocessingDocument.new()
    other.protect("readOnly", password="secret")
    assert other.settings.find(f"{{{URI['w']}}}documentProtection").get(
        f"{{{URI['w']}}}salt") != attrs["salt"]


def test_protect_sha512_and_no_password_and_replace():
    doc = WordprocessingDocument.new()
    doc.protect("comments")
    assert doc.protection().password is False
    assert doc.check_protection_password("")
    doc.protect("trackedChanges", password="pw", algorithm="SHA-512", salt=SALT, spin_count=10)
    assert len(doc.settings.findall(f"{{{URI['w']}}}documentProtection")) == 1
    info = doc.protection()
    assert info.edit == "trackedChanges" and info.algorithm == "SHA-512"
    assert doc.check_protection_password("pw") and not doc.check_protection_password("p")


def test_check_password_reads_word_2010_attributes():
    """Word 2010+ may write algorithmName/hashValue/saltValue/spinCount."""
    from oox.ns import w

    doc = WordprocessingDocument.new()
    salt = base64.b64decode("pH1TDVHSfGBxkd3Q88UNhQ==")
    el = w.DocumentProtection(edit="readOnly", enforcement=True)
    for k, v in {"algorithmName": "SHA-512", "spinCount": "100000",
                 "saltValue": "pH1TDVHSfGBxkd3Q88UNhQ==",
                 "hashValue": word_password_hash("johnjohn", salt, "SHA-512")}.items():
        el.set(f"{{{URI['w']}}}{k}", v)
    doc.settings.add(el, replace=True)
    assert doc.protection().algorithm == "SHA-512"
    assert doc.check_protection_password("johnjohn")
    assert not doc.check_protection_password("john")


def test_protect_rejects_bad_arguments():
    doc = WordprocessingDocument.new()
    with pytest.raises(ValueError):
        doc.protect("everything")
    with pytest.raises(ValueError):
        doc.protect("readOnly", password="")
    with pytest.raises(ValueError):
        doc.protect("readOnly", password="x", algorithm="MD5")
    assert doc.protection() is None


# -- Excel --------------------------------------------------------------------------
def test_workbook_protection_round_trip():
    book = SpreadsheetDocument.new("A")
    book.add_sheet("B")
    assert book.workbook_protection() is None
    book.protect_workbook(password="secret")
    back = _reopen(book)
    assert back.workbook_protection() == WorkbookProtection(structure=True, password=True)
    assert back.check_workbook_password("secret") and not back.check_workbook_password("x")
    assert back.validate(target="Office2007") == {}
    # x:workbookProtection sits after x:bookViews' predecessors, before x:sheets
    names = [c.tag.split("}")[1] for c in back.workbook]
    assert names.index("workbookProtection") < names.index("sheets")
    back.unprotect_workbook()
    assert back.workbook_protection() is None


def test_workbook_protection_windows_without_password():
    book = SpreadsheetDocument.new()
    book.protect_workbook(structure=False, windows=True)
    assert book.workbook_protection() == WorkbookProtection(windows=True)
    assert book.check_workbook_password("anything")


def test_check_workbook_password_reads_excel_2010_attributes():
    book = SpreadsheetDocument.new()
    book.protect_workbook()
    el = book.workbook.first("x:workbookProtection")
    el.set("workbookAlgorithmName", "SHA-512")
    el.set("workbookSpinCount", "1000")
    el.set("workbookSaltValue", base64.b64encode(SALT).decode())
    el.set("workbookHashValue", agile_password_hash("secret", SALT, "SHA-512", 1000))
    assert book.workbook_protection().password
    assert book.check_workbook_password("secret") and not book.check_workbook_password("x")


# -- text boxes ---------------------------------------------------------------------
MC, V, W = URI["mc"], URI["v"], URI["w"]


def test_text_box_has_drawingml_and_vml_copies():
    doc = WordprocessingDocument.new()
    doc.add_paragraph("Body")
    run = doc.add_text_box("Note\nSecond", width=Cm(5), height=Cm(2), x=Cm(10))
    ac = run.find(f"{{{MC}}}AlternateContent")
    assert ac is not None
    assert ac.find(f"{{{MC}}}Choice").get("Requires") == "wps"
    assert ac.find(f".//{{{MC}}}Fallback/{{{W}}}pict/{{{V}}}shape") is not None
    assert len(list(ac.iter(f"{{{W}}}txbxContent"))) == 2
    assert doc.text_boxes() == [TextBox("Note\nSecond", paragraphs=2)]
    for level in (None, "full"):
        kwargs = {"level": level} if level else {}
        assert doc.validate(**kwargs) == {}
    assert doc.validate(target="Office2007") == {}


def test_text_boxes_round_trip_and_edit_both_copies():
    doc = WordprocessingDocument.new()
    doc.add_text_box("One")
    doc.add_text_box("Two", h_relative="page", v_relative="page", wrap="topAndBottom",
                     fill=None, line=None)
    back = _reopen(doc)
    assert [b.text for b in back.text_boxes()] == ["One", "Two"]
    back.set_text_box_text(1, "Changed\nTwice")
    assert [b.text for b in back.text_boxes()] == ["One", "Changed\nTwice"]
    from oox import teletype

    copies = list(back.body.iter(f"{{{W}}}txbxContent"))[2:]
    assert all(["".join(teletype.extractText(p)) for p in c] == ["Changed", "Twice"]
               for c in copies)
    with pytest.raises(IndexError):
        back.set_text_box_text(2, "x")


def test_text_box_vml_shape_type_written_once():
    doc = WordprocessingDocument.new()
    for i in range(3):
        doc.add_text_box(f"Box {i}")
    types = list(doc.document.iter(f"{{{V}}}shapetype"))
    assert len(types) == 1 and types[0].get("id") == "_x0000_t202"
    ids = [d.get("id") for d in doc.document.iter(f"{{{URI['wp']}}}docPr")]
    assert len(set(ids)) == 3


def test_text_box_with_table_content_and_existing_paragraph():
    from oox.ns import w

    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("Anchor here")
    tbl = doc.add_table([["a", "b"], ["c", "d"]])
    doc.body.remove(tbl.element)  # a Table object is accepted as content
    run = doc.add_text_box([w.P(w.R(w.T("Above the table"))), tbl, w.P()], paragraph=p)
    assert run.getparent() is p
    assert len(list(doc.document.iter(f"{{{W}}}tbl"))) == 2  # one per copy
    assert doc.text_boxes()[0].paragraphs == 2
    assert doc.validate() == {}


def test_vml_only_text_boxes_are_read():
    doc = WordprocessingDocument.new()
    doc.add_text_box("Legacy")
    for ac in list(doc.document.iter(f"{{{MC}}}AlternateContent")):
        ac.addprevious(ac.find(f"{{{MC}}}Fallback")[0])
        ac.getparent().remove(ac)
    assert doc.text_boxes() == [TextBox("Legacy", paragraphs=1, vml_only=True)]


def test_text_box_rejects_bad_options():
    doc = WordprocessingDocument.new()
    with pytest.raises(ValueError):
        doc.add_text_box("x", wrap="tight")
    with pytest.raises(ValueError):
        doc.add_text_box("x", h_relative="sky")
    with pytest.raises(ValueError):
        doc.add_text_box("x", v_align="middle")


def test_text_box_alignment_in_both_copies():
    doc = WordprocessingDocument.new()
    run = doc.add_text_box("Right", h_align="right", h_relative="margin", v_align="center",
                           v_relative="page")
    assert [e.text for e in run.iter(f"{{{URI['wp']}}}align")] == ["right", "center"]
    style = next(run.iter(f"{{{V}}}shape")).get("style")
    assert "mso-position-horizontal:right" in style and "mso-position-vertical:center" in style
    assert doc.validate() == {}


def test_text_extraction_counts_a_text_box_once(capsys, tmp_path):
    """Found by the packaging smoke test: ``oox text`` printed a text box's
    text three times (both copies inside the anchoring paragraph, then the
    box's own paragraphs from both copies)."""
    from oox import teletype
    from oox.__main__ import main

    doc = WordprocessingDocument.new()
    p = doc.add_paragraph("Before ")
    doc.add_text_box("In box\nline 2", paragraph=p)
    doc.add_paragraph("After")
    assert teletype.extractText(doc.body) == "Before In boxline 2After"
    assert teletype.extractText(p, text_boxes=False) == "Before "
    plain = doc.add_paragraph("No boxes here")
    assert teletype.extractText(plain, text_boxes=False) == "No boxes here"
    doc.save(tmp_path / "t.docx")
    assert main(["text", str(tmp_path / "t.docx")]) == 0
    assert capsys.readouterr().out.splitlines() == [
        "Before ", "In box", "line 2", "After", "No boxes here"]
