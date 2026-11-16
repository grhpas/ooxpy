"""python -m oox: validate, info, text, downgrade."""

import json

import pytest

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument
from oox.__main__ import main
from oox.ns import x


@pytest.fixture
def files(tmp_path):
    doc = WordprocessingDocument.new()
    doc.add_heading("Report", 1)
    doc.add_paragraph("Hello world")
    doc.save(tmp_path / "a.docx")
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", ["Region", "Q1"])
    book.append_row("Data", ["North", 120])
    book.save(tmp_path / "b.xlsx")
    old = SpreadsheetDocument.new("Data")
    old.workbook.add(x.WorkbookPr(dateCompatibility=True))
    old.save(tmp_path / "c.xlsx")
    deck = PresentationDocument.new()
    s = deck.add_slide("Title Only")
    deck.set_placeholder_text(s, "title", "Agenda")
    deck.set_notes(s, "Say hello")
    deck.save(tmp_path / "d.pptx")
    return tmp_path


def test_validate(files, capsys):
    assert main(["validate", str(files / "a.docx")]) == 0
    assert "no problems" in capsys.readouterr().out
    assert main(["validate", str(files / "c.xlsx"), "--target", "Office2007"]) == 1
    out = capsys.readouterr().out
    assert "dateCompatibility needs Office2010" in out and "1 problem(s)" in out
    assert main(["validate", str(files / "c.xlsx"), "--target", "Office2007", "--json"]) == 1
    assert list(json.loads(capsys.readouterr().out)) == ["/xl/workbook.xml"]


def test_info(files, capsys):
    assert main(["info", str(files / "b.xlsx"), "--json"]) == 0
    info = json.loads(capsys.readouterr().out)
    assert info["kind"] == "SpreadsheetDocument" and info["sheets"]["Data"]["rows"] == 2
    assert main(["info", str(files / "d.pptx")]) == 0
    out = capsys.readouterr().out
    assert "PresentationDocument" in out and "'layout': 'Title Only'" in out and "'notes': True" in out


def test_text(files, capsys):
    main(["text", str(files / "a.docx")])
    assert capsys.readouterr().out.splitlines()[-2:] == ["Report", "Hello world"]
    main(["text", str(files / "b.xlsx")])
    assert capsys.readouterr().out.splitlines() == ["# Data", "Region\tQ1", "North\t120"]
    main(["text", str(files / "d.pptx")])
    assert capsys.readouterr().out.splitlines() == ["# Slide 1", "Agenda", "[notes] Say hello"]


def test_downgrade(files, capsys):
    out_file = files / "c2007.xlsx"
    assert main(["downgrade", str(files / "c.xlsx"), str(out_file)]) == 0
    assert "1 change(s)" in capsys.readouterr().out
    assert main(["validate", str(out_file), "--target", "Office2007"]) == 0


def test_unreadable_file_exits_2(tmp_path, capsys):
    (tmp_path / "x.docx").write_bytes(b"not a zip")
    with pytest.raises(SystemExit) as err:
        main(["info", str(tmp_path / "x.docx")])
    assert err.value.code == 2 and "cannot open" in capsys.readouterr().err
