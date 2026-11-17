"""Interop with the other Python OOXML libraries, both ways: files they write
pass oox's interop check (open, unchanged save, validation, readers, edit), and
files oox writes read back with them as written."""

from __future__ import annotations

import pytest

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument
from oox.units import Cm, Inches
from tools.interop import check
from tools.writers import WRITERS, _png, build, readback


@pytest.mark.parametrize("module", sorted(WRITERS))
def test_files_written_by_other_libraries_pass_interop(module, tmp_path):
    pytest.importorskip(module)
    files = WRITERS[module](tmp_path)
    assert files
    for f in files:
        result = check(str(f))
        assert result["ok"], (f.name, result["problems"])
        assert result["notes"][0] == "validate: 0 basic, 0 more at full", (f.name, result)


def test_openpyxl_chart_values_come_from_the_cells(tmp_path):
    pytest.importorskip("openpyxl")
    from oox import open_document

    path = next(f for f in build(tmp_path) if f.name == "openpyxl_features.xlsx")
    book = open_document(path)
    charts = book.charts("Sales")
    assert [c.kind for c in charts] == ["column", "line", "pie"]
    assert [s.name for s in charts[0].series] == ["North", "South"]
    assert charts[0].series[1].values == [7, 9, 11, 8]
    assert charts[0].categories == ["M1", "M2", "M3", "M4"]


def test_xlsxwriter_secondary_axis_and_combination_are_read(tmp_path):
    pytest.importorskip("xlsxwriter")
    from oox import open_document

    path = next(f for f in build(tmp_path) if f.name == "xlsxwriter_features.xlsx")
    charts = open_document(path).charts("Data")
    line, combo = charts[1], charts[2]
    assert [s.axis for s in line.series] == ["primary", "secondary"]
    assert [(s.kind, s.axis) for s in combo.series] == [
        ("column", "primary"), ("line", "secondary")]


def _save(doc, path):
    doc.save(path)
    return path


def test_python_docx_reads_oox_output(tmp_path):
    pytest.importorskip("docx")
    doc = WordprocessingDocument.new()
    doc.properties.title = "From oox"
    doc.add_heading("Heading", 1)
    doc.add_paragraph("Body text")
    doc.add_table([["a", "b"], ["1", "2"]])
    doc.add_picture(_png(), width=Cm(2))
    doc.add_section(landscape=True)
    doc.add_paragraph("Second section")
    got = readback(_save(doc, tmp_path / "oox.docx"))
    assert got["paragraphs"][:2] == ["Heading", "Body text"]
    assert got["styles"][:2] == ["Heading 1", "Normal"]
    assert got["tables"] == [[["a", "b"], ["1", "2"]]]
    assert got["pictures"] == 1 and got["sections"] == 2
    assert got["title"] == "From oox"


def test_openpyxl_reads_oox_output(tmp_path):
    pytest.importorskip("openpyxl")
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", ["k", "v", "w"])
    for row in (["a", 1, 2.5], ["b", 2, 3.5]):
        book.append_row("Data", row)
    book.merge_cells("Data", "E1:F1")
    book.add_table("Data", "A1:C3", name="Facts")
    book.add_chart("Data", "column", "A1:C3", kinds={"w": "line"}, secondary=["w"])
    book.add_sheet("Other")
    book.append_row("Other", [True, None, "x"])
    got = readback(_save(book, tmp_path / "oox.xlsx"))
    assert got["sheets"] == ["Data", "Other"]
    assert [r[:3] for r in got["values"]["Data"]] == [["k", "v", "w"], ["a", 1, 2.5],
                                                       ["b", 2, 3.5]]
    assert got["values"]["Other"] == [[True, None, "x"]]
    assert got["merged"]["Data"] == ["E1:F1"]
    assert got["tables"]["Data"] == ["Facts"]


def test_python_pptx_reads_oox_output(tmp_path):
    pytest.importorskip("pptx")
    deck = PresentationDocument.new(layouts="basic")
    s1 = deck.add_slide("Title Slide")
    deck.set_placeholder_text(s1, "ctrTitle", "From oox")
    s2 = deck.add_slide("Blank")
    deck.add_table(s2, [["h1", "h2"], ["1", "2"]], x=Inches(1), y=Inches(1), width=Inches(4))
    deck.add_chart(s2, "column", ["a", "b"], {"S": [1, 2]})
    deck.add_picture(s2, _png(), x=Inches(1), y=Inches(4), width=Inches(1))
    deck.set_notes(s2, "Speaker notes")
    got = readback(_save(deck, tmp_path / "oox.pptx"))
    assert got["texts"][0] == ["From oox", ""]          # the empty subtitle placeholder
    assert got["tables"] == [0, 1] and got["charts"] == [0, 1]
    assert got["notes"] == [None, "Speaker notes"]
    assert "Blank" in got["layouts"]


def test_the_libraries_read_the_office_corpus(tmp_path):
    """Every corpus case but the malformed controls opens in the other libraries."""
    pytest.importorskip("docx")
    pytest.importorskip("openpyxl")
    pytest.importorskip("pptx")
    from tools.officecheck import generate

    failures = []
    ext = {"word": "docx", "excel": "xlsx", "powerpoint": "pptx"}
    for case in generate.cases():
        if case.expect_status != ("ok",):        # the malformed controls
            continue
        if case.requires:
            try:
                __import__(case.requires)
            except ImportError:
                continue
        path = tmp_path / f"{case.id}.{ext[case.app]}"
        path.write_bytes(case.build())
        try:
            readback(path)
        except Exception as e:
            failures.append(f"{case.id}: {type(e).__name__}: {e}")
    assert failures == []


def test_build_skips_missing_libraries(tmp_path, monkeypatch):
    import builtins

    real = builtins.__import__

    def fake(name, *a, **k):
        if name == "xlsxwriter":
            raise ImportError(name)
        return real(name, *a, **k)

    monkeypatch.setattr(builtins, "__import__", fake)
    names = [f.name for f in build(tmp_path)]
    assert not any(n.startswith("xlsxwriter") for n in names)
