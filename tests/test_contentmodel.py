"""validate(level="full"): whole content models (oox.contentmodel)."""

import glob
from pathlib import Path

import pytest

from oox import contentmodel, open_document
from oox.element import lenient
from oox.ns import a, p, w


def test_exclusive_choice():
    with lenient():
        ln = a.Ln(a.NoFill(), a.SolidFill(a.SrgbClr(val="FF0000")))
    assert ln.validate() == []                         # each child is fine on its own
    (msg,) = ln.validate(level="full")
    assert "<a:solidFill> is not allowed here (child 2 of 2)" in msg
    assert a.Ln(a.SolidFill(a.SrgbClr(val="FF0000"))).validate(level="full") == []


def test_missing_and_misplaced():
    with lenient():
        tbl = w.Tbl(w.TblPr(), w.Tr(w.Tc(w.P())))
        sp = p.Sp(p.NvSpPr(p.CNvPr(id=1, name="x"), p.CNvSpPr(), p.NvPr()))
    full = tbl.validate(level="full")
    assert full[0] == "<w:tbl>: missing required child <w:tblGrid>"
    assert "<w:tr> is not allowed here (child 2 of 2); expected <w:tblGrid>" in full[1]
    assert sp.validate(level="full") == ["<p:sp>: missing required child <p:spPr>"]  # no repeat


def test_unknown_and_alternate_content_children_are_left_to_basic_checks():
    from lxml import etree

    rpr = w.RPr(w.B())
    etree.SubElement(rpr, "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent")
    assert [m for m in rpr.validate(level="full") if "AlternateContent" in m] == []


def test_compiled_models_are_cached_and_unknown_types_skip():
    key = "w:CT_Tbl/w:tbl"
    assert contentmodel.compiled(key) is contentmodel.compiled(key)
    assert contentmodel.compiled("w:CT_Nope/w:nope") is None
    assert contentmodel.check("w:CT_Nope/w:nope", "w:nope", ["w:p"]) is None


def test_level_argument():
    with pytest.raises(ValueError, match="level"):
        w.P().validate(level="paranoid")


def test_corpus_and_third_party_files_are_clean_at_full_level(tmp_path):
    from tools.officecheck.generate import generate

    generate(tmp_path)
    files = [f for f in glob.glob(str(tmp_path / "*" / "*.*x"))
             if "control" not in f and "observe" not in f]
    for pkg in ("docx", "pptx"):
        mod = pytest.importorskip(pkg)
        files += glob.glob(str(Path(mod.__file__).parent / "templates" / f"*.{pkg}"))
    assert len(files) > 45
    for f in files:
        doc = open_document(f)
        assert doc.validate(level="full") == doc.validate(), f
