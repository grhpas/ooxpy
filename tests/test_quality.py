"""Guards that don't fit elsewhere: syntax floor, clear errors on odd structure."""

import ast
import io
from pathlib import Path

import pytest

from oox import PackageError, PresentationDocument, SpreadsheetDocument, open_document
from oox.ns import p, w

ROOT = Path(__file__).parents[1]


def test_syntax_floor():
    """The package and tools parse as Python 3.10; Windows-side scripts as 3.8."""
    files = [*ROOT.joinpath("src").rglob("*.py"), *ROOT.joinpath("tools").rglob("*.py"),
             *ROOT.joinpath("examples").glob("*.py")]
    for f in files:
        floor = (3, 8) if f.name in ("check_office.py", "report.py") else (3, 10)
        ast.parse(f.read_text(encoding="utf-8"), feature_version=floor)


def test_need_names_the_missing_child():
    assert w.P(w.PPr()).need("w:pPr") is not None
    with pytest.raises(PackageError, match=r"<w:p> has no <w:pPr>"):
        w.P().need("w:pPr")


def test_foreign_master_without_layout_list_gets_one():
    deck = PresentationDocument.new()
    master = deck.masters[0].element
    master.remove(master.first("p:sldLayoutIdLst"))  # optional in the schema
    deck.add_layout(p.SldLayout(p.CSld(p.SpTree(
        p.NvGrpSpPr(p.CNvPr(id=1, name=""), p.CNvGrpSpPr(), p.NvPr()), p.GrpSpPr()), name="Extra")))
    assert master.first("p:sldLayoutIdLst") is not None
    assert "Extra" in deck.layout_names


def test_broken_sheet_relationship_is_a_clear_error():
    book = SpreadsheetDocument.new()
    book.main_part.rels.items.clear()
    buf = io.BytesIO()
    book.save(buf)
    again = open_document(io.BytesIO(buf.getvalue()))
    with pytest.raises(PackageError, match="broken relationship"):
        again.sheet("Sheet1")


def test_ci_matrix_matches_python_floor():
    yaml = pytest.importorskip("yaml")
    import re

    floor = re.search(r'requires-python = ">=(\d+\.\d+)"',
                      ROOT.joinpath("pyproject.toml").read_text()).group(1)
    ci = yaml.safe_load(ROOT.joinpath(".github/workflows/ci.yml").read_text())
    assert ci["jobs"]["test"]["strategy"]["matrix"]["python"][0] == floor
    office = yaml.safe_load(ROOT.joinpath(".github/workflows/office-check.yml").read_text())
    assert {"corpus", "office"} <= set(office["jobs"])
