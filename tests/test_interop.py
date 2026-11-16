"""tools/interop.py on files from other producers that ship with the test
dependencies (python-docx and python-pptx templates) and on our own output."""

from pathlib import Path

import pytest

from oox import SpreadsheetDocument
from tools.interop import check


@pytest.mark.parametrize("pkg,ext", [("docx", "docx"), ("pptx", "pptx")])
def test_third_party_templates_pass_interop(pkg, ext):
    mod = pytest.importorskip(pkg)
    files = sorted((Path(mod.__file__).parent / "templates").glob(f"*.{ext}"))
    assert files
    for f in files:
        result = check(str(f))
        assert result["ok"], result["problems"]


def test_interop_reports_problems(tmp_path):
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", ["a", 1])
    path = tmp_path / "b.xlsx"
    book.save(path)
    result = check(str(path))
    assert result["ok"] and "1 sheets" in result["notes"]
    (tmp_path / "bad.xlsx").write_bytes(b"not a zip")
    bad = check(str(tmp_path / "bad.xlsx"))
    assert not bad["ok"] and bad["problems"][0].startswith("open:")
