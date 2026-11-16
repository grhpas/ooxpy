"""Examples and tutorials must keep working: run every script and every
``<code class="run">`` block of each tutorial, top to bottom."""

import html
import io
import re
import runpy
import sys
from pathlib import Path

import pytest

EXAMPLES = Path(__file__).parents[1] / "examples"


def png_bytes() -> bytes:
    from tools.officecheck.assets import image

    return image("PNG")


@pytest.mark.parametrize("tutorial", ["hello.html", "edit_existing.html", "showcase.html", "report.html",
                                      "tables_review.html", "forms_fields.html",
                                      "excel_report.html", "excel_dashboard.html",
                                      "charts.html", "slides.html"])
def test_tutorial_blocks_run(tutorial, tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "chart.png").write_bytes(png_bytes())
    source = (EXAMPLES / tutorial).read_text(encoding="utf-8")
    blocks = [html.unescape(b) for b in re.findall(r'<code class="run">(.*?)</code>', source, re.S)]
    assert blocks, "no runnable blocks found"
    namespace: dict = {}
    for i, block in enumerate(blocks):
        exec(compile(block, f"{tutorial}#block{i}", "exec"), namespace)
    printed = capsys.readouterr().out
    # spot-check the outputs the pages show
    shown = [html.unescape(o) for o in re.findall(r'<pre class="out">(.*?)</pre>', source, re.S)]
    for out in shown:
        for line in out.splitlines():
            if "…" not in line:  # elided lines are abbreviated on purpose
                assert line in printed, f"{tutorial}: shown output not produced: {line!r}"


@pytest.mark.parametrize("script,args", [
    ("hello.py", ["out.docx"]),
    ("edit_existing.py", ["outdir"]),
    ("showcase.py", ["outdir"]),
    ("report.py", ["report.docx"]),
    ("tables_review.py", ["outdir"]),
    ("forms_fields.py", ["outdir"]),
    ("excel_report.py", ["report.xlsx"]),
    ("excel_dashboard.py", ["dashboard.xlsx"]),
    ("charts.py", ["outdir"]),
    ("slides.py", ["plan.pptx"]),
])
def test_example_scripts_run(script, args, tmp_path, monkeypatch):
    if script == "showcase.py":
        pytest.importorskip("PIL")
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [script, *args])
    monkeypatch.setattr(sys, "stdout", io.StringIO())
    runpy.run_path(str(EXAMPLES / script), run_name="__main__")
