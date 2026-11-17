"""The documentation is tested like code.

- Every ```python block in the guides runs, in order, in one namespace per page,
  inside a temporary directory that has ``logo.png`` and ``chart.png``.  A block
  preceded by ``<!-- no-run -->`` is shown but not run (e.g. it needs a file the
  reader has).
- The API reference (docs/reference) matches the docstrings
  (``python -m tools.apidocs``) and places every public method.
- ``mkdocs build --strict`` succeeds (when MkDocs is installed).
"""

from __future__ import annotations

import os
import re
import shutil
import struct
import subprocess
import sys
import zlib
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
RUN = ["index.md", "getting-started.md", "concepts.md", "guide/word.md", "guide/excel.md",
       "guide/powerpoint.md", "guide/charts.md", "guide/existing-files.md", "guide/reading.md",
       "guide/validation.md", "cookbook.md", "odfpy.md", "faq.md"]
BLOCK = re.compile(r"(<!-- no-run -->\s*\n)?```python[^\n]*\n(.*?)\n```", re.S)


def png(rgb: bytes, w: int = 96, h: int = 64) -> bytes:
    raw = b"".join(b"\x00" + rgb * w for _ in range(h))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def blocks(page: str) -> list[tuple[int, str]]:
    text = (DOCS / page).read_text(encoding="utf-8")
    out = []
    for m in BLOCK.finditer(text):
        if m.group(1):
            continue
        line = text.count("\n", 0, m.start(2)) + 1
        out.append((line, m.group(2)))
    return out


@pytest.mark.parametrize("page", RUN)
def test_page_code_runs(page, tmp_path, monkeypatch):
    found = blocks(page)
    if page not in ("faq.md",):
        assert found, f"{page} has no runnable code"
    monkeypatch.chdir(tmp_path)
    (tmp_path / "logo.png").write_bytes(png(b"\x2f\x54\x96"))
    (tmp_path / "chart.png").write_bytes(png(b"\x70\xad\x47", 160, 90))
    namespace: dict = {"__name__": f"docs_{Path(page).stem}"}
    for line, code in found:
        try:
            exec(compile(code, f"docs/{page}:{line}", "exec"), namespace)
        except Exception as e:  # pragma: no cover - the message is the point
            raise AssertionError(f"docs/{page}, block at line {line}: {type(e).__name__}: {e}") from e


def test_reference_is_current():
    from tools import apidocs

    pages, problems = apidocs.render()
    assert not problems, problems
    stale = [n for n, text in pages.items()
             if (DOCS / "reference" / n).read_text(encoding="utf-8") != text]
    assert not stale, f"docs/reference is stale ({stale}): run python -m tools.apidocs"


def test_every_guide_page_is_in_the_nav():
    nav = (ROOT / "mkdocs.yml").read_text(encoding="utf-8")
    for page in RUN:
        assert page in nav, page


@pytest.mark.skipif(shutil.which("mkdocs") is None and
                    subprocess.run([sys.executable, "-c", "import mkdocs"],
                                   capture_output=True, check=False).returncode,
                    reason="MkDocs not installed (pip install .[docs])")
def test_site_builds_strict(tmp_path):
    env = dict(os.environ, PYTHONPATH=os.pathsep.join([str(ROOT / "src"), str(ROOT)]))
    run = subprocess.run([sys.executable, "-m", "mkdocs", "build", "--strict", "-q",
                          "-d", str(tmp_path / "site")], cwd=ROOT, capture_output=True,
                         text=True, env=env, check=False)
    assert run.returncode == 0, run.stdout + run.stderr
    assert (tmp_path / "site" / "guide" / "word" / "index.html").is_file()


def test_readme_links_are_absolute():
    """PyPI shows the README without the repository around it: relative links break."""
    text = (ROOT / "README.md").read_text(encoding="utf-8")
    links = re.findall(r"\]\(([^)]+)\)|(?:src|href)=\"([^\"]+)\"", text)
    relative = [a or b for a, b in links if not re.match(r"https?://|mailto:|#", a or b)]
    assert relative == []
