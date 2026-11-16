"""Check an installed ooxpy from outside the repository.

    python tools/smoke_install.py            # with the interpreter that has ooxpy installed

Run it from a directory that isn't the repository (so ``import oox`` finds the
installed package, not ``src/``).  It builds one file per format with the
main features, reopens them, validates at every level and for Office 2007,
checks that the type stubs and ``py.typed`` were installed, and runs the
``oox`` command.  The release workflow runs it against the built wheel on the
oldest and newest supported Python.  Standard library plus oox only.
"""

from __future__ import annotations

import importlib.resources
import os
import shutil
import subprocess
import sys
import tempfile


def main() -> int:
    import oox
    from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument, open_document
    from oox.units import Cm

    here = os.path.dirname(os.path.abspath(oox.__file__))
    if os.path.isdir(os.path.join(here, "..", "..", "tools")) and "site-packages" not in here:
        print(f"oox was imported from a source tree ({here}); run from outside the repository")
        return 2
    print(f"oox {oox.__version__} from {here}")
    out = tempfile.mkdtemp(prefix="oox-smoke-")
    try:
        doc = WordprocessingDocument.new()
        doc.add_heading("Smoke test", 1)
        doc.add_paragraph("Installed from a wheel.")
        doc.add_table([["Region", "Units"], ["North", 120]], header_rows=1)
        doc.add_text_box("A text box", width=Cm(5), height=Cm(2))
        doc.add_chart("column", ["Q1", "Q2"], {"North": [1, 2]})
        doc.protect("readOnly", password="secret")
        doc.save(os.path.join(out, "smoke.docx"))

        book = SpreadsheetDocument.new("Sales")
        book.append_row("Sales", ["Region", "Units"])
        book.append_row("Sales", ["North", 120])
        book.append_row("Sales", ["South", 80])
        book.add_chart("Sales", "column", "A1:B3")
        book.protect_workbook(password="secret")
        book.save(os.path.join(out, "smoke.xlsx"))

        deck = PresentationDocument.new()
        slide = deck.add_slide("Title and Content")
        deck.set_placeholder_text(slide, "title", "Smoke test")
        deck.add_table(slide, [["a", "b"], [1, 2]], Cm(2), Cm(5))
        deck.save(os.path.join(out, "smoke.pptx"))

        failures = []
        for name in ("smoke.docx", "smoke.xlsx", "smoke.pptx"):
            back = open_document(os.path.join(out, name))
            for label, found in (("basic", back.validate()),
                                 ("full", back.validate(level="full")),
                                 ("Office2007", back.validate(target="Office2007"))):
                if found:
                    failures.append(f"{name} ({label}): {found}")
            print(f"{name}: {type(back).__name__}, valid")

        files = importlib.resources.files("oox")
        for res in ("py.typed",):
            if not (files / res).is_file():
                failures.append(f"missing {res}")
        if not (importlib.resources.files("oox.ns") / "w.pyi").is_file():
            failures.append("missing type stubs (oox/ns/*.pyi)")

        cli = shutil.which("oox")
        command = [cli] if cli else [sys.executable, "-m", "oox"]
        run = subprocess.run([*command, "validate", os.path.join(out, "smoke.docx"),
                              "--target", "Office2007"], capture_output=True, text=True,
                             check=False)
        print(f"{' '.join(os.path.basename(c) for c in command)} validate: exit {run.returncode}")
        if run.returncode != 0:
            failures.append(f"oox validate: {run.stdout}{run.stderr}")
        if cli is None:
            failures.append("the `oox` command isn't on the PATH")

        for f in failures:
            print("FAIL", f)
        print("ok" if not failures else f"{len(failures)} failure(s)")
        return 1 if failures else 0
    finally:
        shutil.rmtree(out, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
