"""Interop check over third-party files: open, save unchanged (byte-identical
parts), validate (basic and full), run every reader, edit and save again.

    python -m tools.interop DIR [DIR ...]          # summary + problems
    python -m tools.interop --build OUT            # make a corpus with LibreOffice from
                                                   # its templates (and odfpy's test files
                                                   # if vendor/odfpy exists)
    python -m tools.writers OUT                    # or with python-docx, openpyxl,
                                                   # XlsxWriter and python-pptx

Nothing here is committed as test data: the corpus is rebuilt from what a
LibreOffice installation ships, so its licence stays with LibreOffice.
"""

from __future__ import annotations

import argparse
import glob
import hashlib
import io
import os
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile
from collections import Counter
from pathlib import Path

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument, open_document


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist() if not n.endswith("/")}


def _readers(doc) -> list[str]:
    """Run the readers; return what each found (for the summary)."""
    found = []
    if isinstance(doc, WordprocessingDocument):
        from oox.ns import w
        from oox.teletype import extractText

        text = extractText(doc.body)
        found.append(f"{len(text)} chars")
        found.append(f"{len(doc.charts())} charts")
        found.append(f"{len(doc.tables())} tables")
        found.append(f"{len(doc.sections)} sections")
        doc.fields()
        doc.revisions()
        doc.pictures()
        doc.content_controls()
        doc.body.getElementsByType(w.P)
    elif isinstance(doc, SpreadsheetDocument):
        for name in doc.sheet_names:
            doc.values(name, dates=True)
            doc.charts(name)
            doc.tables(name)
            doc.conditional_formats(name)
            doc.comments(name)
            doc.hyperlinks(name)
            doc.validations(name)
            doc.merged_cells(name)
            doc.page_setup_info(name)
        doc.defined_names()
        found.append(f"{len(doc.sheet_names)} sheets")
    elif isinstance(doc, PresentationDocument):
        for slide in doc.slides:
            doc.shapes(slide)
            doc.tables(slide)
            doc.charts(slide)
            doc.notes(slide)
            doc.slide_layout(slide)
            doc.placeholders(slide)
        found.append(f"{len(doc.slides)} slides, {len(doc.layout_names)} layouts")
    return found


def _edit(doc) -> None:
    if isinstance(doc, WordprocessingDocument):
        doc.add_paragraph("Added by oox")
    elif isinstance(doc, SpreadsheetDocument):
        doc.append_row(doc.sheet_names[0], ["Added by oox", 1])
    elif isinstance(doc, PresentationDocument):
        layout = "Title Only" if "Title Only" in doc.layout_names else None
        doc.add_slide(layout)


def check(path: str) -> dict:
    """One file -> {"ok": bool, "problems": [...], "notes": [...]}"""
    out: dict = {"file": path, "problems": [], "notes": [], "ok": False}
    original = Path(path).read_bytes()
    try:
        doc = open_document(io.BytesIO(original))
    except Exception as e:
        out["problems"].append(f"open: {type(e).__name__}: {e}")
        return out
    # 1. unchanged save keeps every part
    buf = io.BytesIO()
    try:
        doc.save(buf)
        before, after = _parts(original), _parts(buf.getvalue())
        changed = sorted(n for n in before if n in after and before[n] != after[n])
        missing = sorted(set(before) - set(after))
        if changed:
            out["problems"].append(f"unchanged save rewrote {len(changed)} part(s): {changed[:4]}")
        if missing:
            out["problems"].append(f"unchanged save dropped {missing[:4]}")
    except Exception as e:
        out["problems"].append(f"save: {type(e).__name__}: {e}")
    # 2. validation
    try:
        basic = doc.validate()
        full = doc.validate(level="full")
        nb, nf = sum(map(len, basic.values())), sum(map(len, full.values()))
        out["basic"], out["full_only"] = basic, {
            k: [m for m in v if m not in basic.get(k, [])] for k, v in full.items()}
        out["notes"].append(f"validate: {nb} basic, {nf - nb} more at full")
    except Exception as e:
        out["problems"].append(f"validate: {type(e).__name__}: {e}")
    # 3. readers
    try:
        out["notes"] += _readers(doc)
    except Exception as e:
        out["problems"].append(f"reader: {type(e).__name__}: {e} "
                               f"({traceback.extract_tb(e.__traceback__)[-1].name})")
    # 4. edit, save, reopen, no new problems
    try:
        doc2 = open_document(io.BytesIO(original))
        _edit(doc2)
        b2 = io.BytesIO()
        doc2.save(b2)
        b2.seek(0)
        again = open_document(b2)
        new = {k: [m for m in v if m not in out.get("basic", {}).get(k, [])]
               for k, v in again.validate().items()}
        new = {k: v for k, v in new.items() if v}
        if new:
            out["problems"].append(f"edit added problems: {dict(list(new.items())[:2])}")
    except Exception as e:
        out["problems"].append(f"edit: {type(e).__name__}: {e} "
                               f"({traceback.extract_tb(e.__traceback__)[-1].name})")
    out["ok"] = not out["problems"]
    return out


def build(out_dir: str) -> int:
    """Convert LibreOffice's templates (and odfpy's test files) to OOXML."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice not found", file=sys.stderr)
        return 2
    roots = [r for r in ("/usr/lib/libreoffice/share/template",
                         "/Applications/LibreOffice.app/Contents/Resources/template",
                         "vendor/odfpy") if os.path.isdir(r)]
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        groups: dict[str, list[str]] = {"docx": [], "xlsx": [], "pptx": []}
        for root in roots:
            for f in glob.glob(f"{root}/**/*.o[dt][tsp]", recursive=True):
                ext = {"t": "docx", "s": "xlsx", "p": "pptx"}[f[-1]]
                tag = hashlib.sha1(f.encode()).hexdigest()[:6]   # same name on every run
                name = f"{tag}_{os.path.basename(f).replace(' ', '_')}"
                shutil.copy(f, os.path.join(tmp, name))
                groups[ext].append(os.path.join(tmp, name))
        for ext, files in groups.items():
            if files:
                subprocess.run([soffice, "--headless", "--convert-to", ext, "--outdir", str(out),
                                *files], capture_output=True, timeout=1800, check=False)
    print(f"{len(list(out.iterdir()))} files in {out}")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("dirs", nargs="*")
    ap.add_argument("--build", metavar="OUT")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    if args.build:
        return build(args.build)
    files = sorted(f for d in args.dirs for f in glob.glob(os.path.join(d, "*.*x")))
    results = [check(f) for f in files]
    bad = [r for r in results if not r["ok"]]
    kinds: Counter = Counter()
    full_only: Counter = Counter()
    basic_msgs: Counter = Counter()
    for r in results:
        for p in r["problems"]:
            kinds[p.split(":")[0]] += 1
        for msgs in r.get("full_only", {}).values():
            for m in msgs:
                full_only[m.split(" (child")[0][:120]] += 1
        for msgs in r.get("basic", {}).values():
            for m in msgs:
                basic_msgs[m[:120]] += 1
    print(f"{len(results)} files, {len(results) - len(bad)} without problems")
    for r in bad:
        print(f"- {os.path.basename(r['file'])}")
        for p in r["problems"]:
            print(f"    {p[:300]}")
    if basic_msgs:
        print("\nmost common validation findings (basic):")
        for m, n in basic_msgs.most_common(15):
            print(f"  {n:4d}  {m}")
    if full_only:
        print("\nmost common findings only at level='full':")
        for m, n in full_only.most_common(15):
            print(f"  {n:4d}  {m}")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
