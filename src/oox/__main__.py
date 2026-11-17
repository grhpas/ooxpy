"""Command line: ``python -m oox <command> FILE ...``

    python -m oox validate report.docx                  # schema problems per part
    python -m oox validate book.xlsx --target Office2007 --full
    python -m oox info deck.pptx                        # kind, parts, sheets/slides, properties
    python -m oox text report.docx                      # paragraphs / cells / slide texts
    python -m oox downgrade in.xlsx out.xlsx --target Office2007   # save(strip=True)

Exit status: 0 = fine, 1 = problems found, 2 = the file can't be read.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Sequence

from . import open_document
from .errors import OoxError


def _open(path: str):
    try:
        return open_document(path)
    except (OoxError, OSError, ValueError) as e:
        print(f"{path}: cannot open: {e}", file=sys.stderr)
        raise SystemExit(2) from None


def cmd_validate(args) -> int:
    doc = _open(args.file)
    report = doc.validate(target=args.target, level="full" if args.full else "basic")
    if args.json:
        print(json.dumps(report, indent=2, ensure_ascii=False))
        return 1 if report else 0
    total = sum(len(v) for v in report.values())
    for part, problems in report.items():
        print(part)
        for p in problems[: args.limit]:
            print(f"  - {p}")
        if len(problems) > args.limit:
            print(f"  ... {len(problems) - args.limit} more")
    what = f"valid{' for ' + args.target if args.target else ''}"
    print(f"{args.file}: {'no problems, ' + what if not total else f'{total} problem(s) in {len(report)} part(s)'}")
    return 1 if total else 0


def _summary(doc) -> dict:
    from .presentation import PresentationDocument
    from .spreadsheet import SpreadsheetDocument

    info: dict = {"kind": type(doc).__name__, "conformance": doc.conformance,
                  "parts": len(doc.package.parts), "main": doc.main_part.partname}
    props = doc.properties
    for key in ("title", "creator", "modified"):
        value = getattr(props, key, None)
        if value:
            info[key] = str(value)
    if isinstance(doc, SpreadsheetDocument):
        info["sheets"] = {name: {"rows": len(doc.values(name)), "charts": len(doc.charts(name)),
                                 "tables": len(doc.tables(name))} for name in doc.sheet_names}
    elif isinstance(doc, PresentationDocument):
        info["slides"] = [{"layout": doc.slide_layout(s),
                           "shapes": len(doc.shapes(s)),
                           "notes": doc.notes(s) is not None} for s in doc.slides]
        info["layouts"] = doc.layout_names
    else:
        from .walk import TBL, P, iter_story

        info["paragraphs"] = sum(1 for _ in iter_story(doc.body, P))
        info["tables"] = sum(1 for _ in iter_story(doc.body, TBL))
        info["charts"] = len(doc.charts())
    return info


def cmd_info(args) -> int:
    info = _summary(_open(args.file))
    if args.json:
        print(json.dumps(info, indent=2, ensure_ascii=False))
        return 0
    for key, value in info.items():
        if isinstance(value, dict):
            print(f"{key}:")
            for k, v in value.items():
                print(f"  {k}: {v}")
        elif isinstance(value, list) and value and isinstance(value[0], dict):
            print(f"{key}:")
            for i, v in enumerate(value, 1):
                print(f"  {i}: {v}")
        else:
            print(f"{key}: {value}")
    return 0


def cmd_text(args) -> int:
    from .document import WordprocessingDocument
    from .presentation import PresentationDocument
    from .spreadsheet import SpreadsheetDocument
    from .teletype import extractText

    doc = _open(args.file)
    out = sys.stdout
    if isinstance(doc, SpreadsheetDocument):
        for name in doc.sheet_names:
            print(f"# {name}", file=out)
            for row in doc.values(name, dates=True):
                print("\t".join("" if v is None else str(v) for v in row), file=out)
    elif isinstance(doc, PresentationDocument):
        for i, slide in enumerate(doc.slides, 1):
            print(f"# Slide {i}", file=out)
            for shape in doc.shapes(slide):
                if shape.text:
                    print(shape.text, file=out)
            notes = doc.notes(slide)
            if notes:
                print(f"[notes] {notes}", file=out)
    elif isinstance(doc, WordprocessingDocument):
        from .walk import P, iter_story

        # each paragraph once, as Word shows it: a text box's paragraphs
        # follow the one that anchors it, and the VML copy of the box is skipped
        for para in iter_story(doc.body, P):
            print(extractText(para, text_boxes=False), file=out)
    return 0


def cmd_downgrade(args) -> int:
    doc = _open(args.file)
    doc.save(args.output, target=args.target, strip=True)
    changes = sum(len(v) for v in doc.stripped.values())
    for part, msgs in doc.stripped.items():
        print(part)
        for m in msgs:
            print(f"  - {m}")
    print(f"{args.output}: written for {args.target}, {changes} change(s)")
    return 0


def main(argv: Sequence[str] | None = None, prog: str = "python -m oox") -> int:
    from .grammar import OFFICE_VERSIONS

    ap = argparse.ArgumentParser(prog=prog, description=__doc__.split("\n\n")[0]
                                 if __doc__ else None)
    sub = ap.add_subparsers(dest="command", required=True)
    v = sub.add_parser("validate", help="schema problems per part")
    v.add_argument("file")
    v.add_argument("--target", choices=OFFICE_VERSIONS, help="also report markup this Office can't read")
    v.add_argument("--full", action="store_true",
                   help="also match whole content models (choices, occurrence counts)")
    v.add_argument("--json", action="store_true")
    v.add_argument("--limit", type=int, default=20, help="problems shown per part (default 20)")
    v.set_defaults(fn=cmd_validate)
    i = sub.add_parser("info", help="kind, parts, sheets/slides/paragraphs, properties")
    i.add_argument("file")
    i.add_argument("--json", action="store_true")
    i.set_defaults(fn=cmd_info)
    t = sub.add_parser("text", help="paragraphs, cell values (tab-separated) or slide texts")
    t.add_argument("file")
    t.set_defaults(fn=cmd_text)
    d = sub.add_parser("downgrade", help="save for an older Office, removing what it can't read")
    d.add_argument("file")
    d.add_argument("output")
    d.add_argument("--target", choices=OFFICE_VERSIONS, default="Office2007")
    d.set_defaults(fn=cmd_downgrade)
    args = ap.parse_args(argv)
    try:
        return args.fn(args)
    except BrokenPipeError:              # e.g. | head: stop quietly
        import os

        os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
        return 0


def console_main() -> None:
    """The ``oox`` command that installing the package puts on the PATH."""
    raise SystemExit(main(prog="oox"))


if __name__ == "__main__":
    raise SystemExit(main())
