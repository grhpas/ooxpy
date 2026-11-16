"""Tables with merged cells, a floating logo, and a reviewed (tracked) edit.

    python examples/tables_review.py outdir
"""

import datetime as dt
import sys
from pathlib import Path

from oox import WordprocessingDocument, open_document
from oox.units import Cm


def logo_png() -> bytes:
    """A small PNG drawn without any imaging library (a solid blue square)."""
    import struct
    import zlib

    size = 64
    raw = b"".join(b"\x00" + b"\x2f\x54\x96" * size for _ in range(size))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def build(path: Path) -> None:
    doc = WordprocessingDocument.new()
    title = doc.add_heading("Sales by region", 0)
    doc.add_floating_picture(logo_png(), width=Cm(2), h_align="right", h_relative="margin",
                             paragraph=title, wrap="square", descr="Company logo")
    doc.add_paragraph("Figures for the first half of the year. The total row spans the "
                      "label columns.")
    t = doc.add_table([["Region", "Product", "Q1", "Q2"],
                       ["North", "Widgets", 120, 135],
                       [None, "Gadgets", 80, 95],
                       ["South", "Widgets", 95, 110],
                       ["Total", None, 295, 340]],
                      header_rows=1, widths=[Cm(4), Cm(4), Cm(3), Cm(3)])
    t.merge(1, 0, 2, 0)        # "North" covers both of its rows
    t.merge(4, 0, 4, 1)        # "Total" spans two columns
    for c in range(4):
        t.format_cell(0, c, fill="D9E2F3", align="center")
    t.format_cell(1, 0, valign="center")
    t.format_cell(4, 0, align="right", bold=True)
    doc.add_paragraph("Sales grew by 8% in 2025, driven by the north.")
    doc.save(path)


def review(src: Path, dst: Path) -> None:
    doc = open_document(src)
    p = doc.body.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")[-1]
    when = dt.datetime(2026, 11, 14, 9, 0)
    doc.suggest_replace(p, "8%", "9%", author="Finance", date=when)
    doc.suggest_insert(p, " Figures are unaudited.", author="Finance", date=when)
    doc.add_comment(p, "Updated with the final numbers.", author="Finance")
    doc.track_changes = True        # later edits in Word are tracked too
    doc.save(dst)
    print(f"{len(doc.revisions())} tracked changes in {dst}")


def main(outdir: str) -> None:
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    build(out / "sales.docx")
    review(out / "sales.docx", out / "sales-reviewed.docx")
    final = open_document(out / "sales-reviewed.docx")
    final.accept_all()
    assert not final.validate(target="Office2007")
    final.save(out / "sales-final.docx")
    print(f"wrote {out / 'sales.docx'}, sales-reviewed.docx, sales-final.docx")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "out")
