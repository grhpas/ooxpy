"""Rough timings for building and reading large documents.

    python -m tools.bench            # 50,000 rows x 6 cells, 20,000 paragraphs
"""

from __future__ import annotations

import io
import sys
import time

from oox import SpreadsheetDocument, WordprocessingDocument, open_document


def _t(label: str, fn):
    start = time.perf_counter()
    out = fn()
    print(f"{label:<40} {time.perf_counter() - start:7.2f} s")
    return out


def main(rows: int = 50000, paragraphs: int = 20000) -> None:
    book = SpreadsheetDocument.new("Data")

    def fill():
        for i in range(rows):
            book.append_row("Data", [f"row {i}", i, i * 1.5, i % 7 == 0, None, "x"])

    _t(f"Excel: append {rows:,} rows x 6", fill)
    buf = io.BytesIO()
    _t("Excel: save", lambda: book.save(buf))
    print(f"{'':<40} {len(buf.getvalue()) / 1e6:7.2f} MB")
    buf.seek(0)
    again = _t("Excel: open", lambda: open_document(buf))
    _t("Excel: values()", lambda: again.values("Data"))
    refs = [f"{'ABCDEF'[i % 6]}{1 + (i * 7919) % rows}" for i in range(10000)]
    _t("Excel: first value() (builds caches)", lambda: again.value("Data", "A1"))
    _t("Excel: 10,000 random value()", lambda: [again.value("Data", r) for r in refs])
    doc = WordprocessingDocument.new()
    _t(f"Word: {paragraphs:,} paragraphs",
       lambda: [doc.add_paragraph(f"Paragraph {i}") for i in range(paragraphs)])
    out = io.BytesIO()
    _t("Word: save", lambda: doc.save(out))


if __name__ == "__main__":
    main(*(int(a) for a in sys.argv[1:3]))
