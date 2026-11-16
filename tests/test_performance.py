"""Guards against quadratic building.  The thresholds are ~10x today's times
(generous for slow CI machines) but far below the quadratic code they replaced:
10,000 rows took 32 s and 10,000 Word paragraphs 16 s before 2026-11-15."""

import time

from oox import PresentationDocument, SpreadsheetDocument, WordprocessingDocument
from oox.spreadsheet import cell_ref


def _timed(fn) -> float:
    start = time.perf_counter()
    fn()
    return time.perf_counter() - start


def test_append_rows_is_linear():
    book = SpreadsheetDocument.new("S")

    def fill():
        for i in range(10000):
            book.append_row("S", [f"row {i}", i, i * 1.5, None, True])

    assert _timed(fill) < 15
    assert book.value("S", "B10000") == 9999


def test_set_value_in_order_is_linear():
    book = SpreadsheetDocument.new("S")

    def fill():
        for r in range(1, 3001):
            for c in range(1, 6):
                book.set_value("S", cell_ref(c, r), r * c)

    assert _timed(fill) < 15
    assert book.value("S", "E3000") == 15000


def test_word_paragraphs_before_sectpr_are_linear():
    doc = WordprocessingDocument.new()

    def fill():
        for i in range(20000):
            doc.add_paragraph(f"Paragraph {i}")

    assert _timed(fill) < 10
    assert doc.body[-1].tag.endswith("}sectPr")    # section properties stay last


def test_many_slides():
    deck = PresentationDocument.new()
    assert _timed(lambda: [deck.add_slide("Title Only") for _ in range(300)]) < 15
    assert len(deck.slides) == 300


# -- single-cell reads: cached, and the caches follow edits ----------------------------
def _book(rows=2000):
    book = SpreadsheetDocument.new("S")
    for i in range(1, rows + 1):
        book.append_row("S", [f"r{i}", i])
    return book


def test_single_cell_reads_are_fast_after_the_first():
    import io

    from oox import open_document

    book = _book(10000)
    book.share_strings()
    buf = io.BytesIO()
    book.save(buf)
    buf.seek(0)
    again = open_document(buf)
    again.value("S", "A1")                                  # builds the caches
    refs = [cell_ref(1 + i % 2, 1 + (i * 7919) % 10000) for i in range(2000)]
    assert _timed(lambda: [again.value("S", r) for r in refs]) < 2   # ~0.05 ms each
    assert again.value("S", "A10000") == "r10000" and again.value("S", "B7") == 7


def test_row_index_follows_oox_edits():
    book = _book(10)
    book.cells("S")
    assert book.value("S", "B5") == 5
    book.set_value("S", "B20", 20)                          # appended row
    book.set_value("S", "C5", "new")                         # new cell in a known row
    assert (book.value("S", "B20"), book.value("S", "C5")) == (20, "new")
    data = book.sheet_data("S")
    rows = list(data)
    for r in rows[2:]:                                     # renumber rows 3.. down by 2
        n = int(r.get("r")) + 2
        r.set("r", str(n))
        for c in r:
            c.set("r", c.get("r").rstrip("0123456789") + str(n))
    book.set_value("S", "A3", "inserted")                   # into the gap
    assert book.value("S", "A3") == "inserted"
    assert book.value("S", "B7") == 5 and book.value("S", "B5") == 3
    assert book.value("S", "A4") is None


def test_row_index_follows_direct_tree_edits():
    from oox.ns import x

    book = _book(10)
    assert book.value("S", "A6") == "r6"
    data = book.sheet_data("S")
    six = next(r for r in data if r.get("r") == "6")
    data.remove(six)                                        # removed
    assert book.value("S", "A6") is None
    seven = next(r for r in data if r.get("r") == "7")
    seven.addprevious(x.Row(x.C(x.V(text="66"), r="B6"), r="6"))   # inserted in the middle
    assert book.value("S", "B6") == 66
    data.remove(next(iter(data)))                           # row 1 gone, row 2 is first
    assert book.value("S", "A1") is None and book.value("S", "A2") == "r2"


def test_rows_without_numbers_fall_back_to_scanning():
    from oox.ns import x

    book = SpreadsheetDocument.new("S")
    data = book.sheet_data("S")
    data.append(x.Row(x.C(x.V(text="1"))))
    data.append(x.Row(x.C(x.V(text="2")), x.C(x.V(text="3"))))
    assert book.value("S", "A2") == 2 and book.value("S", "B2") == 3


def test_shared_strings_and_date_styles_follow_edits():
    import datetime as dt

    book = _book(3)
    assert book.value("S", "A1") == "r1"
    book.share_strings()
    book.set_value("S", "C1", "later")
    book.share_strings()
    assert (book.value("S", "A1"), book.value("S", "C1")) == ("r1", "later")
    book.set_value("S", "D1", 45000)
    assert book.value("S", "D1") == 45000
    book.set_value("S", "D1", 45000, style={"number_format": "yyyy-mm-dd"})
    assert book.value("S", "D1") == dt.datetime(2023, 3, 15)
