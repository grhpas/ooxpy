"""Excel tables (Insert > Table): a styled range with a header row, filter
buttons, structured references and an optional totals row.

    book.add_table("Sales", "A1:D20", name="Sales", style="TableStyleMedium2",
                   totals={"Region": "Total", "Amount": "sum", "Share": "average"})
    book.set_value("Sales", "F1", Formula("SUM(Sales[Amount])"))   # structured reference
    book.tables("Sales")   # [{"name": "Sales", "ref": "A1:D21", "columns": [...], ...}]

The first row of the range is the header; its texts become the column names
(empty or repeated headers are replaced by "Column1", "Column2" ... and written
into the cells, as Excel requires).  ``totals`` adds a row below the range: a
string is a label, a function name (sum, average, count, count_numbers, min,
max, stddev, var) a ``SUBTOTAL`` formula, whose result oox computes and caches.
"""

from __future__ import annotations

import math
import re
from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import x

if TYPE_CHECKING:
    from .spreadsheet import SpreadsheetDocument

X = URI["x"]
R = URI["r"]
# totals function -> (totalsRowFunction, SUBTOTAL code)
FUNCTIONS = {"sum": ("sum", 109), "average": ("average", 101), "count": ("count", 103),
             "count_numbers": ("countNums", 102), "max": ("max", 104), "min": ("min", 105),
             "stddev": ("stdDev", 107), "var": ("var", 110)}
STYLES_HINT = "TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11"


def _table_parts(book: SpreadsheetDocument):
    for part in book.package.parts.values():
        if part.content_type.endswith("spreadsheetml.table+xml"):
            yield part


def _sheet_tables(book: SpreadsheetDocument, sheet: str):
    ws = book.sheet(sheet)
    part = book.part_of(ws)
    return [(p, p.element) for p in book.related_parts(part, "TableDefinitionPart")]


def _validate_name(book: SpreadsheetDocument, name: str) -> None:
    if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.\\]*", name) or \
            re.fullmatch(r"[A-Za-z]{1,3}\d+", name) or name.upper() in ("R", "C"):
        raise ValueError(f"invalid table name {name!r} (letters, digits, _ and .; "
                         "not a cell reference; no spaces)")
    taken = {(p.element.get("name") or "").lower() for p in _table_parts(book)}
    taken |= {k.rpartition("!")[2].lower() for k in book.defined_names()}
    if name.lower() in taken:
        raise ValueError(f"the name {name!r} is already used by a table or defined name")


def _subtotal(values: list[Any], func: str) -> float | int | None:
    nums = [v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if func == "count":  # COUNTA: every non-empty cell, an empty string included
        return len([v for v in values if v is not None])
    if func == "count_numbers":
        return len(nums)
    if not nums:
        return 0 if func == "sum" else None
    if func == "sum":
        total = math.fsum(nums)
        return int(total) if total.is_integer() else total
    if func == "average":
        return math.fsum(nums) / len(nums)
    if func == "max":
        return max(nums)
    if func == "min":
        return min(nums)
    if len(nums) < 2:
        return None
    mean = math.fsum(nums) / len(nums)
    var = math.fsum((v - mean) ** 2 for v in nums) / (len(nums) - 1)
    return math.sqrt(var) if func == "stddev" else var


def add_table(book: SpreadsheetDocument, sheet: str, ref: str, *, name: str | None = None,
              style: str | None = "TableStyleMedium2", totals: dict[str, str] | None = None,
              stripes: bool = True, first_column: bool = False,
              last_column: bool = False) -> OoxElement:
    from .spreadsheet import Formula, cell_ref
    from .xlsheet import auto_filter_range, merged_cells, parse_range, range_ref

    c1, r1, c2, r2 = parse_range(ref)
    if r2 <= r1:
        raise ValueError("a table needs a header row and at least one data row")
    existing = [t.ref for t in tables(book, sheet)]
    af = auto_filter_range(book.sheet(sheet))
    for other in existing + merged_cells(book.sheet(sheet)) + ([af] if af else []):
        d1, s1, d2, s2 = parse_range(other)
        last = r2 + (1 if totals else 0)
        if not (c2 < d1 or d2 < c1 or last < s1 or s2 < r1):
            raise ValueError(f"{ref} overlaps {other} (another table, a merge or the filter)")
    ids = [int(p.element.get("id", "0")) for p in _table_parts(book)]
    tid = max(ids, default=0) + 1
    name = name or f"Table{tid}"
    _validate_name(book, name)
    # column names from the header row
    names: list[str] = []
    for col in range(c1, c2 + 1):
        text = book.display_text(sheet, cell_ref(col, r1)).strip()
        if not text or text.lower() in (n.lower() for n in names):
            text = f"Column{col - c1 + 1}"
            while text.lower() in (n.lower() for n in names):
                text += "_"
        if book.value(sheet, cell_ref(col, r1)) != text:
            book.set_value(sheet, cell_ref(col, r1), text)
        names.append(text)
    columns = x.TableColumns(count=len(names))
    totals = dict(totals or {})
    unknown = set(totals) - set(names)
    if unknown:
        raise ValueError(f"totals for unknown columns {sorted(unknown)}; columns: {names}")
    for i, col_name in enumerate(names, start=1):
        tc = x.TableColumn(id=i, name=col_name)
        spec = totals.get(col_name)
        if spec is not None:
            if spec in FUNCTIONS:
                tc.set("totalsRowFunction", FUNCTIONS[spec][0])
            else:
                tc.set("totalsRowLabel", spec)
        columns.add(tc)
    last_row = r2 + (1 if totals else 0)
    attrs: dict[str, Any] = {"id": tid, "name": name, "displayName": name,
                             "ref": range_ref(c1, r1, c2, last_row)}
    if totals:
        attrs["totalsRowCount"] = 1
    else:
        attrs["totalsRowShown"] = False
    table = x.Table(x.AutoFilter(ref=range_ref(c1, r1, c2, r2)), columns, **attrs)
    if style:
        table.add(x.TableStyleInfo(name=style, showFirstColumn=first_column,
                                   showLastColumn=last_column, showRowStripes=stripes,
                                   showColumnStripes=False))
    ws = book.sheet(sheet)
    _, rid = book.add_part("TableDefinitionPart", book.part_of(ws), element=table)
    tparts = ws.first("x:tableParts")
    if tparts is None:
        tparts = x.TableParts()
        ws.add(tparts)
    tparts.add(x.TablePart(r_id=rid))
    tparts.set("count", str(len(tparts)))
    # the totals row: labels and SUBTOTAL formulas with computed results
    for i, col_name in enumerate(names):
        spec = totals.get(col_name)
        if spec is None:
            continue
        ref_ = cell_ref(c1 + i, last_row)
        if spec in FUNCTIONS:
            data = [book.value(sheet, cell_ref(c1 + i, r), dates=False) for r in range(r1 + 1, r2 + 1)]
            formula = f"SUBTOTAL({FUNCTIONS[spec][1]},{name}[{_escape(col_name)}])"
            book.set_value(sheet, ref_, Formula(formula, _subtotal(data, spec)))
        else:
            book.set_value(sheet, ref_, spec)
    return table


def _escape(column: str) -> str:
    """Structured-reference escaping: ' [ ] # get a leading apostrophe."""
    return re.sub(r"(['\[\]#])", r"'\1", column)


def tables(book: SpreadsheetDocument, sheet: str) -> list:
    """An ``ExcelTable`` for each table on the sheet."""
    from .records import ExcelTable

    out = []
    for _, t in _sheet_tables(book, sheet):
        cols = t.first("x:tableColumns")
        info = t.first("x:tableStyleInfo")
        def shown(attr: str, default: str = "0", info=info) -> bool:
            return (info.get(attr, default) if info is not None else default) in ("1", "true")

        out.append(ExcelTable(
            name=t.get("name"), ref=t.get("ref"),
            columns=[c.get("name") for c in cols.iterchildren(f"{{{X}}}tableColumn")]
            if cols is not None else [],
            style=info.get("name") if info is not None else None,
            stripes=shown("showRowStripes", "1"), first_column=shown("showFirstColumn"),
            last_column=shown("showLastColumn"),
            header=t.get("headerRowCount", "1") != "0",
            totals_row=int(t.get("totalsRowCount", "0")) > 0))
    return out
