# Excel

`SpreadsheetDocument`, cell values and the stylesheet.

## `SpreadsheetDocument`

*Based on* `OoxDocument`. The common methods are under [Documents](documents.md).

Base class.  Subclasses set `MAIN_CONTENT_TYPES` and implement `new`.

### Workbook and sheets

#### `new`

```python
new(sheet: str | None = 'Sheet1', kind: str = '.xlsx') -> SpreadsheetDocument
```

*classmethod*

A new workbook with one worksheet named `sheet` (Excel refuses a workbook
without one).  `kind` ".xlsm", ".xltx" ... sets the content type.

#### `workbook`

*read-only property*

The root `x:workbook` element of the main part.

#### `spreadsheet`

*read-only property*

The root `x:workbook` element of the main part.

#### `add_sheet`

```python
add_sheet(name: str) -> OoxElement
```

Append a worksheet; returns its `<x:worksheet>` root.

#### `sheet`

```python
sheet(name: str) -> OoxElement
```

Root `<x:worksheet>` of the named sheet.

#### `shared_strings_part`

*read-only property*

The shared strings part (`xl/sharedStrings.xml`), or None.

#### `styles_part`

*read-only property*

The workbook styles part (`xl/styles.xml`), or None.

#### `theme_part`

*read-only property*

The theme part (`xl/theme/theme1.xml`), or None.

#### `sheet_names`

*read-only property*

The worksheets, in tab order (chart sheets: `chart_sheet_names`).

#### `tab_names`

*read-only property*

Every tab in order: worksheets and chart sheets (`localSheetId`
and the active tab count these).

#### `chart_sheet_names`

*read-only property*

Names of the chart sheets, in tab order (`tab_names` has every tab).

#### `worksheets`

*read-only property*

The `x:worksheet` elements, in tab order (chart sheets excluded).

#### `sheet_data`

```python
sheet_data(sheet: str | OoxElement) -> OoxElement
```

The sheet's `x:sheetData` element (the rows).

#### `cell_element`

```python
cell_element(sheet: str | OoxElement, ref: str) -> OoxElement | None
```

The `x:c` element at `ref`, or None if the cell isn't stored.

### Values

#### `set_value`

```python
set_value(
    sheet: str | OoxElement,
    ref: str,
    value: object,
    *,
    style: int | dict | None = None,
) -> OoxElement
```

Write one cell, creating its row if needed and keeping row/column
order.  Dates and times get a date format.  `style`: an index from
`cell_style` or a dict of its options; without it an existing
cell's style is kept.

#### `append_row`

```python
append_row(
    sheet: str | OoxElement,
    values: Iterable[object],
    start_col: int = 1,
    *,
    style: int | dict | None = None,
) -> OoxElement
```

Add a row after the last one; returns the `<x:row>`.  Dates get a
date format; `style` (an index from `cell_style` or a dict of its
options) applies to every cell.

#### `value`

```python
value(
    sheet: str | OoxElement,
    ref: str,
    *,
    formulas: bool = False,
    dates: bool = True,
) -> object
```

Python value of one cell (`None` if empty or absent).

Formula cells give their cached result; files written by oox (or
openpyxl) have no cached result until Excel saves them, so you get
`None` -- pass `formulas=True` to get `Formula` objects instead.
Numbers in a date or time format come back as `datetime` (pass
`dates=False` for the serial number).

#### `values`

```python
values(
    sheet: str | OoxElement,
    *,
    formulas: bool = False,
    dates: bool = True,
) -> list[list[object]]
```

All values as a dense list of rows, starting at A1.

#### `cells`

```python
cells(
    sheet: str | OoxElement,
    *,
    formulas: bool = False,
    dates: bool = True,
) -> dict[str, object]
```

Non-empty cells as `{"A1": value, ...}` in sheet order.

#### `display_text`

```python
display_text(sheet: str | OoxElement, ref: str) -> str
```

The text Excel shows in the cell (number format applied; formula
cells show their cached result, if any).

#### `share_strings`

```python
share_strings() -> int
```

Move inline strings into the shared-strings table, as Excel stores
them (smaller files; some readers expect it).  Equal strings share one
entry.  Returns the number of cells converted.

### Cell styles

#### `cell_style`

```python
cell_style(**fmt) -> int
```

A cell style index for `set_value(..., style=)`: number_format,
bold, italic, underline, strike, size, font, color, fill (hex), border
("thin", "medium" ... or {"bottom": "double"}), border_color, align,
valign, wrap, indent, rotation.  Equal styles share one index.

#### `style_range`

```python
style_range(
    sheet: str | OoxElement,
    ref: str,
    style: int | dict | None = None,
    **fmt,
) -> None
```

Apply a style to every cell in a range (empty cells are created).

#### `cell_format`

```python
cell_format(sheet: str | OoxElement, ref: str) -> CellFormat
```

The cell's formatting as a `CellFormat` with `cell_style`'s options
(number_format, bold, font, fill, border, align ...).

### Layout

#### `set_column_width`

```python
set_column_width(sheet: str, cols: str, chars: float | None) -> None
```

Width in characters, as Excel's Column Width dialog shows it;
`cols` is "B" or "B:D"; None resets.

#### `column_width`

```python
column_width(sheet: str, col: str) -> float | None
```

Width of column `col` ("B") in characters as Excel's dialog shows it, or None
for the default.

#### `autofit_columns`

```python
autofit_columns(sheet: str, cols: str | None = None, **options) -> dict[str, float]
```

Widen columns to their longest displayed text (an estimate), so
numbers and dates don't show as ####; see `oox.xlsheet.autofit_columns`.

**Options** (keywords)

```python
min_width: float = 8.43
max_width: float = 80.0
padding: float = 1.5
```

Set column widths from the longest displayed text in each column (an
estimate: characters of the default font, bold or larger text not
weighted).  Excel shows `####` for numbers and dates that don't fit,
so this is worth calling after filling a sheet.  Returns the widths set.

#### `set_row_height`

```python
set_row_height(sheet: str, row_index: int, points: float | None) -> None
```

Set row `row_index` (1-based) to `points`; None returns it to the default.

#### `row_height`

```python
row_height(sheet: str, row_index: int) -> float | None
```

Height of row `row_index` (1-based) in points, or None for the default.

#### `merge_cells`

```python
merge_cells(sheet: str, ref: str) -> None
```

Merge the range `ref` ("A1:C1"); the top-left cell's value is shown.

#### `unmerge_cells`

```python
unmerge_cells(sheet: str, ref: str) -> None
```

Undo the merge of `ref`.

#### `merged_cells`

```python
merged_cells(sheet: str) -> list[str]
```

The sheet's merged ranges, in document order.

#### `set_frozen_panes`

```python
set_frozen_panes(sheet: str, ref: str | None) -> None
```

Freeze the rows above and columns left of `ref` ("B2"); None unfreezes.

#### `frozen_panes`

```python
frozen_panes(sheet: str) -> str | None
```

The top-left unfrozen cell ("B2") if panes are frozen, else None.

#### `auto_filter`

```python
auto_filter(sheet: str, ref: str | None) -> None
```

Filter buttons on the header row of `ref`; None removes them.

#### `auto_filter_range`

```python
auto_filter_range(sheet: str) -> str | None
```

The sheet's autofilter range ("A1:D20"), or None.

#### `define_name`

```python
define_name(name: str, sheet: str, ref: str, *, local: bool = False) -> OoxElement
```

A named range ("Sales" -> Sheet1!$B$2:$D$20); `local` scopes it to the sheet.

#### `defined_names`

```python
defined_names() -> dict[str, str]
```

`{name: refers_to}`; sheet-scoped names as `"Sheet!name"`; Excel's hidden
names (`_xlnm.*`) included.

#### `add_validation`

```python
add_validation(sheet: str, ref: str, **options) -> OoxElement
```

Data validation: `values=[...]` (drop-down), `source="$F$1:$F$5"`,
`whole=(1, 10)` or `decimal=(0, 1)`; `prompt=`, `error=`.

**Options** (keywords)

```python
values: list[str] | None = None
source: str | None = None
whole: tuple[int, int] | None = None
decimal: tuple[float, float] | None = None
allow_blank: bool = True
prompt: str | None = None
error: str | None = None
```

A drop-down list (`values` or a `source` range like "$F$1:$F$5"),
or a whole/decimal number between two bounds.

#### `validations`

```python
validations(sheet: str) -> list[Validation]
```

A `Validation` per data-validation rule, with `add_validation`'s options.

### Conditional formatting

#### `conditional_format`

```python
conditional_format(sheet: str, ref: str, rule: str, **options) -> OoxElement
```

Add a conditional format; see `oox.xlcond` for the rules and options.

**Details** (from `oox.xlcond`)

Excel conditional formatting.

```python
book.conditional_format("Sheet1", "B2:B20", "cell", operator=">", value=100,
                        style={"fill": "FFC7CE", "color": "9C0006"})
book.conditional_format("Sheet1", "C2:C20", "color_scale")          # red-yellow-green
book.conditional_format("Sheet1", "D2:D20", "data_bar", color="638EC6")
book.conditional_format("Sheet1", "E2:E20", "icon_set", icons="3TrafficLights1")
book.conditional_format("Sheet1", "A2:A20", "duplicates", style={"fill": "FFEB9C"})
book.conditional_format("Sheet1", "A2:A20", "text", text="North", style={"bold": True})
book.conditional_format("Sheet1", "A2:E20", "formula", formula="$E2>$D2",
                        style={"fill": "C6EFCE"})
book.conditional_formats("Sheet1")   # [ConditionalFormat(ref="B2:B20", rule="cell", ...)]
```

Rules: `cell` (operator `>`, `>=`, `<`, `<=`, `=`, `!=`, `between`,
`not between`), `formula`, `color_scale` (2 or 3 colours), `data_bar`,
`icon_set`, `top` / `bottom` (`n`, `percent=`), `above_average` /
`below_average`, `duplicates` / `unique`, `text` (`how=` contains,
not contains, begins, ends) and `blanks` / `no_blanks`.  Highlight styles are
differential formats (font colour, bold, italic, fill, border, number format).
Formulas are written relative to the range's top-left cell, as Excel does.

#### `conditional_formats`

```python
conditional_formats(sheet: str) -> list[ConditionalFormat]
```

A `ConditionalFormat` per rule, in priority order, with
`conditional_format`'s options where they apply.

### Tables

#### `add_table`

```python
add_table(
    sheet: str,
    ref: str,
    *,
    name: str | None = None,
    style: str | None = 'TableStyleMedium2',
    totals: dict | None = None,
    stripes: bool = True,
    first_column: bool = False,
    last_column: bool = False,
) -> OoxElement
```

Make `ref` an Excel table (header row = first row); `totals`
adds a totals row: {"Amount": "sum", "Region": "Total"}.  See `oox.xltables`.

**Details** (from `oox.xltables`)

Excel tables (Insert > Table): a styled range with a header row, filter
buttons, structured references and an optional totals row.

```python
book.add_table("Sales", "A1:D20", name="Sales", style="TableStyleMedium2",
               totals={"Region": "Total", "Amount": "sum", "Share": "average"})
book.set_value("Sales", "F1", Formula("SUM(Sales[Amount])"))   # structured reference
book.tables("Sales")   # [{"name": "Sales", "ref": "A1:D21", "columns": [...], ...}]
```

The first row of the range is the header; its texts become the column names
(empty or repeated headers are replaced by "Column1", "Column2" ... and written
into the cells, as Excel requires).  `totals` adds a row below the range: a
string is a label, a function name (sum, average, count, count_numbers, min,
max, stddev, var) a `SUBTOTAL` formula, whose result oox computes and caches.

#### `tables`

```python
tables(sheet: str) -> list[ExcelTable]
```

An `ExcelTable` (name, ref, columns, style, switches) for each table on the
sheet.

### Comments and hyperlinks

#### `add_comment`

```python
add_comment(
    sheet: str,
    ref: str,
    text: str,
    *,
    author: str = 'oox',
    visible: bool = False,
) -> OoxElement
```

A cell comment (note); replaces an existing one on that cell.

#### `comments`

```python
comments(sheet: str) -> dict[str, Comment]
```

`{ref: Comment}` for the sheet's comments; the "Author:" line Excel puts
first is removed.

#### `add_hyperlink`

```python
add_hyperlink(
    sheet: str,
    ref: str,
    target: str,
    *,
    text: str | None = None,
    tooltip: str | None = None,
    style: bool = True,
) -> OoxElement
```

A link in a cell: a URL, or "#Sheet!A1" / "#Name" inside the workbook.

#### `hyperlinks`

```python
hyperlinks(sheet: str) -> dict[str, Hyperlink]
```

`{ref: Hyperlink}`: `target` a URL, or `#Sheet!A1` for a link inside
the workbook.

### Printing and sheet options

#### `set_page_setup`

```python
set_page_setup(sheet: str, **options) -> None
```

Orientation, paper, fit to pages, scale, margins, print area,
repeated rows/columns, header and footer, gridlines, centring; see
`oox.xlprint.set_page_setup`.

**Options** (keywords)

```python
orientation: str | None = None
paper: str | None = None
fit_width: int | None = None
fit_height: int | None = None
scale: int | None = None
margins: dict[str, float] | None = None
print_area: str | None = None
repeat_rows: str | tuple[int, int] | None = None
repeat_cols: str | None = None
header: str | dict | None = None
footer: str | dict | None = None
gridlines: bool | None = None
center_horizontally: bool | None = None
center_vertically: bool | None = None
```

#### `page_setup`

```python
page_setup(sheet: str) -> PageSetup
```

The sheet's print setup as a `PageSetup` (`set_page_setup`'s options;
None where the sheet doesn't say; margins in points).

#### `set_sheet_options`

```python
set_sheet_options(sheet: str, **options) -> None
```

tab_color, hidden (True/False/"very"), zoom, gridlines.

**Options** (keywords)

```python
tab_color: str | None = None
hidden: bool | str | None = None
zoom: int | None = None
gridlines: bool | None = None
```

`tab_color` (hex, "" removes), `hidden` (True, False or "very" --
hidden from Excel's Unhide list too), `zoom` (percent), `gridlines`
(shown on screen).

#### `sheet_options`

```python
sheet_options(sheet: str) -> SheetOptions
```

The sheet's `SheetOptions` (tab_color, hidden, zoom, gridlines, protected).

### Protection

#### `protect_sheet`

```python
protect_sheet(
    sheet: str,
    password: str | None = None,
    *,
    allow: tuple[str, ...] = (),
) -> None
```

Protect the sheet's cells (Review > Protect Sheet), optionally with Excel's legacy
password hash.  `allow`: actions still permitted, e.g. ("sort", "autoFilter",
"formatCells").

#### `unprotect_sheet`

```python
unprotect_sheet(sheet: str) -> None
```

Remove the sheet's protection.

#### `protect_workbook`

```python
protect_workbook(
    *,
    password: str | None = None,
    structure: bool = True,
    windows: bool = False,
) -> None
```

Lock the structure (sheets can't be added, deleted, moved or renamed)
and/or windows; see `oox.protection.protect_workbook`.

**Details** (from `oox.protection`)

Document protection: Word's Review > Restrict Editing and Excel's
Review > Protect Workbook.

```python
doc.protect("readOnly", password="secret")    # or comments, trackedChanges, forms
doc.protection()       # Protection(edit="readOnly", enforced=True, password=True, ...)
doc.check_protection_password("secret")        # True
doc.unprotect()

book.protect_workbook(password="secret")       # lock the structure (sheets)
book.workbook_protection()                     # WorkbookProtection(structure=True, ...)
```

Word's password hash (ECMA-376 Part 1 §17.15.1.29, Part 4 §14.8.1): the
password is first reduced to Word 97's 32-bit key (a high-order word from an
encryption matrix, a low-order word like Excel's 16-bit hash), its four bytes
are reversed and written as an uppercase hex string, and that string
(UTF-16LE) is hashed with the salt prepended, then again `spin_count`
times with a 4-byte little-endian counter appended.  Checked against the
standard's worked example ("Example" -> 0x64CEED7E) and against a hash
Word 2019 wrote.  The hash goes into the attributes Word 2007 understands
(`w:cryptAlgorithmSid` …); SHA-1 by default, which every Word version can
check.  Protection is an editing restriction, not encryption: the text stays
readable in the file.

#### `workbook_protection`

```python
workbook_protection() -> WorkbookProtection | None
```

The workbook's `WorkbookProtection` (structure, windows, password), or None.

#### `check_workbook_password`

```python
check_workbook_password(password: str) -> bool
```

Whether `password` unlocks the workbook structure (True without a password);
checks Excel 2007's 16-bit hash and Excel 2010's salted one.

#### `unprotect_workbook`

```python
unprotect_workbook() -> None
```

Remove workbook structure and window protection.

### Pictures and charts

#### `add_picture`

```python
add_picture(
    sheet: str | OoxElement,
    image,
    anchor: str = 'A1',
    *,
    width: float | None = None,
    height: float | None = None,
    descr: str = '',
    fallback=None,
) -> OoxElement
```

Float a picture over the sheet with its top-left corner at `anchor`.

Size in points (or `Length`); one dimension keeps the aspect ratio.  An
SVG needs `fallback=` (a PNG) for Office before 2016.

#### `add_chart`

```python
add_chart(sheet: str, kind: str, data: str, **options) -> OoxElement
```

Chart of a cell range (header row = series names, first column =
categories); see `oox.xlcharts.add_chart`.

**Options** (keywords)

```python
anchor: str = 'E2'
width: float = 5400000
height: float = 2700000
series_in: str = 'columns'
name: str | None = None
```

Add a chart of `data` with its top-left corner at `anchor`; returns
its `c:chartSpace`.

#### `add_chart_sheet`

```python
add_chart_sheet(
    name: str,
    kind: str,
    sheet: str,
    data: str,
    *,
    series_in: str = 'columns',
    **options,
) -> OoxElement
```

A tab holding one chart of `data` on `sheet` (a worksheet), as
Excel's Move Chart > New sheet; options as `add_chart`.  Returns the
chart's `c:chartSpace`.

A chart sheet (its own tab) showing a chart of `data` on `sheet`.

#### `chart_sheet`

```python
chart_sheet(name: str) -> OoxElement
```

Root `<x:chartsheet>` of the named chart sheet.

#### `charts`

```python
charts(sheet: str) -> list[Chart]
```

A `Chart` for each chart on a worksheet or chart sheet, in drawing order,
with its `anchor` (top-left cell; None on a chart sheet) and `name`.

## `Formula`

```python
Formula(expression: str, cached: object = None) -> None
```

A cell formula, without the leading `=`: `Formula("SUM(A1:A3)")`.

`cached` is the result to store with it (a number, str or bool), so
readers that don't calculate see a value; Excel recalculates anyway.

Fields: `expression`, `cached`

## `CellError`

An error value read from a cell, e.g. `CellError("#DIV/0!")`.

## `StyleSheet`

The workbook's `xl/styles.xml` with de-duplicating builders.

### `describe`

```python
describe(index: int) -> dict[str, Any]
```

The formatting of cellXfs[index] as style options.

### `describe_dxf`

```python
describe_dxf(index: int) -> dict[str, Any]
```

The differential format (conditional-format highlight) at `index` as keywords
like `cell_style` takes.

### `dxf`

```python
dxf(**fmt: Any) -> int
```

Index into `dxfs`: a differential format (only what changes), as
conditional formatting and tables use.  Options: bold, italic,
underline, strike, color, fill, border, border_color, number_format.

### `number_format_code`

```python
number_format_code(num_fmt_id: int) -> str
```

The format code for number format id `num_fmt_id` (built-in or custom).

### `number_format_id`

```python
number_format_id(code: str) -> int
```

The id for format `code`: a locale-invariant built-in id, or a custom format
added on first use.

### `xf`

```python
xf(**fmt: Any) -> int
```

Index into cellXfs for this combination of formatting.

## `oox.spreadsheet`

SpreadsheetML (.xlsx): the `OpenDocumentSpreadsheet()` equivalent.

```python
from oox.spreadsheet import SpreadsheetDocument, Formula

book = SpreadsheetDocument.new()            # one sheet, "Sheet1"
book.append_row("Sheet1", ["Item", "Qty", "Price", "Total"])
book.append_row("Sheet1", ["Pens", 3, 1.5, Formula("B2*C2")])
book.add_picture("Sheet1", "logo.png", anchor="F2")
book.save("report.xlsx")

book = SpreadsheetDocument.open("from_excel.xlsx")   # reading works on any producer
book.value("Sheet1", "B2")          # 3, "text", True, 1.5, CellError("#DIV/0!") ...
book.values("Sheet1")               # [[...], [...]] dense, from A1
book.set_value("Sheet1", "D7", 42)  # keeps the cell's existing style
```

Strings are written as inline strings (`t="inlineStr"`), which every
consumer reads and which avoid managing a shared-string table.  A
shared-strings and number-format/style manager is roadmap phase 2.

### `cell`

```python
cell(ref: str, value: object) -> OoxElement
```

Build `<x:c>` for a Python value.

`None` -> empty cell, `bool` -> boolean, `int`/`float` -> number,
`str` -> inline string, `Formula` -> formula.  Dates need a date
number format, which lives in the workbook's stylesheet: use the
document's `set_value` / `append_row`, which add one.

### `row`

```python
row(index: int, values: Iterable[object], start_col: int = 1) -> OoxElement
```

Build `<x:row r=index>` with one cell per value (`None` cells are skipped).

### `cell_ref`

```python
cell_ref(col: int, row: int) -> str
```

### `column_letter`

```python
column_letter(col: int) -> str
```

1 -> "A", 27 -> "AA".

### `split_ref`

```python
split_ref(ref: str) -> tuple[int, int]
```

"B3" -> (2, 3)  (1-based column, row).

### `excel_serial`

```python
excel_serial(value: _dt.date | _dt.datetime) -> float
```

Date/datetime -> Excel serial number (1900 date system).

### `from_excel_serial`

```python
from_excel_serial(serial: float) -> _dt.datetime
```

Excel serial number (1900 date system) -> datetime.
