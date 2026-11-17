# Excel

`SpreadsheetDocument` creates, reads and edits .xlsx files. Sheets are addressed by
name and cells by A1 references. Every block below runs as written, in order. The
[reference](../reference/excel.md) lists every option.

## Workbooks and sheets

```python
from oox import SpreadsheetDocument

book = SpreadsheetDocument.new("Sales")         # one worksheet, as Excel requires
book.add_sheet("Lookup")
print(book.sheet_names)                         # ['Sales', 'Lookup']
```

`sheet_names` lists worksheets; `tab_names` lists every tab, chart sheets included
(`chart_sheet_names` has just those). `book.sheet(name)` returns a worksheet's root
element for anything the conveniences don't cover.

## Writing values

```python
import datetime as dt

from oox.spreadsheet import Formula

book.append_row("Sales", ["Region", "Target", "Actual", "Updated"])
book.append_row("Sales", ["North", 100, 120, dt.date(2026, 11, 2)])
book.append_row("Sales", ["South", 100, 30.5, dt.datetime(2026, 11, 9, 14, 30)])
book.append_row("Sales", ["West", 80, 84, None])
book.set_value("Sales", "E1", "Hit rate")
# formula with its cached result
book.set_value("Sales", "E2", Formula("C2/B2", 1.2))
book.set_value("Sales", "F1", True)
```

Strings, numbers, booleans, dates, datetimes, times, `None` (an empty cell) and
`Formula` are written with the right cell type; dates and times also get a date
format. `append_row` adds below the last row (`start_col=3` starts at column C);
`set_value` writes anywhere, keeping rows and cells in the order Excel requires, and
keeps an existing cell's style unless you pass one.

A `Formula`'s cached result is what readers that don't calculate see (and it spares
you Excel's "save changes?" prompt). Tables' totals rows are computed by oox.

## Reading values

```python
print(book.value("Sales", "C3"))                 # 30.5
print(book.value("Sales", "D2"))                 # 2026-11-02 00:00:00 (a datetime)
print(book.values("Sales")[1])
# ['North', 100, 120, datetime(2026, 11, 2, 0, 0), 1.2, None]
print(book.cells("Sales")["E2"])                 # 1.2: the cached result
print(book.value("Sales", "E2", formulas=True))  # Formula(expression='C2/B2')
```

`values` gives a dense list of rows from A1 (as wide as the widest row); `cells`
gives `{ref: value}` for the stored cells. Date-formatted numbers come back as `datetime` (or `time`) objects;
`dates=False` gives the serial number. Error values come back as `CellError("#DIV/0!")`.
Single-cell reads are fast (cached and checked on every lookup), so reading cell by
cell from a large sheet is fine.

These readers work on any workbook, including shared strings with rich text,
phonetic runs and escaped characters, as Excel and LibreOffice write them.

## Cell styles and number formats

```python
header = book.cell_style(bold=True, color="FFFFFF", fill="2F5496", align="center",
                         border="thin")
money = book.cell_style(number_format="#,##0.00;[Red]-#,##0.00", border="thin")

book.style_range("Sales", "A1:F1", header)
book.style_range("Sales", "B2:C4", money)
book.style_range("Sales", "D2:D4", number_format="d mmm yyyy")
book.set_value("Sales", "E3", Formula("C3/B3", 0.305), style={"number_format": "0.0%"})

print(book.display_text("Sales", "C3"))          # 30.50
print(book.display_text("Sales", "E3"))          # 30.5%
print(book.cell_format("Sales", "A1").fill)   # 2F5496
```

`cell_style` returns a shared style index: equal styles share one entry, and entries
already in an Excel-made stylesheet are reused. Wherever a style is accepted you can
pass the index or a dict of the same options: `number_format`, `bold`, `italic`,
`underline`, `strike`, `size`, `font`, `color`, `fill`, `border` (a style name, or
`{"bottom": "double"}`), `border_color`, `align`, `valign`, `wrap`, `indent`,
`rotation`.

`display_text` renders a cell the way Excel shows it in an English-locale Excel:
sections, thousands separators, percentages, scientific notation, dates and times,
literal text and Excel's 11-character General format.

## Layout

```python
book.set_column_width("Sales", "A", 14)           # characters, as in Excel's dialog
book.set_column_width("Sales", "B:C", 11)
book.set_row_height("Sales", 1, 24)               # points
# widths from the longest displayed text
book.autofit_columns("Sales", "D:F")
book.merge_cells("Sales", "A6:C6")
book.set_value("Sales", "A6", "Preliminary figures")
book.set_frozen_panes("Sales", "B2")                  # first row and column stay visible
book.auto_filter("Sales", "A1:F4")
book.define_name("Actuals", "Sales", "C2:C4")
book.add_validation("Sales", "A2:A4", values=["North", "South", "West", "East"])
book.add_validation("Sales", "B2:B4", whole=(0, 1000), error="Targets are 0 to 1000")

print(book.frozen_panes("Sales"), book.merged_cells("Sales"), book.defined_names())
```

## Conditional formatting

```python
book.conditional_format("Sales", "C2:C4", "formula", formula="$C2<$B2",
                        style={"fill": "FFC7CE", "color": "9C0006"})
book.conditional_format("Sales", "E2:E4", "data_bar", color="638EC6")
book.conditional_format("Sales", "B2:B4", "icon_set", icons="3Arrows")
book.conditional_format("Sales", "A2:A4", "text", text="North", style={"bold": True})

print([rule.rule for rule in book.conditional_formats("Sales")])
```

Rules: `cell` (with `operator` and `value`), `formula`, `color_scale`, `data_bar`,
`icon_set`, `top`/`bottom`, `above_average`/`below_average`, `duplicates`/`unique`,
`text` and `blanks`/`no_blanks`. Formulas are relative to the range's top-left cell,
as in Excel's dialog.

## Tables

```python
orders = SpreadsheetDocument.new("Orders")
orders.append_row("Orders", ["Region", "Units", "Amount"])
for row in (["North", 12, 1250.0], ["South", 8, 980.5], ["West", 5, 410.0]):
    orders.append_row("Orders", row)
orders.add_table("Orders", "A1:C4", name="Orders2026",
                 totals={"Region": "Total", "Units": "sum", "Amount": "sum"})
orders.set_value("Orders", "E1", Formula("SUM(Orders2026[Amount])", 2640.5))

print(orders.tables("Orders")[0].ref)          # A1:C5: the totals row was added
print(orders.value("Orders", "C5"))               # 2640.5, computed by oox
```

Header names come from the first row (empty or repeated ones are fixed the way Excel
requires); table names are unique per workbook. Structured references
(`Orders2026[Amount]`) work in formulas.

## Comments and hyperlinks

```python
book.add_comment("Sales", "C3", "Storm damage closed two stores.", author="Ada")
book.add_hyperlink("Sales", "H1", "https://ecma-international.org/", text="About OOXML")
book.add_hyperlink("Sales", "H2", "#Lookup!A1", text="Region codes")
note = book.comments("Sales")["C3"]
print(note.author, note.text)        # Ada Storm damage closed two stores.
print({cell: link.target for cell, link in book.hyperlinks("Sales").items()})
# {'H1': 'https://ecma-international.org/', 'H2': '#Lookup!A1'}
```

## Printing and sheet options

```python
book.set_page_setup("Sales", orientation="landscape", paper="A4", fit_width=1,
                repeat_rows="1:1", footer={"right": "Page {page} of {pages}"})
book.set_sheet_options("Sales", tab_color="2F5496", zoom=120)
book.set_sheet_options("Lookup", hidden=True)          # "very" hides it from Unhide too
print(book.page_setup("Sales").orientation, book.sheet_options("Lookup"))
```

## Protection

```python
# cells locked, sorting allowed
book.protect_sheet("Sales", allow=("sort", "autoFilter"))
# sheets can't be added or renamed
book.protect_workbook(password="secret")
print(book.workbook_protection())
# WorkbookProtection(structure=True, windows=False, password=True)
print(book.check_workbook_password("secret"))
```

As in Excel, protection is an editing restriction stored with Excel's password hash,
not encryption.

## Pictures and charts

```python
from oox.units import Cm

book.add_picture("Sales", "logo.png", anchor="J2", width=Cm(4))
book.add_chart("Sales", "column", "A1:C4", anchor="J8", title="Target and actual",
               legend="bottom")
print(book.charts("Sales")[0].kind)            # column
```

Charts take their data from cell ranges: the header row names the series and the
first column holds the categories. See [Charts](charts.md).

## Saving

```python
# text into the shared-strings table, as Excel stores it
book.share_strings()
assert not book.validate(target="Office2007")
book.save("sales.xlsx")
```

`share_strings()` is optional: inline strings are valid and every reader handles them,
but Excel itself always uses the shared-strings table.
