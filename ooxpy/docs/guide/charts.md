# Charts

Charts work the same way in all three formats; only where the data comes from differs.
In Excel a chart refers to cell ranges. In Word and PowerPoint you pass the data from
Python and oox stores it in an embedded workbook, so Office's "Edit Data" opens it.
Every chart has a reader too.

## Kinds

| `kind` | `grouping` | Notes |
|---|---|---|
| `column`, `bar`, `line`, `area` | `clustered` (default for column/bar), `standard`, `stacked`, `percent` | lines have markers by default (`markers=False`) |
| `pie`, `doughnut` | | one series; `colors` per slice, `hole_size`, `first_slice` |
| `scatter` | | numeric categories; markers only unless `lines=True`, `smooth=True` |

## In Excel

```python
from oox import SpreadsheetDocument

book = SpreadsheetDocument.new("Sales")
book.append_row("Sales", ["Region", "Q1", "Q2", "Q3"], style={"bold": True})
for row in [("North", 120, 135, 150), ("South", 95, 110, 99), ("West", 80, 84.5, 91)]:
    book.append_row("Sales", list(row))

book.add_chart("Sales", "column", "A1:D4", anchor="F2", title="Quarterly sales",
               x_title="Region", y_title="EUR k", legend="bottom", labels="value")
book.add_chart("Sales", "line", "A1:D4", anchor="F18", series_in="rows", title="Trend")
book.add_chart("Sales", "pie", "A1:B4", anchor="N2", title="Q1 share", labels="percent")
book.add_chart_sheet("Overview", "area", "Sales", "A1:D4", grouping="percent",
                     title="Share of quarter", number_format="0%")

print([c["type"] for c in book.charts("Sales")])
# ['column', 'line_markers', 'pie']
print(book.tab_names)                                # ['Sales', 'Overview']
book.save("charts.xlsx")
```

The first row of the range names the series and the first column holds the categories
(`series_in="rows"` reads series across rows instead). The chart's cached values are
filled from the cells, so viewers that don't calculate show it right away. `anchor`
is the top-left cell; `width` and `height` are in EMU. A chart sheet is a tab of its own.

## In Word and PowerPoint

```python
from oox import PresentationDocument, WordprocessingDocument
from oox.units import Cm, Inches

doc = WordprocessingDocument.new()
doc.add_paragraph("Sales by quarter:")
doc.add_chart("column", ["Q1", "Q2", "Q3"], {"North": [120, 135, 150],
                                             "South": [95, 110, 99]},
              title="Sales", legend="bottom")
doc.add_chart("doughnut", ["Web", "Shop", "Phone"], {"Channel": [55, 30, 15]},
              title="Channels", labels="percent", width=Cm(9), height=Cm(7))

deck = PresentationDocument.new()
slide = deck.add_slide("Blank")
deck.add_chart(slide, "scatter", [1, 2, 3, 4, 5],
               {"Measured": [2.1, 3.9, 6.2, 7.8, 10.1]},
               Inches(1), Inches(1), Inches(8), Inches(5), lines=True, smooth=True,
               legend=None, title="Fit")
print([c["type"] for c in doc.charts()], [c["type"] for c in deck.charts(slide)])
# ['column', 'doughnut'] ['scatter_smooth']
```

The arguments are the kind, the categories, then the series as `{name: values}` (or a
list of `oox.charts.Series` for per-series colours and kinds). In Word the chart is an
inline picture in a new paragraph (`paragraph=` puts it in an existing one); on a slide
it's a graphic frame at `x, y` with size `cx, cy`.

## Options

| Option | Meaning |
|---|---|
| `title`, `x_title`, `y_title` | chart and axis titles (no title: none is shown) |
| `legend` | `"right"` (default), `"left"`, `"top"`, `"bottom"`, `"top_right"` or `None` |
| `labels` | data labels: `"value"`, `"percent"`, `"category"`, `"series"` or a list of them; `label_format` |
| `number_format` | for the value axis and cached values, e.g. `"#,##0"` or `"0%"` |
| `y_min`, `y_max`, `gridlines` | value axis bounds and major gridlines |
| `colors` | RRGGBB per series (per slice for pie and doughnut); the default palette is the theme's accent colours |
| `gap_width`, `hole_size`, `first_slice`, `markers`, `lines`, `smooth` | per kind |

Colours are always written explicitly, so every reader draws the same colours.

## Combination charts and a secondary axis

```python
book.append_row("Sales", ["Margin", 0.21, 0.24, 0.26])
book.add_chart("Sales", "column", "A1:D5", anchor="N18", series_in="rows",
               kinds={"Margin": "line"}, secondary=["Margin"],
               y_title="EUR k", y2_title="Margin", y2_number_format="0%",
               title="Sales and margin")
combo = book.charts("Sales")[-1]
print([(s["name"], s["type"], s["axis"]) for s in combo["series"]][-2:])
# [('West', 'column', 'primary'), ('Margin', 'line_markers', 'secondary')]
```

`kinds` gives series another kind (columns, lines and areas combine; bars only with
bars); `secondary` puts series on a value axis on the right, with its own title, bounds
and number format.

## Reading charts

```python
info = book.charts("Sales")[0]
print(info["title"], info["legend"], info["series"][0]["values"])
# Quarterly sales bottom [120, 95, 80]
```

`charts()` works on charts from any producer (Excel, LibreOffice, openpyxl, XlsxWriter)
and reports the type in Excel's own vocabulary (`column_stacked`, `line_markers`,
`scatter_smooth` …), with the title, legend, axis titles, number format, data labels and
each series' name, values, categories, references, colour, type and axis. Charts that
were saved without cached values (openpyxl writes them that way) are read from the
cells.
