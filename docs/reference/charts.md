# Charts

Chart data and the builder and reader the document methods use.

## `Series`

```python
Series(
    name: str | None,
    values: Sequence[float | int | None],
    categories: Sequence | None = None,
    color: str | None = None,
    name_ref: str | None = None,
    values_ref: str | None = None,
    categories_ref: str | None = None,
    kind: str | None = None,
    axis: str = 'primary',
    point_colors: Sequence[str | None] = (),
) -> None
```

One data series.  `categories` overrides the chart's categories (for
scatter: the x values).  The `*_ref` fields are cell references
(`Sheet1!$B$2:$B$5`); without them the values are written as literals.

Fields: `name`, `values`, `categories`, `color`, `name_ref`, `values_ref`, `categories_ref`, `kind`, `axis`, `point_colors`

## `oox.charts`

Charts: DrawingML chart parts (`c:chartSpace`), shared by Excel, Word and
PowerPoint.

```python
from oox.charts import Series, build_chart, chart_info

space = build_chart("column", ["Q1", "Q2", "Q3"],
                    [Series("North", [120, 135, 150]), Series("South", [95, 110, 99])],
                    title="Sales", y_title="EUR k", legend="bottom", labels="value")
chart_info(space).kind           # "column"
```

The applications wrap this: `SpreadsheetDocument.add_chart` (series from a
cell range), `WordprocessingDocument.add_chart` and
`PresentationDocument.add_chart` (series from Python lists, stored in an
embedded workbook so "Edit Data" works).  Every series carries cached values,
so viewers draw the chart without recalculating anything.

Kinds and groupings:

=========  ====================================================  ==============
kind       grouping                                              Excel type
=========  ====================================================  ==============
column     clustered (default), stacked, percent                 51, 52, 53
bar        clustered (default), stacked, percent (horizontal)    57, 58, 59
line       standard (default), stacked, percent; markers=        65 / 4, ...
area       standard (default), stacked, percent                  1, 76, 77
pie        (one series; colors= per slice)                       5
doughnut   (hole_size=)                                          -4120
scatter    lines=, markers=, smooth=  (categories are x values)  -4169, 74, ...
=========  ====================================================  ==============

`chart_info` reports the type in the same vocabulary for any chart part --
oox's own, Excel's or LibreOffice's -- so it can be compared with
`Chart.ChartType` in Office (see `EXCEL_CHART_TYPES`).

### `build_chart`

```python
build_chart(
    kind: str,
    categories: Sequence,
    series: Iterable[Series],
    *,
    categories_ref: str | None = None,
    grouping: str | None = None,
    title: str | None = None,
    x_title: str | None = None,
    y_title: str | None = None,
    legend: str | None = 'right',
    labels=None,
    label_format: str | None = None,
    colors: Sequence[str] | None = None,
    number_format: str | None = None,
    markers: bool = True,
    lines: bool | None = None,
    smooth: bool = False,
    gridlines: bool = True,
    y_min: float | None = None,
    y_max: float | None = None,
    hole_size: int = 50,
    gap_width: int = 150,
    first_slice: int = 0,
    kinds: Mapping[str | int, str] | None = None,
    secondary: Iterable[str | int] = (),
    y2_title: str | None = None,
    y2_min: float | None = None,
    y2_max: float | None = None,
    y2_number_format: str | None = None,
    external_rid: str | None = None,
) -> OoxElement
```

A complete `c:chartSpace`.

`colors` are RRGGBB strings: one per series, or one per slice for pie
and doughnut; missing ones come from `PALETTE`.  `labels` is one of / a list of "value", "percent",
"category", "series".  `number_format` applies to the value axis and the
cached values; `label_format` to the data labels.  `lines` (scatter)
defaults to False: markers only.  `external_rid` links the embedded
workbook (Word, PowerPoint).

Combination charts: `kinds` maps series (by name or 0-based index) to
another kind, e.g. `{"Margin": "line"}` on a column chart; `secondary`
names the series drawn against a second value axis on the right, titled
`y2_title`.  Columns, lines and areas combine; bars only with bars.
A `Series` can also carry `kind` and `axis` itself.

### `chart_info`

```python
chart_info(space: OoxElement)
```

Describe a `c:chartSpace` as a `Chart` record (`oox.records`):
kind (Excel's vocabulary, see `EXCEL_CHART_TYPES`), title, legend
position, axis titles, value number format, data labels, and its
`Series` with names, values, categories, refs, colours, kind (its
plot's, for combination charts) and axis ("primary" or "secondary"; the
secondary value axis's title is `y2_title`).  Series come in plot order
(`c:order`).

### `PALETTE`

```python
PALETTE = ('4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300', '255E91', '43682B', '698ED0', 'F1975A', 'B7B7B7', 'FFCD33', '7CAFDD', '8CC168')
```
