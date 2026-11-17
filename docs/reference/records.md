# Records

What the readers return: records with named fields, and the formatting values the writers take.

## `oox.records`

What the readers return: small frozen records with named fields.

```python
for link in doc.hyperlinks():
    print(link.text, link.target)
book.comments("Sheet1")["B2"].author
deck.shapes(slide)[0].kind
```

One concept has one class in all three formats (a `Hyperlink` from Word,
Excel or PowerPoint, a `Comment` from Word or Excel).  Where a writer
exists, field names are its keyword names, so a record can be written back:
`book.set_page_setup("T", **book.page_setup("S").as_dict())`.  Lengths
are in points.

Fields are only ever added, never renamed or removed, within a major
version; build records with keywords if you build them at all.  `as_dict()`
gives a plain dict (without the XML elements some records carry), for JSON.

### `Record`

*class*

Base class: `as_dict()` for JSON and for passing back to writers.

## `Hyperlink`

*Based on* `Record`.

```python
Hyperlink(
    text: str | None,
    target: str | None = None,
    anchor: str | int | None = None,
    tooltip: str | None = None,
    kind: str = 'url',
    cell: str | None = None,
    shape: str | None = None,
) -> None
```

A link.  `target` is the address as the format's writer takes it (a
URL; in Excel also `#Sheet!A1`; in PowerPoint also a jump such as
"next"); `anchor` a place inside the document (a Word bookmark, a
PowerPoint slide number); `kind` is "url", "bookmark", "location",
"slide", "jump" or "other"; `cell` (Excel) and `shape` (PowerPoint)
say where it is.

Fields: `text`, `target`, `anchor`, `tooltip`, `kind`, `cell`, `shape`

## `Comment`

*Based on* `Record`.

```python
Comment(
    text: str,
    author: str | None = None,
    date: str | None = None,
    initials: str | None = None,
    anchor_text: str | None = None,
    cell: str | None = None,
    id: int | None = None,
) -> None
```

A comment.  `anchor_text` is the commented text (Word); `cell` the
commented cell (Excel).

Fields: `text`, `author`, `date`, `initials`, `anchor_text`, `cell`, `id`

## `Chart`

*Based on* `Record`.

```python
Chart(
    kind: str | None,
    plot_kinds: list = <factory>,
    title: str | None = None,
    legend: str | None = None,
    x_title: str | None = None,
    y_title: str | None = None,
    y2_title: str | None = None,
    number_format: str | None = None,
    labels: list = <factory>,
    series: list = <factory>,
    categories: list | None = None,
    anchor: str | None = None,
    name: str | None = None,
    rid: str | None = None,
) -> None
```

A chart, in Excel's `XlChartType` vocabulary: `kind` ("column",
"pie" ...), `plot_kinds` (one per plot of a combination chart);
`series` are `oox.charts.Series` (each with its `kind` and
`axis`).  `anchor` is the top-left cell and `name` the chart's
name (Excel); `rid` the chart part's relationship id.

Fields: `kind`, `plot_kinds`, `title`, `legend`, `x_title`, `y_title`, `y2_title`, `number_format`, `labels`, `series`, `categories`, `anchor`, `name`, `rid`

## `CellFormat`

*Based on* `Record`.

```python
CellFormat(
    number_format: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strike: bool | None = None,
    size: float | None = None,
    font: str | None = None,
    color: str | None = None,
    fill: str | None = None,
    border: Any = None,
    align: str | None = None,
    valign: str | None = None,
    wrap: bool | None = None,
) -> None
```

A cell's (or a highlight's) formatting, as `cell_style` options;
None where the format doesn't say.

Fields: `number_format`, `bold`, `italic`, `underline`, `strike`, `size`, `font`, `color`, `fill`, `border`, `align`, `valign`, `wrap`

## `Note`

*Based on* `Record`.

```python
Note(id: int, text: str) -> None
```

A footnote or endnote.

Fields: `id`, `text`

## `ContentControl`

*Based on* `Record`.

```python
ContentControl(
    tag: str | None,
    title: str | None = None,
    type: str | None = None,
    placeholder: bool = False,
    value: Any = None,
    items: list | None = None,
    date: Any = None,
) -> None
```

A content control (form field): `type` is "text", "rich_text",
"dropdown", "combo", "date", "checkbox" ...; `placeholder` True while it
shows its placeholder text.

Fields: `tag`, `title`, `type`, `placeholder`, `value`, `items`, `date`

## `Picture`

*Based on* `Record`.

```python
Picture(
    kind: str,
    name: str | None = None,
    descr: str | None = None,
    width: float | None = None,
    height: float | None = None,
    floating: bool = False,
    wrap: str | None = None,
    embed: str | None = None,
    link: str | None = None,
    svg: str | None = None,
    element: OoxElement | None = None,
) -> None
```

A picture.  `kind` "drawing" (DrawingML) or "vml"; `width` and
`height` in points; `wrap` "inline" for inline pictures; `embed` /
`link` / `svg` are relationship ids (`doc.image_of(picture)` gives
the bytes).

Fields: `kind`, `name`, `descr`, `width`, `height`, `floating`, `wrap`, `embed`, `link`, `svg`, `element`

## `Protection`

*Based on* `Record`.

```python
Protection(
    edit: str | None,
    enforced: bool = False,
    password: bool = False,
    algorithm: str | None = None,
    formatting: bool = False,
) -> None
```

Word's editing restriction: `edit` "readOnly", "comments",
"trackedChanges" or "forms"; `enforced`; `password` whether one is set.

Fields: `edit`, `enforced`, `password`, `algorithm`, `formatting`

## `Revision`

*Based on* `Record`.

```python
Revision(
    kind: str,
    author: str | None = None,
    date: str | None = None,
    text: str = '',
) -> None
```

A tracked change: `kind` "insert", "delete", "move-from", "move-to",
"paragraph-insert", "paragraph-delete", "row-insert", "row-delete" or
"format".

Fields: `kind`, `author`, `date`, `text`

## `TextBox`

*Based on* `Record`.

```python
TextBox(
    text: str,
    paragraphs: int = 0,
    vml_only: bool = False,
    name: str | None = None,
) -> None
```

A text box, counted once (`vml_only`: no DrawingML shape, as in
files from Word 2007 and older).

Fields: `text`, `paragraphs`, `vml_only`, `name`

## `ConditionalFormat`

*Based on* `Record`.

```python
ConditionalFormat(
    ref: str,
    rule: str | None,
    type: str | None = None,
    priority: int = 0,
    operator: str | None = None,
    value: Any = None,
    values: list | None = None,
    formula: str | None = None,
    formulas: list = <factory>,
    colors: list | None = None,
    color: str | None = None,
    icons: str | None = None,
    n: int | None = None,
    percent: bool | None = None,
    text: str | None = None,
    how: str | None = None,
    stop: bool = False,
    style: CellFormat | None = None,
) -> None
```

A conditional formatting rule, as `conditional_format` options:
`rule` is oox's name ("cell", "data_bar", "top" ...; None for rules
oox doesn't write), `type` Excel's; `style` the highlight.

Fields: `ref`, `rule`, `type`, `priority`, `operator`, `value`, `values`, `formula`, `formulas`, `colors`, `color`, `icons`, `n`, `percent`, `text`, `how`, `stop`, `style`

## `Validation`

*Based on* `Record`.

```python
Validation(
    ref: str,
    type: str,
    values: list | None = None,
    source: str | None = None,
    whole: tuple | None = None,
    decimal: tuple | None = None,
    allow_blank: bool = True,
    prompt: str | None = None,
    error: str | None = None,
) -> None
```

A data validation, as `add_validation` options: `values` (a
list), `source` (a range), `whole` or `decimal` (min, max).

Fields: `ref`, `type`, `values`, `source`, `whole`, `decimal`, `allow_blank`, `prompt`, `error`

## `ExcelTable`

*Based on* `Record`.

```python
ExcelTable(
    name: str | None,
    ref: str,
    columns: list = <factory>,
    style: str | None = None,
    stripes: bool = True,
    first_column: bool = False,
    last_column: bool = False,
    header: bool = True,
    totals_row: bool = False,
) -> None
```

An Excel table (Insert > Table): `ref`, column names, style and its
switches as `add_table` takes them; `totals_row` whether it has one.

Fields: `name`, `ref`, `columns`, `style`, `stripes`, `first_column`, `last_column`, `header`, `totals_row`

## `PageSetup`

*Based on* `Record`.

```python
PageSetup(
    orientation: str | None = None,
    paper: Any = None,
    fit_width: int | None = None,
    fit_height: int | None = None,
    scale: int | None = None,
    margins: dict | None = None,
    print_area: str | None = None,
    repeat_rows: str | None = None,
    repeat_cols: str | None = None,
    header: Any = None,
    footer: Any = None,
    gridlines: bool | None = None,
    center_horizontally: bool | None = None,
    center_vertically: bool | None = None,
) -> None
```

A sheet's print setup, as `set_page_setup` options (None: not set);
`margins` in points.

Fields: `orientation`, `paper`, `fit_width`, `fit_height`, `scale`, `margins`, `print_area`, `repeat_rows`, `repeat_cols`, `header`, `footer`, `gridlines`, `center_horizontally`, `center_vertically`

## `SheetOptions`

*Based on* `Record`.

```python
SheetOptions(
    tab_color: str | None = None,
    hidden: bool | str = False,
    zoom: int = 100,
    gridlines: bool = True,
    protected: bool = False,
) -> None
```

A sheet's options, as `set_sheet_options` takes them, plus
`protected`.

Fields: `tab_color`, `hidden`, `zoom`, `gridlines`, `protected`

## `WorkbookProtection`

*Based on* `Record`.

```python
WorkbookProtection(
    structure: bool = False,
    windows: bool = False,
    password: bool = False,
) -> None
```

Workbook protection, as `protect_workbook` takes it; `password`
whether one is set.

Fields: `structure`, `windows`, `password`

## `Placeholder`

*Based on* `Record`.

```python
Placeholder(
    type: str,
    idx: int,
    name: str | None = None,
    element: OoxElement | None = None,
) -> None
```

A placeholder on a slide: `type` ("title", "body", "obj", "pic"
...) and `idx`.

Fields: `type`, `idx`, `name`, `element`

## `Shape`

*Based on* `Record`.

```python
Shape(
    kind: str,
    name: str | None = None,
    text: str | None = None,
    placeholder: Placeholder | None = None,
    geometry: str | None = None,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    element: OoxElement | None = None,
) -> None
```

Something on a slide: `kind` "placeholder", "text_box", "shape",
"picture", "table", "chart", "group" or "other"; position and size in
points (None for a placeholder that inherits them from its layout).

Fields: `kind`, `name`, `text`, `placeholder`, `geometry`, `x`, `y`, `width`, `height`, `element`

## `SlideTableInfo`

*Based on* `Record`.

```python
SlideTableInfo(
    name: str | None,
    rows: list = <factory>,
    merges: list = <factory>,
    style: str | None = None,
    style_name: str | None = None,
    header: bool = False,
    banded: bool = False,
    first_column: bool = False,
) -> None
```

A table on a slide (`deck.table(slide, i)` edits it): cell texts,
merges as "A1:B1" ranges, style and switches as `add_table` takes them.

Fields: `name`, `rows`, `merges`, `style`, `style_name`, `header`, `banded`, `first_column`

## `SlideSection`

*Based on* `Record`.

```python
SlideSection(name: str, slides: list = <factory>) -> None
```

A section of slides: its name and slide numbers (1-based).

Fields: `name`, `slides`

## `ListItem`

```python
ListItem(
    paragraph: OoxElement,
    label: str,
    level: int,
    num_id: int,
    abstract_id: int,
    format: str,
    font: str | None,
    suffix: str,
    raw_label: str,
    picture: OoxElement | None = None,
    story: str = 'body',
) -> None
```

A list paragraph and the label Word shows for it.

`label` is readable text ("2.", "1.3.", "•"); `raw_label` keeps the
characters as stored, which for bullets are often symbol-font codes
("\uf0b7" in `font` Symbol).  `suffix` is what follows the label:
"tab", "space" or "nothing".  `picture` is the image element of a
picture bullet (`doc.image_of(item.picture)` gives its bytes).
`story` is "body", "header", "footer", "footnotes", "endnotes" or
"comments".

Fields: `paragraph`, `label`, `level`, `num_id`, `abstract_id`, `format`, `font`, `suffix`, `raw_label`, `picture`, `story`

## `ImageData`

An image's bytes and where they live; see `OoxDocument.image_of`.

## `Formatting`

Resolves effective formatting for one document; see the module docs.
Make one with `doc.formatting()`.

### `mark`

```python
mark(p: OoxElement) -> RunFormat
```

The formatting of `p`'s paragraph mark (what an empty paragraph,
and by default a list label, looks like).

### `paragraph`

```python
paragraph(p: OoxElement) -> ParagraphFormat
```

The paragraph formatting Word applies to `p`.

### `run`

```python
run(r: OoxElement) -> RunFormat
```

The character formatting Word applies to the run `r`.

## `ParagraphFormat`

```python
ParagraphFormat(
    style: str | None = None,
    align: str = 'left',
    space_before: float = 0.0,
    space_after: float = 0.0,
    line_spacing: LineSpacing = LineSpacing(rule='auto', value=1.0),
    indent_left: float = 0.0,
    indent_right: float = 0.0,
    first_line: float = 0.0,
    keep_with_next: bool = False,
    keep_together: bool = False,
    page_break_before: bool = False,
    widow_control: bool = False,
    contextual_spacing: bool = False,
    outline_level: int | None = None,
    shading: str | None = None,
    borders: Borders = <factory>,
    tabs: tuple[Tab, ...] = (),
    numbering: tuple[int, int] | None = None,
) -> None
```

Effective paragraph formatting, in points.  `first_line` is negative
for a hanging indent; `outline_level` is 0-8 or None for body text;
`style` is the paragraph style's name; `numbering` is
`(num_id, level)` for a list paragraph.

Fields: `style`, `align`, `space_before`, `space_after`, `line_spacing`, `indent_left`, `indent_right`, `first_line`, `keep_with_next`, `keep_together`, `page_break_before`, `widow_control`, `contextual_spacing`, `outline_level`, `shading`, `borders`, `tabs`, `numbering`

## `RunFormat`

```python
RunFormat(
    bold: bool = False,
    italic: bool = False,
    underline: str | None = None,
    strike: bool = False,
    double_strike: bool = False,
    size: float = 10.0,
    font: str | None = 'Times New Roman',
    font_east_asia: str | None = None,
    font_complex: str | None = None,
    color: str | None = None,
    highlight: str | None = None,
    superscript: bool = False,
    subscript: bool = False,
    caps: bool = False,
    small_caps: bool = False,
    shading: str | None = None,
    letter_spacing: float = 0.0,
    hidden: bool = False,
    language: str | None = None,
    border: Border | None = None,
    style: str | None = None,
) -> None
```

Effective run (character) formatting.  `size`, `letter_spacing` in
points; colors as hex ("C00000") or None for automatic; `underline` is
None or Word's underline style ("single", "double" ...); `style` is the
character style's name.

Fields: `bold`, `italic`, `underline`, `strike`, `double_strike`, `size`, `font`, `font_east_asia`, `font_complex`, `color`, `highlight`, `superscript`, `subscript`, `caps`, `small_caps`, `shading`, `letter_spacing`, `hidden`, `language`, `border`, `style`

## `LineSpacing`

Line spacing as Word stores it.  `rule` is "auto" (`value` is a
multiple of single spacing), "exact" or "at_least" (`value` in points).

```python
LineSpacing.multiple(1.15)   LineSpacing.exactly(12)   LineSpacing.at_least(14)
```

### `at_least`

```python
at_least(points: float) -> LineSpacing
```

*classmethod*

At least `points` per line (more for larger text).

### `exactly`

```python
exactly(points: float) -> LineSpacing
```

*classmethod*

Exactly `points` per line.

### `multiple`

```python
multiple(n: float) -> LineSpacing
```

*classmethod*

`n` times single spacing (1.15, 1.5, 2).

## `Border`

One border line: `style` ("single", "double", "dotted", ...), `width`
and `space` (gap to the text) in points, `color` as hex or None (auto).

## `Borders`

```python
Borders(
    top: Border | None = None,
    left: Border | None = None,
    bottom: Border | None = None,
    right: Border | None = None,
    between: Border | None = None,
    bar: Border | None = None,
) -> None
```

Paragraph borders; each side is a `Border` or None.

Fields: `top`, `left`, `bottom`, `right`, `between`, `bar`

## `Tab`

A tab stop: `position` in points from the left indent's origin,
`align` ("left", "center", "right", "decimal", "bar"), `leader`
("none", "dot", "hyphen", "underscore", "heavy", "middleDot").
