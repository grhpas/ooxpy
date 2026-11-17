# PowerPoint

`PresentationDocument` and `SlideTable`.

## `PresentationDocument`

*Based on* `OoxDocument`. The common methods are under [Documents](documents.md).

Base class.  Subclasses set `MAIN_CONTENT_TYPES` and implement `new`.

### Deck, slides and layouts

#### `new`

```python
new(
    width: float = Emu(12192000),
    height: float = Emu(6858000),
    kind: str = '.pptx',
    layouts: str = 'standard',
) -> PresentationDocument
```

*classmethod*

New deck (default 13.333 x 7.5 in, i.e. 16:9; size in points or
`Length`) with master, theme and
Office's eleven standard layouts (Title Slide, Title and Content, Section
Header, Two Content, Comparison, Title Only, Blank, Content with Caption,
Picture with Caption, Title and Vertical Text, Vertical Title and Text);
`layouts="basic"` gives only Title Slide, Title and Content, Blank.

#### `presentation`

*read-only property*

The root `p:presentation` element of the main part.

#### `slides`

*read-only property*

Slide roots (`p:sld`) in presentation order.

#### `add_slide`

```python
add_slide(layout: Part | str | None = None) -> OoxElement
```

Append a slide based on `layout` (a layout part or its name).

The default is the layout named "Blank" if there is one, else the
first.  The layout's placeholders are copied onto the slide (empty,
inheriting position and style), except date/footer/slide number --
the same as PowerPoint's "New Slide".

#### `masters`

*read-only property*

The slide master parts.

#### `layouts`

*read-only property*

Layout parts in the masters' own order (`p:sldLayoutIdLst`).

#### `layout_names`

*read-only property*

Names of the slide layouts, in master order.

#### `layout`

```python
layout(name: str) -> Part
```

The slide layout part named `name` ("Title and Content"); KeyError lists the
available names.

#### `add_layout`

```python
add_layout(layout: OoxElement, master: Part | None = None) -> Part
```

Add a `p:sldLayout` under `master` (default: the first master).

#### `slide_layout`

```python
slide_layout(slide: OoxElement) -> str
```

Name of the layout a slide is based on.

#### `slide_size`

*read-only property*

(width, height) of the slides in points.

#### `content_area`

```python
content_area(slide: OoxElement) -> tuple[float, float, float, float]
```

(x, y, width, height) in points of the slide's content area: the body
placeholder of its master (where "Title and Content" puts its text).
Shapes, tables, charts and pictures go there when no position is given.

### Placeholders

#### `placeholders`

```python
placeholders(slide: OoxElement) -> list[Placeholder]
```

A `Placeholder` (type, idx, name, element) for each placeholder on a
slide (or layout).

#### `placeholder`

```python
placeholder(slide: OoxElement, which: str | int) -> OoxElement
```

Find a placeholder by idx (int), by type ("ctrTitle", "obj", ...) or by
the aliases "title" (title/ctrTitle), "subtitle", "body" (body/obj).

#### `set_placeholder_text`

```python
set_placeholder_text(
    slide: OoxElement,
    which: str | int,
    text: str | list,
    *,
    lang: str = 'en-US',
) -> OoxElement
```

Replace a placeholder's text.  `text` is a string (`\n` splits
paragraphs) or a list of strings / `(level, string)` for indented
bullets (level 0 = top).  Formatting comes from the layout and master.

#### `set_placeholder_picture`

```python
set_placeholder_picture(
    slide: OoxElement,
    which: str | int,
    image,
    *,
    crop: bool = True,
    descr: str = '',
) -> OoxElement
```

Fill a picture (or content) placeholder, cropped to fill like PowerPoint.

### Text boxes, shapes and pictures

#### `add_text_box`

```python
add_text_box(
    slide: OoxElement,
    text: str,
    *,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    size: float | None = None,
    bold: bool = False,
    color: str | None = None,
    align: str | None = None,
    lang: str = 'en-US',
) -> OoxElement
```

Add a text box; `
` starts a new paragraph.  Position and size in

```python
    points (or `Length`), by default the content area's top, as wide as
    it and as tall as the text; `size` in points, `color` as hex RGB,
    `align` one of left/center/right/justify.
```

#### `add_shape`

```python
add_shape(
    slide: OoxElement,
    shape: str,
    *,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    **options,
) -> OoxElement
```

An auto shape with fill, line and text; position and size in points
(or `Length`), by default the content area.  See `oox.slides.add_shape`.

**Options** (keywords)

```python
text: str | None = None
fill: str | None = '4472C4'
line: str | None = None
line_width: float = 1.0
size: float | None = None
bold: bool = False
color: str | None = 'FFFFFF'
align: str | None = 'center'
name: str | None = None
```

An auto shape (`shape` = preset geometry: rect, roundRect, ellipse,
rightArrow, chevron, ...).  `fill`/`line` as hex RGB or `None`;
`line_width` in points; text centred and white by default.  `box` is
(x, y, cx, cy) in EMU (`deck.add_shape` takes points).

#### `add_picture`

```python
add_picture(
    slide: OoxElement,
    image,
    *,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    descr: str = '',
    fallback=None,
) -> OoxElement
```

Place a picture on a slide.  Position and size in points (or
`Length`); by default at the content area's top left, at the image's
natural size; one dimension keeps the aspect ratio.  An SVG needs
`fallback=` (a PNG) for Office before 2016.

#### `shapes`

```python
shapes(slide: OoxElement) -> list[Shape]
```

A `Shape` for each thing on the slide, with position and size in points;
see `oox.slides.shapes`.

**Details** (from `oox.slides`)

Slide content beyond text boxes: shapes with fills, tables, speaker notes,
picture placeholders, and a reader for what a slide holds.

```python
deck.add_shape(slide, "roundRect", x=72, y=144, width=216, height=72,
               text="Step 1", fill="2F5496")
deck.add_table(slide, [["Region", "Q1"], ["North", 120]], y=Inches(3.5),
               width=Inches(6), style="Medium Style 2 - Accent 1")
deck.set_notes(slide, "Mention the North region first.")
deck.set_placeholder_picture(slide, "pic", "photo.jpg")   # crops to fill, like PowerPoint
deck.shapes(slide)    # [Shape(kind="table", name=..., text=...), ...]
```

Tables use PowerPoint's built-in table styles, referenced by GUID
(`TABLE_STYLES` names a few; any GUID works).  Notes need a notes master,
which is created on first use (decks from PowerPoint usually have one).

### Tables

#### `add_table`

```python
add_table(
    slide: OoxElement,
    rows,
    *,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    **options,
) -> OoxElement
```

A table in a built-in PowerPoint style; position and size in points
(or `Length`): by default at the content area's top left, as wide as
it, rows of `row_height`.  See `oox.slides.add_table`.

**Options** (keywords)

```python
header: bool = True
banded: bool = True
first_column: bool = False
style: str | None = 'Medium Style 2 - Accent 1'
widths: list[float] | None = None
row_height: float = Emu(365760)
size: float | None = None
merges=()
name: str | None = None
```

A table from a list of rows (values become text; `None` is empty).

`style` is a name from `TABLE_STYLES` or a GUID; `header`,
`banded`, `first_column` switch the style's options.  `merges` are
`"A1:B1"`-style ranges (columns A.., rows 1..).  `box` is (x, y, cx,
cy) in EMU, cy None for rows x `row_height`; `widths` (per column) and
`row_height` in points or `Length`.  Returns the `p:graphicFrame`.

#### `tables`

```python
tables(slide: OoxElement) -> list[SlideTableInfo]
```

A `SlideTableInfo` (name, rows, merges, style ...) for each table on the
slide (`deck.table(slide, i)` edits one).

#### `table`

```python
table(slide: OoxElement, which: int | str = 0) -> SlideTable
```

A table on the slide, to edit: by position among the slide's
tables (0-based) or by shape name; see `oox.slidetable`.

**Details** (from `oox.slidetable`)

Editing tables already on a slide (made by oox, PowerPoint or anyone else).

```python
t = deck.table(slide)                 # first table, or by index / shape name
t.values()                            # [["Region", "Q1"], ["North", "120"], ...]
t.set_text(1, 1, 125)                 # keeps the cell's font, size and alignment
t.add_row(["West", 80])               # formatted like the last row
t.delete_row(2)
t.add_column(["Q2", 130, 95], index=2)
t.merge(0, 0, 0, 1)
```

Rows and columns are 0-based, as in `oox.tables.Table` (Word).  Merged cells
follow DrawingML: the top-left cell carries `gridSpan`/`rowSpan`, the
cells it covers `hMerge`/`vMerge`; inserting a row or column inside a
merge is refused (unmerge first), deleting one shrinks the merge.

### Charts

#### `add_chart`

```python
add_chart(
    slide: OoxElement,
    kind: str,
    categories,
    series,
    *,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    name: str | None = None,
    **options,
) -> OoxElement
```

Place a chart on a slide; returns its `p:graphicFrame`.  Position
and size in points (or `Length`), by default the content area.

`series` is `{"North": [1, 2, 3], ...}` (or `oox.charts.Series`
objects); the data also goes into an embedded workbook ("Edit Data").
Options are those of `oox.charts.build_chart`.

#### `charts`

```python
charts(slide: OoxElement) -> list[Chart]
```

A `Chart` (`oox.records`) for each chart on the slide, in order.

### Notes, links and sections

#### `set_notes`

```python
set_notes(slide: OoxElement, text: str) -> OoxElement
```

Set a slide's speaker notes (creates the notes master if needed).

#### `notes`

```python
notes(slide: OoxElement) -> str | None
```

A slide's speaker notes, or None if it has none.

#### `set_link`

```python
set_link(
    slide: OoxElement,
    shape,
    target,
    *,
    text: str | None = None,
    tooltip: str | None = None,
) -> OoxElement
```

Link a shape (element or name), or `text` in it, to a URL, a slide,
or "next"/"previous"/"first"/"last"/"end"; `None` removes it.  See
`oox.slidelinks`.

**Details** (from `oox.slidelinks`)

Hyperlinks and click actions on slides.

```python
deck.set_link(slide, shape, "https://example.org/")             # click the shape
deck.set_link(slide, shape, "https://example.org/", text="docs")  # just that text
deck.set_link(slide, shape, other_slide)                        # jump to a slide
deck.set_link(slide, shape, "next")                             # next/previous/first/last/end
deck.set_link(slide, shape, None)                               # remove
deck.links(slide)   # [{"shape": "TextBox 2", "text": "docs", "kind": "url",
                    #   "target": "https://example.org/", "tooltip": None}, ...]
```

A whole-shape link is `a:hlinkClick` in the shape's `p:cNvPr`; a text link
is `a:hlinkClick` in the runs' `a:rPr` (PowerPoint colours and underlines
them from the theme).  URLs (`http`, `https`, `mailto`, `file`) are
external relationships, slide jumps relationships to the slide part with the
`ppaction://hlinksldjump` action, and show jumps `ppaction://hlinkshowjump`.

#### `links`

```python
links(slide: OoxElement) -> list[Hyperlink]
```

A `Hyperlink` for each click link on the slide; see `oox.slidelinks.links`.

#### `set_sections`

```python
set_sections(sections) -> None
```

Slide sections: `[(name, first slide), ...]` (slide elements or
1-based numbers) or `SlideSection` records as `sections()` reads them;
`[]` removes them.  See `oox.sections_pml`.

**Details** (from `oox.sections_pml`)

Slide sections (PowerPoint 2010's Home > Section), stored as `p14:sectionLst`
in the presentation's extension list.

```python
deck.set_sections([("Introduction", s1), ("Results", s4)])
deck.sections()          # [SlideSection("Introduction", [1, 2, 3]), ...]
deck.set_sections([])    # remove them
```

Each section starts at the given slide and runs to the next one; slides before
the first section go into a "Default Section", as PowerPoint does.  New slides
join the last section.  The list sits in `p:extLst`, which PowerPoint 2007
and other readers that don't know it skip.  Section ids are GUIDs derived
from the names and positions, so the same sections give the same bytes.

#### `sections`

```python
sections() -> list[SlideSection]
```

A `SlideSection` (name, slide numbers) per section.

## `SlideTable`

A table graphic frame on a slide (`p:graphicFrame` holding `a:tbl`).

### `add_column`

```python
add_column(
    values: list[Any] | None = None,
    *,
    index: int | None = None,
    width: float | None = None,
    like: int | None = None,
) -> None
```

Insert a column before `index` (default: at the end), formatted
like column `like` (default: the one left of it); `width` in points
or `Length` (default: that column's).  The frame grows by the width.

### `add_row`

```python
add_row(
    values: list[Any] | None = None,
    *,
    index: int | None = None,
    like: int | None = None,
) -> OoxElement
```

Insert a row before `index` (default: at the end), formatted like
row `like` (default: the row above, or the first row); returns it.

### `cell`

```python
cell(row: int, col: int) -> OoxElement
```

The `a:tc` shown at (row, col): a merge's top-left cell.

### `delete_column`

```python
delete_column(col: int) -> None
```

Remove a column; merges through it shrink.

### `delete_row`

```python
delete_row(row: int) -> None
```

Remove a row; merges through it shrink (a merge starting in it moves
down with its text).

### `merge`

```python
merge(row1: int, col1: int, row2: int, col2: int) -> OoxElement
```

Merge the block (inclusive, 0-based); the texts of the merged cells
join the top-left one, as PowerPoint does.  Returns that cell.

### `merges`

```python
merges() -> list[tuple[int, int, int, int]]
```

`(row1, col1, row2, col2)` of each merged block.

### `n_cols`

*read-only property*

Number of grid columns.

### `n_rows`

*read-only property*

Number of rows.

### `rows`

*read-only property*

The `a:tr` elements.

### `set_style`

```python
set_style(
    style: str | None = None,
    *,
    header: bool | None = None,
    banded: bool | None = None,
    first_column: bool | None = None,
) -> None
```

A built-in style by name (`oox.slides.TABLE_STYLES`) or GUID, and
its header row / banded rows / first column switches.

### `set_text`

```python
set_text(
    row: int,
    col: int,
    text: Any,
    *,
    size: float | None = None,
    bold: bool | None = None,
    color: str | None = None,
    align: str | None = None,
) -> OoxElement
```

Replace the text of the cell shown at (row, col); `\n` starts a
paragraph.  The first run's font, size and colour and the first
paragraph's alignment are kept unless given.

### `set_widths`

```python
set_widths(widths: list[float]) -> None
```

Column widths in points (or `Length`); the frame's width follows.

### `text`

```python
text(row: int, col: int) -> str
```

Text of the cell at (`row`, `col`); paragraphs are joined with `\n`.

### `unmerge`

```python
unmerge(row: int, col: int) -> None
```

Split the merged block shown at (row, col) back into cells.

### `values`

```python
values() -> list[list[str | None]]
```

Text per grid position; cells covered by a merge are `None`.

### `widths`

*read-only property*

Column widths in points.
