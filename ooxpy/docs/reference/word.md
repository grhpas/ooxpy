# Word

`WordprocessingDocument` and the objects its methods return.

## `WordprocessingDocument`

*Based on* `OoxDocument`. The common methods are under [Documents](documents.md).

A .docx package.

```pycon
>>> from oox import WordprocessingDocument
>>> from oox.ns import w
>>> doc = WordprocessingDocument.new()
>>> doc.body.add(w.P(w.R(w.T("Hello"))))
<OoxElement w:body ...>
```

### Creating and structure

#### `new`

```python
new(
    kind: str = '.docx',
    *,
    minimal: bool = False,
    page: str = 'A4',
    landscape: bool = False,
    lang: str = 'en-US',
) -> WordprocessingDocument
```

*classmethod*

A new document, set up the way Word sets one up.

By default it gets styles (Normal, Title, Heading 1-3, List
Paragraph, Hyperlink, Table Grid), settings, a font table, the
Office theme and a section with `page` size ("A4" or "Letter")
and 1-inch margins.  `minimal=True` gives just `word/document.xml`
(smallest valid package; Word fills in its own defaults).

#### `body`

*read-only property*

The `w:body` element, where paragraphs and tables go.

#### `text`

*read-only property*

The `w:body` element, where paragraphs and tables go.

#### `document`

*read-only property*

The root `w:document` element of the main part.

#### `settings`

*read-only property*

Root `<w:settings>`, created on first use.

#### `styles`

*read-only property*

Root `<w:styles>` of the styles part, created on first use.

odfpy: `doc.styles`.  OOXML has no "automatic styles"; direct
formatting lives in `w:pPr`/`w:rPr` instead.

### Text and styles

#### `add_heading`

```python
add_heading(text: str, level: int = 1) -> OoxElement
```

Heading paragraph: level 0 is "Title", 1-9 "Heading N".

#### `add_paragraph`

```python
add_paragraph(text: str = '', style: str | None = None) -> OoxElement
```

Append a paragraph (before the final section properties); returns it.
`text` may contain \t and \n (tab, line break).

#### `add_run`

```python
add_run(
    paragraph: OoxElement,
    text: str,
    *,
    style: str | None = None,
    **fmt,
) -> OoxElement
```

Append a run with formatting (bold, italic, underline, strike, size,
font, color, highlight, superscript, subscript, caps, small_caps);
`style` is a character style (name or id).

#### `format_paragraph`

```python
format_paragraph(paragraph: OoxElement, **fmt) -> OoxElement
```

Alignment, spacing (pt), indents (EMU or twips), line spacing, keeps;
see `oox.formatting.format_paragraph`.

Paragraph formatting: align (left/center/right/justify), space_before /
space_after (pt), line_spacing (multiple or "12pt"), indent_left /
indent_right / first_line (EMU or twips; negative first_line = hanging),
keep_with_next, keep_together, page_break_before.

#### `add_style`

```python
add_style(
    name: str,
    *,
    kind: str = 'paragraph',
    base: str | None = None,
    next_style: str | None = None,
    quick: bool = True,
    **fmt,
) -> str
```

A custom paragraph or character style; returns its id.

Add a custom style and return its id.  `fmt` takes the run and
paragraph options (paragraph styles may use both).

#### `style_id`

```python
style_id(style: str) -> str
```

Style id for a style given by id ("Heading1") or name ("Heading 1",
"heading 1"; Word stores built-in names in lower case).

#### `ensure_style`

```python
ensure_style(style_id: str, text_width: int = 9026) -> str
```

Make sure a style exists, adding Word's definition for the known
feature styles (Header, FootnoteText, TOC1, Hyperlink ...).

#### `effective_format`

```python
effective_format(paragraph: OoxElement, run: OoxElement | None = None) -> dict
```

The formatting Word applies to `run` (or the paragraph mark),
resolved through defaults, styles and direct formatting.

#### `outline_level`

```python
outline_level(paragraph: OoxElement) -> int
```

1-9 for headings (direct w:outlineLvl or via the style chain), 10 for body text.

#### `add_page_break`

```python
add_page_break() -> OoxElement
```

A paragraph holding a page break, at the end of the body.

### Lists

#### `new_list`

```python
new_list(kind: str = 'bullet', start: int = 1)
```

A new list: "bullet", "decimal" (1. a. i.) or "legal" (1. 1.1. 1.1.1.).
Returns a `ListHandle`; `handle.add(text, level)` appends items.

#### `list_labels`

```python
list_labels() -> list[tuple[OoxElement, str]]
```

`(paragraph, label)` for every list paragraph, as Word numbers them.

#### `numbering`

*read-only property*

Root `<w:numbering>` of the numbering part, created on first use.

### Sections, headers and footers

#### `sections`

*read-only property*

`Section` objects in order (page setup, headers, footers).

#### `add_section`

```python
add_section(
    start: str = 'nextPage',
    *,
    landscape: bool | None = None,
    page: str | None = None,
)
```

End the current section and start a new one (default: on a new page).

#### `even_and_odd_headers`

*property*

Whether even and odd pages have different headers and footers (a document-wide
setting).

### Links, bookmarks, notes and comments

#### `add_hyperlink`

```python
add_hyperlink(
    paragraph: OoxElement,
    text: str,
    *,
    url: str | None = None,
    anchor: str | None = None,
    tooltip: str | None = None,
) -> OoxElement
```

A hyperlink with `text` at the end of `paragraph`: to `url` (an external
relationship) or to the bookmark `anchor`.  Returns the `w:hyperlink`.

#### `hyperlinks`

```python
hyperlinks() -> list[tuple[str, str]]
```

`(text, target)` for each hyperlink in the body; the target is the URL or
`#bookmark`.

#### `add_bookmark`

```python
add_bookmark(name: str, paragraph: OoxElement) -> OoxElement
```

Bookmark the whole content of `paragraph` as `name` (the target of
`add_hyperlink(..., anchor=name)` and cross-references); returns the bookmark
start.

#### `bookmarks`

```python
bookmarks() -> list[str]
```

Bookmark names, without Word's hidden ones (`_GoBack`, `_Toc...`).

#### `add_footnote`

```python
add_footnote(paragraph: OoxElement, text: str) -> OoxElement
```

A footnote with `text`, referenced at the end of `paragraph`; the footnotes
part is created on first use.  Returns the reference run.

#### `footnotes`

```python
footnotes() -> list[tuple[int, str]]
```

`(id, text)` for each footnote (Word's separator notes excluded).

#### `add_endnote`

```python
add_endnote(paragraph: OoxElement, text: str) -> OoxElement
```

An endnote with `text`, referenced at the end of `paragraph`; the endnotes
part is created on first use.  Returns the reference run.

#### `endnotes`

```python
endnotes() -> list[tuple[int, str]]
```

`(id, text)` for each endnote (Word's separator notes excluded).

#### `add_comment`

```python
add_comment(
    paragraph: OoxElement,
    text: str,
    *,
    author: str = 'oox',
    initials: str | None = None,
    date=None,
) -> OoxElement
```

Comment on the whole content of `paragraph`; `date` defaults to now.  Returns
the comment's `w:comment`.

#### `comments`

```python
comments() -> list[dict]
```

`{"id", "author", "text", "anchor"}` for each comment; `anchor` is the
commented text in the body.

### Tables

#### `add_table`

```python
add_table(
    rows: list[list],
    *,
    widths: list[int] | None = None,
    style: str | None = 'Table Grid',
    header_rows: int = 0,
    bold_header: bool = True,
    align: str | None = None,
    before: OoxElement | None = None,
)
```

A table from rows of values; returns a `Table` (merge, widths,
formatting, reading).  See `oox.tables`.

**Details** (from `oox.tables`)

Word tables: building, merged cells, widths, header rows, reading.

```python
t = doc.add_table([["Region", "Q1", "Q2"], ["North", 120, 135]],
                  header_rows=1, widths=[Cm(5), Cm(3), Cm(3)])
t.merge(2, 0, 2, 1)                 # cells (2,0)-(2,1) become one
t.format_cell(0, 0, fill="D9E2F3", align="center")
t.values()                          # [["Region", "Q1", "Q2"], ...]
```

Coordinates are 0-based *grid* positions, so they don't shift when cells are
merged: after `merge(2, 0, 2, 1)`, both `cell(2, 0)` and `cell(2, 1)`
return the merged cell.

How OOXML stores merges: a horizontal merge is one `w:tc` with
`w:gridSpan`; a vertical merge keeps a `w:tc` in every row, the first with
`w:vMerge val="restart"` and the ones below with `w:vMerge` (continue).
Widths are in twips (1/1440 inch); EMU values from `oox.units` are
converted when they are obviously EMU (> 100000).

Word doesn't count a row whose cells all continue vertical merges (Word 2007
reports 3 rows for such a 4-row table, and the row takes no height); LibreOffice
shows it as a thin row.  `Table.word_row_count` gives Word's count; avoid
merges that cover a whole row if the layout matters.

#### `tables`

```python
tables()
```

`Table` objects for the body's top-level tables.

### Pictures, charts and text boxes

#### `add_picture`

```python
add_picture(
    image,
    *,
    width: int | None = None,
    height: int | None = None,
    paragraph: OoxElement | None = None,
    descr: str = '',
    fallback=None,
) -> OoxElement
```

Insert an inline picture; returns the `w:r` holding it.

Size is in EMU (`oox.units`); give one dimension to keep the aspect
ratio, none for the image's natural size.  Without `paragraph` the
picture goes into a new paragraph at the end of the body.  An SVG
needs `fallback=` (a PNG) for Office before 2016.

#### `add_floating_picture`

```python
add_floating_picture(image, **options) -> OoxElement
```

A picture positioned on the page with text wrapping; see
`oox.floating.add_floating_picture` for the options.

**Options** (keywords)

```python
width: int | None = None
height: int | None = None
x: int = 0
y: int = 0
h_relative: str = 'column'
v_relative: str = 'paragraph'
h_align: str | None = None
v_align: str | None = None
wrap: str = 'square'
side: str = 'bothSides'
paragraph: OoxElement | None = None
descr: str = ''
margin: int = 114300
fallback: Any = None
```

Insert a floating picture; returns the `w:r` holding it.

`x`/`y` are offsets in EMU from the `h_relative`/`v_relative`
reference; `h_align` (left, center, right, inside, outside) and
`v_align` (top, center, bottom, inside, outside) replace them.
`margin` is the gap between picture and text on the left and right.

#### `pictures`

```python
pictures() -> list[dict]
```

Inline and floating pictures in the body, in order.

#### `add_chart`

```python
add_chart(
    kind: str,
    categories,
    series,
    *,
    width: int = 5400000,
    height: int = 2700000,
    paragraph: OoxElement | None = None,
    descr: str = '',
    **options,
) -> OoxElement
```

Insert an inline chart; returns the `w:r` holding it.

`series` is `{"North": [1, 2, 3], ...}` (or `oox.charts.Series`
objects); the data is also stored in an embedded workbook, so Word's
"Edit Data" works.  Size in EMU (default 15 x 7.5 cm).  Options are
those of `oox.charts.build_chart` (title, legend, labels, ...).

#### `charts`

```python
charts() -> list[dict]
```

Charts in the body, in order: `oox.charts.chart_info` plus "rid".

#### `add_text_box`

```python
add_text_box(content='', **options) -> OoxElement
```

A floating text box (DrawingML, with a VML copy for Word 2007);
returns its `w:r`.  See `oox.textboxes.add_text_box`.

**Options** (keywords)

```python
width: int = 2160000
height: int = 1080000
x: int = 0
y: int = 0
h_relative: str = 'column'
v_relative: str = 'paragraph'
h_align: str | None = None
v_align: str | None = None
wrap: str = 'square'
fill: str | None = 'FFFFFF'
line: str | None = '000000'
line_width: float = 0.75
paragraph: OoxElement | None = None
name: str | None = None
```

A floating text box; returns the `w:r` holding it.

`content` is text (`\n` starts a paragraph) or a list of `w:p`/
`w:tbl` elements (or `Table` objects); they are copied in.  `x`/`y`
are EMU offsets from `h_relative` (column, page, margin, character) and
`v_relative` (paragraph, page, margin, line); `h_align` (left, center,
right, inside, outside) and `v_align` (top, center, bottom, inside,
outside) replace them.  `fill`/`line` are hex colours or None,
`line_width` in points.  Without `paragraph` the box is anchored in a
new paragraph at the end of the body.

#### `text_boxes`

```python
text_boxes() -> list[dict]
```

Text boxes in the body, each once (`oox.textboxes`).

**Details** (from `oox.textboxes`)

Word text boxes, written the way Word 2010 and later write them.

```python
run = doc.add_text_box("Note\nSecond paragraph", width=Cm(6), height=Cm(3),
                       x=Cm(10), y=0, wrap="square", fill="FFF2CC")
doc.text_boxes()          # [{"text": "Note\nSecond paragraph", ...}]
doc.set_text_box_text(0, "New text")
```

Each text box is an `mc:AlternateContent` in a run: the choice is a
DrawingML shape (`wps:wsp`, Word 2010+) and the fallback the same box as
VML (`v:shape` with `v:textbox`), which is what Word 2007 reads.  Both
hold the text in a `w:txbxContent` (paragraphs, tables), so the readers
count each box once and the editors change both copies.

#### `set_text_box_text`

```python
set_text_box_text(index: int, content) -> None
```

Replace the content of text box `index` (in `text_boxes()` order) with text
(`\n` starts a paragraph) or a list of paragraphs, in both copies Word keeps.

### Fields, captions, contents and templates

#### `add_field`

```python
add_field(paragraph: OoxElement, instr: str, result: str = '') -> OoxElement
```

Any field (`"TITLE"`, `'DOCPROPERTY "Company"'`, `'DATE \@ "d MMMM yyyy"'`);
the result is computed now when oox can.

#### `fields`

```python
fields(all_stories: bool = False) -> list[dict]
```

`{"name", "code", "result"}` for the body's fields; `all_stories=True` adds
headers, footers, notes and comments.

#### `update_fields`

```python
update_fields(*, now=None) -> dict[str, int]
```

Recompute field results that don't need page layout (SEQ, REF,
document properties, dates ...); see `oox.fieldcalc`.

**Details** (from `oox.fieldcalc`)

Captions, cross-references, document-information fields, and computing
field results without Word.

```python
cap = doc.add_caption("Sales by region", label="Figure", after=picture_paragraph)
p = doc.add_paragraph("As ")
doc.add_cross_reference(p, cap)                  # "Figure 1"
doc.add_cross_reference(p, cap, what="page")     # page estimate
doc.add_field(p, "TITLE")                        # the document title
doc.update_fields()                              # renumber, refresh references
doc.fields()                                     # [{"name": "SEQ", "code": ..., "result": "1"}, ...]
```

Word stores a field's last result in the file and only recomputes it on F9 or
printing, so the file has to carry correct results.  `update_fields`
computes what doesn't need page layout:

* `SEQ` (caption numbers; `\r n` restart, `\c` repeat, `\h` hidden,
  `\* ARABIC|ROMAN|roman|ALPHABETIC|alphabetic`)
* `REF` (bookmark text; `\p` gives "above"/"below"), `PAGEREF` (an
  *estimate* from explicit page and section breaks, or from the page breaks
  Word recorded at its last layout)
* `TITLE`, `SUBJECT`, `AUTHOR`, `KEYWORDS`, `COMMENTS`,
  `LASTSAVEDBY`, `DOCPROPERTY` (core, extended and custom properties),
  `CREATEDATE`/`SAVEDATE`/`PRINTDATE` with a `\@` date picture, and
  `DATE`/`TIME` when `now=` is given
* case switches `\* Upper|Lower|Caps|FirstCap`

`PAGE`, `NUMPAGES`, `TOC` and other layout-dependent fields keep their
cached results.

#### `add_caption`

```python
add_caption(
    text: str,
    label: str = 'Figure',
    *,
    before: OoxElement | None = None,
    after: OoxElement | None = None,
    sep: str = ': ',
)
```

A numbered caption ("Figure 1: text"); returns a `Caption` to
cross-reference.  See `oox.fieldcalc`.

#### `add_cross_reference`

```python
add_cross_reference(
    paragraph: OoxElement,
    target,
    *,
    what: str = 'text',
    hyperlink: bool = True,
) -> OoxElement
```

A cross-reference at the end of `paragraph` to a caption (from `add_caption`),
a bookmark name or a paragraph such as a heading (which gets a hidden bookmark).
`what`: "text" ("Figure 1", the heading's text), "page" (estimated until Word
updates fields) or "position" ("above"/"below").

#### `add_toc`

```python
add_toc(
    levels: tuple[int, int] = (1, 3),
    title: str | None = 'Contents',
    *,
    before: OoxElement | None = None,
    update_on_open: bool = False,
) -> list[OoxElement]
```

Insert a table of contents (a TOC field pre-filled with the current
headings; Word adds page numbers on update).  `before` places it
before that body element; default: at the end.

#### `replace_text`

```python
replace_text(mapping: dict[str, str] | str, new: str | None = None) -> dict[str, int]
```

Replace text across runs in body, headers, footers, notes and comments:
`replace_text({"{{name}}": "Ada"})` or `replace_text("old", "new")`.

### Tracked changes

#### `suggest_insert`

```python
suggest_insert(
    paragraph: OoxElement,
    text: str,
    *,
    author: str,
    date=None,
    after: str | None = None,
) -> OoxElement
```

Insert `text` as a tracked change by `author`: at the end of `paragraph`, or
right after the first occurrence of `after`.

#### `suggest_delete`

```python
suggest_delete(
    paragraph: OoxElement,
    text: str,
    *,
    author: str,
    date=None,
    occurrence: int = 1,
) -> OoxElement
```

Delete the `occurrence`-th `text` in `paragraph` (across runs) as a tracked
change by `author`.

#### `suggest_replace`

```python
suggest_replace(
    paragraph: OoxElement,
    old: str,
    new: str,
    *,
    author: str,
    date=None,
    occurrence: int = 1,
)
```

Replace the `occurrence`-th `old` in `paragraph` with `new` as a tracked
deletion plus insertion; the new text keeps the old text's formatting.

#### `track_changes`

*property*

Word's "Track Changes" switch (`w:trackRevisions` in settings).

#### `revisions`

```python
revisions() -> list[dict]
```

Tracked changes in document order (body first, then other stories): `{"kind",
"author", "date", "text"}`; `kind` is insert, delete, move-from, move-to,
paragraph-insert, paragraph-delete, row-insert, row-delete or format.

#### `accept_all`

```python
accept_all() -> int
```

Accept every tracked change (body, headers, footers, notes, comments).

#### `reject_all`

```python
reject_all() -> int
```

Reject every tracked change in every story; returns how many there were.

### Forms (content controls)

#### `add_text_control`

```python
add_text_control(
    paragraph: OoxElement,
    *,
    tag: str,
    title: str | None = None,
    text: str | None = None,
    placeholder: str = 'Click here to enter text.',
    rich: bool = False,
    multiline: bool = False,
) -> OoxElement
```

A text (or rich text) content control; see `oox.controls`.

**Details** (from `oox.controls`)

Content controls (Word forms): text, rich text, drop-down, combo box,
date and check box; reading and filling them by tag.

```python
p = doc.add_paragraph("Name: ")
doc.add_text_control(p, tag="name", title="Name", placeholder="Your name")
doc.add_dropdown(p, tag="dept", items=["Sales", "Finance", "IT"])
doc.add_date_control(p, tag="start", fmt="d MMMM yyyy")
doc.add_checkbox(p, tag="agree", checked=False)

doc.fill_controls({"name": "Ada", "dept": "Finance",
                   "start": date(2026, 11, 14), "agree": True})
doc.content_controls()   # [{"tag": "name", "type": "text", "value": "Ada", ...}, ...]
```

Controls are inline (inside a paragraph).  Readers also see block-level ones
(controls holding whole paragraphs, as Word templates often use).  Check
boxes are an Office 2010 feature (`w14:checkbox`): Word 2007 shows them
as a control holding the ☐/☒ character, which `fill_controls` keeps in
step.

#### `add_dropdown`

```python
add_dropdown(
    paragraph: OoxElement,
    *,
    tag: str,
    items: list,
    value=None,
    title: str | None = None,
    combo: bool = False,
    placeholder: str = 'Choose an item.',
) -> OoxElement
```

A drop-down content control (`combo=True`: a combo box, which also takes typed
text).  `items` are display texts or `(display, value)` pairs.

#### `add_date_control`

```python
add_date_control(
    paragraph: OoxElement,
    *,
    tag: str,
    value=None,
    fmt: str = 'd MMMM yyyy',
    title: str | None = None,
    lang: str = 'en-US',
    placeholder: str = 'Click here to enter a date.',
) -> OoxElement
```

A date-picker content control at the end of `paragraph`; `fmt` is Word's date
picture (`dd/MM/yyyy` ...), `value` a date or None for the placeholder.

#### `add_checkbox`

```python
add_checkbox(
    paragraph: OoxElement,
    *,
    tag: str,
    checked: bool = False,
    title: str | None = None,
) -> OoxElement
```

A check-box content control at the end of `paragraph` (Word 2010+; Word 2007
shows the ☐/☒ character).

#### `content_controls`

```python
content_controls() -> list[dict]
```

Every content control (body, headers, footers, notes) in order: `{"tag", "title",
"type", "value", "placeholder", ...}`.  `value` is the text (None while the
placeholder shows), a bool for check boxes; drop-downs add `items`, dates add
`date`.

#### `fill_controls`

```python
fill_controls(values: dict, *, strict: bool = True) -> int
```

Fill content controls by tag; see `oox.controls.fill_controls`.

### Protection

#### `protect`

```python
protect(edit: str = 'readOnly', *, password: str | None = None, **options) -> None
```

Restrict editing (readOnly, comments, trackedChanges, forms),
optionally with a password; see `oox.protection.protect`.

**Options** (keywords)

```python
algorithm: str = 'SHA-1'
spin_count: int = 100000
salt: bytes | None = None
formatting: bool | None = None
```

Restrict editing to `edit` (readOnly, comments, trackedChanges,
forms), optionally with a password (hashed; `salt` for reproducible
files).  `formatting=True` also limits formatting to allowed styles.

#### `protection`

```python
protection() -> dict | None
```

The editing restriction: `{"edit", "enforced", "password", "algorithm",
"formatting"}`, or None.

#### `check_protection_password`

```python
check_protection_password(password: str) -> bool
```

Whether `password` lifts the editing restriction (True when there is no
password).  Checks Word 2007's and Word 2010's hash attributes.

#### `unprotect`

```python
unprotect() -> None
```

Remove the editing restriction (a password is not needed: protection is not
encryption).

## `Table`

### `add_row`

```python
add_row(values: list[Any] | None = None) -> OoxElement
```

Append a row (one cell per grid column) and return it.

### `cell`

```python
cell(row: int, col: int) -> OoxElement
```

The `w:tc` shown at grid position (row, col), 0-based.

### `format_cell`

```python
format_cell(
    row: int,
    col: int,
    *,
    fill: str | None = None,
    valign: str | None = None,
    align: str | None = None,
    bold: bool | None = None,
) -> OoxElement
```

Shading (hex colour), vertical alignment (top/center/bottom),
paragraph alignment (left/center/right/justify), bold.

### `header_rows`

*read-only property*

How many leading rows repeat as a header on each page.

### `merge`

```python
merge(row1: int, col1: int, row2: int, col2: int) -> OoxElement
```

Merge the rectangle (row1, col1)-(row2, col2), inclusive; returns the
merged cell.  Text of the merged cells is kept, as in Word.

### `n_cols`

*read-only property*

Number of grid columns.

### `n_rows`

*read-only property*

Number of rows.

### `rows`

*read-only property*

The `w:tr` elements.

### `set_header_rows`

```python
set_header_rows(n: int) -> None
```

Repeat the first `n` rows at the top of each page.

### `set_row_height`

```python
set_row_height(row: int, twips: int, rule: str = 'atLeast') -> None
```

Height of `row` in twips (or EMU from `oox.units`); `rule` "atLeast" or
"exact".

### `set_text`

```python
set_text(row: int, col: int, text: Any, *, bold: bool = False) -> OoxElement
```

Replace the cell's content with one paragraph of `text`.

### `set_widths`

```python
set_widths(widths: list[int]) -> None
```

Fixed column widths (twips or EMU); updates the grid, every cell and
the table width, and switches to a fixed layout.

### `text`

```python
text(row: int, col: int) -> str
```

Text of the cell shown at grid position (`row`, `col`); paragraphs are joined
with `\n`.

### `values`

```python
values() -> list[list[str | None]]
```

Text per grid position; positions covered by a merge (other than its
top-left) are `None`.

### `widths`

*read-only property*

Column widths from the table grid, in twips.

### `word_row_count`

*read-only property*

Rows as Word counts them: rows made only of vertical-merge
continuation cells don't count.

## `Section`

### `different_first_page`

*property*

Whether the section's first page has its own header and footer (assignable).

### `footer`

```python
footer(kind: str = 'default', create: bool = True) -> OoxElement | None
```

The section's footer of `kind` ("default", "first", "even") as a `w:ftr`
element, created if missing (`create=False`: None instead).

### `footer_text`

```python
footer_text(text: str, kind: str = 'default', align: str = 'left') -> OoxElement
```

Set the footer of `kind` to `text`; `{PAGE}`, `{NUMPAGES}` and
`{SECTIONPAGES}` become fields, `\t` aligns to the centre and right tab stops.

### `header`

```python
header(kind: str = 'default', create: bool = True) -> OoxElement | None
```

The `w:hdr` of this section (creating its own if needed).

### `header_text`

```python
header_text(text: str, kind: str = 'default', align: str = 'left') -> OoxElement
```

Replace the header with one paragraph of `text`; `{PAGE}`,
`{NUMPAGES}`, `{SECTIONPAGES}` become fields.  Returns the paragraph.

### `linked_to_previous`

```python
linked_to_previous(what: str = 'header', kind: str = 'default') -> bool
```

True if this section has no own header/footer of that kind and
shows the previous section's.

### `margins`

*read-only property*

Twips: top, right, bottom, left, header, footer, gutter.

### `orientation`

*property*

"portrait" or "landscape" (assignable; assigning swaps width and height).

### `page_height`

*read-only property*

Page height in twips.

### `page_width`

*read-only property*

Page width in twips.

### `set_margins`

```python
set_margins(**twips: int) -> None
```

Set margins in twips (or EMU from `oox.units`): `top`, `right`, `bottom`,
`left`, `header`, `footer`, `gutter`.

### `shown`

```python
shown(what: str = 'header', kind: str = 'default') -> OoxElement | None
```

The header/footer this section actually shows, following "Link to
Previous" back through earlier sections; None if there is none.

### `start`

*property*

How the section starts: nextPage (default), continuous, evenPage, oddPage.

### `text_width`

*read-only property*

Width between the margins in twips (the space a full-width table gets).

## `ListHandle`

One list instance (a `w:num`); `add` appends items to the body.

### `add`

```python
add(text: str = '', level: int = 0, style: str | None = 'List Paragraph') -> OoxElement
```

Append a list item to the end of the body; returns the paragraph.

### `num_pr`

```python
num_pr(level: int = 0) -> OoxElement
```

A `w:numPr` putting a paragraph in this list at `level` (0-8), for paragraphs
you build yourself.

### `paragraph`

```python
paragraph(
    text: str = '',
    level: int = 0,
    style: str | None = 'List Paragraph',
) -> OoxElement
```

A list paragraph, not yet placed (add it wherever you like).

### `restart`

```python
restart(start: int = 1) -> ListHandle
```

A new list instance with the same format, numbering from `start`.

## `Field`

```python
Field(
    code: str,
    kind: str,
    element: OoxElement,
    separate: OoxElement | None = None,
    end: OoxElement | None = None,
    result: str = '',
) -> None
```

Fields: `code`, `kind`, `element`, `separate`, `end`, `result`

### `args`

```python
args() -> list[str]
```

The field code's arguments and switches, tokenized as Word does (quotes
respected).

### `name`

*read-only property*

The field type in upper case ("SEQ", "REF", "PAGE" ...).

## `Caption`

```python
Caption(paragraph: OoxElement, bookmark: str, label: str) -> None
```

Fields: `paragraph`, `bookmark`, `label`

### `number`

*read-only property*

The caption's current number, as its SEQ field shows it.
