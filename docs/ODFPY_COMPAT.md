# odfpy compatibility

oox aims to make odfpy programs **port mechanically**. The odfpy names are a
**permanent**, tested part of the API (ADR-002); they aren't being deprecated. You change the imports
and the element names. The structure of the code, and the methods you call,
stay the same. `tests/test_odfpy_examples.py` holds four of odfpy's own
example scripts (helloworld, loadsave, manualpagebreak, passwd-as-odt),
translated line by line, and checks the output with python-docx (and, for the
page break, with LibreOffice rendering).

## Same name, same behavior

| Area | Supported |
|---|---|
| Documents | `OpenDocumentText()`, `OpenDocumentSpreadsheet()`, `OpenDocumentPresentation()`, `load(file)` (any format), `doc.text`, `doc.styles`, `doc.save(name, addsuffix=True)`, `doc.save("-")` (stdout), `doc.getElementsByType(factory)` |
| Factory keywords | `text=`, `attributes={...}`, `qattributes={(uri, local): v}`, `parent=element`, `check_grammar=False` |
| Attribute names | odfpy spellings (lowercase, no dashes: `firstline` → `w:firstLine`), plus the exact local name (`firstLine`), the prefixed form (`w_firstLine`) and `**{"w:firstLine": …}` |
| Style references | Pass the style element as the value: `w.PStyle(val=mystyle)` uses its `w:styleId`. The odfpy equivalent is `P(stylename=mystyle)`. |
| Grammar-checked building | `addElement(el, check_grammar=True)`, `addText`, `setAttribute`, `getAttribute`, `removeAttribute` |
| DOM | `childNodes` (including text nodes with `.data`), `parentNode`, `firstChild`, `lastChild`, `nextSibling`, `previousSibling`, `hasChildNodes`, `appendChild`, `insertBefore`, `removeChild`, `replaceChild`, `setAttrNS`, `getAttrNS`, `removeAttrNS`, `allowed_attributes()`, `getElementsByType`, `isInstanceOf`, `toXml(level, f)` |
| Identity | `qname` is `(namespace, local)`, `tagName` is `"w:p"`, `nodeType`, and `ELEMENT_NODE`/`TEXT_NODE` |
| Pictures | `addPicture(filename, mediatype, content)`, `addPictureFromFile`, `addPictureFromString` |
| teletype | `from oox import teletype`, then `extractText(el)` and `addTextToElement(el, s)` (`\t` → `w:tab`, `\n` → `w:br`, spaces preserved) |

As in odfpy, `addElement` checks the grammar. In oox it also inserts the child
in schema order. `appendChild` and `insertBefore` are raw DOM operations that
do neither, so `element.validate()` is there to check trees built that way.

## Different because the formats differ

These can't be made identical by any library. They're properties of OOXML.

| odfpy / ODF | oox / OOXML | Why |
|---|---|---|
| `P(text="Hello")` | `w.P(w.R(w.T("Hello")))`, or `teletype.addTextToElement(p, "Hello")` | Word text must sit in a run inside the paragraph; `w:p` can't hold text directly. |
| `p.addText("x")` on a paragraph | `teletype.addTextToElement(p, "x")` | Same reason: you get `InvalidChild: <w:p> cannot contain text`. |
| `P(stylename=s)` | `w.P(w.PPr(w.PStyle(val=s)))` | Paragraph properties are a child element in OOXML, not an attribute. |
| `textdoc.automaticstyles` | direct formatting (`w:pPr`/`w:rPr`), or `doc.styles` | OOXML has no automatic styles. |
| `Style(name=…, family=…)` | `w.Style(w.Name(val=…), type=…, styleId=…)` | The display name and the ID are separate; references use the ID. |
| `List(stylename=…)` with nested `ListItem(P(…))` | `lst = doc.new_list("decimal")`; `lst.add(text, level)` | OOXML has no list container: each paragraph points at a numbering definition (`w:numPr`), and the level is a number, not nesting. |
| `Note(noteclass="footnote")` with `NoteCitation` and `NoteBody` inline in the paragraph | `doc.add_footnote(p, text)`, `doc.add_endnote(p, text)` | Note text lives in a separate part (`footnotes.xml`); the paragraph only holds a reference by id. |
| `A(href=…)` | `doc.add_hyperlink(p, text, target=…)`; `anchor=` for a bookmark | The URL is an external relationship, referenced by `r:id`. |
| `Annotation(creator=…)` inline | `doc.add_comment(p, text, author=…)` | Comments are a separate part; the paragraph marks the range (`commentRangeStart`/`End`). |
| `TableOfContent` with a `TableOfContentSource` | `doc.add_toc(levels=(1, 3))` | A Word TOC is a field (`TOC \o "1-3"`) whose cached result is the entries. |
| Header and footer in the master page (`style.Header`, `PageNumber`) | `doc.sections[i].header_text("… {PAGE} …")` | Headers are separate parts referenced per section, with first/even variants; page numbers are fields. |
| `TableCell(numbercolumnsspanned=2, numberrowsspanned=2)` plus `CoveredTableCell()` | `table.merge(r1, c1, r2, c2)` | A horizontal merge is one `w:tc` with `w:gridSpan` (no covered cells); a vertical merge keeps a `w:tc` in each row, marked `w:vMerge`. |
| `draw:frame` with `anchortype="paragraph"` and wrap in its graphic style | `doc.add_floating_picture(..., wrap="square")` | Position and wrapping are elements of `wp:anchor`, not a style. |
| `text:tracked-changes` with `ChangedRegion`, plus change marks in the text | `doc.suggest_insert / suggest_delete / suggest_replace`, `doc.revisions()`, `accept_all()` | OOXML wraps changed runs in place (`w:ins`, `w:del`); there is no separate change list. |
| `text:sequence` with `text:sequence-ref` / `text:bookmark-ref` | `doc.add_caption(...)`, `doc.add_cross_reference(...)` | Word numbers captions with SEQ fields and refers to bookmarks with REF fields; results are cached in the file. |
| `text:user-defined` / `meta:user-defined` | `doc.custom_properties["Name"] = value` and `doc.add_field(p, 'DOCPROPERTY "Name"')` | Custom properties live in `docProps/custom.xml`, typed. |
| Form controls (`form:text`, `form:listbox` …) | `add_text_control`, `add_dropdown`, `add_date_control`, `add_checkbox`; `fill_controls` | Word forms are content controls (`w:sdt`) in the text, not a separate form layer. |
| `style:text-properties` in automatic styles | `doc.add_run(p, text, bold=True, color=...)`, or `doc.add_style(name, ...)` | Direct formatting sits on the run (`w:rPr`); named styles are separate. |
| `TableColumn(stylename=width)` | `w.TblGrid(w.GridCol(w=twips), …)` | Column widths live in the table grid, in twips; `doc.add_table(widths=)` takes points. |
| `numbercolumnsrepeated=n` | repeat the element | OOXML has no repeat attribute. |
| file suffix `.odt`/`.ods`/`.odp` | `.docx`/`.xlsx`/`.pptx` | `addsuffix=True` picks the right one. |
| `addPicture()` returns a manifest path for `draw:image href=` | returns an `r:id` for `a:blip r:embed=`; usually you'd call `add_picture()`, which also builds the drawing XML | OOXML references media through relationships, not paths. |
| `number:number-style` / `number:date-style` plus a cell style with `datastylename` | `book.cell_style(number_format="#,##0.00")` | Excel uses format codes in the stylesheet; a cell refers to a combined style by index. |
| `TableColumn(stylename=...)` with `style:column-width` | `book.set_column_width(sheet, "B", 12)` | Widths are in characters of the default font, stored in `x:cols`. |
| `TableCell(numbercolumnsspanned=..., numberrowsspanned=...)` + `CoveredTableCell` | `book.merge_cells(sheet, "A1:C1")` | Merges are listed per sheet (`x:mergeCells`); covered cells simply stay empty. |
| `table:named-range`, `table:content-validation` | `book.define_name(...)`, `book.add_validation(...)` | Names live in the workbook, validations in the sheet. |
| `table:database-range` | `book.add_table(sheet, "A1:D20", totals=...)` | An Excel table is its own part with named columns, a style and an optional totals row. |
| `calcext:conditional-formats` | `book.conditional_format(sheet, ref, rule, ...)` | Rules sit in the sheet; highlight styles are differential formats in the stylesheet. |
| `office:annotation` in a cell | `book.add_comment(sheet, ref, text, author=)` | The text is in a comments part, the note box in a legacy VML drawing. |
| `text:a` in a cell | `book.add_hyperlink(sheet, ref, url_or_#location)` | Links are listed per sheet; external targets are relationships. |
| `table:print-ranges`, `style:page-layout` | `book.set_page_setup(sheet, ...)` | Print area and titles are defined names (`_xlnm.Print_Area`); page settings sit in the sheet. |
| `table:protected`, `table:display="false"` | `book.protect_sheet(...)`, `book.set_sheet_options(sheet, hidden=True)` | |
| `doc.spreadsheet.addElement(Table(name="S"))` | `book.add_sheet("S")` | Each worksheet is its own part with a relationship, not an element in one content.xml. |
| `TableRow`/`TableCell` with `text=` | `book.append_row(sheet, [values])`, or `x.Row(x.C(x.Is(x.T(...)), r="A1", t="inlineStr"))` | Cells carry their A1 reference and value type explicitly. |
| `OpenDocumentSpreadsheet()` has no tables | starts with "Sheet1" | Excel refuses a workbook with no sheets. |
| `doc.presentation.addElement(Page(masterpagename=…))` | `deck.add_slide()` | Slides are parts tied to a layout and master through relationships. |

## Not yet (roadmap)
- `addObject` (embedded sub-documents) and `addThumbnail`.
- Excel cell styles and number formats (phase 4).
- `doc.meta` as odfpy's element tree: use `doc.properties.title`, `.creator`, … instead (done in phase 2).
- `addCDATA`: not meaningful in OOXML.

## Name clashes to know about
`text` and `parent` are odfpy factory keywords, but two OOXML elements also
have attributes with those local names. The keyword meaning wins. Reach those
attributes with the prefixed form (`x_parent=…`) or `**{"x:parent": …}`.
