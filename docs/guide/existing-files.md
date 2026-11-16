# Existing files

ooxpy is built to edit files that other programs made: Word, Excel, PowerPoint,
LibreOffice, python-docx, openpyxl, XlsxWriter, python-pptx. This page covers opening
them, reading them, changing them without disturbing the rest, and reaching parts the
conveniences don't cover.

## Opening

```python
from oox import SpreadsheetDocument, open_document

SpreadsheetDocument.new("Stock").save("stock.xlsx")       # a stand-in input

# WordprocessingDocument, SpreadsheetDocument or
book = open_document("stock.xlsx")
# PresentationDocument, from the file's content
print(type(book).__name__)
```

`open_document` accepts a path or a binary file object, and recognises macro-enabled
files and templates (.docm, .xlsm, .xltx, .potx …). The format classes have their own
`open()` if you know what to expect. Strict-conformance files are read as well; they
are converted to the usual Transitional form in memory and written back as Strict.

## Reading

Every convenience has a reader, and the readers are written against what other
producers actually write (shared strings with rich text, LibreOffice's relationship
variants, charts without cached values, Word's hidden bookmarks …):

```python
book.append_row("Stock", ["Item", "Qty", "Unit price"])
book.append_row("Stock", ["Pens", 120, 0.8])
book.append_row("Stock", ["Paper", 40, 4.5])

for row in book.values("Stock"):
    print(row)
print(book.properties.as_dict())
```

| Word | Excel | PowerPoint |
|---|---|---|
| `teletype.extractText(doc.body)`, `doc.tables()`, `doc.list_labels()`, `doc.sections`, `doc.hyperlinks()`, `doc.bookmarks()`, `doc.footnotes()`, `doc.comments()`, `doc.revisions()`, `doc.fields()`, `doc.content_controls()`, `doc.pictures()`, `doc.text_boxes()`, `doc.charts()`, `doc.protection()` | `book.values()`, `book.cells()`, `book.value()`, `book.display_text()`, `book.cell_format()`, `book.merged_cells()`, `book.defined_names()`, `book.validations()`, `book.conditional_formats()`, `book.tables()`, `book.comments()`, `book.hyperlinks()`, `book.charts()`, `book.page_setup_info()`, `book.sheet_options_info()` | `deck.layout_names`, `deck.slide_layout()`, `deck.placeholders()`, `deck.shapes()`, `deck.tables()`, `deck.notes()`, `deck.links()`, `deck.sections()`, `deck.charts()` |

## Changing a file without disturbing it

```python
import zipfile

from oox.spreadsheet import Formula

book.save("stock.xlsx")
book = open_document("stock.xlsx")
book.set_value("Stock", "A5", "Total qty")
book.set_value("Stock", "B5", Formula("SUM(B2:B4)", 160))
book.save("stock-total.xlsx")

before, after = zipfile.ZipFile("stock.xlsx"), zipfile.ZipFile("stock-total.xlsx")
changed = [n for n in before.namelist() if before.read(n) != after.read(n)]
print(changed)
# ['xl/worksheets/sheet1.xml']: only the edited sheet
```

Parts are parsed only when you touch them. On saving, a part whose XML you didn't
change is written back exactly as it was read, even if you read it, and so are the
relationships and `[Content_Types].xml`. What oox doesn't model (macros, custom XML,
SmartArt, pivot tables, ActiveX controls …) travels along untouched.

## Untrusted files

Opening is bounded. `Limits` caps the number of parts, the size of each part, the total
size and the compression ratio, so a zip bomb fails fast instead of filling memory.
Duplicate zip entries and DTDs are refused, and encrypted or legacy binary files (.doc,
.xls, .ppt) are recognised. Every failure is a `PackageError`:

```python
from oox import Limits, PackageError, Package

with open("not-office.docx", "wb") as f:
    f.write(b"just some text")
try:
    open_document("not-office.docx")
except PackageError as e:
    print("refused:", e)

tight = Limits(max_parts=500, max_part_size=50_000_000)
package = Package.open("stock.xlsx", limits=tight)
print(len(package.parts))
```

`Limits.unlimited()` lifts the limits for large files you trust. Validation never runs
on opening: a file that breaks the schema still opens, and `validate()` tells you how.

## Parts and relationships

When you need a part the conveniences don't cover, the document gives you the package:

```python
from oox import WordprocessingDocument

doc = WordprocessingDocument.new()
doc.add_paragraph("Hello")
settings = doc.related_parts(doc.main_part, "DocumentSettingsPart")[0]
print(settings.partname, settings.element.tag_name)
# /word/settings.xml w:settings

for name, part in sorted(doc.package.parts.items())[:3]:
    print(name, part.content_type.split(".")[-1])
```

`add_part(part_type, source)` creates any part the format defines (custom XML, a
glossary, a VBA project …) with the right content type, relationship and folder,
named from the Open XML SDK's part table:

```python
from oox.ns import w

part, rel_id = doc.add_part("WordprocessingCommentsPart",
                            element=w.Comments())
print(part.partname, rel_id)                               # /word/comments.xml rId...
```

## Working with the tree

Every element is an lxml element, so XPath and the rest of lxml are available. The
`oox.grammar.URI` table maps prefixes to namespace URIs:

```python
from oox.grammar import URI

W = URI["w"]
for p in doc.body.iter(f"{{{W}}}p"):
    print(p.tag_name, len(p))
print(doc.body.xpath("count(.//w:t)", namespaces={"w": W}))
```

## Other Python libraries

Files written by python-docx, openpyxl, XlsxWriter and python-pptx open, read, edit and
save cleanly, and files oox writes open in those libraries. The test suite checks both
directions on feature-rich files (styles, tables, merges, charts, comments, links,
validation, pictures, notes).
