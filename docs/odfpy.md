# Coming from odfpy

ooxpy is to Office Open XML what [odfpy](https://github.com/eea/odfpy) is to
OpenDocument: one factory per element, a grammar to check against, `addElement`,
`teletype`, `load()`. odfpy programs port mechanically: you change the imports and the
element names, and the structure of the code stays. The odfpy names are a permanent,
tested part of the API; they won't be deprecated.

## Hello world, both ways

odfpy:

<!-- no-run -->
```python
from odf.opendocument import OpenDocumentText
from odf.text import P

textdoc = OpenDocumentText()
p = P(text="Hello World!")
textdoc.text.addElement(p)
textdoc.save("helloworld", True)          # helloworld.odt
```

oox:

```python
from oox import teletype
from oox.ns import w
from oox.odfcompat import OpenDocumentText

textdoc = OpenDocumentText()
p = w.P()
teletype.addTextToElement(p, "Hello World!")
textdoc.text.addElement(p)
textdoc.save("helloworld", True)          # helloworld.docx
```

The one real difference: Word text lives in a run inside the paragraph (`w:p` →
`w:r` → `w:t`), so a paragraph can't take text directly. `teletype.addTextToElement`
builds the run for you, as it handles tabs and line breaks in odfpy.

## Loading and reading

```python
from oox.odfcompat import load

doc = load("helloworld.docx")             # any of .docx, .xlsx, .pptx
for para in doc.getElementsByType(w.P):
    print(teletype.extractText(para))
```

## Concept map

| odfpy | oox |
|---|---|
| `odf.text`, `odf.table`, `odf.style` … | `oox.ns.w`, `oox.ns.x`, `oox.ns.p`, `oox.ns.a` … (one module per prefix) |
| `text.P(stylename=s, text="…")` | `w.P(w.PPr(w.PStyle(val=s)), w.R(w.T("…")))`, or `doc.add_paragraph("…", style=…)` |
| `grammar.py` (allowed children and attributes) | generated grammar tables, plus schema order and value checks |
| `addElement`, `addText`, `setAttribute`, DOM (`childNodes`, `appendChild` …) | the same names, plus `add()`, `add_text()`, `set_attr()` |
| `attrconverters.py` | `oox.attrconverters` |
| `OpenDocumentText()` / `OpenDocumentSpreadsheet()` / `OpenDocumentPresentation()` | the same in `oox.odfcompat`, or `WordprocessingDocument.new()` … |
| `addPicture(...)` | `addPicture(...)` returns an `r:id`; `add_picture(...)` places the picture too |
| `load(path)` | `oox.odfcompat.load(path)` or `oox.open_document(path)` |
| `teletype` | `oox.teletype` |
| `doc.styles`, `doc.automaticstyles` | `doc.styles`; OOXML has no automatic styles (direct formatting sits on the paragraph or run) |

## Where the formats differ

Some things can't look the same in any library, because the formats differ:

| ODF | OOXML in oox |
|---|---|
| nested `List` / `ListItem` | each paragraph points at a numbering definition: `lst = doc.new_list("decimal")`, `lst.add(text, level)` |
| notes inline in the paragraph | notes live in their own part: `doc.add_footnote(p, text)` |
| `A(href=…)` | a relationship plus `w:hyperlink`: `doc.add_hyperlink(p, text, url=…)` |
| page styles with header and footer | per section, with first/even variants: `doc.sections[0].header_text("… {PAGE} …")` |
| `numbercolumnsspanned` + covered cells | `table.merge(r1, c1, r2, c2)` in Word, `book.merge_cells("A1:C1")` in Excel |
| `number:number-style` data styles | number format codes: `book.cell_style(number_format="#,##0.00")` |

[odfpy compatibility details](ODFPY_COMPAT.md) has the complete tables: every
supported odfpy name and keyword, and every structural difference with the oox call
that replaces it.
