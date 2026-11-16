# Getting started

## Install

```bash
pip install ooxpy
```

ooxpy needs Python 3.10 or newer and installs [lxml](https://lxml.de/). The package
is `ooxpy`; you import `oox`.

## A Word document

```python
from oox import WordprocessingDocument

# styles, theme, A4 page, like Word's blank document
doc = WordprocessingDocument.new()
doc.add_heading("Meeting notes", 1)
p = doc.add_paragraph("Present: Ada, Grace and ")
doc.add_run(p, "Charles", bold=True)
doc.add_run(p, ".")

items = doc.new_list("decimal")
items.add("Budget approved")
items.add("Launch moved to May", level=0)
items.add("Marketing to confirm the date", level=1)

doc.properties.title = "Meeting notes"
doc.save("notes.docx")
```

`new()` sets up the document the way Word does: Normal, Title and Heading styles,
the Office theme, a font table and an A4 page with 1-inch margins (`page="Letter"`
for US paper).

## An Excel workbook

```python
import datetime as dt

from oox import SpreadsheetDocument
from oox.spreadsheet import Formula

book = SpreadsheetDocument.new("Orders")
book.append_row("Orders", ["Date", "Customer", "Amount"], style={"bold": True})
book.append_row("Orders", [dt.date(2026, 11, 2), "Acme", 1250.0])
book.append_row("Orders", [dt.date(2026, 11, 9), "Globex", 980.5])
book.set_value("Orders", "B4", "Total")
# with its cached result
book.set_value("Orders", "C4", Formula("SUM(C2:C3)", 2230.5))
book.style_range("Orders", "C2:C4", number_format="#,##0.00")
book.autofit_columns("Orders")
book.save("orders.xlsx")
```

Dates get a date format automatically, and reading them back gives `datetime`
objects again. A `Formula` can carry its result, so readers that don't calculate
(and Excel's "save changes?" prompt) see a value.

## A PowerPoint deck

```python
from oox import PresentationDocument
from oox.units import Inches

deck = PresentationDocument.new()           # 16:9, Office's eleven standard layouts
cover = deck.add_slide("Title Slide")
deck.set_placeholder_text(cover, "title", "Project Falcon")
deck.set_placeholder_text(cover, "subtitle", "Status, November 2026")

plan = deck.add_slide("Title and Content")
deck.set_placeholder_text(plan, "title", "Plan")
deck.set_placeholder_text(plan, "body", ["Research", (1, "12 interviews"), "Prototype"])
deck.set_notes(plan, "The pilot moved two weeks.")

deck.add_picture(plan, "logo.png", Inches(11), Inches(0.3), width=Inches(1.5))
deck.save("falcon.pptx")
```

Placeholders are filled by type (`"title"`, `"body"`, `"subtitle"`, `"pic"`), by
name or by index, and keep the layout's formatting.

## Open, read and change a file

`open_document` picks the right class from the file's content, whoever made it:

```python
from oox import open_document

book = open_document("orders.xlsx")
print(book.sheet_names)                    # ['Orders']
print(book.values("Orders"))               # rows as Python values
print(book.value("Orders", "C4"))          # 2230.5

book.set_value("Orders", "D1", "Checked")
book.save("orders-checked.xlsx")
```

Everything you didn't change is written back byte for byte, so editing a file made by
Excel or LibreOffice doesn't disturb the rest of it.

## Check a file

```python
doc = open_document("notes.docx")
print(doc.validate())                      # {} : nothing to report
print(doc.validate(target="Office2007"))   # also what Office 2007 couldn't read
```

From a shell:

```bash
oox validate notes.docx --target Office2007
oox text falcon.pptx
```

## Next

- [Concepts](concepts.md) explains elements, factories, validation and units, which
  every guide builds on.
- The guides go through each format: [Word](guide/word.md), [Excel](guide/excel.md),
  [PowerPoint](guide/powerpoint.md), [Charts](guide/charts.md).
