# Cookbook

Complete recipes for common jobs. Each one runs as written (they share a working
directory, so later recipes use files from earlier ones).

## Letters from a template and a list of people

A Word template with `{{placeholders}}`, filled once per record. `replace_text` matches
across runs, so placeholders can be formatted and still be found.

```python
from pathlib import Path

from oox import WordprocessingDocument, open_document

template = WordprocessingDocument.new()
template.add_paragraph("{{date}}")
greeting = template.add_paragraph("Dear ")
template.add_run(greeting, "{{name}}", bold=True)
template.add_run(greeting, ",")
template.add_paragraph("Your order {{order}} has shipped and should arrive by {{eta}}.")
template.sections[0].footer_text("Example Ltd · {{city}} office")
template.save("letter-template.docx")

people = [
    {"name": "Ada Lovelace", "order": "A-1042", "eta": "18 November", "city": "London"},
    {"name": "Grace Hopper", "order": "A-1043", "eta": "20 November", "city": "New York"},
]
out = Path("letters")
out.mkdir(exist_ok=True)
for person in people:
    letter = open_document("letter-template.docx")
    letter.replace_text({"{{date}}": "16 November 2026",
                         **{"{{%s}}" % k: v for k, v in person.items()}})
    letter.save(out / f"letter-{person['order']}.docx")
print(sorted(p.name for p in out.iterdir()))
```

## A Word report from data

```python
from oox import WordprocessingDocument
from oox.units import Cm

sales = {"North": [120, 135, 150], "South": [95, 110, 99], "West": [80, 84, 91]}
quarters = ["Q1", "Q2", "Q3"]

doc = WordprocessingDocument.new()
doc.properties.title = "Sales report"
first = doc.add_heading("Sales report", 0)
doc.add_heading("Figures", 1)
rows = [["Region", *quarters, "Total"]]
rows += [[region, *values, sum(values)] for region, values in sales.items()]
table = doc.add_table(rows, header_rows=1, widths=[Cm(4)] + [Cm(2.5)] * 4)
for col in range(5):
    table.format_cell(0, col, fill="D9E2F3", bold=True)
caption = doc.add_caption("Sales by region and quarter", label="Table",
                          after=table.element)

doc.add_heading("Trend", 1)
doc.add_chart("line", quarters, sales, title="Sales", y_title="EUR k", legend="bottom")
p = doc.add_paragraph("Details are in ")
doc.add_cross_reference(p, caption)
doc.add_run(p, ".")
doc.add_toc(levels=(1, 1), before=doc.body[1])
doc.save("sales-report.docx")
```

## Rows of data to a formatted Excel sheet

```python
import datetime as dt

from oox import SpreadsheetDocument

orders = [
    {"date": dt.date(2026, 11, 2), "customer": "Acme", "amount": 1250.0, "paid": True},
    {"date": dt.date(2026, 11, 9), "customer": "Globex", "amount": 980.5, "paid": False},
    {"date": dt.date(2026, 11, 13), "customer": "Initech", "amount": -120.0, "paid": True},
]
book = SpreadsheetDocument.new("Orders")
book.append_row("Orders", ["Date", "Customer", "Amount", "Paid"])
for o in orders:
    book.append_row("Orders", [o["date"], o["customer"], o["amount"], o["paid"]])
last = len(orders) + 1
book.add_table("Orders", f"A1:D{last}", name="Orders",
               totals={"Date": "Total", "Amount": "sum"})
book.style_range("Orders", f"C2:C{last + 1}", number_format="#,##0.00;[Red]-#,##0.00")
book.freeze_panes("Orders", "A2")
book.autofit_columns("Orders")
book.save("orders.xlsx")
print(book.display_text("Orders", f"C{last + 1}"))      # 2,110.50
```

## An Excel sheet into a list of dicts

```python
from oox import open_document

book = open_document("orders.xlsx")
header, *rows = book.values("Orders")
records = [dict(zip(header, row)) for row in rows if row[0] != "Total"]
print(records[0])
# {'Date': datetime.datetime(2026, 11, 2, 0, 0), 'Customer': 'Acme', 'Amount': 1250.0,
#  'Paid': True}
```

## One slide per item

```python
from oox import PresentationDocument

projects = [("Falcon", "On track", ["Pilot in North", "Launch in May"]),
            ("Heron", "At risk", ["Supplier delayed", "Decision on 1 Dec"])]
deck = PresentationDocument.new()
cover = deck.add_slide("Title Slide")
deck.set_placeholder_text(cover, "title", "Project status")
deck.set_placeholder_text(cover, "subtitle", "November 2026")
for name, status, points in projects:
    slide = deck.add_slide("Title and Content")
    deck.set_placeholder_text(slide, "title", f"{name}: {status}")
    deck.set_placeholder_text(slide, "body", points)
    deck.set_notes(slide, f"Owner of {name} to present.")
deck.save("status.pptx")
print(len(deck.slides))                                   # 3
```

## All the text in any Office file

```python
from oox import (PresentationDocument, SpreadsheetDocument, WordprocessingDocument,
                 open_document)
from oox import teletype


def text_of(path) -> str:
    doc = open_document(path)
    if isinstance(doc, WordprocessingDocument):
        return teletype.extractText(doc.body)
    if isinstance(doc, SpreadsheetDocument):
        return "\n".join("\t".join("" if v is None else str(v) for v in row)
                         for name in doc.sheet_names for row in doc.values(name))
    assert isinstance(doc, PresentationDocument)
    return "\n".join(s["text"] for slide in doc.slides for s in doc.shapes(slide)
                     if s["text"])


for name in ("sales-report.docx", "orders.xlsx", "status.pptx"):
    print(name, len(text_of(name)), "characters")
```

The `oox text FILE` command does the same from a shell.

## Find and replace in a folder of documents

```python
from pathlib import Path

from oox import open_document

changed = {}
for path in sorted(Path("letters").glob("*.docx")):
    doc = open_document(path)
    counts = doc.replace_text("Example Ltd", "Example Group")
    if any(counts.values()):
        doc.save(path)
        changed[path.name] = counts
print(changed)
```

Headers, footers, notes and comments are searched too, and untouched parts stay
byte-identical.

## Make a file safe for Office 2007

```python
from oox import open_document

doc = open_document("sales-report.docx")
findings = doc.validate(target="Office2007")
if findings:
    doc.save("sales-report-2007.docx", target="Office2007", strip=True)
    print(doc.stripped)
else:
    print("already fine for Office 2007")
```

`oox downgrade IN OUT` does the same from a shell.

## Accept every tracked change in a batch of files

```python
import datetime as dt

from oox import WordprocessingDocument, open_document, teletype

draft = WordprocessingDocument.new()
p = draft.add_paragraph("Revenue grew by 8%.")
draft.suggest_replace(p, "8%", "9%", author="Finance", date=dt.datetime(2026, 11, 14))
draft.save("draft.docx")

final = open_document("draft.docx")
n = final.accept_all()
final.save("final.docx")
print(n, teletype.extractText(final.body))           # 2 Revenue grew by 9%.
```
