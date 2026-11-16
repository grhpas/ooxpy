# ooxpy

**Office Open XML for Python, the odfpy way.** ooxpy (imported as `oox`) reads and
writes Word (.docx), Excel (.xlsx) and PowerPoint (.pptx) files. It has one factory
function per element of the format, generated from the schemas, and a runtime that
checks your markup as you build it: wrong attribute values, children in the wrong
place and elements a given Office version can't read are reported before the file
reaches anyone.

On top of the elements sit conveniences for the everyday work: headings, lists,
tables, pictures, charts, tracked changes and forms in Word; values, styles,
conditional formats, tables and charts in Excel; layouts, placeholders, tables and
notes in PowerPoint. Every convenience has a reader, so the same calls work on files
made by Office, LibreOffice or other libraries.

```python
from oox import WordprocessingDocument

doc = WordprocessingDocument.new()
doc.add_heading("Quarterly report", 1)
doc.add_paragraph("Revenue grew in every region.")
table = doc.add_table([["Region", "Q1", "Q2"], ["North", 120, 135], ["South", 95, 110]],
                      header_rows=1)
doc.add_chart("column", ["Q1", "Q2"], {"North": [120, 135], "South": [95, 110]},
              title="Revenue")
doc.save("report.docx")
```

## Why ooxpy

- **The whole vocabulary.** Every element of ECMA-376 and Microsoft's extensions
  (w14, x14, p14 …) has a factory: `oox.ns.w.P`, `oox.ns.x.C`, `oox.ns.a.SolidFill`.
  When a convenience doesn't exist, you can still build exactly the markup you need,
  and your IDE completes the attribute names and values.
- **Hard to get wrong.** Children are kept in schema order, values are checked, and
  `doc.validate(target="Office2007")` tells you what an older Office won't read.
- **Safe with existing files.** Parts you don't touch are written back byte for byte;
  untrusted files are opened with limits (zip bombs, DTDs, encrypted files give a clear
  error).
- **Checked against real Office.** The Word, Excel and PowerPoint features have cases
  in an Office check corpus that is run on Office 2007 and 2016
  ([Office compatibility](office.md)).
- **odfpy-compatible names.** `addElement`, `teletype.extractText`, `load()` and
  `OpenDocumentText()` work as in odfpy ([Coming from odfpy](odfpy.md)).

## Install

```bash
pip install ooxpy
```

Python 3.10 or newer; the only dependency is [lxml](https://lxml.de/).

## Where next

- [Getting started](getting-started.md): one file in each format, reading and editing.
- [Concepts](concepts.md): elements, factories, validation, units.
- Guides for [Word](guide/word.md), [Excel](guide/excel.md),
  [PowerPoint](guide/powerpoint.md) and [charts](guide/charts.md).
- [Cookbook](cookbook.md): complete recipes.
- [API reference](reference/index.md).
