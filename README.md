# ooxpy (`import oox`)

A schema-driven, **odfpy-style** library for Office Open XML (.docx, .xlsx, .pptx).
It has one factory function per element, generated from the Open XML SDK's
schema metadata, plus a runtime that validates your markup and keeps children
in schema order.

```python
from oox import WordprocessingDocument
from oox.ns import w

doc = WordprocessingDocument.new()
doc.body.add(
    w.P(w.PPr(w.Jc(val="center")),                 # order is fixed for you
        w.R(w.RPr(w.Sz(val=48), w.B()), w.T("Hello, OOXML"))),
)
doc.save("hello.docx")

w.Jc(val="middle")   # InvalidAttributeValue: w:val='middle' is not one of: left, start, center, right, end, both, ...
                     # (and your IDE offers those values: every factory ships with typed stubs)
w.P(w.Tbl())         # InvalidChild: <w:tbl> is not allowed in <w:p>; allowed: ...
```

Excel and PowerPoint work the same way:

```python
from oox import SpreadsheetDocument, PresentationDocument
from oox.spreadsheet import Formula
from oox.units import Inches, Cm

book = SpreadsheetDocument.new("Sales")
book.append_row("Sales", ["Region", "Units", "Price", "Revenue"])
book.append_row("Sales", ["North", 120, 9.5, Formula("B2*C2")])
book.add_picture("Sales", "chart.png", anchor="F2", width=Cm(12))
book.save("sales.xlsx")

deck = PresentationDocument.new()                      # master, layout and theme included
slide = deck.add_slide()
deck.add_text_box(slide, "Regional sales", x=Inches(1), y=Inches(0.5), width=Inches(11),
                  height=Inches(1), size=40)
deck.add_picture(slide, "chart.png", x=Inches(1), y=Inches(2), width=Inches(6))
deck.save("sales.pptx")
```

Opening untrusted files is bounded (zip bombs, DTDs, legacy/encrypted files all give a
clear `PackageError`), Strict-conformance files are read and written, and
`doc.validate(target="Office2007")` checks a file against an older Office version.

Pictures (PNG, JPEG, GIF, BMP, TIFF, EMF, WMF, SVG) work in all three formats, and so
do charts (`book.add_chart("Sales", "column", "A1:D5")`, `doc.add_chart(...)`,
`deck.add_chart(...)`). `doc.save(path, target="Office2007")` refuses markup Office 2007
can't read; `strip=True` removes it.

From the command line:

```bash
python -m oox validate report.docx --target Office2007   # problems per part; exit 1 if any
python -m oox info deck.pptx                             # kind, parts, sheets/slides, properties
python -m oox text book.xlsx                             # paragraphs / cells / slide texts and notes
python -m oox downgrade from-libreoffice.xlsx for-2007.xlsx
```

**Coming from odfpy?** The same method names work (`addElement`, `childNodes`,
`getElementsByType`, `teletype`, `load`, `save(name, True)`). See
[docs/odfpy.md](https://github.com/grhpas/ooxpy/blob/main/docs/odfpy.md).

## Install

```bash
pip install ooxpy
```

Python 3.10 or newer; the only dependency is [lxml](https://lxml.de/).

## Documentation

The documentation lives in [`docs/`](https://github.com/grhpas/ooxpy/blob/main/docs/index.md): getting started, concepts, guides
for Word, Excel, PowerPoint and charts, a cookbook, the command line, migrating from
odfpy, Office compatibility and the API reference. Every code example in it runs in the
test suite. To read it as a site:

```bash
pip install ".[docs]"
mkdocs serve
```

The [`examples/`](https://github.com/grhpas/ooxpy/tree/main/examples/) folder has ten runnable scripts, each with a short HTML
tutorial next to it (`examples/index.html`).

**Status:** alpha, on PyPI since 0.5.2: the API may still change before
1.0 ([docs/STABILITY.md](https://github.com/grhpas/ooxpy/blob/main/docs/STABILITY.md)); changes are listed in
[CHANGELOG.md](https://github.com/grhpas/ooxpy/blob/main/CHANGELOG.md). The design and decision record is
[docs/ARCHITECTURE.md](https://github.com/grhpas/ooxpy/blob/main/docs/ARCHITECTURE.md), the plan [docs/ROADMAP.md](https://github.com/grhpas/ooxpy/blob/main/docs/ROADMAP.md).

## Development

```bash
pip install -e ".[dev,docs]"
pytest                               # includes the docs' examples and a strict site build
python -m tools.apidocs              # after changing docstrings: regenerate docs/reference
python -m tools.api_surface --write  # after a deliberate public API change

# regenerate the grammar from a pinned Open XML SDK checkout
git clone --depth 1 --filter=blob:none --sparse https://github.com/dotnet/Open-XML-SDK.git ../sdk
git -C ../sdk sparse-checkout set data
python -m tools.gen --sdk-data ../sdk/data
OOX_SDK_DATA=../sdk/data pytest      # also checks committed code == regenerated code
```

Generated code in `src/oox/_generated` and `src/oox/ns` comes from Open XML SDK
data (© .NET Foundation and Contributors, MIT). Every source, library and
specification we use is credited in [docs/ATTRIBUTIONS.md](https://github.com/grhpas/ooxpy/blob/main/docs/ATTRIBUTIONS.md).
