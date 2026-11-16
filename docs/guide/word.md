# Word

`WordprocessingDocument` creates and edits .docx files. This guide builds one document
section by section; every block runs as written (they share one document `doc`).
The [reference](../reference/word.md) lists every option.

## Creating a document

```python
from oox import WordprocessingDocument

# A4 portrait, like Word's blank document
doc = WordprocessingDocument.new()
letter = WordprocessingDocument.new(page="Letter", landscape=True, lang="en-GB")
bare = WordprocessingDocument.new(minimal=True)        # just word/document.xml
```

`new()` adds Word's defaults: the Normal, Title, Heading 1–3, List Paragraph,
Hyperlink and Table Grid styles, settings, a font table, the Office theme and a
section with 1-inch margins. Other built-in styles (Heading 4–9, Caption, TOC 1–9,
Footnote Text …) are added the first time something uses them, so a document carries
only the styles it needs.

## Paragraphs, runs and formatting

```python
doc.add_heading("Annual report", 0)                    # level 0 = Title
doc.add_heading("Summary", 1)                          # Heading 1
p = doc.add_paragraph("Revenue grew in every region ")
doc.add_run(p, "except the west", bold=True, color="C00000")
doc.add_run(p, ". Figures are in ")
doc.add_run(p, "EUR thousands", italic=True, font="Georgia", size=10)
doc.add_run(p, ".")
doc.format_paragraph(p, align="justify", space_after=12, line_spacing=1.15,
                     first_line=360)
```

`add_run` takes `bold`, `italic`, `underline`, `strike`, `size` (points), `font`,
`color`, `highlight`, `superscript`, `subscript`, `caps`, `small_caps` and a character
`style`. `format_paragraph` takes `align`, `space_before`/`space_after` (points),
`line_spacing` (a multiple, or `"12pt"` for exact), `indent_left`, `indent_right`,
`first_line` (twips or EMU; negative for a hanging indent), `keep_with_next`,
`keep_together` and `page_break_before`.

**Styles** are better than direct formatting for anything repeated:

```python
note = doc.add_style("Note", italic=True, color="595959", indent_left=720,
                     space_before=6)                    # returns the style id
n = doc.add_paragraph("All figures are unaudited.", style="Note")
fmt = doc.effective_format(n)                           # what Word will apply
print(fmt["italic"], fmt["color"], fmt["font"])          # True 595959 Calibri
```

`effective_format(paragraph, run=None)` resolves defaults, the style chain, character
styles, direct formatting and theme fonts, the way Word does.

## Lists

```python
# or "bullet", "legal" (1. 1.1. 1.1.1.)
steps = doc.new_list("decimal")
steps.add("Collect the figures")
steps.add("Check the North region", level=1)
steps.add("Check the South region", level=1)
steps.add("Publish")
# same format, numbering from 1 again
again = steps.restart()
again.add("A new numbered list")

print([label for _, label in doc.list_labels()])
# ['1.', 'a.', 'b.', '2.', '1.']
```

`list_labels()` computes the labels Word displays for every list paragraph in the
document, including lists that come from styles and from other producers' files.

## Sections, headers and footers

```python
first = doc.sections[0]
# no header on the cover page
first.different_first_page = True
# tabs: centre and right stops
first.header_text("Example Ltd\t\tConfidential")
first.footer_text("Page {PAGE} of {NUMPAGES}", align="center")
first.set_margins(left=1800, right=1800)                  # twips (or EMU)

# a landscape section for wide material
wide = doc.add_section(landscape=True)
doc.add_heading("Appendix", 1)
wide.header_text("Appendix", align="right")
print(wide.orientation, wide.text_width > first.text_width)   # landscape True
```

A new section shares the previous section's headers and footers until you write its
own, as in Word. `{PAGE}`, `{NUMPAGES}` and `{SECTIONPAGES}` become fields. For full
control, `section.header()` returns the header part's root element.

## Links, bookmarks, notes and comments

```python
p = doc.add_paragraph("The method follows ")
doc.add_hyperlink(p, "ECMA-376", url="https://ecma-international.org/")
doc.add_run(p, "; details are in ")
doc.add_hyperlink(p, "the appendix", anchor="appendix")
doc.add_run(p, ".")
doc.add_bookmark("appendix", doc.add_paragraph("Appendix text."))

doc.add_footnote(p, "Fourth edition, 2012.")
doc.add_endnote(p, "See the sources list.")
doc.add_comment(p, "Check the edition with legal.", author="Ada Lovelace")

print(doc.hyperlinks())
# [('ECMA-376', 'https://ecma-international.org/'), ('the appendix', '#appendix')]
print(doc.footnotes())          # [(1, 'Fourth edition, 2012.')]
print(doc.comments()[0]["author"])
```

## Tables

```python
from oox.units import Cm

table = doc.add_table([["Region", "Product", "Q1", "Q2"],
                       ["North", "Widgets", 120, 135],
                       [None, "Gadgets", 80, 95],
                       ["Total", None, 200, 230]],
                      header_rows=1, widths=[Cm(4), Cm(4), Cm(2.5), Cm(2.5)])
table.merge(1, 0, 2, 0)                     # "North" spans two rows
table.merge(3, 0, 3, 1)                     # "Total" spans two columns
for col in range(4):
    table.format_cell(0, col, fill="D9E2F3", align="center")
table.format_cell(3, 0, align="right", bold=True)
table.add_row(["Note", "Preliminary", None, None])

print(table.values()[0])                    # ['Region', 'Product', 'Q1', 'Q2']
print(table.n_rows, table.n_cols)           # 5 4
```

Coordinates are grid positions, so they don't shift when cells merge. `header_rows`
repeat on every page. `doc.tables()` returns a `Table` for each table in any
document, so the same methods read and edit tables made by Word.

## Pictures

```python
# inline, in a new paragraph
doc.add_picture("logo.png", width=Cm(3))
title = doc.add_heading("Sales by region", 1)
doc.add_floating_picture("logo.png", width=Cm(2), h_align="right", h_relative="margin",
                         paragraph=title, wrap="square", descr="Company logo")
print([pic["wrap"] for pic in doc.pictures()])             # ['inline', 'square']
```

Images can be file names, paths, bytes or file objects; PNG, JPEG, GIF, BMP, TIFF, EMF,
WMF and SVG (with a PNG `fallback=` for Office before 2016) are read without any
imaging library. Give one dimension to keep the aspect ratio. Wrapping can be
`square`, `tight`, `through`, `topAndBottom`, `front` or `behind`; positions are
offsets (`x`, `y` in EMU) or alignments (`h_align`, `v_align`) relative to the column,
page, margin, paragraph or line.

## Text boxes

```python
box = doc.add_text_box("Note\nPreliminary figures", width=Cm(5), height=Cm(2.5),
                       h_align="right", h_relative="margin", fill="FFF2CC", line="BF9000")
print(doc.text_boxes()[0]["text"].splitlines())
# ['Note', 'Preliminary figures']
doc.set_text_box_text(0, "Note\nFinal figures")
```

The content can also be a list of paragraphs and tables. Each box is written the way
Word 2010 and later write it, a DrawingML shape, with a VML copy that Word 2007
reads; `text_boxes()` counts each box once and `set_text_box_text` changes both copies.

## Charts

```python
doc.add_chart("column", ["Q1", "Q2", "Q3"], {"North": [120, 135, 150],
                                             "South": [95, 110, 99]},
              title="Sales", legend="bottom", labels="value")
```

The data is also stored in an embedded workbook, so Word's "Edit Data" works. See
[Charts](charts.md) for every kind and option.

## Captions, cross-references and fields

```python
figure = doc.add_caption("Sales by quarter", label="Figure")
# "Figure 1: Sales by quarter"
p = doc.add_paragraph("As ")
doc.add_cross_reference(p, figure)                                 # "Figure 1"
doc.add_run(p, " shows (")
doc.add_cross_reference(p, figure, what="position")                # "above"
doc.add_run(p, "), the trend holds.")

doc.properties.title = "Annual report"
# filled from the properties
doc.add_field(doc.add_paragraph(), "TITLE")
print([(f["name"], f["result"]) for f in doc.fields()][:3])
# [('SEQ', '1'), ('REF', 'Figure 1'), ('REF', 'above')]
```

Word recomputes fields only when asked, so oox computes their results itself: every
method that adds a field calls `doc.update_fields()`, which renumbers SEQ fields and
refreshes REF, document property and date fields. Results that need page layout
(PAGE, NUMPAGES, page cross-references) are left for Word.

## Table of contents and templates

```python
# pre-filled with the headings
doc.add_toc(levels=(1, 2), before=doc.body[0])

template = WordprocessingDocument.new()
template.add_paragraph("Dear {{name}}, your order {{order}} has shipped.")
template.sections[0].footer_text("{{company}}")
counts = template.replace_text({"{{name}}": "Ada", "{{order}}": "A-1042",
                                "{{company}}": "Example Ltd"})
print(counts)
# {'{{name}}': 1, '{{order}}': 1, '{{company}}': 1}
```

The TOC is a real TOC field filled with the current headings; Word adds page numbers
when it updates the field (`update_on_open=True` asks Word to do that on opening).
`replace_text` matches across run boundaries, the classic problem with Word
templates, in the body, headers, footers, notes and comments; the replacement takes the
formatting of the run where the match starts.

## Tracked changes

```python
import datetime as dt

p = doc.add_paragraph("Sales grew by 8% in 2025.")
when = dt.datetime(2026, 11, 14, 9, 0)
doc.suggest_replace(p, "8%", "9%", author="Finance", date=when)
doc.suggest_insert(p, " Figures are unaudited.", author="Finance", date=when)
# edits made later in Word are tracked too
doc.track_changes = True

print([(r["kind"], r["text"]) for r in doc.revisions()])
# [('delete', '8%'), ('insert', '9%'), ('insert', ' Figures are unaudited.')]
```

`accept_all()` and `reject_all()` resolve every change, including Word's moves,
paragraph marks, table rows and formatting changes, in every story.

## Forms

```python
form = WordprocessingDocument.new()
form.add_heading("Leave request", 1)
form.add_text_control(form.add_paragraph("Name: "), tag="name", placeholder="Your name")
form.add_dropdown(form.add_paragraph("Type: "), tag="type",
                  items=["Annual leave", "Sick leave", "Unpaid"])
form.add_date_control(form.add_paragraph("From: "), tag="from", fmt="d MMMM yyyy")
form.add_checkbox(form.add_paragraph("Manager informed: "), tag="informed")
form.save("leave-form.docx")

form.fill_controls({"name": "Ada Lovelace", "type": "Annual leave",
                    "from": dt.date(2026, 12, 21), "informed": True})
print({c["tag"]: c["value"] for c in form.content_controls()})
# {'name': 'Ada Lovelace', 'type': 'Annual leave', 'from': '21 December 2026',
#  'informed': True}
```

Content controls are Word's form fields. `fill_controls` fills them by tag (unknown
tags raise unless `strict=False`), so a template made in Word can be filled from data.

## Protection

```python
# or "readOnly", "comments", "trackedChanges"
form.protect("forms", password="secret")
print(form.protection()["edit"])              # forms
print(form.check_protection_password("secret"))
form.unprotect()
```

This is Word's Restrict Editing: the password is stored as Word's salted hash (SHA-1 by
default, which every Word version checks; `algorithm="SHA-512"` is also understood by
Word 2007). It restricts editing in Word; it doesn't encrypt the file.

## Properties

```python
doc.properties.title = "Annual report"
doc.properties.creator = "Finance"
doc.extended_properties.company = "Example Ltd"
doc.custom_properties["Quarter"] = "Q3"
doc.custom_properties["Approved"] = True
print(doc.properties.as_dict()["title"], doc.custom_properties.as_dict())
```

## Saving

```python
problems = doc.validate(target="Office2007")
assert not problems, problems
doc.save("annual-report.docx")
```

On save, oox applies the rules a single element can't enforce: a body or table cell
that ends with a table gets a closing paragraph, as Word requires.
