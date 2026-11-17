# Reading a document as Word shows it

A Word file says more than Word shows, and less. More, because some content is
stored twice: a text box written by Word 2010 or later is a modern shape *and* a VML
copy for Word 2007, and other producers do the same with charts and equations.
Less, because much of what you see is not written down at all: a list number, the
font of a paragraph that only names a style, the page size of a document with no
section properties. oox's readers report what Word shows. This page explains how,
and where the limits are.

## The rules every reader follows

- **One copy of everything.** Where the file offers alternatives
  (`mc:AlternateContent`), readers look only at the branch Word uses: the first
  `mc:Choice` whose required namespaces Word understands, else the `mc:Fallback`.
  Text, fields, bookmarks, hyperlinks, captions, revisions, content controls,
  tables, charts and list labels are counted once.
- **Content in wrappers counts.** Paragraphs and tables inside content controls
  (`w:sdt`) and custom XML are part of the document; so are section breaks there.
- **Reading never writes.** Missing parts and settings are reported as Word's
  defaults instead of being created (a document without section properties reads
  as Letter with 1-inch margins). The `ensure_*` methods and setters are what
  write. Saving after reading gives the same bytes as saving without it.

`oox.walk` is the walker the readers share. Use it for your own traversals and they
will agree with oox's:

```python
from oox import WordprocessingDocument, teletype
from oox.ns import qn
from oox.walk import iter_blocks, iter_story

doc = WordprocessingDocument.new()
doc.add_heading("Quarterly report", 1)
doc.add_paragraph("Sales rose in every region.")
doc.add_text_box("Note: figures are unaudited.")

shown = list(iter_story(doc.body, qn("w:p")))        # paragraphs Word shows
stored = list(doc.body.iter(qn("w:p")))               # every copy in the file
print(len(shown), len(stored))                        # 4 5
print(teletype.extractText(doc.body).count("unaudited"))  # 1
blocks = list(iter_blocks(doc.body, qn("w:p"), qn("w:tbl")))  # top-level blocks
```

`iter_story(root, *tags, text_boxes=False)` leaves out text-box content, for the text
of a paragraph itself; `teletype.extractText(p, text_boxes=False)` does the same for
text.

## Formatting

`doc.formatting()` gives a resolver that answers "how does this look?" for
paragraphs and runs:

```python
from oox.formatting import LineSpacing

p = doc.add_paragraph("", style="Heading 2")
r = doc.add_run(p, "Results", italic=True)
fmt = doc.formatting()
para, run = fmt.paragraph(p), fmt.run(r)
print(para.style, para.keep_with_next, para.space_before)  # heading 2 True 2.0
print(run.font, run.size, run.italic, run.color)          # Calibri Light 13.0 True 2F5496
print(para.line_spacing)                                  # LineSpacing(rule='auto', value=1.0791666666666666)
```

It applies the layers Word applies, in Word's order:

1. document defaults;
2. the table style, for text in a table: its whole-table formatting, then the
   conditional formats that apply to the cell (banded columns, banded rows,
   first and last row, first and last column, corner cells), as the table's
   `tblLook` switches them on;
3. the paragraph style and the styles it is based on;
4. the list level's indents and tabs. When the numbering comes from the paragraph
   style, the level is applied before the style, so the style's own indents win;
5. the character style;
6. direct formatting.

Bold, italic, caps, small caps, strike, double strike and hidden text are *toggle*
properties. A character or table style that turns one on flips what is below it:
the Strong character style applied to text in a bold heading gives regular text,
as it does in Word. A paragraph style sets the value, and so does direct formatting.

One compatibility rule of Word's is applied too: in a table, a font size of 11 or
12 pt or left alignment that comes from the Normal style doesn't override the table
style (unless the document's `overrideTableStyleFontSizeAndJustification`
compatibility setting says otherwise).

All lengths are in points, and the values come in the forms the writers take, so
you can copy formatting from one paragraph to another:

```python
q = doc.add_paragraph("Copied spacing.")
doc.format_paragraph(q, line_spacing=LineSpacing.exactly(14), space_after=6)
r2 = doc.add_paragraph("Same spacing.")
source = doc.formatting().paragraph(q)
doc.format_paragraph(r2, line_spacing=source.line_spacing, space_after=source.space_after)
print(doc.formatting().paragraph(r2).line_spacing)  # LineSpacing(rule='exact', value=14.0)
```

A resolver reads the styles once when it is made; after changing styles, make a new
one. `doc.effective_format(p, r)` is a one-off `fmt.run(r)` (or `fmt.mark(p)`
without a run).

What isn't modelled: complex-script variants (`w:bCs`, `w:szCs`), indents and
spacing given in character units or lines, East Asian layout options, frames,
conditional formatting recorded only in `w:cnfStyle`, and the effect of
`contextualSpacing` between neighbouring paragraphs (the flag is reported; the
spacing it suppresses is not subtracted).

## Lists

```python
steps = doc.new_list("decimal")
steps.add("Collect")
steps.add("by region", level=1)
steps.add("Report")
for item in doc.list_labels():
    print(item.level, repr(item.label), item.format, item.suffix)
# 0 '1.' decimal tab
# 1 'a.' lowerLetter tab
# 0 '2.' decimal tab
```

`list_labels()` numbers lists as Word does: numbering from styles and from
numbering styles (`numStyleLink`), restarts (`startOverride`), `lvlRestart`, legal
numbering, levels linked to styles (how headings are usually numbered), and
`numId 0` switching a style's numbering off. Lists that share a definition share a
counter, as in Word. Headers, footers, notes and comments are included with
`list_labels(all_stories=True)`; each part has its own counters.

Bullets are often stored as codes in the Symbol or Wingdings fonts; `label` gives
the character they look like ("•", "▪", "➢") and `raw_label` the stored code.
Picture bullets have the image element in `item.picture`.

Formats Word draws from other writing systems (Hebrew, Arabic, CJK and others)
read as decimal numbers.

## Sections, headers and footers

```python
section = doc.sections[0]
print(section.page_width, section.orientation)     # from the document's page setup
print(section.header() is None)                    # True: no header of its own
print(section.shown("header") is None)             # True: Word shows none
```

`header(kind)` is the section's own header or None; `shown(kind)` follows "Link to
Previous" back through earlier sections, the way Word decides what a page shows.
`ensure_header(kind)` creates one.

## Pictures

```python
doc.add_picture("logo.png")
for picture in doc.pictures():
    image = doc.image_of(picture)
    print(picture.kind, image.content_type, len(image.data))   # drawing image/png ...
```

`pictures()` lists pictures only (not charts or text boxes), DrawingML and VML.
`image_of` also takes any element that holds an image reference.

## Checking it

The rules on this page are tested against files from Word, LibreOffice and other
libraries (`tests/`), and against Office itself with the [Office check kit](../office.md).
If a reader disagrees with what Word shows for one of your files, that is a bug
worth reporting with the file.
