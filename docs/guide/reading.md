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
paragraphs, runs and table cells:

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

`fmt.explain(p)` (or a run, or a cell) says which layer each value came from, for a
"reveal formatting" view or for debugging:

```python
print(fmt.explain(p)["keep_with_next"])   # paragraph style 'heading 2'
print(fmt.explain(r)["italic"])           # direct formatting
```

Bold, italic, caps, small caps, strike, double strike, hidden, outline, shadow,
emboss and imprint text are *toggle*
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

Where Word and the standard disagree, the values are Word's. A paragraph with no
`w:widowControl` anywhere in its hierarchy has widow control *on*, as Word shows it
(the standard says off). Tab stop positions (`Tab.position`) are measured from the
left page margin, not from the paragraph's indent. Outline levels and list levels are
0-8, as stored; Word's dialogs show 1-9. A highlight is Word's colour name
(`run.highlight`, "yellow"); `run.highlight_color` gives the hex.

### Fonts for each script

Word keeps four fonts per run and picks one for each character: `font` for ASCII,
`font_high_ansi` for the rest of Latin, Greek, Cyrillic and most symbols,
`font_east_asia` and `font_complex` (Arabic, Hebrew, Thai, Indic scripts) for
theirs. A Greek document often gives Greek text another font than ASCII, so a
renderer should ask per character:

```python
g = doc.add_paragraph()
gr = doc.add_run(g, "Καλημέρα, World", font="Calibri", font_high_ansi="Tahoma")
greek = doc.formatting().run(gr)
print(greek.font_for("Κ"), greek.font_for("W"))   # Tahoma Calibri
```

`font_hint` ("eastAsia" or "cs") moves the characters two scripts share, and a
right-to-left or complex-script run uses `font_complex` throughout, with
`bold_complex`, `italic_complex` and `size_complex`. The choice works on whole
Unicode blocks, which is what Word does except for a few single characters.

### Table cells

`fmt.cell(tc)` resolves a cell's shading, vertical alignment, margins and borders
through the same layers: the table style and its conditional formats, the row's
exceptions (`w:tblPrEx`), the table's own properties and the cell's:

```python
t = doc.add_table([["Region", "Q1"], ["North", "120"]], header_rows=1)
t.format_cell(0, 0, fill="DDEBF7")
cell = doc.formatting().cell(t.cell(0, 0))
print(cell.fill, cell.valign, cell.margins["left"])   # DDEBF7 top 5.4
print(cell.borders.bottom)                            # Border(style='single', width=0.5, color=None, space=0.0)
```

`borders` holds the line Word draws on each edge. Where two cells meet, Word draws
one line, chosen by MS-OI29500's rules: `nil` wins, a missing border loses, then the
heavier border (width times a number for the style; dotted and dashed weigh 1), then
the style earlier in Word's list, then the darker colour. With cell spacing there is
no conflict. The names match `Table.format_cell(fill=, valign=, borders=,
margins=)`, so a cell's look can be copied.

### What isn't modelled

Indents and spacing given in character units or lines, East Asian layout options,
frames, conditional formatting recorded only in `w:cnfStyle`, right-to-left tables
(`w:bidiVisual`; left and right are not swapped), diagonal cell borders, the merged
look of vertically merged cells (each `w:tc` is reported as stored), theme fonts
given per script (`a:font script=...`), and the effect of `contextualSpacing`
between neighbouring paragraphs (the flag is reported; the spacing it suppresses is
not subtracted).

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

`item.num_id` and `item.abstract_id` are ints, the same ids
`ParagraphFormat.numbering` gives and `oox.numbering.Definitions(doc).levels()` takes.

Bullets are often stored as codes in symbol fonts; `label` gives the character they
look like ("•", "▪", "➢", "✓") and `raw_label` the stored code. Symbol and the bullet
shapes of Wingdings, Wingdings 2, Wingdings 3 and Webdings are known. When a code
isn't, `item.needs_symbol_font` is True: draw `label` in `item.font`.
Picture bullets have the image element in `item.picture`.

Besides the Western formats, labels can be Russian, Hebrew or Arabic letters, Hebrew
numerals, Arabic-Indic or Thai digits, Japanese kana (aiueo, iroha), CJK digits and
counting (一〇, 二十一), the heavenly stems and earthly branches, and Korean ganada
and chosung. Hindi letters, Thai letters and counting, Vietnamese and Korean
counting and the Japanese legal formats read as decimal numbers.

## Sections, headers and footers

```python
section = doc.sections[0]
print(section.page_width, section.orientation)     # from the document's page setup
print(section.header() is None)                    # True: no header of its own
print(section.shown("header") is None)             # True: Word shows none
```

`header(kind)` is the section's own header or None; `shown(kind)` follows "Link to
Previous" back through earlier sections, the way Word decides what a page shows.
It walks back from the section itself, so reading every header of a long document
stays fast. `ensure_header(kind)` creates one.

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
