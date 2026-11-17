# Changelog

All notable changes to ooxpy. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
[semantic versioning](https://semver.org/) from 1.0 (see [API stability](https://github.com/grhpas/ooxpy/blob/main/docs/STABILITY.md)).

## [0.6.1] — 2026-11-17

Fixes and additions from the first review of 0.6.0 in use (Petals Office, a word
processor that reads .docx through oox). Some of them add to the API or correct a
default; by the maintainer's decision they ship as a patch release, because 0.6.0
had a single user (see [API stability](https://github.com/grhpas/ooxpy/blob/main/docs/STABILITY.md)). Given the same settings, the files oox writes
are unchanged byte for byte.

### Fixed

- **List ids are ints everywhere.** `ListItem.num_id` and
  `ParagraphFormat.numbering` gave ints, but `numbering.Definitions.levels()`
  looked ids up as strings, so passing it what the other two give returned None
  (no error, no list). `levels()` now takes an int (or the XML's string form), and
  it and `numbering_of()` return int ids.
- **Lengths stay lengths in arithmetic.** `Cm(2) + Cm(1)`, `Cm(2) * 2`, `-Cm(1)`,
  `abs()` and `sum([...])` gave a plain `int`, which oox then read as points:
  12,700 times too big. They now give a `Length`; `Length / Length` is a plain
  ratio. **Changed:** a `Length` plus a plain number other than 0 raises
  `TypeError` (the number's unit is ambiguous: write `Cm(2) + Pt(6)`), and
  `repr(Cm(1))` is `Emu(360000)`.
- **Widow control is on by default**, as Word shows it: `ParagraphFormat.widow_control`
  is True unless something in the hierarchy turns it off (the standard says off
  when absent; Word writes `w:val="0"` to turn it off). **Changed** default.
- `Section.shown()` walked the whole body twice on every call (to find the
  section, then to count sections sharing the header, which only matters when
  creating one). It now walks back from the section itself, so reading every
  header and footer of a long document is linear.
- `Part.related_part()` on a part outside a package said the target was missing,
  which read as a broken file; it now says the part is not in a package.
- `Record.as_dict()` mangled `Border`, `Tab` and `LineSpacing` values inside a
  record; they stay as they are, ready to be written back.
- `Tab.position` is documented as measured from the left page margin (as
  `w:tab`'s `pos` is), not from the paragraph's indent.

### Added

- **The font for Greek, Cyrillic and accented text.** `RunFormat` gains
  `font_high_ansi` (`w:hAnsi`, the font Word uses outside ASCII for Latin, Greek,
  Cyrillic and symbols) and `font_hint`, resolved through the same layers and theme
  fonts as `font`; `rf.font_for(char)` gives the font Word uses for a character
  (and `oox.resolve.script_slot` which of the four). The writers take
  `font_high_ansi`, `font_east_asia`, `font_complex` and `font_hint` to set one
  script's font.
- **Table cells in the resolver.** `formatting().cell(tc)` returns a
  `TableCellFormat`: shading, vertical alignment, cell margins and the border Word
  draws on each edge, through the table style and its conditional formats, the
  row's exceptions, the table's and the cell's own properties, with border
  conflicts between neighbouring cells resolved as MS-OI29500 describes Word doing
  it (`oox.resolve.resolve_conflict`). `Table.format_cell` takes `borders` and
  `margins` under the same names.
- **Where each value comes from.** `formatting().explain(el)` maps each field of a
  paragraph, run or cell to its layer ("paragraph style 'Heading 1'", "table style
  'Grid', first row", "list 3, level 0", "direct formatting" ...).
- More properties, read and written: raised/lowered text (`position`), character
  scale, kerning, emphasis marks, outline, shadow, emboss, imprint, right-to-left
  and complex-script runs with their own bold, italic and size; right-to-left
  paragraphs and "Don't hyphenate" (`suppress_hyphens`).
- `RunFormat.highlight_color` gives the hex of Word's highlight name, and
  `oox.formatting.HIGHLIGHT_COLORS` the whole table; the writers take a highlight
  as a name or its hex.
- `RunFormat` and `ParagraphFormat` are records (`as_dict()`).
- More number formats: Russian, Hebrew (letters and numerals) and Arabic (alphabet
  and abjad) letters, Arabic-Indic and Thai digits, the Japanese kana orders (aiueo,
  iroha, half and full width), CJK digits and Japanese and Chinese counting, the
  heavenly stems and earthly branches, Korean ganada and chosung. The formats that
  still read as decimal are listed in the reading guide.
- Symbol fonts: `readable_text()` knows Symbol's Greek letters and signs, and the
  bullet shapes of Wingdings 2, Wingdings 3 and Webdings besides Wingdings;
  `numbering.needs_symbol_font(text)` and `ListItem.needs_symbol_font` say when a
  character couldn't be converted and must be drawn in its symbol font.
- The reading guide covers fonts per script, table cells, `explain()`, where Word
  and the standard disagree, and what still isn't modelled.

## [0.6.0] — 2026-11-17

This release follows a review of reading existing files, and settles the API
questions left open for 1.0 ([API decisions](https://github.com/grhpas/ooxpy/blob/main/docs/API_DECISIONS.md)). Nobody depended on 0.5 yet,
so the changes replace the old shapes rather than adding a second way to do things;
every change below names the replacement. Given the same settings, the files oox
writes are unchanged byte for byte.

### The API's conventions (see [API stability](https://github.com/grhpas/ooxpy/blob/main/docs/STABILITY.md))

- **Changed: readers return records.** About thirty readers that returned dicts or
  tuples now return frozen records from `oox.records`, with named fields and
  `as_dict()`: `Hyperlink` (Word, Excel and PowerPoint links), `Comment` (Word and
  Excel), `Note` (footnotes, endnotes), `Chart` with `charts.Series` (every format's
  `charts()` and `chart_info()`), `Picture`, `TextBox`, `ContentControl`, `Revision`,
  `Protection`, `Field` (from `fields()`), `CellFormat`, `ConditionalFormat`,
  `Validation`, `ExcelTable`, `PageSetup`, `SheetOptions`, `WorkbookProtection`,
  `Placeholder`, `Shape`, `SlideTableInfo` and `SlideSection`. Excel readers keyed
  by cell (`comments`, `hyperlinks`) stay dicts of records. Field names are the
  writers' keywords, so a record can be written back
  (`book.set_page_setup("T", **book.page_setup("S").as_dict())`,
  `deck.set_sections(deck.sections())`); `conditional_formats()` now reports oox's
  rule names and options (`rule="cell"`, `operator=">="`, `value=100`) as well as
  Excel's `type`, and `page_setup()` gives `print_area`, `repeat_rows` and
  `repeat_cols` in the form `set_page_setup` takes.
- **Changed: a plain number is points everywhere**; anything else is a `Length`
  from `oox.units` (`Cm(2)`, `Inches(1)`, `Emu(n)`). This now also covers Word page
  size and margins (`Section.page_width`, `margins`, `text_width`, `set_margins`;
  new `set_page_size`), Word table widths and row heights, picture, chart and
  text-box sizes and offsets in all three formats, Excel page margins (were inches),
  PowerPoint positions and slide size, and `SlideTable` widths. Readers report
  points. A plain number far too large to be points (a raw EMU or twips value)
  raises `ValueError` suggesting `Emu()`. Excel column widths stay in characters.
  `ensure_style(text_width=...)` is in points too (it was twips: an old value is
  now too large and raises). In 0.5.2 plain numbers in page setup were twips (see
  below). Note: in 0.6.0, arithmetic on lengths (`Cm(2) + Cm(1)`) gives a plain
  `int`, which reads as points; fixed in 0.6.1.
- **Changed: PowerPoint geometry is keywords.** `add_text_box`, `add_shape`,
  `add_table`, `add_chart` and `add_picture` take `x`, `y`, `width`, `height` as
  optional keywords (`cx`/`cy` are gone; `add_table`'s `col_widths` is `widths`).
  Left out, things go in the slide's content area (`deck.content_area(slide)`; new
  `deck.slide_size`). The text or data comes right after the slide:
  `deck.add_text_box(slide, "Hello")`.
- **Renamed:** PowerPoint's `add_textbox` → `add_text_box` (as in Word); Excel's
  `freeze_panes` → `set_frozen_panes`, `page_setup` → `set_page_setup` (and
  `page_setup_info` → `page_setup`), `sheet_options` → `set_sheet_options` (and
  `sheet_options_info` → `sheet_options`); Word's `add_hyperlink(url=)` →
  `add_hyperlink(target=)`, as in Excel and PowerPoint.
- Style writers ignore options that are None, so a `CellFormat` read back can be
  passed to `cell_style` or as a conditional format's `style`.

### Reading the way Word does

- One walker (`oox.walk`) now decides what a reader sees, and every reader uses
  it: in `mc:AlternateContent` only the branch Word shows (text boxes are stored
  twice, as a Word 2010 shape and a VML copy for Word 2007), and paragraphs inside
  content controls and custom XML. Before, fields, bookmarks, hyperlinks, comment
  anchors, revisions, content controls, tables, charts, list labels and `oox info`
  counted text-box content twice.
- `update_fields()` numbered a caption in a text box twice (a second caption came
  out as "Figure 3"); it now numbers the shown copy and writes the same result into
  the hidden one.
- `sections()` missed section breaks inside content controls.
- `pictures()` returns pictures only (it also listed charts, text boxes and other
  drawings), includes VML pictures, and gives each one's `kind`, `embed`/`link`
  ids and `svg`.

### Reading never changes a document

- A rule now, enforced by a test over every public member (see [API stability](https://github.com/grhpas/ooxpy/blob/main/docs/STABILITY.md)).
- `doc.styles`, `doc.settings` and `doc.numbering` return None when the part is
  missing; `ensure_styles()`, `ensure_settings()` and `ensure_numbering()` create it.
- A document without a final `w:sectPr` has one virtual section reporting Word's
  defaults (Letter, 1-inch margins); it's written only when you set something.
- `Section.page_width`, `page_height`, `orientation` and `margins` report defaults
  without writing; `margins` always has all seven keys.
- `Section.header(kind)` and `footer(kind)` only read (the `create` argument is
  gone); `ensure_header(kind)` and `ensure_footer(kind)` create.
- The `track_changes` and `even_and_odd_headers` getters no longer create the
  settings part.
- `Section.shown()` on a section that is no longer in the document raises
  `OoxError` saying so (it failed with a bare `StopIteration`); its docstring says
  what None means for a first-page header.

### Formatting

- **Changed:** `doc.formatting()` returns a resolver: `.paragraph(p)` gives a
  `ParagraphFormat`, `.run(r)` and `.mark(p)` a `RunFormat` (frozen dataclasses).
  `doc.effective_format(p, r)` now returns a `RunFormat` instead of a dict;
  paragraph values moved to `formatting().paragraph(p)`.
- The resolver applies Word's layers in Word's order: document defaults, the
  table style with its conditional formats (header row, banding, first and last
  row and column, corners) as switched on by `tblLook`, the paragraph style chain,
  numbering indents (before the style when the numbering comes from the style),
  character styles, direct formatting. Bold, italic and the other toggle properties
  follow Word's rules (a bold character style on bold text gives regular text), as
  does Word's rule that the Normal style's 11/12 pt and left alignment don't
  override a table style.
- New properties, read and written: shading, letter spacing, hidden text,
  language, double strike and a border for runs; outline level, contextual
  spacing, widow control, shading, borders and tab stops for paragraphs.
- **Changed:** every length is in points, indents included (they were twips; a
  `Length` from `oox.units` such as `Cm(1)` still works, and the unit helpers now
  return `Length`, an `int` that knows it is EMU). Line spacing is a multiple or a
  `LineSpacing`: `LineSpacing.exactly(12)`, `.at_least(12)`, `.multiple(1.5)`;
  the `"12pt"` string is gone (it could not say "at least", which read back the
  same as "exactly").

### Lists

- **Changed:** `list_labels()` returns `ListItem`s (`paragraph`, `label`, `level`,
  `num_id`, `abstract_id`, `format`, `font`, `suffix`, `raw_label`, `picture`,
  `story`) instead of `(paragraph, label)` pairs.
- Lists that use a numbering style (`numStyleLink`) are numbered (their labels were
  empty), sharing the style's counters; `lvlRestart`, levels linked to a style
  (`w:lvl/w:pStyle`), the default paragraph style and `numId 0` are honoured.
- More number formats: ordinal ("1st"), cardinal and ordinal text ("Twenty-one",
  "Twenty-first"), hex, chicago (* † ‡ §), enclosed and full-width digits.
- Symbol-font bullets read as the character they look like ("•" instead of
  "\uf0b7"); `raw_label` keeps the stored characters. Picture bullets give the
  image element (`doc.image_of(item.picture)`).
- `list_labels(all_stories=True)` adds headers, footers, footnotes, endnotes and
  comments.

### Packages

- `oox.ns` has a constant per namespace (`from oox.ns import W, A, R, MC`) and
  `qn("w:p")`; use them instead of `oox.grammar.URI`, which stays internal.
- **Changed:** `part.rels` is a mapping (`rels[rid]`, `rels.get(rid)`,
  `rels.items()`); the `.items` attribute is gone.
- `part.related_part(rid)`, `part.blob`, `part.package`.
- `doc.image_of(picture)` gives `(data, content_type, partname, external)` for a
  picture, a blip, VML image data or a picture bullet, in all three formats.
- Named parts: `styles_part`, `settings_part`, `numbering_part`, `theme_part`,
  `footnotes_part`, `endnotes_part`, `comments_part`, `font_table_part` (Word);
  `shared_strings_part`, `styles_part`, `theme_part` (Excel).
- `doc.part_of(element)` is a dictionary lookup instead of a scan of every part.

### Also

- A guide, ["Reading a document as Word shows it"](https://github.com/grhpas/ooxpy/blob/main/docs/guide/reading.md), covers the walker, the
  formatting layers, lists, sections and pictures, and what isn't modelled.
- A PyInstaller hook, registered through the `pyinstaller40` entry point, bundles
  the modules oox imports by name; `pyinstaller app.py` needs no options.
- A corpus of .docx files written by LibreOffice from known sources
  (`tools/locorpus.py`, `tests/data/libreoffice/`), checked against the sources and
  LibreOffice's own reading: list labels, sections, pictures, text boxes, tables.
- The README's links are absolute (PyPI shows the README without the repository
  around it), and the package metadata links the repository, issues and changelog.

### Errors you may meet when upgrading from 0.5

- `AttributeError: 'function' object has no attribute 'get'`: `part.rels.items.get(rid)`;
  `rels` is a mapping now: `part.rels.get(rid)`.
- `cannot unpack non-iterable ListItem object`: `for p, label in doc.list_labels()`;
  use `for item in doc.list_labels()` with `item.paragraph` and `item.label`.
- `ValueError: ... is read as points ...`: a twips or EMU number where points are
  meant; divide twips by 20, or wrap EMU in `Emu()`.
- `TypeError: ... unexpected keyword argument 'url'`: `add_hyperlink(url=)` is
  `add_hyperlink(target=)`.
- `AttributeError` on a reader's result used as a dict or tuple: readers return
  records; use the field names (`link.target`), or `record.as_dict()`.

### Known issues (fixed in 0.6.1)

- `numbering.Definitions.levels()` takes the id as a string, while
  `ListItem.num_id` is an int: pass `str(item.num_id)`.
- Arithmetic on lengths gives a plain `int` (read as points): wrap sums in `Emu()`.
- `RunFormat.font` is the ASCII font only; Greek, Cyrillic and accented text may use
  another (`w:hAnsi`), which 0.6.0 doesn't report.
- `widow_control` reads False where Word shows it on.
- `Section.shown()` walks the whole body twice per call.

## [0.5.2] — 2026-11-17 (first public release)

The first release on PyPI (0.5.0 and 0.5.1 were never published: the version
numbers were used up while setting up trusted publishing). It covers the
roadmap's phases 0–2 and 4, and most of phase 3.
The Word, Excel and PowerPoint features below have cases in the Office check
corpus, all confirmed on Office 2007; Office 2016 opened every case without a
repair prompt (`tools/officecheck/reports/`). Strict-conformance files have been
checked with LibreOffice only.

### Core

- One factory per element for the whole Open XML SDK vocabulary (Word, Excel,
  PowerPoint, DrawingML, VML and Microsoft's extension namespaces), with typed
  stubs for IDE completion.
- Validation on construction (attributes, values, children, order), `validate()`
  reports at two levels (`level="full"` matches whole content models), and
  `validate(target="Office2007")` / `save(target=..., strip=True)` for older Office.
- Lossless editing: untouched parts are written back byte for byte; saves are
  deterministic.
- Strict-conformance files are read and written; untrusted files are bounded
  (`Limits`: zip bombs, DTDs, legacy and encrypted files give a `PackageError`).
- odfpy-compatible names (`addElement`, `teletype`, `load`, `OpenDocumentText()` …).
- The `oox` command: `validate`, `info`, `text`, `downgrade`.

### Word

- Default styles, settings, fonts, theme and page setup; headings, paragraphs,
  runs with formatting, custom styles and effective formatting.
- Lists with Word's own numbering and the labels Word displays.
- Sections, headers and footers, page fields; hyperlinks, bookmarks, footnotes,
  endnotes, comments; table of contents; find and replace across runs.
- Tables (merges, widths, header rows, cell formatting); inline and floating pictures
  (PNG, JPEG, GIF, BMP, TIFF, EMF, WMF, SVG); text boxes.
- Tracked changes (suggest, list, accept, reject); fields (captions,
  cross-references, document properties, dates) recomputed by oox; content
  controls; document protection with Word's password hash.
- Charts with an embedded workbook.

### Excel

- Typed cells, formulas with cached results, dates, shared strings; fast reads.
- Cell styles and number formats (with `display_text` as Excel renders them),
  column widths, row heights, merges, frozen panes, filters, names, data validation.
- Conditional formatting, tables with totals rows, comments, hyperlinks, print
  setup, sheet options, sheet and workbook protection.
- Charts from cell ranges (combinations, secondary axes) and chart sheets.

### PowerPoint

- Office's eleven standard layouts; placeholders (text and pictures), text boxes,
  shapes, pictures, tables (creating and editing), speaker notes, click links,
  sections, charts.

### Fixed since the last internal build

- Text extraction (`teletype.extractText`, `oox text`) counted a Word text box's
  text twice: once from Word 2010's shape, once from the VML copy kept for Word
  2007. Fallback copies are now skipped, and `extractText(p, text_boxes=False)`
  gives a paragraph's own text, as Word does.
- `TargetError` (raised by `save(target=...)`) is exported from `oox`.
- `oox.lenient()` now also accepts invalid attribute values when building elements, as
  documented (it used to turn off only the checks for unknown attributes and children).
  `validate()` reports such values whether or not it runs inside the block.
- A `Section` object taken before `add_section()` now keeps describing its own section
  (it used to switch to the new one).
- `Section.set_margins` accepts EMU from `oox.units` (`left=Cm(2)`), as Word tables do;
  plain numbers are still twips.

[0.6.1]: https://github.com/grhpas/ooxpy/compare/v0.6.0...v0.6.1
[0.6.0]: https://github.com/grhpas/ooxpy/compare/v0.5.2...v0.6.0
[0.5.2]: https://github.com/grhpas/ooxpy/releases/tag/v0.5.2
