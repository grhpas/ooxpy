# API stability

This page says what you can rely on between versions, and how changes are made.

## Versions

oox follows [semantic versioning](https://semver.org/) from 1.0 on. Until then,
minor versions (0.5 → 0.6) may still change the public API; every such change is
listed under **Changed** or **Removed** in the [changelog](changelog.md), with the
replacement. Patch versions (0.5.0 → 0.5.1) never change the API, with one
exception: 0.6.1 carries the fixes from the first review of 0.6.0 in use, some of
which add to the API or correct a default (each listed under **Changed** in the
changelog), because 0.6.0 had a single user, who asked for it.

## What is public

The public API is pinned in `tests/api/public_api.txt`, so any change to it fails
the test suite until it's reviewed and the file is updated deliberately
(`python -m tools.api_surface --write`). It contains:

- everything `oox` exports: the document classes, `open_document`, `OoxElement`,
  `lenient`, `Package`, `Part`, `Limits` and the exceptions;
- every public method and property of `WordprocessingDocument`,
  `SpreadsheetDocument`, `PresentationDocument`, `OoxDocument` and `OoxElement`;
- the objects those methods return: `Table`, `Section`, `ListHandle`, `ListItem`,
  `SlideTable`, `Field`, `Caption`, the property sets, `StyleSheet`, `Series`,
  `Formula`, `CellError`, `Formatting`, `RunFormat`, `ParagraphFormat`,
  `TableCellFormat`, `ImageData`,
  `Relationships` and `Relationship`;
- the records the readers return (`oox.records`: `Hyperlink`, `Comment`, `Chart`,
  `PageSetup`, `Shape` …);
- the formatting value types `LineSpacing`, `Border`, `Borders` and `Tab`
  (`oox.formatting`);
- the utility modules `oox.units`, `oox.teletype`, `oox.odfcompat`, `oox.errors`,
  and the helpers `cell`, `row`, `cell_ref`, `column_letter`, `split_ref`,
  `excel_serial`, `from_excel_serial` (`oox.spreadsheet`), `build_chart`,
  `chart_info`, `PALETTE` (`oox.charts`);
- `oox.ns`: the namespace constants (`W`, `A`, `R` …), `qn` and the element
  factories (see below).

## Conventions the API keeps

- **Writers are verbs, readers are nouns.** `set_page_setup` writes and
  `page_setup` reads; `add_validation` adds and `validations` reads. The same idea
  has the same name in all three formats (`add_text_box`, `hyperlinks`, `charts`).
- **Readers return records** (`oox.records`): frozen objects with named fields.
  Fields are only ever added, never renamed or removed within a major version;
  construct records with keywords, not positions, if you construct them at all.
  Where a writer exists, field names are its keyword names. Readers whose result is
  data (`values()`, `defined_names()`) return plain lists and dicts.
- **A plain number is points**; anything else is a `Length` from `oox.units`
  (`Cm(2)`, `Inches(1)`, `Emu(n)`). Readers report points. The one exception is
  Excel's column width, in characters as in Excel. Values far too large to be
  points are refused with an error that names `Emu()`. Arithmetic on lengths
  gives lengths; a length plus a plain number raises `TypeError`, since the
  number's unit would be ambiguous.
- **Ids are ints** where the file stores numbers (`ListItem.num_id`,
  `ParagraphFormat.numbering`), and the methods that take them take the same type.
- **Readers report what Word shows.** Where Word and the standard disagree (a
  missing `w:widowControl` is on in Word), the value is Word's, and the docs say so.
- **Positions and sizes are keywords** (`x`, `y`, `width`, `height`), optional
  where a sensible default exists.

## Reading never changes a document

Every reader — properties such as `doc.styles` or `section.page_width`, methods
such as `doc.tables()`, `doc.formatting()` or `section.header()` — leaves the
document exactly as it was: saving after any amount of reading writes the same
bytes as saving without it. Readers report what Word shows when a part or setting
is missing (Letter page, 1-inch margins, no header) instead of creating it; the
`ensure_*` methods (`ensure_styles()`, `ensure_header()` …) and setters are the
ones that write. `tests/test_reading_is_pure.py` classifies every public member as
a reader or a writer and checks each reader on bare and feature-rich files.

**Not public:** the modules behind the document methods (`oox.xlsheet`,
`oox.floating`, `oox.protection`, `oox.textboxes` …), `oox.grammar`,
`oox.lexical`, `oox.strict`, `oox._generated` and anything starting with an
underscore. Use them if you need to, but they may change in any release.

## Element factories

`oox.ns` is generated from the Open XML SDK's schema data, pinned to one revision
(`oox.GRAMMAR_SOURCE`). Moving to a newer SDK revision is a minor release and
normally only *adds* elements, attributes and enumeration values. If a newer
revision removes or renames something, that is called out in the changelog as a
breaking change.

## Validation

- **Relaxations** (accepting markup oox used to reject) can come in any release.
  They are made only with evidence that Office opens the file
  (`tools/officecheck/`), and listed in the changelog.
- **New checks** (rejecting markup oox used to accept) are breaking for code that
  builds such markup. They come only in minor releases, only when Office is
  shown to break on the markup, and are listed under **Changed**.
- `validate()` reports may gain new findings in any release: it's a report, not
  an API.

## Output

Saving is deterministic: the same oox version, given the same input, writes the
same bytes. Between versions the bytes may change (attribute order, defaults oox
adds), but a file that opened in Office keeps opening. Untouched parts of an
existing file are always written back byte for byte.

## odfpy names

The odfpy-compatible names (`addElement`, `getElementsByType`, `teletype.extractText`,
`OpenDocumentText()` …) are permanent: they are never deprecated or removed
(ADR-002 in the [architecture](ARCHITECTURE.md)).

## Deprecation (from 1.0)

A public name that is renamed or removed first keeps working for at least one
minor release and six months, emitting a `DeprecationWarning` that names the
replacement. It's removed only in the next major version.

## Python versions

The floor is Python 3.10. A Python version is dropped only after its end of life,
and only in a minor release.
