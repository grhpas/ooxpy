# API stability

This page says what you can rely on between versions, and how changes are made.

## Versions

oox follows [semantic versioning](https://semver.org/) from 1.0 on. Until then,
minor versions (0.5 → 0.6) may still change the public API; every such change is
listed under **Changed** or **Removed** in the [changelog](changelog.md), with the
replacement. Patch versions (0.5.0 → 0.5.1) never change the API.

## What is public

The public API is pinned in `tests/api/public_api.txt`, so any change to it fails
the test suite until it's reviewed and the file is updated deliberately
(`python -m tools.api_surface --write`). It contains:

- everything `oox` exports: the document classes, `open_document`, `OoxElement`,
  `lenient`, `Package`, `Part`, `Limits` and the exceptions;
- every public method and property of `WordprocessingDocument`,
  `SpreadsheetDocument`, `PresentationDocument`, `OoxDocument` and `OoxElement`;
- the objects those methods return: `Table`, `Section`, `ListHandle`, `SlideTable`,
  `Field`, `Caption`, the property sets, `StyleSheet`, `Series`, `Formula`,
  `CellError`;
- the utility modules `oox.units`, `oox.teletype`, `oox.odfcompat`, `oox.errors`,
  and the helpers `cell`, `row`, `cell_ref`, `column_letter`, `split_ref`,
  `excel_serial`, `from_excel_serial` (`oox.spreadsheet`), `build_chart`,
  `chart_info`, `PALETTE` (`oox.charts`);
- the element factories in `oox.ns` (see below).

The dictionaries the readers return (`doc.comments()`, `book.tables(sheet)`,
`deck.shapes(slide)` …) are public too: keys are only added, never renamed or
removed within a major version.

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
