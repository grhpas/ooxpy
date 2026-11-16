# Changelog

All notable changes to ooxpy. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
[semantic versioning](https://semver.org/) from 1.0 (see docs/STABILITY.md).

## [0.5.0] — unreleased (first public release)

The first release on PyPI. It covers the roadmap's phases 0–4 and most of phase 3.
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
