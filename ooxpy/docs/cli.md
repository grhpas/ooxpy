# Command line

Installing ooxpy puts an `oox` command on your PATH; `python -m oox` is the same
program. It has four commands; each takes a .docx, .xlsx or .pptx file (and the macro
and template variants).

## `oox validate`

Checks every part of a file against the schema and prints the problems per part.

```text
$ oox validate report.docx --target Office2007
report.docx: no problems, valid for Office2007

$ oox validate from-libreoffice.xlsx --target Office2007
/xl/workbook.xml
  - <x:workbookPr>: attribute dateCompatibility needs Office2010 (not understood by Office2007)
from-libreoffice.xlsx: 1 problem(s) in 1 part(s)
```

| Option | |
|---|---|
| `--target VERSION` | also report markup that Office version can't read: `Office2007` … `Microsoft365` |
| `--full` | also match whole content models (exclusive choices, occurrence counts) |
| `--json` | machine-readable output: `{part: [problems]}` |
| `--limit N` | problems shown per part (default 20) |

The exit status is 0 when there's nothing to report, 1 when there are problems and 2
when the file can't be opened, so the command fits in CI:

```bash
for f in build/*.docx; do oox validate "$f" --target Office2007 || exit 1; done
```

## `oox info`

The kind of document, its parts and properties, and a summary per sheet or slide.

```text
$ oox info sales-report.docx
kind: WordprocessingDocument
conformance: transitional
parts: 8
main: /word/document.xml
title: Sales report
paragraphs: 29
tables: 1
charts: 1
```

`--json` gives the same as JSON (for a workbook, rows, charts and tables per sheet).

## `oox text`

The text of a document: paragraphs for Word (a text box's paragraphs follow the
paragraph that anchors it), tab-separated rows per sheet for Excel, and the text of
each slide with its notes for PowerPoint.

```text
$ oox text status.pptx
# Slide 1
Project status
November 2026
# Slide 2
Falcon: On track
Pilot in North
Launch in May
[notes] Owner of Falcon to present.
```

## `oox downgrade`

Saves a copy for an older Office version, removing (or, for Word 2010's `w:start` and
`w:end`, renaming) the markup it can't read, and lists what changed.

```text
$ oox downgrade from-libreoffice.xlsx for-2007.xlsx
/xl/workbook.xml
  - <x:workbookPr>: attribute dateCompatibility needs Office2010 (not understood by Office2007)
for-2007.xlsx: written for Office2007, 1 change(s)
```

`--target` picks another version (default `Office2007`).
