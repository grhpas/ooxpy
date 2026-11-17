# Limitations and FAQ

## What isn't supported

- **Rendering and conversion.** ooxpy reads and writes the files; it doesn't lay out
  pages, render to PDF or convert between formats. Use Office or LibreOffice for that.
- **Formula calculation.** Formulas are stored, not calculated. Pass a cached result
  (`Formula("SUM(A1:A3)", 6)`) when readers should see a value; Excel recalculates on
  opening. Table totals rows are the exception: oox computes their results.
- **Layout-dependent results.** Page numbers (PAGE, NUMPAGES, the page numbers in a
  table of contents) are filled in by Word when it updates fields; page
  cross-references are estimates until then.
- **Encrypted files.** Password-encrypted files (an encrypted container around the
  package) are recognised and refused with a clear error. Editing restrictions
  (`doc.protect`, `book.protect_sheet`) are supported: they aren't encryption.
- **Legacy binary formats** (.doc, .xls, .ppt) are recognised and refused.
- **No helpers yet** for equations, SmartArt, mail merge fields and pivot tables, and
  for PowerPoint's transitions, animations and handout master. Files containing them
  open and save with that content untouched, and the element factories can build it
  (`oox.ns.m` for equations, `oox.ns.dgm` for SmartArt).

## Behaviour worth knowing

- **Excel display text** (`display_text`) renders what an English-locale Excel shows;
  Excel itself uses the viewer's regional separators and names unless the format has a
  locale prefix such as `[$-409]`. Fraction formats and conditional sections
  (`[>100]`) aren't rendered. Column widths assume Calibri 11.
- **Tracked edits** (`suggest_*`) split runs that hold plain text; runs that mix text
  with tabs, breaks or fields are left whole.
- **Single-cell reads** are cached and checked on every lookup. If you rewrite a
  shared string's text in place through the raw tree, call a structural edit (or open
  the file again) before reading.
- **Strict files** keep Word's own Strict spellings (`w:jc="start"`); Transitional
  accepts both.

## FAQ

**How does it compare with python-docx, openpyxl and python-pptx?** Those libraries
offer a hand-written object model for one format each, covering what their authors
modelled. ooxpy covers the whole vocabulary of all three formats with generated,
validated element factories, and adds conveniences on top; you can always drop down to
the exact markup. Files move freely between them: each opens the others' output.

**Why is the package `ooxpy` but the import `oox`?** `oox` is short to type, and
`ooxpy` says what the project is on PyPI.

**Which Office versions are supported?** Files are Office 2007-compatible unless you
use a newer feature; `validate(target="Office2007")` and `save(target=...)` tell you
which ones. See [Office compatibility](office.md) for what has been checked.

**Is the output deterministic?** Yes: the same document saves to the same bytes, and
untouched parts of an existing file are written back byte for byte.

**Does it work without LibreOffice or Office installed?** Yes. It's pure Python plus
lxml; Office and LibreOffice are only used to verify the library.

**Can I use it on untrusted uploads?** Opening is bounded by `Limits` (zip bombs, huge
parts, DTDs are refused) and every failure is a `PackageError`. See
[Existing files](guide/existing-files.md).

**How big is it, and does it work with PyInstaller?** The installed package is about
12 MB: about 5.5 MB of grammar tables, 2 MB of type stubs for editors, and the
element factories. The grammar for a namespace is loaded the first time it's used, so
a script that only writes Word documents doesn't load Excel's tables. oox imports
those modules by name, which PyInstaller can't see on its own; oox ships a PyInstaller
hook that lists them, found automatically, so `pyinstaller app.py` needs no extra
options. The stubs are not bundled.
