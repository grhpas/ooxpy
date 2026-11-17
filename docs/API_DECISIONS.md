# API decisions before 1.0

The questions about the public API that were left open after the first release,
and what was decided (2026-11-17). All of them were made in 0.6.0, at once and
without compatibility aliases: nobody depended on oox yet. The rules they produce
are kept in [API stability](STABILITY.md#conventions-the-api-keeps).

The guiding rule: **writers are verbs, readers are nouns**, and the same idea has
the same name and units in all three formats.

## 1. Excel reader/writer pairs: renamed

The pairs followed three patterns (`set_column_width`/`column_width`,
`freeze_panes`/`frozen_panes`, `page_setup`/`page_setup_info`). Now:

| Writer | Reader |
|---|---|
| `set_frozen_panes(sheet, ref)` | `frozen_panes(sheet)` |
| `set_page_setup(sheet, **options)` | `page_setup(sheet)` |
| `set_sheet_options(sheet, **options)` | `sheet_options(sheet)` |
| `add_validation(sheet, ref, …)` | `validations(sheet)` (unchanged) |
| `set_column_width`, `set_row_height` | `column_width`, `row_height` (unchanged) |

Considered against: `freeze_panes` is Excel's and openpyxl's word. The docstring of
`set_frozen_panes` says "freeze" so searches still find it.

## 2. `add_textbox` → `add_text_box` in PowerPoint

Word already had `add_text_box`, `text_boxes()` and `set_text_box_text()`; oox's
own consistency outweighs python-pptx's spelling.

## 3. Readers return records

Readers returned a mix of tuples (which can't grow without breaking unpacking) and
dicts (no completion, misspelt keys fail late), and the same concept had different
shapes in Word and Excel. Options were records everywhere, dicts everywhere, or
leaving it.

**Decided: records** (`oox.records`): frozen dataclasses with named fields and
`as_dict()`. One class per concept across formats (`Hyperlink`, `Comment`, `Chart`
…). Excel readers keyed by cell stay dicts, of records. Field names are the
writers' keyword names, so what is read can be written back. Readers whose result
is the data itself (`values()`, `defined_names()`) stay plain lists and dicts.
Records carrying an XML element leave it out of `as_dict()`.

## 4. PowerPoint geometry: keywords with defaults

`add_shape` and `add_text_box` required four positional numbers, `add_table` two,
while `add_picture` and `add_chart` had defaults.

**Decided:** `x`, `y`, `width`, `height` are optional keywords everywhere (the
names `add_picture` and Word use; DrawingML's `cx`/`cy` are gone). The default is
the slide's content area: the box of its master's body placeholder, so a deck
opened from a file uses its own template. The content comes right after the slide.

## 5. Units: a plain number is points

Units were mixed: points for formatting, twips for Word page setup and table
widths, EMU for pictures, charts and PowerPoint positions, inches for Excel
margins.

**Decided:** a plain number is points everywhere, as in Office's dialogs; anything
else is a `Length` (`Cm(2)`, `Inches(1)`, `Emu(n)`). Readers report points. Excel
column widths stay in characters (Excel's own unit). Code ported from libraries
that pass raw EMU, or old oox code passing twips, is caught where it matters: each
setting checks a plausible range (Word pages at most 22 inches, drawings at most
56 inches; margins must leave room on the page) and the error names `Emu()`.

**Added in 0.6.1:** arithmetic keeps a `Length` a `Length` (`Cm(2) + Cm(1)`,
`Cm(2) * 1.5`, `sum(...)`); in 0.6.0 it gave a plain `int`, which then read as
points. A `Length` plus a plain number raises `TypeError`, since nothing says
whether the number is points or EMU.

## 6. Smaller ones

- **6a.** `doc.effective_format(p, r)` stays as the obvious shortcut for
  `doc.formatting().run(r)`; no more shortcuts are added.
- **6b.** `oox.grammar.URI` stays internal; `oox.ns` constants (`W`, `A`, `R` …)
  and `qn()` are the public way.

## 7. Added in 0.6.1, after the first review of 0.6.0 in use

- **7a. Readers report what Word shows, where Word and the standard differ.** A
  paragraph without `w:widowControl` anywhere in its hierarchy has widow control
  on (Word), not off (the standard). Such cases are documented where they apply.
- **7b. Ids are ints** where the file stores numbers, and every method that takes
  an id takes the type the readers give (it also accepts the XML's string form).
- **7c. One resolver for paragraphs, runs and table cells**, with `explain()`
  saying which layer each value came from; cell records use `Table.format_cell`'s
  names (`fill`, `valign`, `borders`, `margins`).
