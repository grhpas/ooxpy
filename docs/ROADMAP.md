# oox: Roadmap

Every phase ends in something you can release or demo. Each phase is usable on
its own, and later phases build on earlier ones without rewriting them.

## Pace and estimates (revised 2026-11-14, updated 2026-11-15)

The original sizes assumed one engineer working alone (Phases 0–2 alone came
to roughly 6–9 weeks). In practice, Phases 0–2 and Excel step 1 were done in
**one long working session**, with about a dozen Office 2007 verification
rounds in between. The bottleneck is no longer writing code; it's
verification on real Office (each round takes a few minutes of someone's
time at the Windows machine) and reviewing decisions.

Remaining work, in working sessions like the first one:

| Work | Sessions | Notes |
|---|---|---|
| ~~Excel step 2: conditional formatting, tables (`x:table`), shared strings~~ | 1 | ✅ done in one session (2026-11-15), plus comments, links, print setup and sheet options |
| ~~Charts (DrawingML `c:`), shared by Excel, Word and PowerPoint~~ | 1–2 | ✅ done in part of one unattended session (2026-11-15); Office run pending |
| ~~PowerPoint depth: standard layouts, tables, speaker notes, picture placeholders~~ | 1–2 | ✅ done in the same unattended session (2026-11-15), plus shapes with fills; Office run pending |
| ~~Phase 1 leftovers: performance (large sheets), `save(target=...)`~~ | 0.5–1 | ✅ done in the same unattended session (2026-11-15) |
| Phase 3: ~~full validation automata, CLI~~, XSD frontend and IR diff | 1.5–2 | automata and CLI ✅ (2026-11-15); the XSD frontend is the heavy part left |
| Interop corpus in CI (python-docx/SDK test files, Word-authored samples) | 0.5 | started 2026-11-15: `tools/interop.py` (149 LibreOffice-converted templates + one Word-authored file, all pass open / unchanged save / readers / edit); `tools/writers.py` adds python-docx, openpyxl, XlsxWriter and python-pptx output, both ways (in the test suite); Word-authored samples from you still wanted |
| ~~Charts step 2 (combinations, secondary axes, chart sheets), PowerPoint step 2 (table editing, links, sections), fast single-cell reads~~ | 1 | ✅ done in one unattended session (2026-11-15); complete on Office 2007 the same day |
| ~~Phase 5: packaging, PyPI, API freeze, full documentation~~ | 2 | ✅ done in one session (2026-11-16) except the upload itself (needs the public repository and PyPI's trusted publisher) and the API decisions listed under Phase 5 |
| Office 2016 full run, Mac/LibreOffice run, fixes they reveal | 0.5–1 | user-side runs dominate. First 2016 run done 2026-11-16 on an unactivated Office: no file prompted, every answered probe passed; probes that change documents or read layout were refused by the licence and need a rerun on an activated Office. Rerun the same evening: `word-03`, `word-16` (behind-text wrapping confirmed), `excel-20` ok; five Word cases (accept/reject, paragraph formatting, unprotect) still wait for activation; the Excel screenshot probe was reworked for Excel 2013+ |

Update 2026-11-15: one unattended hour covered charts, PowerPoint depth, the
Phase 1 leftovers (performance, `save(target=)`), Phase 3's content-model
automata and CLI, and the start of the interop corpus. What remains is mostly
the XSD frontend, Office runs (2016, Mac/LibreOffice) and Phase 5 (packaging
and full documentation): **about 4–6 sessions to 1.0**, still bounded by
verification on real Office rather than by writing code.

The original estimate was **about 9–13 sessions to 1.0**. Wall-clock time depends mostly on
how often Office runs and reviews can happen. Each estimate covers the
engineering; findings from Office 2016 or real-world files may add work.

---

## Phase 0: Prototype ✅ (this delivery, v0.0.1)
**Deliverable:** a working generator and runtime for all three formats; 79 passing tests (plus an opt-in LibreOffice run); three runnable examples with HTML tutorials; the Office check kit.

- Generator: IR, Open XML SDK frontend and emitter, deterministic output, pinned to the SDK revision.
- Grammar tables for **all** 4,315 types / 3,702 elements / 139 namespaces (Word, Excel, PowerPoint, DrawingML, VML, Office extensions).
- `OoxElement`: context-resolved content models, schema-ordered insertion, construction-time validation, `validate()` reports, odfpy-compatible method names.
- OPC: open/save, lazy parsing, byte-identical round-trip of untouched parts, deterministic zips, automatic `mc:Ignorable`.
- **Documents:** `WordprocessingDocument`, `SpreadsheetDocument` and `PresentationDocument` on a shared `OoxDocument` base; `open_document()` picks the right one.
  - **xlsx:**
    - sheets, A1 helpers, typed cells (numbers, booleans, inline strings, formulas), `append_row`
    - **reading** (`value`/`values`/`cells`): shared strings incl. rich text, phonetic runs and `_xHHHH_` escapes; errors; cached formula results
    - `set_value`: ordered insertion; keeps the cell's style
  - **pptx:**
    - the full master → layout → theme skeleton, with "Title Slide", "Title and Content" and "Blank" layouts
    - slides that copy the layout's placeholders, like PowerPoint's "New Slide"
    - `set_placeholder_text` with bullet levels, by alias, type or index; works on PowerPoint-authored decks
    - text boxes
- **Pictures in all three formats:** PNG/JPEG/GIF/BMP read from file headers (no Pillow); natural size from DPI, or scaled with the aspect ratio kept; identical images stored once; `oox.units` (EMU, inches, cm, pt, px).
- **Generic part creation** from the SDK part table (content type, relationship, folder and numbering).
- **Partial minOccurs check:** `validate()` reports missing mandatory children.
- **First documented SDK overrides** (`tools/gen/overrides.py`): `x:CT_Font` and `x:CT_RPrElt` are unordered choices in ECMA-376.
- **odfpy compatibility layer:** odfpy DOM methods, factory keywords, attribute spellings, style references, `load()`, `teletype`, `doc.styles`. Four of odfpy's own example scripts are translated as tests (docs/ODFPY_COMPAT.md).
- **Evidence:**
  - Zero validation false positives on a corpus of about 30k elements: Word (python-docx's template), PowerPoint (python-pptx's template), LibreOffice (docx/xlsx/pptx), openpyxl and python-pptx output.
  - Output reads back in python-docx, openpyxl and python-pptx, and renders in LibreOffice.
  - **Microsoft Office:**
    - The showcase .docx/.xlsx/.pptx opened without errors on the team's Windows machines (Office 2007 and 2016).
    - **Office 2007, scripted (2026-11-14): complete.** 18/18 cases ok and 3/3 negative controls detected, including every content probe and the w14/mc:Ignorable case (`tools/officecheck/reports/`).
- **Office check kit** (`tools/officecheck/`):
  - a deterministic corpus: 18 cases with content probes, plus 3 negative controls (malformed files Office must reject, which prove the dialog detection works)
  - a manual checklist
  - a standalone Windows/COM checker (Python 3.8, dialog watchdog)
  - a LibreOffice runner, which passes all 18 cases
  - a report merger
- **Tutorials:** `examples/*.html`, three short pages next to their scripts. Every code block runs in the test suite.
- **Attributions:** `docs/ATTRIBUTIONS.md`.
- **Protective round** (Phase 1): see below.
- **Lossless editing of existing files:** read-only round trips are byte-identical per zip entry, and edits change only the affected parts, their relationships and the content-types file.

## Phase 1: Harden the core ✅ (protective round done 2026-11-14) → **0.1.0 alpha**
0. **Office regression runs** (`tools/officecheck/`):
   - Office 2007: complete, 18/18 ok and 3/3 controls detected.
   - The corpus now has 27 cases: 18 normal, 3 negative controls, and 6 **observe** cases (well-formed but schema-invalid, recorded without pass/fail to learn how strict Office is).
   - Office 2007 observe results: invalid values, Excel row order, a newer attribute in an old namespace, and a missing required element were all **fatal**; out-of-order properties and unknown elements were tolerated. Every fatal case is something `validate()` reports (`tools/officecheck/reports/README.md`).
   - Office 2016 (2026-11-16, unactivated): all 58 normal cases opened without a repair prompt and the 3 controls were detected; probes refused by the missing licence still to confirm on an activated Office. Still to run: LibreOffice on the Mac mini.
   - ✅ `save(target="Office2007")` refuses markup the target can't read (`TargetError` with findings per part); `strip=True` removes it, renaming Word 2010's `w:start`/`w:end` back to `w:left`/`w:right`. Motivated by Excel 2007 refusing LibreOffice's `dateCompatibility`. Also `python -m oox downgrade`.
1. ✅ **CI** (`.github/workflows/ci.yml`):
   - tests on Python 3.10–3.14
   - ruff and pyright (both clean)
   - a generated-code job: fetch the pinned SDK (`python -m tools.gen.fetch_sdk`; revision in `tools/gen/SDK_REVISION`, which the generator now enforces) and regenerate
   - a LibreOffice probe job
   - `office-check.yml`: a manual run on self-hosted Windows runners with Office
2. ✅ **Lexical validation**: the SDK's String/Number/Enum validators compiled into value checks, including unions ("integer or `1.5in`") and hexBinary byte lengths. Messages say what's expected ("6 hex digits"). Zero false positives on 31 real files.
3. ✅ **Typed kwargs** (2026-11-15): the generator writes `.pyi` stubs next to every factory module (`py.typed` included). Each factory lists its attribute keywords (the union over its content models, named exactly as the runtime names them) with hints: enumerations as `Literal[...]`, booleans, numbers, dates. IDEs complete them and type checkers catch e.g. a float passed to `w:jc/@w:val`. `**attrs` stays for the other spellings the runtime accepts, so nothing that runs today becomes a type error. Plain stubs need no newer Python, so the floor stays at 3.10.
4. ✅ **Office version targeting**: `doc.validate(target="Office2007")` reports markup that version won't understand (newer elements or attributes in old namespaces), and skips what it safely ignores (newer namespaces, `extLst`, `mc:AlternateContent`). First finding: LibreOffice writes Office 2010's `x:workbookPr/@dateCompatibility`.
5. ✅ **Strict ↔ Transitional**: Strict files are read into Transitional trees and written back as Strict. Unchanged parts stay byte-identical, and `doc.conformance = "strict"` converts. LibreOffice reads our Strict output in all three formats. Value forms (2026-11-15): DrawingML percentages (`"50%"` ↔ 50000) and ST_OnOff (`on`/`off` → `true`/`false`) are translated too; LibreOffice reads a Strict crop of 18.339% as the same crop.
6. ✅ **Hostile input**:
   - `Limits` (part count, part size, total size, compression ratio; zip bombs are caught by default)
   - duplicate zip entries rejected
   - DTDs refused
   - encrypted and legacy (.doc/.xls/.ppt) files get a clear message
   - every failure is a `PackageError`
   - a deterministic 400-case mutation fuzz test
7. ✅ **Performance** (2026-11-15): building is linear. `append_row`/`set_value` find the last row and cell directly; schema-ordered insertion scans from the end and inserts with `addnext` (lxml's `len()` and indexing walk every child, which made appending quadratic, and inserting before a trailing `w:sectPr` too); `qname_of` is memoized; `trusted()` skips re-validating oox's own ambiguous elements (`x:row`, `x:c`). 10,000 rows × 6: 32 s → 1.1 s; 50,000 rows: 5.7 s; 10,000 Word paragraphs: 16 s → 0.24 s. Corpus bytes unchanged. Guards in `tests/test_performance.py`, numbers from `python -m tools.bench`. Single-cell reads (`value()`) were 40 ms each on 50,000 rows (a scan per call, the whole shared-strings table decoded each time); now 0.02 ms: a cached row index checked around each lookup (the cached neighbours must still be neighbours, so rows added, removed or renumbered on the tree rebuild it), and cached shared strings and date styles with cheap fingerprints (counting lxml children costs 0.4 ms at that size, so the big caches use their last child).

Also done in this round:
- `need()`, for required structure in files from other producers: a clear error instead of `NoneType`.
- Semantic rules for constraints the XSD can't express (`oox/semantics.py`; first: Excel row and cell order).
- `doc.validate()` across all parts.

**Exit criteria for 0.1.0:** Office 2016 run complete; observe results recorded; TestPyPI upload.

## Phase 2: Word in depth ✅ (complete on Office 2007, 2026-11-14) → **0.2.0**
Progress:
- ✅ **Step 1, default parts.** `WordprocessingDocument.new()` now sets a document up the way Word does: styles (Normal, Title, Heading 1–3, List Paragraph, Hyperlink, Table Grid, plus the defaults), settings, a font table, the theme, and an A4/Letter section. All of it validates against Office 2007. `new(minimal=True)` keeps the bare package.
- ✅ `add_heading`, `add_paragraph(style=...)` and `style_id(name)`.
- ✅ **Document properties** for all formats (`doc.properties.title = ...`). Reading also accepts LibreOffice's variant relationship type, which the LibreOffice run uncovered.
- Office corpus: `word-10-defaults`, `excel-06-properties`, `ppt-06-properties`, with language-independent probes (outline levels, built-in properties).
- ✅ **Step 2, lists.** `doc.new_list("bullet" | "decimal" | "legal", start=)`, `handle.add(text, level)` and `handle.restart()`, using Word's own level definitions.
  - `doc.list_labels()` computes the labels Word displays, from direct and style-based numbering, level overrides and restarts.
  - Counters are shared per abstractNum, as in Word. The LibreOffice cross-check exposed this when my first model kept counters per instance.
  - Office corpus: `word-11-lists`, with 13 probes on Word's own `ListString`.
- ✅ **Step 3, document structure** (one batch):
  - **Sections:** `doc.sections`, `doc.add_section(start=, landscape=, page=)`; per section: orientation, page size, margins, `text_width`, start type, `different_first_page`. A new section links to the previous section's headers and footers and drops "different first page", which would otherwise repeat a cover header.
  - **Headers and footers:** `section.header_text(text, kind="default" | "first" | "even", align=)` and `footer_text`; `{PAGE}`, `{NUMPAGES}`, `{SECTIONPAGES}` become fields. `header(kind)` returns the part's root for full control. A header part shared between sections is split when one of them is written. `linked_to_previous()` and `shown()` follow Word's inheritance.
  - **Links:** `add_hyperlink(p, text, url= | anchor=)`, `add_bookmark(name, p)`; readers `hyperlinks()`, `bookmarks()` (Word's hidden `_Toc…` ones excluded).
  - **Notes and comments:** `add_footnote`, `add_endnote` (parts created with Word's separator notes and registered in settings), `add_comment(p, text, author=, date=)`; readers `footnotes()`, `endnotes()`, `comments()` (with the commented text).
  - **Table of contents:** `add_toc(levels=(1, 3), before=)` is a real TOC field pre-filled with the headings; Word adds page numbers on update. `outline_level(p)` follows the style chain.
  - **Templates:** `replace_text({"{{name}}": "Ada"})` matches across runs in body, headers, footers, notes and comments; the replacement takes the formatting of the run where the match starts.
  - Styles these features need (Header, Footer, Hyperlink, Footnote Text, TOC 1–9, Heading 4–9 …) are added on first use, so existing documents and corpus bytes don't change. Also `add_page_break()`.
  - Office corpus: `word-12-sections`, `word-13-links-notes`, `word-14-toc-template`, with new probe kinds (counts, orientation, header/footer text and fields, hyperlink targets, note and comment text). LibreOffice: 31/31.
  - Office 2007 (2026-11-14): all 15 Word cases ok, including these three; every probe matched Word's own object model.
- ✅ **Step 4, tables, floating pictures, tracked changes** (one batch):
  - **Tables:** `doc.add_table(rows, header_rows=, widths=, style=, align=)` returns a `Table`: `merge(r1, c1, r2, c2)` (horizontal, vertical, rectangular; text kept as Word does), `format_cell` (shading, vertical and horizontal alignment, bold), `set_widths` (fixed layout), `set_header_rows`, `set_row_height`, `add_row`, `set_text`. Grid coordinates don't shift after merges. `doc.tables()` and `Table.values()` read any document's tables, including `gridBefore` rows. Consecutive tables get a separating paragraph, and on save a body or cell never ends in a table.
  - **Floating pictures:** `doc.add_floating_picture(image, h_relative=, v_relative=, x=, y=, h_align=, v_align=, wrap=)` with square, tight, through, top-and-bottom, in-front and behind wrapping. `doc.pictures()` lists inline and floating pictures with their wrapping.
  - **Tracked changes:** `suggest_insert`, `suggest_delete`, `suggest_replace` (across runs, keeping formatting), `doc.track_changes`, `doc.revisions()`, `accept_all()`, `reject_all()`. Accept and reject also cover Word's moves, paragraph-mark and table-row revisions, and formatting changes (`rPrChange`, `pPrChange` …), in every story.
  - Core fix found on the way: `add(..., replace=True)` now replaces where the schema allows repeats too (`w:trPr` is an unbounded choice, so it used to append a second `w:tblHeader`).
  - Office corpus: `word-15-tables`, `word-16-floating-pictures`, `word-17-tracked-changes`, `word-18-tracked-changes-reject`. The last two accept or reject all changes inside Word and compare the text. LibreOffice: 35/35.
  - Office 2007 (2026-11-14): all four cases ok after two probe expectations were corrected (Word doesn't count rows made only of vertical-merge continuations; Word 2007 reports behind and in-front wrapping as the same value).
- ✅ **Step 5, formatting, fields, forms, properties, images** (one batch):
  - **Formatting:** `doc.add_run(p, text, bold=, italic=, underline=, strike=, size=, font=, color=, highlight=, superscript=, subscript=, caps=, small_caps=, style=)`, `doc.format_paragraph(p, align=, space_before=, space_after=, line_spacing=, indent_left=, indent_right=, first_line=, keep_with_next=, …)`, `doc.add_style(name, kind=, base=, **format)`. `doc.effective_format(p, run)` resolves what Word shows: defaults, style chains, character style, direct formatting, theme fonts. (Superseded in 0.6.0 by `doc.formatting()`: table styles, numbering, toggles, points throughout.)
  - **Fields:** `add_caption` (numbered SEQ, bookmarked label and number), `add_cross_reference` (text, page estimate, above/below), `add_field` (TITLE, DOCPROPERTY …), `fields()`, `update_fields()` (SEQ with restart/repeat/hidden/formats, REF, PAGEREF estimate, property fields, dates with Word date pictures, case switches). Field codes are tokenized Word's way, including LibreOffice's `\@"d\ MMMM"`.
  - **Forms:** content controls (text, rich text, drop-down, combo, date, check box) with `content_controls()` and `fill_controls({tag: value})`, inline and block-level.
  - **Properties:** `extended_properties` (`app.xml`: company, manager, template …) and `custom_properties` (a typed mapping) for all three formats.
  - **Images:** TIFF, EMF, WMF and SVG (with a raster `fallback=`, in Word, Excel and PowerPoint); vector formats get their physical size from the file.
  - Office corpus: `word-19` … `word-22`, `excel-07`, `ppt-07`. `word-20` makes Word recompute every field and compares the results with oox's. LibreOffice: 41/41. Finding: LibreOffice drops `\* Upper` switches from DOCPROPERTY fields on export.
  - Office 2007 (2026-11-14): all six cases ok, including Word's own recompute of every field matching oox's results. **Word work complete.**
- ✅ **Step 6, text boxes and protection** (2026-11-15, after the gap review):
  - **Text boxes** (`textboxes.py`): `doc.add_text_box(text or [paragraphs, tables], width=, height=, x=, y=, h_align=, v_align=, h_relative=, v_relative=, wrap=, fill=, line=, line_width=, paragraph=)` writes what Word 2010+ writes: an `mc:AlternateContent` whose choice is a DrawingML shape (`wps:wsp`, `txBox="1"`) and whose fallback is the same box in VML (`v:shape` + `v:textbox`, shape type 202), which is what Word 2007 reads. `text_boxes()` counts each box once (also VML-only boxes from old files); `set_text_box_text(i, …)` edits both copies.
  - **Word protection** (`protection.py`): `doc.protect("readOnly" | "comments" | "trackedChanges" | "forms", password=, algorithm=, salt=, formatting=)`, `protection()`, `check_protection_password()`, `unprotect()`. The password hash is Word's: the Word 97 32-bit key (ECMA-376 Part 4 §14.8.1), reversed as uppercase hex, salted SHA-1 (default) … SHA-512, 100,000 iterations, written in the attributes Word 2007 reads (`w:cryptAlgorithmSid` …). Checked against the standard's example ("Example" → `0x64CEED7E`) and a hash Word 2019 wrote; the reader also checks Word 2010's `algorithmName`/`hashValue` set.
  - **Excel workbook protection:** `book.protect_workbook(password=, structure=, windows=)`, `workbook_protection()`, `check_workbook_password()` (Excel 2007's 16-bit hash and Excel 2010's salted SHA-512), `unprotect_workbook()`.
  - Office corpus: `word-24-text-boxes` (text through `Shapes(i).TextFrame`), `word-25-protection-read-only`, `word-26-protection-tracked-sha512`, `word-27-protection-comments` (`ProtectionType`; then `Unprotect` must refuse a wrong password and accept the right one), `excel-20-workbook-protection` (`ProtectStructure`, `Workbook.Unprotect`). LibreOffice: all ok; LibreOffice 24.2 doesn't import tracked-changes protection or workbook structure protection, so those are not judged there (the runner checks the original file to tell a drop from a bug). **Office 2007 (2026-11-15): every probe ok**, including the SHA-512 password and removing protection from a document opened read-only; Word 2007 reads the VML copies as text boxes. Word's extra "password is incorrect" dialog for the wrong attempt is now recorded as expected by the checker.
- Word leftovers: TOC refresh with page numbers needs layout, so it stays Word's job.

1. **Part conveniences** on the generic `add_part` (done): hyperlinks (external relationships), headers/footers, footnotes, comments, numbering.
2. ✅ **Default parts / template** (step 1; `app.xml` in step 5): styles, settings, fontTable, theme, `docProps/core.xml` + `app.xml` (Dublin Core, modeled by hand), mirroring odfpy's default styles.
3. **teletype extensions** (the basics are done; page-number fields and footnote references exist as document methods since step 3): simple formatting markers.
4. ✅ **Template workflows:** find/replace placeholders across run boundaries, the classic docx pain point (step 3).
5. **Interop corpus in CI:** python-docx / python-docx-ng test files (MIT), LibreOffice output, the Open XML SDK test assets (MIT), plus Word-authored samples you provide. Gate: zero false positives. `openxml-audit` (a pure-Python port of the SDK validator) runs as an external oracle on everything we generate.
6. ✅ **More image formats:** TIFF, EMF, WMF and SVG with a caller-supplied raster fallback (`asvg:svgBlip`), step 5; parsers written from the specifications. Floating (anchored) pictures in Word: ✅ step 4.

**Deliverable:** realistic reports (styles, lists, tables, images, headers, footers) generated end to end, with a sample gallery.

## Phase 3: Full validation and the second frontend → **0.3.0** (2–3 sessions; originally 3–4 weeks)
1. ✅ **Content-model automata** (2026-11-15): the generator emits particle trees; `oox.contentmodel` compiles them into Thompson NFAs; `validate(level="full")` checks choice exclusivity, required alternatives, counts and unfinished groups. No false positives on the corpus, LibreOffice re-saves or the python-docx/pptx templates.
2. ✅ **CLI** (2026-11-15): `python -m oox validate FILE [--target Office2007] [--full] [--json]`, `info`, `text`, `downgrade`. Exit status 0/1/2. Found on its first run: the SDK closes `w:characterSet` to a list without windows-1252, although ISO/IEC 29500-1 §17.8.3.2 types it `ST_String` (now an open enum).
3. **XSD frontend** (ECMA-376 Transitional + Strict, plus the MS extension XSDs) into the same IR, with an **IR-diff tool** that compares it to the SDK frontend. Every deviation gets triaged and documented. This also fills gaps the SDK has (Visio, Strict). **Started 2026-11-15** (`docs/XSD_DIFF.md`): `tools/gen/xsd_frontend.py` and `tools/gen/irdiff.py` compare the SDK data with the ISO/IEC 29500-4:2016 Transitional schemas (plus Microsoft's Word extension XSDs): 2,637 type pairs, findings by category and direction (SDK stricter: possible false positives in oox; XSD stricter: leniency), full list in `docs/reports/xsd-diff.md`. It rediscovers both deviations Office checks had already found (`word-94`'s values, `overrides.py`'s unordered fonts). Three new observe cases (`word-96`, `excel-93`, `excel-94`) put the most plausible SDK-stricter findings before Office. Not wired into the emitter: the generated code stays the SDK's. Strict schemas still to compare.
4. Differential testing of our full validator against openxml-audit / the .NET SDK over the corpus.
5. ✅ **Interop with the other Python libraries** (2026-11-15): `python -m tools.writers OUT` writes feature-rich files with python-docx, openpyxl, XlsxWriter and python-pptx (styles, tables, merges, charts including XlsxWriter combinations and secondary axes, chart sheets, comments, links, validation, pictures, notes, groups); all pass the interop check with zero validation findings at both levels. The other direction: every corpus case opens in python-docx, openpyxl and python-pptx, and tests compare what they read with what oox wrote. Found: workbook readers failed on chart sheets (now `sheet_names` lists worksheets, `tab_names` every tab); openpyxl writes charts without cached values (oox now reads them from the cells).
6. ✅ **Interop check** (2026-11-15): `python -m tools.interop --build OUT` converts LibreOffice's own templates (and odfpy's test documents) to OOXML; `python -m tools.interop DIR` opens each file, saves it unchanged (every part must come back byte-identical), validates at both levels, runs every reader, edits and saves again. First run, 149 files (99 docx, 24 xlsx, 26 pptx, plus a Word-authored docx with math): all pass. Validation findings are real LibreOffice export problems (duplicate `w:i`/`w:sz` in one `w:rPr`, `w:pStyle` before `w:numFmt`, `p:bgPr` without a fill) plus four values ISO Transitional allows and the SDK doesn't (`w:lvlJc`/table `w:jc` "start"/"end", `w:cantSplit`/`w:tblHeader` "true"). Two new observe cases ask Office 2007 about both kinds (`word-94`, `word-95`) before oox relaxes anything.

**Deliverable:** a validator CLI good enough to use in other people's CI.

## Phase 4: Excel and PowerPoint depth → **0.4.0 / 0.5.0** (in progress)
The document classes, skeletons and pictures are done (Phase 0).
- ✅ **Excel step 1, styles and layout** (one batch):
  - **Stylesheet manager** (`xlstyles.py`): `book.cell_style(number_format=, bold=, italic=, underline=, size=, font=, color=, fill=, border=, border_color=, align=, valign=, wrap=, indent=, rotation=)` returns a shared index; equal styles share one entry, and entries in an Excel-authored stylesheet are reused (structural comparison). The stylesheet is created on first use with Excel's defaults. `set_value(..., style=)`, `append_row(..., style=)`, `style_range`, `cell_format`.
  - **Dates:** dates, datetimes and times written with a date format; `value()`/`values()`/`cells()` return `datetime`/`time` for date-formatted cells in any workbook (`dates=False` for serials).
  - **Display text:** `display_text()` renders the number format the way Excel's `Range.Text` does (sections, grouping, percent, scientific, literals, currency, text sections, dates and times, Excel's 11-character General, half-away-from-zero rounding).
  - **Layout** (`xlsheet.py`): column widths in characters as Excel's dialog shows them, row heights, merged cells, frozen panes, autofilter (with Excel's hidden `_FilterDatabase` name), defined names (sheet names quoted), data validation (lists, ranges, whole/decimal bounds), each with a reader.
  - Office corpus: `excel-08-styles-formats` (19 `Range.Text` probes), `excel-09-layout`, `excel-10-style-existing`. LibreOffice: 44/44 (oox renders LibreOffice's re-saved format codes, e.g. `yyyy\-mm\-dd`).
  - First Office 2007 run (Greek regional settings) found that built-in format ids 14–22 and 37–40 render regionally: oox now maps only locale-invariant codes to built-in ids. Added `autofit_columns()` (dates showed `####` in default-width columns) and `excel-11-names-autofit` (English names on any locale via `[$-409]`).
  - **Complete on Office 2007** (Greek regional settings) after three checker fixes for regional settings.
- ✅ **Excel step 2** (one uninterrupted batch, 2026-11-15):
  - **Conditional formatting** (`xlcond.py`): cell value (all operators), formula, 2/3-colour scales, data bars, icon sets, top/bottom (n or %), above/below average, duplicates/unique, text (contains, not contains, begins, ends), blanks; highlight styles are shared differential formats (dxfs); `conditional_formats()` reads any workbook's rules.
  - **Tables** (`xltables.py`): `add_table` with header names taken from the first row (empty or repeated ones fixed as Excel requires), styles, banding, a totals row (labels and `SUBTOTAL` formulas whose results oox computes and caches), workbook-wide unique names, overlap checks; structured references in formulas. `Formula(expr, cached)` stores a result with any formula.
  - **Comments** (`xlnotes.py`): a comments part plus the legacy VML note boxes Excel needs, edited as text (Excel's VML isn't always well-formed), hidden or always visible, per-sheet shape-id blocks. **Hyperlinks**: external (with tooltips) and inside the workbook. Readers for both.
  - **Shared strings**: `share_strings()` moves inline strings into the shared-strings table, one entry per distinct string.
  - **Print setup and sheet options** (`xlprint.py`): orientation, paper, fit to pages, scale, margins, print area, repeated rows/columns, header/footer with `{page}`-style fields, gridlines, centring; tab colour, hidden and very hidden sheets (the active tab moves to a visible one), zoom, on-screen gridlines; sheet protection with Excel's legacy password hash (`secret` → `DAA7`, the value LibreOffice also writes).
  - Core fix found on the way: unqualified attributes named like odfpy keywords (`cfRule/@text`) are reachable as `x_text=`, as ODFPY_COMPAT always said.
  - Office corpus: `excel-12-conditional`, `excel-13-tables`, `excel-14-comments-links`, `excel-15-print-options`. LibreOffice: 49/49, with "not judged" for what LibreOffice doesn't keep (table styles, comment authors, tab colours).
- ✅ **Charts** (one unattended batch, 2026-11-15), shared by all three formats:
  - **Builder and reader** (`charts.py`): column, bar, line, area (clustered/standard, stacked, 100%), pie, doughnut, scatter (markers, lines, smooth); titles, axis titles, legend position, data labels (value, percent, category, series), value number format, axis min/max, gridlines, colours per series or per slice. Colours are always explicit (default palette = the theme's accents): LibreOffice draws "automatic" series without a chart style as empty. No title sets `autoTitleDeleted`, so Excel doesn't invent one from a single series' name; `roundedCorners` is written off (Excel's default is on). `chart_info()` reads any chart part (oox's, Excel's, LibreOffice's) and reports the type in Excel's `XlChartType` vocabulary, so the Office check compares the two directly.
  - **Excel** (`xlcharts.py`): `add_chart(sheet, kind, "A1:D5", anchor=, series_in=)` with references to the cells (sheet names quoted) and cached values; two-cell anchors computed from the sheet's column widths and row heights; `charts(sheet)`.
  - **Word and PowerPoint:** `add_chart` from Python data (`{name: values}`) as an inline chart or a slide graphic frame, with an embedded workbook built by oox holding the data, so "Edit Data" works; `charts()` readers. `add_part(folder=)` puts parts where Office does (`xl/charts`, `word/embeddings`, ...).
  - Office corpus: `excel-16-charts` (5 charts, 21 probes through Excel's `Chart` object, plus chart images via `Chart.Export`), `word-23-charts`, `ppt-08-charts` (with slide images). Word/PowerPoint 2007 have a chart object model only from SP2; without it, chart details are not judged there. LibreOffice: 52/52.
  - Tutorial 9 (`examples/charts.html`).
- ✅ **Charts, step 2** (2026-11-15): combination charts (`kinds={"Price": "line"}`; columns, lines and areas combine, bars with bars; areas drawn behind columns behind lines, `c:order` following the plots as Office and LibreOffice list them) and a secondary value axis (`secondary=`, `y2_title=`, `y2_number_format=`, `y2_min/max=`; `Series.kind/axis`), for all three formats; `chart_info` reports each series' type and axis and `y2_title`. **Chart sheets**: `add_chart_sheet(name, kind, sheet, range)`, `chart_sheet()`, `charts()` on a chart sheet. Office corpus: `excel-17-combo-charts` (series types and axis groups through COM), `excel-18-edit-openpyxl-file`, `excel-19-chart-sheet`. LibreOffice: all ok. **Complete on Office 2007** (2026-11-15): series types, axis groups and the secondary axis title confirmed through Excel's `Chart` object, after two checker fixes for dynamic COM.
1. **Excel, still to do:** pivot tables (not planned); parity check with odfpy's table API.
- ✅ **PowerPoint depth** (same unattended session, 2026-11-15; `slides.py`):
  - **Layouts:** `new()` builds Office's eleven standard layouts with Office's placeholder geometry (scaled to the slide size) and layout types (`secHead`, `twoObj`, `twoTxTwoObj`, `picTx`, …); `new(layouts="basic")` keeps the earlier three (the corpus builders use it, so existing files are byte-identical). `slide_layout(slide)` reader.
  - **Tables:** `add_table(slide, rows, ...)` in PowerPoint's built-in styles (by name or GUID), header/banding/first-column switches, column widths, merges as `"A4:B4"` ranges (`gridSpan`/`rowSpan` + `hMerge`/`vMerge`); `tables(slide)` reader.
  - **Shapes:** `add_shape` with preset geometries (rect, roundRect, ellipse, arrows, chevron, …), fill, line and centred text.
  - **Speaker notes:** `set_notes`/`notes`, creating the notes master (with its own theme) on first use and reusing an existing one.
  - **Picture placeholders:** `set_placeholder_picture` replaces a picture or content placeholder with a `p:pic` that keeps the placeholder link, cropped to fill (`srcRect`) from the inherited position, as PowerPoint does.
  - `shapes(slide)` lists what any slide holds (placeholders, text boxes, shapes, pictures, tables, charts, groups) with text.
  - Office corpus: `ppt-09-layouts-tables-notes` (layout names, slide layouts, table cells, notes, auto-shape types, filled placeholder, slide images). LibreOffice: 53/53 (it drops table styles and saves a filled placeholder as a plain picture: not judged there).
  - Tutorial 10 (`examples/slides.html`).
- ✅ **PowerPoint, step 2** (2026-11-15): **editing existing tables** (`deck.table(slide, which)` → `SlideTable`: `values`, `set_text` keeping the cell's run and paragraph properties, `add_row`/`add_column` copying a neighbour's formatting, `delete_row`/`delete_column` shrinking merges or moving a merge's origin, `merge`/`unmerge` as PowerPoint writes them, `set_widths`, `set_style`; the frame follows the grid). **Click links** (`set_link(slide, shape, target, text=)` on a shape or on text, split out of its runs; URLs, mail, slides, show jumps; `links()`). **Sections** (`set_sections`, `sections`; `p14:sectionLst`, ignored by 2007). Fix found on the way: `table_info` counted PowerPoint's repeated `gridSpan` below a merge as a second merge. Office corpus: `ppt-10-links-table-edit` (Slide.Hyperlinks, ActionSettings, table cells), `ppt-11-edit-python-pptx-file`, `ppt-12-sections` (SectionProperties, 2010+). LibreOffice: all ok (sections not judged: LibreOffice drops them). **Complete on Office 2007** (2026-11-15), sections not judged there; the ppt-09 picture-placeholder fix confirmed. Links are read through `ActionSettings`: PowerPoint 2007's `Slide.Hyperlinks` misses text-run links.
2. **PowerPoint, still to do:** handout master, slide transitions and animations (not planned).
3. ~~**Shared DrawingML:** charts (`c:` namespace) from Python data, usable from all three formats.~~ ✅ see above.

**Deliverable:** report-grade xlsx and pptx (styled tables, charts, placeholders) with a round-trip corpus per format.

## Phase 5: Conveniences, docs, 1.0 (ongoing)
- An optional thin convenience layer (`paragraph(text, style=…)`, `table(rows)`) on top of the factories. Never a parallel object model.
- ✅ **Packaging** (2026-11-16): version 0.5.0 (read from `oox.__version__`), PEP 639 licence metadata, classifiers, an `oox` command, a slim wheel (the package, stubs and `py.typed`; 1.1 MB) and an sdist that carries the tests, tools, examples and docs so its test suite runs (`MANIFEST.in`). Built offline and installed into a clean environment, where `tools/smoke_install.py` (three formats, every validation level, stubs, the command) passes. `.github/workflows/release.yml` builds once, tests the wheel on Python 3.10 and 3.14 (Linux and Windows) and the sdist's suite, and publishes with trusted publishing: TestPyPI by hand, PyPI on a `vX.Y.Z` tag that must equal the version. `RELEASING.md` has the one-time setup and the checklist. A PyPI search found no `ooxpy` or `oox`; confirm on pypi.org before the first upload.
- ✅ **API inventory and stability policy** (2026-11-16): `docs/STABILITY.md` defines the public API (exports, document classes, the objects they return, utility modules, factories) and the rules for changing it (semver from 1.0, deprecation for a minor release and six months, odfpy names permanent, relaxations only with Office evidence). `tests/api/public_api.txt` pins it: any change fails `tests/test_public_api.py` until the snapshot is updated with `python -m tools.api_surface --write`, so the diff is the API review. 115 public methods that had no docstring got one. `TargetError` is now exported.
- ✅ **Full documentation** (2026-11-16): an MkDocs (Material) site in `docs/`: home, getting started, concepts, guides for Word, Excel, PowerPoint, charts, existing files and validation, a cookbook of nine recipes, the command line, coming from odfpy, Office compatibility, limitations and FAQ, the stability policy, the changelog, and the design documents. The API reference (`docs/reference/`) is generated from the docstrings by `python -m tools.apidocs`, grouped by topic, with each method's keyword options and the details from the module it delegates to. `tests/test_docs.py` runs every Python block in the guides, checks that the reference is current and places every public method, and builds the site with `--strict`; CI has a docs job. Found while writing: `oox text` printed a text box's text three times; a `Section` taken before `add_section()` switched to the new section; `lenient()` still rejected invalid values. All three fixed, with tests.
- ✅ **Released 0.5.2 on PyPI** (2026-11-17) from `grhpas/ooxpy` with trusted publishing.
- ✅ **0.6.0, reading as Word shows it** (2026-11-17, after an external review of reading real files; ADR-003):
  - **Shared walker** (`oox.walk`): every reader sees the `mc:AlternateContent` branch Word shows and content inside content controls. Fixed double counting in fields, bookmarks, hyperlinks, revisions, controls, tables, charts, list labels and `oox info`; caption numbering in text boxes; sections in content controls; `pictures()` now lists pictures only (DrawingML and VML).
  - **Reading never writes**: `styles`/`settings`/`numbering` return None when missing (`ensure_*` create); a virtual default section; read-only `header()`/`footer()` with `ensure_header()`/`ensure_footer()`; `margins` complete. Enforced by `tests/test_reading_is_pure.py` over every public member.
  - **Formatting resolver** (`resolve.py`): `doc.formatting()` → `ParagraphFormat`/`RunFormat` through defaults, table styles with conditional formats, paragraph styles, numbering, character styles and direct formatting, with Word's toggle rules and MS-DOCX's table rule. Points everywhere; `LineSpacing`, `Border`, `Tab` shared by readers and writers; new properties (shading, letter spacing, hidden, language, borders, tabs, outline level, contextual spacing).
  - **Lists**: `ListItem`s, numbering styles (`numStyleLink`), `lvlRestart`, style-linked levels, more formats, readable bullets, picture bullets, all stories.
  - **Packages**: namespace constants and `qn` in `oox.ns`; relationships as a mapping; `related_part`, `blob`, named parts, `image_of`, a fast `part_of`.
  - Also: the guide "Reading a document as Word shows it"; a PyInstaller hook (entry point `pyinstaller40`); a LibreOffice-written .docx corpus with ground truth (`tools/locorpus.py`, `tests/data/libreoffice/`: sources we wrote, plus LibreOffice's own text export of its labels).
  - The Office corpus is byte-identical (Office results stay valid); LibreOffice 69/69; 525 tests.
- ✅ **Decisions on the API before 1.0** (2026-11-17, in 0.6.0; `docs/API_DECISIONS.md`): writers are verbs and readers nouns (`set_frozen_panes`/`frozen_panes`, `set_page_setup`/`page_setup`, `set_sheet_options`/`sheet_options`); `add_text_box` in PowerPoint as in Word; readers return records (`oox.records`, one class per concept across formats, field names = writer keywords, `as_dict()`); PowerPoint geometry as optional `x`/`y`/`width`/`height` keywords defaulting to the content area; a plain number is points everywhere, anything else a `Length`, with range checks that catch raw EMU/twips. The Office corpus stayed byte-identical; LibreOffice 69/69; 526 tests.
- Freeze the 1.0 API, with a regeneration cadence that follows SDK releases.

---

## Risks and mitigations
| Risk | Mitigation |
|---|---|
| SDK JSON format changes (it's not a public contract) | IR boundary; pinned revision; frontend tests on synthetic input; the XSD frontend as a fallback (Phase 3). |
| SDK data deviates from ECMA-376 | IR-diff against XSD; documented overrides file applied in the frontend. |
| Ambiguous elements built detached confuse users | Reorder and re-validate on attach (done); SDK aliases; docs; maybe a warning when a detached ambiguous element gets children. |
| Generated code size (about 4 MB) | Lazy per-namespace import (done). Could move to compressed data plus `__getattr__` factories if packaging size matters. |
| Name clash with existing packages (`docx`, `ooxml`) | Distribution name `ooxpy` (to be checked), import name `oox`. |
| Word accepts things the schema rejects, and the reverse | The corpus of real Word files is the arbiter; `lenient()` as an escape hatch. |
| Office is stricter than LibreOffice and the Python readers we test with | Phase 1 item 0: manual Office checks, then automation (e.g. a Windows CI runner with Office, or openxml-audit as a proxy). |
| Ambiguous chains defer validation (e.g. `x:sheetData → x:row → x:c` resolves only at `x:worksheet`) | Detached elements are still checked against the union of their models (enum values merged), so most mistakes fail fast; the rest fail on attach with a precise message. |

## Decisions
1. **Name:** distribution `ooxpy`, import `oox`. Agreed 2026-11-14 (PyPI availability still to check before the first upload).
2. **License:** MIT. Agreed.
3. **Python floor:** 3.10, raised to 3.12 if a feature needs it. Agreed. Typed kwargs turned out not to need PEP 692: they're generated `.pyi` stubs (2026-11-15). The Windows-side check scripts stay Python 3.8-compatible for Windows 7 machines.
4. **odfpy names:** kept permanently as a supported, tested compatibility surface. See ADR-002 in ARCHITECTURE.md.
5. **Office testing:** Windows machines with Office 2007 and 2016; a Mac mini without Office (LibreOffice and Apple apps for now). The showcase files opened cleanly. Excel's save prompt on close is expected (formulas are calculated on open). Caching formula results to avoid it would need a formula evaluator; not planned.
