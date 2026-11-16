# Office check kit

A repeatable check that files written by oox open cleanly in Microsoft Office
and LibreOffice, with their content intact.

| Script | Runs on | Needs |
|---|---|---|
| `generate.py` | any machine with oox | Python 3.10+, oox |
| `check_office.py` | Windows with Office 2007 or newer | Python 3.8+ and `pywin32`; no oox |
| `check_libreoffice.py` | Linux, macOS, Windows | Python 3.10+, oox, LibreOffice |
| `report.py` | anywhere | Python 3.8+ |

## 1. Generate the corpus (once per oox change)

```bash
python -m tools.officecheck.generate officecheck-corpus
```

This writes 72 files (Word, Excel and PowerPoint): 58 normal cases, 3 negative
controls and 11 observe cases. It also writes:
- `manifest.json`: hashes and probes
- `CHECKLIST.md` and `manual-results.csv`: for manual runs
- copies of `check_office.py` and `report.py`

Copy the whole folder to each test machine.

The output is deterministic, so results from different machines and dates
are comparable. Adding cases doesn't change existing files, so earlier results
stay valid.

### Negative controls

Three cases (`*-90-control-malformed`) are **broken on purpose**: one part is
truncated into malformed XML. A control passes (✅ detected) when either:
- Office complained (prompt or error), or
- Office repaired the file silently and content was lost: the probes check text that sits after the cut, so this shows up as a mismatch.

The checker turns alerts on explicitly in all three apps. PowerPoint's
automation default is "no alerts, take the default answer", which hid the
PowerPoint control's prompt on the first Office 2007 run.

An all-green column means something only if the controls are green too. If a
control shows "not detected (ok)", either the checker isn't seeing that Office
version's dialogs, or Office silently repaired the file. "timeout" usually
means a dialog appeared that the watchdog didn't recognise, so add its window
class to `DIALOG_CLASSES` in `check_office.py`.

### Observe cases

Eleven cases (`*-91…96-observe-*`) are well-formed XML that breaks the schema in
one specific way: properties out of order, an unknown element, invalid values,
rows out of order, an Office 2010 attribute Office 2007 can't know, a
missing required element, and two found in LibreOffice's output by
`tools/interop.py`: values ISO Transitional allows but the SDK rejects
(`word-94`: `lvlJc`/table `jc` "start", `cantSplit`/`tblHeader` "true") and
structure the schema forbids (`word-95`: `pStyle` before `numFmt`, duplicate run
properties), and three from the XSD cross-check (`docs/XSD_DIFF.md`): markup
ISO/IEC 29500-4 allows and the SDK data rejects (`excel-93`: a doughnut without
`holeSize`, `barDir` without `val`; `excel-94`: `pageSetup` paper width and
height; `word-96`: table band sizes in a body table). **They aren't judged** (👁 in the report). They
record how strict each Office version is, which tells us how serious each
kind of `validate()` finding is in practice. Run them with `--only observe`.

To run only the controls on a machine that already has results:
`py check_office.py . --only control`. The LibreOffice runner skips them,
because headless LibreOffice shows no dialogs.

## 2a. Manual run (works today, any Office version)

1. Open each file listed in `CHECKLIST.md` in the matching app.
2. Watch for repair or recovery prompts, errors, and missing content.
3. Copy `manual-results.csv` to `manual-results-<label>.csv` (e.g. `manual-results-office2007.csv`).
4. Fill in `status` (ok, prompt, error, mismatch) and `notes`.

## 2b. Scripted run (Windows)

```bat
py -m pip install pywin32
py check_office.py .                 :: all apps
py check_office.py . --apps excel    :: one app
py check_office.py . --only ppt-02   :: one case
py check_office.py . --only word-15,word-16   :: several cases
```

For each file, the script:
- starts a fresh Office instance
- opens the file read-only
- records any dialog: a watchdog thread captures its text and closes it
- runs the probes through Office itself: table and picture counts, **computed** cell values, slide texts, list labels, outline levels, section orientation, header and footer text and fields, hyperlink targets, footnote/endnote/comment text, and counts of sections, bookmarks, notes, comments and tables of contents; table rows, cells (after merges), header rows and cell text; floating shapes and their wrapping; tracked changes (type, author, text), and the text after Word accepts or rejects them all (this changes the open, read-only document, so it is always the last probe of its case); character and paragraph formatting as Word resolves it; field results, before and after Word recomputes every field (also a last probe); content controls by tag (text, type, placeholder); custom document properties; in Excel, the text a cell displays (`Range.Text`, i.e. the number format applied), cell font, fill, alignment and number format, column widths, row heights, merge areas, frozen panes, autofilter, defined names and validation lists, conditional-format rule types per cell, tables (names, ranges, totals, columns, styles), comments (author and text) and hyperlinks, page setup (orientation, paper, fit, print area and titles, header and footer), tab colour, hidden state and protection; charts in all three apps through Office's own `Chart` object (type as `XlChartType`, title, legend position, axis titles including the secondary axis, series names, values, categories, data labels, each series' type and axis group), Excel tab order and chart sheets (`Workbook.Charts`); Word text boxes (`Shapes(i).TextFrame`), Word protection (`ProtectionType`) and Excel workbook protection (`ProtectStructure`, `ProtectWindows`), then `Unprotect` with a wrong password (must be refused) and the right one (must be accepted; a last probe, since it changes the document; Word also shows a "password is incorrect" box for the wrong one, which the report lists as an expected dialog, not a prompt); in PowerPoint, the layout names (`CustomLayouts`), each slide's layout, table rows, columns and cell text, speaker notes (`NotesPage`), auto-shape types, filled picture placeholders, hyperlinks (`Slide.Hyperlinks`), click actions (`ActionSettings`) and sections (`SectionProperties`, 2010+).

An adapter returns `NotJudged("reason")` for something its application can't answer; the probe then counts as ok and the report shows the reason. The LibreOffice runner uses it for what LibreOffice doesn't keep on export: table styles, comment authors, tab colours, slide sections, tracked-changes protection and workbook structure protection (for the last two it checks the original file, so a protection oox failed to write still fails). Word and PowerPoint 2007 got a chart object model only with SP2: without it, their chart details are not judged (the error is in the note); chart counts are judged everywhere.
- quits Office

**Regional settings.** Excel shows numbers, dates and lists in the machine's
regional format, and through pywin32 it also answers format codes and list
separators in it. The Excel probes swap the regional decimal, thousands and
list separators back to ".", "," and ",", and probes that depend on month
or day names (marked `names`) are judged only on English Excel. Cases meant
to test names everywhere use a `[$-409]` locale prefix in the format code.

Results go to `results-<host>-<apps>-<date-time>.json`, one file per run, so
a rerun of a few cases never overwrites earlier results. `report.py` merges all
of them, and for each case the newest result wins. (Before 2026-11-14 the
name had no date and time, so each rerun replaced the previous file.)

Some cases also save pictures into `screens/` in the corpus folder, listed
under *Screenshots* in the report: `excel.chart_image` and `ppt.slide_image`
(Office's own `Chart.Export` / `Slide.Export`, sharp), and `excel.screenshot`,
a range as Excel displays it (`excel-12-conditional`). The range picture uses `Range.CopyPicture`, a temporary chart
and `Chart.Export`, so it needs no extra packages, but it does use the
clipboard while it runs. It captures as a metafile at 2x (the note says
`picture, x2`); the bitmap fallback renders text poorly. Send the PNGs along with the report.

Before running:
- **Close Word, Excel and PowerPoint.** PowerPoint allows only one instance.
- **Open each app once by hand first**, so first-run, activation and privacy dialogs don't count as prompts.
- **Copy the corpus via USB or a network share, not a browser or email.** Downloaded files carry the "mark of the web", and Office 2010+ then opens them in Protected View, which blocks COM. If that happens, run `Unblock-File` in PowerShell, or right-click the file, open Properties and tick *Unblock*. Office 2007 has no Protected View.
- **Use Python 3.8 on Windows 7.** It's the last Python version that runs there.
- **Office not activated?** Word and Excel then show "Sign in to set up Office" and run in reduced functionality mode, refusing calls such as `Unprotect`, `AcceptAll` and paragraph formatting. The checker records those dialogs and refusals apart (🔒 in the report: the file opened cleanly, some probes couldn't run) and they don't count, but the result is only complete on an activated Office.
- **Several Pythons installed?** `py` starts the newest one (or the one set in `py.ini`), `py -3.14` a specific one; `py -0p` lists them. Each needs its own `pywin32`. Every results file records the interpreter, its pywin32 build and whether COM was bound early (makepy) or dynamically, and the report lists them under *Runs*.

First run: Office 2007, 2026-11-14, with all 18 cases ok
(`reports/2026-11-14-office2007.md`). The negative controls were added
afterwards, to confirm that the dialog detection works on each Office version.
If it's ever unsure, the manual run is the reference.

## 2c. LibreOffice run (the Mac mini, Linux)

```bash
python -m tools.officecheck.check_libreoffice officecheck-corpus   # finds soffice, incl. macOS app
```

LibreOffice converts each file, and the probes run on what it wrote back. This
uses an isolated LibreOffice profile, so it doesn't disturb a running instance.

## 3. Report

Collect every `results-*.json` and filled-in `manual-results-*.csv` into the
corpus folder, then run:

```bash
python report.py officecheck-corpus     # -> REPORT.md
```

You get one column per machine and Office version (✅ ok · ⚠️ prompt · ❌
error or mismatch · ⏱️ timeout), plus details: dialog texts, failed probes
with expected and actual values, and notes.

## CI

`.github/workflows/office-check.yml` runs the corpus on a self-hosted Windows
runner with Office. Register a runner on each Office machine with a label such
as `office2007` or `office2016`, then start the workflow manually and choose the
label. The results and `REPORT.md` are uploaded as artifacts.

## Adding a case

Add a builder function and a `Case(...)` entry in `generate.py`, with probes for
what must survive. Probe kinds are implemented twice: through COM in
`check_office.py` (keep it Python 3.8) and through the oox readers in
`check_libreoffice.py`. Never change an existing builder's output: add a new case
instead, so earlier reports stay comparable (the manifest hashes show any change).
Run `check_libreoffice` locally before taking it to the Office machines. A bug found in Office becomes a new case.
