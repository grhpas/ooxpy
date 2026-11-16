# Office compatibility

Schema validity isn't enough: what matters is that Office opens the file without a
repair prompt and shows what you wrote. ooxpy is checked against real Office with a
corpus of test files and a scripted checker that drives Office itself.

## How it's checked

The Office check kit (`tools/officecheck/` in the repository) generates a deterministic
corpus:

- **58 normal cases** (26 Word, 20 Excel, 12 PowerPoint), one or more per feature. Each
  carries **probes**: what Office must report through its own object model, such as a
  table's cell text, Word's list labels, a chart's type and axis titles, a conditional
  format's rule type, `Range.Text` for a number format, or the result of Word
  recomputing every field.
- **3 controls**: broken files Office must reject, which prove the checker notices
  prompts.
- **11 observe cases**: files that break the schema in one specific way, recorded
  without pass or fail, to learn how strict each Office version is.

`check_office.py` opens each file in Word, Excel or PowerPoint through COM, watches
for any dialog, runs the probes, saves chart and slide images, and writes a results
file; `report.py` merges runs from several machines into one table. The same probes
run through LibreOffice (`check_libreoffice.py`) on every CI build.

## Results

| | Office 2007 (12.0) | Office 2016 (16.0) | LibreOffice 24.2 |
|---|---|---|---|
| Word, 26 cases | all ok | all opened cleanly; 21 fully confirmed, 5 with probes still to run on an activated Office | all ok |
| Excel, 20 cases | all ok | all opened cleanly and every probe matched; one screenshot couldn't be taken (a checker problem, fixed in the kit) | all ok |
| PowerPoint, 12 cases | all ok (sections aren't shown by 2007) | all ok | all ok (sections aren't kept) |
| Controls | all detected | all detected | — |

Office 2007 was run with Greek regional settings, which also checks that number
formats, dates and probes don't assume an English locale. The Office 2016 machine
wasn't activated, so Word refused calls that change a document (accepting tracked
changes, removing protection) or read paragraph formatting; those five cases opened
without a prompt and are confirmed on Office 2007.

## What the observe cases taught

| Deviation | Office 2007 | Office 2016 | oox |
|---|---|---|---|
| paragraph properties out of order, unknown elements | opened | opened | still reported by `validate()`: other readers may refuse them |
| invalid attribute values | refused | refused | rejected while building |
| Excel rows out of order | refused | refused | a semantic rule |
| a required element missing (PowerPoint) | prompt, slide blanked | refused | rejected by `validate()` |
| an Office 2010 attribute in a 2007 namespace | refused | opened | `validate(target="Office2007")` reports it |
| values ISO/IEC 29500 allows and the SDK data doesn't (`w:jc="start"` …) | opened | opened | accepted |
| chart elements the SDK data requires and ISO makes optional | — | refused | kept required |

Relaxations are made only with this kind of evidence, and always at run time, so the
generated grammar stays identical to the pinned Open XML SDK data
([XSD cross-check](XSD_DIFF.md)).

## What isn't checked yet

- Office 2016 on an activated installation for the five Word cases above, and Microsoft
  365's current builds (the 2016 run reports version 16.0, which 2016, 2019 and 365
  share).
- Office for Mac.
- Strict-conformance files in Office (checked with LibreOffice only).
