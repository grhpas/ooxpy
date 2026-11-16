# Office check report

Corpus: oox 0.0.1, Open XML SDK data @ 431ab05cf160248cc3885a4a766026d4f8243792 2026-08-18.

✅ ok · ⚠️ prompt (repair/recovery dialog) · ❌ error or content mismatch · ⏱️ timeout · · not run

*Control* cases are broken on purpose: ✅ means the checker noticed, either because Office complained or because it repaired silently and content went missing.

*Observe* cases (👁) break the schema in one specific way and are not judged: they record how strict each Office version is.

| Case | gregory_citadel (Office 2016+) |
|---|---|
| `word-01-minimal` | ✅ |
| `word-02-formatting` | ✅ |
| `word-03-table` | ⚠️ |
| `word-04-pictures` | ⚠️ |
| `word-05-styles` | ⚠️ |
| `word-06-w14` | ⚠️ |
| `word-07-odfpy-passwd` | ⚠️ |
| `word-08-edit-word-file` | ⚠️ |
| `excel-01-types` | ⚠️ |
| `excel-02-formulas` | ⚠️ |
| `excel-03-sheets` | ✅ |
| `excel-04-pictures` | ⚠️ |
| `excel-05-edit-existing` | ⚠️ |
| `ppt-01-minimal` | ✅ |
| `ppt-02-placeholders` | ✅ |
| `ppt-03-textboxes-pictures` | ✅ |
| `ppt-04-4x3` | ✅ |
| `ppt-05-edit-powerpoint-file` | ✅ |
| `word-90-control-malformed` *(control)* | ✅ detected |
| `excel-90-control-malformed` *(control)* | ✅ detected |
| `word-10-defaults` | ✅ |
| `word-11-lists` | ✅ |
| `word-12-sections` | ⚠️ |
| `word-13-links-notes` | ⚠️ |
| `word-14-toc-template` | ⚠️ |
| `word-15-tables` | ✅ |
| `word-16-floating-pictures` | ⚠️ |
| `word-17-tracked-changes` | ⚠️ |
| `word-18-tracked-changes-reject` | ⚠️ |
| `word-19-formatting-styles` | ⚠️ |
| `word-20-fields-captions` | ⚠️ |
| `word-21-content-controls` | ⚠️ |
| `word-22-more-images` | ⚠️ |
| `excel-07-properties-images` | ⚠️ |
| `ppt-07-properties-images` | ✅ |
| `excel-08-styles-formats` | ⚠️ |
| `excel-09-layout` | ⚠️ |
| `excel-10-style-existing` | ⚠️ |
| `excel-11-names-autofit` | ✅ |
| `excel-12-conditional` | ⚠️ |
| `excel-13-tables` | ⚠️ |
| `excel-14-comments-links` | ✅ |
| `excel-15-print-options` | ⚠️ |
| `excel-16-charts` | ⚠️ |
| `excel-17-combo-charts` | ⚠️ |
| `excel-20-workbook-protection` | ⚠️ |
| `excel-19-chart-sheet` | ⚠️ |
| `excel-18-edit-openpyxl-file` | ✅ |
| `word-24-text-boxes` | ⚠️ |
| `word-25-protection-read-only` | ⚠️ |
| `word-26-protection-tracked-sha512` | ⚠️ |
| `word-27-protection-comments` | ✅ |
| `word-23-charts` | ⚠️ |
| `ppt-08-charts` | ✅ |
| `ppt-09-layouts-tables-notes` | ✅ |
| `ppt-10-links-table-edit` | ✅ |
| `ppt-12-sections` | ✅ |
| `ppt-11-edit-python-pptx-file` | ✅ |
| `excel-06-properties` | ⚠️ |
| `ppt-06-properties` | ✅ |
| `word-91-observe-order` *(observe)* | 👁 prompt |
| `word-92-observe-unknown-element` *(observe)* | 👁 prompt |
| `word-93-observe-bad-values` *(observe)* | 👁 prompt |
| `word-94-observe-sdk-values` *(observe)* | 👁 prompt |
| `word-96-observe-iso-band-size` *(observe)* | 👁 prompt |
| `word-95-observe-libreoffice-structure` *(observe)* | 👁 prompt |
| `excel-91-observe-row-order` *(observe)* | 👁 error |
| `excel-93-observe-iso-optional` *(observe)* | 👁 error |
| `excel-94-observe-iso-paper-size` *(observe)* | 👁 prompt |
| `excel-92-observe-newer-attribute` *(observe)* | 👁 prompt |
| `ppt-91-observe-missing-required` *(observe)* | 👁 error |
| `ppt-90-control-malformed` *(control)* | ✅ detected |

## Details

- **word-03-table** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.tables`: expected `1`, got `None` (com_error(-2147418111, 'Call was rejected by callee.', None, None))
  - probe `word.text_contains`: expected `None`, got `None` (com_error(-2147418111, 'Call was rejected by callee.', None, None))
- **word-04-pictures** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-05-styles** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-06-w14** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-07-odfpy-passwd** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-08-edit-word-file** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-01-types** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-02-formulas** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-04-pictures** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-05-edit-existing** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-90-control-malformed** on gregory_citadel (Office 2016+): prompt
  - dialog `Microsoft Word`: OK / Word experienced an error trying to open the file.
Try these suggestions.
* Check the file permissions for the document or drive.
* Make sure there is sufficient free memory and disk space.
* Open the file with the Text Recovery converter. (C:\...\word-90-control-malformed.docx) / 200494
  - error: `com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'Word experienced an error trying to open the file.\nTry these suggestions.\n* Check the file permissions for the document or drive.\n* Make sure there is sufficient free memory and disk space.\n* Open the file with the Text Recover`
- **word-12-sections** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-13-links-notes** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-14-toc-template** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-16-floating-pictures** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.shape_wraps`: expected `['behind', 'square', 'tight']`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The Type method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
- **word-17-tracked-changes** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.text_after`: expected `None`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
- **word-18-tracked-changes-reject** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.text_after`: expected `None`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
- **word-19-formatting-styles** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.paragraph`: expected `6`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
  - probe `word.paragraph`: expected `36`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
  - probe `word.paragraph`: expected `'justify'`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
  - probe `word.paragraph`: expected `12`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
  - probe `word.paragraph`: expected `18`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
  - probe `word.paragraph`: expected `'right'`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
  - probe `word.paragraph`: expected `36`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'This method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None))
- **word-20-fields-captions** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-21-content-controls** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-22-more-images** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-07-properties-images** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-08-styles-formats** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-09-layout** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-10-style-existing** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-12-conditional** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `excel.cf_types` sheet=Rules ref=B2: ok, got `['cellIs', 'cellIs', 'expression', 'top10']`
    - route 1 (sheet rules by AppliesTo) ok
  - probe `excel.cf_types` sheet=Rules ref=C2: ok, got `['colorScale', 'expression']`
    - route 1 (sheet rules by AppliesTo) ok
  - probe `excel.cf_types` sheet=Rules ref=D2: ok, got `['dataBar', 'expression']`
    - route 1 (sheet rules by AppliesTo) ok
  - probe `excel.cf_types` sheet=Rules ref=E2: ok, got `['expression', 'iconSet']`
    - route 1 (sheet rules by AppliesTo) ok
  - probe `excel.cf_types` sheet=Rules ref=A3: ok, got `['containsText', 'duplicateValues', 'expression']`
    - route 1 (sheet rules by AppliesTo) ok
  - probe `excel.screenshot` sheet=Rules ref=A1:E11: expected `'picture'`, got `None` (RuntimeError('screenshot failed: com_error 0x800A03EC'))
- **excel-13-tables** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-15-print-options** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-16-charts** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-17-combo-charts** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-20-workbook-protection** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `excel.unprotect_workbook`: expected `[True, True, False]`, got `[True, False, True]`
    - Unprotect(wrong): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'This command is unavailable because the licence to use this application has expired.', 'xlmain11.chm', 0, -2146827284), None)
    - Unprotect(password): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'This command is unavailable because the licence to use this application has expired.', 'xlmain11.chm', 0, -2146827284), None)
- **excel-19-chart-sheet** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-24-text-boxes** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.text_boxes`: ok, got `['Note\nSecond paragraph', 'Page box', 'Plain box']`
    - Shapes(1).Type = 17
    - Shapes(2).Type = 17
    - Shapes(3).Type = 17
- **word-25-protection-read-only** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.unprotect`: expected `[True, True, 'none']`, got `[True, False, 'readOnly']`
    - Unprotect(wrong): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The Unprotect method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None)
    - Unprotect(password): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The Unprotect method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None)
- **word-26-protection-tracked-sha512** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.unprotect`: expected `[True, True, 'none']`, got `[True, False, 'trackedChanges']`
    - Unprotect(wrong): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The Unprotect method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None)
    - Unprotect(password): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The Unprotect method or property is not available because the licence to use this application has expired..', 'wdmain11.chm', 37373, -2146823683), None)
- **word-23-charts** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
  - probe `word.charts`: ok, got `3`
    - InlineShapes(1).Type = 12
    - InlineShapes(2).Type = 12
    - InlineShapes(3).Type = 12
- **ppt-08-charts** on gregory_citadel (Office 2016+): ok
  - probe `ppt.charts`: ok, got `1`
    - slide 1 shape 1: HasChart=False
    - slide 1 shape 2: HasChart=True
  - probe `ppt.charts`: ok, got `2`
    - slide 2 shape 1: HasChart=False
    - slide 2 shape 2: HasChart=True
    - slide 2 shape 3: HasChart=True
  - probe `ppt.chart`: ok, got `'line_markers'`
    - slide 1 shape 1: HasChart=False
    - slide 1 shape 2: HasChart=True
  - probe `ppt.chart`: ok, got `'Visits and orders'`
    - slide 1 shape 1: HasChart=False
    - slide 1 shape 2: HasChart=True
  - probe `ppt.chart`: ok, got `'doughnut'`
    - slide 2 shape 1: HasChart=False
    - slide 2 shape 2: HasChart=True
    - slide 2 shape 3: HasChart=True
  - probe `ppt.chart`: ok, got `'column_percent'`
    - slide 2 shape 1: HasChart=False
    - slide 2 shape 2: HasChart=True
    - slide 2 shape 3: HasChart=True
- **ppt-10-links-table-edit** on gregory_citadel (Office 2016+): ok
  - probe `ppt.links`: ok, got `['https://example.org/docs', 'mailto:team@example.org', 'slide:3']`
    - Slide.Hyperlinks: [('https://example.org/docs', ''), ('mailto:team@example.org', ''), ('', '-1,-1,NEXT'), ('', '258,3, ')]
- **ppt-11-edit-python-pptx-file** on gregory_citadel (Office 2016+): ok
  - probe `ppt.links`: ok, got `['https://example.org/']`
    - Slide.Hyperlinks: [('https://example.org/', '')]
- **excel-06-properties** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-91-observe-order** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-92-observe-unknown-element** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-93-observe-bad-values** on gregory_citadel (Office 2016+): prompt
  - dialog `Microsoft Word`: OK / Word experienced an error trying to open the file.
Try these suggestions.
* Check the file permissions for the document or drive.
* Make sure there is sufficient free memory and disk space.
* Open the file with the Text Recovery converter. (C:\...\word-93-observe-bad-values.docx) / 200494
  - error: `com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'Word experienced an error trying to open the file.\nTry these suggestions.\n* Check the file permissions for the document or drive.\n* Make sure there is sufficient free memory and disk space.\n* Open the file with the Text Recover`
- **word-94-observe-sdk-values** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-96-observe-iso-band-size** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **word-95-observe-libreoffice-structure** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-91-observe-row-order** on gregory_citadel (Office 2016+): error
  - error: `com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Open method of Workbooks class failed', 'xlmain11.chm', 0, -2146827284), None)`
- **excel-93-observe-iso-optional** on gregory_citadel (Office 2016+): error
  - error: `com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Open method of Workbooks class failed', 'xlmain11.chm', 0, -2146827284), None)`
- **excel-94-observe-iso-paper-size** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **excel-92-observe-newer-attribute** on gregory_citadel (Office 2016+): prompt
  - dialog `Sign in to set up Office`: 
- **ppt-91-observe-missing-required** on gregory_citadel (Office 2016+): error
  - error: `com_error(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2147467259), None)`

## Screenshots

Saved in the corpus folder; send them along with this report.

- `ppt-03-textboxes-pictures` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-03-textboxes-pictures-slide1.png
- `ppt-07-properties-images` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-07-properties-images-slide1.png
- `excel-12-conditional` on gregory_citadel (Office 2016+): RuntimeError('screenshot failed: com_error 0x800A03EC')
- `excel-16-charts` on gregory_citadel (Office 2016+): chart image saved: screens/gregory_citadel-excel-16-charts-Sales-chart1.png
- `excel-16-charts` on gregory_citadel (Office 2016+): chart image saved: screens/gregory_citadel-excel-16-charts-Sales-chart2.png
- `excel-16-charts` on gregory_citadel (Office 2016+): chart image saved: screens/gregory_citadel-excel-16-charts-Sales-chart3.png
- `excel-16-charts` on gregory_citadel (Office 2016+): chart image saved: screens/gregory_citadel-excel-16-charts-Sales-chart4.png
- `excel-16-charts` on gregory_citadel (Office 2016+): chart image saved: screens/gregory_citadel-excel-16-charts-Points-chart1.png
- `excel-17-combo-charts` on gregory_citadel (Office 2016+): chart image saved: screens/gregory_citadel-excel-17-combo-charts-Combo-chart1.png
- `excel-17-combo-charts` on gregory_citadel (Office 2016+): chart image saved: screens/gregory_citadel-excel-17-combo-charts-Combo-chart2.png
- `excel-18-edit-openpyxl-file` on gregory_citadel (Office 2016+): chart image saved: screens/gregory_citadel-excel-18-edit-openpyxl-file-Data-chart2.png
- `ppt-08-charts` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-08-charts-slide1.png
- `ppt-08-charts` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-08-charts-slide2.png
- `ppt-09-layouts-tables-notes` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-09-layouts-tables-notes-slide1.png
- `ppt-09-layouts-tables-notes` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-09-layouts-tables-notes-slide2.png
- `ppt-09-layouts-tables-notes` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-09-layouts-tables-notes-slide3.png
- `ppt-09-layouts-tables-notes` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-09-layouts-tables-notes-slide4.png
- `ppt-09-layouts-tables-notes` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-09-layouts-tables-notes-slide5.png
- `ppt-10-links-table-edit` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-10-links-table-edit-slide1.png
- `ppt-10-links-table-edit` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-10-links-table-edit-slide2.png
- `ppt-10-links-table-edit` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-10-links-table-edit-slide3.png
- `ppt-11-edit-python-pptx-file` on gregory_citadel (Office 2016+): slide image saved: screens/gregory_citadel-ppt-11-edit-python-pptx-file-slide1.png

## Not judged

Counted as ok; the application can't answer these.

- `excel-08-styles-formats` on gregory_citadel (Office 2016+): `excel.text` sheet=Formats ref=B13: month/day names follow the regional settings
- `excel-08-styles-formats` on gregory_citadel (Office 2016+): `excel.text` sheet=Formats ref=B14: month/day names follow the regional settings
- `excel-08-styles-formats` on gregory_citadel (Office 2016+): `excel.text` sheet=Formats ref=B15: month/day names follow the regional settings

## Runs

- `results-gregory_citadel-word16-excel16-powerpoint16-20261116-183434.json`: Python 3.13.7, 64-bit, `C:\Python313\python.exe`; pywin32 311; COM binding dynamic

Sources: `results-gregory_citadel-word16-excel16-powerpoint16-20261116-183434.json`
