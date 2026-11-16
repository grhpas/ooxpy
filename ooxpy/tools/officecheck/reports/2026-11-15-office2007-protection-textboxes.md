# Office check report

Corpus: oox 0.0.1, Open XML SDK data @ 431ab05cf160248cc3885a4a766026d4f8243792 2026-08-18.

✅ ok · ⚠️ prompt (repair/recovery dialog) · ❌ error or content mismatch · ⏱️ timeout · · not run

*Control* cases are broken on purpose: ✅ means the checker noticed, either because Office complained or because it repaired silently and content went missing.

*Observe* cases (👁) break the schema in one specific way and are not judged: they record how strict each Office version is.

| Case | mosyxlon3 (Office 2007) |
|---|---|
| `word-01-minimal` | · |
| `word-02-formatting` | · |
| `word-03-table` | · |
| `word-04-pictures` | · |
| `word-05-styles` | · |
| `word-06-w14` | · |
| `word-07-odfpy-passwd` | · |
| `word-08-edit-word-file` | · |
| `excel-01-types` | · |
| `excel-02-formulas` | · |
| `excel-03-sheets` | · |
| `excel-04-pictures` | · |
| `excel-05-edit-existing` | · |
| `ppt-01-minimal` | · |
| `ppt-02-placeholders` | · |
| `ppt-03-textboxes-pictures` | · |
| `ppt-04-4x3` | · |
| `ppt-05-edit-powerpoint-file` | · |
| `word-90-control-malformed` *(control)* | · |
| `excel-90-control-malformed` *(control)* | · |
| `word-10-defaults` | · |
| `word-11-lists` | · |
| `word-12-sections` | · |
| `word-13-links-notes` | · |
| `word-14-toc-template` | · |
| `word-15-tables` | · |
| `word-16-floating-pictures` | · |
| `word-17-tracked-changes` | · |
| `word-18-tracked-changes-reject` | · |
| `word-19-formatting-styles` | · |
| `word-20-fields-captions` | · |
| `word-21-content-controls` | · |
| `word-22-more-images` | · |
| `excel-07-properties-images` | · |
| `ppt-07-properties-images` | · |
| `excel-08-styles-formats` | · |
| `excel-09-layout` | · |
| `excel-10-style-existing` | · |
| `excel-11-names-autofit` | · |
| `excel-12-conditional` | · |
| `excel-13-tables` | · |
| `excel-14-comments-links` | · |
| `excel-15-print-options` | · |
| `excel-16-charts` | · |
| `excel-17-combo-charts` | · |
| `excel-20-workbook-protection` | ✅ |
| `excel-19-chart-sheet` | · |
| `excel-18-edit-openpyxl-file` | · |
| `word-24-text-boxes` | ✅ |
| `word-25-protection-read-only` | ⚠️ |
| `word-26-protection-tracked-sha512` | ⚠️ |
| `word-27-protection-comments` | ✅ |
| `word-23-charts` | · |
| `ppt-08-charts` | · |
| `ppt-09-layouts-tables-notes` | · |
| `ppt-10-links-table-edit` | · |
| `ppt-12-sections` | · |
| `ppt-11-edit-python-pptx-file` | · |
| `excel-06-properties` | · |
| `ppt-06-properties` | · |
| `word-91-observe-order` *(observe)* | · |
| `word-92-observe-unknown-element` *(observe)* | · |
| `word-93-observe-bad-values` *(observe)* | · |
| `word-94-observe-sdk-values` *(observe)* | · |
| `word-96-observe-iso-band-size` *(observe)* | · |
| `word-95-observe-libreoffice-structure` *(observe)* | · |
| `excel-91-observe-row-order` *(observe)* | · |
| `excel-93-observe-iso-optional` *(observe)* | · |
| `excel-94-observe-iso-paper-size` *(observe)* | · |
| `excel-92-observe-newer-attribute` *(observe)* | · |
| `ppt-91-observe-missing-required` *(observe)* | · |
| `ppt-90-control-malformed` *(control)* | · |

## Details

- **excel-20-workbook-protection** on mosyxlon3 (Office 2007): ok
  - probe `excel.unprotect_workbook`: ok, got `[True, True, False]`
    - Unprotect(wrong): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Office Excel', 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.', 'C:\\Program Files (x86)\\Microsoft Office\\Office12\\1033\\XLMAIN11.CHM', 0, -2146827284), None)
- **word-24-text-boxes** on mosyxlon3 (Office 2007): ok
  - probe `word.text_boxes`: ok, got `['Note\nSecond paragraph', 'Page box', 'Plain box']`
    - Shapes(1).Type = 17
    - Shapes(2).Type = 17
    - Shapes(3).Type = 17
- **word-25-protection-read-only** on mosyxlon3 (Office 2007): prompt
  - dialog `Microsoft Office Word`: OK / The password is incorrect. / Show H&elp >> / &Open in Help Window / 200919
  - probe `word.unprotect`: ok, got `[True, True, 'none']`
    - Unprotect(wrong): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The password is incorrect.', 'C:\\Program Files (x86)\\Microsoft Office\\Office12\\1033\\WDMAIN11.CHM', 24965, -2146822803), None)
- **word-26-protection-tracked-sha512** on mosyxlon3 (Office 2007): prompt
  - dialog `Microsoft Office Word`: OK / The password is incorrect. / Show H&elp >> / &Open in Help Window / 200919
  - probe `word.unprotect`: ok, got `[True, True, 'none']`
    - Unprotect(wrong): com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The password is incorrect.', 'C:\\Program Files (x86)\\Microsoft Office\\Office12\\1033\\WDMAIN11.CHM', 24965, -2146822803), None)

## Runs

- `results-mosyxlon3-word12-excel12-20261115-145704.json`: Python 3.14.6, 64-bit, `C:\Python314\python.exe`; pywin32 312; COM binding dynamic

Sources: `results-mosyxlon3-word12-excel12-20261115-145704.json`
