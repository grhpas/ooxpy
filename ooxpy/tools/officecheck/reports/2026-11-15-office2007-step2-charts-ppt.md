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
| `ppt-03-textboxes-pictures` | ✅ |
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
| `ppt-07-properties-images` | ✅ |
| `excel-08-styles-formats` | · |
| `excel-09-layout` | · |
| `excel-10-style-existing` | · |
| `excel-11-names-autofit` | · |
| `excel-12-conditional` | · |
| `excel-13-tables` | · |
| `excel-14-comments-links` | · |
| `excel-15-print-options` | · |
| `excel-16-charts` | · |
| `excel-17-combo-charts` | ❌ |
| `excel-19-chart-sheet` | ✅ |
| `excel-18-edit-openpyxl-file` | ✅ |
| `word-23-charts` | · |
| `ppt-08-charts` | · |
| `ppt-09-layouts-tables-notes` | ✅ |
| `ppt-10-links-table-edit` | ❌ |
| `ppt-12-sections` | ✅ |
| `ppt-11-edit-python-pptx-file` | ✅ |
| `excel-06-properties` | · |
| `ppt-06-properties` | · |
| `word-91-observe-order` *(observe)* | · |
| `word-92-observe-unknown-element` *(observe)* | · |
| `word-93-observe-bad-values` *(observe)* | · |
| `word-94-observe-sdk-values` *(observe)* | · |
| `word-95-observe-libreoffice-structure` *(observe)* | · |
| `excel-91-observe-row-order` *(observe)* | · |
| `excel-92-observe-newer-attribute` *(observe)* | · |
| `ppt-91-observe-missing-required` *(observe)* | · |
| `ppt-90-control-malformed` *(control)* | · |

## Details

- **excel-17-combo-charts** on mosyxlon3 (Office 2007): mismatch
  - probe `excel.chart` sheet=Combo: expected `'EUR'`, got `None` (com_error(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2147467259), None))
- **ppt-10-links-table-edit** on mosyxlon3 (Office 2007): mismatch
  - probe `ppt.links`: expected `['https://example.org/docs', 'mailto:team@example.org', 'slide:3']`, got `['slide:-1', 'slide:3']`

## Screenshots

Saved in the corpus folder; send them along with this report.

- `ppt-03-textboxes-pictures` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-03-textboxes-pictures-slide1.png
- `ppt-07-properties-images` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-07-properties-images-slide1.png
- `excel-17-combo-charts` on mosyxlon3 (Office 2007): chart image saved: screens/mosyxlon3-excel-17-combo-charts-Combo-chart1.png
- `excel-17-combo-charts` on mosyxlon3 (Office 2007): chart image saved: screens/mosyxlon3-excel-17-combo-charts-Combo-chart2.png
- `excel-18-edit-openpyxl-file` on mosyxlon3 (Office 2007): chart image saved: screens/mosyxlon3-excel-18-edit-openpyxl-file-Data-chart2.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide1.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide2.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide3.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide4.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide5.png
- `ppt-10-links-table-edit` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-10-links-table-edit-slide1.png
- `ppt-10-links-table-edit` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-10-links-table-edit-slide2.png
- `ppt-10-links-table-edit` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-10-links-table-edit-slide3.png
- `ppt-11-edit-python-pptx-file` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-11-edit-python-pptx-file-slide1.png

## Not judged

Counted as ok; the application can't answer these.

- `ppt-12-sections` on mosyxlon3 (Office 2007): `ppt.sections`: sections are PowerPoint 2010+; 2007 ignores them

## Runs

- `results-mosyxlon3-excel12-powerpoint12-20261115-123044.json`: Python 3.14.6, 64-bit, `C:\Python314\python.exe`; pywin32 312; COM binding dynamic

Sources: `results-mosyxlon3-excel12-powerpoint12-20261115-123044.json`
