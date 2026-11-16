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
| `excel-16-charts` | ✅ |
| `word-23-charts` | ✅ |
| `ppt-08-charts` | ✅ |
| `ppt-09-layouts-tables-notes` | ✅ |
| `excel-06-properties` | · |
| `ppt-06-properties` | · |
| `word-91-observe-order` *(observe)* | · |
| `word-92-observe-unknown-element` *(observe)* | · |
| `word-93-observe-bad-values` *(observe)* | · |
| `word-94-observe-sdk-values` *(observe)* | 👁 ok |
| `word-95-observe-libreoffice-structure` *(observe)* | 👁 ok |
| `excel-91-observe-row-order` *(observe)* | · |
| `excel-92-observe-newer-attribute` *(observe)* | · |
| `ppt-91-observe-missing-required` *(observe)* | · |
| `ppt-90-control-malformed` *(control)* | · |

## Details

- **word-23-charts** on mosyxlon3 (Office 2007): ok
  - probe `word.charts`: ok, got `3`
    - InlineShapes(1).Type = 12
    - InlineShapes(2).Type = 12
    - InlineShapes(3).Type = 12
- **ppt-08-charts** on mosyxlon3 (Office 2007): ok
  - probe `ppt.charts`: ok, got `1`
    - slide 1 shape 1: HasChart=False
    - slide 1 shape 2: HasChart=True
  - probe `ppt.charts`: ok, got `2`
    - slide 2 shape 1: HasChart=False
    - slide 2 shape 2: HasChart=True
    - slide 2 shape 3: HasChart=True
  - probe `ppt.chart`: ok, got `None` (not judged: not in Office 2007's object model (charts need SP2): AttributeError('<unknown>.Chart'))
    - slide 1 shape 1: HasChart=False
    - slide 1 shape 2: HasChart=True
  - probe `ppt.chart`: ok, got `None` (not judged: not in Office 2007's object model (charts need SP2): AttributeError('<unknown>.Chart'))
    - slide 1 shape 1: HasChart=False
    - slide 1 shape 2: HasChart=True
  - probe `ppt.chart`: ok, got `None` (not judged: not in Office 2007's object model (charts need SP2): AttributeError('<unknown>.Chart'))
    - slide 2 shape 1: HasChart=False
    - slide 2 shape 2: HasChart=True
    - slide 2 shape 3: HasChart=True
  - probe `ppt.chart`: ok, got `None` (not judged: not in Office 2007's object model (charts need SP2): AttributeError('<unknown>.Chart'))
    - slide 2 shape 1: HasChart=False
    - slide 2 shape 2: HasChart=True
    - slide 2 shape 3: HasChart=True
- **word-94-observe-sdk-values** on mosyxlon3 (Office 2007): ok
- **word-95-observe-libreoffice-structure** on mosyxlon3 (Office 2007): ok

## Screenshots

Saved in the corpus folder; send them along with this report.

- `excel-16-charts` on mosyxlon3 (Office 2007): chart image saved: screens/mosyxlon3-excel-16-charts-Sales-chart1.png
- `excel-16-charts` on mosyxlon3 (Office 2007): chart image saved: screens/mosyxlon3-excel-16-charts-Sales-chart2.png
- `excel-16-charts` on mosyxlon3 (Office 2007): chart image saved: screens/mosyxlon3-excel-16-charts-Sales-chart3.png
- `excel-16-charts` on mosyxlon3 (Office 2007): chart image saved: screens/mosyxlon3-excel-16-charts-Sales-chart4.png
- `excel-16-charts` on mosyxlon3 (Office 2007): chart image saved: screens/mosyxlon3-excel-16-charts-Points-chart1.png
- `ppt-08-charts` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-08-charts-slide1.png
- `ppt-08-charts` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-08-charts-slide2.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide1.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide2.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide3.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide4.png
- `ppt-09-layouts-tables-notes` on mosyxlon3 (Office 2007): slide image saved: screens/mosyxlon3-ppt-09-layouts-tables-notes-slide5.png

## Runs

- `results-mosyxlon3-word12-excel12-powerpoint12-20261115-114420.json`: Python 3.14.6, 64-bit, `C:\Python314\python.exe`; pywin32 312; COM binding dynamic

Sources: `results-mosyxlon3-word12-excel12-powerpoint12-20261115-114420.json`
