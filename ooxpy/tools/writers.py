"""Files written by other Python libraries, for interop checks in both directions.

    python -m tools.writers OUT        # write a corpus with python-docx, openpyxl,
                                       # XlsxWriter and python-pptx (whichever are installed)
    python -m tools.interop OUT        # then check it with oox

`build(out)` writes each file with a feature-rich script per library; `readback(path)`
opens an oox-written file with the matching library and returns what it found, so
the tests can check the other direction too.
"""

from __future__ import annotations

import datetime as dt
import io
import struct
import sys
import zlib
from pathlib import Path


def _png(w: int = 4, h: int = 3) -> bytes:
    raw = b"".join(b"\x00" + bytes((200, 60, 60)) * w for _ in range(h))

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _python_docx(out: Path) -> list[Path]:
    import docx
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.shared import Inches, Pt, RGBColor

    d = docx.Document()
    d.core_properties.title = "python-docx interop"
    d.core_properties.author = "tools.writers"
    d.add_heading("Interop from python-docx", 0)
    d.add_heading("Level one", 1)
    p = d.add_paragraph("Plain, ")
    p.add_run("bold").bold = True
    p.add_run(", ")
    r = p.add_run("italic red 14pt")
    r.italic, r.font.size, r.font.color.rgb = True, Pt(14), RGBColor(0xC0, 0, 0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    d.add_paragraph("first bullet", style="List Bullet")
    d.add_paragraph("second bullet", style="List Bullet")
    d.add_paragraph("first number", style="List Number")
    d.add_paragraph("A quote", style="Intense Quote")
    t = d.add_table(rows=3, cols=3)
    t.style = "Light Grid Accent 1"
    for i, row in enumerate(t.rows):
        for j, cell in enumerate(row.cells):
            cell.text = f"r{i}c{j}"
    t.cell(1, 0).merge(t.cell(2, 0))
    d.add_picture(io.BytesIO(_png()), width=Inches(1))
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    s = d.add_section()
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = s.page_height, s.page_width
    s.header.paragraphs[0].text = "Landscape header"
    s.footer.paragraphs[0].text = "Footer"
    d.add_paragraph("In the landscape section.")
    if hasattr(d, "add_comment"):                       # python-docx >= 1.2
        d.add_comment(d.paragraphs[2].runs, text="A comment", author="tools", initials="t")
    path = out / "python-docx_features.docx"
    d.save(path)

    empty = out / "python-docx_empty.docx"
    docx.Document().save(empty)
    return [path, empty]


def _openpyxl(out: Path) -> list[Path]:
    import openpyxl
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.comments import Comment
    from openpyxl.drawing.image import Image
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Month", "North", "South", "Date"])
    for i, (n, s) in enumerate([(10, 7), (12, 9), (9, 11), (15, 8)]):
        ws.append([f"M{i + 1}", n, s, dt.date(2026, i + 1, 1)])
    ws["F1"], ws["F2"] = "Total", "=SUM(B2:C5)"
    ws["A1"].font = Font(bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="4472C4")
    ws["B2"].border = Border(bottom=Side(style="thin"))
    ws["A1"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws["D2"].number_format = "yyyy-mm-dd"
    ws.column_dimensions["A"].width = 14
    ws.freeze_panes = "B2"
    ws.merge_cells("H1:I2")
    ws["H1"] = "merged"
    ws["B3"].comment = Comment("checked", "tools")
    ws["A7"] = "link"
    ws["A7"].hyperlink = "https://example.org/"
    tab = Table(displayName="SalesTable", ref="A1:D5")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)
    ws.conditional_formatting.add("B2:B5", CellIsRule(operator="greaterThan", formula=["11"],
                                                      fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.conditional_formatting.add("C2:C5", ColorScaleRule(start_type="min", start_color="FFFFFF",
                                                          end_type="max", end_color="63BE7B"))
    dv = DataValidation(type="list", formula1='"yes,no"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E2:E5")
    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=5)
    cats = Reference(ws, min_col=1, min_row=2, max_row=5)
    for kind, anchor in ((BarChart, "K2"), (LineChart, "K18")):
        ch = kind()
        ch.title = kind.__name__
        ch.add_data(data, titles_from_data=True)
        ch.set_categories(cats)
        ws.add_chart(ch, anchor)
    pie = PieChart()
    pie.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    pie.set_categories(cats)
    ws.add_chart(pie, "K34")
    ws.add_image(Image(io.BytesIO(_png())), "H5")
    other = wb.create_sheet("Notes")
    other["A1"] = "Second sheet"
    other["A2"] = 3.5
    other["A3"] = True
    other.sheet_properties.tabColor = "00B050"
    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names["Months"] = DefinedName("Months", attr_text="Sales!$A$2:$A$5")
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    sheet_chart = BarChart()
    sheet_chart.add_data(data, titles_from_data=True)
    wb.create_chartsheet("Chart sheet").add_chart(sheet_chart)
    path = out / "openpyxl_features.xlsx"
    wb.save(path)
    empty = out / "openpyxl_empty.xlsx"
    openpyxl.Workbook().save(empty)
    return [path, empty]


def _xlsxwriter(out: Path) -> list[Path]:
    import xlsxwriter

    path = out / "xlsxwriter_features.xlsx"
    wb = xlsxwriter.Workbook(str(path), {"default_date_format": "yyyy-mm-dd"})
    bold = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
    ws = wb.add_worksheet("Data")
    ws.write_row(0, 0, ["Item", "Q1", "Q2", "When"], bold)
    for i, (a, b) in enumerate([(5, 6), (7, 4), (3, 9)]):
        ws.write_string(i + 1, 0, f"I{i}")
        ws.write_number(i + 1, 1, a)
        ws.write_number(i + 1, 2, b)
        ws.write_datetime(i + 1, 3, dt.datetime(2026, 3, i + 1))
    ws.write_formula("E2", "=B2+C2", None, 11)
    ws.write_boolean("F2", True)
    ws.write_url("A6", "https://example.org/", string="example")
    ws.write_comment("B2", "a note")
    ws.merge_range("H1:J1", "Merged", bold)
    ws.set_column("A:A", 12)
    ws.freeze_panes(1, 1)
    ws.add_table("A9:C11", {"columns": [{"header": "x"}, {"header": "y"}, {"header": "z"}],
                            "data": [[1, 2, 3], [4, 5, 6]]})
    ws.conditional_format("B2:C4", {"type": "3_color_scale"})
    ws.conditional_format("B2:B4", {"type": "cell", "criteria": ">", "value": 4, "format": bold})
    ws.data_validation("G2:G4", {"validate": "list", "source": ["a", "b"]})
    for kind, cell in (("column", "L2"), ("line", "L18")):
        ch = wb.add_chart({"type": kind})
        ch.add_series({"name": "=Data!$B$1", "categories": "=Data!$A$2:$A$4",
                       "values": "=Data!$B$2:$B$4"})
        ch.add_series({"name": "=Data!$C$1", "categories": "=Data!$A$2:$A$4",
                       "values": "=Data!$C$2:$C$4", "y2_axis": kind == "line"})
        ch.set_title({"name": kind})
        ws.insert_chart(cell, ch)
    combo = wb.add_chart({"type": "column"})
    combo.add_series({"values": "=Data!$B$2:$B$4"})
    line = wb.add_chart({"type": "line"})
    line.add_series({"values": "=Data!$C$2:$C$4", "y2_axis": True})
    combo.combine(line)
    ws.insert_chart("L34", combo)
    ws.insert_image("H5", "image.png", {"image_data": io.BytesIO(_png())})
    ws.write_rich_string("A8", "plain ", bold, "bold", " plain")
    ws.write_array_formula("F5:F7", "{=B2:B4*2}")
    wb.define_name("Items", "=Data!$A$2:$A$4")
    wb.add_worksheet("Empty")
    cs = wb.add_chartsheet("Pie")
    pie = wb.add_chart({"type": "pie"})
    pie.add_series({"categories": "=Data!$A$2:$A$4", "values": "=Data!$B$2:$B$4"})
    cs.set_chart(pie)
    ws.set_landscape()
    ws.repeat_rows(0)
    wb.close()
    return [path]


def _python_pptx(out: Path) -> list[Path]:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "Interop from python-pptx"
    s.placeholders[1].text = "subtitle"
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Bullets"
    tf = s.placeholders[1].text_frame
    tf.text = "first"
    para = tf.add_paragraph()
    para.text, para.level = "nested", 1
    s.notes_slide.notes_text_frame.text = "Speaker notes"
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Shapes, table, chart, picture"
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5),
                             Inches(2), Inches(1))
    box.text = "rounded"
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x44, 0x72, 0xC4)
    run = box.text_frame.paragraphs[0].runs[0]
    run.font.size, run.font.bold = Pt(18), True
    run.hyperlink.address = "https://example.org/"
    tb = s.shapes.add_textbox(Inches(0.5), Inches(3), Inches(2), Inches(0.5))
    tb.text_frame.text = "textbox"
    table = s.shapes.add_table(3, 2, Inches(3), Inches(1.5), Inches(3), Inches(1.2)).table
    for i in range(3):
        for j in range(2):
            table.cell(i, j).text = f"{i},{j}"
    table.cell(1, 0).merge(table.cell(2, 0))
    cd = CategoryChartData()
    cd.categories = ["A", "B", "C"]
    cd.add_series("S1", (1, 3, 2))
    cd.add_series("S2", (2, 1, 4))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.2), Inches(1.5), Inches(3.5),
                       Inches(2.5), cd)
    s.shapes.add_picture(io.BytesIO(_png()), Inches(3), Inches(3.2), Inches(1))
    grp = s.shapes.add_group_shape()
    grp.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), Inches(4), Inches(1), Inches(1))
    grp.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.6), Inches(4), Inches(1), Inches(1))
    s.shapes.add_connector(1, Inches(4), Inches(5), Inches(6), Inches(6))
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pie = CategoryChartData()
    pie.categories = ["x", "y"]
    pie.add_series("Share", (0.3, 0.7))
    s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(1), Inches(1), Inches(4), Inches(4), pie)
    prs.core_properties.title = "python-pptx interop"
    path = out / "python-pptx_features.pptx"
    prs.save(path)
    empty = out / "python-pptx_empty.pptx"
    Presentation().save(empty)
    return [path, empty]


WRITERS = {"docx": _python_docx, "openpyxl": _openpyxl, "xlsxwriter": _xlsxwriter,
           "pptx": _python_pptx}


def build(out_dir: str | Path) -> list[Path]:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for module, write in WRITERS.items():
        try:
            __import__(module)
        except ImportError:
            print(f"skipped {module} (not installed)", file=sys.stderr)
            continue
        written += write(out)
    return written


def readback(path: str | Path) -> dict:
    """Open an oox-written file with the matching library; return what it found."""
    path = str(path)
    if path.endswith(".docx"):
        import docx

        d = docx.Document(path)
        return {"paragraphs": [p.text for p in d.paragraphs],
                "styles": [p.style.name if p.style is not None else None for p in d.paragraphs],
                "tables": [[[c.text for c in r.cells] for r in t.rows] for t in d.tables],
                "sections": len(d.sections),
                "pictures": len(d.inline_shapes),
                "title": d.core_properties.title}
    if path.endswith(".xlsx"):
        import openpyxl

        wb = openpyxl.load_workbook(path)
        return {"sheets": wb.sheetnames,
                "values": {ws.title: [list(r) for r in ws.iter_rows(values_only=True)]
                           for ws in wb.worksheets},
                "charts": {ws.title: len(ws._charts) for ws in wb.worksheets},
                "merged": {ws.title: [str(m) for m in ws.merged_cells.ranges]
                           for ws in wb.worksheets},
                "tables": {ws.title: list(ws.tables) for ws in wb.worksheets}}
    if path.endswith(".pptx"):
        from pptx import Presentation

        prs = Presentation(path)
        return {"slides": [[sh.shape_type and str(sh.shape_type).split(".")[-1].split(" ")[0]
                            for sh in s.shapes] for s in prs.slides],
                "texts": [[sh.text_frame.text for sh in s.shapes if sh.has_text_frame]
                          for s in prs.slides],
                "charts": [sum(1 for sh in s.shapes if sh.has_chart) for s in prs.slides],
                "tables": [sum(1 for sh in s.shapes if sh.has_table) for s in prs.slides],
                "notes": [s.notes_slide.notes_text_frame.text if s.has_notes_slide else None
                          for s in prs.slides],
                "layouts": [ly.name for ly in prs.slide_layouts]}
    raise ValueError(path)


if __name__ == "__main__":
    files = build(sys.argv[1] if len(sys.argv) > 1 else "writers-out")
    print(f"{len(files)} files")
