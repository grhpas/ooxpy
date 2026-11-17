"""Charts: the shared builder/reader and the Excel, Word and PowerPoint wrappers."""

from __future__ import annotations

import io

import pytest

from oox.charts import EXCEL_CHART_TYPES, KINDS, PALETTE, C, Series, build_chart, chart_info
from oox.document import open_document
from oox.spreadsheet import SpreadsheetDocument
from oox.units import Cm
from oox.xlcharts import quote_sheet


def _reload(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return open_document(buf)


def _all_parts_valid(doc):
    problems = []
    for part in doc.package.parts.values():
        if part.content_type.endswith("chart+xml") or "drawing" in part.content_type:
            problems += part.element.validate(target="Office2007")
    return problems


@pytest.mark.parametrize("kind", KINDS)
def test_every_kind_builds_valid_xml_and_reads_back(kind):
    cats = [1, 2, 3.5] if kind == "scatter" else ["Q1", "Q2", "Q3"]
    series = [Series("Share", [50, 30, 20])] if kind == "pie" else \
        [Series("North", [120, 135, 150.5]), Series("South", [95, None, 99])]
    space = build_chart(kind, cats, series, title="Sales", x_title="X", y_title="Y",
                        legend="bottom", labels="value", number_format="0.0")
    assert space.validate(target="Office2007") == []
    info = chart_info(space)
    assert info.kind in EXCEL_CHART_TYPES.values()
    assert info.title == "Sales" and info.legend == "bottom" and info.labels == ["value"]
    assert info.categories == cats
    assert info.series[0].values == list(series[0].values)
    if kind not in ("pie", "doughnut"):
        assert (info.x_title, info.y_title) == ("X", "Y")
        assert info.number_format == "0.0"
        assert info.series[1].values == [95, None, 99]   # a gap stays a gap


@pytest.mark.parametrize("kind,options,expected", [
    ("column", {}, "column"), ("column", {"grouping": "stacked"}, "column_stacked"),
    ("column", {"grouping": "percent"}, "column_percent"), ("bar", {}, "bar"),
    ("bar", {"grouping": "stacked"}, "bar_stacked"), ("line", {}, "line_markers"),
    ("line", {"markers": False}, "line"), ("line", {"grouping": "stacked"}, "line_markers_stacked"),
    ("area", {}, "area"), ("area", {"grouping": "percent"}, "area_percent"),
    ("pie", {}, "pie"), ("doughnut", {}, "doughnut"), ("scatter", {}, "scatter"),
    ("scatter", {"lines": True}, "scatter_lines"),
    ("scatter", {"lines": True, "markers": False}, "scatter_lines_no_markers"),
    ("scatter", {"lines": True, "smooth": True}, "scatter_smooth"),
    ("scatter", {"lines": True, "smooth": True, "markers": False}, "scatter_smooth_no_markers"),
])
def test_type_vocabulary_matches_excel_chart_types(kind, options, expected):
    cats = [1, 2, 3] if kind == "scatter" else ["a", "b", "c"]
    space = build_chart(kind, cats, [Series("s", [1, 2, 3])], **options)
    assert chart_info(space).kind == expected
    assert expected in EXCEL_CHART_TYPES.values()


def test_colours_are_always_explicit():
    """Without colours, LibreOffice draws empty bars: oox writes the palette."""
    info = chart_info(build_chart("column", ["a"], [Series("x", [1]), Series("y", [2])]))
    assert [s.color for s in info.series] == list(PALETTE[:2])
    pie = chart_info(build_chart("pie", ["a", "b", "c"], [Series("x", [1, 2, 3])],
                                 colors=["FF0000"]))
    assert pie.series[0].point_colors == ["FF0000", PALETTE[1], PALETTE[2]]


def test_no_title_suppresses_excels_automatic_title():
    space = build_chart("pie", ["a"], [Series("Only series", [1])])
    assert chart_info(space).title is None
    auto = space.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}autoTitleDeleted")
    assert auto.get("val") == "1"


def test_errors():
    with pytest.raises(ValueError, match="chart kind"):
        build_chart("radar", ["a"], [Series("s", [1])])
    with pytest.raises(ValueError, match="one series"):
        build_chart("pie", ["a"], [Series("s", [1]), Series("t", [2])])
    with pytest.raises(ValueError, match="2 values for 3 categories"):
        build_chart("line", ["a", "b", "c"], [Series("s", [1, 2])])
    with pytest.raises(ValueError, match="legend"):
        build_chart("line", ["a"], [Series("s", [1])], legend="middle")
    with pytest.raises(ValueError, match="labels"):
        build_chart("line", ["a"], [Series("s", [1])], labels="everything")


def test_quote_sheet():
    assert quote_sheet("Sales") == "Sales"
    assert quote_sheet("Sales Data") == "'Sales Data'"
    assert quote_sheet("Bob's") == "'Bob''s'"
    assert quote_sheet("A1") == "'A1'"            # looks like a cell reference
    assert quote_sheet("2024") == "'2024'"


def _sales_book(sheet="Sales Data"):
    book = SpreadsheetDocument.new(sheet)
    book.append_row(sheet, ["Region", "Q1", "Q2"])
    for row in [("North", 120, 135), ("South", 95, 110), ("West", 80, 84.5)]:
        book.append_row(sheet, list(row))
    return book


def test_excel_chart_from_range_round_trip():
    book = _sales_book()
    book.add_chart("Sales Data", "column", "A1:C4", anchor="E2", title="Sales",
                   y_title="EUR k", legend="bottom")
    book.add_chart("Sales Data", "line", "A1:C4", anchor="E20", series_in="rows", markers=False)
    assert _all_parts_valid(book) == []
    back = _reload(book)
    first, second = back.charts("Sales Data")
    assert first.anchor == "E2" and first.name == "Chart 1" and first.kind == "column"
    assert [s.name for s in first.series] == ["Q1", "Q2"]
    assert first.series[1].values == [135, 110, 84.5]
    assert first.series[0].values_ref == "'Sales Data'!$B$2:$B$4"
    assert first.series[0].categories_ref == "'Sales Data'!$A$2:$A$4"
    assert first.categories == ["North", "South", "West"]
    assert second.kind == "line" and second.categories == ["Q1", "Q2"]
    assert [s.name for s in second.series] == ["North", "South", "West"]
    assert second.series[2].values_ref == "'Sales Data'!$B$4:$C$4"
    assert sorted(p for p in back.package.parts if "chart" in p) == \
        ["/xl/charts/chart1.xml", "/xl/charts/chart2.xml"]


def test_excel_chart_anchor_follows_column_widths():
    book = _sales_book("S")
    book.add_chart("S", "pie", "A1:B4", anchor="B2", width=Cm(5), height=Cm(3))
    wide = _sales_book("S")
    wide.set_column_width("S", "B:H", 30)
    wide.add_chart("S", "pie", "A1:B4", anchor="B2", width=Cm(5), height=Cm(3))

    def to_col(doc):
        d = next(p for p in doc.package.parts.values() if p.partname.startswith("/xl/drawings/"))
        to = d.element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}to")
        return int(to[0].text)

    # 5 cm = 189 px: from B over 64-px columns it ends inside D (0-based 3);
    # over 30-character (215-px) columns it ends inside B (0-based 1)
    assert to_col(book) == 3 and to_col(wide) == 1


def test_excel_chart_values_come_from_cells():
    from oox.spreadsheet import Formula

    book = SpreadsheetDocument.new("S")
    book.append_row("S", ["", "Amount"])
    book.append_row("S", ["a", Formula("1+1", 2)])
    book.append_row("S", ["b", "n/a"])                 # text: a gap in the series
    book.add_chart("S", "column", "A1:B3")
    (info,) = book.charts("S")
    assert info.series[0].values == [2, None]
    assert info.series[0].name == "Amount"


def test_excel_chart_needs_a_header_and_label_column():
    book = _sales_book("S")
    with pytest.raises(ValueError, match="header row"):
        book.add_chart("S", "column", "B2:B4")
    with pytest.raises(ValueError, match="series_in"):
        book.add_chart("S", "column", "A1:C4", series_in="diagonal")


def test_word_chart_with_embedded_workbook():
    from oox.document import WordprocessingDocument

    doc = WordprocessingDocument.new()
    doc.add_chart("column", ["Q1", "Q2", "Q3"], {"North": [120, 135, 150], "South": [95, None, 99]},
                  title="Sales", legend="bottom", labels="value")
    doc.add_chart("pie", ["A", "B"], [("Share", [60, 40])], width=Cm(8), height=Cm(6))
    assert _all_parts_valid(doc) == []
    back = _reload(doc)
    col, pie = back.charts()
    assert col.kind == "column" and col.title == "Sales"
    south = col.series[1]
    assert (south.name, list(south.values), south.values_ref, south.categories_ref) == (
        "South", [95, None, 99], "Sheet1!$C$2:$C$4", "Sheet1!$A$2:$A$4")
    assert pie.kind == "pie" and pie.series[0].name == "Share"
    # the embedded workbook holds the same data, at the referenced cells
    space = back.package.parts["/word/charts/chart1.xml"].element
    ext = space.find("{http://schemas.openxmlformats.org/drawingml/2006/chart}externalData")
    part = back.package.parts["/word/charts/chart1.xml"]
    rel = part.rels[ext.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")]
    book = open_document(io.BytesIO(back.package.parts[part.resolve(rel.target)].to_bytes()))
    assert book.values("Sheet1") == [["", "North", "South"], ["Q1", 120, 95],
                                     ["Q2", 135, None], ["Q3", 150, 99]]
    # inline drawing with the chart size
    ext_el = back.document.find(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
    assert (ext_el.get("cx"), ext_el.get("cy")) == ("5400000", "2700000")


def test_powerpoint_chart_per_slide():
    from oox.presentation import PresentationDocument

    pres = PresentationDocument.new()
    s1, s2 = pres.add_slide(), pres.add_slide()
    pres.add_chart(s1, "line", ["Jan", "Feb"], {"Visits": [10, 14], "Orders": [3, 5]},
                   title="Traffic", y_title="count")
    pres.add_chart(s2, "doughnut", ["a", "b", "c"], {"Mix": [1, 2, 3]}, labels="percent")
    pres.add_chart(s2, "scatter", [1, 2, 3], {"y": [2, 4, 9]}, lines=True, smooth=True)
    assert _all_parts_valid(pres) == []
    back = _reload(pres)
    a, b = back.slides
    (line,) = back.charts(a)
    assert line.kind == "line_markers" and line.y_title == "count"
    assert [s.name for s in line.series] == ["Visits", "Orders"]
    assert [c.kind for c in back.charts(b)] == ["doughnut", "scatter_smooth"]
    assert back.charts(b)[1].categories == [1, 2, 3]
    assert sorted(p for p in back.package.parts if "/charts/" in p or "/embeddings/" in p) == [
        "/ppt/charts/chart1.xml", "/ppt/charts/chart2.xml", "/ppt/charts/chart3.xml",
        "/ppt/embeddings/package1.xlsx", "/ppt/embeddings/package2.xlsx",
        "/ppt/embeddings/package3.xlsx"]


def test_embedded_series_input_forms():
    from oox.charts import normalize_series

    assert [s.name for s in normalize_series({"a": [1], "b": [2]})] == ["a", "b"]
    assert normalize_series([("a", (1, 2))])[0].values == [1, 2]
    s = Series("z", [1])
    assert normalize_series([s])[0] is s


# -- combination charts and the secondary axis -----------------------------------------
def test_column_and_line_on_a_secondary_axis():
    series = [Series("Revenue", [120, 135, 150]), Series("Margin", [0.12, 0.15, 0.11])]
    space = build_chart("column", ["Q1", "Q2", "Q3"], series, kinds={"Margin": "line"},
                        secondary=["Margin"], y_title="EUR k", y2_title="Margin",
                        y2_number_format="0%", title="Revenue and margin")
    assert space.validate(target="Office2007") == []
    assert space.validate(level="full") == []
    info = chart_info(space)
    assert info.plot_kinds == ["column", "line_markers"]
    assert [(s.name, s.kind, s.axis) for s in info.series] == [
        ("Revenue", "column", "primary"), ("Margin", "line_markers", "secondary")]
    assert (info.y_title, info.y2_title, info.x_title) == ("EUR k", "Margin", None)
    # series keep their index and order, so the legend keeps the given order
    idx = [s.first("c:idx").get("val") for s in space.iter(f"{{{C}}}ser")]
    assert idx == ["0", "1"]


def test_areas_behind_columns_behind_lines_on_one_axis():
    series = [Series("L", [1, 2, 3], kind="line"), Series("C", [2, 3, 4]),
              Series("A", [3, 4, 5], kind="area")]
    info = chart_info(build_chart("column", ["a", "b", "c"], series))
    assert info.plot_kinds == ["area", "column", "line_markers"]
    assert {s.axis for s in info.series} == {"primary"}
    assert [s.name for s in info.series] == ["A", "C", "L"]     # as drawn


def test_secondary_by_index_and_same_kind_on_both_axes():
    series = [Series("Big", [1000, 2000]), Series("Small", [1, 2])]
    space = build_chart("line", ["x", "y"], series, secondary=[1])
    info = chart_info(space)
    assert info.plot_kinds == ["line_markers", "line_markers"]
    assert [s.axis for s in info.series] == ["primary", "secondary"]
    assert info.y2_title is None


@pytest.mark.parametrize("kind,options,message", [
    ("pie", {"kinds": {0: "line"}}, "doesn't combine"),
    ("scatter", {"secondary": [0]}, "doesn't combine"),
    ("bar", {"kinds": {1: "line"}}, "bars combine only with bars"),
    ("column", {"kinds": {1: "pie"}}, "don't combine"),
    ("column", {"kinds": {"Nope": "line"}}, "no series"),
    ("column", {"secondary": [0, 1]}, "primary axis"),
    ("column", {"y2_title": "x"}, "no series is on the secondary axis"),
])
def test_combination_errors(kind, options, message):
    series = [Series("A", [1, 2]), Series("B", [3, 4])]
    if kind == "pie":
        series = series[:1]
    with pytest.raises(ValueError, match=message):
        build_chart(kind, [1, 2] if kind == "scatter" else ["x", "y"], series, **options)


def test_excel_combination_chart_from_a_range():
    book = SpreadsheetDocument.new("Data")
    book.append_row("Data", ["Month", "Units", "Price"])
    for row in (["Jan", 10, 2.5], ["Feb", 14, 2.25], ["Mar", 9, 2.75]):
        book.append_row("Data", row)
    book.add_chart("Data", "column", "A1:C4", kinds={"Price": "line"}, secondary=["Price"],
                   y2_title="EUR")
    again = _reload(book)
    assert _all_parts_valid(again) == []
    info = again.charts("Data")[0]
    assert [(s.name, s.kind, s.axis) for s in info.series] == [
        ("Units", "column", "primary"), ("Price", "line_markers", "secondary")]
    assert info.series[1].values == [2.5, 2.25, 2.75]


def test_word_and_powerpoint_combination_charts():
    from oox import PresentationDocument, WordprocessingDocument

    doc = WordprocessingDocument.new()
    doc.add_chart("column", ["a", "b"], {"Units": [3, 4], "Share": [0.4, 0.6]},
                  kinds={"Share": "line"}, secondary=["Share"])
    deck = PresentationDocument.new(layouts="basic")
    slide = deck.add_slide("Blank")
    deck.add_chart(slide, "area", ["a", "b"], {"Units": [3, 4], "Share": [0.4, 0.6]},
                   kinds={"Units": "column"}, secondary=["Share"])
    w_info = _reload(doc).charts()[0]
    deck2 = _reload(deck)
    p_info = deck2.charts(deck2.slides[0])[0]
    assert [s.axis for s in w_info.series] == ["primary", "secondary"]
    assert [(s.kind, s.axis) for s in p_info.series] == [
        ("column", "primary"), ("area", "secondary")]


def test_excel_charts_without_caches_read_their_cells():
    """openpyxl writes references only; the values come from the sheet."""
    book = SpreadsheetDocument.new("S")
    book.append_row("S", ["k", "v"])
    book.append_row("S", ["a", 1])
    book.append_row("S", ["b", 2])
    book.add_chart("S", "column", "A1:B3")
    space = next(p.element for p in book.package.parts.values()
                 if p.content_type.endswith("chart+xml"))
    for cache in list(space.iter(f"{{{C}}}numCache", f"{{{C}}}strCache")):
        cache.getparent().remove(cache)
    info = book.charts("S")[0]
    assert info.series[0].name == "v"
    assert info.series[0].values == [1, 2]
    assert info.categories == ["a", "b"]


# -- chart sheets ------------------------------------------------------------------------
def test_chart_sheet_is_a_tab_but_not_a_worksheet():
    book = SpreadsheetDocument.new("Data")
    for row in (["k", "v"], ["a", 1], ["b", 2]):
        book.append_row("Data", row)
    book.add_chart_sheet("Chart", "pie", "Data", "A1:B3", title="Share", labels="percent")
    book.add_sheet("Notes")
    book.define_name("Here", "Notes", "A1", local=True)
    assert book.tab_names == ["Data", "Chart", "Notes"]
    assert book.sheet_names == ["Data", "Notes"]
    assert book.chart_sheet_names == ["Chart"]
    with pytest.raises(KeyError, match="chart sheet"):
        book.sheet("Chart")
    with pytest.raises(ValueError, match="duplicate"):
        book.add_sheet("chart")
    with pytest.raises(ValueError, match="duplicate"):
        book.add_chart_sheet("Data", "pie", "Data", "A1:B3")
    again = _reload(book)
    assert again.validate(level="full") == {}
    info = again.charts("Chart")[0]
    assert (info.kind, info.title, info.anchor) == ("pie", "Share", None)
    assert info.series[0].values == [1, 2]
    assert again.defined_names() == {"Notes!Here": "Notes!$A$1"}   # localSheetId counts tabs
    assert again.chart_sheet("Chart").tag.endswith("}chartsheet")


def test_workbook_readers_skip_chart_sheets_from_other_writers(tmp_path):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    path = tmp_path / "cs.xlsx"
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Data")
    ws.write_column("A1", [1, 2, 3])
    cs = wb.add_chartsheet("Graph")
    chart = wb.add_chart({"type": "line"})
    chart.add_series({"values": "=Data!$A$1:$A$3"})
    cs.set_chart(chart)
    wb.close()
    book = open_document(path)
    assert book.sheet_names == ["Data"] and book.tab_names == ["Data", "Graph"]
    assert [c.kind for c in book.charts("Graph")] == ["line"]
    from tools.interop import check

    assert check(str(path))["ok"]
