"""Charts: DrawingML chart parts (``c:chartSpace``), shared by Excel, Word and
PowerPoint.

    from oox.charts import Series, build_chart, chart_info

    space = build_chart("column", ["Q1", "Q2", "Q3"],
                        [Series("North", [120, 135, 150]), Series("South", [95, 110, 99])],
                        title="Sales", y_title="EUR k", legend="bottom", labels="value")
    chart_info(space)["type"]        # "column"

The applications wrap this: ``SpreadsheetDocument.add_chart`` (series from a
cell range), ``WordprocessingDocument.add_chart`` and
``PresentationDocument.add_chart`` (series from Python lists, stored in an
embedded workbook so "Edit Data" works).  Every series carries cached values,
so viewers draw the chart without recalculating anything.

Kinds and groupings:

=========  ====================================================  ==============
kind       grouping                                              Excel type
=========  ====================================================  ==============
column     clustered (default), stacked, percent                 51, 52, 53
bar        clustered (default), stacked, percent (horizontal)    57, 58, 59
line       standard (default), stacked, percent; markers=        65 / 4, ...
area       standard (default), stacked, percent                  1, 76, 77
pie        (one series; colors= per slice)                       5
doughnut   (hole_size=)                                          -4120
scatter    lines=, markers=, smooth=  (categories are x values)  -4169, 74, ...
=========  ====================================================  ==============

``chart_info`` reports the type in the same vocabulary for any chart part --
oox's own, Excel's or LibreOffice's -- so it can be compared with
``Chart.ChartType`` in Office (see ``EXCEL_CHART_TYPES``).
"""

from __future__ import annotations

import io
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass

from .element import OoxElement
from .grammar import URI
from .ns import a, c

C = URI["c"]
A = URI["a"]
CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"
KINDS = ("column", "bar", "line", "area", "pie", "doughnut", "scatter")
LEGEND = {"right": "r", "left": "l", "top": "t", "bottom": "b", "top_right": "tr"}
LABELS = ("value", "percent", "category", "series")
CAT_AX, VAL_AX = 500000001, 500000002
CAT2_AX, VAL2_AX = 500000003, 500000004   # secondary axes
COMBINABLE = ("column", "line", "area")
RT_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
# Excel's XlChartType -> chart_info()["type"]
EXCEL_CHART_TYPES = {
    51: "column", 52: "column_stacked", 53: "column_percent",
    57: "bar", 58: "bar_stacked", 59: "bar_percent",
    4: "line", 65: "line_markers", 63: "line_stacked", 66: "line_markers_stacked",
    64: "line_percent", 67: "line_markers_percent",
    1: "area", 76: "area_stacked", 77: "area_percent",
    5: "pie", -4120: "doughnut",
    -4169: "scatter", 74: "scatter_lines", 75: "scatter_lines_no_markers",
    72: "scatter_smooth", 73: "scatter_smooth_no_markers",
}
# The default theme's accent colours (oox.drawing.default_theme), darker and
# lighter variants after six series.  Written explicitly: viewers other than
# Office draw "automatic" series without a chart style as empty.
PALETTE = ("4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
           "264478", "9E480E", "636363", "997300", "255E91", "43682B",
           "698ED0", "F1975A", "B7B7B7", "FFCD33", "7CAFDD", "8CC168")
_GROUPING = {"clustered": "clustered", "standard": "standard", "stacked": "stacked",
             "percent": "percentStacked"}


@dataclass
class Series:
    """One data series.  ``categories`` overrides the chart's categories (for
    scatter: the x values).  The ``*_ref`` fields are cell references
    (``Sheet1!$B$2:$B$5``); without them the values are written as literals."""

    name: str | None
    values: Sequence[float | int | None]
    categories: Sequence | None = None
    color: str | None = None
    name_ref: str | None = None
    values_ref: str | None = None
    categories_ref: str | None = None
    kind: str | None = None          # another kind than the chart's (combination charts)
    axis: str = "primary"            # or "secondary": its own value axis on the right


# -- building -----------------------------------------------------------------------
def _el(factory, *kids, **attrs) -> OoxElement:
    """``factory(*kids)`` without the ``None`` entries."""
    return factory(*[k for k in kids if k is not None], **attrs)


def _num(v: float | int) -> str:
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return str(int(v))
    return repr(v) if isinstance(v, float) else str(v)


def _is_number(v: object) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _num_data(values: Sequence, ref: str | None, fmt: str | None) -> OoxElement:
    """``c:numRef`` (with cache) or ``c:numLit``."""
    pts = [c.Pt(c.V(text=_num(v)), idx=i) for i, v in enumerate(values) if v is not None]
    if ref:
        return c.NumRef(c.F(text=ref), c.NumCache(c.FormatCode(text=fmt or "General"),
                                                  c.PtCount(val=len(values)), *pts))
    return c.NumLit(c.FormatCode(text=fmt or "General"), c.PtCount(val=len(values)), *pts)


def _str_data(values: Sequence, ref: str | None) -> OoxElement:
    pts = [c.Pt(c.V(text=str(v)), idx=i) for i, v in enumerate(values) if v is not None]
    if ref:
        return c.StrRef(c.F(text=ref), c.StrCache(c.PtCount(val=len(values)), *pts))
    return c.StrLit(c.PtCount(val=len(values)), *pts)


def _axis_data(wrapper, values: Sequence, ref: str | None) -> OoxElement:
    """``c:cat`` / ``c:xVal``: numbers stay numbers, anything else is text."""
    if values and all(_is_number(v) or v is None for v in values):
        return wrapper(_num_data(values, ref, None))
    return wrapper(_str_data(["" if v is None else v for v in values], ref))


def _rich(text: str, rot: int | None = None, size: int | None = None,
          bold: bool | None = None) -> OoxElement:
    def drop(**kw):
        return {k: v for k, v in kw.items() if v is not None}

    rpr = a.RPr(lang="en-US", **drop(sz=size, b=bold))
    body = a.BodyPr(**drop(rot=rot, vert="horz" if rot is not None else None))
    return c.Rich(body, a.LstStyle(),
                  a.P(a.PPr(a.DefRPr(**drop(sz=size, b=bold))), a.R(rpr, a.T(text=text))))


def _title(text: str, rot: int | None = None, size: int | None = None) -> OoxElement:
    return c.Title(c.Tx(_rich(text, rot, size)), c.Overlay(val=False))


def _solid(color: str) -> OoxElement:
    return a.SolidFill(a.SrgbClr(val=color.lstrip("#").upper()))


def _series_name(s: Series) -> OoxElement | None:
    if s.name is None:
        return None
    if s.name_ref:
        return c.Tx(_str_data([s.name], s.name_ref))
    return c.Tx(c.V(text=s.name))


def _labels(labels, fmt: str | None) -> OoxElement | None:
    """``c:dLbls`` showing the requested parts ("value", "percent", ...)."""
    if not labels:
        return None
    wanted = {labels} if isinstance(labels, str) else set(labels)
    bad = wanted - set(LABELS)
    if bad:
        raise ValueError(f"labels: unknown {sorted(bad)}; use {LABELS}")
    parts = []
    if fmt:
        parts.append(c.NumFmt(formatCode=fmt, sourceLinked=False))
    parts += [c.ShowLegendKey(val=False), c.ShowVal(val="value" in wanted),
              c.ShowCatName(val="category" in wanted), c.ShowSerName(val="series" in wanted),
              c.ShowPercent(val="percent" in wanted), c.ShowBubbleSize(val=False)]
    return c.DLbls(*parts)


def _series(kind: str, i: int, s: Series, categories: Sequence, categories_ref: str | None,
            *, order: int | None = None, color: str | None, point_colors: Sequence[str] | None, markers: bool,
            lines: bool, smooth: bool, fmt: str | None) -> OoxElement:
    cats = s.categories if s.categories is not None else categories
    cref = s.categories_ref if s.categories is not None else categories_ref
    if len(cats) != len(s.values):
        raise ValueError(f"series {s.name!r}: {len(s.values)} values for {len(cats)} categories")
    kids: list[OoxElement | None] = [c.Idx(val=i), c.Order(val=i if order is None else order),
                                     _series_name(s)]
    color = s.color or color
    if kind in ("line", "scatter"):
        ln = (a.Ln(_solid(color), w=28575) if color else None) if lines else a.Ln(a.NoFill())
        kids.append(c.SpPr(ln) if ln is not None else None)
        if not markers:
            kids.append(c.Marker(c.Symbol(val="none")))
        elif color:
            kids.append(c.Marker(c.Symbol(val="circle"), c.Size(val=6),
                                 c.SpPr(_solid(color), a.Ln(_solid(color)))))
    elif color:
        kids.append(c.SpPr(_solid(color)))
    if kind in ("column", "bar"):
        kids.append(c.InvertIfNegative(val=False))
    for j, pc in enumerate(point_colors or ()):
        if j < len(s.values):
            kids.append(c.DPt(c.Idx(val=j), c.Bubble3D(val=False), c.SpPr(_solid(pc))))
    if kind == "scatter":
        kids += [_axis_data(c.XVal, cats, cref), c.YVal(_num_data(s.values, s.values_ref, fmt)),
                 c.Smooth(val=smooth)]
    else:
        kids += [_axis_data(c.Cat, cats, cref), c.Val(_num_data(s.values, s.values_ref, fmt))]
        if kind == "line":
            kids.append(c.Smooth(val=smooth))
    return _el(c.Ser, *kids)


def _secondary_axes(kind: str, *, title: str | None, fmt: str | None,
                    y_min: float | None, y_max: float | None) -> list[OoxElement]:
    """A hidden category axis and a value axis on the right (top for bars)
    crossing it at its maximum, for the series on the secondary axis."""
    horizontal = kind == "bar"
    rot = None if horizontal else -5400000
    cat = c.CatAx(c.AxId(val=CAT2_AX), c.Scaling(c.Orientation(val="minMax")),
                  c.Delete(val=True), c.AxPos(val="l" if horizontal else "b"),
                  c.MajorTickMark(val="out"), c.MinorTickMark(val="none"),
                  c.TickLblPos(val="nextTo"), c.CrossAx(val=VAL2_AX), c.Crosses(val="autoZero"),
                  c.Auto(val=True), c.LblAlgn(val="ctr"), c.LblOffset(val=100),
                  c.NoMultiLvlLbl(val=False))
    val = _el(c.ValAx, c.AxId(val=VAL2_AX),
              _el(c.Scaling, c.Orientation(val="minMax"),
                  c.Max(val=y_max) if y_max is not None else None,
                  c.Min(val=y_min) if y_min is not None else None),
              c.Delete(val=False), c.AxPos(val="t" if horizontal else "r"),
              _title(title, rot) if title else None,
              c.NumFmt(formatCode=fmt or "General", sourceLinked=fmt is None),
              c.MajorTickMark(val="out"), c.MinorTickMark(val="none"),
              c.TickLblPos(val="nextTo"), c.CrossAx(val=CAT2_AX), c.Crosses(val="max"),
              c.CrossBetween(val="midCat" if kind == "area" else "between"))
    return [val, cat]


def _axes(kind: str, *, x_title: str | None, y_title: str | None, fmt: str | None,
          gridlines: bool, y_min: float | None, y_max: float | None) -> list[OoxElement]:
    horizontal = kind == "bar"
    cat_pos, val_pos = ("l", "b") if horizontal else ("b", "l")

    def scaling(lo=None, hi=None):
        return _el(c.Scaling, c.Orientation(val="minMax"),
                   c.Max(val=hi) if hi is not None else None,
                   c.Min(val=lo) if lo is not None else None)

    y_rot = None if horizontal else -5400000
    x_rot = -5400000 if horizontal else None
    val_kids = [c.AxId(val=VAL_AX), scaling(y_min, y_max), c.Delete(val=False),
                c.AxPos(val=val_pos),
                c.MajorGridlines() if gridlines else None,
                _title(y_title, y_rot) if y_title else None,
                c.NumFmt(formatCode=fmt or "General", sourceLinked=fmt is None),
                c.MajorTickMark(val="out"), c.MinorTickMark(val="none"),
                c.TickLblPos(val="nextTo"), c.CrossAx(val=CAT_AX), c.Crosses(val="autoZero"),
                c.CrossBetween(val="midCat" if kind in ("scatter", "area") else "between")]
    if kind == "scatter":
        x_ax = _el(c.ValAx, c.AxId(val=CAT_AX), scaling(), c.Delete(val=False), c.AxPos(val="b"),
                       _title(x_title) if x_title else None,
                       c.NumFmt(formatCode="General", sourceLinked=True),
                       c.MajorTickMark(val="out"), c.MinorTickMark(val="none"),
                       c.TickLblPos(val="nextTo"), c.CrossAx(val=VAL_AX),
                       c.Crosses(val="autoZero"), c.CrossBetween(val="midCat"))
    else:
        x_ax = _el(c.CatAx, c.AxId(val=CAT_AX), scaling(), c.Delete(val=False), c.AxPos(val=cat_pos),
                       _title(x_title, x_rot) if x_title else None,
                       c.NumFmt(formatCode="General", sourceLinked=True),
                       c.MajorTickMark(val="out"), c.MinorTickMark(val="none"),
                       c.TickLblPos(val="nextTo"), c.CrossAx(val=VAL_AX),
                       c.Crosses(val="autoZero"), c.Auto(val=True), c.LblAlgn(val="ctr"),
                       c.LblOffset(val=100), c.NoMultiLvlLbl(val=False))
    return [x_ax, _el(c.ValAx, *val_kids)]


_Z_ORDER = {"area": 0, "column": 1, "bar": 1, "line": 2}


def _combination(kind: str, series: list[Series], kinds, secondary
                 ) -> list[tuple[str, str, list[tuple[int, Series]]]]:
    """Group the series into plots: ``[(kind, axis, [(index, series)])]``,
    areas behind columns behind lines, primary before secondary."""
    def key_of(i: int, s: Series, keys) -> bool:
        return i in keys or (s.name is not None and s.name in keys)

    kinds = dict(kinds or {})
    secondary = set(secondary)
    names = {s.name for s in series} | set(range(len(series)))
    for given, what in ((kinds, "kinds"), (secondary, "secondary")):
        unknown = [k for k in given if k not in names]
        if unknown:
            raise ValueError(f"{what}: no series {unknown}")
    groups: dict[tuple[str, str], list[tuple[int, Series]]] = {}
    for i, s in enumerate(series):
        k = next((v for key, v in kinds.items() if key_of(i, s, {key})), None) or s.kind or kind
        axis = "secondary" if key_of(i, s, secondary) else s.axis
        if axis not in ("primary", "secondary"):
            raise ValueError(f"series {s.name!r}: axis={axis!r}: use 'primary' or 'secondary'")
        if k not in KINDS:
            raise ValueError(f"series {s.name!r}: kind {k!r}: use one of {KINDS}")
        groups.setdefault((k, axis), []).append((i, s))
    used = {k for k, _ in groups}
    secondary_used = any(axis == "secondary" for _, axis in groups)
    if used != {kind} or secondary_used:
        if kind in ("pie", "doughnut", "scatter"):
            raise ValueError(f"a {kind} chart doesn't combine kinds or use a secondary axis")
        if "bar" in used and used != {"bar"}:
            raise ValueError("bars combine only with bars (columns, lines and areas combine)")
        bad = used - set(COMBINABLE) - {"bar"}
        if bad:
            raise ValueError(f"{sorted(bad)} don't combine; use {COMBINABLE}")
    if all(axis == "secondary" for _, axis in groups):
        raise ValueError("at least one series must stay on the primary axis")
    order = sorted(groups, key=lambda ka: (ka[1] == "secondary", _Z_ORDER.get(ka[0], 1)))
    return [(k, axis, groups[(k, axis)]) for k, axis in order]


def _plot(kind: str, grouping: str, ser: list[OoxElement], dl: OoxElement | None,
          ids: tuple[int, int], *, gap_width: int, first_slice: int, hole_size: int
          ) -> OoxElement:
    labels = [dl] if dl is not None else []
    axes_ids = [c.AxId(val=ids[0]), c.AxId(val=ids[1])]
    if kind in ("column", "bar"):
        stacked = grouping in ("stacked", "percent")
        return c.BarChart(c.BarDir(val="bar" if kind == "bar" else "col"),
                          c.Grouping(val=_GROUPING[grouping]), c.VaryColors(val=False), *ser,
                          *labels, c.GapWidth(val=gap_width),
                          *([c.Overlap(val=100)] if stacked else []), *axes_ids)
    if kind == "line":
        return c.LineChart(c.Grouping(val=_GROUPING[grouping]), c.VaryColors(val=False), *ser,
                           *labels, c.Marker(val=True), *axes_ids)
    if kind == "area":
        return c.AreaChart(c.Grouping(val=_GROUPING[grouping]), c.VaryColors(val=False), *ser,
                           *labels, *axes_ids)
    if kind == "scatter":
        return c.ScatterChart(c.ScatterStyle(val="lineMarker"), c.VaryColors(val=False), *ser,
                              *labels, *axes_ids)
    if kind == "pie":
        return c.PieChart(c.VaryColors(val=True), *ser, *labels, c.FirstSliceAng(val=first_slice))
    return c.DoughnutChart(c.VaryColors(val=True), *ser, *labels,
                           c.FirstSliceAng(val=first_slice), c.HoleSize(val=hole_size))


def build_chart(kind: str, categories: Sequence, series: Iterable[Series], *,
                categories_ref: str | None = None, grouping: str | None = None,
                title: str | None = None, x_title: str | None = None,
                y_title: str | None = None, legend: str | None = "right",
                labels=None, label_format: str | None = None,
                colors: Sequence[str] | None = None, number_format: str | None = None,
                markers: bool = True, lines: bool | None = None, smooth: bool = False,
                gridlines: bool = True, y_min: float | None = None, y_max: float | None = None,
                hole_size: int = 50, gap_width: int = 150, first_slice: int = 0,
                kinds: Mapping[str | int, str] | None = None,
                secondary: Iterable[str | int] = (), y2_title: str | None = None,
                y2_min: float | None = None, y2_max: float | None = None,
                y2_number_format: str | None = None,
                external_rid: str | None = None) -> OoxElement:
    """A complete ``c:chartSpace``.

    ``colors`` are RRGGBB strings: one per series, or one per slice for pie
    and doughnut; missing ones come from ``PALETTE``.  ``labels`` is one of / a list of "value", "percent",
    "category", "series".  ``number_format`` applies to the value axis and the
    cached values; ``label_format`` to the data labels.  ``lines`` (scatter)
    defaults to False: markers only.  ``external_rid`` links the embedded
    workbook (Word, PowerPoint).

    Combination charts: ``kinds`` maps series (by name or 0-based index) to
    another kind, e.g. ``{"Margin": "line"}`` on a column chart; ``secondary``
    names the series drawn against a second value axis on the right, titled
    ``y2_title``.  Columns, lines and areas combine; bars only with bars.
    A ``Series`` can also carry ``kind`` and ``axis`` itself.
    """
    if kind not in KINDS:
        raise ValueError(f"chart kind {kind!r}: use one of {KINDS}")
    series = list(series)
    if not series:
        raise ValueError("a chart needs at least one series")
    pie = kind in ("pie", "doughnut")
    if pie and len(series) > 1 and kind == "pie":
        raise ValueError("a pie chart shows one series")
    if legend is not None and legend not in LEGEND:
        raise ValueError(f"legend={legend!r}: use one of {sorted(LEGEND)} or None")
    if lines is None:
        lines = kind != "scatter"
    if grouping is None:
        grouping = "clustered" if kind in ("column", "bar") else "standard"
    if grouping not in _GROUPING:
        raise ValueError(f"grouping={grouping!r}: use one of {sorted(_GROUPING)}")
    if kind in ("column", "bar") and grouping == "standard":
        grouping = "clustered"
    if kind in ("line", "area") and grouping == "clustered":
        grouping = "standard"

    def pick(i: int) -> str:
        if colors and i < len(colors):
            return colors[i]
        return PALETTE[i % len(PALETTE)]

    groups = _combination(kind, series, kinds, secondary)
    if y2_title is not None and not any(axis == "secondary" for _, axis, _ in groups):
        raise ValueError("y2_title: no series is on the secondary axis")
    n_points = max(len(s.values) for s in series)
    plots: list[OoxElement] = []
    position = 0     # c:order follows the plots (areas, columns, lines), as Office and
    for k, axis, members in groups:     # LibreOffice list them; c:idx keeps the colour
        ser = [_series(k, i, s, categories, categories_ref, order=position + n,
                       color=None if pie else pick(i),
                       point_colors=[pick(j) for j in range(n_points)] if pie else None,
                       markers=markers, lines=lines, smooth=smooth,
                       fmt=y2_number_format if axis == "secondary" else number_format)
               for n, (i, s) in enumerate(members)]
        position += len(members)
        g = grouping if k == kind else ("clustered" if k in ("column", "bar") else "standard")
        if k in ("line", "area") and g == "clustered":
            g = "standard"
        ids = (CAT2_AX, VAL2_AX) if axis == "secondary" else (CAT_AX, VAL_AX)
        plots.append(_plot(k, g, ser, _labels(labels, label_format), ids, gap_width=gap_width,
                           first_slice=first_slice, hole_size=hole_size))
    area = c.PlotArea(c.Layout(), *plots)
    if not pie:
        for ax in _axes(kind, x_title=x_title, y_title=y_title, fmt=number_format,
                        gridlines=gridlines, y_min=y_min, y_max=y_max):
            area.append(ax)
        if any(axis == "secondary" for _, axis, _ in groups):
            for ax in _secondary_axes(kind, title=y2_title, fmt=y2_number_format,
                                      y_min=y2_min, y_max=y2_max):
                area.append(ax)
    chart = c.Chart(*([_title(title, size=1400)] if title else []),
                    c.AutoTitleDeleted(val=title is None), area)
    if legend is not None:
        chart.append(c.Legend(c.LegendPos(val=LEGEND[legend]), c.Overlay(val=False)))
    chart.append(c.PlotVisOnly(val=True))
    chart.append(c.DispBlanksAs(val="gap"))
    space = c.ChartSpace(c.Date1904(val=False), c.Lang(val="en-US"),
                         c.RoundedCorners(val=False), chart)
    if external_rid:
        space.append(c.ExternalData(c.AutoUpdate(val=False), r_id=external_rid))
    return space


# -- reading ------------------------------------------------------------------------
def _q(local: str, ns: str = C) -> str:
    return f"{{{ns}}}{local}"


def _val(el, local: str, default=None):
    if el is None:
        return default
    child = el.find(_q(local))
    return child.get("val", default) if child is not None else default


def _on(el, local: str, default: bool = False) -> bool:
    v = _val(el, local)
    if v is None:
        return default if el is None or el.find(_q(local)) is None else True
    return v in ("1", "true")


def _text(el) -> str | None:
    """Text of a ``c:title`` / ``c:tx``: rich runs or a string reference cache."""
    if el is None:
        return None
    runs = [t.text or "" for t in el.iter(_q("t", A))]
    if runs:
        return "".join(runs)
    vs = [v.text or "" for v in el.iter(_q("v"))]
    return "".join(vs) if vs else None


def _points(data) -> list:
    """Values of a numRef/numLit/strRef/strLit (cache), with None for gaps."""
    if data is None:
        return []
    cache = data.find(_q("numCache")) if data.tag == _q("numRef") else \
        data.find(_q("strCache")) if data.tag == _q("strRef") else data
    if cache is None:
        return []
    count = int(_val(cache, "ptCount", "0"))
    out: list = [None] * count
    numeric = data.tag in (_q("numRef"), _q("numLit"))
    for pt in cache.findall(_q("pt")):
        idx = int(pt.get("idx", "0"))
        v = pt.find(_q("v"))
        if v is None or idx >= count:
            continue
        text = v.text or ""
        if numeric:
            try:
                f = float(text)
                out[idx] = int(f) if f.is_integer() and "." not in text and "E" not in text.upper() else f
            except ValueError:
                out[idx] = text
        else:
            out[idx] = text
    return out


def _ref(data) -> str | None:
    if data is None:
        return None
    f = data.find(_q("f"))
    return f.text if f is not None else None


def _first_data(wrapper):
    if wrapper is None:
        return None
    for child in wrapper:
        if isinstance(child.tag, str):
            return child
    return None


def _series_info(ser) -> dict:
    tx = ser.find(_q("tx"))
    cat = ser.find(_q("cat"))
    if cat is None:
        cat = ser.find(_q("xVal"))
    val = ser.find(_q("val"))
    if val is None:
        val = ser.find(_q("yVal"))
    cdata, vdata = _first_data(cat), _first_data(val)
    color = None
    sppr = ser.find(_q("spPr"))
    if sppr is not None:
        clr = sppr.find(f"{_q('solidFill', A)}/{_q('srgbClr', A)}")
        if clr is None:
            clr = sppr.find(f"{_q('ln', A)}/{_q('solidFill', A)}/{_q('srgbClr', A)}")
        color = clr.get("val") if clr is not None else None
    points = []
    for dpt in ser.findall(_q("dPt")):
        clr = dpt.find(f"{_q('spPr')}/{_q('solidFill', A)}/{_q('srgbClr', A)}")
        if clr is not None:
            points.append(clr.get("val"))
    return {"name": _text(tx), "values": _points(vdata), "categories": _points(cdata),
            "values_ref": _ref(vdata), "categories_ref": _ref(cdata), "color": color,
            "point_colors": points, "name_ref": _ref(_first_data(tx)) if tx is not None else None}


def _has_lines(ser) -> bool:
    ln = ser.find(f"{_q('spPr')}/{_q('ln', A)}")
    return ln is None or ln.find(_q("noFill", A)) is None


def _has_markers(ser, chart_level: bool = True) -> bool:
    sym = ser.find(f"{_q('marker')}/{_q('symbol')}")
    if sym is not None:
        return sym.get("val") != "none"
    return chart_level


def _type_of(plot) -> str:
    local = plot.tag.split("}")[1]
    grouping = _val(plot, "grouping", "standard")
    suffix = {"stacked": "_stacked", "percentStacked": "_percent"}.get(grouping, "")
    sers = plot.findall(_q("ser"))
    if local == "barChart":
        base = "bar" if _val(plot, "barDir", "col") == "bar" else "column"
        return base + suffix
    if local == "lineChart":
        markers = any(_has_markers(s, _on(plot, "marker", True)) for s in sers)
        return ("line_markers" if markers else "line") + suffix
    if local == "areaChart":
        return "area" + suffix
    if local == "pieChart":
        return "pie"
    if local == "doughnutChart":
        return "doughnut"
    if local == "scatterChart":
        lines = any(_has_lines(s) for s in sers)
        markers = any(_has_markers(s) for s in sers)
        smooth = any(_on(s, "smooth") for s in sers)
        if not lines:
            return "scatter"
        if smooth:
            return "scatter_smooth" if markers else "scatter_smooth_no_markers"
        return "scatter_lines" if markers else "scatter_lines_no_markers"
    return local


def _shown_labels(dl) -> list[str]:
    return sorted(name for name, tag in (("value", "showVal"), ("percent", "showPercent"),
                                         ("category", "showCatName"), ("series", "showSerName"))
                  if _val(dl, tag) in ("1", "true"))


def chart_info(space: OoxElement) -> dict:
    """Describe a ``c:chartSpace``: type (Excel's vocabulary, see
    ``EXCEL_CHART_TYPES``), title, legend position, axis titles, value number
    format, data labels, and per series its name, values, categories, refs,
    colour, "type" (its plot's, for combination charts) and "axis" ("primary"
    or "secondary"; the secondary value axis's title is "y2_title").  Series
    come in plot order (``c:order``)."""
    chart = space.find(_q("chart"))
    area = chart.find(_q("plotArea")) if chart is not None else None
    plots = [el for el in (area if area is not None else [])
             if isinstance(el.tag, str) and el.tag.endswith("Chart")]
    title_el = chart.find(_q("title")) if chart is not None else None
    title = _text(title_el.find(_q("tx"))) if title_el is not None else None
    legend = chart.find(_q("legend")) if chart is not None else None
    legend_pos = None
    if legend is not None:
        code = _val(legend, "legendPos", "r")
        legend_pos = {v: k for k, v in LEGEND.items()}.get(code, code)
    axes: dict[str, dict] = {}
    for ax in (area if area is not None else []):
        if not isinstance(ax.tag, str) or not ax.tag.endswith("Ax"):
            continue
        t = ax.find(_q("title"))
        fmt = ax.find(_q("numFmt"))
        axes[_val(ax, "axId", "")] = {
            "kind": ax.tag.split("}")[1], "pos": _val(ax, "axPos", "b"),
            "title": _text(t.find(_q("tx"))) if t is not None else None,
            "format": fmt.get("formatCode") if fmt is not None
            and fmt.get("sourceLinked") in ("0", "false")
            and fmt.get("formatCode") != "General" else None,
        }
    plot_ids = [[a.get("val") for a in p.findall(_q("axId"))] for p in plots]
    primary = next((ids for ids in plot_ids if ids), [])

    def x_and_y(ids: list) -> tuple[dict, dict]:
        """The category (or scatter x) axis and the value axis of a plot."""
        found = [axes[i] for i in ids if i in axes]
        cat = [a for a in found if a["kind"] in ("catAx", "dateAx")]
        if cat:
            x = cat[0]
        else:   # scatter: two value axes, x is the horizontal one
            x = next((a for a in found if a["pos"] in ("b", "t")), found[0] if found else {})
        y = next((a for a in found if a is not x and a["kind"] == "valAx"), {})
        return x, y

    x_axis, y_axis = x_and_y(primary)
    y2_axis: dict = {}
    series: list = []
    types = []
    for p, ids in zip(plots, plot_ids, strict=True):
        kind = _type_of(p)
        types.append(kind)
        on = "secondary" if ids and ids != primary else "primary"
        if on == "secondary" and not y2_axis:
            y2_axis = x_and_y(ids)[1]
        for s in p.findall(_q("ser")):
            info = _series_info(s)
            info["type"], info["axis"] = kind, on
            order = _val(s, "order")
            series.append((int(order) if order and order.isdigit() else len(series), info))
    # in plot order (c:order), as the legend and Select Data list them; combination
    # charts keep their series in one plot per kind
    series = [info for _, info in sorted(series, key=lambda oi: oi[0])]
    # labels for the whole chart type, or (LibreOffice) on the first series
    dl = plots[0].find(_q("dLbls")) if plots else None
    if (dl is None or not _shown_labels(dl)) and plots:
        first = plots[0].find(_q("ser"))
        if first is not None and first.find(_q("dLbls")) is not None:
            dl = first.find(_q("dLbls"))
    labels = _shown_labels(dl) if dl is not None else []
    return {"type": types[0] if types else None, "types": types, "title": title,
            "legend": legend_pos, "x_title": x_axis.get("title"), "y_title": y_axis.get("title"),
            "y2_title": y2_axis.get("title"),
            "number_format": y_axis.get("format"), "labels": labels, "series": series,
            "categories": series[0]["categories"] if series else []}


def chart_spaces_of(doc, part) -> list[tuple[str, OoxElement]]:
    """``(relationship id, chartSpace)`` for each chart part related from ``part``."""
    out = []
    for r in part.rels.of_type(RT_CHART):
        target = part.resolve(r.target)
        if target in doc.package.parts:
            out.append((r.rid, doc.package.parts[target].element))
    return out


# -- embedded charts (Word, PowerPoint) ---------------------------------------------
XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
C_CHART_REF = f"{{{C}}}chart"
R_ID = f"{{{URI['r']}}}id"


def normalize_series(series) -> list[Series]:
    """``{"North": [1, 2]}``, ``[("North", [1, 2])]`` or ``[Series(...)]``."""
    if isinstance(series, dict):
        return [Series(str(k), list(v)) for k, v in series.items()]
    out = []
    for s in series:
        out.append(s if isinstance(s, Series) else Series(str(s[0]), list(s[1])))
    return out


def embedded_workbook(kind: str, categories: Sequence, series: list[Series]) -> bytes:
    """The chart's data as a workbook (``Sheet1``: categories in A, one column
    per series) and the series' references set to its cells."""
    from .spreadsheet import SpreadsheetDocument, column_letter

    book = SpreadsheetDocument.new("Sheet1")
    book.append_row("Sheet1", ["X" if kind == "scatter" else ""] + [s.name or "" for s in series])
    n = len(categories)
    for i, cat in enumerate(categories):
        book.append_row("Sheet1", [cat] + [s.values[i] if i < len(s.values) else None
                                           for s in series])
    for j, s in enumerate(series):
        col = column_letter(j + 2)
        s.name_ref = f"Sheet1!${col}$1" if s.name is not None else None
        s.values_ref = f"Sheet1!${col}$2:${col}${n + 1}"
    buf = io.BytesIO()
    book.save(buf)
    return buf.getvalue()


def add_embedded_chart(doc, source, kind: str, categories: Sequence, series, **options
                       ) -> tuple[str, OoxElement]:
    """Create a chart part related from ``source`` (a Word document or a slide
    part) with an embedded workbook holding its data; ``(rid, chartSpace)``."""
    series = normalize_series(series)
    if any(s.categories is not None for s in series):
        raise ValueError("embedded charts share one category list")
    blob = embedded_workbook(kind, categories, series)
    cref = f"Sheet1!$A$2:$A${len(categories) + 1}"
    space = build_chart(kind, categories, series, categories_ref=cref, **options)
    top = "/" + source.partname.strip("/").split("/")[0]
    chart_part, rid = doc.add_part("ChartPart", source, element=space, folder=f"{top}/charts")
    _, xrid = doc.add_part("EmbeddedPackagePart", chart_part, blob=blob, content_type=XLSX_CT,
                           ext="xlsx", folder=f"{top}/embeddings")
    space.append(c.ExternalData(c.AutoUpdate(val=False), r_id=xrid))
    return rid, space


def charts_in(doc, part, root) -> list[dict]:
    """``chart_info`` for each chart referenced under ``root`` (a document body
    or slide tree) of ``part``, in document order, with its "rid"."""
    out = []
    for ref in root.iter(C_CHART_REF):
        rid = ref.get(R_ID)
        rel = part.rels.items.get(rid or "")
        if rel is None or rel.reltype != RT_CHART:
            continue
        target = part.resolve(rel.target)
        if target in doc.package.parts:
            info = chart_info(doc.package.parts[target].element)
            info["rid"] = rid
            out.append(info)
    return out
