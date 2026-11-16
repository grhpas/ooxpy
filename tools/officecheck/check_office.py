"""Open every corpus file in Microsoft Office and record what happens.

Windows only.  Standalone: needs Python 3.8+ (works on Windows 7) and pywin32
(`py -m pip install pywin32`); it does NOT need oox.

    py check_office.py CORPUS_DIR                 # all apps
    py check_office.py CORPUS_DIR --apps word     # one app
    py check_office.py CORPUS_DIR --only excel-02 # matching case ids

For each file it starts a fresh Office instance, opens the file read-only and
records:
  * dialogs   -- any dialog Office shows while opening (repair / "unreadable
                 content" / recovery prompts).  A watchdog thread captures the
                 dialog text and closes it.
  * error     -- Office refused to open the file.
  * probes    -- facts read back through Office itself (table counts, computed
                 cell values, slide texts).  Catches silent repairs that drop
                 content, and checks that formulas calculate.
Status per file: ok | prompt | error | mismatch | timeout.  Normal cases pass
on "ok"; CONTROL cases (deliberately broken files) pass on "prompt"/"error"
(Office complained) or "mismatch" (Office repaired silently and content was
lost), which proves the checker notices broken files on this machine.

Close PowerPoint before running: it allows only one instance, so a running
PowerPoint would be reused (and could be closed by the run).

Results go to CORPUS_DIR/results-<host>-<apps+versions>-<date-time>.json (one
file per run, so reruns never overwrite earlier results); turn one or many
into a table with report.py.
"""

import argparse
import datetime
import hashlib
import json
import math
import os
import platform
import socket
import subprocess
import sys
import threading
import time

APPS = ("word", "excel", "powerpoint")
PROGID = {"word": "Word.Application", "excel": "Excel.Application",
          "powerpoint": "PowerPoint.Application"}
EXE = {"word": "WINWORD.EXE", "excel": "EXCEL.EXE", "powerpoint": "POWERPNT.EXE"}
OFFICE_NAMES = {"12": "Office 2007", "14": "Office 2010", "15": "Office 2013",
                "16": "Office 2016+"}  # 2016/2019/2021/365 all report 16.x
DIALOG_CLASSES = ("#32770", "NUIDialog")  # message boxes, modern Office dialogs
DIALOG_PREFIXES = ("bosa_sdm",)  # classic Office dialogs (Word, Excel)
WM_CLOSE = 0x0010


# --------------------------------------------------------------------------------------
# Portable logic (unit-tested on any OS)
# --------------------------------------------------------------------------------------
def is_dialog_class(cls):
    return cls in DIALOG_CLASSES or any(cls.startswith(p) for p in DIALOG_PREFIXES)


def office_name(version):
    return OFFICE_NAMES.get(str(version).split(".")[0], "Office " + str(version))


def normalize_text(s):
    """Word uses \\r for paragraphs, \\x0b for line breaks, \\x07 for cell ends."""
    for ch in ("\r", "\x0b", "\x07"):
        s = s.replace(ch, "\n")
    return s


def values_equal(expect, got):
    if isinstance(expect, bool) or isinstance(got, bool):
        return bool(expect) == bool(got) and type(got) in (bool, int, float)
    if isinstance(expect, (int, float)) and isinstance(got, (int, float)):
        return math.isclose(float(expect), float(got), rel_tol=1e-9, abs_tol=1e-9)
    if isinstance(expect, str) and isinstance(got, str):
        return normalize_text(expect) == normalize_text(got)
    return expect == got


class NotJudged(str):
    """Returned by an adapter for a probe its application can't answer (the
    reason is the text); the probe counts as ok and is reported as not judged."""


_PROBE_KEYS = ("sheet", "ref", "slide", "paragraph", "name", "what", "section", "part",
               "type", "index", "table", "row", "col", "char", "attr", "update", "series")


def _copy_probe_keys(probe, result):
    for key in _PROBE_KEYS:
        if key in probe:
            result[key] = probe[key]


def _attach_trace(adapter, result):
    """Steps an adapter recorded for this probe (``adapter.trace``)."""
    trace = getattr(adapter, "trace", None)
    if trace:
        result["trace"] = list(trace)


RPC_E_CALL_REJECTED = "-2147418111"  # "Call was rejected by callee": the app is busy


def evaluate_probe(probe, adapter, retries=5, pause=1.0):
    """Run one manifest probe against an app adapter -> result dict.

    Retried while Office rejects calls because it is busy (a dialog of its own,
    such as "Sign in to set up Office", is up; the watchdog closes it).  A
    rejected call never ran, so retrying is safe even for probes that change
    the document."""
    result = _evaluate_probe_once(probe, adapter)
    for _ in range(retries):
        if RPC_E_CALL_REJECTED not in str(result.get("error", "")):
            break
        time.sleep(pause)
        result = _evaluate_probe_once(probe, adapter)
    return result


# Office's own account and activation windows: about the installation, not the file
ENVIRONMENT_DIALOGS = ("Sign in to set up Office", "Activate Office",
                       "Microsoft Office Activation Wizard")


def is_environment_dialog(dialog):
    title = (dialog.get("title") or "").strip()
    return any(title.startswith(t) for t in ENVIRONMENT_DIALOGS)


def is_licence_error(text):
    """Word and Excel without an activated licence ("reduced functionality")
    refuse many calls: "... not available because the licence to use this
    application has expired" (English Office only)."""
    t = (text or "").lower()
    return "licen" in t and "expired" in t


def classify_environment(result):
    """Separate what an unactivated Office does from what the file does.

    Sign-in/activation dialogs move to ``environment_dialogs``; probes that
    failed because Office refused the call for want of a licence get
    ``refused: True`` and don't count.  A case whose remaining evidence is
    clean gets status "unactivated" (if probes were refused) or "ok".
    Idempotent; report.py also applies it to older results files."""
    dialogs = result.get("dialogs") or []
    env = [d for d in dialogs if is_environment_dialog(d)]
    if env:
        result["dialogs"] = [d for d in dialogs if not is_environment_dialog(d)]
        result["environment_dialogs"] = (result.get("environment_dialogs") or []) + env
    refused = 0
    for p in result.get("probes") or []:
        if not p.get("ok") and (is_licence_error(p.get("error")) or
                                any(is_licence_error(t) for t in p.get("trace") or [])):
            p["refused"] = True
        refused += bool(p.get("refused"))
    if result.get("status") in ("prompt", "mismatch", "unactivated", "ok"):
        judged = [p for p in result.get("probes") or [] if not p.get("refused")]
        status = decide_status(result.get("error"), result["dialogs"] if env else dialogs,
                               judged, False)
        result["status"] = "unactivated" if status == "ok" and refused else status
    return result


def _evaluate_probe_once(probe, adapter):
    kind = probe["kind"]
    result = {"kind": kind}
    try:
        adapter.trace = []
    except Exception:
        pass
    try:
        if kind == "word.text_contains":
            got = adapter.text()
            ok = normalize_text(probe["value"]) in normalize_text(got)
            result.update(value=probe["value"], ok=ok)
            if not ok:
                result["got"] = got[:300]
            return result
        if kind == "word.text_lacks":
            got = adapter.text()
            ok = normalize_text(probe["value"]) not in normalize_text(got)
            result.update(value=probe["value"], ok=ok)
            if not ok:
                result["got"] = got[:300]
            return result
        if kind == "word.text_after":
            got = adapter.text_after(probe["action"])
            ok = normalize_text(probe["value"]) in normalize_text(got)
            if probe.get("lacks"):
                ok = ok and normalize_text(probe["lacks"]) not in normalize_text(got)
            result.update(action=probe["action"], value=probe["value"], ok=ok)
            if not ok:
                result["got"] = got[:300]
            return result
        if kind == "excel.comment":
            author, got = adapter.comment(probe["sheet"], probe["ref"])
            ok = normalize_text(probe["value"]) in normalize_text(got or "")
            if "author" in probe and isinstance(author, NotJudged):
                result["note"] = "author not judged: %s" % author
            elif "author" in probe:
                ok = ok and normalize_text(author or "") == normalize_text(probe["author"])
            result.update(sheet=probe["sheet"], ref=probe["ref"], value=probe["value"],
                          author=author, ok=ok)
            if not ok:
                result["got"] = (got or "")[:300]
            return result
        if kind == "word.control":
            info = adapter.control(probe["tag"])
            got = info["text"]
            ok = normalize_text(probe["value"]) in normalize_text(got)
            if "type" in probe:
                ok = ok and info["type"] == probe["type"]
            if "placeholder" in probe:
                ok = ok and bool(info["placeholder"]) == bool(probe["placeholder"])
            result.update(tag=probe["tag"], value=probe["value"], ok=ok,
                          type=info["type"], placeholder=info["placeholder"])
            if not ok:
                result["got"] = got[:300]
            return result
        if kind in ("word.header_footer", "word.note", "word.comment", "word.table_cell",
                    "word.revision"):
            # "contains" probes on one story: header/footer, a note, a comment
            author = None
            if kind == "word.header_footer":
                got = adapter.header_footer_text(probe["part"], probe["section"], probe["type"])
                where = {k: probe[k] for k in ("part", "section", "type")}
            elif kind == "word.table_cell":
                got = adapter.table_cell(probe["table"], probe["row"], probe["col"])
                where = {k: probe[k] for k in ("table", "row", "col")}
            elif kind == "word.revision":
                rtype, author, got = adapter.revision(probe["index"])
                where = {"index": probe["index"], "type": rtype}
                if rtype != probe["type"]:
                    got = "[type %s] %s" % (rtype, got)
            elif kind == "word.note":
                got = adapter.note_text(probe["type"], probe["index"])
                where = {"type": probe["type"], "index": probe["index"]}
            else:
                author, got = adapter.comment(probe["index"])
                where = {"index": probe["index"]}
            ok = normalize_text(probe["value"]) in normalize_text(got)
            if kind == "word.revision":
                ok = ok and where["type"] == probe["type"]
            if "author" in probe:
                where["author"] = author
                ok = ok and normalize_text(probe["author"]) == normalize_text(author or "")
            result.update(where, value=probe["value"], ok=ok)
            if not ok:
                result["got"] = got[:300]
            return result
        if kind == "word.tables":
            got = adapter.tables()
        elif kind == "word.inline_shapes":
            got = adapter.inline_shapes()
        elif kind == "word.pages":
            got = adapter.pages()
        elif kind == "word.outline_level":
            got = adapter.outline_level(probe["paragraph"])
        elif kind == "word.list_string":
            got = adapter.list_string(probe["paragraph"])
        elif kind == "word.font":
            got = adapter.font(probe["paragraph"], probe["char"], probe["attr"])
        elif kind == "word.paragraph":
            got = adapter.paragraph_format(probe["paragraph"], probe["attr"])
        elif kind == "word.field_results":
            got = adapter.field_results(probe["codes"], probe.get("update", False))
        elif kind == "doc.custom_property":
            got = adapter.custom_property(probe["name"])
        elif kind == "word.table":
            got = adapter.table(probe["table"], probe["what"])
        elif kind == "word.table_header":
            got = adapter.table_header(probe["table"])
        elif kind == "word.shapes":
            got = adapter.shapes_count()
        elif kind == "word.shape_wraps":
            got = adapter.shape_wraps()
            if "none" in got:  # Word 2007 can't tell behind from in front
                probe = dict(probe, expect=sorted(
                    "none" if x in ("front", "behind") else x for x in probe["expect"]))
        elif kind == "word.text_boxes":
            got = adapter.text_boxes()
        elif kind == "word.protection":
            got = adapter.protection()
        elif kind == "word.unprotect":
            got = adapter.unprotect(probe.get("wrong"), probe.get("password"))
        elif kind == "excel.workbook_protection":
            got = adapter.workbook_protection(probe["attr"])
        elif kind == "excel.unprotect_workbook":
            got = adapter.unprotect_workbook(probe.get("wrong"), probe.get("password"))
        elif kind == "word.count":
            got = adapter.count(probe["what"])
        elif kind == "word.orientation":
            got = adapter.orientation(probe["section"])
        elif kind == "word.fields":
            got = adapter.fields(probe["part"], probe["section"], probe["type"])
        elif kind == "word.hyperlink":
            got = adapter.hyperlink(probe["index"])
        elif kind == "doc.property":
            got = adapter.property(probe["name"])
        elif kind == "excel.sheets":
            got = adapter.sheets()
        elif kind == "excel.cell":
            got = adapter.cell(probe["sheet"], probe["ref"])
        elif kind == "excel.text":
            if probe.get("names") and hasattr(adapter, "english_names") and \
                    not adapter.english_names():
                result.update(sheet=probe["sheet"], ref=probe["ref"], ok=True,
                              note="not judged: month/day names follow the regional settings")
                return result
            got = adapter.text(probe["sheet"], probe["ref"])
        elif kind == "excel.font":
            got = adapter.font(probe["sheet"], probe["ref"], probe["attr"])
        elif kind == "excel.fill":
            got = adapter.fill(probe["sheet"], probe["ref"])
        elif kind == "excel.align":
            got = adapter.align(probe["sheet"], probe["ref"])
        elif kind == "excel.number_format":
            got = adapter.number_format(probe["sheet"], probe["ref"])
        elif kind == "excel.column_width":
            got = adapter.column_width(probe["sheet"], probe["col"])
        elif kind == "excel.row_height":
            got = adapter.row_height(probe["sheet"], probe["row"])
        elif kind == "excel.merge_area":
            got = adapter.merge_area(probe["sheet"], probe["ref"])
        elif kind == "excel.frozen":
            got = adapter.frozen(probe["sheet"])
        elif kind == "excel.autofilter":
            got = adapter.autofilter(probe["sheet"])
        elif kind == "excel.defined_name":
            got = adapter.defined_name(probe["name"])
        elif kind == "excel.validation":
            got = adapter.validation(probe["sheet"], probe["ref"])
        elif kind == "excel.charts":
            got = adapter.chart_count(probe["sheet"])
        elif kind == "excel.chart":
            got = adapter.chart(probe["sheet"], probe["index"], probe["attr"], probe.get("series", 1))
        elif kind == "excel.tabs":
            got = adapter.tabs()
        elif kind == "excel.chart_sheet":
            got = adapter.chart_sheet(probe["name"], probe["attr"], probe.get("series", 1))
        elif kind == "excel.chart_image":
            got = adapter.chart_image(probe["sheet"], probe["index"])
        elif kind == "word.charts":
            got = adapter.chart_count()
        elif kind == "word.chart":
            got = adapter.chart(probe["index"], probe["attr"], probe.get("series", 1))
        elif kind == "ppt.charts":
            got = adapter.chart_count(probe["slide"])
        elif kind == "ppt.chart":
            got = adapter.chart(probe["slide"], probe["index"], probe["attr"], probe.get("series", 1))
        elif kind == "ppt.layouts":
            got = sorted(adapter.layouts())
        elif kind == "ppt.slide_layout":
            got = adapter.slide_layout(probe["slide"])
        elif kind == "ppt.table":
            got = adapter.table(probe["slide"], probe["attr"], probe.get("row"), probe.get("col"))
        elif kind == "ppt.notes":
            got = adapter.notes(probe["slide"])
        elif kind == "ppt.sections":
            got = adapter.sections()
        elif kind == "ppt.links":
            got = sorted(adapter.links(probe["slide"]))
        elif kind == "ppt.action":
            got = adapter.action(probe["slide"], probe["shape"])
        elif kind == "ppt.shape_types":
            got = adapter.shape_types(probe["slide"])
        elif kind == "ppt.picture_placeholder":
            got = adapter.picture_placeholder(probe["slide"])
        elif kind == "ppt.slide_image":
            got = adapter.slide_image(probe["slide"])
        elif kind == "excel.screenshot":
            got = adapter.screenshot(probe["sheet"], probe["ref"])
        elif kind == "excel.cf_count":
            got = adapter.cf_count(probe["sheet"])
        elif kind == "excel.cf_types":
            got = adapter.cf_types(probe["sheet"], probe["ref"])
            if "formatCondition" in got:
                want = sorted("formatCondition" if k in _PLAIN_CF else k for k in probe["expect"])
                result.update(expect=probe["expect"], got=got, ok=got == want,
                              note="plain rule kinds (comparison, formula, text, blanks) "
                                   "not judged: this Excel misreads FormatCondition.Type")
                _copy_probe_keys(probe, result)
                _attach_trace(adapter, result)
                return result
        elif kind == "excel.tables":
            got = adapter.table_names(probe["sheet"])
        elif kind == "excel.table":
            got = adapter.table(probe["sheet"], probe["name"], probe["attr"])
        elif kind == "excel.comments":
            got = adapter.comment_count(probe["sheet"])
        elif kind == "excel.hyperlink":
            got = adapter.hyperlink(probe["sheet"], probe["ref"])
        elif kind == "excel.page":
            got = adapter.page(probe["sheet"], probe["attr"])
        elif kind == "excel.sheet":
            got = adapter.sheet_option(probe["sheet"], probe["attr"])
        elif kind == "excel.shapes":
            got = adapter.shapes(probe["sheet"])
        elif kind == "ppt.slides":
            got = adapter.slides()
        elif kind == "ppt.shapes":
            got = adapter.shapes(probe["slide"])
        elif kind == "ppt.text_contains":
            got = adapter.slide_text(probe["slide"])
            ok = normalize_text(probe["value"]) in normalize_text(got)
            result.update(value=probe["value"], slide=probe["slide"], ok=ok)
            if not ok:
                result["got"] = got[:300]
            return result
        else:
            result.update(ok=False, error="unknown probe kind")
            return result
    except Exception as e:
        result.update(ok=False, error=repr(e))
        if "expect" in probe:
            result["expect"] = probe["expect"]
        _copy_probe_keys(probe, result)
        _attach_trace(adapter, result)
        return result
    _attach_trace(adapter, result)
    for key in ("sheet", "ref", "slide", "paragraph", "name", "what", "section", "part",
                "type", "index", "table", "row", "col", "char", "attr", "update", "name"):
        if key in probe:
            result[key] = probe[key]
    if isinstance(got, NotJudged):
        result.update(expect=probe["expect"], ok=True, note="not judged: %s" % got)
        return result
    result.update(expect=probe["expect"], got=got, ok=values_equal(probe["expect"], got))
    return result


def split_expected(dialogs, windows, grace=2.0):
    """Dialogs a probe provoked on purpose (shown inside one of ``windows``,
    ``(start, end)`` times of calls expected to fail, plus ``grace`` seconds)
    -> (unexpected, expected).  Word shows "The password is incorrect" as well
    as raising the COM error; that is the probe working, not the file."""
    unexpected, expected = [], []
    for d in dialogs:
        t = d.get("time")
        hit = t is not None and any(a <= t <= b + grace for a, b in windows)
        (expected if hit else unexpected).append(d)
    return unexpected, expected


def decide_status(open_error, dialogs, probe_results, timed_out):
    if timed_out:
        return "timeout"
    if dialogs:
        return "prompt"
    if open_error:
        return "error"
    if any(not p.get("ok") for p in probe_results):
        return "mismatch"
    return "ok"


def passed(case, status):
    """A case passes if its status is the expected one (controls expect a prompt/error).
    Observe cases have no expectation: they always "pass" and are just recorded."""
    if case.get("observe"):
        return True
    return status in case.get("expect_status", ["ok"])


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


# --------------------------------------------------------------------------------------
# COM adapters (thin; the probes above call these)
# --------------------------------------------------------------------------------------
def _international(app, index):
    """Application.International(index).  Through pywin32's dynamic dispatch the
    property can come back as the tuple of all settings (VBA's 1-based list),
    so both shapes are handled."""
    value = app.International
    if isinstance(value, (tuple, list)):
        return value[index - 1]
    if callable(value):
        return value(index)
    return app.International(index)


def _english(obj, name):
    """A property read with LCID 1033 (as VBA does).  pywin32 passes the user's
    locale, and Excel then answers in it: Greek Excel turns NumberFormat
    "yyyy-mm-dd" into "εεεε-μμ-ηη".  Falls back to a plain read."""
    try:
        import pythoncom

        ole = obj._oleobj_
        dispid = ole.GetIDsOfNames(name)
        return ole.Invoke(dispid, 0x409, pythoncom.DISPATCH_PROPERTYGET, True)
    except Exception:
        return getattr(obj, name)


_CF_CLASS = {"ColorScale": "colorScale", "Databar": "dataBar",
             "IconSetCondition": "iconSet", "Top10": "top10",
             "AboveAverage": "aboveAverage"}  # COM interface -> cfRule type


def _interface_name(obj):
    """Name of the COM interface behind ``obj`` ("FormatCondition", "ColorScale", ...)."""
    try:
        return obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    except Exception:
        return None


_PLAIN_CF = frozenset({"cellIs", "expression", "containsText", "notContainsText",
                       "beginsWith", "endsWith", "containsBlanks", "notContainsBlanks",
                       "containsErrors", "notContainsErrors", "timePeriod"})
_TEXT_OPERATOR = {0: "containsText", 1: "notContainsText", 2: "beginsWith", 3: "endsWith"}


def _cf_kind(fc, trust_type=True, trace=None, label=""):
    """cfRule type of one FormatConditions item.  Special rules come from
    the COM interface.  For plain FormatCondition rules ``Type`` is read
    only if ``trust_type``; Excel 2007 raises (sheet's rules) or answers with
    the cell's first rule (cell's rules), so there the rule counts as
    "formatCondition" and a kind derived from other properties goes into
    ``trace`` for information."""
    name = _interface_name(fc)
    if name in _CF_CLASS:
        return _CF_CLASS[name]
    if name == "UniqueValues":
        return "duplicateValues" if int(_english(fc, "DupeUnique")) == 1 else "uniqueValues"
    if trust_type:
        typ = int(_english(fc, "Type"))
        return _CF_TYPE.get(typ, str(typ))
    if trace is not None:
        trace.append("%s plain rule: %s" % (label, _plain_hint(fc)))
    return "formatCondition"


def _plain_hint(fc):
    """Read Operator, Formula1, TextOperator, Text, and guess the kind:
    text rule if TextOperator and Text read, comparison if Operator reads,
    formula if only Formula1 reads.  Diagnostics only."""
    got, readable = {}, set()
    for prop in ("Operator", "Formula1", "TextOperator", "Text"):
        try:
            got[prop] = _english(fc, prop)
            readable.add(prop)
        except Exception as e:
            got[prop] = _error_text(e)
    if {"TextOperator", "Text"} <= readable and got["Text"]:
        guess = _TEXT_OPERATOR.get(got["TextOperator"], "text?")
    elif "Operator" in readable:
        guess = "cellIs"
    elif "Formula1" in readable:
        guess = "expression"
    else:
        guess = "?"
    return "%s -> derived %s" % (", ".join("%s=%r" % kv for kv in got.items()), guess)


SCREENSHOT_SCALE = 2


def _scale_pasted(chart, w, h, scale):
    """Enlarge the chart and the picture pasted into it; no border.  Best
    effort: returns the scale reached (1 if Excel refused)."""
    try:
        chart.Chart.ChartArea.Format.Line.Visible = False
    except Exception:
        pass
    try:
        pic = chart.Chart.Shapes(1)
        chart.Width, chart.Height = w * scale + 4, h * scale + 4
        pic.Left, pic.Top, pic.Width, pic.Height = 0, 0, w * scale, h * scale
        return scale
    except Exception:
        try:
            chart.Width, chart.Height = w + 4, h + 4
        except Exception:
            pass
        return 1


# XlChartType -> oox.charts vocabulary (kept equal to oox.charts.EXCEL_CHART_TYPES
# by a test: this script stays standalone)
CHART_TYPES = {
    51: "column", 52: "column_stacked", 53: "column_percent",
    57: "bar", 58: "bar_stacked", 59: "bar_percent",
    4: "line", 65: "line_markers", 63: "line_stacked", 66: "line_markers_stacked",
    64: "line_percent", 67: "line_markers_percent",
    1: "area", 76: "area_stacked", 77: "area_percent",
    5: "pie", -4120: "doughnut",
    -4169: "scatter", 74: "scatter_lines", 75: "scatter_lines_no_markers",
    72: "scatter_smooth", 73: "scatter_smooth_no_markers",
}
# preset geometry -> MsoAutoShapeType (kept equal to oox.slides.SHAPES by a test)
AUTO_SHAPES = {"rect": 1, "roundRect": 5, "ellipse": 9, "triangle": 7, "diamond": 4,
               "hexagon": 10, "octagon": 6, "rightArrow": 33, "leftArrow": 34, "upArrow": 35,
               "downArrow": 36, "homePlate": 51, "chevron": 52}
# PpActionType -> oox.slidelinks' jump names ("hyperlink": a URL or a slide)
PP_ACTIONS = {0: "none", 1: "next", 2: "previous", 3: "first", 4: "last", 5: "last_viewed",
              6: "end", 7: "hyperlink"}
_LEGEND_POS = {-4107: "bottom", -4131: "left", -4152: "right", -4160: "top", 2: "top_right"}


def _plain_number(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def com_chart(ch, attr, series=1):
    """One attribute of an Office ``Chart`` object (Excel, Word and PowerPoint
    share the chart object model), in oox.charts.chart_info's terms."""
    if attr == "type":
        t = int(ch.ChartType)
        return CHART_TYPES.get(t, str(t))
    if attr == "title":
        return str(ch.ChartTitle.Text) if ch.HasTitle else None
    if attr == "legend":
        return _LEGEND_POS.get(int(ch.Legend.Position), str(ch.Legend.Position)) \
            if ch.HasLegend else None
    if attr in ("x_title", "y_title"):
        ax = ch.Axes(1 if attr == "x_title" else 2)  # xlCategory, xlValue
        return str(ax.AxisTitle.Text) if ax.HasTitle else None
    if attr == "y2_title":
        # Not HasAxis(2, 2): through dynamic COM a parameterised property is
        # fetched without its arguments, and Excel 2007 answers E_FAIL. Nor
        # Axes().Item(i): Item takes (Type, AxisGroup), not a position.
        # Axes(xlValue, xlSecondary) is a method call; if it fails, the
        # collection is enumerated (_NewEnum) to tell "no such axis" apart.
        try:
            ax = ch.Axes(2, 2)
        except Exception as first:
            try:
                found = [a for a in ch.Axes()
                         if int(a.Type) == 2 and int(a.AxisGroup) == 2]
            except Exception:
                raise first from None
            if not found:
                return None
            ax = found[0]
        return str(ax.AxisTitle.Text) if ax.HasTitle else None
    coll = ch.SeriesCollection()
    if attr == "series":
        return [str(coll.Item(i).Name) for i in range(1, coll.Count + 1)]
    ser = coll.Item(series)
    if attr == "series_type":
        t = int(ser.ChartType)
        return CHART_TYPES.get(t, str(t))
    if attr == "axis":
        return {1: "primary", 2: "secondary"}.get(int(ser.AxisGroup), str(ser.AxisGroup))
    if attr == "values":
        return [_plain_number(v) for v in ser.Values]
    if attr == "categories":
        return [_plain_number(v) for v in ser.XValues]
    if attr == "labels":
        if not ser.HasDataLabels:
            return []
        dl = ser.DataLabels()
        return sorted(name for name, prop in (("value", "ShowValue"), ("percent", "ShowPercentage"),
                                              ("category", "ShowCategoryName"),
                                              ("series", "ShowSeriesName"))
                      if getattr(dl, prop))
    raise ValueError("unknown chart attribute %r" % attr)


def _contains_picture(shape):
    """ContainedType (PowerPoint 2010+) = msoPicture; on 2007, a picture's crop
    values (PictureFormat answers only for a placeholder holding a picture)."""
    try:
        return int(shape.PlaceholderFormat.ContainedType) == 13  # msoPicture
    except Exception:
        fmt = shape.PictureFormat
        return any(float(getattr(fmt, side)) > 0
                   for side in ("CropLeft", "CropRight", "CropTop", "CropBottom"))


def _chart_om(app_version, fn, *args):
    """Object-model features Office 2007 may lack (Word/PowerPoint charts
    before SP2, PlaceholderFormat.ContainedType): on 2007 a failure is
    reported as not judged (with the error); from 2010 on it is an error."""
    try:
        return fn(*args)
    except Exception as e:
        if app_version < 14:
            return NotJudged("not in Office 2007's object model (charts need SP2): %s"
                             % _error_text(e))
        raise


def _error_text(e):
    """Short text for an exception; COM errors with their Excel scode in hex."""
    try:
        scode = e.args[2][5]
        if scode:
            return "%s 0x%08X" % (type(e).__name__, scode & 0xFFFFFFFF)
    except Exception:
        pass
    return repr(e)[:120]


def _step(trace, label, fn, *args):
    """Run one COM call; on failure note which step failed, then re-raise."""
    try:
        return fn(*args)
    except Exception as e:
        trace.append("%s: %s" % (label, _error_text(e)))
        raise


def _cell_rc(ref):
    """"B12" / "$B$12" -> (12, 2)."""
    import re

    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", ref.strip())
    if not m:
        raise ValueError("not a cell reference: %r" % ref)
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


def _covers(address, row, col):
    """Does an Excel address ("$A$2:$E$11,$G$1"; areas split on , or ;)
    include the cell?"""
    import re

    for area in re.split(r"[,;]", address):
        area = area.strip().split("!")[-1]
        if not area:
            continue
        first, _, last = area.partition(":")
        r1, c1 = _cell_rc(first)
        r2, c2 = _cell_rc(last) if last else (r1, c1)
        if min(r1, r2) <= row <= max(r1, r2) and min(c1, c2) <= col <= max(c1, c2):
            return True
    return False


def _is_date_code(code):
    """Date/time number format (letters d m y h s outside quotes and [...])."""
    import re

    bare = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", code or "")
    return bool(re.search(r"[dmyhs]", bare, re.I)) and code not in ("General", "@")


def _custom_property(doc, name):
    """CustomDocumentProperties(name).Value; None if missing."""
    try:
        return doc.CustomDocumentProperties(name).Value
    except Exception:
        return None


def _hex_color(value):
    """Word's Font.Color (0x00BBGGRR, or wdColorAutomatic) -> "RRGGBB" / None."""
    if value is None or value in (-16777216, 0xFF000000) or value < 0:
        return None
    return "%02X%02X%02X" % (value & 0xFF, (value >> 8) & 0xFF, (value >> 16) & 0xFF)


_CF_TYPE = {1: "cellIs", 2: "expression", 3: "colorScale", 4: "dataBar", 5: "top10",
            6: "iconSet", 8: "duplicateValues", 9: "containsText", 10: "containsBlanks",
            12: "aboveAverage", 13: "notContainsBlanks"}  # XlFormatConditionType
_XL_ALIGN = {1: "general", -4131: "left", -4108: "center", -4152: "right", 5: "fill",
             -4130: "justify", 7: "centerContinuous", -4117: "distributed"}  # XlHAlign
_ALIGNMENT = {0: "left", 1: "center", 2: "right", 3: "justify"}  # WdParagraphAlignment
_CONTROL = {0: "rich", 1: "text", 2: "picture", 3: "combo", 4: "dropdown", 5: "gallery",
            6: "date", 7: "group", 8: "checkbox"}  # WdContentControlType


def _builtin_property(doc, name):
    """BuiltInDocumentProperties by English name ("Title", "Author", ...);
    the names are not localized.  Unset properties read as None."""
    try:
        value = doc.BuiltInDocumentProperties(name).Value
    except Exception:
        return None
    return value if value != "" else None


_WRAP = {0: "square", 1: "tight", 2: "through", 3: "front", 4: "topAndBottom", 5: "behind",
         7: "inline"}  # WdWrapType
_REVISION = {1: "insert", 2: "delete", 3: "format", 10: "paragraph-format"}  # WdRevisionType
_HF_INDEX = {"default": 1, "first": 2, "even": 3}  # wdHeaderFooterPrimary/FirstPage/EvenPages


_PROTECTION = {-1: "none", 0: "trackedChanges", 1: "comments", 2: "forms", 3: "readOnly"}


def _unprotect(adapter, unprotect, wrong, password, after):
    """Try ``wrong`` (must fail), then ``password`` (None: no argument).  The
    wrong attempt's time window goes to ``adapter.expected_windows``: a dialog
    shown then is the probe's doing (see ``split_expected``)."""
    trace = adapter.trace
    refused = None
    if wrong is not None:
        start = time.time()
        try:
            unprotect(wrong)
            return [False, None, after()]
        except Exception as e:
            trace.append("Unprotect(wrong): %r" % (e,))
            refused = True
        finally:
            windows = getattr(adapter, "expected_windows", None)
            if windows is None:
                windows = adapter.expected_windows = []
            windows.append((start, time.time()))
    try:
        if password is None:
            unprotect()
        else:
            unprotect(password)
        accepted = True
    except Exception as e:
        trace.append("Unprotect(password): %r" % (e,))
        accepted = False
    return [refused, accepted, after()]


class WordAdapter:
    def __init__(self, doc):
        self.doc = doc
        self.trace = []

    def _version(self):
        try:
            return int(float(str(self.doc.Application.Version)))
        except Exception:
            return 99

    def chart_count(self):
        """Inline shapes (the chart case has no pictures); their types are traced."""
        shapes = self.doc.InlineShapes
        n = int(shapes.Count)
        self.trace = ["InlineShapes(%d).Type = %s" % (i, shapes(i).Type) for i in range(1, n + 1)]
        return n

    def chart(self, index, attr, series=1):
        return _chart_om(self._version(),
                         lambda: com_chart(self.doc.InlineShapes(index).Chart, attr, series))

    def text(self):
        return self.doc.Content.Text

    def tables(self):
        return self.doc.Tables.Count

    def inline_shapes(self):
        return self.doc.InlineShapes.Count

    def pages(self):
        return self.doc.ComputeStatistics(2)  # wdStatisticPages

    def outline_level(self, paragraph):
        """1-9 for headings, 10 for body text -- language-independent."""
        return self.doc.Paragraphs(paragraph).OutlineLevel

    def property(self, name):
        return _builtin_property(self.doc, name)

    def list_string(self, paragraph):
        """The list label Word displays ("2.", "1.3."), "" if not a list item."""
        return self.doc.Paragraphs(paragraph).Range.ListFormat.ListString

    def count(self, what):
        d = self.doc
        return {"sections": lambda: d.Sections.Count, "hyperlinks": lambda: d.Hyperlinks.Count,
                "bookmarks": lambda: d.Bookmarks.Count,  # hidden ones (_Toc...) not counted
                "footnotes": lambda: d.Footnotes.Count, "endnotes": lambda: d.Endnotes.Count,
                "comments": lambda: d.Comments.Count, "revisions": lambda: d.Revisions.Count,
                "controls": lambda: d.ContentControls.Count,
                "toc": lambda: d.TablesOfContents.Count}[what]()

    def orientation(self, section):
        return ("portrait", "landscape")[self.doc.Sections(section).PageSetup.Orientation]

    def _hf(self, part, section, kind):
        sec = self.doc.Sections(section)
        coll = sec.Headers if part == "header" else sec.Footers
        return coll(_HF_INDEX[kind])  # linked ones return the previous section's

    def header_footer_text(self, part, section, kind):
        return self._hf(part, section, kind).Range.Text

    def fields(self, part, section, kind):
        fs = self._hf(part, section, kind).Range.Fields
        return [fs(i).Code.Text.split()[0].upper() for i in range(1, fs.Count + 1)]

    def hyperlink(self, index):
        h = self.doc.Hyperlinks(index)
        return h.Address if h.Address else "#" + h.SubAddress

    def note_text(self, kind, index):
        coll = self.doc.Footnotes if kind == "footnote" else self.doc.Endnotes
        return coll(index).Range.Text

    def comment(self, index):
        c = self.doc.Comments(index)
        return c.Author, c.Range.Text

    def custom_property(self, name):
        return _custom_property(self.doc, name)

    def font(self, paragraph, char, attr):
        f = self.doc.Paragraphs(paragraph).Range.Characters(char).Font
        getters = {
            "bold": lambda: bool(f.Bold), "italic": lambda: bool(f.Italic),
            "underline": lambda: f.Underline != 0, "strike": lambda: bool(f.StrikeThrough),
            "subscript": lambda: bool(f.Subscript), "superscript": lambda: bool(f.Superscript),
            "small_caps": lambda: bool(f.SmallCaps), "size": lambda: round(f.Size, 1),
            "font": lambda: f.Name, "color": lambda: _hex_color(f.Color),
        }
        return getters[attr]()

    def paragraph_format(self, paragraph, attr):
        p = self.doc.Paragraphs(paragraph)
        getters = {
            "style": lambda: p.Style.NameLocal, "align": lambda: _ALIGNMENT.get(p.Alignment),
            "space_before": lambda: round(p.SpaceBefore, 1),
            "space_after": lambda: round(p.SpaceAfter, 1),
            "indent_left": lambda: round(p.LeftIndent, 1),
            "first_line": lambda: round(p.FirstLineIndent, 1),
        }
        return getters[attr]()

    def field_results(self, codes, update=False):
        """Results of the main text's fields whose name is in ``codes``;
        ``update`` makes Word recompute every field first (changes the
        document, so it runs as the last probe)."""
        fields = self.doc.Fields
        if update:
            fields.Update()
        out = []
        for i in range(1, fields.Count + 1):
            f = fields(i)
            words = f.Code.Text.split()
            if words and words[0].upper() in codes:
                out.append(f.Result.Text)
        return out

    def control(self, tag):
        cc = self.doc.SelectContentControlsByTag(tag)(1)
        return {"text": cc.Range.Text, "type": _CONTROL.get(cc.Type, str(cc.Type)),
                "placeholder": bool(cc.ShowingPlaceholderText)}

    def table(self, table, what):
        t = self.doc.Tables(table)
        return {"rows": lambda: t.Rows.Count, "cells": lambda: t.Range.Cells.Count}[what]()

    def table_header(self, table):
        return self.doc.Tables(table).Rows(1).HeadingFormat  # -1 (True) or 0

    def table_cell(self, table, row, col):
        """Word's Cell(row, col): col counts the cells *in that row*."""
        return self.doc.Tables(table).Cell(row, col).Range.Text

    def shapes_count(self):
        return self.doc.Shapes.Count  # floating; inline ones are InlineShapes

    def shape_wraps(self):
        """Wrap names, sorted.  Word 2007 (12) reports "behind text" and "in front
        of text" both as 3 (wdWrapNone): that becomes "none"."""
        shapes = self.doc.Shapes
        old = float(self.doc.Application.Version) < 14
        names = []
        for i in range(1, shapes.Count + 1):
            t = shapes(i).WrapFormat.Type
            names.append("none" if (t == 3 and old) else _WRAP.get(t, str(t)))
        return sorted(names)

    def text_boxes(self):
        """Texts of the floating shapes that hold text, in order (types traced)."""
        shapes = self.doc.Shapes
        out = []
        for i in range(1, shapes.Count + 1):
            s = shapes(i)
            self.trace.append("Shapes(%d).Type = %s" % (i, s.Type))
            try:
                frame = s.TextFrame
                if not frame.HasText:
                    continue
                out.append(normalize_text(str(frame.TextRange.Text)).rstrip("\n"))
            except Exception as e:
                self.trace.append("Shapes(%d).TextFrame: %r" % (i, e))
        return out

    def protection(self):
        t = int(self.doc.ProtectionType)
        return _PROTECTION.get(t, str(t))

    def unprotect(self, wrong, password):
        """[wrong password refused, right one accepted, protection after].
        This changes the document, so it runs as the last probe."""
        return _unprotect(self, self.doc.Unprotect, wrong, password, self.protection)

    def revision(self, index):
        r = self.doc.Revisions(index)
        return _REVISION.get(r.Type, str(r.Type)), r.Author, r.Range.Text

    def text_after(self, action):
        """Accept or reject all changes in the (read-only) open document, then
        its text.  This changes the document, so it runs as the last probe."""
        if action == "accept":
            self.doc.Revisions.AcceptAll()
        else:
            self.doc.Revisions.RejectAll()
        return self.doc.Content.Text


class ExcelAdapter:
    def __init__(self, wb):
        self.wb = wb
        self.trace = []  # COM steps of the current probe (shown in the report)
        self.screens: tuple = (None, None)  # (folder, file prefix) for screenshot()

    def sheets(self):
        return [self.wb.Worksheets(i).Name for i in range(1, self.wb.Worksheets.Count + 1)]

    def cell(self, sheet, ref):
        return self.wb.Worksheets(sheet).Range(ref).Value

    def shapes(self, sheet):
        return self.wb.Worksheets(sheet).Shapes.Count

    def property(self, name):
        return _builtin_property(self.wb, name)

    def custom_property(self, name):
        return _custom_property(self.wb, name)

    def _separators(self):
        """Excel's decimal, thousands and list separators (regional settings)."""
        app = self.wb.Application
        if app.UseSystemSeparators:
            dec, thou = _international(app, 3), _international(app, 4)  # xlDecimal/ThousandsSeparator
        else:
            dec, thou = app.DecimalSeparator, app.ThousandsSeparator
        return dec, thou, _international(app, 5)  # xlListSeparator

    def _to_english(self, text):
        """Swap regional decimal/thousands separators for "." and ","."""
        dec, thou, _ = self._separators()
        if (dec, thou) == (".", ","):
            return text
        return text.replace(thou, "\0").replace(dec, ".").replace("\0", ",")

    def english_names(self):
        """True if month/day names come out in English: the regional setting
        (xlCountrySetting, 2), not Excel's own language (xlCountryCode, 1)."""
        return _international(self.wb.Application, 2) in (1, 44, 61, 64, 353)  # country codes

    def text(self, sheet, ref):
        """What the cell shows (number format applied), with the regional
        separators of numeric cells swapped to "." and ","."""
        rng = self.wb.Worksheets(sheet).Range(ref)
        text = rng.Text
        value = rng.Value2
        # plain numbers only: in dates and times "," and "." are literal text
        if isinstance(value, (int, float)) and not isinstance(value, bool) and \
                not _is_date_code(_english(rng, "NumberFormat")):
            text = self._to_english(text)
        return text

    def font(self, sheet, ref, attr):
        f = self.wb.Worksheets(sheet).Range(ref).Font
        return {"bold": lambda: bool(f.Bold), "italic": lambda: bool(f.Italic),
                "underline": lambda: f.Underline not in (-4142, None, 0),  # xlUnderlineStyleNone
                "size": lambda: round(f.Size, 1), "font": lambda: f.Name,
                "color": lambda: _hex_color(int(f.Color))}[attr]()

    def fill(self, sheet, ref):
        interior = self.wb.Worksheets(sheet).Range(ref).Interior
        if interior.Pattern == -4142:  # xlNone
            return None
        return _hex_color(int(interior.Color))

    def align(self, sheet, ref):
        return _XL_ALIGN.get(self.wb.Worksheets(sheet).Range(ref).HorizontalAlignment)

    def number_format(self, sheet, ref):
        """Asked in English (LCID 1033): a plain read through pywin32 gets the
        regional variant ("0,0%", "εεεε-μμ-ηη")."""
        rng = self.wb.Worksheets(sheet).Range(ref)
        code = _english(rng, "NumberFormat")
        return code if hasattr(rng, "_oleobj_") else self._to_english(code)

    def column_width(self, sheet, col):
        return round(self.wb.Worksheets(sheet).Columns(col).ColumnWidth, 2)

    def row_height(self, sheet, row):
        return round(self.wb.Worksheets(sheet).Rows(row).RowHeight, 2)

    def merge_area(self, sheet, ref):
        return self.wb.Worksheets(sheet).Range(ref).MergeArea.Address.replace("$", "")

    def frozen(self, sheet):
        """Top-left cell of the scrolling area if panes are frozen (the sheet
        is activated first: freezing is a window property)."""
        self.wb.Worksheets(sheet).Activate()
        win = self.wb.Windows(1)
        if not win.FreezePanes:
            return None
        col, row = int(win.SplitColumn) + 1, int(win.SplitRow) + 1
        letters = ""
        while col:
            col, rem = divmod(col - 1, 26)
            letters = chr(65 + rem) + letters
        return "%s%d" % (letters, row)

    def autofilter(self, sheet):
        ws = self.wb.Worksheets(sheet)
        return ws.AutoFilter.Range.Address.replace("$", "") if ws.AutoFilterMode else None

    def defined_name(self, name):
        return self.wb.Names(name).RefersTo.lstrip("=")

    def cf_count(self, sheet):
        return self.wb.Worksheets(sheet).Cells.FormatConditions.Count

    def chart_count(self, sheet):
        return int(self.wb.Worksheets(sheet).ChartObjects().Count)

    def chart(self, sheet, index, attr, series=1):
        return com_chart(self.wb.Worksheets(sheet).ChartObjects(index).Chart, attr, series)

    def workbook_protection(self, attr):
        return bool({"structure": lambda: self.wb.ProtectStructure,
                     "windows": lambda: self.wb.ProtectWindows}[attr]())

    def unprotect_workbook(self, wrong, password):
        """[wrong password refused, right one accepted, structure locked after]."""
        return _unprotect(self, self.wb.Unprotect, wrong, password,
                          lambda: bool(self.wb.ProtectStructure))

    def tabs(self):
        """Every tab in order (worksheets and chart sheets)."""
        return [str(self.wb.Sheets(i).Name) for i in range(1, self.wb.Sheets.Count + 1)]

    def chart_sheet(self, name, attr, series=1):
        return com_chart(self.wb.Charts(name), attr, series)

    def chart_image(self, sheet, index):
        """Chart.Export: the chart as Excel draws it (no clipboard)."""
        folder, prefix = self.screens
        if not folder:
            return NotJudged("no screenshot folder")
        os.makedirs(folder, exist_ok=True)
        name = "%s-%s-chart%d.png" % (prefix, sheet, index)
        self.wb.Worksheets(sheet).ChartObjects(index).Chart.Export(os.path.join(folder, name), "PNG")
        return NotJudged("chart image saved: screens/%s" % name)

    def screenshot(self, sheet, ref):
        """Save ``ref`` as Excel displays it (conditional formats included) to
        screens/<host>-<case>-<sheet>.png: Range.CopyPicture(xlScreen), paste
        into a temporary chart, Chart.Export.  Pure COM, no extra packages;
        the workbook is read-only and closed unsaved.  Uses the clipboard."""
        folder, prefix = self.screens
        if not folder:
            return NotJudged("no screenshot folder")
        os.makedirs(folder, exist_ok=True)
        name = "%s-%s.png" % (prefix, sheet)
        path = os.path.join(folder, name)
        ws = self.wb.Worksheets(sheet)
        rng = ws.Range(ref)
        trace = self.trace
        try:  # newer Excel copies a picture only from the active sheet
            ws.Activate()
        except Exception as e:
            trace.append("Worksheet.Activate: %s" % _error_text(e))
        last = None
        # Two orders: chart first, then copy, then activate the chart and paste
        # (Excel 2013+ loses the clipboard when a chart object is added after
        # the copy, and pastes only into an active chart), and Excel 2007's
        # copy-then-chart.  xlPicture (a metafile: text drawn properly, scales
        # cleanly) before xlBitmap (text without smoothing, looks bold).
        for order in ("chart-first", "copy-first"):
            for fmt, label in ((-4147, "picture"), (2, "bitmap")):
                for attempt in range(2):
                    chart, step = None, "start"
                    try:
                        w, h = rng.Width, rng.Height
                        if order == "chart-first":
                            step = "ChartObjects.Add"
                            chart = ws.ChartObjects().Add(rng.Left, rng.Top, w + 4, h + 4)
                            step = "Range.CopyPicture"
                            rng.CopyPicture(1, fmt)  # xlScreen
                            step = "ChartObject.Activate"
                            chart.Activate()
                        else:
                            step = "Range.CopyPicture"
                            rng.CopyPicture(1, fmt)
                            step = "ChartObjects.Add"
                            chart = ws.ChartObjects().Add(rng.Left, rng.Top, w + 4, h + 4)
                        step = "Chart.Paste"
                        chart.Chart.Paste()
                        step = "scale"
                        scale = _scale_pasted(chart, w, h, SCREENSHOT_SCALE)
                        step = "Chart.Export"
                        chart.Chart.Export(path, "PNG")
                        trace.append("%s, %s: ok" % (order, label))
                        return NotJudged("screenshot saved: screens/%s (%s, %s, x%s)"
                                         % (name, label, order, scale))
                    except Exception as e:
                        last = e
                        trace.append("%s, %s, try %d: %s failed: %s"
                                     % (order, label, attempt + 1, step, _error_text(e)))
                        time.sleep(0.5)
                    finally:
                        if chart is not None:
                            try:
                                chart.Delete()
                            except Exception:
                                pass
        raise RuntimeError("screenshot failed: %s" % _error_text(last))

    def _excel_major(self):
        """Excel's major version (12 = 2007); 99 if unknown (fakes, tests)."""
        try:
            return int(float(str(self.wb.Application.Version)))
        except Exception:
            return 99

    def cf_types(self, sheet, ref):
        """Types of the rules covering ``ref``, by two routes.

        Excel 2007 misreported the type of the last rule when a cell's
        ``FormatConditions`` were indexed, and the sheet-wide walk raised
        0x800A03EC.  Route 1 indexes the sheet's rules and tests
        ``AppliesTo.Address`` in Python; if any step fails, route 2 indexes
        the cell's rules.  Both take the kind from the rule object's COM
        interface (ColorScale, Databar, ...) and read Type, through the
        object's own IDispatch, only for plain FormatCondition rules.  Route 2
        refuses duplicate priorities (the same rule returned twice).  Every
        step goes into ``self.trace``, which the report shows."""
        trace = self.trace = []
        trust = self._excel_major() >= 14  # Excel 2007's FormatCondition.Type is unreliable
        if not trust:
            trace.append("Excel %s: plain rules counted as formatCondition" % self._excel_major())
        ws = self.wb.Worksheets(sheet)
        row, col = _cell_rc(ref)
        try:
            fcs = _step(trace, "sheet rules", getattr, ws.Cells, "FormatConditions")
            n = int(_step(trace, "sheet Count", getattr, fcs, "Count"))
            types = []
            for i in range(1, n + 1):
                fc = _step(trace, "sheet Item(%d)" % i, fcs.Item, i)
                where = _step(trace, "sheet Item(%d).AppliesTo" % i, getattr, fc, "AppliesTo")
                address = str(_step(trace, "sheet Item(%d).AppliesTo.Address" % i,
                                    _english, where, "Address"))
                if _covers(address, row, col):
                    types.append(_step(trace, "sheet Item(%d) kind" % i, _cf_kind, fc, trust,
                                       trace, "sheet Item(%d)" % i))
            trace.append("route 1 (sheet rules by AppliesTo) ok")
            return sorted(types)
        except Exception:
            pass
        rng = ws.Range(ref)
        fcs = _step(trace, "cell rules", getattr, rng, "FormatConditions")
        n = int(_step(trace, "cell Count", getattr, fcs, "Count"))
        types, seen = [], {}
        for i in range(1, n + 1):
            fc = _step(trace, "cell Item(%d)" % i, fcs.Item, i)
            kind = _step(trace, "cell Item(%d) kind" % i, _cf_kind, fc, trust,
                         trace, "cell Item(%d)" % i)
            try:
                prio = int(_english(fc, "Priority"))
            except Exception as e:
                prio = None
                trace.append("cell Item(%d).Priority: %s" % (i, _error_text(e)))
            trace.append("cell Item(%d): %s %s priority %s"
                         % (i, _interface_name(fc), kind, prio))
            if prio is not None and prio in seen:
                raise RuntimeError("cell Item(%d) repeats Item(%d) (priority %d)"
                                   % (i, seen[prio], prio))
            if prio is not None:
                seen[prio] = i
            types.append(kind)
        trace.append("route 2 (cell rules) used")
        return sorted(types)

    def table_names(self, sheet):
        los = self.wb.Worksheets(sheet).ListObjects
        return sorted(los(i).Name for i in range(1, los.Count + 1))

    def table(self, sheet, name, attr):
        lo = self.wb.Worksheets(sheet).ListObjects(name)
        if attr == "ref":
            return lo.Range.Address.replace("$", "")
        if attr == "totals":
            return bool(lo.ShowTotals)
        if attr == "columns":
            cols = lo.ListColumns
            return [cols(i).Name for i in range(1, cols.Count + 1)]
        if attr == "style":
            style = lo.TableStyle
            return getattr(style, "Name", style)
        raise ValueError(attr)

    def comment_count(self, sheet):
        return self.wb.Worksheets(sheet).Comments.Count

    def comment(self, sheet, ref):
        c = self.wb.Worksheets(sheet).Range(ref).Comment
        if c is None:
            return None, None
        return c.Author, c.Text()

    def hyperlink(self, sheet, ref):
        links = self.wb.Worksheets(sheet).Range(ref).Hyperlinks
        if links.Count == 0:
            return None
        link = links(1)
        return link.Address if link.Address else "#" + link.SubAddress

    def page(self, sheet, attr):
        ps = self.wb.Worksheets(sheet).PageSetup
        if attr == "orientation":
            return {1: "portrait", 2: "landscape"}.get(ps.Orientation)
        if attr == "paper":
            return {1: "letter", 5: "legal", 8: "a3", 9: "a4", 11: "a5"}.get(ps.PaperSize, ps.PaperSize)
        if attr == "fit_width":
            return ps.FitToPagesWide
        if attr == "print_area":
            return (_english(ps, "PrintArea") or "").replace("$", "")
        if attr == "print_titles":
            return (_english(ps, "PrintTitleRows") or "").replace("$", "")
        return {"header_center": lambda: ps.CenterHeader, "footer_left": lambda: ps.LeftFooter,
                "footer_right": lambda: ps.RightFooter}[attr]()

    def sheet_option(self, sheet, attr):
        ws = self.wb.Worksheets(sheet)
        if attr == "tab_color":
            color = ws.Tab.Color
            return None if color in (False, None) else _hex_color(int(color))
        if attr == "hidden":
            return {-1: False, 0: True, 2: "very"}.get(ws.Visible, ws.Visible)
        if attr == "protected":
            return bool(ws.ProtectContents)
        raise ValueError(attr)

    def validation(self, sheet, ref):
        """List items, with the regional list separator turned into ","."""
        validation = self.wb.Worksheets(sheet).Range(ref).Validation
        f1 = _english(validation, "Formula1")
        sep = self._separators()[2]
        return f1.replace(sep, ",") if sep != "," else f1


class PowerPointAdapter:
    def __init__(self, pres):
        self.pres = pres
        self.trace = []
        self.screens: tuple = (None, None)

    def _version(self):
        try:
            return int(float(str(self.pres.Application.Version)))
        except Exception:
            return 99

    def _chart_shapes(self, slide):
        """Shapes holding a chart: HasChart, or (if PowerPoint lacks it) Type 3 (msoChart)."""
        shapes = self.pres.Slides(slide).Shapes
        found = []
        self.trace = []
        for i in range(1, shapes.Count + 1):
            sh = shapes(i)
            try:
                has = bool(sh.HasChart)
                how = "HasChart=%s" % has
            except Exception as e:
                has = int(sh.Type) == 3
                how = "HasChart failed (%s); Type=%s" % (_error_text(e), sh.Type)
            self.trace.append("slide %d shape %d: %s" % (slide, i, how))
            if has:
                found.append(sh)
        return found

    def chart_count(self, slide):
        return len(self._chart_shapes(slide))

    def chart(self, slide, index, attr, series=1):
        return _chart_om(self._version(),
                         lambda: com_chart(self._chart_shapes(slide)[index - 1].Chart, attr, series))

    def layouts(self):
        lays = self.pres.SlideMaster.CustomLayouts
        return [str(lays(i).Name) for i in range(1, lays.Count + 1)]

    def slide_layout(self, slide):
        return str(self.pres.Slides(slide).CustomLayout.Name)

    def _table(self, slide):
        shapes = self.pres.Slides(slide).Shapes
        for i in range(1, shapes.Count + 1):
            if shapes(i).HasTable:
                return shapes(i).Table
        raise LookupError("no table on slide %d" % slide)

    def table(self, slide, attr, row=None, col=None):
        t = self._table(slide)
        if attr == "rows":
            return int(t.Rows.Count)
        if attr == "cols":
            return int(t.Columns.Count)
        if attr == "cell":
            return str(t.Cell(row, col).Shape.TextFrame.TextRange.Text)
        raise ValueError("unknown table attribute %r" % attr)

    def sections(self):
        """[[name, first slide, slide count], ...] (SectionProperties, PowerPoint 2010+)."""
        def read():
            sp = self.pres.SectionProperties
            return [[str(sp.Name(i)), int(sp.FirstSlide(i)), int(sp.SlidesCount(i))]
                    for i in range(1, sp.Count + 1)]
        if self._version() < 14:
            return NotJudged("sections are PowerPoint 2010+; 2007 ignores them")
        return read()

    def links(self, slide):
        """Click hyperlinks on the slide's shapes and text runs, read through
        ActionSettings (ppActionHyperlink): the URL, or "slide:N" for a jump
        to slide N (SubAddress "id,N,title").  Slide.Hyperlinks is only
        traced: on PowerPoint 2007 it listed the shape links and missed the
        text runs' (ppt-10, 2026-11-15)."""
        sl = self.pres.Slides(slide)
        self.trace = []
        try:
            coll = sl.Hyperlinks
            self.trace.append("Slide.Hyperlinks: %s" % [
                (str(coll.Item(i).Address or ""), str(coll.Item(i).SubAddress or ""))
                for i in range(1, coll.Count + 1)])
        except Exception as e:
            self.trace.append("Slide.Hyperlinks failed: %s" % _error_text(e))

        def target(settings):
            if int(settings.Action) != 7:          # not ppActionHyperlink
                return None
            h = settings.Hyperlink
            address = str(h.Address or "")
            if address:
                return address
            parts = str(h.SubAddress or "").split(",")
            return "slide:%s" % (parts[1] if len(parts) > 1 else parts[0])

        out = []
        shapes = sl.Shapes
        for i in range(1, shapes.Count + 1):
            sh = shapes(i)
            got = target(sh.ActionSettings(1))
            if got:
                out.append(got)
            if not sh.HasTextFrame or not sh.TextFrame.HasText:
                continue
            runs = sh.TextFrame.TextRange.Runs()
            previous = None
            for j in range(1, runs.Count + 1):
                got = target(sh.TextFrame.TextRange.Runs(j, 1).ActionSettings(1))
                if got and got != previous:        # one link across runs counts once
                    out.append(got)
                previous = got
        return out

    def action(self, slide, shape):
        """The shape's click action (ActionSettings(ppMouseClick).Action)."""
        act = int(self.pres.Slides(slide).Shapes(shape).ActionSettings(1).Action)
        return PP_ACTIONS.get(act, str(act))

    def notes(self, slide):
        """Text of the notes page's body placeholder (ppPlaceholderBody = 2)."""
        shapes = self.pres.Slides(slide).NotesPage.Shapes
        for i in range(1, shapes.Count + 1):
            sh = shapes(i)
            if int(sh.Type) == 14 and int(sh.PlaceholderFormat.Type) == 2:
                return str(sh.TextFrame.TextRange.Text).replace("\r", "\n") or None
        return None

    def shape_types(self, slide):
        """Preset geometry names of the auto shapes (msoAutoShape = 1)."""
        names = {v: k for k, v in AUTO_SHAPES.items()}
        shapes = self.pres.Slides(slide).Shapes
        out = []
        for i in range(1, shapes.Count + 1):
            sh = shapes(i)
            if int(sh.Type) == 1:
                out.append(names.get(int(sh.AutoShapeType), str(sh.AutoShapeType)))
        return out

    def picture_placeholder(self, slide):
        """A picture placeholder (ppPlaceholderPicture = 18) holding a picture
        (ContainedType = msoPicture, PowerPoint 2010+)."""
        shapes = self.pres.Slides(slide).Shapes
        for i in range(1, shapes.Count + 1):
            sh = shapes(i)
            if int(sh.Type) == 14 and int(sh.PlaceholderFormat.Type) == 18:
                return _chart_om(self._version(), _contains_picture, sh)
        return False

    def slide_image(self, slide):
        folder, prefix = self.screens
        if not folder:
            return NotJudged("no screenshot folder")
        os.makedirs(folder, exist_ok=True)
        name = "%s-slide%d.png" % (prefix, slide)
        # keep the slide's proportions (a 4:3 deck exported at 1280x720 came out stretched)
        width = 1280
        height = round(width * float(self.pres.PageSetup.SlideHeight)
                       / float(self.pres.PageSetup.SlideWidth))
        self.pres.Slides(slide).Export(os.path.join(folder, name), "PNG", width, height)
        return NotJudged("slide image saved: screens/%s" % name)

    def slides(self):
        return self.pres.Slides.Count

    def shapes(self, slide):
        return self.pres.Slides(slide).Shapes.Count

    def property(self, name):
        return _builtin_property(self.pres, name)

    def custom_property(self, name):
        return _custom_property(self.pres, name)

    def slide_text(self, slide):
        parts = []
        shapes = self.pres.Slides(slide).Shapes
        for i in range(1, shapes.Count + 1):
            sh = shapes(i)
            if sh.HasTextFrame:
                parts.append(sh.TextFrame.TextRange.Text)
        return "\n".join(parts)


# --------------------------------------------------------------------------------------
# Windows machinery
# --------------------------------------------------------------------------------------
def process_ids(exe):
    out = subprocess.run(["tasklist", "/FO", "CSV", "/NH", "/FI", "IMAGENAME eq " + exe],
                         capture_output=True, text=True, check=False).stdout
    pids = set()
    for line in out.splitlines():
        cols = [c.strip('"') for c in line.split('","')]
        if len(cols) > 1 and cols[0].lower() == exe.lower():
            pids.add(int(cols[1]))
    return pids


def kill(pids):
    for pid in pids:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True, check=False)


class Watchdog(threading.Thread):
    """Captures and closes Office dialogs; kills Office after ``timeout``."""

    def __init__(self, pids, timeout):
        threading.Thread.__init__(self, daemon=True)
        self.pids, self.timeout = set(pids), timeout
        self.dialogs, self.timed_out = [], False
        self._seen, self._stop = set(), threading.Event()

    def stop(self):
        self._stop.set()

    def run(self):
        import win32gui
        import win32process

        start = time.time()
        while not self._stop.is_set():
            def visit(hwnd, _):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid in self.pids and win32gui.IsWindowVisible(hwnd):
                        cls = win32gui.GetClassName(hwnd)
                        if is_dialog_class(cls):
                            self._record(hwnd, cls, win32gui)
                except Exception:
                    pass
                return True
            win32gui.EnumWindows(visit, None)
            if time.time() - start > self.timeout:
                self.timed_out = True
                kill(self.pids)
                return
            time.sleep(0.5)

    def _record(self, hwnd, cls, win32gui):
        if hwnd not in self._seen:
            self._seen.add(hwnd)
            texts = []

            def child(h, _):
                t = win32gui.GetWindowText(h)
                if t and t not in texts:
                    texts.append(t)
                return True
            try:
                win32gui.EnumChildWindows(hwnd, child, None)
            except Exception:
                pass
            self.dialogs.append({"title": win32gui.GetWindowText(hwnd), "class": cls,
                                 "texts": texts, "time": time.time()})
        win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)  # like pressing Esc / No


def _pywin32_version():
    """pywin32 build ("311"), or None."""
    try:
        from importlib import metadata

        return metadata.version("pywin32")
    except Exception:
        try:
            import win32api

            info = win32api.GetFileVersionInfo(win32api.__file__, "\\")
            return str(info["FileVersionLS"] >> 16)
        except Exception:
            return None


def _binding(obj):
    """"early" if pywin32 wrapped ``obj`` with a makepy (gen_py) class, else "dynamic"."""
    return "early" if type(obj).__module__.startswith("win32com.gen_py") else "dynamic"


def run_case(case, root, timeout, visible):
    import pythoncom
    import win32com.client

    app_name = case["app"]
    path = os.path.abspath(os.path.join(root, case["file"]))
    result = {"id": case["id"], "app": app_name, "file": case["file"]}
    if case.get("sha256") and sha256(path) != case["sha256"]:
        result.update(status="error", error="file does not match manifest sha256 (corpus copied "
                      "incompletely, or regenerated without updating the manifest)")
        return result, None
    pythoncom.CoInitialize()
    before = process_ids(EXE[app_name])
    started = time.time()
    app = win32com.client.DispatchEx(PROGID[app_name])
    new_pids = process_ids(EXE[app_name]) - before
    version = str(app.Version)
    result["binding"] = _binding(app)
    dog = Watchdog(new_pids, timeout)
    doc, adapter, open_error, probes = None, None, None, []
    try:
        # Alerts must be ON for the watchdog to see repair prompts.  Set them
        # explicitly: PowerPoint's automation default is ppAlertsNone, which
        # silently takes the default answer (it hid the PowerPoint control's
        # prompt on the first Office 2007 run).
        if app_name == "word":
            app.Visible = visible
            app.DisplayAlerts = -1  # wdAlertsAll
            dog.start()
            doc = app.Documents.Open(path, False, True, False)  # ConfirmConversions, ReadOnly
            adapter = WordAdapter(doc)
        elif app_name == "excel":
            app.Visible = visible
            app.DisplayAlerts = True
            dog.start()
            doc = app.Workbooks.Open(path, 0, True)  # UpdateLinks=0, ReadOnly
            app.CalculateFull()
            adapter = ExcelAdapter(doc)
            adapter.screens = (os.path.join(os.path.abspath(root), "screens"),
                               "%s-%s" % (socket.gethostname(), case["id"]))
        else:
            app.DisplayAlerts = 2  # ppAlertsAll (default is ppAlertsNone = 1)
            dog.start()
            doc = app.Presentations.Open(path, True, False, False)  # ReadOnly, Untitled, WithWindow
            adapter = PowerPointAdapter(doc)
            adapter.screens = (os.path.join(os.path.abspath(root), "screens"),
                               "%s-%s" % (socket.gethostname(), case["id"]))
        time.sleep(1.0)  # let late dialogs surface
        probes = [evaluate_probe(p, adapter) for p in case.get("probes", [])]
    except Exception as e:
        open_error = repr(e)
    finally:
        try:
            if doc is not None:
                if app_name == "word":
                    doc.Close(0)  # wdDoNotSaveChanges
                elif app_name == "excel":
                    doc.Close(False)
                else:
                    doc.Close()
            app.Quit()
        except Exception:
            pass
        dog.stop()
        time.sleep(0.5)
        kill(process_ids(EXE[app_name]) - before)  # never leave our instance behind
        pythoncom.CoUninitialize()
    dialogs, expected = split_expected(dog.dialogs, getattr(adapter, "expected_windows", []))
    result.update(
        status=decide_status(open_error, dialogs, probes, dog.timed_out),
        dialogs=dialogs, probes=probes, seconds=round(time.time() - started, 1))
    if expected:
        result["expected_dialogs"] = expected
    if open_error:
        result["error"] = open_error
    classify_environment(result)
    return result, version


def select_cases(cases, apps, only):
    """Cases for these apps whose id contains any of the comma-separated
    ``only`` terms (all cases when ``only`` is empty)."""
    terms = [t.strip() for t in only.split(",") if t.strip()]
    return [c for c in cases if c["app"] in apps and "skipped" not in c
            and (not terms or any(t in c["id"] for t in terms))]


def default_out(corpus, host, label, when):
    """results-<host>-<apps+versions>-<YYYYmmdd-HHMMSS>.json: every run gets its
    own file, and report.py merges them (newest result per case wins)."""
    return os.path.join(corpus, "results-%s-%s-%s.json"
                        % (host, label or "none", when.strftime("%Y%m%d-%H%M%S")))


def main(argv=None):
    ap = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    ap.add_argument("corpus", help="folder with manifest.json")
    ap.add_argument("--apps", default=",".join(APPS), help="comma list: word,excel,powerpoint")
    ap.add_argument("--only", default="",
                    help="run case ids containing this text; comma list for several "
                         "(word-15,word-16) or a kind (control, observe)")
    ap.add_argument("--timeout", type=int, default=90, help="seconds per file before killing Office")
    ap.add_argument("--hidden", action="store_true", help="keep Word/Excel windows hidden")
    ap.add_argument("--out", help="results file (default: CORPUS/results-<host>-<versions>-"
                                  "<date-time>.json, so reruns never overwrite earlier results)")
    args = ap.parse_args(argv)
    if sys.platform != "win32":
        sys.exit("check_office.py drives Microsoft Office through COM and only runs on Windows.")
    with open(os.path.join(args.corpus, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    apps = [a for a in args.apps.split(",") if a]
    if "powerpoint" in apps and process_ids(EXE["powerpoint"]):
        sys.exit("Close PowerPoint first (it is single-instance and would be reused).")

    results, versions = [], {}
    started = datetime.datetime.now()
    todo = select_cases(manifest["cases"], apps, args.only)
    for n, case in enumerate(todo, 1):
        print("[%d/%d] %-34s" % (n, len(todo), case["id"]), end=" ", flush=True)
        res, version = run_case(case, args.corpus, args.timeout, not args.hidden)
        if version:
            versions[case["app"]] = version
        res["passed"] = passed(case, res["status"])
        if case.get("control"):
            res["control"] = True
        if case.get("observe"):
            res["observe"] = True
            extra_obs = " | observed (no expectation)"
        else:
            extra_obs = ""
        results.append(res)
        extra = ""
        if res.get("dialogs"):
            extra = " | dialog: " + "; ".join(d["title"] or d["class"] for d in res["dialogs"])
        elif res.get("error"):
            extra = " | " + res["error"][:100]
        elif res["status"] == "mismatch":
            extra = " | probes failed: " + ", ".join(p["kind"] for p in res["probes"] if not p.get("ok"))
        if case.get("control"):
            extra += " | control: " + ("detected as expected" if res["passed"]
                                       else "NOT DETECTED - checker or Office behaviour needs a look")
        print(res["status"] + extra + extra_obs)

    label = "-".join("%s%s" % (a, versions[a].split(".")[0]) for a in apps if a in versions)
    out = args.out or default_out(args.corpus, socket.gethostname(), label, started)
    doc = {
        "kind": "oox-officecheck-results", "format": 1,
        "created": datetime.datetime.now().isoformat(timespec="seconds"),
        "machine": {"host": socket.gethostname(), "os": platform.platform(),
                    "python": platform.python_version(), "bits": 64 if sys.maxsize > 2**32 else 32,
                    "executable": sys.executable, "pywin32": _pywin32_version()},
        "apps": {a: {"version": v, "office": office_name(v)} for a, v in versions.items()},
        "manifest": {"oox_version": manifest.get("oox_version"),
                     "grammar_source": manifest.get("grammar_source")},
        "results": results,
    }
    with open(out, "w", encoding="utf-8") as f:
        json.dump(doc, f, indent=2, ensure_ascii=False)
    bad = [r for r in results if not r["passed"]]
    print("\n%d passed, %d failed -> %s" % (len(results) - len(bad), len(bad), out))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
