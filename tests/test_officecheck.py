"""The Office check kit: generator, portable checker logic, report."""

import ast
import csv
import json
import os
import shutil
from pathlib import Path

import pytest

from oox import Package
from tools.officecheck import check_office, report
from tools.officecheck.generate import cases, generate

KIT = Path(__file__).parents[1] / "tools" / "officecheck"


@pytest.mark.parametrize("script", ["check_office.py", "report.py"])
def test_windows_side_scripts_parse_as_python_38(script):
    """These run on the Office machines (Windows 7 tops out at Python 3.8)."""
    ast.parse((KIT / script).read_text(encoding="utf-8"), feature_version=(3, 8))
    src = (KIT / script).read_text(encoding="utf-8")
    assert "import oox" not in src and "from oox" not in src  # standalone


# observe cases whose values oox accepts since Office opened them
# (grammar.EXTRA_ENUM_VALUES; Office 2007, 2026-11-15)
ACCEPTED_AFTER_OBSERVING = {"word-94-observe-sdk-values"}


def test_corpus_is_deterministic_and_valid(tmp_path):
    a, b = generate(tmp_path / "a"), generate(tmp_path / "b")
    assert [c.get("sha256") for c in a["cases"]] == [c.get("sha256") for c in b["cases"]]
    controls = [c for c in a["cases"] if c.get("control")]
    assert len(controls) == 3
    assert all(c["expect_status"] == ["prompt", "error", "mismatch"] and c["probes"] for c in controls)
    for c in a["cases"]:
        if "skipped" in c:
            continue
        if c.get("observe"):  # schema-invalid on purpose: validate() must say so
            from oox import open_document

            report = open_document(tmp_path / "a" / c["file"]).validate(target="Office2007")
            if c["id"] in ACCEPTED_AFTER_OBSERVING:
                assert not report, f"{c['id']}: values accepted since Office opened them"
                continue
            assert report, f"{c['id']} should break the schema in some way"
            continue
        if c.get("control"):  # broken on purpose: the malformed part must not parse
            pkg = Package.open(tmp_path / "a" / c["file"])
            with pytest.raises(Exception):
                [p.element for p in pkg.parts.values() if p.is_xml]
            continue
        assert c["probes"], c["id"]
        pkg = Package.open(tmp_path / "a" / c["file"])
        for part in pkg.parts.values():
            if part.is_xml:
                assert part.element.validate() == [], (c["id"], part.partname)
    assert (tmp_path / "a" / "CHECKLIST.md").read_text(encoding="utf-8").count("☐") == len(cases())
    assert sum(1 for c in a["cases"] if c.get("observe")) == 11
    with open(tmp_path / "a" / "manual-results.csv", encoding="utf-8") as f:
        assert [r["id"] for r in csv.DictReader(f)] == [c["id"] for c in a["cases"] if "skipped" not in c]


class FakeWord:
    def text(self):
        return "Title\rBody with   spaces\x07cell\x0bbreak"

    def tables(self):
        return 1

    def inline_shapes(self):
        return 5

    def pages(self):
        return 2


class FakeExcel:
    def sheets(self):
        return ["A", "B"]

    def cell(self, sheet, ref):
        return {"B1": 42.0, "B4": True, "B5": "x"}.get(ref)

    def shapes(self, sheet):
        raise RuntimeError("sheet not found")


def test_probe_evaluation():
    ev = check_office.evaluate_probe
    assert ev({"kind": "word.text_contains", "value": "Body with   spaces"}, FakeWord())["ok"]
    assert ev({"kind": "word.text_contains", "value": "cell\nbreak"}, FakeWord())["ok"]
    assert not ev({"kind": "word.text_contains", "value": "missing"}, FakeWord())["ok"]
    assert ev({"kind": "word.tables", "expect": 1}, FakeWord())["ok"]
    bad = ev({"kind": "word.inline_shapes", "expect": 4}, FakeWord())
    assert (bad["ok"], bad["got"]) == (False, 5)
    assert ev({"kind": "excel.cell", "sheet": "A", "ref": "B1", "expect": 42}, FakeExcel())["ok"]
    assert ev({"kind": "excel.cell", "sheet": "A", "ref": "B4", "expect": True}, FakeExcel())["ok"]
    assert not ev({"kind": "excel.cell", "sheet": "A", "ref": "B5", "expect": 5}, FakeExcel())["ok"]
    assert ev({"kind": "excel.sheets", "expect": ["A", "B"]}, FakeExcel())["ok"]
    err = ev({"kind": "excel.shapes", "sheet": "Z", "expect": 1}, FakeExcel())
    assert not err["ok"] and "sheet not found" in err["error"]
    assert not ev({"kind": "nonsense"}, FakeExcel())["ok"]


def test_status_and_helpers():
    st = check_office.decide_status
    assert st(None, [], [{"ok": True}], False) == "ok"
    assert st(None, [], [{"ok": False}], False) == "mismatch"
    assert st("boom", [], [], False) == "error"
    assert st("boom", [{"title": "Repair?"}], [], False) == "prompt"
    assert st(None, [], [], True) == "timeout"
    assert check_office.is_dialog_class("#32770") and check_office.is_dialog_class("bosa_sdm_msword")
    assert not check_office.is_dialog_class("OpusApp")
    assert check_office.office_name("12.0") == "Office 2007"
    assert check_office.office_name("16.0") == "Office 2016+"
    control = {"control": True, "expect_status": ["prompt", "error", "mismatch"]}
    assert check_office.passed(control, "prompt") and check_office.passed(control, "error")
    assert check_office.passed(control, "mismatch")  # silent repair that lost content
    assert not check_office.passed(control, "ok") and not check_office.passed(control, "timeout")
    assert check_office.passed({}, "ok") and not check_office.passed({}, "prompt")
    assert check_office.passed({"observe": True}, "prompt")  # recorded, never a failure


def test_dialogs_provoked_by_a_probe_are_expected():
    """Word shows "The password is incorrect" besides raising the COM error
    (Office 2007, word-25/26, 2026-11-15): a dialog inside the wrong-password
    attempt's window is the probe working, not a prompt about the file."""
    class Doc:
        ProtectionType = 3

        def Unprotect(self, pw=None):
            if pw != "secret":
                raise RuntimeError("The password is incorrect.")
            self.ProtectionType = -1

    adapter = check_office.WordAdapter(Doc())
    got = check_office.evaluate_probe({"kind": "word.unprotect", "wrong": "wrong",
                                       "password": "secret", "expect": [True, True, "none"]},
                                      adapter)
    assert got["ok"]
    (start, end), = adapter.expected_windows
    during = {"title": "Microsoft Office Word", "texts": ["The password is incorrect."],
              "time": end + 0.5}
    at_open = {"title": "Repair?", "time": start - 5}
    undated = {"title": "Old result"}
    unexpected, expected = check_office.split_expected([at_open, during, undated],
                                                       adapter.expected_windows)
    assert expected == [during] and unexpected == [at_open, undated]
    assert check_office.decide_status(None, unexpected[1:], [got], False) == "prompt"
    assert check_office.decide_status(None, [], [got], False) == "ok"
    assert check_office.split_expected([during], []) == ([during], [])


LICENCE = ("com_error(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The Unprotect "
           "method or property is not available because the licence to use this application "
           "has expired..', 'wdmain11.chm', 37373, -2146823683), None)")
SIGN_IN = {"title": "Sign in to set up Office", "class": "NUIDialog", "texts": []}


def test_unactivated_office_is_told_apart_from_the_files(tmp_path):
    """Office 2016 run, 2026-11-16: an unactivated Office showed "Sign in to set
    up Office" in most cases and refused calls "because the licence ... has
    expired"; neither is about the file."""
    classify = check_office.classify_environment
    only_signin = classify({"status": "prompt", "dialogs": [SIGN_IN],
                            "probes": [{"kind": "word.text_boxes", "ok": True}]})
    assert only_signin["status"] == "ok" and only_signin["dialogs"] == []
    assert only_signin["environment_dialogs"] == [SIGN_IN]
    refused = classify({"status": "prompt", "dialogs": [SIGN_IN], "probes": [
        {"kind": "word.protection", "ok": True},
        {"kind": "word.unprotect", "ok": False, "got": [True, False, "readOnly"],
         "trace": ["Unprotect(wrong): " + LICENCE]}]})
    assert refused["status"] == "unactivated" and refused["probes"][1]["refused"]
    real = classify({"status": "prompt", "dialogs": [SIGN_IN], "probes": [
        {"kind": "word.tables", "ok": False, "got": 2, "expect": 1}]})
    assert real["status"] == "mismatch"  # a real mismatch still shows
    repair = classify({"status": "prompt", "dialogs": [
        SIGN_IN, {"title": "Microsoft Word", "texts": ["Word found unreadable content"]}],
        "probes": []})
    assert repair["status"] == "prompt" and len(repair["dialogs"]) == 1
    control = classify({"status": "prompt", "error": "com_error(...)", "probes": [],
                        "dialogs": [{"title": "Microsoft Word", "texts": ["error"]}]})
    assert control["status"] == "prompt"
    assert classify(dict(refused)) == refused  # idempotent
    observe = classify({"status": "error", "error": "Open method failed", "dialogs": [],
                        "probes": []})
    assert observe["status"] == "error"

    # the report applies it to results files written before the checker did
    ids = ["word-24-text-boxes", "word-25-protection-read-only", "word-12-sections"]
    manifest = {"oox_version": "x", "grammar_source": "y",
                "cases": [{"id": i, "app": "word"} for i in ids]}
    (tmp_path / "results-pc-word16-20261116-183434.json").write_text(json.dumps({
        "created": "2026-11-16T18:34:34", "machine": {"host": "pc"},
        "apps": {"word": {"version": "16.0", "office": "Office 2016+"}},
        "results": [
            {"id": ids[0], "status": "prompt", "dialogs": [SIGN_IN],
             "probes": [{"kind": "word.text_boxes", "ok": True}]},
            {"id": ids[1], "status": "prompt", "dialogs": [SIGN_IN],
             "probes": [{"kind": "word.unprotect", "ok": False, "got": [True, False, "x"],
                         "trace": ["Unprotect(wrong): " + LICENCE]}]},
            {"id": ids[2], "status": "ok", "dialogs": [], "probes": []}]}),
        encoding="utf-8")
    text = report.render(manifest, report.load_sources(str(tmp_path)))
    rows = {line.split("|")[1].strip(): line for line in text.splitlines()
            if line.startswith("| `")}
    assert "✅" in rows[f"`{ids[0]}`"] and "🔒" in rows[f"`{ids[1]}`"]
    assert "Office isn't activated" in text and "2 cases and 1 probes" in text
    assert "refused, Office isn't activated" in text and "Sign in" not in text


def test_busy_office_calls_are_retried():
    calls = []

    class Busy:
        def text(self):
            calls.append(1)
            if len(calls) < 3:
                raise RuntimeError("com_error(-2147418111, 'Call was rejected by callee.')")
            return "hello"

    got = check_office.evaluate_probe({"kind": "word.text_contains", "value": "hello"},
                                      Busy(), pause=0)
    assert got["ok"] and len(calls) == 3


@pytest.mark.skipif(os.name == "nt", reason="non-Windows guard")
def test_checker_refuses_to_run_off_windows(tmp_path):
    with pytest.raises(SystemExit, match="only runs on Windows"):
        check_office.main([str(tmp_path)])


def test_report_merges_scripted_and_manual(tmp_path):
    manifest = generate(tmp_path)
    first, second = manifest["cases"][0]["id"], manifest["cases"][1]["id"]
    (tmp_path / "results-pc1-word16.json").write_text(json.dumps({
        "machine": {"host": "pc1"}, "apps": {"word": {"version": "16.0", "office": "Office 2016+"}},
        "results": [
            {"id": first, "status": "ok", "dialogs": [], "probes": []},
            {"id": second, "status": "prompt", "probes": [],
             "dialogs": [{"title": "Microsoft Word", "class": "#32770",
                          "texts": ["Word found unreadable content"]}]},
        ]}), encoding="utf-8")
    with open(tmp_path / "manual-results-office2007.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["id", "status", "notes"])
        wr.writerow([first, "mismatch", "picture missing"])
    text = report.render(manifest, report.load_sources(str(tmp_path)))
    assert "pc1 (Office 2016+)" in text and "office2007 (manual)" in text
    assert "Word found unreadable content" in text and "picture missing" in text
    row = next(line for line in text.splitlines() if f"`{first}`" in line)
    assert "✅" in row and "❌" in row


def test_report_judges_controls_by_expectation(tmp_path):
    manifest = generate(tmp_path)
    ctrl = [c["id"] for c in manifest["cases"] if c.get("control")]
    (tmp_path / "results-pc2-all12.json").write_text(json.dumps({
        "machine": {"host": "pc2"}, "apps": {"word": {"version": "12.0", "office": "Office 2007"}},
        "results": [
            {"id": ctrl[0], "status": "prompt", "probes": [],
             "dialogs": [{"title": "Microsoft Office Word", "class": "#32770", "texts": ["unreadable"]}]},
            {"id": ctrl[1], "status": "ok", "dialogs": [], "probes": []},  # NOT detected
            {"id": ctrl[2], "status": "mismatch", "dialogs": [],
             "probes": [{"kind": "ppt.slides", "expect": 2, "got": 1, "ok": False}]},
        ]}), encoding="utf-8")
    text = report.render(manifest, report.load_sources(str(tmp_path)))
    rows = {line.split("|")[1].strip(): line for line in text.splitlines() if line.startswith("| `")}
    assert "✅ detected" in rows[f"`{ctrl[0]}` *(control)*"]
    assert "❌ not detected (ok)" in rows[f"`{ctrl[1]}` *(control)*"]
    assert "✅ detected (silent repair, content lost)" in rows[f"`{ctrl[2]}` *(control)*"]
    assert "expected `2`, got `1`" in text  # the evidence is shown in the details


@pytest.mark.skipif(not (shutil.which("soffice") and os.environ.get("OOX_LIBREOFFICE")),
                    reason="set OOX_LIBREOFFICE=1 (needs LibreOffice; ~1 min)")
def test_libreoffice_passes_every_probe(tmp_path):
    from tools.officecheck.check_libreoffice import find_soffice, run

    generate(tmp_path)
    doc = run(tmp_path, find_soffice(None))
    assert {r["id"]: r["status"] for r in doc["results"]} == \
        {r["id"]: "ok" for r in doc["results"]}


def test_report_merges_reruns_into_one_column(tmp_path):
    manifest = generate(tmp_path)
    ids = [c["id"] for c in manifest["cases"]]
    ppt = [i for i in ids if i.startswith("ppt-")]

    def write(name, created, apps, results):
        (tmp_path / name).write_text(json.dumps({
            "created": created, "machine": {"host": "pc"},
            "apps": {a: {"version": "12.0", "office": "Office 2007"} for a in apps},
            "results": results}), encoding="utf-8")

    # full run: every case, the PowerPoint control "not detected"
    write("results-pc-word12-excel12-powerpoint12.json", "2026-11-14T02:20:00",
          ["word", "excel", "powerpoint"],
          [{"id": i, "status": "ok", "dialogs": [], "probes": []} for i in ids])
    # later rerun of PowerPoint only: the control is now detected
    write("results-pc-powerpoint12.json", "2026-11-14T02:44:00", ["powerpoint"],
          [{"id": i, "status": "prompt" if "control" in i else "ok", "probes": [],
            "dialogs": [{"title": "Microsoft Office PowerPoint", "texts": ["blank slides"]}]
            if "control" in i else []} for i in ppt])
    sources = report.load_sources(str(tmp_path))
    assert [s["label"] for s in sources] == ["pc (Office 2007)"]  # one column
    text = report.render(manifest, sources)
    rows = [line for line in text.splitlines() if line.startswith("| `")]
    assert "·" not in "".join(rows)  # nothing missing
    assert "✅ detected" in next(line for line in rows if "ppt-90-control" in line)
    assert "results-pc-word12-excel12-powerpoint12.json" in text and "results-pc-powerpoint12.json" in text


def test_shape_wraps_on_word_2007_and_newer():
    """Word 2007 reports behind and in-front both as wdWrapNone (3)."""

    class Shape:
        def __init__(self, t):
            self.WrapFormat = type("WF", (), {"Type": t})()

    class Shapes:
        def __init__(self, types):
            self.items, self.Count = [Shape(t) for t in types], len(types)

        def __call__(self, i):
            return self.items[i - 1]

    def adapter(version, types):
        doc = type("Doc", (), {})()
        doc.Shapes = Shapes(types)
        doc.Application = type("App", (), {"Version": version})()
        return check_office.WordAdapter(doc)

    probe = {"kind": "word.shape_wraps", "expect": ["behind", "square", "tight"]}
    old = check_office.evaluate_probe(probe, adapter("12.0", [0, 1, 3]))
    assert old["ok"] and old["got"] == ["none", "square", "tight"]
    new = check_office.evaluate_probe(probe, adapter("16.0", [0, 1, 5]))
    assert new["ok"] and new["got"] == ["behind", "square", "tight"]
    wrong = check_office.evaluate_probe(probe, adapter("16.0", [0, 1, 3]))
    assert not wrong["ok"]  # in front where behind was expected


def test_case_selection_and_unique_result_names():
    import datetime as dt

    cases = [{"id": "word-15-tables", "app": "word"}, {"id": "word-16-floating", "app": "word"},
             {"id": "word-90-control-malformed", "app": "word", "control": True},
             {"id": "excel-01-types", "app": "excel"}, {"id": "word-99", "app": "word",
                                                         "skipped": "no asset"}]
    ids = lambda sel: [c["id"] for c in sel]
    sel = check_office.select_cases
    assert ids(sel(cases, ["word"], "")) == ["word-15-tables", "word-16-floating",
                                             "word-90-control-malformed"]
    assert ids(sel(cases, ["word", "excel"], "word-15, word-16")) == ["word-15-tables",
                                                                     "word-16-floating"]
    assert ids(sel(cases, ["word"], "control")) == ["word-90-control-malformed"]
    a = check_office.default_out("c", "host", "word12", dt.datetime(2026, 11, 14, 12, 0, 0))
    b = check_office.default_out("c", "host", "word12", dt.datetime(2026, 11, 14, 12, 0, 1))
    assert a != b and a.endswith("results-host-word12-20261114-120000.json")


def test_excel_regional_settings_normalised():
    """Greek Excel 2007 through pywin32: International comes back as a tuple
    of all settings, separators are "," and "."."""
    settings = [1, 30, ",", ".", ";"]  # Excel's language (English), region (Greece), ...
    app = type("App", (), {"International": tuple(settings), "UseSystemSeparators": True})()

    class Rng:
        def __init__(self, text, value):
            self.Text, self.Value2, self.NumberFormat = text, value, "0,0%"
            self.Validation = type("V", (), {"Formula1": "Yes;No"})()

    cells = {"B2": Rng("1.234,50", 1234.5), "B10": Rng("a.b,c", "a.b,c"),
             "B4": Rng("Nov 14, 2026", 46340.0)}
    cells["B4"].NumberFormat = "[$-409]mmm d, yyyy"
    ws = type("WS", (), {"Range": lambda self, ref: cells[ref]})()
    wb = type("WB", (), {"Application": app, "Worksheets": lambda self, name: ws})()
    a = check_office.ExcelAdapter(wb)
    assert a.text("S", "B2") == "1,234.50"
    assert a.text("S", "B10") == "a.b,c"            # text cells are left alone
    assert a.text("S", "B4") == "Nov 14, 2026"      # dates too: "," is literal there
    assert a.number_format("S", "B2") == "0.0%"
    assert a.validation("S", "B2") == "Yes,No"
    assert not a.english_names()
    names = check_office.evaluate_probe(
        {"kind": "excel.text", "sheet": "S", "ref": "B2", "expect": "x", "names": True}, a)
    assert names["ok"] and "not judged" in names["note"]
    # the callable shape works too
    app2 = type("App", (), {"International": staticmethod(lambda i: settings[i - 1])})()
    assert check_office._international(app2, 3) == ","


class _Ole:
    def __init__(self, name):
        self.name = name

    def GetTypeInfo(self):
        return type("TI", (), {"GetDocumentation": lambda _s, i: (self.name, "", 0, "")})()


class _Rule:
    def __init__(self, typ, address, iface=None, priority=None, **extra):
        self.Type, self.Priority = typ, priority
        self._address = address
        if iface:
            self._oleobj_ = _Ole(iface)
        self.__dict__.update(extra)

    @property
    def AppliesTo(self):
        if self._address is None:
            raise RuntimeError("0x800A03EC")
        return type("R", (), {"Address": self._address})()


class _FCs:
    def __init__(self, rules):
        self.rules, self.Count = rules, len(rules)

    def Item(self, i):
        return self.rules[i - 1]


def _cf_adapter(sheet_rules, cell_rules=None, version="16.0"):
    class Sheet:
        Cells = type("Cells", (), {"FormatConditions": _FCs(sheet_rules)})()

        def Range(self, ref):
            return type("Rng", (), {"FormatConditions": _FCs(cell_rules[ref])})()

    app = type("App", (), {"Version": version})()
    wb = type("WB", (), {"Worksheets": lambda self, name: Sheet(), "Application": app})()
    return check_office.ExcelAdapter(wb)


def test_cf_types_route1_sheet_rules_and_interface():
    """Route 1 indexes the sheet's rules, matches AppliesTo.Address in Python
    and trusts the COM interface name over Type for the special classes."""
    rules = [_Rule(1, "$B$2:$B$11"),
             _Rule(1, "$C$2:$C$11", "ColorScale"),          # Type misread: interface wins
             _Rule(8, "$A$2:$A$11", "UniqueValues", DupeUnique=1),
             _Rule(2, "$A$2:$E$11", "FormatCondition"),
             _Rule(5, "$G$1,$B$3", "Top10")]
    a = _cf_adapter(rules)
    assert a.cf_types("S", "B2") == ["cellIs", "expression"]
    assert a.cf_types("S", "C2") == ["colorScale", "expression"]
    assert a.cf_types("S", "A3") == ["duplicateValues", "expression"]
    assert a.cf_types("S", "B3") == ["cellIs", "expression", "top10"]
    assert a.cf_types("S", "G1") == ["top10"]
    assert a.cf_types("S", "F2") == []
    assert a.trace[-1].startswith("route 1")
    assert check_office._covers("$A$1:$B$2;$D$4", 4, 4)
    assert check_office._cell_rc("$AA$10") == (10, 27)


def test_cf_types_route2_when_sheet_walk_fails():
    """Excel 2007 raised 0x800A03EC in the sheet walk: the cell's own rules
    are used, and a repeated priority (same rule twice) is refused."""
    sheet = [_Rule(3, None, "ColorScale")]                  # AppliesTo raises
    cell = {"C2": [_Rule(3, "", "ColorScale", priority=4),
                   _Rule(3, "", "FormatCondition", priority=9)]}  # Type misread...
    cell["C2"][1].Type = 2                                  # ...read via own IDispatch: 2
    a = _cf_adapter(sheet, cell)
    assert a.cf_types("S", "C2") == ["colorScale", "expression"]
    assert any("AppliesTo: RuntimeError" in step for step in a.trace)
    assert a.trace[-1] == "route 2 (cell rules) used"

    dup = {"C2": [_Rule(3, "", "ColorScale", priority=4), _Rule(3, "", "ColorScale", priority=4)]}
    a = _cf_adapter(sheet, dup)
    result = check_office.evaluate_probe(
        {"kind": "excel.cf_types", "sheet": "S", "ref": "C2", "expect": ["colorScale", "expression"]}, a)
    assert not result["ok"] and "repeats Item(1)" in result["error"]
    assert result["expect"] == ["colorScale", "expression"] and result["ref"] == "C2"
    assert any("priority 4" in step for step in result["trace"])


def test_report_shows_probe_traces(tmp_path):
    """Step traces appear for failed probes and for ok probes that carry one
    (to see which cf_types route Excel allowed)."""
    manifest = generate(tmp_path)
    cid = "excel-12-conditional"
    (tmp_path / "results-pc3-excel12.json").write_text(json.dumps({
        "machine": {"host": "pc3"}, "apps": {"excel": {"version": "12.0", "office": "Office 2007"}},
        "results": [{"id": cid, "status": "mismatch", "dialogs": [], "probes": [
            {"kind": "excel.cf_types", "sheet": "Rules", "ref": "B2", "ok": False,
             "expect": ["cellIs"], "error": "com_error(...)",
             "trace": ["sheet Item(1).AppliesTo: com_error 0x800A03EC"]},
            {"kind": "excel.cf_types", "sheet": "Rules", "ref": "C2", "ok": True,
             "expect": ["colorScale"], "got": ["colorScale"],
             "trace": ["route 2 (cell rules) used"]}]}]}), encoding="utf-8")
    text = report.render(manifest, report.load_sources(str(tmp_path)))
    assert "probe `excel.cf_types` sheet=Rules ref=B2: expected `['cellIs']`" in text
    assert "    - sheet Item(1).AppliesTo: com_error 0x800A03EC" in text
    assert "ref=C2: ok, got `['colorScale']`" in text and "route 2 (cell rules) used" in text


class _BadType(_Rule):
    """A plain rule as Excel 2007 serves it: Type raises, or echoes another rule."""

    def __init__(self, address, priority, echo=None, **extra):
        super().__init__(None, address, "FormatCondition", priority, **extra)
        self._echo = echo

    @property
    def Type(self):
        if self._echo is None:
            raise RuntimeError("com_error 0x800A03EC")
        return self._echo

    @Type.setter
    def Type(self, value):
        pass


def test_cf_types_excel2007_groups_plain_rules():
    """Replays the 2026-11-15 Excel 2007 trace: sheet-wide Type raised,
    cell-level Type echoed the cell's first rule.  Plain rules count as a
    group, special kinds and counts are still judged, and the derived kind
    is only traced."""
    sheet = [_BadType("$B$2:$B$11", 1, Operator=5, Formula1="=95"),
             _BadType("$B$2:$B$11", 2, Operator=1, Formula1="=30"),
             _Rule(5, "$B$2:$B$11", "Top10", 3),
             _Rule(3, "$C$2:$C$11", "ColorScale", 4),
             _BadType("$A$2:$E$11", 9, Formula1="=$E2>$D2")]
    a = _cf_adapter(sheet, version="12.0")
    probe = {"kind": "excel.cf_types", "sheet": "Rules", "ref": "B2",
             "expect": ["cellIs", "cellIs", "expression", "top10"]}
    result = check_office.evaluate_probe(probe, a)
    assert result["ok"], result
    assert result["got"] == ["formatCondition"] * 3 + ["top10"]
    assert "not judged" in result["note"]
    trace = "\n".join(result["trace"])
    assert "route 1" in trace and "derived cellIs" in trace and "derived expression" in trace

    # a wrong special kind or a missing rule still fails
    bad = dict(probe, expect=["cellIs", "cellIs", "colorScale", "expression"])
    assert not check_office.evaluate_probe(bad, a)["ok"]
    short = dict(probe, expect=["cellIs", "expression", "top10"])
    assert not check_office.evaluate_probe(short, a)["ok"]

    # cell route as in the report (sheet walk broken): Type echoes rule 1
    broken = [_Rule(3, None, "ColorScale")]
    cell = {"C2": [_Rule(3, "", "ColorScale", 4), _BadType("", 9, echo=3, Formula1="=$E2>$D2")]}
    a = _cf_adapter(broken, cell, version="12.0")
    result = check_office.evaluate_probe(
        {"kind": "excel.cf_types", "sheet": "Rules", "ref": "C2",
         "expect": ["colorScale", "expression"]}, a)
    assert result["ok"] and result["got"] == ["colorScale", "formatCondition"]
    assert result["trace"][-1] == "route 2 (cell rules) used"

    # newer Excel: Type is trusted, a misread type fails
    a = _cf_adapter(broken, cell, version="16.0")
    result = check_office.evaluate_probe(
        {"kind": "excel.cf_types", "sheet": "Rules", "ref": "C2",
         "expect": ["colorScale", "expression"]}, a)
    assert not result["ok"] and result["got"] == ["colorScale", "colorScale"]


def test_excel_screenshot_via_chart_export(tmp_path):
    """Chart first, CopyPicture, activate the chart, paste, Export (Excel 2016
    pastes only into an active chart: 0x800A03EC otherwise, 2026-11-16); then
    Excel 2007's copy-first order; picture format before bitmap; every step
    traced; the chart is always deleted."""
    log = []

    class FakeChart:
        def __init__(self, owner):
            self.owner = owner

        def Paste(self):
            if not self.owner.active:
                raise RuntimeError("0x800A03EC")
            log.append("paste")

        def Export(self, path, fmt):
            assert fmt == "PNG"
            with open(path, "wb") as f:
                f.write(b"\x89PNG fake")

    class ChartObject:
        def __init__(self):
            self.active = False
            self.Chart = FakeChart(self)

        def Activate(self):
            self.active = True

        def Delete(self):
            log.append("delete")

    class Rng:
        Left, Top, Width, Height = 0, 0, 300, 200
        fail_picture = True

        def CopyPicture(self, appearance, fmt):
            log.append(("copy", appearance, fmt))
            if fmt == -4147 and Rng.fail_picture:
                raise RuntimeError("clipboard busy")

    class Sheet:
        def Activate(self):
            log.append("activate sheet")

        def Range(self, ref):
            return Rng()

        def ChartObjects(self):
            return type("COs", (), {"Add": lambda self, *a: ChartObject()})()

    wb = type("WB", (), {"Worksheets": lambda self, name: Sheet()})()
    a = check_office.ExcelAdapter(wb)
    a.screens = (str(tmp_path / "screens"), "pc-excel-12-conditional")
    probe = {"kind": "excel.screenshot", "sheet": "Rules", "ref": "A1:E11", "expect": "picture"}
    check_office.time.sleep, sleep = (lambda s: None), check_office.time.sleep
    try:
        result = check_office.evaluate_probe(probe, a)
        assert result["ok"] and "screens/pc-excel-12-conditional-Rules.png" in result["note"]
        assert "(bitmap, chart-first, x1)" in result["note"]  # the fake has no Shapes
        assert log[0] == "activate sheet" and log.count(("copy", 1, -4147)) == 2
        assert log[-2:] == ["paste", "delete"]
        assert "chart-first, picture, try 1: Range.CopyPicture failed" in result["trace"][0]
        Rng.fail_picture = False
        good = check_office.evaluate_probe(probe, a)
        assert "(picture, chart-first, x1)" in good["note"]
        # without Activate (an Excel that refuses it) the copy-first order still runs
        ChartObject.Activate = lambda self: (_ for _ in ()).throw(RuntimeError("no"))
        bad = check_office.evaluate_probe(probe, a)
        assert not bad["ok"] and "Chart.Paste failed" in bad["trace"][-1]
        assert "ChartObject.Activate failed" in bad["trace"][0]
        assert len(bad["trace"]) == 8  # 2 orders x 2 formats x 2 tries, each step named
    finally:
        check_office.time.sleep = sleep
    assert (tmp_path / "screens" / "pc-excel-12-conditional-Rules.png").read_bytes().startswith(b"\x89PNG")

    # the report lists screenshots, also for ok cases
    manifest = generate(tmp_path)
    (tmp_path / "results-pc-excel12.json").write_text(json.dumps({
        "machine": {"host": "pc"}, "apps": {"excel": {"version": "12.0", "office": "Office 2007"}},
        "results": [{"id": "excel-12-conditional", "status": "ok", "dialogs": [],
                     "probes": [result]}]}), encoding="utf-8")
    text = report.render(manifest, report.load_sources(str(tmp_path)))
    assert "## Screenshots" in text
    assert "`excel-12-conditional` on pc (Office 2007): screenshot saved: screens/" in text


def test_report_lists_interpreter_per_run(tmp_path):
    """Several Pythons on one machine: each results file says which one ran."""
    manifest = generate(tmp_path)
    for name, py, binding in (("results-pc-a.json", "3.8.10", "dynamic"),
                              ("results-pc-b.json", "3.14.0", "early")):
        (tmp_path / name).write_text(json.dumps({
            "created": "2026-11-15T03:%s:00" % ("10" if py == "3.8.10" else "20"),
            "machine": {"host": "pc", "python": py, "bits": 64,
                        "executable": "C:\\Python\\python.exe", "pywin32": "311"},
            "apps": {"excel": {"version": "12.0", "office": "Office 2007"}},
            "results": [{"id": "excel-12-conditional", "status": "ok", "dialogs": [],
                         "probes": [], "binding": binding}]}), encoding="utf-8")
    text = report.render(manifest, report.load_sources(str(tmp_path)))
    assert "## Runs" in text
    assert "`results-pc-a.json`: Python 3.8.10, 64-bit" in text and "binding dynamic" in text
    assert "`results-pc-b.json`: Python 3.14.0" in text and "pywin32 311; COM binding early" in text


def test_screenshot_scales_pasted_picture():
    class Pic:
        Left = Top = Width = Height = 0

    pic = Pic()
    chart = type("CO", (), {"Width": 0, "Height": 0})()
    chart.Chart = type("C", (), {"Shapes": staticmethod(lambda i: pic)})()
    assert check_office._scale_pasted(chart, 100, 50, 2) == 2
    assert (pic.Width, pic.Height, chart.Width, chart.Height) == (200, 100, 204, 104)


def test_chart_type_table_matches_oox():
    """check_office.py can't import oox, so it keeps its own copy."""
    from oox.charts import EXCEL_CHART_TYPES

    assert check_office.CHART_TYPES == EXCEL_CHART_TYPES


class _FakeSeries:
    def __init__(self, name, values, xvalues, labels=None):
        self.Name, self.Values, self.XValues = name, values, xvalues
        self.HasDataLabels = labels is not None
        self._labels = labels or {}

    def DataLabels(self):
        return type("DL", (), {k: self._labels.get(k, False) for k in
                               ("ShowValue", "ShowPercentage", "ShowCategoryName",
                                "ShowSeriesName")})()


class _FakeChart:
    def __init__(self):
        self.ChartType, self.HasTitle, self.HasLegend = 51, True, True
        self.ChartTitle = type("T", (), {"Text": "Sales"})()
        self.Legend = type("L", (), {"Position": -4107})()
        self._series = [_FakeSeries("Q1", (120.0, 95.0), ("North", "South"),
                                    {"ShowPercentage": True})]

    def Axes(self, which):
        title = {1: "Region", 2: None}[which]
        return type("Ax", (), {"HasTitle": title is not None,
                               "AxisTitle": type("AT", (), {"Text": title})()})()

    def SeriesCollection(self):
        series = self._series
        return type("SC", (), {"Count": len(series),
                               "Item": staticmethod(lambda i: series[i - 1])})()


def test_com_chart_reads_office_chart_objects():
    ch = _FakeChart()
    assert check_office.com_chart(ch, "type") == "column"
    assert check_office.com_chart(ch, "title") == "Sales"
    assert check_office.com_chart(ch, "legend") == "bottom"
    assert check_office.com_chart(ch, "x_title") == "Region"
    assert check_office.com_chart(ch, "y_title") is None
    assert check_office.com_chart(ch, "series") == ["Q1"]
    assert check_office.com_chart(ch, "values") == [120, 95]       # floats -> plain numbers
    assert check_office.com_chart(ch, "categories") == ["North", "South"]
    assert check_office.com_chart(ch, "labels") == ["percent"]
    ch.ChartType = 999
    assert check_office.com_chart(ch, "type") == "999"


def test_word_and_powerpoint_charts_gated_on_office_2007():
    """No chart object model in Word/PowerPoint 2007 before SP2: not judged
    there (with the error), an error from 2010 on."""
    def broken():
        raise AttributeError("InlineShape has no attribute Chart")

    got = check_office._chart_om(12, broken)
    assert isinstance(got, check_office.NotJudged) and "SP2" in got and "AttributeError" in got
    with pytest.raises(AttributeError):
        check_office._chart_om(14, broken)
    assert check_office._chart_om(12, lambda: "column") == "column"


def test_powerpoint_chart_shapes_and_slide_image(tmp_path):
    exported = []

    class Shape:
        def __init__(self, chart, has=True):
            self.Chart, self.Type = chart, 3 if chart else 1
            if has is not None:
                self.HasChart = bool(chart)

    shapes = [Shape(None), Shape(_FakeChart())]

    class Slide:
        Shapes = type("Shapes", (), {"Count": 2, "__call__": lambda self, i: shapes[i - 1]})()

        def Export(self, path, fmt, w, h):
            exported.append((path, fmt, w, h))

    setup = type("PS", (), {"SlideWidth": 960.0, "SlideHeight": 540.0})()
    pres = type("Pres", (), {"Slides": staticmethod(lambda i: Slide()), "PageSetup": setup,
                             "Application": type("App", (), {"Version": "12.0"})()})()
    a = check_office.PowerPointAdapter(pres)
    assert a.chart_count(1) == 1
    assert a.chart(1, 1, "type") == "column"
    assert any("HasChart=True" in t for t in a.trace)
    a.screens = (str(tmp_path), "pc-ppt-08")
    note = a.slide_image(2)
    assert isinstance(note, check_office.NotJudged) and "pc-ppt-08-slide2.png" in note
    assert exported[0][1:] == ("PNG", 1280, 720)
    setup.SlideWidth, setup.SlideHeight = 720.0, 540.0     # a 4:3 deck keeps its shape
    a.slide_image(1)
    assert exported[1][1:] == ("PNG", 1280, 960)


def test_auto_shape_table_matches_oox():
    from oox.slides import SHAPES

    assert check_office.AUTO_SHAPES == SHAPES


def test_powerpoint_depth_probes_with_a_fake_powerpoint():
    def coll(items):
        return type("Coll", (), {"Count": len(items),
                                 "__call__": lambda self, i: items[i - 1]})()

    def ph(kind, text="", contained=None):
        fmt = type("PF", (), {"Type": kind})()
        if contained is not None:
            fmt.ContainedType = contained
        tf = type("TF", (), {"TextRange": type("TR", (), {"Text": text})()})()
        return type("Sh", (), {"Type": 14, "PlaceholderFormat": fmt, "TextFrame": tf,
                               "HasTable": False})()

    cell = type("Cell", (), {"Shape": type("S", (), {"TextFrame": type("TF", (), {
        "TextRange": type("TR", (), {"Text": "120"})()})()})()})()
    table = type("T", (), {"Rows": type("R", (), {"Count": 4})(),
                           "Columns": type("C", (), {"Count": 3})(),
                           "Cell": staticmethod(lambda r, c: cell)})()
    table_shape = type("Sh", (), {"Type": 19, "HasTable": True, "Table": table})()
    auto = [type("Sh", (), {"Type": 1, "AutoShapeType": t, "HasTable": False})()
            for t in (5, 52, 9)]
    slides = {
        1: [ph(1, "Title"), table_shape, *auto],
        2: [ph(1, "Picture"), ph(18, contained=13)],
        3: [ph(18)],                                    # 2007: no ContainedType
    }
    notes = {1: [ph(1), ph(2, "Line 1\rLine 2")]}

    class Slide:
        def __init__(self, i):
            self.Shapes = coll(slides[i])
            self.NotesPage = type("NP", (), {"Shapes": coll(notes.get(i, []))})()
            self.CustomLayout = type("L", (), {"Name": "Title Only"})()

    layouts = coll([type("L", (), {"Name": n})() for n in ("Title Slide", "Blank")])
    pres = type("Pres", (), {
        "Slides": staticmethod(lambda i: Slide(i)),
        "SlideMaster": type("M", (), {"CustomLayouts": layouts})(),
        "Application": type("App", (), {"Version": "12.0"})()})()
    a = check_office.PowerPointAdapter(pres)
    assert a.layouts() == ["Title Slide", "Blank"]
    assert a.slide_layout(1) == "Title Only"
    assert (a.table(1, "rows"), a.table(1, "cols"), a.table(1, "cell", 2, 2)) == (4, 3, "120")
    assert a.notes(1) == "Line 1\nLine 2" and a.notes(2) is None
    assert a.shape_types(1) == ["roundRect", "chevron", "ellipse"]
    assert a.picture_placeholder(2) is True
    got = a.picture_placeholder(3)
    assert isinstance(got, check_office.NotJudged) and "Office 2007" in got


def test_picture_placeholder_probe_falls_back_to_crop_on_2007():
    fmt = type("PF", (), {"Type": 18})()
    pic = type("PicF", (), {"CropLeft": 12.5, "CropRight": 12.5, "CropTop": 0.0,
                            "CropBottom": 0.0})()
    shape = type("Sh", (), {"Type": 14, "PlaceholderFormat": fmt, "PictureFormat": pic})()
    assert check_office._contains_picture(shape) is True


def test_report_lists_not_judged_probes(tmp_path):
    manifest = generate(tmp_path)
    (tmp_path / "results-pc-w.json").write_text(json.dumps({
        "machine": {"host": "pc"}, "apps": {"word": {"version": "12.0", "office": "Office 2007"}},
        "results": [{"id": "word-23-charts", "status": "ok", "dialogs": [], "probes": [
            {"kind": "word.chart", "index": 1, "attr": "type", "ok": True, "got": None,
             "note": "not judged: not in Office 2007's object model"}]}]}), encoding="utf-8")
    text = report.render(manifest, report.load_sources(str(tmp_path)))
    assert "## Not judged" in text
    assert "`word.chart` index=1 attr=type: not in Office 2007's object model" in text


def test_links_probe_reads_action_settings_not_slide_hyperlinks():
    """PowerPoint 2007 (ppt-10, 2026-11-15): Slide.Hyperlinks gave two entries
    and missed the text runs' links; the probe reads ActionSettings instead."""
    def coll(items):
        return type("Coll", (), {"Count": len(items),
                                 "__call__": lambda self, i: items[i - 1],
                                 "Item": lambda self, i: items[i - 1]})()

    def settings(action, address="", sub=""):
        link = type("H", (), {"Address": address, "SubAddress": sub})()
        return type("AS", (), {"Action": action, "Hyperlink": link})()

    def shape(click, runs=None):
        runs = runs or []
        rng = type("TR", (), {
            "Runs": lambda self, *a: (type("R", (), {"Count": len(runs)})() if not a
                                      else type("Run", (), {"ActionSettings": (
                                          lambda self2, k, r=runs[a[0] - 1]: r)})())})()
        return type("Sh", (), {"ActionSettings": lambda self, k: click,
                               "HasTextFrame": bool(runs),
                               "TextFrame": type("TF", (), {"HasText": bool(runs),
                                                            "TextRange": rng})()})()

    none = settings(0)
    docs = settings(7, "https://example.org/docs")
    shapes = [shape(none),                                      # title
              shape(none, [none, docs, docs, none, settings(7, "mailto:t@x.org"), none]),
              shape(settings(1)),                               # next slide: not a link
              shape(settings(7, "", "258,3,Summary"))]
    slide = type("Slide", (), {"Shapes": coll(shapes),
                               "Hyperlinks": coll([settings(7, "", "-1").Hyperlink])})()
    pres = type("Pres", (), {"Slides": staticmethod(lambda i: slide),
                             "Application": type("App", (), {"Version": "12.0"})()})()
    a = check_office.PowerPointAdapter(pres)
    assert a.links(1) == ["https://example.org/docs", "mailto:t@x.org", "slide:3"]
    assert "Slide.Hyperlinks" in a.trace[0]


def test_y2_title_probe_uses_the_axes_method():
    """Excel 2007: HasAxis(2, 2) failed through dynamic COM (E_FAIL), and
    Axes().Item(i) takes (Type, AxisGroup); Axes(2, 2) is a method call."""
    def axis(kind, group, title=None):
        return type("Ax", (), {"Type": kind, "AxisGroup": group, "HasTitle": title is not None,
                               "AxisTitle": type("T", (), {"Text": title})()})()

    items = [axis(1, 1), axis(2, 1, "Units"), axis(1, 2), axis(2, 2, "EUR")]

    class Chart:
        direct = True

        def Axes(self, *args):
            if not args:
                return iter(items)
            if not self.direct:
                raise RuntimeError("E_FAIL")
            return next(a for a in items if (a.Type, a.AxisGroup) == args)

        @property
        def HasAxis(self):
            raise RuntimeError("E_FAIL: fetched without arguments")

    chart = Chart()
    assert check_office.com_chart(chart, "y2_title") == "EUR"
    chart.direct = False                                    # falls back to the enumeration
    assert check_office.com_chart(chart, "y2_title") == "EUR"
    items.pop()                                             # no secondary axis at all
    assert check_office.com_chart(chart, "y2_title") is None
