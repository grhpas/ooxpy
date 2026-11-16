"""Merge Office check results into one Markdown table.

    python report.py CORPUS_DIR                 # writes CORPUS_DIR/REPORT.md
    python report.py CORPUS_DIR --stdout

Reads every ``results-*.json`` (from check_office.py) and every filled-in
``manual-results*.csv`` (from a manual run; name it e.g.
``manual-results-office2007.csv``) in the corpus folder.  Python 3.8+, no
dependencies; uses check_office.py next to it (as in the kit) to tell an
unactivated Office's sign-in dialogs and refused calls from file problems,
also in results files written before the checker did that itself.
"""

import argparse
import csv
import glob
import importlib.util
import json
import os
import sys


def _classifier():
    """check_office.classify_environment, from the package or from the file
    next to this one (the kit; also under ``python -I``), or None."""
    try:
        from .check_office import classify_environment
        return classify_environment
    except ImportError:
        pass
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "check_office.py")
    try:
        spec = importlib.util.spec_from_file_location("_oox_check_office", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module.classify_environment
    except Exception:
        return None


classify_environment = _classifier()


ICON = {"ok": "✅", "prompt": "⚠️", "error": "❌", "mismatch": "❌", "timeout": "⏱️",
        "unactivated": "🔒", "": "·"}


def cell(case, result):
    """Icon for one case on one machine.  Controls pass when Office complained."""
    status = (result or {}).get("status", "")
    if not status:
        return "·"
    if case.get("observe"):
        return "👁 " + status  # recorded, not judged
    if case.get("control"):
        if status not in case.get("expect_status", []):
            return "❌ not detected (%s)" % status
        return "✅ detected (silent repair, content lost)" if status == "mismatch" else "✅ detected"
    return ICON.get(status, "?")


def _run_line(doc):
    """One line per results file: interpreter, pywin32, COM binding."""
    m = doc.get("machine", {})
    bindings = sorted({r["binding"] for r in doc.get("results", []) if r.get("binding")})
    bits = (", %s-bit" % m["bits"]) if m.get("bits") else ""
    exe = (", `%s`" % m["executable"]) if m.get("executable") else ""
    return "- `%s`: Python %s%s%s; pywin32 %s; COM binding %s" % (
        doc["_path"], m.get("python", "?"), bits, exe, m.get("pywin32") or "?",
        "/".join(bindings) or "?")


def load_sources(corpus):
    """One column per machine + Office version.

    Several runs on the same machine and version (a full run, then a rerun of
    one app or one case) merge into one column; for each case the newest run
    wins.
    """
    sources = []
    by_label = {}
    docs = []
    for path in sorted(glob.glob(os.path.join(corpus, "results-*.json"))):
        with open(path, encoding="utf-8") as f:
            doc = json.load(f)
        doc["_path"] = os.path.basename(path)
        docs.append(doc)
    for doc in sorted(docs, key=lambda d: d.get("created", "")):  # oldest first
        offices = sorted({a["office"] for a in doc.get("apps", {}).values()})
        label = "%s (%s)" % (doc["machine"]["host"], ", ".join(offices) or "?")
        src = by_label.get(label)
        if src is None:
            src = by_label[label] = {"label": label, "kind": "scripted", "paths": [],
                                     "results": {}}
            sources.append(src)
        src["paths"].append(doc["_path"])
        if classify_environment is not None:
            for r in doc["results"]:
                classify_environment(r)
        src["results"].update({r["id"]: r for r in doc["results"]})  # newer wins
        src.setdefault("runs", []).append(_run_line(doc))
    for path in sorted(glob.glob(os.path.join(corpus, "manual-results*.csv"))):
        with open(path, newline="", encoding="utf-8") as f:
            rows = [r for r in csv.DictReader(f) if (r.get("status") or "").strip()]
        if not rows:
            continue  # the blank template
        stem = os.path.splitext(os.path.basename(path))[0]
        label = stem.replace("manual-results", "").strip("-_ ") or "manual"
        sources.append({"label": label + " (manual)", "kind": "manual",
                        "paths": [os.path.basename(path)],
                        "results": {r["id"]: {"id": r["id"], "status": r["status"].strip().lower(),
                                              "notes": r.get("notes", "")} for r in rows}})
    return sources


def render(manifest, sources):
    cases = [c for c in manifest["cases"] if "skipped" not in c]
    lines = ["# Office check report", "",
             "Corpus: oox %s, %s." % (manifest.get("oox_version"), manifest.get("grammar_source")),
             "", "✅ ok · ⚠️ prompt (repair/recovery dialog) · ❌ error or content mismatch · "
             "⏱️ timeout · 🔒 opened cleanly, but Office isn't activated and refused some "
             "probes · · not run", "",
             "*Control* cases are broken on purpose: ✅ means the checker noticed, either "
             "because Office complained or because it repaired silently and content went "
             "missing.", "",
             "*Observe* cases (👁) break the schema in one specific way and are not judged: "
             "they record how strict each Office version is.", ""]
    if not sources:
        lines.append("_No results yet: run check_office.py or fill in manual-results.csv._")
        return "\n".join(lines) + "\n"
    header = "| Case | " + " | ".join(s["label"] for s in sources) + " |"
    lines += [header, "|---|" + "---|" * len(sources)]
    for c in cases:
        cells = [cell(c, s["results"].get(c["id"])) for s in sources]
        tag = " *(control)*" if c.get("control") else " *(observe)*" if c.get("observe") else ""
        name = "`%s`%s" % (c["id"], tag)
        lines.append("| %s | %s |" % (name, " | ".join(cells)))
    for s in sources:
        results = [r for r in s["results"].values()]
        signin = sum(1 for r in results if r.get("environment_dialogs"))
        refused = sum(1 for r in results for p in r.get("probes", []) if p.get("refused"))
        if signin or refused:
            lines += ["", "**%s: Office isn't activated** (reduced functionality mode). "
                      "Its sign-in/activation dialog appeared in %d cases and %d probes were "
                      "refused for want of a licence; neither says anything about the files, "
                      "and neither counts. Activate Office and rerun the cases marked 🔒 for "
                      "a complete result." % (s["label"], signin, refused)]
    details = []
    for s in sources:
        for c in cases:
            r = s["results"].get(c["id"])
            traced = any(p.get("trace") for p in (r or {}).get("probes", [])) or \
                bool((r or {}).get("expected_dialogs"))
            if not r or (r.get("status") == "ok" and not c.get("control")
                         and not c.get("observe") and not traced):
                continue
            if c.get("control") and r.get("status") in c.get("expect_status", []) \
                    and r.get("status") != "mismatch" and not r.get("dialogs"):
                continue  # detected via an error: nothing more to show
            item = ["- **%s** on %s: %s" % (c["id"], s["label"], r.get("status"))]
            for d in r.get("dialogs", []):
                item.append("  - dialog `%s`: %s" % (d.get("title") or d.get("class"),
                                                    " / ".join(d.get("texts", []))[:400]))
            for d in r.get("expected_dialogs", []):
                item.append("  - expected dialog (provoked by a probe, not counted) `%s`: %s"
                            % (d.get("title") or d.get("class"),
                               " / ".join(d.get("texts", []))[:200]))
            if r.get("error"):
                item.append("  - error: `%s`" % r["error"][:300])
            for p in r.get("probes", []):
                if p.get("ok") and p.get("trace"):
                    where = "".join(" %s=%s" % (k, p[k]) for k in ("sheet", "ref") if k in p)
                    item.append("  - probe `%s`%s: ok, got `%r`%s" % (
                        p["kind"], where, p.get("got"),
                        (" (" + p["note"] + ")") if p.get("note") else ""))
                    for step in p["trace"][:40]:
                        item.append("    - " + step)
                if p.get("refused"):
                    item.append("  - probe `%s`: refused, Office isn't activated (not counted)"
                                % p["kind"])
                    continue
                if not p.get("ok"):
                    want = p.get("expect", p.get("value"))
                    where = "".join(" %s=%s" % (k, p[k]) for k in ("sheet", "ref") if k in p)
                    item.append("  - probe `%s`%s: expected `%r`, got `%r`%s" % (
                        p["kind"], where, want, p.get("got"),
                        (" (" + p["error"] + ")") if p.get("error") else ""))
                    for step in p.get("trace", [])[:40]:
                        item.append("    - " + step)
            if r.get("notes"):
                item.append("  - notes: " + r["notes"])
            details += item
    if details:
        lines += ["", "## Details", "", *details]
    shots = []
    for s in sources:
        for c in cases:
            for p in (s["results"].get(c["id"]) or {}).get("probes", []):
                if p.get("kind") in ("excel.screenshot", "excel.chart_image", "ppt.slide_image"):
                    what = p.get("note", "").replace("not judged: ", "") or p.get("error", "?")
                    shots.append("- `%s` on %s: %s" % (c["id"], s["label"], what))
    if shots:
        lines += ["", "## Screenshots", "",
                  "Saved in the corpus folder; send them along with this report.", "", *shots]
    skipped = []
    for s in sources:
        for c in cases:
            for p in (s["results"].get(c["id"]) or {}).get("probes", []):
                note = p.get("note", "")
                if p.get("ok") and note.startswith("not judged") and \
                        p.get("kind") not in ("excel.screenshot", "excel.chart_image",
                                              "ppt.slide_image"):
                    where = "".join(" %s=%s" % (k, p[k]) for k in
                                    ("sheet", "slide", "index", "attr", "ref") if k in p)
                    skipped.append("- `%s` on %s: `%s`%s: %s" % (
                        c["id"], s["label"], p["kind"], where, note[len("not judged: "):]))
    if skipped:
        lines += ["", "## Not judged", "",
                  "Counted as ok; the application can't answer these.", "", *skipped]
    runs = [line for s in sources for line in s.get("runs", [])]
    if runs:
        lines += ["", "## Runs", "", *runs]
    lines += ["", "Sources: " + ", ".join("`%s`" % p for s in sources for p in s["paths"])]
    return "\n".join(lines) + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    ap.add_argument("corpus")
    ap.add_argument("--stdout", action="store_true")
    args = ap.parse_args(argv)
    with open(os.path.join(args.corpus, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    text = render(manifest, load_sources(args.corpus))
    if args.stdout:
        sys.stdout.write(text)
    else:
        out = os.path.join(args.corpus, "REPORT.md")
        with open(out, "w", encoding="utf-8") as f:
            f.write(text)
        print("wrote " + out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
