"""Run the Office check probes against LibreOffice (Linux, macOS, Windows).

    python -m tools.officecheck.check_libreoffice CORPUS_DIR [--soffice PATH]

LibreOffice converts each file to the same format (and Word files also to
PDF for page counts); the probes are then evaluated on what LibreOffice wrote
back, i.e. on what it understood.  A failed conversion counts as an error.
There are no dialogs in headless mode, so "prompt" never occurs here.

Unlike check_office.py this uses oox to read the converted files.  Results use
the same format, so report.py puts LibreOffice next to the Office columns.
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import platform
import re
import shutil
import socket
import subprocess
import sys
import tempfile
from pathlib import Path

from oox import open_document, teletype
from oox.grammar import URI
from oox.ns import w

from .check_office import NotJudged, decide_status, evaluate_probe, sha256

_EXT = {"word": "docx", "excel": "xlsx", "powerpoint": "pptx"}
_MAC = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


def find_soffice(explicit: str | None) -> str:
    for cand in (explicit, shutil.which("soffice"), shutil.which("libreoffice"), _MAC):
        if cand and Path(cand).exists():
            return cand
    sys.exit("LibreOffice not found; pass --soffice PATH")


def convert(soffice: str, files: list[Path], fmt: str, outdir: Path, profile: Path) -> None:
    subprocess.run([soffice, f"-env:UserInstallation={profile.as_uri()}", "--headless",
                    "--convert-to", fmt, "--outdir", str(outdir), *map(str, files)],
                   capture_output=True, timeout=600, check=False)  # failures show as missing output


def pdf_pages(path: Path) -> int:
    return len(re.findall(rb"/Type\s*/Page[^s]", path.read_bytes()))


_PROPS = {"Title": "title", "Author": "creator", "Subject": "subject", "Keywords": "keywords",
          "Comments": "description", "Last Author": "last_modified_by", "Category": "category"}
_EXT_PROPS = {"Company": "company", "Manager": "manager", "Template": "template"}


def _property(doc, name):
    if name in _EXT_PROPS:
        return getattr(doc.extended_properties, _EXT_PROPS[name])
    return getattr(doc.properties, _PROPS[name])


def _continues(tc):
    tcpr = tc.first("w:tcPr")
    vm = tcpr.first("w:vMerge") if tcpr is not None else None
    return vm is not None and vm.get(f"{{{URI['w']}}}val") != "restart"


def _chart_attr(info, attr, series=1):
    """A Chart record -> the value a chart probe asks for (same terms as check_office.com_chart)."""
    if attr == "type":
        return info.kind
    if attr in ("title", "legend", "x_title", "y_title", "y2_title", "labels"):
        return getattr(info, attr)
    if attr == "series":
        return [s.name for s in info.series]
    ser = info.series[series - 1]
    if attr == "series_type":
        return ser.kind
    if attr == "axis":
        return ser.axis
    if attr == "values":
        return list(ser.values)
    if attr == "categories":
        return list(ser.categories) if ser.categories is not None else None
    raise ValueError("unknown chart attribute %r" % attr)


def _unprotect(check, unprotect, wrong, password, after):
    """The COM probe's answer from oox's password checks."""
    refused = None
    if wrong is not None:
        if check(wrong):
            unprotect()
            return [False, None, after()]
        refused = True
    accepted = password is None or check(password)
    if accepted:
        unprotect()
    return [refused, accepted, after()]


class LoWord:
    source: Path | None = None

    def __init__(self, doc, pdf: Path | None):
        self.doc, self.pdf = doc, pdf

    def chart_count(self):
        return len(self.doc.charts())

    def chart(self, index, attr, series=1):
        return _chart_attr(self.doc.charts()[index - 1], attr, series)

    def text(self):
        return "\n".join(teletype.extractText(q) for q in self.doc.document.iter(f"{{{URI['w']}}}p"))

    def tables(self):
        return len(self.doc.getElementsByType(w.Tbl))

    def inline_shapes(self):
        return sum(1 for e in self.doc.document.iter("{*}inline"))

    def property(self, name):
        return _property(self.doc, name)

    def list_string(self, paragraph):
        ps = list(self.doc.document.iter(f"{{{URI['w']}}}p"))
        labels = {id(item.paragraph): item.raw_label for item in self.doc.list_labels()}
        return labels.get(id(ps[paragraph - 1]), "")

    def outline_level(self, paragraph):
        """Direct w:outlineLvl, else the paragraph style's (following basedOn); 10 = body."""
        ps = list(self.doc.document.iter(f"{{{URI['w']}}}p"))[paragraph - 1]
        root = self.doc.styles
        styles = {} if root is None else {
            s.get(f"{{{URI['w']}}}styleId"): s for s in root.iterchildren(f"{{{URI['w']}}}style")}

        def level_of(el, depth=0):
            ppr = el.first("w:pPr")
            lvl = ppr.first("w:outlineLvl") if ppr is not None else None
            if lvl is not None:
                return int(lvl.get(f"{{{URI['w']}}}val")) + 1
            ref = (ppr.first("w:pStyle") if ppr is not None and el.tag_name == "w:p"
                   else el.first("w:basedOn"))
            target = styles.get(ref.get(f"{{{URI['w']}}}val")) if ref is not None else None
            return level_of(target, depth + 1) if target is not None and depth < 10 else 10
        return level_of(ps)

    def count(self, what):
        from oox.fields import field_codes

        d = self.doc
        return {"sections": lambda: len(d.sections), "hyperlinks": lambda: len(d.hyperlinks()),
                "bookmarks": lambda: len(d.bookmarks()), "footnotes": lambda: len(d.footnotes()),
                "endnotes": lambda: len(d.endnotes()), "comments": lambda: len(d.comments()),
                "revisions": lambda: len([r for r in d.revisions() if r.kind in
                                          ("insert", "delete")]),
                "controls": lambda: len(d.content_controls()),
                "toc": lambda: field_codes(d.body).count("TOC")}[what]()

    def orientation(self, section):
        return self.doc.sections[section - 1].orientation

    def _hf(self, part, section, kind):
        root = self.doc.sections[section - 1].shown(part, kind)
        if root is None:
            raise LookupError(f"no {kind} {part} in section {section}")
        return root

    def header_footer_text(self, part, section, kind):
        return teletype.extractText(self._hf(part, section, kind))

    def fields(self, part, section, kind):
        from oox.fields import field_codes

        return field_codes(self._hf(part, section, kind))

    def hyperlink(self, index):
        link = self.doc.hyperlinks()[index - 1]
        return link.target if link.target is not None else "#" + str(link.anchor)

    def note_text(self, kind, index):
        found = self.doc.footnotes() if kind == "footnote" else self.doc.endnotes()
        return found[index - 1].text

    def comment(self, index):
        c = self.doc.comments()[index - 1]
        return c.author, c.text

    def custom_property(self, name):
        return self.doc.custom_properties.get(name)

    def _paragraph(self, paragraph):
        return list(self.doc.document.iter(f"{{{URI['w']}}}p"))[paragraph - 1]

    def font(self, paragraph, char, attr):
        p = self._paragraph(paragraph)
        pos, target = 0, None
        for r in p.iter(f"{{{URI['w']}}}r"):
            n = len(teletype.extractText(r))
            if pos < char <= pos + n:
                target = r
                break
            pos += n
        fmt = self.doc.effective_format(p, target)
        if attr == "underline":
            return fmt.underline is not None
        if attr in ("subscript", "superscript", "small_caps", "strike", "bold", "italic"):
            return bool(getattr(fmt, attr))
        if attr == "size":
            return round(fmt.size, 1)
        return getattr(fmt, attr)

    def paragraph_format(self, paragraph, attr):
        fmt = self.doc.formatting().paragraph(self._paragraph(paragraph))
        value = getattr(fmt, attr)
        return round(value, 1) if isinstance(value, float) else value

    def field_results(self, codes, update=False):
        """Cached results as LibreOffice wrote them back.  ``update`` is not
        re-run here: LibreOffice drops some switches on export (``\\* Upper`` on
        DOCPROPERTY), so recomputing would test LibreOffice's export rather
        than oox.  oox's own recompute is covered by the unit tests; the
        Word side compares it with Word's."""
        return [f.result for f in self.doc.fields() if f.name in codes]

    def control(self, tag):
        info = next(c for c in self.doc.content_controls() if c.tag == tag)
        text = info.value
        if info.type == "checkbox":
            text = "\u2612" if info.value else "\u2610"
        elif text is None:  # placeholder showing: the placeholder text
            sdt = next(s for s in self.doc.document.iter(f"{{{URI['w']}}}sdt")
                       if s.find(f".//{{{URI['w']}}}tag").get(f"{{{URI['w']}}}val") == tag)
            text = teletype.extractText(sdt.first("w:sdtContent"))
        return {"text": text, "type": info.type, "placeholder": info.placeholder}

    def table(self, table, what):
        t = self.doc.tables()[table - 1]
        if what == "rows":
            return t.word_row_count
        # cells Word counts: one per w:tc, minus vertical-merge continuations
        return sum(1 for tc in t.element.iter(f"{{{URI['w']}}}tc")
                   if tc.getparent().getparent() is t.element and not _continues(tc))

    def table_header(self, table):
        return self.doc.tables()[table - 1].header_rows >= 1

    def table_cell(self, table, row, col):
        """Like Word's Cell(row, col): the col-th w:tc of that row."""
        tr = self.doc.tables()[table - 1].rows[row - 1]
        tc = list(tr.iterchildren(f"{{{URI['w']}}}tc"))[col - 1]
        return teletype.extractText(tc)

    def shapes_count(self):
        return sum(1 for p in self.doc.pictures() if p.floating)

    def shape_wraps(self):
        return sorted(p.wrap for p in self.doc.pictures() if p.floating)

    def text_boxes(self):
        return [b.text for b in self.doc.text_boxes() if b.text]

    def _dropped_tracked_changes(self):
        """LibreOffice (24.2) doesn't import tracked-changes protection, so its
        re-save has none: check the original instead of failing."""
        if self.doc.protection() is not None or self.source is None:
            return None
        original = open_document(self.source).protection()
        if original and original.edit == "trackedChanges":
            return NotJudged("LibreOffice doesn't keep tracked-changes protection")
        return None

    def protection(self):
        info = self.doc.protection()
        return self._dropped_tracked_changes() or (
            info.edit if info and info.enforced else "none")

    def unprotect(self, wrong, password):
        return self._dropped_tracked_changes() or _unprotect(
            self.doc.check_protection_password, self.doc.unprotect, wrong, password,
            self.protection)

    def revision(self, index):
        r = [x for x in self.doc.revisions() if x.kind in ("insert", "delete")][index - 1]
        return r.kind, r.author, r.text

    def text_after(self, action):
        if action == "accept":
            self.doc.accept_all()
        else:
            self.doc.reject_all()
        return self.text()

    def pages(self):
        if self.pdf is None or not self.pdf.exists():
            raise RuntimeError("no PDF produced")
        return pdf_pages(self.pdf)


class LoExcel:
    source: Path | None = None

    def custom_property(self, name):
        return self.book.custom_properties.get(name)

    def text(self, sheet, ref):
        return self.book.display_text(sheet, ref)

    def font(self, sheet, ref, attr):
        fmt = self.book.cell_format(sheet, ref)
        return round(fmt.size, 1) if attr == "size" else getattr(fmt, attr)

    def fill(self, sheet, ref):
        return self.book.cell_format(sheet, ref).fill

    def align(self, sheet, ref):
        return self.book.cell_format(sheet, ref).align or "general"

    def number_format(self, sheet, ref):
        """LibreOffice escapes literal punctuation on export (yyyy\\-mm\\-dd):
        the same code, so compare it unescaped."""
        import re

        code = self.book.cell_format(sheet, ref).number_format or "General"
        return re.sub(r"\\([^A-Za-z0-9])", r"\1", code)

    def column_width(self, sheet, col):
        """LibreOffice re-saves widths in its own units: compare whole characters."""
        width = self.book.column_width(sheet, col)
        return round(width) if width is not None else None

    def row_height(self, sheet, row):
        h = self.book.row_height(sheet, row)
        return round(h) if h is not None else None

    def merge_area(self, sheet, ref):
        from oox.spreadsheet import split_ref
        from oox.xlsheet import parse_range

        col, rw = split_ref(ref)
        for m in self.book.merged_cells(sheet):
            c1, r1, c2, r2 = parse_range(m)
            if c1 <= col <= c2 and r1 <= rw <= r2:
                return m
        return ref

    def frozen(self, sheet):
        return self.book.frozen_panes(sheet)

    def autofilter(self, sheet):
        return self.book.auto_filter_range(sheet)

    def defined_name(self, name):
        return self.book.defined_names().get(name)

    def _covers(self, ranges, ref):
        from oox.spreadsheet import split_ref
        from oox.xlsheet import parse_range

        col, rw = split_ref(ref)
        for part in ranges.split():
            c1, r1, c2, r2 = parse_range(part)
            if c1 <= col <= c2 and r1 <= rw <= r2:
                return True
        return False

    def screenshot(self, sheet, ref):
        return NotJudged("no screenshot from LibreOffice's re-save")

    def chart_count(self, sheet):
        return len(self.book.charts(sheet))

    def chart(self, sheet, index, attr, series=1):
        return _chart_attr(self.book.charts(sheet)[index - 1], attr, series)

    def chart_image(self, sheet, index):
        return NotJudged("no chart image from LibreOffice's re-save")

    def cf_count(self, sheet):
        return len(self.book.conditional_formats(sheet))

    def cf_types(self, sheet, ref):
        alias = {"uniqueValues": "duplicateValues"}
        return sorted(str(alias.get(r.type, r.type)) for r in self.book.conditional_formats(sheet)
                      if self._covers(r.ref, ref))

    def tabs(self):
        return self.book.tab_names

    def _dropped_workbook_protection(self):
        """LibreOffice (24.2) doesn't import workbook structure protection: its
        re-save has an empty x:workbookProtection.  Check the original instead."""
        info = self.book.workbook_protection()
        if (info and (info.structure or info.windows or info.password)) or self.source is None:
            return None
        original = open_document(self.source).workbook_protection()
        if original and (original.structure or original.windows or original.password):
            return NotJudged("LibreOffice doesn't keep workbook structure protection")
        return None

    def workbook_protection(self, attr):
        return self._dropped_workbook_protection() or bool(
            getattr(self.book.workbook_protection(), attr, False))

    def unprotect_workbook(self, wrong, password):
        return self._dropped_workbook_protection() or _unprotect(
            self.book.check_workbook_password, self.book.unprotect_workbook, wrong, password,
            lambda: self.workbook_protection("structure"))

    def chart_sheet(self, name, attr, series=1):
        return _chart_attr(self.book.charts(name)[0], attr, series)

    def table_names(self, sheet):
        return sorted(t.name for t in self.book.tables(sheet))

    def table(self, sheet, name, attr):


        if attr == "style":
            return NotJudged("LibreOffice doesn't keep table styles")
        t = next(t for t in self.book.tables(sheet) if t.name == name)
        return {"ref": t.ref, "totals": t.totals_row, "columns": t.columns,
                "style": t.style}[attr]

    def comment_count(self, sheet):
        return len(self.book.comments(sheet))

    def comment(self, sheet, ref):


        found = self.book.comments(sheet).get(ref)
        text = found.text if found is not None else None
        return NotJudged("LibreOffice writes comment authors as 'Unknown Author'"), text

    def hyperlink(self, sheet, ref):
        found = self.book.hyperlinks(sheet).get(ref)
        return found.target if found is not None else None

    def page(self, sheet, attr):
        info = self.book.page_setup(sheet)
        if attr == "print_area":
            return info.print_area or ""
        if attr == "print_titles":
            return ",".join(v for v in (info.repeat_cols, info.repeat_rows) if v)
        if attr.startswith(("header_", "footer_")):
            kind, _, side = attr.partition("_")
            text = (getattr(info, kind) or {}).get(side, "")
            # back to Excel's codes, which COM returns
            text = text.replace("&", "&&")
            for name, code in (("{page}", "&P"), ("{pages}", "&N"), ("{date}", "&D"),
                               ("{sheet}", "&A"), ("{file}", "&F"), ("{time}", "&T")):
                text = text.replace(name, code)
            return text
        return getattr(info, attr)

    def sheet_option(self, sheet, attr):
        if attr == "tab_color":


            return NotJudged("LibreOffice drops tab colours on export")
        return getattr(self.book.sheet_options(sheet), attr)

    def validation(self, sheet, ref):
        from oox.spreadsheet import split_ref
        from oox.xlsheet import parse_range

        col, rw = split_ref(ref)
        for v in self.book.validations(sheet):
            for part in v.ref.split():
                c1, r1, c2, r2 = parse_range(part)
                if c1 <= col <= c2 and r1 <= rw <= r2:
                    return ",".join(v.values) if v.values is not None else v.source
        return None

    def __init__(self, book):
        self.book = book

    def sheets(self):
        return self.book.sheet_names

    def cell(self, sheet, ref):
        return self.book.value(sheet, ref)

    def property(self, name):
        return _property(self.book, name)

    def shapes(self, sheet):
        ws = self.book.sheet(sheet)
        parts = self.book.related_parts(self.book.part_of(ws), "DrawingsPart")
        if not parts:
            return 0
        return sum(1 for e in parts[0].element if e.tag.rpartition("}")[2].endswith("Anchor"))


class LoPowerPoint:
    def custom_property(self, name):
        return self.deck.custom_properties.get(name)

    def __init__(self, deck):
        self.deck = deck

    def chart_count(self, slide):
        return len(self.deck.charts(self.deck.slides[slide - 1]))

    def chart(self, slide, index, attr, series=1):
        return _chart_attr(self.deck.charts(self.deck.slides[slide - 1])[index - 1], attr, series)

    def slide_image(self, slide):
        return NotJudged("no slide image from LibreOffice's re-save")

    def _slide(self, slide):
        return self.deck.slides[slide - 1]

    def layouts(self):
        return self.deck.layout_names

    def slide_layout(self, slide):
        return self.deck.slide_layout(self._slide(slide))

    def table(self, slide, attr, row=None, col=None):
        rows = self.deck.tables(self._slide(slide))[0].rows
        if attr == "rows":
            return len(rows)
        if attr == "cols":
            return max(len(r) for r in rows)
        if attr == "cell":
            if row is None or col is None:
                raise ValueError("a cell probe needs row and col")
            return rows[row - 1][col - 1]
        raise ValueError("unknown table attribute %r" % attr)

    def sections(self):
        found = self.deck.sections()
        if not found:
            return NotJudged("LibreOffice doesn't keep slide sections")
        return [[s.name, s.slides[0] if s.slides else 0, len(s.slides)] for s in found]

    def links(self, slide):
        out = []
        for lk in self.deck.links(self._slide(slide)):
            if lk.kind == "url":
                out.append(lk.target)
            elif lk.kind == "slide":
                out.append("slide:%s" % lk.anchor)
        return out

    def action(self, slide, shape):
        for lk in self.deck.links(self._slide(slide)):
            if lk.shape == shape and lk.text is None:
                return lk.target if lk.kind == "jump" else "hyperlink"
        return "none"

    def notes(self, slide):
        return self.deck.notes(self._slide(slide))

    def shape_types(self, slide):
        return [s.geometry for s in self.deck.shapes(self._slide(slide)) if s.kind == "shape"]

    def picture_placeholder(self, slide):
        kinds = [s.kind for s in self.deck.shapes(self._slide(slide))]
        if "picture" in kinds:
            return NotJudged("LibreOffice saves a filled placeholder as a plain picture")
        return False

    def slides(self):
        return len(self.deck.slides)

    def property(self, name):
        return _property(self.deck, name)

    def _shapes(self, slide):
        tree = self.deck.slides[slide - 1].first("p:cSld").first("p:spTree")
        return [e for e in tree if e.tag.rpartition("}")[2] not in ("nvGrpSpPr", "grpSpPr", "extLst")]

    def shapes(self, slide):
        return len(self._shapes(slide))

    def slide_text(self, slide):
        out = []
        for sh in self._shapes(slide):
            for para in sh.iter("{*}p"):
                if para.tag.endswith("}p") and para.getparent().tag.endswith("}txBody"):
                    out.append("".join(t.text or "" for t in para.iter("{*}t")))
        return "\n".join(out)


def lo_version(soffice: str) -> str:
    out = subprocess.run([soffice, "--version"], capture_output=True, text=True, timeout=120,
                         check=False).stdout
    m = re.search(r"(\d+\.\d+(\.\d+)*)", out)
    return m.group(1) if m else "?"


def run(corpus: str | Path, soffice: str) -> dict:
    corpus = Path(corpus)
    manifest = json.loads((corpus / "manifest.json").read_text(encoding="utf-8"))
    version = lo_version(soffice)
    results = []
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        profile = tmp / "profile"  # isolated LibreOffice profile: no clashes with a running one
        # controls test Office's dialogs; headless LibreOffice has none
        cases = [c for c in manifest["cases"] if "skipped" not in c and not c.get("control")]
        for app, ext in _EXT.items():
            files = [corpus / c["file"] for c in cases if c["app"] == app]
            if files:
                convert(soffice, files, ext, tmp / app, profile)
                if app == "word":
                    convert(soffice, files, "pdf", tmp / "pdf", profile)
        for case in cases:
            src = corpus / case["file"]
            res = {"id": case["id"], "app": case["app"], "file": case["file"]}
            if case.get("sha256") and sha256(src) != case["sha256"]:
                res.update(status="error", error="file does not match manifest sha256")
                results.append(res)
                continue
            out = tmp / case["app"] / src.name
            error, probes = None, []
            try:
                if not out.exists():
                    raise RuntimeError("LibreOffice could not convert the file")
                doc = open_document(out)
                if case["app"] == "word":
                    adapter = LoWord(doc, tmp / "pdf" / (src.stem + ".pdf"))
                elif case["app"] == "excel":
                    adapter = LoExcel(doc)
                else:
                    adapter = LoPowerPoint(doc)
                adapter.source = src  # the original, for what LibreOffice drops
                probes = [evaluate_probe(pr, adapter) for pr in case.get("probes", [])]
            except Exception as e:
                error = repr(e)
            res.update(status=decide_status(error, [], probes, False), dialogs=[], probes=probes)
            if error:
                res["error"] = error
            results.append(res)
    office = f"LibreOffice {version}"
    return {
        "kind": "oox-officecheck-results", "format": 1,
        "created": datetime.datetime.now().isoformat(timespec="seconds"),
        "machine": {"host": socket.gethostname(), "os": platform.platform(),
                    "python": platform.python_version()},
        "apps": {a: {"version": version, "office": office} for a in _EXT},
        "manifest": {"oox_version": manifest.get("oox_version"),
                     "grammar_source": manifest.get("grammar_source")},
        "results": results,
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    ap.add_argument("corpus")
    ap.add_argument("--soffice")
    ap.add_argument("--out")
    args = ap.parse_args(argv)
    doc = run(args.corpus, find_soffice(args.soffice))
    ver = doc["apps"]["word"]["version"].split(".")[0]
    out = args.out or os.path.join(args.corpus, f"results-{socket.gethostname()}-libreoffice{ver}.json")
    Path(out).write_text(json.dumps(doc, indent=2, ensure_ascii=False), encoding="utf-8")
    bad = [r for r in doc["results"] if r["status"] != "ok"]
    for r in bad:
        print(f"{r['id']}: {r['status']} {r.get('error', '')}",
              [pr for pr in r.get("probes", []) if not pr.get("ok")])
    print(f"{len(doc['results']) - len(bad)} ok, {len(bad)} not ok -> {out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
