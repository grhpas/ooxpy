"""Generate the API reference pages (docs/reference/*.md) from the docstrings.

    python -m tools.apidocs            # write docs/reference/
    python -m tools.apidocs --check    # exit 1 if the pages are out of date

The pages cover exactly the public API (tools/api_surface.py, docs/STABILITY.md).
Document methods are grouped by topic; a public method missing from TOPICS
fails ``--check`` (and the test suite), so new methods get a place.
"""

from __future__ import annotations

import argparse
import importlib
import inspect
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "reference"

COMMON = ["open", "save", "validate", "properties", "extended_properties", "custom_properties",
          "conformance", "add_picture_image", "add_image", "add_part", "related_parts", "relate",
          "part_of", "getElementsByType", "addPicture", "addPictureFromFile",
          "addPictureFromString", "MAIN_CONTENT_TYPES"]

TOPICS: dict[str, list[tuple[str, list[str]]]] = {
    "WordprocessingDocument": [
        ("Creating and structure", ["new", "body", "text", "document", "settings", "styles"]),
        ("Text and styles", ["add_heading", "add_paragraph", "add_run", "format_paragraph",
                             "add_style", "style_id", "ensure_style", "effective_format",
                             "outline_level", "add_page_break"]),
        ("Lists", ["new_list", "list_labels", "numbering"]),
        ("Sections, headers and footers", ["sections", "add_section", "even_and_odd_headers"]),
        ("Links, bookmarks, notes and comments", [
            "add_hyperlink", "hyperlinks", "add_bookmark", "bookmarks", "add_footnote",
            "footnotes", "add_endnote", "endnotes", "add_comment", "comments"]),
        ("Tables", ["add_table", "tables"]),
        ("Pictures, charts and text boxes", [
            "add_picture", "add_floating_picture", "pictures", "add_chart", "charts",
            "add_text_box", "text_boxes", "set_text_box_text"]),
        ("Fields, captions, contents and templates", [
            "add_field", "fields", "update_fields", "add_caption", "add_cross_reference",
            "add_toc", "replace_text"]),
        ("Tracked changes", ["suggest_insert", "suggest_delete", "suggest_replace",
                             "track_changes", "revisions", "accept_all", "reject_all"]),
        ("Forms (content controls)", ["add_text_control", "add_dropdown", "add_date_control",
                                      "add_checkbox", "content_controls", "fill_controls"]),
        ("Protection", ["protect", "protection", "check_protection_password", "unprotect"]),
    ],
    "SpreadsheetDocument": [
        ("Workbook and sheets", ["new", "workbook", "spreadsheet", "add_sheet", "sheet",
                                 "sheet_names", "tab_names", "chart_sheet_names", "worksheets",
                                 "sheet_data", "cell_element"]),
        ("Values", ["set_value", "append_row", "value", "values", "cells", "display_text",
                    "share_strings"]),
        ("Cell styles", ["cell_style", "style_range", "cell_format"]),
        ("Layout", ["set_column_width", "column_width", "autofit_columns", "set_row_height",
                    "row_height", "merge_cells", "unmerge_cells", "merged_cells",
                    "freeze_panes", "frozen_panes", "auto_filter", "auto_filter_range",
                    "define_name", "defined_names", "add_validation", "validations"]),
        ("Conditional formatting", ["conditional_format", "conditional_formats"]),
        ("Tables", ["add_table", "tables"]),
        ("Comments and hyperlinks", ["add_comment", "comments", "add_hyperlink", "hyperlinks"]),
        ("Printing and sheet options", ["page_setup", "page_setup_info", "sheet_options",
                                        "sheet_options_info"]),
        ("Protection", ["protect_sheet", "unprotect_sheet", "protect_workbook",
                        "workbook_protection", "check_workbook_password", "unprotect_workbook"]),
        ("Pictures and charts", ["add_picture", "add_chart", "add_chart_sheet", "chart_sheet",
                                 "charts"]),
    ],
    "PresentationDocument": [
        ("Deck, slides and layouts", ["new", "presentation", "slides", "add_slide", "masters",
                                      "layouts", "layout_names", "layout", "add_layout",
                                      "slide_layout"]),
        ("Placeholders", ["placeholders", "placeholder", "set_placeholder_text",
                          "set_placeholder_picture"]),
        ("Text boxes, shapes and pictures", ["add_textbox", "add_shape", "add_picture",
                                             "shapes"]),
        ("Tables", ["add_table", "tables", "table"]),
        ("Charts", ["add_chart", "charts"]),
        ("Notes, links and sections", ["set_notes", "notes", "set_link", "links",
                                       "set_sections", "sections"]),
    ],
}

PAGES = [
    ("documents", "Documents (common)", "Opening, saving, validating and the parts of any document.",
     [("func", "oox.document.open_document"), ("class", "oox.document.OoxDocument", COMMON)]),
    ("word", "Word", "`WordprocessingDocument` and the objects its methods return.",
     [("class", "oox.document.WordprocessingDocument", None), ("class", "oox.tables.Table", None),
      ("class", "oox.sections.Section", None), ("class", "oox.numbering.ListHandle", None),
      ("class", "oox.fieldcalc.Field", None), ("class", "oox.fieldcalc.Caption", None)]),
    ("excel", "Excel", "`SpreadsheetDocument`, cell values and the stylesheet.",
     [("class", "oox.spreadsheet.SpreadsheetDocument", None),
      ("class", "oox.spreadsheet.Formula", None), ("class", "oox.spreadsheet.CellError", None),
      ("class", "oox.xlstyles.StyleSheet", None),
      ("module", "oox.spreadsheet", ["cell", "row", "cell_ref", "column_letter", "split_ref",
                                     "excel_serial", "from_excel_serial"])]),
    ("powerpoint", "PowerPoint", "`PresentationDocument` and `SlideTable`.",
     [("class", "oox.presentation.PresentationDocument", None),
      ("class", "oox.slidetable.SlideTable", None)]),
    ("charts", "Charts", "Chart data and the builder and reader the document methods use.",
     [("class", "oox.charts.Series", None),
      ("module", "oox.charts", ["build_chart", "chart_info", "PALETTE"])]),
    ("elements", "Elements", "The element class, factories, `lenient()`, plain text and odfpy names.",
     [("class", "oox.element.OoxElement", None), ("func", "oox.element.lenient"),
      ("module", "oox.teletype", None), ("module", "oox.odfcompat", None)]),
    ("package", "Package and properties", "The zip package, parts, limits and document properties.",
     [("class", "oox.opc.Package", None), ("class", "oox.opc.Part", None),
      ("class", "oox.opc.Limits", None), ("class", "oox.properties.CoreProperties", None),
      ("class", "oox.properties.ExtendedProperties", None),
      ("class", "oox.properties.CustomProperties", None)]),
    ("utilities", "Units and errors", "Length units and the exception hierarchy.",
     [("module", "oox.units", None), ("module", "oox.errors", None)]),
]


def _resolve(dotted: str):
    module, _, name = dotted.rpartition(".")
    return getattr(importlib.import_module(module), name)


def _md(doc: str) -> str:
    """Docstring (RST-flavoured) -> Markdown: ``x`` -> `x`, indented blocks and
    doctest sessions -> code blocks."""
    doc = re.sub(r"``(.+?)``", r"`\1`", doc, flags=re.S)
    out, in_code, in_doctest = [], False, False
    for line in doc.split("\n"):
        if line.startswith(">>>") and not in_code and not in_doctest:
            if out and out[-1].strip():
                out.append("")
            out.append("```pycon")
            in_doctest = True
        if in_doctest:
            if not line.strip():
                out += ["```", ""]
                in_doctest = False
            else:
                out.append(re.sub(r"\s*# doctest:.*$", "", line))
            continue
        code = line.startswith("    ")
        if code and not in_code:
            if out and out[-1].strip():
                out.append("")
            out.append("```python")
            in_code = True
        elif not code and in_code and line.strip():
            while out and not out[-1].strip():
                out.pop()
            out.append("```")
            out.append("")
            in_code = False
        out.append(line[4:] if in_code and code else line)
    if in_code or in_doctest:
        while out and not out[-1].strip():
            out.pop()
        out.append("```")
    return "\n".join(out).strip()


def _delegate(func):
    """The feature-module function a document method hands its ``**options`` to."""
    try:
        source = inspect.getsource(func)
    except (OSError, TypeError):
        return None
    params = inspect.signature(func).parameters.values()
    if not any(p.kind is p.VAR_KEYWORD for p in params):
        return None
    for m in re.finditer(r"from \.(\w+) import (\w+)", source):
        if not re.search(rf"\b{m.group(2)}\(\s*self\b", source):
            continue
        target = getattr(importlib.import_module(f"oox.{m.group(1)}"), m.group(2), None)
        if inspect.isfunction(target):
            return target
    return None


PUBLIC_MODULES = {"oox", "oox.units", "oox.teletype", "oox.odfcompat", "oox.errors",
                  "oox.spreadsheet", "oox.charts", "oox.ns"}
_shown_modules: set[str] = set()  # internal module docs already on the current page


def _module_details(doc: str, delegate) -> list[str]:
    """Docs of the internal modules a docstring points to ("see `oox.xlcond`"),
    once per page, so the options they describe are in the reference too."""
    lines = []
    for name in dict.fromkeys(re.findall(r"`(oox\.[a-z_]+)(?:\.[a-z_]+)?`", doc)):
        if name in PUBLIC_MODULES or name in _shown_modules:
            continue
        if delegate is not None and delegate.__module__ == name and not re.search(
                rf"`{re.escape(name)}`", doc):
            continue  # pointed at the delegate function, already shown
        try:
            module = importlib.import_module(name)
        except ImportError:
            continue
        text = inspect.getdoc(module)
        if text:
            _shown_modules.add(name)
            lines += [f"**Details** (from `{name}`)", "", _md(text), ""]
    return lines


def _options(func) -> list[str]:
    """Keyword options a method takes through ``**options``, from its delegate."""
    target = _delegate(func)
    if target is None:
        return _module_details(inspect.getdoc(func) or "", None)
    own = set(inspect.signature(func).parameters)
    extra = [p for p in inspect.signature(target).parameters.values()
             if p.name not in own and p.default is not p.empty and p.kind is not p.VAR_KEYWORD]
    lines = []
    if extra:
        lines += ["**Options** (keywords)", "", "```python"]
        lines += [re.sub(r"(: )'([^']*)'", r"\1\2", str(p)) for p in extra]
        lines += ["```", ""]
    doc = inspect.getdoc(target)
    own_doc = inspect.getdoc(func) or ""
    if doc and doc.split("\n")[0] not in own_doc:
        lines += [_md(doc), ""]
    return lines + _module_details(own_doc + "\n" + (doc or ""), target)


def _signature(name: str, obj) -> str:
    try:
        sig = str(inspect.signature(obj))
    except (TypeError, ValueError):
        return f"{name}(...)"
    sig = re.sub(r"\(self, ?|\(cls, ?", "(", sig).replace("(self)", "()").replace("(cls)", "()")
    sig = re.sub(r"(: |-> )'([^']*)'", r"\1\2", sig)
    return f"{name}{sig}"


def _wrap_sig(text: str) -> str:
    if len(text) <= 88:
        return text
    head, _, rest = text.partition("(")
    params, _, tail = rest.rpartition(")")
    parts, depth, cur = [], 0, ""
    for ch in params:
        if ch == "," and depth == 0:
            parts.append(cur.strip())
            cur = ""
            continue
        depth += ch in "([{"
        depth -= ch in ")]}"
        cur += ch
    if cur.strip():
        parts.append(cur.strip())
    return head + "(\n" + "".join(f"    {p},\n" for p in parts) + ")" + tail


def _member(cls, name: str, level: str) -> list[str]:
    raw = inspect.getattr_static(cls, name)
    lines = [f"{level} `{name}`", ""]
    if isinstance(raw, property):
        kind = "property" if raw.fset else "read-only property"
        lines += [f"*{kind}*", ""]
        doc = inspect.getdoc(raw.fget) or ""
    elif isinstance(raw, (classmethod, staticmethod)):
        lines += ["```python", _wrap_sig(_signature(name, getattr(cls, name))), "```",
                  "", f"*{type(raw).__name__}*", ""]
        doc = inspect.getdoc(raw.__func__) or ""
    elif inspect.isfunction(raw):
        lines += ["```python", _wrap_sig(_signature(name, raw)), "```", ""]
        doc = inspect.getdoc(raw) or ""
        if doc:
            lines += [_md(doc), ""]
        return lines + _options(raw)
    else:
        lines += [f"*attribute* = `{raw!r}`"[:200], ""]
        doc = ""
    if doc:
        lines += [_md(doc), ""]
    return lines


def _public(cls) -> list[str]:
    names = []
    for name in sorted(vars(cls)):
        if name.startswith("_"):
            continue
        raw = vars(cls)[name]
        if isinstance(raw, (property, classmethod, staticmethod)) or inspect.isfunction(raw):
            names.append(name)
    return names


def _class(dotted: str, only: list[str] | None, problems: list[str]) -> list[str]:
    cls = _resolve(dotted)
    short = cls.__name__
    bases = [b.__name__ for b in cls.__bases__ if b is not object and b.__module__.startswith("oox")]
    lines = [f"## `{short}`", ""]
    if bases:
        lines += [f"*Based on* `{', '.join(bases)}`." + (
            " The common methods are under [Documents](documents.md)." if "OoxDocument" in bases
            else ""), ""]
    if dataclass_fields := getattr(cls, "__dataclass_fields__", None):
        lines += ["```python", _wrap_sig(_signature(short, cls)), "```", ""]
        doc = inspect.getdoc(cls)
        if doc and not doc.startswith(short + "("):
            lines += [_md(doc), ""]
        lines += ["Fields: " + ", ".join(f"`{f}`" for f in dataclass_fields), ""]
    else:
        doc = inspect.getdoc(cls)
        if doc:
            lines += [_md(doc), ""]
    names = only if only is not None else _public(cls)
    if short in TOPICS:
        placed = {n for _, ns in TOPICS[short] for n in ns}
        own = set(_public(cls)) - set(COMMON)
        for missing in sorted(own - placed):
            problems.append(f"{short}.{missing} has no topic in tools/apidocs.py TOPICS")
        for gone in sorted(placed - set(dir(cls))):
            problems.append(f"{short}.{gone} is in TOPICS but doesn't exist")
        for title, members in TOPICS[short]:
            lines += [f"### {title}", ""]
            for n in members:
                if hasattr(cls, n):
                    lines += _member(cls, n, "####")
    else:
        for n in names:
            if n == "__init__" or not hasattr(cls, n):
                continue
            lines += _member(cls, n, "###")
    for n in names:
        if inspect.getattr_static(cls, n, None) is not None and n in vars(cls) and n != "__init__":
            raw = vars(cls)[n]
            f = raw.fget if isinstance(raw, property) else getattr(raw, "__func__", raw)
            if callable(f) and not inspect.getdoc(f) and not n[0].isupper():
                problems.append(f"{short}.{n} has no docstring")
    return lines


def _module(name: str, only: list[str] | None) -> list[str]:
    module = importlib.import_module(name)
    lines = [f"## `{name}`", ""]
    doc = inspect.getdoc(module)
    if doc:
        lines += [_md(doc), ""]
    names = only or sorted(n for n, v in vars(module).items() if not n.startswith("_")
                           and getattr(v, "__module__", None) == name
                           and (inspect.isfunction(v) or inspect.isclass(v)))
    for n in names:
        value = getattr(module, n)
        lines += [f"### `{n}`", ""]
        if inspect.isclass(value):
            bases = ", ".join(b.__name__ for b in value.__bases__ if b is not object)
            lines += [f"*exception class* ({bases})" if issubclass(value, BaseException)
                      else "*class*", ""]
            doc = inspect.getdoc(value)
        elif inspect.isfunction(value):
            lines += ["```python", _wrap_sig(_signature(n, value)), "```", ""]
            doc = inspect.getdoc(value)
        else:
            lines += ["```python", f"{n} = {value!r}"[:300], "```", ""]
            doc = None
        if doc and not (inspect.isclass(value) and doc.startswith(n + "(")):
            lines += [_md(doc), ""]
    return lines


def _func(dotted: str) -> list[str]:
    f = _resolve(dotted)
    name = dotted.rsplit(".", 1)[1]
    lines = [f"## `{name}`", "", "```python", _wrap_sig(_signature(name, f)), "```", ""]
    doc = inspect.getdoc(f)
    if doc:
        lines += [_md(doc), ""]
    return lines


def render() -> tuple[dict[str, str], list[str]]:
    problems: list[str] = []
    pages: dict[str, str] = {}
    index = ["# API reference", "",
             "Generated from the docstrings by `python -m tools.apidocs`; it covers exactly the "
             "public API described in [API stability](../STABILITY.md). Element factories "
             "(`oox.ns.w.P`, `oox.ns.x.C` …) follow the schemas: see [Concepts](../concepts.md).",
             ""]
    for slug, title, blurb, items in PAGES:
        _shown_modules.clear()
        lines = [f"# {title}", "", blurb, ""]
        for item in items:
            kind = item[0]
            if kind == "class":
                lines += _class(item[1], item[2], problems)
            elif kind == "module":
                lines += _module(item[1], item[2])
            else:
                lines += _func(item[1])
        pages[f"{slug}.md"] = "\n".join(lines).rstrip() + "\n"
        index.append(f"- [{title}]({slug}.md): {blurb}")
    pages["index.md"] = "\n".join(index) + "\n"
    return pages, problems


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--check", action="store_true", help="fail if docs/reference is stale")
    args = ap.parse_args(argv)
    pages, problems = render()
    for p in problems:
        print("problem:", p, file=sys.stderr)
    stale = [n for n, text in pages.items()
             if not (OUT / n).is_file() or (OUT / n).read_text(encoding="utf-8") != text]
    if args.check:
        for n in stale:
            print(f"stale: docs/reference/{n} (run python -m tools.apidocs)", file=sys.stderr)
        return 1 if stale or problems else 0
    OUT.mkdir(parents=True, exist_ok=True)
    for n, text in pages.items():
        (OUT / n).write_text(text, encoding="utf-8")
    print(f"wrote {len(pages)} pages to docs/reference/ ({len(stale)} changed)")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())


