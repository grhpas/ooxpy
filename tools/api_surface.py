"""The public API of oox, one line per name, for the API snapshot test.

    python -m tools.api_surface            # print
    python -m tools.api_surface --write    # update tests/api/public_api.txt

What counts as public (docs/STABILITY.md):

- the names ``oox`` exports (``oox.__all__``);
- every public method and property of the document classes, ``OoxElement``,
  ``Package``/``Part``/``Limits`` and of the objects their methods return
  (``Table``, ``Section``, ``ListHandle``, ``SlideTable``, ``Field``, ``Caption``,
  the property sets, ``StyleSheet``, ``Series``, ``Formula``, ``CellError``);
- the utility modules ``oox.units``, ``oox.teletype``, ``oox.odfcompat``, the
  module-level helpers of ``oox.spreadsheet`` and ``oox.charts``, ``oox.errors``.

The element factories in ``oox.ns`` are generated from the pinned Open XML
SDK data and are covered by the regeneration test instead.  Everything else
(the feature modules behind the document methods: ``xlsheet``, ``floating``,
``protection`` ...) is implementation and may change without notice.
"""

from __future__ import annotations

import argparse
import dataclasses
import inspect
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SNAPSHOT = ROOT / "tests" / "api" / "public_api.txt"

CLASSES = [
    "oox.document.OoxDocument", "oox.document.WordprocessingDocument",
    "oox.spreadsheet.SpreadsheetDocument", "oox.presentation.PresentationDocument",
    "oox.element.OoxElement", "oox.opc.Package", "oox.opc.Part", "oox.opc.Limits",
    "oox.tables.Table", "oox.sections.Section", "oox.numbering.ListHandle",
    "oox.slidetable.SlideTable", "oox.fieldcalc.Field", "oox.fieldcalc.Caption",
    "oox.properties.CoreProperties", "oox.properties.ExtendedProperties",
    "oox.properties.CustomProperties", "oox.xlstyles.StyleSheet", "oox.charts.Series",
    "oox.spreadsheet.Formula", "oox.spreadsheet.CellError",
    "oox.formatting.LineSpacing", "oox.formatting.Border", "oox.formatting.Borders",
    "oox.formatting.Tab", "oox.resolve.Formatting", "oox.resolve.RunFormat",
    "oox.resolve.ParagraphFormat", "oox.numbering.ListItem", "oox.document.ImageData",
    "oox.opc.Relationships", "oox.opc.Relationship",
    *(f"oox.records.{n}" for n in (
        "Record", "Hyperlink", "Comment", "Chart", "CellFormat", "Note", "ContentControl",
        "Picture", "Protection", "Revision", "TextBox", "ConditionalFormat", "Validation",
        "ExcelTable", "PageSetup", "SheetOptions", "WorkbookProtection", "Placeholder", "Shape",
        "SlideTableInfo", "SlideSection")),
]
# module -> the public names (None: every class and function defined there)
MODULES = {
    "oox.units": None, "oox.teletype": None, "oox.odfcompat": None, "oox.errors": None,
    "oox.spreadsheet": ["cell", "row", "cell_ref", "column_letter", "split_ref",
                        "excel_serial", "from_excel_serial"],
    "oox.charts": ["build_chart", "chart_info", "PALETTE"],
    "oox.ns": ["qn", "W", "A", "R", "MC", "WP", "PIC", "V", "O", "X", "P", "C"],
}
# lxml's own API on OoxElement (append, find, iter ...) is lxml's, not ours
LXML_BASE = "lxml.etree.ElementBase"


def _sig(obj) -> str:
    try:
        return str(inspect.signature(obj))
    except (TypeError, ValueError):
        return "(...)"


def _resolve(dotted: str):
    import importlib

    module, _, name = dotted.rpartition(".")
    return getattr(importlib.import_module(module), name)


def _members(cls) -> list[str]:
    from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

    inherited_from_lxml = set(dir(etree.ElementBase))
    out = []
    fields = {}
    if dataclasses.is_dataclass(cls):
        fields = {f.name: f for f in dataclasses.fields(cls)}
        for f in fields.values():
            out.append(f"  .{f.name}: {f.type}  field")
    for name in sorted(dir(cls)):
        if name.startswith("_") and name != "__init__":
            continue
        if name in inherited_from_lxml and not any(name in vars(k) for k in cls.__mro__
                                                   if k.__module__.startswith("oox")):
            continue
        owner = next((k for k in cls.__mro__ if name in vars(k)), None)
        if owner is None or not owner.__module__.startswith("oox"):
            continue
        if name in fields:
            continue
        raw = inspect.getattr_static(cls, name)
        if isinstance(raw, property):
            kind = "property" + ("" if raw.fset else " (read-only)")
            out.append(f"  .{name}  {kind}")
        elif isinstance(raw, (staticmethod, classmethod)):
            out.append(f"  .{name}{_sig(getattr(cls, name))}  {type(raw).__name__}")
        elif inspect.isfunction(raw):
            if name == "__init__" and owner is not cls:
                continue
            out.append(f"  .{name}{_sig(raw)}")
        elif not callable(raw):
            out.append(f"  .{name}  attribute")
    return out


def surface() -> list[str]:
    import importlib

    import oox

    lines = [f"oox.__all__ = {sorted(oox.__all__)}", ""]
    for dotted in CLASSES:
        cls = _resolve(dotted)
        bases = ", ".join(b.__name__ for b in cls.__bases__ if b is not object)
        lines.append(f"class {dotted}({bases})")
        lines += _members(cls)
        lines.append("")
    listed = set(CLASSES)
    for name, public in MODULES.items():
        module = importlib.import_module(name)
        lines.append(f"module {name}")
        for attr in sorted(public if public is not None else vars(module)):
            value = getattr(module, attr)
            if public is not None and not (inspect.isclass(value) or inspect.isfunction(value)):
                lines.append(f"  {attr}  constant")
                continue
            if attr.startswith("_") or getattr(value, "__module__", None) != name:
                continue
            if inspect.isclass(value):
                if f"{name}.{attr}" not in listed:
                    bases = ", ".join(b.__name__ for b in value.__bases__ if b is not object)
                    lines.append(f"  class {attr}({bases})")
            elif inspect.isfunction(value):
                lines.append(f"  {attr}{_sig(value)}")
        lines.append("")
    return lines


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--write", action="store_true", help=f"update {SNAPSHOT.relative_to(ROOT)}")
    args = ap.parse_args(argv)
    text = "\n".join(surface()).rstrip() + "\n"
    if args.write:
        SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
        SNAPSHOT.write_text(text, encoding="utf-8")
        print(f"wrote {SNAPSHOT.relative_to(ROOT)} ({text.count(chr(10))} lines)")
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
