"""Excel conditional formatting.

    book.conditional_format("Sheet1", "B2:B20", "cell", operator=">", value=100,
                            style={"fill": "FFC7CE", "color": "9C0006"})
    book.conditional_format("Sheet1", "C2:C20", "color_scale")          # red-yellow-green
    book.conditional_format("Sheet1", "D2:D20", "data_bar", color="638EC6")
    book.conditional_format("Sheet1", "E2:E20", "icon_set", icons="3TrafficLights1")
    book.conditional_format("Sheet1", "A2:A20", "duplicates", style={"fill": "FFEB9C"})
    book.conditional_format("Sheet1", "A2:A20", "text", text="North", style={"bold": True})
    book.conditional_format("Sheet1", "A2:E20", "formula", formula="$E2>$D2",
                            style={"fill": "C6EFCE"})
    book.conditional_formats("Sheet1")   # [ConditionalFormat(ref="B2:B20", rule="cell", ...)]

Rules: ``cell`` (operator ``>``, ``>=``, ``<``, ``<=``, ``=``, ``!=``, ``between``,
``not between``), ``formula``, ``color_scale`` (2 or 3 colours), ``data_bar``,
``icon_set``, ``top`` / ``bottom`` (``n``, ``percent=``), ``above_average`` /
``below_average``, ``duplicates`` / ``unique``, ``text`` (``how=`` contains,
not contains, begins, ends) and ``blanks`` / ``no_blanks``.  Highlight styles are
differential formats (font colour, bold, italic, fill, border, number format).
Formulas are written relative to the range's top-left cell, as Excel does.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .element import OoxElement
from .grammar import URI
from .ns import x

if TYPE_CHECKING:
    from .spreadsheet import SpreadsheetDocument

X = URI["x"]
OPERATORS = {">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan",
             "<=": "lessThanOrEqual", "=": "equal", "==": "equal", "!=": "notEqual",
             "<>": "notEqual", "between": "between", "not between": "notBetween"}
ICON_SETS = ("3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3TrafficLights2",
             "3Signs", "3Symbols", "3Symbols2", "4Arrows", "4ArrowsGray", "4RedToBlack",
             "4Rating", "4TrafficLights", "5Arrows", "5ArrowsGray", "5Rating", "5Quarters")
RULES = ("cell", "formula", "color_scale", "data_bar", "icon_set", "top", "bottom",
         "above_average", "below_average", "duplicates", "unique", "text", "blanks",
         "no_blanks")


def _formula_value(v: Any) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return repr(v) if isinstance(v, float) else str(v)
    s = str(v)
    return s[1:] if s.startswith("=") else '"' + s.replace('"', '""') + '"'


def _rgb(c: str) -> str:
    from .xlstyles import _rgb as rgb

    return rgb(c)


def _next_priority(ws: OoxElement) -> int:
    pr = [int(r.get("priority", "0")) for r in ws.iter(f"{{{X}}}cfRule")]
    return max(pr, default=0) + 1


def conditional_format(book: SpreadsheetDocument, sheet: str, ref: str, rule: str,
                       **opt: Any) -> OoxElement:
    from .xlsheet import parse_range, range_ref
    from .xlstyles import StyleSheet

    if rule not in RULES:
        raise ValueError(f"rule must be one of {RULES}")
    ws = book.sheet(sheet)
    ranges = [range_ref(*parse_range(r)) for r in ref.split()]
    top_left = ranges[0].split(":")[0]
    sqref = " ".join(ranges)
    attrs: dict[str, Any] = {"priority": _next_priority(ws)}
    if opt.get("stop"):
        attrs["stopIfTrue"] = True
    kids: list[OoxElement] = []
    style = opt.get("style")
    if style is not None and hasattr(style, "as_dict"):  # a CellFormat, as read back
        style = style.as_dict()

    def need_style() -> None:
        if not style:
            raise ValueError(f"the {rule!r} rule needs style={{...}} (fill, color, bold ...)")
        attrs["dxfId"] = StyleSheet(book).dxf(**style)

    if rule == "cell":
        op = OPERATORS.get(str(opt.get("operator", "")).lower())
        if op is None:
            raise ValueError(f"operator must be one of {list(OPERATORS)}")
        attrs.update(type="cellIs", operator=op)
        if op in ("between", "notBetween"):
            lo, hi = opt["values"]
            kids += [x.Formula(text=_formula_value(lo)), x.Formula(text=_formula_value(hi))]
        else:
            kids.append(x.Formula(text=_formula_value(opt["value"])))
        need_style()
    elif rule == "formula":
        attrs["type"] = "expression"
        kids.append(x.Formula(text=str(opt["formula"]).lstrip("=")))
        need_style()
    elif rule == "color_scale":
        colors = [_rgb(c) for c in opt.get("colors", ["F8696B", "FFEB84", "63BE7B"])]
        if len(colors) not in (2, 3):
            raise ValueError("a colour scale has 2 or 3 colours")
        attrs["type"] = "colorScale"
        cfvos = [x.Cfvo(type="min")]
        if len(colors) == 3:
            cfvos.append(x.Cfvo(type="percentile", val="50"))
        cfvos.append(x.Cfvo(type="max"))
        kids.append(x.ColorScale(*cfvos, *(x.Color(rgb=c) for c in colors)))
    elif rule == "data_bar":
        attrs["type"] = "dataBar"
        kids.append(x.DataBar(x.Cfvo(type="min"), x.Cfvo(type="max"),
                              x.Color(rgb=_rgb(opt.get("color", "638EC6")))))
    elif rule == "icon_set":
        icons = opt.get("icons", "3TrafficLights1")
        if icons not in ICON_SETS:
            raise ValueError(f"icons must be one of {ICON_SETS}")
        n = int(icons[0])
        steps = [round(100 * i / n) for i in range(n)]
        attrs["type"] = "iconSet"
        box = x.IconSet(*(x.Cfvo(type="percent", val=str(v)) for v in steps), iconSet=icons)
        if opt.get("reverse"):
            box.set("reverse", "1")
        kids.append(box)
    elif rule in ("top", "bottom"):
        attrs.update(type="top10", rank=int(opt.get("n", 10)))
        if rule == "bottom":
            attrs["bottom"] = True
        if opt.get("percent"):
            attrs["percent"] = True
        need_style()
    elif rule in ("above_average", "below_average"):
        attrs["type"] = "aboveAverage"
        if rule == "below_average":
            attrs["aboveAverage"] = False
        need_style()
    elif rule in ("duplicates", "unique"):
        attrs["type"] = "duplicateValues" if rule == "duplicates" else "uniqueValues"
        need_style()
    elif rule == "text":
        text = str(opt["text"])
        how = opt.get("how", "contains")
        q = text.replace('"', '""')
        kinds = {"contains": ("containsText", "containsText",
                              f'NOT(ISERROR(SEARCH("{q}",{top_left})))'),
                 "not contains": ("notContainsText", "notContains",
                                  f'ISERROR(SEARCH("{q}",{top_left}))'),
                 "begins": ("beginsWith", "beginsWith",
                            f'LEFT({top_left},{len(text)})="{q}"'),
                 "ends": ("endsWith", "endsWith", f'RIGHT({top_left},{len(text)})="{q}"')}
        if how not in kinds:
            raise ValueError(f"how must be one of {list(kinds)}")
        typ, op, formula = kinds[how]
        # "text" is also odfpy's keyword for element text: use the prefixed form
        attrs.update(type=typ, operator=op, x_text=text)
        kids.append(x.Formula(text=formula))
        need_style()
    elif rule in ("blanks", "no_blanks"):
        attrs["type"] = "containsBlanks" if rule == "blanks" else "notContainsBlanks"
        f = f"LEN(TRIM({top_left}))=0" if rule == "blanks" else f"LEN(TRIM({top_left}))>0"
        kids.append(x.Formula(text=f))
        need_style()
    cf_rule = x.CfRule(*kids, **attrs)
    cf = x.ConditionalFormatting(cf_rule, sqref=sqref)
    ws.add(cf)
    return cf_rule


_RULE_OF = {"expression": "formula", "colorScale": "color_scale", "dataBar": "data_bar",
            "iconSet": "icon_set", "duplicateValues": "duplicates", "uniqueValues": "unique",
            "containsBlanks": "blanks", "notContainsBlanks": "no_blanks"}
_HOW_OF = {"containsText": "contains", "notContainsText": "not contains",
           "beginsWith": "begins", "endsWith": "ends"}
_SYMBOL_OF = {"greaterThan": ">", "greaterThanOrEqual": ">=", "lessThan": "<",
              "lessThanOrEqual": "<=", "equal": "=", "notEqual": "!=", "between": "between",
              "notBetween": "not between"}


def _value_of(formula: str | None) -> Any:
    if formula is None:
        return None
    if len(formula) >= 2 and formula[0] == formula[-1] == '"':
        return formula[1:-1].replace('""', '"')
    try:
        number = float(formula)
    except ValueError:
        return "=" + formula
    return int(number) if number.is_integer() and "." not in formula else number


def conditional_formats(book: SpreadsheetDocument, sheet: str) -> list:
    """A ``ConditionalFormat`` per rule, in priority order, with the options
    ``conditional_format`` takes (``rule``, ``operator``, ``value(s)``,
    ``formula``, ``colors``, ``color``, ``icons``, ``n``, ``percent``,
    ``text``, ``how``, ``style``) where they apply."""
    from .records import CellFormat, ConditionalFormat
    from .xlstyles import StyleSheet

    st = StyleSheet(book, create=False)
    out = []
    for cf in book.sheet(sheet).iterchildren(f"{{{X}}}conditionalFormatting"):
        for r in cf.iterchildren(f"{{{X}}}cfRule"):
            kind = r.get("type")
            formulas = [f.text or "" for f in r.iterchildren(f"{{{X}}}formula")]
            opt: dict[str, Any] = {"ref": cf.get("sqref"), "type": kind,
                                   "priority": int(r.get("priority", "0")),
                                   "formulas": formulas,
                                   "stop": r.get("stopIfTrue") in ("1", "true")}
            rule = _RULE_OF.get(kind or "")
            if kind == "cellIs":
                rule = "cell"
                opt["operator"] = _SYMBOL_OF.get(r.get("operator") or "", r.get("operator"))
                if len(formulas) == 2:
                    opt["values"] = [_value_of(f) for f in formulas]
                elif formulas:
                    opt["value"] = _value_of(formulas[0])
            elif kind == "expression" and formulas:
                opt["formula"] = "=" + formulas[0]
            elif kind == "top10":
                rule = "bottom" if r.get("bottom") in ("1", "true") else "top"
                opt["n"] = int(r.get("rank", "10"))
                opt["percent"] = r.get("percent") in ("1", "true")
            elif kind == "aboveAverage":
                rule = ("below_average" if r.get("aboveAverage") in ("0", "false")
                        else "above_average")
            elif kind in _HOW_OF:
                rule, opt["how"], opt["text"] = "text", _HOW_OF[kind], r.get("text")
            scale = r.first("x:colorScale")
            if scale is not None:
                opt["colors"] = [(c.get("rgb") or "")[-6:]
                                 for c in scale.iterchildren(f"{{{X}}}color")]
            bar = r.first("x:dataBar")
            if bar is not None and bar.first("x:color") is not None:
                opt["color"] = (bar.first("x:color").get("rgb") or "")[-6:]
            icons = r.first("x:iconSet")
            if icons is not None:
                opt["icons"] = icons.get("iconSet", "3TrafficLights1")
            if r.get("dxfId") is not None:
                opt["style"] = CellFormat(**st.describe_dxf(int(r.get("dxfId"))))
            out.append(ConditionalFormat(rule=rule, **opt))
    return sorted(out, key=lambda c: c.priority)
