"""Excel cell styles and number formats.

    s = book.cell_style(bold=True, fill="D9E2F3", border="thin", align="center")
    book.set_value("Sheet1", "A1", "Region", style=s)
    book.set_value("Sheet1", "B2", 1234.5, style=book.cell_style(number_format="#,##0.00"))
    book.set_value("Sheet1", "C2", date(2026, 11, 14))         # a date style is applied
    book.display_text("Sheet1", "B2")                          # "1,234.50", as Excel shows it

A cell's ``s`` attribute indexes ``cellXfs`` in ``xl/styles.xml``; each xf
combines a font, fill, border, number format and alignment.  oox keeps one
entry per distinct combination, so styling a million cells the same way adds
one xf.  The stylesheet is created on first use with Excel's defaults
(Calibri 11, the two required fills), so workbooks that are never styled stay
unchanged.

``format_value`` renders a value with a number-format code the way Excel's
``Range.Text`` does for the common codes: General, ``0``, ``0.00``,
``#,##0.00``, percentages, scientific, literal text and currency symbols,
sections (positive;negative;zero;text), and dates and times (``yyyy-mm-dd``,
``d mmm yyyy``, ``hh:mm:ss``, ``AM/PM`` ...).  English month and day names,
"." decimal and "," thousands separators.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from typing import TYPE_CHECKING, Any

from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

from .element import OoxElement
from .grammar import URI
from .ns import x

if TYPE_CHECKING:
    from .spreadsheet import SpreadsheetDocument

X = URI["x"]

BUILTIN_FORMATS: dict[int, str] = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
    11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm",
    17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss",
    47: "mmss.0", 48: "##0.0E+0", 49: "@",
}
# Codes written as built-in ids.  Only the ones that mean the same code in
# every regional setting: Excel draws ids 14-22 (dates) and 37-40 with the
# viewer's regional variant (Greek Excel 2007 showed id 40 as "-98,25 "
# instead of "(98.25)", 2026-11-14), so those codes are written out as
# custom formats (ids from 164).
_INVARIANT_IDS = (0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 45, 46, 47, 48, 49)
_BUILTIN_IDS = {BUILTIN_FORMATS[i]: i for i in _INVARIANT_IDS}
DATE_FORMAT = "yyyy-mm-dd"
DATETIME_FORMAT = "yyyy-mm-dd hh:mm"
TIME_FORMAT = "hh:mm:ss"
_BORDER_STYLES = ("thin", "medium", "thick", "dashed", "dotted", "double", "hair",
                  "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
                  "mediumDashDotDot", "slantDashDot")
_H_ALIGN = ("general", "left", "center", "right", "fill", "justify", "centerContinuous",
            "distributed")
_V_ALIGN = ("top", "center", "bottom", "justify", "distributed")
STYLE_OPTIONS = ("number_format", "bold", "italic", "underline", "strike", "size", "font",
                 "color", "fill", "border", "border_color", "align", "valign", "wrap", "indent",
                 "rotation")


def _rgb(value: str) -> str:
    v = value.lstrip("#").upper()
    if len(v) == 6:
        v = "FF" + v
    if len(v) != 8 or any(c not in "0123456789ABCDEF" for c in v):
        raise ValueError(f"colour must be hex RRGGBB: {value!r}")
    return v


_ATTR_DEFAULTS = {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"}


def _key(el: etree._Element) -> tuple:
    """Structural identity: namespace URIs not prefixes, booleans normalised,
    default ids dropped -- so an entry from Excel's file equals the same
    entry built by oox."""
    attrs = []
    for k, v in el.attrib.items():
        if v in ("true", "false"):
            v = "1" if v == "true" else "0"
        if _ATTR_DEFAULTS.get(k) == v or (k.startswith("apply") and v == "0"):
            continue
        attrs.append((k, v))
    return (el.tag, tuple(sorted(attrs)), (el.text or "").strip(),
            tuple(_key(c) for c in el if isinstance(c.tag, str)))


class StyleSheet:
    """The workbook's ``xl/styles.xml`` with de-duplicating builders."""

    def __init__(self, book: SpreadsheetDocument, create: bool = True) -> None:
        self.book = book
        found = book.related_parts(book.main_part, "WorkbookStylesPart")
        if found:
            self.root: OoxElement | None = found[0].element
        elif create:
            root = x.StyleSheet(
                x.Fonts(x.Font(x.Sz(val=11), x.Color(theme=1), x.Name(val="Calibri"),
                               x.Family(val=2), x.Scheme(val="minor")), count=1),
                x.Fills(x.Fill(x.PatternFill(patternType="none")),
                        x.Fill(x.PatternFill(patternType="gray125")), count=2),
                x.Borders(x.Border(x.Left(), x.Right(), x.Top(), x.Bottom(), x.Diagonal()),
                          count=1),
                x.CellStyleXfs(x.Xf(numFmtId=0, fontId=0, fillId=0, borderId=0), count=1),
                x.CellXfs(x.Xf(numFmtId=0, fontId=0, fillId=0, borderId=0, xfId=0), count=1),
                x.CellStyles(x.CellStyle(name="Normal", xfId=0, builtinId=0), count=1),
                x.Dxfs(count=0),
                x.TableStyles(count=0, defaultTableStyle="TableStyleMedium2",
                              defaultPivotStyle="PivotStyleLight16"),
            )
            book.add_part("WorkbookStylesPart", element=root)
            self.root = root
        else:
            self.root = None

    # -- lists -------------------------------------------------------------------------
    def _list(self, tag: str, factory) -> OoxElement:
        assert self.root is not None
        el = self.root.first(f"x:{tag}")
        if el is None:
            el = factory()
            self.root.add(el)
        return el

    def _index(self, tag: str, factory, item: OoxElement) -> int:
        """Index of an equal entry in a list (fonts, fills ...), adding it if new."""
        lst = self._list(tag, factory)
        key = _key(item)
        for i, existing in enumerate(lst):
            if _key(existing) == key:
                return i
        lst.add(item)
        lst.set("count", str(len(lst)))
        return len(lst) - 1

    # -- number formats -------------------------------------------------------------------
    def number_format_id(self, code: str) -> int:
        """The id for format ``code``: a locale-invariant built-in id, or a custom format
        added on first use."""
        if code in _BUILTIN_IDS:
            return _BUILTIN_IDS[code]
        fmts = self._list("numFmts", x.NumFmts)
        for nf in fmts.iterchildren(f"{{{X}}}numFmt"):
            if nf.get("formatCode") == code:
                return int(nf.get("numFmtId", "0"))
        ids = [int(nf.get("numFmtId", "0")) for nf in fmts.iterchildren(f"{{{X}}}numFmt")]
        new = max([163, *ids]) + 1
        fmts.add(x.NumFmt(numFmtId=new, formatCode=code))
        fmts.set("count", str(len(fmts)))
        return new

    def number_format_code(self, num_fmt_id: int) -> str:
        """The format code for number format id ``num_fmt_id`` (built-in or custom)."""
        if self.root is not None:
            fmts = self.root.first("x:numFmts")
            if fmts is not None:
                for nf in fmts.iterchildren(f"{{{X}}}numFmt"):
                    if int(nf.get("numFmtId", "-1")) == num_fmt_id:
                        return nf.get("formatCode") or "General"
        return BUILTIN_FORMATS.get(num_fmt_id, "General")

    # -- building ------------------------------------------------------------------------
    def _default_font(self) -> OoxElement:
        import copy

        fonts = self._list("fonts", x.Fonts)
        return copy.deepcopy(fonts[0]) if len(fonts) else x.Font(x.Sz(val=11), x.Name(val="Calibri"))

    def _font(self, **fmt: Any) -> int | None:
        if all(fmt.get(k) is None for k in ("bold", "italic", "underline", "strike", "size",
                                            "font", "color")):
            return None
        font = self._default_font()
        props = {child.tag.rpartition("}")[2]: child for child in font}

        def put(tag: str, el: OoxElement | None) -> None:
            if tag in props:
                font.remove(props.pop(tag))
            if el is not None:
                font.add(el)
                props[tag] = el

        for key, tag, factory in (("bold", "b", x.B), ("italic", "i", x.I),
                                  ("strike", "strike", x.Strike)):
            if fmt.get(key) is not None:
                put(tag, factory() if fmt[key] else None)
        if fmt.get("underline") is not None:
            u = fmt["underline"]
            put("u", None if u is False else (x.U() if u is True else x.U(val=u)))
        if fmt.get("size") is not None:
            put("sz", x.Sz(val=float(fmt["size"])))
        if fmt.get("color") is not None:
            put("color", x.Color(rgb=_rgb(fmt["color"])))
        if fmt.get("font") is not None:
            put("name", x.Name(val=fmt["font"]))
            put("scheme", None)  # an explicit font isn't the theme's
        # Excel's order: b i strike u sz color name family scheme
        order = ["b", "i", "strike", "u", "sz", "color", "name", "family", "charset", "scheme"]
        def rank(child: OoxElement) -> int:
            local = str(child.tag).rpartition("}")[2]
            return order.index(local) if local in order else 99

        kids = sorted(list(font), key=rank)
        for k in kids:
            font.append(k)
        return self._index("fonts", x.Fonts, font)

    def _fill(self, fill: str | None) -> int | None:
        if fill is None:
            return None
        el = x.Fill(x.PatternFill(x.FgColor(rgb=_rgb(fill)), x.BgColor(indexed=64),
                                  patternType="solid"))
        return self._index("fills", x.Fills, el)

    def _border(self, border: str | dict | None, color: str | None) -> int | None:
        if border is None:
            return None
        sides = border if isinstance(border, dict) else \
            {s: border for s in ("left", "right", "top", "bottom")}
        el = x.Border()
        for side, factory in (("left", x.Left), ("right", x.Right), ("top", x.Top),
                              ("bottom", x.Bottom), ("diagonal", x.Diagonal)):
            style = sides.get(side)
            if style is None:
                el.add(factory())
                continue
            if style not in _BORDER_STYLES:
                raise ValueError(f"border style must be one of {_BORDER_STYLES}")
            col = x.Color(rgb=_rgb(color)) if color else x.Color(indexed=64)
            el.add(factory(col, style=style))
        return self._index("borders", x.Borders, el)

    def xf(self, **fmt: Any) -> int:
        """Index into cellXfs for this combination of formatting."""
        fmt = {k: v for k, v in fmt.items() if v is not None}  # None: not set (as read)
        unknown = set(fmt) - set(STYLE_OPTIONS)
        if unknown:
            raise TypeError(f"unknown cell style option {sorted(unknown)}; "
                            f"known: {', '.join(STYLE_OPTIONS)}")
        attrs: dict[str, Any] = {"numFmtId": 0, "fontId": 0, "fillId": 0, "borderId": 0, "xfId": 0}
        if fmt.get("number_format"):
            attrs["numFmtId"] = self.number_format_id(fmt["number_format"])
            attrs["applyNumberFormat"] = attrs["numFmtId"] != 0
        font = self._font(**fmt)
        if font is not None:
            attrs["fontId"], attrs["applyFont"] = font, True
        fill = self._fill(fmt.get("fill"))
        if fill is not None:
            attrs["fillId"], attrs["applyFill"] = fill, True
        border = self._border(fmt.get("border"), fmt.get("border_color"))
        if border is not None:
            attrs["borderId"], attrs["applyBorder"] = border, True
        align: dict[str, Any] = {}
        if fmt.get("align") is not None:
            if fmt["align"] not in _H_ALIGN:
                raise ValueError(f"align must be one of {_H_ALIGN}")
            align["horizontal"] = fmt["align"]
        if fmt.get("valign") is not None:
            if fmt["valign"] not in _V_ALIGN:
                raise ValueError(f"valign must be one of {_V_ALIGN}")
            align["vertical"] = fmt["valign"]
        if fmt.get("wrap"):
            align["wrapText"] = True
        if fmt.get("indent"):
            align["indent"] = int(fmt["indent"])
        if fmt.get("rotation"):
            align["textRotation"] = int(fmt["rotation"])
        el = x.Xf(**attrs)
        if align:
            el.add(x.Alignment(**align))
            el.set("applyAlignment", "1")
        return self._index("cellXfs", x.CellXfs, el)

    def dxf(self, **fmt: Any) -> int:
        """Index into ``dxfs``: a differential format (only what changes), as
        conditional formatting and tables use.  Options: bold, italic,
        underline, strike, color, fill, border, border_color, number_format."""
        allowed = {"bold", "italic", "underline", "strike", "color", "fill", "border",
                   "border_color", "number_format"}
        fmt = {k: v for k, v in fmt.items() if v is not None}  # None: not set (as read)
        unknown = set(fmt) - allowed
        if unknown:
            raise TypeError(f"unknown highlight option {sorted(unknown)}; known: {sorted(allowed)}")
        el = x.Dxf()
        font = x.Font()
        for key, factory in (("bold", x.B), ("italic", x.I), ("strike", x.Strike)):
            if fmt.get(key) is not None:
                font.add(factory(val=bool(fmt[key])))
        if fmt.get("underline"):
            font.add(x.U())
        if fmt.get("color") is not None:
            font.add(x.Color(rgb=_rgb(fmt["color"])))
        if len(font):
            el.add(font)
        if fmt.get("number_format"):
            code = fmt["number_format"]
            el.add(x.NumFmt(numFmtId=self.number_format_id(code), formatCode=code))
        if fmt.get("fill") is not None:
            # in a dxf the solid colour is the *background* colour
            el.add(x.Fill(x.PatternFill(x.BgColor(rgb=_rgb(fmt["fill"])))))
        if fmt.get("border") is not None:
            sides = fmt["border"] if isinstance(fmt["border"], dict) else \
                {s: fmt["border"] for s in ("left", "right", "top", "bottom")}
            border = x.Border()
            for side, factory in (("left", x.Left), ("right", x.Right), ("top", x.Top),
                                  ("bottom", x.Bottom)):
                style = sides.get(side)
                if style is None:
                    continue
                if style not in _BORDER_STYLES:
                    raise ValueError(f"border style must be one of {_BORDER_STYLES}")
                col = fmt.get("border_color")
                border.add(factory(x.Color(rgb=_rgb(col)) if col else x.Color(indexed=64),
                                   style=style))
            el.add(border)
        return self._index("dxfs", x.Dxfs, el)

    def describe_dxf(self, index: int) -> dict[str, Any]:
        """The differential format (conditional-format highlight) at ``index`` as keywords
        like ``cell_style`` takes."""
        out: dict[str, Any] = {}
        lst = self.root.first("x:dxfs") if self.root is not None else None
        if lst is None or index >= len(lst):
            return out
        d = lst[index]
        f = d.first("x:font")
        if f is not None:
            for key, tag in (("bold", "b"), ("italic", "i"), ("strike", "strike")):
                if f.first(f"x:{tag}") is not None:
                    out[key] = _flag(f.first(f"x:{tag}"))
            if f.first("x:u") is not None:
                out["underline"] = True
            c = f.first("x:color")
            if c is not None and c.get("rgb"):
                out["color"] = c.get("rgb")[-6:]
        fill = d.first("x:fill")
        pf = fill.first("x:patternFill") if fill is not None else None
        if pf is not None:
            bg = pf.first("x:bgColor")
            fg = pf.first("x:fgColor")
            col = bg if bg is not None and bg.get("rgb") else fg
            if col is not None and col.get("rgb"):
                out["fill"] = col.get("rgb")[-6:]
        nf = d.first("x:numFmt")
        if nf is not None:
            out["number_format"] = nf.get("formatCode")
        return out

    # -- reading -------------------------------------------------------------------------
    def describe(self, index: int) -> dict[str, Any]:
        """The formatting of cellXfs[index] as style options."""
        out: dict[str, Any] = {"number_format": "General", "bold": False, "italic": False,
                               "underline": False, "size": 11.0, "font": "Calibri",
                               "color": None, "fill": None, "border": None, "align": None,
                               "valign": None, "wrap": False}
        if self.root is None:
            return out
        xfs = self.root.first("x:cellXfs")
        if xfs is None or index >= len(xfs):
            return out
        xf = xfs[index]
        out["number_format"] = self.number_format_code(int(xf.get("numFmtId", "0")))
        fonts = self.root.first("x:fonts")
        fid = int(xf.get("fontId", "0"))
        if fonts is not None and fid < len(fonts):
            f = fonts[fid]
            out["bold"] = _flag(f.first("x:b"))
            out["italic"] = _flag(f.first("x:i"))
            u = f.first("x:u")
            out["underline"] = u is not None and u.get("val", "single") != "none"
            if f.first("x:sz") is not None:
                out["size"] = float(f.first("x:sz").get("val", "11"))
            if f.first("x:name") is not None:
                out["font"] = f.first("x:name").get("val")
            c = f.first("x:color")
            if c is not None and c.get("rgb"):
                out["color"] = c.get("rgb")[-6:]
        fills = self.root.first("x:fills")
        fill_id = int(xf.get("fillId", "0"))
        if fills is not None and fill_id < len(fills):
            pf = fills[fill_id].first("x:patternFill")
            if pf is not None and pf.get("patternType") == "solid":
                fg = pf.first("x:fgColor")
                if fg is not None and fg.get("rgb"):
                    out["fill"] = fg.get("rgb")[-6:]
        borders = self.root.first("x:borders")
        bid = int(xf.get("borderId", "0"))
        if borders is not None and bid < len(borders):
            styles = {side: el.get("style") for side in ("left", "right", "top", "bottom")
                      for el in [borders[bid].first(f"x:{side}")] if el is not None
                      and el.get("style")}
            if styles:
                vals = set(styles.values())
                out["border"] = vals.pop() if len(vals) == 1 and len(styles) == 4 else styles
        al = xf.first("x:alignment")
        if al is not None:
            out["align"] = al.get("horizontal")
            out["valign"] = al.get("vertical")
            out["wrap"] = al.get("wrapText") in ("1", "true")
        return out


def _flag(el: OoxElement | None) -> bool:
    return el is not None and el.get("val", "1") not in ("0", "false")


def styles(book: SpreadsheetDocument, create: bool = True) -> StyleSheet:
    return StyleSheet(book, create)


# -- number formats: dates ----------------------------------------------------------------
_DATE_IDS = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))


def _strip_literals(code: str) -> str:
    code = re.sub(r'"[^"]*"', "", code)
    code = re.sub(r"\\.", "", code)
    code = re.sub(r"\[(?![hms]+\])[^\]]*\]", "", code, flags=re.I)  # colours, locales
    return code


def is_date_format(code: str, num_fmt_id: int | None = None) -> bool:
    if num_fmt_id is not None and num_fmt_id in _DATE_IDS:
        return True
    if code in ("General", "@"):
        return False
    bare = _strip_literals(code.split(";")[0])
    return bool(re.search(r"[dmyhs]", bare, re.I)) and not re.search(r"[0#?]", bare)


# -- number formats: rendering --------------------------------------------------------------
def _fixed(value: float, decimals: int) -> str:
    """``f"{value:.{decimals}f}"`` with Excel's rounding: halves away from
    zero, on the decimal value Excel shows (repr), not the binary one."""
    from decimal import ROUND_HALF_UP, Decimal

    q = Decimal(1).scaleb(-decimals)
    return str(Decimal(repr(float(value))).quantize(q, rounding=ROUND_HALF_UP))


def _general(value: float) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if float(value).is_integer() and abs(value) < 1e11:
        return str(int(value))
    a = abs(value)
    if 1e-9 <= a < 1e11:  # Excel's General: at most 11 characters (sign aside)
        digits = len(str(int(a)))
        decimals = max(0, 11 - digits - 1)
        text = _fixed(a, decimals).rstrip("0").rstrip(".") if decimals else _fixed(a, 0)
        if text not in ("", "0"):
            return ("-" if value < 0 else "") + text
    text = f"{value:.5e}"
    if "e" in text:
        mant, exp = text.split("e")
        sign = "-" if exp.startswith("-") else "+"
        mant = f"{float(mant):.5f}".rstrip("0").rstrip(".")
        return f"{mant}E{sign}{abs(int(exp)):02d}"
    return text


def _sections(code: str) -> list[str]:
    out, buf, quoted, i = [], "", False, 0
    while i < len(code):
        ch = code[i]
        if ch == '"':
            quoted = not quoted
        elif ch == "\\" and not quoted and i + 1 < len(code):
            buf += code[i:i + 2]
            i += 2
            continue
        elif ch == ";" and not quoted:
            out.append(buf)
            buf = ""
            i += 1
            continue
        buf += ch
        i += 1
    out.append(buf)
    return out


def _literal_tokens(section: str) -> list[tuple[str, str]]:
    """Split a section into ("lit", text) and ("code", text) pieces."""
    out: list[tuple[str, str]] = []
    i = 0
    while i < len(section):
        ch = section[i]
        if ch == '"':
            j = section.index('"', i + 1) if '"' in section[i + 1:] else len(section)
            out.append(("lit", section[i + 1:j]))
            i = j + 1
        elif ch == "\\" and i + 1 < len(section):
            out.append(("lit", section[i + 1]))
            i += 2
        elif ch == "_" and i + 1 < len(section):
            out.append(("lit", " "))
            i += 2
        elif ch == "*" and i + 1 < len(section):
            i += 2  # fill character: no width to fill here
        elif ch == "[":
            j = section.find("]", i)
            inner = section[i + 1:j] if j != -1 else ""
            if re.fullmatch(r"[hms]+", inner, re.I):
                out.append(("code", section[i:j + 1]))
            elif inner.startswith("$"):
                out.append(("lit", inner[1:].split("-")[0]))  # [$€-407] -> €
            i = j + 1 if j != -1 else len(section)
        else:
            out.append(("code", ch))
            i += 1
    return out


def _format_number(value: float, section: str, drop_sign: bool) -> str:
    pieces = _literal_tokens(section)
    code = "".join(t for k, t in pieces if k == "code")
    percent = code.count("%")
    v = abs(value) if drop_sign else value
    v *= 100 ** percent
    sci = re.search(r"E[+-]", code, re.I)
    placeholder = re.compile(r"[0#?.,]")
    first = next((i for i, (k, t) in enumerate(pieces) if k == "code" and placeholder.match(t)),
                 None)
    if first is None:
        return "".join(t if k == "lit" else ("" if t == "%" else t) for k, t in pieces) \
            if not percent else "".join(t for k, t in pieces if k == "lit")
    last = max(i for i, (k, t) in enumerate(pieces)
               if k == "code" and (placeholder.match(t) or (sci and t.upper() in "E+-")))
    number_code = "".join(t for _, t in pieces[first:last + 1])
    prefix = "".join(t if k == "lit" else ("" if t == "%" else t) for k, t in pieces[:first])
    suffix = "".join(t for k, t in pieces[last + 1:])
    if sci:
        mant_code = re.split(r"E[+-]", number_code, flags=re.I)[0]
        exp_digits = len(re.split(r"E[+-]", number_code, flags=re.I)[1]) or 1
        decimals = len(mant_code.split(".")[1]) if "." in mant_code else 0
        if v == 0:
            exp = 0
        else:
            exp = math.floor(math.log10(abs(v)))
        mant = v / (10 ** exp) if v else 0.0
        if abs(float(_fixed(mant, decimals))) >= 10:
            mant, exp = mant / 10, exp + 1
        sign = "-" if exp < 0 else "+"
        body = f"{_fixed(mant, decimals)}E{sign}{abs(exp):0{exp_digits}d}"
        return prefix + body + suffix
    int_code, _, dec_code = number_code.partition(".")
    scale = 0
    while int_code.endswith(","):
        int_code, scale = int_code[:-1], scale + 1
    v /= 1000 ** scale
    grouping = "," in int_code
    min_int = int_code.count("0")
    decimals = sum(1 for c in dec_code if c in "0#?")
    min_dec = sum(1 for c in dec_code if c == "0")
    negative = v < 0
    text = _fixed(abs(v), decimals)
    ip, _, dp = text.partition(".")
    if decimals and min_dec < decimals:
        dp = dp.rstrip("0")
        if len(dp) < min_dec:
            dp = dp.ljust(min_dec, "0")
    ip = ip.lstrip("0")
    ip = ip.rjust(min_int, "0") if min_int else ip
    if grouping and ip:
        ip = f"{int(ip):,}".rjust(len(ip), "0") if ip.isdigit() else ip
    body = ip + ("." + dp if "." in number_code else "")  # "0.##" shows 12 as "12."
    if negative and not drop_sign and float(body.replace(",", "") or 0) != 0:
        body = "-" + body
    return prefix + body + suffix


_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
           "September", "October", "November", "December"]
_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_DATE_TOKEN = re.compile(r"\[h+\]|\[m+\]|\[s+\]|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|"
                         r"ss|s|AM/PM|am/pm|A/P|a/p|\.0+", re.I)


def _format_date(value: _dt.datetime, section: str) -> str:
    pieces = _literal_tokens(section)
    code = "".join(t if k == "code" else "\0" + t + "\0" for k, t in pieces)
    tokens: list[tuple[str, str]] = []
    pos = 0
    for m in _DATE_TOKEN.finditer(code):
        if m.start() > pos:
            tokens.append(("lit", code[pos:m.start()]))
        tokens.append(("tok", m.group(0)))
        pos = m.end()
    tokens.append(("lit", code[pos:]))
    ampm = any(t.lower() in ("am/pm", "a/p") for k, t in tokens if k == "tok")
    out = []
    for i, (kind, t) in enumerate(tokens):
        if kind == "lit":
            out.append(t.replace("\0", ""))
            continue
        low = t.lower()
        if low in ("m", "mm"):
            prev = next((tt.lower() for kk, tt in reversed(tokens[:i]) if kk == "tok"), "")
            nxt = next((tt.lower() for kk, tt in tokens[i + 1:] if kk == "tok"), "")
            if prev.startswith("h") or prev.startswith("[h") or nxt.startswith("s"):
                out.append(f"{value.minute:02d}" if low == "mm" else str(value.minute))
                continue
            out.append(f"{value.month:02d}" if low == "mm" else str(value.month))
            continue
        hour = value.hour % 12 or 12 if ampm else value.hour
        out.append({
            "yyyy": f"{value.year:04d}", "yy": f"{value.year % 100:02d}",
            "mmmmm": _MONTHS[value.month - 1][0], "mmmm": _MONTHS[value.month - 1],
            "mmm": _MONTHS[value.month - 1][:3], "dddd": _DAYS[value.weekday()],
            "ddd": _DAYS[value.weekday()][:3], "dd": f"{value.day:02d}", "d": str(value.day),
            "hh": f"{hour:02d}", "h": str(hour), "ss": f"{value.second:02d}",
            "s": str(value.second),
            "am/pm": ("AM" if value.hour < 12 else "PM") if t == "AM/PM" else
            ("am" if value.hour < 12 else "pm"),
            "a/p": ("A" if value.hour < 12 else "P") if t == "A/P" else
            ("a" if value.hour < 12 else "p"),
        }.get(low, t))
    return "".join(out)


def format_value(value: Any, code: str) -> str:
    """The text Excel shows for ``value`` with number format ``code``."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    sections = _sections(code or "General")
    if isinstance(value, str):
        text_section = sections[3] if len(sections) > 3 else (
            sections[0] if "@" in sections[0] else None)
        if text_section is None:
            return value
        return "".join(value if (k == "code" and t == "@") else t
                       for k, t in _literal_tokens(text_section))
    if isinstance(value, (_dt.datetime, _dt.date)):
        from .spreadsheet import excel_serial

        value = excel_serial(value)
    v = float(value)
    if len(sections) >= 3 and v == 0:
        section, drop = sections[2], True
    elif len(sections) >= 2 and v < 0:
        section, drop = sections[1], True
    else:
        section, drop = sections[0], False
    if re.sub(r"\[[^\]]*\]", "", section).strip().lower() in ("general", ""):
        prefix = "".join(t for k, t in _literal_tokens(section) if k == "lit")
        text = _general(abs(v) if drop else v)
        return prefix + text
    if is_date_format(section):
        from .spreadsheet import from_excel_serial

        when = from_excel_serial(v)
        if not re.search(r"\.0", section):  # Excel rounds to the shown second
            when = (when + _dt.timedelta(microseconds=500000)).replace(microsecond=0)
        return _format_date(when, section)
    return _format_number(v, section, drop)
