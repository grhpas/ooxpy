"""Bulleted and numbered lists (``word/numbering.xml``).

Writing:

    bullets = doc.new_list("bullet")
    bullets.add("First point")
    bullets.add("A sub-point", level=1)

    steps = doc.new_list("decimal")          # 1. / a. / i. ...
    steps.add("Mix")
    again = steps.restart()                  # same format, starts again at 1
    legal = doc.new_list("legal", start=3)   # 3. / 3.1. / 3.1.1. ...

Reading: ``doc.list_labels()`` gives the label Word displays for each list
paragraph ("2.", "1.3.", "c.") -- including numbering inherited from a
paragraph's style, ``lvlOverride/startOverride`` restarts and per-level
formats -- so text extraction can keep list numbers.

How OOXML models it: a ``w:abstractNum`` defines up to nine levels (format,
text such as "%1.%2.", indentation); a ``w:num`` is one *list instance*
pointing at an abstractNum, optionally with per-level overrides; a paragraph
joins a list with ``w:pPr/w:numPr(w:ilvl, w:numId)``.  As in Word, instances
that share an abstractNum share one counter, and a ``startOverride`` resets it
when its instance is first used -- which is how "Restart Numbering" works
(LibreOffice writes restarted lists exactly this way).  The level definitions
below copy Word's own defaults.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from .element import OoxElement
from .grammar import URI
from .ns import w

if TYPE_CHECKING:
    from .document import WordprocessingDocument

W = URI["w"]
KINDS = ("bullet", "decimal", "legal")

# Word's default bullet glyphs: Symbol U+F0B7, Courier New "o", Wingdings U+F0A7
_BULLETS = [("", "Symbol"), ("o", "Courier New"), ("", "Wingdings")]
_DECIMAL_FORMATS = ["decimal", "lowerLetter", "lowerRoman"]
# Word's legal (1. / 1.1. / 1.1.1.) indents: (left, hanging) in twips
_LEGAL_IND = [(360, 360), (792, 432), (1224, 504), (1728, 648), (2232, 792), (2736, 936),
              (3240, 1080), (3744, 1224), (4320, 1440)]


def _level(ilvl: int, kind: str, start: int) -> OoxElement:
    if kind == "bullet":
        char, font = _BULLETS[ilvl % 3]
        return w.Lvl(
            w.Start(val=1), w.NumFmt(val="bullet"), w.LvlText(val=char), w.LvlJc(val="left"),
            w.PPr(w.Ind(left=720 * (ilvl + 1), hanging=360)),
            w.RPr(w.RFonts(ascii=font, hAnsi=font, hint="default")),
            ilvl=ilvl)
    if kind == "decimal":
        fmt = _DECIMAL_FORMATS[ilvl % 3]
        lvl_jc = "right" if fmt == "lowerRoman" else "left"
        return w.Lvl(
            w.Start(val=start if ilvl == 0 else 1), w.NumFmt(val=fmt),
            w.LvlText(val=f"%{ilvl + 1}."), w.LvlJc(val=lvl_jc),
            w.PPr(w.Ind(left=720 * (ilvl + 1), hanging=360 if lvl_jc == "left" else 180)),
            ilvl=ilvl)
    # legal: every level shows the whole path, 1.2.3.
    left, hanging = _LEGAL_IND[ilvl]
    return w.Lvl(
        w.Start(val=start if ilvl == 0 else 1), w.NumFmt(val="decimal"),
        w.LvlText(val="".join(f"%{k}." for k in range(1, ilvl + 2))), w.LvlJc(val="left"),
        w.PPr(w.Ind(left=left, hanging=hanging)),
        ilvl=ilvl)


def abstract_num(abstract_id: int, kind: str, start: int = 1) -> OoxElement:
    if kind not in KINDS:
        raise ValueError(f"list kind must be one of {KINDS}")
    return w.AbstractNum(
        w.MultiLevelType(val="multilevel" if kind == "legal" else "hybridMultilevel"),
        *(_level(i, kind, start) for i in range(9)),
        abstractNumId=abstract_id)


class ListHandle:
    """One list instance (a ``w:num``); ``add`` appends items to the body."""

    def __init__(self, doc: WordprocessingDocument, num_id: int, abstract_id: int) -> None:
        self.doc, self.num_id, self.abstract_id = doc, num_id, abstract_id

    def __repr__(self) -> str:
        return f"<ListHandle numId={self.num_id} abstractNumId={self.abstract_id}>"

    def num_pr(self, level: int = 0) -> OoxElement:
        """A ``w:numPr`` putting a paragraph in this list at ``level`` (0-8), for paragraphs
        you build yourself."""
        if not 0 <= level <= 8:
            raise ValueError("list levels are 0-8")
        return w.NumPr(w.Ilvl(val=level), w.NumId(val=self.num_id))

    def paragraph(self, text: str = "", level: int = 0, style: str | None = "List Paragraph") -> OoxElement:
        """A list paragraph, not yet placed (add it wherever you like)."""
        from . import teletype

        ppr = w.PPr(self.num_pr(level))
        if style:
            ppr.add(w.PStyle(val=self.doc.style_id(style)))
        p = w.P(ppr)
        if text:
            teletype.addTextToElement(p, text)
        return p

    def add(self, text: str = "", level: int = 0, style: str | None = "List Paragraph") -> OoxElement:
        """Append a list item to the end of the body; returns the paragraph."""
        p = self.paragraph(text, level, style)
        self.doc.body.add(p)
        return p

    def restart(self, start: int = 1) -> ListHandle:
        """A new list instance with the same format, numbering from ``start``."""
        numbering = self.doc.numbering
        num_id = _next_id(numbering, "num", "numId")
        numbering.add(w.Num(w.AbstractNumId(val=self.abstract_id),
                            w.LvlOverride(w.StartOverride(val=start), ilvl=0),
                            numId=num_id))
        return ListHandle(self.doc, num_id, self.abstract_id)


def _next_id(numbering: OoxElement, local: str, attr: str) -> int:
    """Next free id: abstractNumId counts from 0, numId from 1 (0 means "no list")."""
    ids = [int(e.get(f"{{{W}}}{attr}", "0")) for e in numbering.iterchildren(f"{{{W}}}{local}")]
    if ids:
        return max(ids) + 1
    return 0 if local == "abstractNum" else 1


def new_list(doc: WordprocessingDocument, kind: str = "bullet", start: int = 1) -> ListHandle:
    numbering = doc.numbering
    abstract_id = _next_id(numbering, "abstractNum", "abstractNumId")
    numbering.add(abstract_num(abstract_id, kind, start))
    num_id = _next_id(numbering, "num", "numId")
    numbering.add(w.Num(w.AbstractNumId(val=abstract_id), numId=num_id))
    return ListHandle(doc, num_id, abstract_id)


# -- reading: the labels Word displays -----------------------------------------------
def _roman(n: int) -> str:
    out = ""
    for v, s in ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
                 (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")):
        while n >= v:
            out, n = out + s, n - v
    return out


def _letters(n: int) -> str:
    # Word's lowerLetter: a..z, aa..zz, aaa.. (repeat the letter)
    return chr(ord("a") + (n - 1) % 26) * ((n - 1) // 26 + 1) if n > 0 else ""


def format_number(n: int, fmt: str) -> str:
    if fmt == "lowerLetter":
        return _letters(n)
    if fmt == "upperLetter":
        return _letters(n).upper()
    if fmt == "lowerRoman":
        return _roman(n)
    if fmt == "upperRoman":
        return _roman(n).upper()
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "none":
        return ""
    return str(n)  # decimal and anything we don't model


def _child(parent: OoxElement | None, qname: str) -> OoxElement | None:
    # Never use `a or b` on lxml elements: an element without children is falsy.
    return parent.first(qname) if parent is not None else None


class _Level:
    __slots__ = ("fmt", "legal", "start", "text")

    def __init__(self, lvl: OoxElement | None) -> None:
        def val(q: str, default: str | None) -> str | None:
            el = lvl.first(q) if lvl is not None else None
            return el.get(f"{{{W}}}val", default) if el is not None else default
        self.start = int(val("w:start", "1") or 1)
        self.fmt = val("w:numFmt", "decimal") or "decimal"
        self.text = val("w:lvlText", "") or ""
        legal = lvl.first("w:isLgl") if lvl is not None else None
        self.legal = legal is not None and legal.get(f"{{{W}}}val", "1") not in ("0", "false", "off")


def list_labels(doc: WordprocessingDocument) -> list[tuple[OoxElement, str]]:
    """``(paragraph, label)`` for every numbered/bulleted paragraph, in order."""
    found = doc.related_parts(doc.main_part, "NumberingDefinitionsPart")
    if not found:
        return []
    numbering = found[0].element
    abstracts = {e.get(f"{{{W}}}abstractNumId"): e for e in numbering.iterchildren(f"{{{W}}}abstractNum")}
    nums = {e.get(f"{{{W}}}numId"): e for e in numbering.iterchildren(f"{{{W}}}num")}
    styles = {}
    st = doc.related_parts(doc.main_part, "StyleDefinitionsPart")
    if st:
        styles = {s.get(f"{{{W}}}styleId"): s for s in st[0].element.iterchildren(f"{{{W}}}style")}

    def style_num_pr(style_id: str | None, depth: int = 0) -> OoxElement | None:
        style = styles.get(style_id) if style_id else None
        if style is None or depth > 10:
            return None
        ppr = style.first("w:pPr")
        npr = ppr.first("w:numPr") if ppr is not None else None
        if npr is not None:
            return npr
        based = style.first("w:basedOn")
        return style_num_pr(based.get(f"{{{W}}}val") if based is not None else None, depth + 1)

    levels_cache: dict[str, tuple[str, list[_Level], dict[int, int]] | None] = {}

    def levels_for(num_id: str) -> tuple[str, list[_Level], dict[int, int]] | None:
        """(abstractNumId, effective levels, {level: startOverride}) for a numId."""
        if num_id in levels_cache:
            return levels_cache[num_id]
        num = nums.get(num_id)
        aref = num.first("w:abstractNumId") if num is not None else None
        abstract_id = aref.get(f"{{{W}}}val") if aref is not None else None
        abstract = abstracts.get(abstract_id)
        if num is None or abstract_id is None:
            levels_cache[num_id] = None
            return None
        lvls = {int(lv.get(f"{{{W}}}ilvl", "0")): lv
                for lv in (abstract.iterchildren(f"{{{W}}}lvl") if abstract is not None else [])}
        result = [_Level(lvls.get(i)) for i in range(9)]
        overrides: dict[int, int] = {}
        for ov in num.iterchildren(f"{{{W}}}lvlOverride"):
            i = int(ov.get(f"{{{W}}}ilvl", "0"))
            if not 0 <= i <= 8:
                continue
            if ov.first("w:lvl") is not None:
                result[i] = _Level(ov.first("w:lvl"))
            so = ov.first("w:startOverride")
            if so is not None:
                overrides[i] = int(so.get(f"{{{W}}}val", "1"))
        levels_cache[num_id] = (abstract_id, result, overrides)
        return levels_cache[num_id]

    seen_nums: set[str] = set()
    counters: dict[str, list[int | None]] = {}
    out: list[tuple[OoxElement, str]] = []
    for p in doc.body.iter(f"{{{W}}}p"):
        ppr = p.first("w:pPr")
        npr = ppr.first("w:numPr") if ppr is not None else None
        style_ref = ppr.first("w:pStyle") if ppr is not None else None
        snpr = style_num_pr(style_ref.get(f"{{{W}}}val") if style_ref is not None else None)
        num_el = _child(npr, "w:numId")
        if num_el is None:
            num_el = _child(snpr, "w:numId")
        if num_el is None:
            continue
        num_id = num_el.get(f"{{{W}}}val")
        if num_id in (None, "0"):  # numId 0 = explicitly not numbered
            continue
        ilvl_el = _child(npr, "w:ilvl")
        if ilvl_el is None:
            ilvl_el = _child(snpr, "w:ilvl")
        ilvl = min(max(int(ilvl_el.get(f"{{{W}}}val", "0")) if ilvl_el is not None else 0, 0), 8)
        found_levels = levels_for(num_id)
        if found_levels is None:
            continue
        abstract_id, levels, overrides = found_levels
        cnt = counters.setdefault(abstract_id, [None] * 9)  # shared per abstractNum
        if num_id not in seen_nums:
            seen_nums.add(num_id)
            for lv, start in overrides.items():  # "Restart Numbering"
                cnt[lv] = start - 1
                for deeper in range(lv + 1, 9):
                    cnt[deeper] = None
        cnt[ilvl] = levels[ilvl].start if cnt[ilvl] is None else cnt[ilvl] + 1  # type: ignore[operator]
        for deeper in range(ilvl + 1, 9):
            cnt[deeper] = None  # lower levels restart under a new parent
        for higher in range(ilvl):
            if cnt[higher] is None:
                cnt[higher] = levels[higher].start
        lvl = levels[ilvl]
        if lvl.fmt == "bullet":
            label = lvl.text
        else:
            label = lvl.text
            for k in range(9, 0, -1):  # %10 doesn't exist; replace %9..%1
                if f"%{k}" in label:
                    fmt = "decimal" if lvl.legal else levels[k - 1].fmt
                    label = label.replace(f"%{k}", format_number(cnt[k - 1] or 0, fmt))
        out.append((p, label))
    return out
