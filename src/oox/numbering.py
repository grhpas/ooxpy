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

from dataclasses import dataclass
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
        numbering = self.doc.ensure_numbering()
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
    numbering = doc.ensure_numbering()
    abstract_id = _next_id(numbering, "abstractNum", "abstractNumId")
    numbering.add(abstract_num(abstract_id, kind, start))
    num_id = _next_id(numbering, "num", "numId")
    numbering.add(w.Num(w.AbstractNumId(val=abstract_id), numId=num_id))
    return ListHandle(doc, num_id, abstract_id)


# -- reading: the labels Word displays -----------------------------------------------
_ONES = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
         "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen",
         "eighteen", "nineteen"]
_TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]
_ORDINAL_WORDS = {"one": "first", "two": "second", "three": "third", "five": "fifth",
                  "eight": "eighth", "nine": "ninth", "twelve": "twelfth"}
_CHICAGO = "*\u2020\u2021\u00a7"  # * † ‡ §, then doubled, tripled ...


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


def _words(n: int) -> str:
    """English cardinal words, as Word writes them ("one hundred twenty-three")."""
    if n == 0:
        return "zero"
    parts = []
    for size, name in ((1000000, "million"), (1000, "thousand"), (100, "hundred")):
        if n >= size:
            parts.append(f"{_words(n // size)} {name}")
            n %= size
    if n >= 20:
        parts.append(_TENS[n // 10] + (f"-{_ONES[n % 10]}" if n % 10 else ""))
    elif n:
        parts.append(_ONES[n])
    return " ".join(parts)


def _ordinal_words(n: int) -> str:
    text = _words(n)
    head, sep, last = text.rpartition("-" if "-" in text.rsplit(" ", 1)[-1] else " ")
    if last in _ORDINAL_WORDS:
        last = _ORDINAL_WORDS[last]
    elif last.endswith("y"):
        last = last[:-1] + "ieth"
    else:
        last += "th"
    return head + sep + last


def _ordinal(n: int) -> str:
    suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


_RUSSIAN = "абвгдежзиклмнопрстуфхцчшщэюя"  # Word skips ё й ъ ы ь
_HEBREW_ALPHABET = "אבגדהוזחטיכלמנסעפצקרשת"
_HEBREW_ONES = " אבגדהוזחט"
_HEBREW_TENS = " יכלמנסעפצ"
_HEBREW_HUNDREDS = " קרשת"
_ARABIC_ALPHA = "أبتثجحخدذرزسشصضطظعغفقكلمنهوي"
_ARABIC_ABJAD = "أبجدهوزحطيكلمنسعفصقرشتثخذضظغ"
_AIUEO = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
_IROHA = "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス"
# full-width katakana -> half-width (ヰ and ヱ have no half-width form)
_HALF_KANA = dict(zip("アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン",
                      "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ", strict=True))
_CJK_DIGITS = "〇一二三四五六七八九"
_STEMS = "甲乙丙丁戊己庚辛壬癸"
_BRANCHES = "子丑寅卯辰巳午未申酉戌亥"
_GANADA = "가나다라마바사아자차카타파하"
_CHOSUNG = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ"
_DIGITS = {"hindiNumbers": "٠١٢٣٤٥٦٧٨٩", "arabicIndic": "٠١٢٣٤٥٦٧٨٩",
           "thaiNumbers": "๐๑๒๓๔๕๖๗๘๙", "ideographDigital": _CJK_DIGITS,
           "decimalFullWidth": "０１２３４５６７８９",
           "decimalFullWidth2": "０１２３４５６７８９"}
_CYCLES = {"aiueoFullWidth": _AIUEO, "irohaFullWidth": _IROHA,
           "aiueo": "".join(_HALF_KANA[c] for c in _AIUEO),
           "iroha": "".join(_HALF_KANA.get(c, c) for c in _IROHA),
           "ideographTraditional": _STEMS, "ideographZodiac": _BRANCHES,
           "ganada": _GANADA, "chosung": _CHOSUNG}


def _repeat_letters(n: int, alphabet: str) -> str:
    """a..z, aa..zz: Word's way of going past the end of an alphabet."""
    if n <= 0:
        return ""
    size = len(alphabet)
    return alphabet[(n - 1) % size] * ((n - 1) // size + 1)


def _hebrew(n: int) -> str:
    """Hebrew numerals (gematria), as Word's hebrew1: 15 is טו, 16 is טז."""
    if not 0 < n < 1000:
        return str(n)
    out = "ת" * (n // 400)
    rest = n % 400
    out += _HEBREW_HUNDREDS[rest // 100].strip()
    rest %= 100
    if rest in (15, 16):
        return out + "ט" + _HEBREW_ONES[rest - 9]
    return out + _HEBREW_TENS[rest // 10].strip() + _HEBREW_ONES[rest % 10].strip()


def _cjk_counting(n: int, chinese: bool) -> str:
    """Counting with 十 百 千: 十一, 二十, 百 (Japanese) or 一百零一 (Chinese)."""
    if not 0 < n < 10000:
        return str(n)
    out, gap = "", False
    for size, unit in ((1000, "千"), (100, "百"), (10, "十")):
        d = n // size % 10
        if d:
            if gap and chinese:
                out += "零"
            lead = d > 1 or (chinese and (out or size > 10))
            out += (_CJK_DIGITS[d] if lead else "") + unit
            gap = False
        elif out:
            gap = True
    if n % 10:
        if gap and chinese and out:
            out += "零"
        out += _CJK_DIGITS[n % 10]
    return out


def format_number(n: int, fmt: str) -> str:
    """``n`` in a ``w:numFmt`` format.  Besides the Western formats this
    knows Russian, Hebrew and Arabic letters and numerals, Arabic-Indic and
    Thai digits, the Japanese kana orders (aiueo, iroha), CJK digits and
    counting, heavenly stems and earthly branches, and Korean ganada and
    chosung.  Others (Hindi letters, Thai letters and counting, Vietnamese
    and Korean counting, Japanese legal ...) fall back to decimal."""
    if fmt in _DIGITS:
        return "".join(_DIGITS[fmt][int(d)] for d in str(n))
    if fmt in _CYCLES:
        alphabet = _CYCLES[fmt]
        return alphabet[(n - 1) % len(alphabet)] if n > 0 else ""
    if fmt == "russianLower":
        return _repeat_letters(n, _RUSSIAN)
    if fmt == "russianUpper":
        return _repeat_letters(n, _RUSSIAN).upper()
    if fmt == "hebrew1":
        return _hebrew(n)
    if fmt == "hebrew2":
        return _repeat_letters(n, _HEBREW_ALPHABET)
    if fmt == "arabicAlpha":
        return _repeat_letters(n, _ARABIC_ALPHA)
    if fmt == "arabicAbjad":
        return _repeat_letters(n, _ARABIC_ABJAD)
    if fmt in ("japaneseCounting", "chineseCounting", "chineseCountingThousand",
               "taiwaneseCounting", "taiwaneseCountingThousand"):
        return _cjk_counting(n, chinese=not fmt.startswith("japanese"))
    if fmt in ("ideographEnclosedCircle",) and 1 <= n <= 10:
        return chr(0x3220 + n - 1)
    if fmt == "decimalEnclosedCircleChinese":
        fmt = "decimalEnclosedCircle"
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
    if fmt == "ordinal":
        return _ordinal(n)
    if fmt == "cardinalText":
        return _words(n).capitalize()
    if fmt == "ordinalText":
        return _ordinal_words(n).capitalize()
    if fmt == "hex":
        return f"{n:X}"
    if fmt == "chicago":
        return _CHICAGO[(n - 1) % 4] * ((n - 1) // 4 + 1) if n > 0 else ""
    if fmt == "decimalEnclosedParen":
        return f"({n})"
    if fmt == "decimalEnclosedFullstop":
        return f"{n}."
    if fmt == "decimalEnclosedCircle" and 1 <= n <= 20:
        return chr(0x2460 + n - 1)
    if fmt == "numberInDash":
        return f"- {n} -"
    if fmt == "none":
        return ""
    return str(n)  # decimal and the formats we don't model


# Characters stored in symbol fonts (in the U+F0xx private range, or as the plain
# code) and the Unicode character that looks the same: Symbol's Greek letters and
# common signs, and the shapes used as bullets in the Wingdings fonts and Webdings
# (after Unicode 7.0's mapping of those fonts).
def _table(family: str, start: int, chars: str) -> dict[tuple[str, int], str]:
    return {(family, start + i): ch for i, ch in enumerate(chars) if ch != " "}


_SYMBOL_GLYPHS: dict[tuple[str, int], str] = {
    **_table("symbol", 0x41, "ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ"),
    **_table("symbol", 0x61, "αβχδεφγηιϕκλμνοπθρστυϖωξψζ"),
    **_table("symbol", 0xA1, "ϒ′≤⁄∞ƒ♣♦♥♠↔←↑→↓°±″≥×∝∂•÷≠≡≈…"),
    **_table("symbol", 0xC0, "ℵℑℜ℘⊗⊕∅∩∪⊃⊇⊄⊂⊆∈∉∠∇"),
    **_table("symbol", 0xD5, "∏√⋅¬∧∨⇔⇐⇑⇒⇓◊〈"),
    ("symbol", 0x22): "∀", ("symbol", 0x24): "∃", ("symbol", 0x27): "∋",
    ("symbol", 0x2A): "∗", ("symbol", 0x2D): "−", ("symbol", 0x40): "≅",
    ("symbol", 0x5C): "∴", ("symbol", 0x5E): "⊥", ("symbol", 0x7E): "∼",
    ("symbol", 0xE5): "∑", ("symbol", 0xF1): "〉", ("symbol", 0xF2): "∫",
    ("wingdings", 0xA7): "\u25aa",   # ▪
    ("wingdings", 0xD8): "\u27a2",   # ➢
    ("wingdings", 0xFC): "\u2714",   # ✔
    ("wingdings", 0x76): "\u2756",   # ❖
    ("wingdings", 0x71): "\u2751",   # ❑
    ("wingdings", 0x6E): "\u25a0",   # ■
    ("wingdings", 0x6C): "\u25cf",   # ●
    **_table("wingdings 2", 0x4F, "🗴✓🗵☑☒☒⮾⮿"),
    **_table("wingdings 2", 0x69, "⓪①②③④⑤⑥⑦⑧⑨⑩⓿❶❷❸❹❺❻❼❽❾❿"),
    **_table("wingdings 2", 0x95, "•●⚫⬤"),
    **_table("wingdings 2", 0x9E, "⦿◾■◼⬛⬜"),
    **_table("wingdings 2", 0xAC, "⬩⬥◆◇"),
    **_table("wingdings 2", 0xB5, "⬪⬧⧫◊"),
    ("wingdings 2", 0xA8): "▣", ("wingdings 2", 0xB1): "◈", ("wingdings 2", 0xBA): "◖",
    ("wingdings 2", 0xBB): "◗", ("wingdings 2", 0xBE): "◼", ("wingdings 2", 0xBF): "⬥",
    ("wingdings 2", 0xE8): "✦", ("wingdings 2", 0xEA): "★", ("wingdings 2", 0xEB): "✶",
    ("wingdings 2", 0xED): "✷", ("wingdings 2", 0xF0): "✹", ("wingdings 2", 0xF3): "✯",
    ("wingdings 2", 0xF8): "※", ("wingdings 2", 0xF9): "⁂",
    **_table("wingdings 3", 0x21, "⭠⭢⭡⭣"),
    **_table("wingdings 3", 0x66, "←→↑↓↖↗↙↘"),
    **_table("wingdings 3", 0x70, "▲▼△▽◄►◁▷◣◢◤◥"),
    **_table("wingdings 3", 0x81, "▲▼◀▶⮜⮞⮝⮟"),
    **_table("webdings", 0x33, "⏴⏵⏶⏷⏪⏩⏮⏭⏸⏹⏺"),
    ("webdings", 0x61): "✔", ("webdings", 0x63): "□", ("webdings", 0x67): "■",
    ("webdings", 0x6E): "⚫", ("webdings", 0x73): "❓", ("webdings", 0x79): "⊖",
    ("webdings", 0x98): "✯",
}
SYMBOL_FONTS = frozenset(family for family, _ in _SYMBOL_GLYPHS)


def readable_text(text: str, font: str | None) -> str:
    """``text`` with symbol-font characters replaced by the Unicode character
    that looks the same ("\\uf0b7" in Symbol becomes "•"; "a" in Symbol "α").
    Knows Symbol, and the bullet shapes of Wingdings, Wingdings 2, Wingdings 3
    and Webdings.  A character it doesn't know stays as it is: check with
    ``needs_symbol_font()`` and draw those in ``font``."""
    family = " ".join((font or "").lower().split())
    out = []
    for ch in text:
        code = ord(ch)
        if 0xF000 <= code <= 0xF0FF:
            code -= 0xF000
        elif family not in SYMBOL_FONTS:
            out.append(ch)
            continue
        out.append(_SYMBOL_GLYPHS.get((family, code), ch))
    return "".join(out)


def needs_symbol_font(text: str) -> bool:
    """Whether ``text`` still holds private-use characters (U+E000-U+F8FF),
    such as a symbol-font code ``readable_text()`` couldn't convert: those
    show correctly only in the original symbol font."""
    return any(0xE000 <= ord(ch) <= 0xF8FF for ch in text)


def _id_int(value: str | None) -> int | None:
    """A numbering id as an int; None when absent or not a number."""
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def _id_key(value: int | str | None) -> str:
    """The key an id is indexed under: ``7``, ``"7"`` and ``"07"`` all give "7"."""
    if isinstance(value, bool) or not isinstance(value, (int, str, type(None))):
        raise TypeError(f"a numbering id is an int, not {type(value).__name__}")
    if isinstance(value, int):
        return str(value)
    number = _id_int(value)
    return str(number) if number is not None else (value or "")


def _child(parent: OoxElement | None, qname: str) -> OoxElement | None:
    # Never use `a or b` on lxml elements: an element without children is falsy.
    return parent.first(qname) if parent is not None else None


def _val(el: OoxElement | None, default: str | None = None) -> str | None:
    return el.get(f"{{{W}}}val", default) if el is not None else default


class Level:
    """One numbering level as defined (after ``lvlOverride``)."""

    __slots__ = ("element", "fmt", "font", "legal", "picture_id", "ppr", "restart", "rpr",
                 "start", "style", "suffix", "text")

    def __init__(self, lvl: OoxElement | None) -> None:
        self.element = lvl
        self.start = int(_val(_child(lvl, "w:start"), "1") or 1)
        self.fmt = _val(_child(lvl, "w:numFmt"), "decimal") or "decimal"
        self.text = _val(_child(lvl, "w:lvlText"), "") or ""
        legal = _child(lvl, "w:isLgl")
        self.legal = legal is not None and _val(legal, "1") not in ("0", "false", "off")
        restart = _val(_child(lvl, "w:lvlRestart"))
        self.restart = int(restart) if restart is not None else None
        self.suffix = _val(_child(lvl, "w:suff"), "tab") or "tab"
        self.style = _val(_child(lvl, "w:pStyle"))
        self.picture_id = _val(_child(lvl, "w:lvlPicBulletId"))
        self.ppr = _child(lvl, "w:pPr")
        self.rpr = _child(lvl, "w:rPr")
        fonts = _child(self.rpr, "w:rFonts")
        self.font = (fonts.get(f"{{{W}}}ascii") or fonts.get(f"{{{W}}}hAnsi")
                     if fonts is not None else None)


class Definitions:
    """The numbering part, indexed: which list (and level) a paragraph is in,
    and the effective levels of each list.  Read-only."""

    def __init__(self, doc: WordprocessingDocument) -> None:
        root = doc.numbering
        self.root = root
        self.abstracts = {_id_key(e.get(f"{{{W}}}abstractNumId")): e for e in
                          (root.iterchildren(f"{{{W}}}abstractNum") if root is not None else [])}
        self.nums = {_id_key(e.get(f"{{{W}}}numId")): e for e in
                     (root.iterchildren(f"{{{W}}}num") if root is not None else [])}
        styles = doc.styles
        self.styles = {s.get(f"{{{W}}}styleId"): s for s in
                       (styles.iterchildren(f"{{{W}}}style") if styles is not None else [])}
        self.default_style = next(
            (sid for sid, st in self.styles.items() if st.get(f"{{{W}}}type") == "paragraph"
             and st.get(f"{{{W}}}default") in ("1", "true", "on")), None)
        self._levels: dict[str, tuple[int, list[Level], dict[int, int]] | None] = {}

    def _abstract_id(self, num_id: str | None) -> str | None:
        found = _val(_child(self.nums.get(_id_key(num_id)), "w:abstractNumId"))
        return _id_key(found) if found is not None else None

    def _resolved_abstract(self, abstract_id: str | None, depth: int = 0) -> str | None:
        """Follow ``numStyleLink`` (a list that uses a numbering style) to the
        abstractNum that holds the levels (the one with ``styleLink``)."""
        abstract = self.abstracts.get(_id_key(abstract_id))
        link = _val(_child(abstract, "w:numStyleLink"))
        if link is None or depth > 5:
            return abstract_id
        style = self.styles.get(link)
        num_id = _val(_child(_child(_child(style, "w:pPr"), "w:numPr"), "w:numId"))
        target = self._abstract_id(num_id)
        return self._resolved_abstract(target, depth + 1) if target is not None else abstract_id

    def levels(self, num_id: int | str) -> tuple[int, list[Level], dict[int, int]] | None:
        """(abstractNumId holding the levels, the nine levels, {level: startOverride})
        for list ``num_id`` -- an int as ``ListItem.num_id`` and
        ``ParagraphFormat.numbering`` give it (the XML's string form works too)."""
        key = _id_key(num_id)
        if key in self._levels:
            return self._levels[key]
        num = self.nums.get(key)
        abstract_key = self._resolved_abstract(self._abstract_id(key))
        abstract = self.abstracts.get(abstract_key or "")
        abstract_id = _id_int(abstract_key)
        if num is None or abstract_id is None or abstract is None:
            self._levels[key] = None
            return None
        defined = {int(lv.get(f"{{{W}}}ilvl", "0")): lv
                   for lv in abstract.iterchildren(f"{{{W}}}lvl")}
        result = [Level(defined.get(i)) for i in range(9)]
        overrides: dict[int, int] = {}
        for ov in num.iterchildren(f"{{{W}}}lvlOverride"):
            i = int(ov.get(f"{{{W}}}ilvl", "0"))
            if not 0 <= i <= 8:
                continue
            if ov.first("w:lvl") is not None:
                result[i] = Level(ov.first("w:lvl"))
            so = ov.first("w:startOverride")
            if so is not None:
                overrides[i] = int(_val(so, "1") or 1)
        self._levels[key] = (abstract_id, result, overrides)
        return self._levels[key]

    def _style_chain(self, style_id: str | None) -> list[OoxElement]:
        chain: list[OoxElement] = []
        while style_id and style_id in self.styles and len(chain) < 20:
            st = self.styles[style_id]
            if any(st is c for c in chain):
                break
            chain.append(st)
            style_id = _val(_child(st, "w:basedOn"))
        return chain  # the style first, then what it is based on

    def numbering_of(self, ppr: OoxElement | None) -> tuple[int, int, bool] | None:
        """(numId, level, from_style) for a paragraph's ``w:pPr``, or None when
        the paragraph isn't in a list.  ``from_style`` says the numbering
        comes from the paragraph style rather than the paragraph."""
        npr = _child(ppr, "w:numPr")
        style_id = _val(_child(ppr, "w:pStyle")) or self.default_style
        chain = self._style_chain(style_id)
        style_npr = [n for n in (_child(_child(st, "w:pPr"), "w:numPr") for st in chain)
                     if n is not None]
        num_el = _child(npr, "w:numId")
        from_style = num_el is None
        if num_el is None:
            num_el = next((e for e in (_child(n, "w:numId") for n in style_npr) if e is not None),
                          None)
        num_id = _id_int(_val(num_el))
        if num_id is None or num_id == 0:  # numId 0 = explicitly not numbered
            return None
        ilvl_el = _child(npr, "w:ilvl")
        if ilvl_el is None:
            ilvl_el = next((e for e in (_child(n, "w:ilvl") for n in style_npr) if e is not None),
                           None)
        if ilvl_el is not None:
            ilvl = int(_val(ilvl_el, "0") or 0)
        else:  # a level linked to the style (w:lvl/w:pStyle), as headings do
            found = self.levels(num_id)
            ids = [st.get(f"{{{W}}}styleId") for st in chain]
            ilvl = next((i for sid in ids for i, lv in enumerate(found[1] if found else [])
                         if lv.style == sid), 0)
        return num_id, min(max(ilvl, 0), 8), from_style

    def picture(self, picture_id: str | None) -> OoxElement | None:
        """The image element (``v:imagedata`` or ``a:blip``) of a picture bullet."""
        if picture_id is None or self.root is None:
            return None
        for pb in self.root.iterchildren(f"{{{W}}}numPicBullet"):
            if pb.get(f"{{{W}}}numPicBulletId") == picture_id:
                for el in pb.iter():
                    if el.tag in (f"{{{URI['v']}}}imagedata", f"{{{URI['a']}}}blip"):
                        return el
        return None


@dataclass(frozen=True)
class ListItem:
    """A list paragraph and the label Word shows for it.

    ``label`` is readable text ("2.", "1.3.", "•"); ``raw_label`` keeps the
    characters as stored, which for bullets are often symbol-font codes
    ("\\uf0b7" in ``font`` Symbol).  ``suffix`` is what follows the label:
    "tab", "space" or "nothing".  ``picture`` is the image element of a
    picture bullet (``doc.image_of(item.picture)`` gives its bytes).
    ``story`` is "body", "header", "footer", "footnotes", "endnotes" or
    "comments".
    """

    paragraph: OoxElement
    label: str
    level: int
    num_id: int
    abstract_id: int
    format: str
    font: str | None
    suffix: str
    raw_label: str
    picture: OoxElement | None = None
    story: str = "body"

    @property
    def is_bullet(self) -> bool:
        return self.format == "bullet"

    @property
    def needs_symbol_font(self) -> bool:
        """Whether ``label`` still has symbol-font codes ``readable_text``
        couldn't convert: draw it in ``font``."""
        return needs_symbol_font(self.label)


_STORIES = (("header", "HeaderPart"), ("footer", "FooterPart"), ("footnotes", "FootnotesPart"),
            ("endnotes", "EndnotesPart"), ("comments", "WordprocessingCommentsPart"))


def _story_roots(doc: WordprocessingDocument, all_stories: bool):
    yield "body", doc.body
    if not all_stories:
        return
    for story, part_type in _STORIES:
        for part in doc.related_parts(doc.main_part, part_type):
            yield story, part.element


def list_labels(doc: WordprocessingDocument, *, all_stories: bool = False) -> list[ListItem]:
    """A ``ListItem`` for every numbered or bulleted paragraph, in order.

    Follows numbering inherited from paragraph styles, numbering styles
    (``numStyleLink``), ``lvlOverride``/``startOverride`` restarts,
    ``lvlRestart`` and legal numbering.  ``all_stories=True`` adds headers,
    footers, footnotes, endnotes and comments; each part numbers its lists
    on its own counters, after the body."""
    from .walk import iter_story

    defs = Definitions(doc)
    if defs.root is None:
        return []
    out: list[ListItem] = []
    for story, root in _story_roots(doc, all_stories):
        seen_nums: set[int] = set()
        counters: dict[int, list[int | None]] = {}
        for p in iter_story(root, f"{{{W}}}p"):
            found = defs.numbering_of(p.first("w:pPr"))
            if found is None:
                continue
            num_id, ilvl, _ = found
            defined = defs.levels(num_id)
            if defined is None:
                continue
            abstract_id, levels, overrides = defined
            cnt = counters.setdefault(abstract_id, [None] * 9)  # shared per abstractNum
            if num_id not in seen_nums:
                seen_nums.add(num_id)
                for lv, start in overrides.items():  # "Restart Numbering"
                    cnt[lv] = start - 1
                    for deeper in range(lv + 1, 9):
                        cnt[deeper] = None
            current = cnt[ilvl]
            cnt[ilvl] = levels[ilvl].start if current is None else current + 1
            for deeper in range(ilvl + 1, 9):
                restart = levels[deeper].restart  # None: after any higher level; 0: never
                if restart is None or (restart != 0 and ilvl < restart):
                    cnt[deeper] = None
            for higher in range(ilvl):
                if cnt[higher] is None:
                    cnt[higher] = levels[higher].start
            lvl = levels[ilvl]
            raw = lvl.text
            if lvl.fmt != "bullet":
                for k in range(9, 0, -1):  # %10 doesn't exist; replace %9..%1
                    if f"%{k}" in raw:
                        fmt = "decimal" if lvl.legal else levels[k - 1].fmt
                        raw = raw.replace(f"%{k}", format_number(cnt[k - 1] or 0, fmt))
            out.append(ListItem(
                paragraph=p, label=readable_text(raw, lvl.font), level=ilvl,
                num_id=num_id, abstract_id=abstract_id, format=lvl.fmt,
                font=lvl.font, suffix=lvl.suffix, raw_label=raw,
                picture=defs.picture(lvl.picture_id), story=story))
    return out
