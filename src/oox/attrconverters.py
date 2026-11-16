"""Python value -> XML lexical value, per attribute kind.

Strings are always accepted as raw lexical values (and validated where the
kind allows it), so anything the converters don't model is still reachable.
Analogous to odfpy's ``attrconverters`` module.
"""

from __future__ import annotations

import base64
import datetime as _dt
import enum
from typing import Any

from .errors import InvalidAttributeValue
from .grammar import AttrSpec, enum_values

_BOOL_LEXICAL = {"0", "1", "true", "false", "on", "off"}
_TF_LEXICAL = {"t", "f", "true", "false"}


_STYLE_ID = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId"


def to_xml(spec: AttrSpec, value: Any) -> str:
    """Convert and check; raises ``InvalidAttributeValue`` with what was expected."""
    text = _convert(spec, value)
    if spec.lexical:
        from .lexical import check

        problem = check(spec.lexical, text)
        if problem:
            raise InvalidAttributeValue(f"{spec.qname}={text!r} is not valid: expected {problem}")
    return text


def to_xml_lenient(spec: AttrSpec, value: Any) -> str:
    """``to_xml`` for building elements: inside ``oox.lenient()`` a value that
    fails the checks is written as ``str(value)`` (markup Office accepts but the
    schema doesn't); ``validate()`` still reports it."""
    from .element import strict

    try:
        return to_xml(spec, value)
    except InvalidAttributeValue:
        if strict():
            raise
        return str(value)


def _convert(spec: AttrSpec, value: Any) -> str:
    if isinstance(value, enum.Enum):
        value = value.value
    # odfpy lets you pass a Style element where a style name is expected
    # (P(stylename=mystyle)).  OOXML equivalent: w:pStyle(val=<w:style>).
    if hasattr(value, "tag") and hasattr(value, "get"):
        style_id = value.get(_STYLE_ID)
        if style_id is None:
            raise InvalidAttributeValue(
                f"{spec.qname}: an element can only be used as a value if it has w:styleId")
        value = style_id
    kind = spec.kind
    if kind == "enum":
        from .lexical import OPEN_ENUMS

        allowed = enum_values(spec.enum) if spec.enum and spec.enum not in OPEN_ENUMS else ()
        if isinstance(value, bool) and allowed:
            # on/off enumerations (m:ST_OnOff, w:ST_OnOff in some versions):
            # Python booleans as the enumeration spells them
            for yes, no in (("1", "0"), ("true", "false"), ("on", "off"), ("t", "f")):
                if yes in allowed and no in allowed:
                    value = yes if value else no
                    break
        s = str(value)
        if allowed and s not in allowed:
            raise InvalidAttributeValue(
                f"{spec.qname}={s!r} is not one of: {', '.join(allowed)}")
        return s
    if kind == "bool":
        if isinstance(value, bool):
            return "1" if value else "0"
        s = str(value)
        if s not in _BOOL_LEXICAL:
            raise InvalidAttributeValue(f"{spec.qname}={s!r} is not a boolean")
        return s
    if kind == "tf":
        if isinstance(value, bool):
            return "t" if value else "f"
        s = str(value)
        if s not in _TF_LEXICAL:
            raise InvalidAttributeValue(f"{spec.qname}={s!r} is not t/f")
        return s
    if kind == "int":
        if isinstance(value, bool):
            raise InvalidAttributeValue(f"{spec.qname} expects an integer, got bool")
        try:
            return str(int(value))
        except (TypeError, ValueError):
            raise InvalidAttributeValue(f"{spec.qname}={value!r} is not an integer") from None
    if kind == "float":
        try:
            f = float(value)
        except (TypeError, ValueError):
            raise InvalidAttributeValue(f"{spec.qname}={value!r} is not a number") from None
        return repr(f) if not f.is_integer() else str(int(f))
    if kind == "datetime" and isinstance(value, (_dt.datetime, _dt.date)):
        return value.isoformat()
    if kind == "list" and isinstance(value, (list, tuple)):
        return " ".join(map(str, value))
    if kind == "hex" and isinstance(value, (bytes, bytearray)):
        return bytes(value).hex().upper()
    if kind == "base64" and isinstance(value, (bytes, bytearray)):
        return base64.b64encode(bytes(value)).decode("ascii")
    return str(value)
