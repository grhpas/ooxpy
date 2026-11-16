"""Image inspection from file headers (no Pillow dependency).

Gives the content type, file extension, pixel size and resolution needed to
place a picture at its natural size.  Supported: PNG, JPEG, GIF, BMP, TIFF,
EMF and WMF (placeable), which every Office version since 2007 displays, and
SVG, which Office shows from 2016 on and which therefore always travels with
a PNG fallback (``add_picture(svg, fallback=png)``).

Vector formats are described with ``dpi`` = units per inch, so
``native_size()`` gives their physical size: EMF frames are in 0.01 mm
(2540 per inch), placeable WMF declares its own units per inch, and SVG
lengths are converted to CSS pixels (96 per inch).
"""

from __future__ import annotations

import os
import struct
from dataclasses import dataclass
from typing import IO

from .units import EMU_PER_INCH

DEFAULT_DPI = 96  # what Office assumes when a file carries no resolution


class UnsupportedImage(ValueError):
    pass


@dataclass(frozen=True)
class ImageInfo:
    content_type: str
    ext: str  # without dot
    width_px: int
    height_px: int
    dpi_x: float = DEFAULT_DPI
    dpi_y: float = DEFAULT_DPI
    vector: bool = False

    @property
    def needs_fallback(self) -> bool:
        """SVG: Office before 2016 needs a raster fallback picture."""
        return self.ext == "svg"

    def native_size(self) -> tuple[int, int]:
        """Size in EMU at the image's own resolution."""
        return (round(self.width_px * EMU_PER_INCH / self.dpi_x),
                round(self.height_px * EMU_PER_INCH / self.dpi_y))

    def scaled(self, width: int | None, height: int | None) -> tuple[int, int]:
        """Requested size in EMU; one missing dimension keeps the aspect ratio."""
        cx, cy = self.native_size()
        if width and height:
            return int(width), int(height)
        if width:
            return int(width), round(cy * width / cx)
        if height:
            return round(cx * height / cy), int(height)
        return cx, cy


def read_blob(image: str | os.PathLike | bytes | IO[bytes]) -> bytes:
    if isinstance(image, (bytes, bytearray)):
        return bytes(image)
    if isinstance(image, (str, os.PathLike)):
        with open(image, "rb") as f:
            return f.read()
    return image.read()


def _dpi(v: float) -> float:
    """Whole dpi: PNG/BMP store pixels per metre as integers (150 dpi -> 5906 ppm)."""
    return float(round(v)) if 1 <= v <= 10000 else DEFAULT_DPI


def _png(b: bytes) -> ImageInfo:
    w, h = struct.unpack(">II", b[16:24])
    dx = dy = DEFAULT_DPI
    pos = 8
    while pos + 8 <= len(b):
        length, ctype = struct.unpack(">I4s", b[pos:pos + 8])
        if ctype == b"pHYs" and length >= 9:
            ppux, ppuy, unit = struct.unpack(">IIB", b[pos + 8:pos + 17])
            if unit == 1:  # per metre
                dx, dy = _dpi(ppux * 0.0254), _dpi(ppuy * 0.0254)
            break
        if ctype in (b"IDAT", b"IEND"):
            break
        pos += 12 + length
    return ImageInfo("image/png", "png", w, h, dx, dy)


_SOF = {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}


def _jpeg(b: bytes) -> ImageInfo:
    pos, dx, dy = 2, DEFAULT_DPI, DEFAULT_DPI
    while pos + 4 <= len(b):
        if b[pos] != 0xFF:
            pos += 1
            continue
        marker = b[pos + 1]
        if marker == 0xFF:  # fill byte
            pos += 1
            continue
        if marker in (0x01, *range(0xD0, 0xDA)):  # standalone markers
            pos += 2
            continue
        (seglen,) = struct.unpack(">H", b[pos + 2:pos + 4])
        seg = b[pos + 4:pos + 2 + seglen]
        if marker == 0xE0 and seg[:5] == b"JFIF\0" and len(seg) >= 12:
            units, xd, yd = seg[7], *struct.unpack(">HH", seg[8:12])
            if units == 1:
                dx, dy = _dpi(xd), _dpi(yd)
            elif units == 2:
                dx, dy = _dpi(xd * 2.54), _dpi(yd * 2.54)
        elif marker in _SOF:
            h, w = struct.unpack(">HH", seg[1:5])
            return ImageInfo("image/jpeg", "jpeg", w, h, dx, dy)
        pos += 2 + seglen
    raise UnsupportedImage("JPEG without a frame header")


def _gif(b: bytes) -> ImageInfo:
    w, h = struct.unpack("<HH", b[6:10])
    return ImageInfo("image/gif", "gif", w, h)


def _bmp(b: bytes) -> ImageInfo:
    (hdr,) = struct.unpack("<I", b[14:18])
    if hdr == 12:  # OS/2 BITMAPCOREHEADER
        w, h = struct.unpack("<HH", b[18:22])
        return ImageInfo("image/bmp", "bmp", w, h)
    w, h = struct.unpack("<ii", b[18:26])
    dx = dy = DEFAULT_DPI
    if hdr >= 40:
        ppmx, ppmy = struct.unpack("<ii", b[38:46])
        dx, dy = _dpi(ppmx * 0.0254), _dpi(ppmy * 0.0254)
    return ImageInfo("image/bmp", "bmp", w, abs(h), dx, dy)


def _tiff(b: bytes) -> ImageInfo:
    end = "<" if b[:2] == b"II" else ">"
    (ifd,) = struct.unpack(end + "I", b[4:8])
    (n,) = struct.unpack(end + "H", b[ifd:ifd + 2])
    tags: dict[int, float] = {}
    sizes = {3: 2, 4: 4, 5: 8}
    for i in range(n):
        entry = b[ifd + 2 + 12 * i: ifd + 14 + 12 * i]
        tag, typ, count = struct.unpack(end + "HHI", entry[:8])
        if typ == 3:      # SHORT
            (val,) = struct.unpack(end + "H", entry[8:10])
        elif typ == 4:    # LONG
            (val,) = struct.unpack(end + "I", entry[8:12])
        elif typ == 5:    # RATIONAL: stored at an offset
            (off,) = struct.unpack(end + "I", entry[8:12])
            num, den = struct.unpack(end + "II", b[off:off + 8])
            val = num / den if den else 0
        else:
            continue
        if count == 1 and typ in sizes:
            tags[tag] = val
    if 256 not in tags or 257 not in tags:
        raise UnsupportedImage("TIFF without width/height tags")
    unit = tags.get(296, 2)  # 2 = inch, 3 = centimetre
    factor = 2.54 if unit == 3 else 1.0
    dx = _dpi(tags.get(282, DEFAULT_DPI) * factor) if 282 in tags else DEFAULT_DPI
    dy = _dpi(tags.get(283, DEFAULT_DPI) * factor) if 283 in tags else DEFAULT_DPI
    return ImageInfo("image/tiff", "tiff", int(tags[256]), int(tags[257]), dx, dy)


def _emf(b: bytes) -> ImageInfo:
    left, top, right, bottom = struct.unpack("<4i", b[24:40])  # rclFrame, 0.01 mm
    return ImageInfo("image/x-emf", "emf", max(right - left, 1), max(bottom - top, 1),
                     2540, 2540, vector=True)


def _wmf(b: bytes) -> ImageInfo:
    if b[:4] == b"\xd7\xcd\xc6\x9a":  # placeable header
        left, top, right, bottom, inch = struct.unpack("<hhhhH", b[6:16])
        return ImageInfo("image/x-wmf", "wmf", abs(right - left) or 1, abs(bottom - top) or 1,
                         inch or 1440, inch or 1440, vector=True)
    # no placeable header: no size information; 1 x 1 inch, give width/height
    return ImageInfo("image/x-wmf", "wmf", 96, 96, vector=True)


_SVG_UNITS = {"px": 1.0, "pt": 96 / 72, "pc": 16.0, "mm": 96 / 25.4, "cm": 96 / 2.54,
              "in": 96.0, "": 1.0}


def _svg_length(value: str | None) -> float | None:
    import re

    if not value:
        return None
    m = re.fullmatch(r"\s*([0-9.]+)\s*(px|pt|pc|mm|cm|in|)\s*", value)
    if not m:
        return None  # percentages and the like
    return float(m.group(1)) * _SVG_UNITS[m.group(2)]


def _svg(b: bytes) -> ImageInfo:
    from lxml import etree  # pyright: ignore[reportAttributeAccessIssue]

    from .element import PARSER

    try:
        root = etree.fromstring(b, PARSER)
    except etree.XMLSyntaxError as e:
        raise UnsupportedImage(f"SVG is not well-formed XML: {e}") from None
    if etree.QName(root).localname != "svg":
        raise UnsupportedImage("not an SVG document")
    w, h = _svg_length(root.get("width")), _svg_length(root.get("height"))
    box = (root.get("viewBox") or "").replace(",", " ").split()
    if (w is None or h is None) and len(box) == 4:
        bw, bh = float(box[2]), float(box[3])
        if w is None and h is None:
            w, h = bw, bh
        elif w is None:
            w = h * bw / bh if bh else bw  # type: ignore[operator]
        else:
            h = w * bh / bw if bw else bh
    return ImageInfo("image/svg+xml", "svg", round(w or 300), round(h or 150), vector=True)


def inspect(blob: bytes) -> ImageInfo:
    if blob[:8] == b"\x89PNG\r\n\x1a\n":
        return _png(blob)
    if blob[:3] == b"\xff\xd8\xff":
        return _jpeg(blob)
    if blob[:6] in (b"GIF87a", b"GIF89a"):
        return _gif(blob)
    if blob[:2] == b"BM":
        return _bmp(blob)
    if blob[:4] in (b"II*\x00", b"MM\x00*"):
        return _tiff(blob)
    if blob[:4] == b"\x01\x00\x00\x00" and blob[40:44] == b" EMF":
        return _emf(blob)
    if blob[:4] == b"\xd7\xcd\xc6\x9a" or blob[:6] in (b"\x01\x00\x09\x00\x00\x03",
                                                       b"\x02\x00\x09\x00\x00\x03"):
        return _wmf(blob)
    head = blob[:1024].lstrip()
    if head[:5] in (b"<?xml", b"<svg ", b"<svg>", b"<!DOC") and b"<svg" in head:
        return _svg(blob)
    raise UnsupportedImage("unrecognised image format "
                           "(supported: PNG, JPEG, GIF, BMP, TIFF, EMF, WMF, SVG)")
