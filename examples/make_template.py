#!/usr/bin/env python3
"""Generate example_brands.xlsx — the exact sheet layout Sheet2Anki expects.

The two placeholder "logos" are simple PNGs drawn with the standard library
(no Pillow needed). Replace them with real photos, then run:

    python3 scripts/feishu_to_anki.py --xlsx examples/example_brands.xlsx \
        -o brands.apkg
"""

import os
import struct
import zlib

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font

HEADERS = [
    "Deck1", "Deck2", "Deck3", "Note Type", "Front", "Back",
    "Fields1_英文名", "Fields2_中文名", "Fields3_品牌级次", "Tags1", "Tags2",
]

ROWS = [
    ["零售", "女装", "中淑装", "Basic", None,
     "入卡旗下实体女装品牌，主打松弛轻熟风，面向 25–40 岁女性。",
     "DULI", "读丽", "C", "韩系", "低价"],
    ["零售", "运动户外", "国际零售", "Basic", None,
     "鬼冢虎（Onitsuka Tiger）1949 年创立于日本神户，经典与街头融合的鞋履品牌。",
     "Onitsuka Tiger", "鬼冢虎", "D", "潮牌", "高端"],
]


def make_png(path, rgb, w=320, h=220):
    """Write a solid-color PNG with a lighter inner rectangle (placeholder)."""
    light = tuple(min(255, c + 60) for c in rgb)

    def pixel(x, y):
        border = 16
        if border <= x < w - border and border <= y < h - border:
            return light
        return rgb

    raw = b""
    for y in range(h):
        raw += b"\x00" + b"".join(bytes(pixel(x, y)) for x in range(w))

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", ihdr)
           + chunk(b"IDAT", zlib.compress(raw, 9))
           + chunk(b"IEND", b""))
    with open(path, "wb") as fh:
        fh.write(png)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    logo1 = os.path.join(here, "_logo1.png")
    logo2 = os.path.join(here, "_logo2.png")
    make_png(logo1, (91, 122, 235))
    make_png(logo2, (235, 122, 91))

    wb = Workbook()
    ws = wb.active
    ws.title = "Anki卡片"
    ws.append(HEADERS)
    for row in ROWS:
        ws.append(row)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for anchor, logo in (("E2", logo1), ("E3", logo2)):
        img = XLImage(logo)
        img.anchor = anchor
        ws.add_image(img)

    widths = {"A": 10, "B": 12, "C": 12, "D": 10, "E": 26,
              "F": 60, "G": 18, "H": 12, "I": 12, "J": 10, "K": 10}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    out = os.path.join(here, "example_brands.xlsx")
    wb.save(out)
    os.remove(logo1)
    os.remove(logo2)
    print(f"created {out}")


if __name__ == "__main__":
    main()
