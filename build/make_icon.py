#!/usr/bin/env python3
"""
make_icon.py — generate the app icon (a simple CBMI mark) as .icns (macOS) and
.ico (Windows) with Pillow, so no binary asset needs committing. Pure-Python,
runs in CI.

Usage: python3 build/make_icon.py
Outputs: build/assets/icon.icns, build/assets/icon.ico, build/assets/icon.png
"""
from __future__ import annotations

import struct
from pathlib import Path

from PIL import Image, ImageDraw

ASSETS = Path(__file__).resolve().parent / "assets"
BG = (11, 12, 16)
ACCENT = (0, 170, 235)
RING = (46, 52, 64)


def _base(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    m = size // 12
    # rounded dark tile
    d.rounded_rectangle([m, m, size - m, size - m], radius=size // 6, fill=BG)
    # confidence ring (open arc) + core dot — echoes the Mission Readiness gauge
    cx = cy = size / 2
    r = size * 0.30
    w = max(2, size // 22)
    d.arc([cx - r, cy - r, cx + r, cy + r], start=0, end=360, fill=RING, width=w)
    d.arc([cx - r, cy - r, cx + r, cy + r], start=-90, end=200, fill=ACCENT, width=w)
    cr = size * 0.10
    d.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=ACCENT)
    return img


def save_png() -> Path:
    ASSETS.mkdir(parents=True, exist_ok=True)
    p = ASSETS / "icon.png"
    _base(512).save(p)
    return p


def save_ico() -> Path:
    p = ASSETS / "icon.ico"
    sizes = [16, 24, 32, 48, 64, 128, 256]
    _base(256).save(p, sizes=[(s, s) for s in sizes])
    return p


def save_icns() -> Path:
    """Minimal .icns writer (PNG-backed OSTypes). Avoids needing iconutil, so
    it runs on Linux CI too."""
    p = ASSETS / "icon.icns"
    ostypes = {"icp4": 16, "icp5": 32, "icp6": 64, "ic07": 128,
               "ic08": 256, "ic09": 512, "ic10": 1024, "ic11": 32,
               "ic12": 64, "ic13": 256, "ic14": 512}
    entries = []
    import io
    for ost, sz in ostypes.items():
        buf = io.BytesIO()
        _base(sz).save(buf, format="PNG")
        data = buf.getvalue()
        entries.append(ost.encode("ascii") + struct.pack(">I", len(data) + 8) + data)
    body = b"".join(entries)
    p.write_bytes(b"icns" + struct.pack(">I", len(body) + 8) + body)
    return p


def main() -> int:
    print("wrote", save_png())
    print("wrote", save_ico())
    print("wrote", save_icns())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
