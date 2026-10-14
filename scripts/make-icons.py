#!/usr/bin/env python3
"""Generate all add-in icons from one source logo PNG (transparent background preferred).

  python3 scripts/make-icons.py path/to/interactio-logo.png

Outputs into docs/assets/:
  btn-{16,20,24,25,32,40,48,50,64,75,80,96,144}.png  grayscale (#919191) button icons, as Microsoft requires
  icon-{64,128}.png                                   colour icons for the add-in list / store

Run without an argument to (re)create neutral placeholder icons.
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageOps

OUT = Path(__file__).resolve().parent.parent / "docs" / "assets"
GREY = (0x91, 0x91, 0x91)
MAGENTA = (0xBD, 0x31, 0xAD)
BTN_SIZES = [16, 20, 24, 25, 32, 40, 48, 50, 64, 75, 80, 96, 144]
COLOR_SIZES = [64, 128]


def placeholder(size=512):
    """Neutral ring + dot mark. Replace with the real Interactio logo."""
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    w = size // 9
    d.ellipse([w, w, size - w, size - w], outline=MAGENTA + (255,), width=w)
    r = size // 7
    c = size // 2
    d.ellipse([c - r, c - r, c + r, c + r], fill=MAGENTA + (255,))
    return im


def square(im):
    im = im.convert("RGBA")
    bbox = im.getchannel("A").getbbox() or (0, 0, *im.size)
    im = im.crop(bbox)
    side = max(im.size)
    canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    canvas.paste(im, ((side - im.width) // 2, (side - im.height) // 2))
    return canvas


def alpha_from_source(im):
    """Mask of the mark: uses alpha if present, else treats near-white as background."""
    a = im.getchannel("A")
    if a.getextrema()[0] < 250:
        return a
    lum = ImageOps.grayscale(im.convert("RGB"))
    return ImageOps.invert(lum).point(lambda v: 255 if v > 40 else int(v * 6))


def main():
    src = square(Image.open(sys.argv[1]) if len(sys.argv) > 1 else placeholder())
    mask = alpha_from_source(src)
    OUT.mkdir(parents=True, exist_ok=True)
    for s in BTN_SIZES:
        pad = max(1, round(s * 0.06))
        inner = s - 2 * pad
        m = mask.resize((inner, inner), Image.LANCZOS)
        tile = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        tile.paste(Image.new("RGBA", (inner, inner), GREY + (255,)), (pad, pad), m)
        tile.save(OUT / f"btn-{s}.png", optimize=True)
    for s in COLOR_SIZES:
        pad = round(s * 0.08)
        inner = s - 2 * pad
        tile = Image.new("RGBA", (s, s), (255, 255, 255, 255))
        logo = src.resize((inner, inner), Image.LANCZOS)
        tile.paste(logo, (pad, pad), logo)
        tile.save(OUT / f"icon-{s}.png", optimize=True)
    print(f"Wrote {len(BTN_SIZES) + len(COLOR_SIZES)} icons to {OUT}")


if __name__ == "__main__":
    main()
