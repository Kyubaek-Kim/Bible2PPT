"""Generate ``run_icon.ico`` — a legible Bible2PPT app icon.

Concept: "Bible → PPT". Two bold, high-contrast elements so it reads even small:

* a **Bible** — a deep-red closed book with a gold **cross** on the cover
  (unmistakably scripture), and
* a white **"PPT" tag** overlapping its corner (the output format).

Text is rendered with the bundled ``NanumGothic-Bold`` font so the icon builds
reproducibly without relying on system fonts. Rendered at 4× and downsampled
for crisp edges, then exported as a multi-resolution ``.ico``.

Run: ``python scripts/make_icon.py``.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parent.parent
FONT_PATH = ROOT / "data" / "fonts" / "NanumGothic-Bold.ttf"
SIZE = 512
SS = 4  # supersampling factor
U = 1000  # design space (drawn in a 1000-unit box, then scaled)

# palette
GRAD_A = (255, 138, 101)   # coral
GRAD_B = (255, 187, 92)    # amber
COVER = (155, 38, 52)      # deep burgundy Bible cover
COVER_DK = (120, 26, 40)   # spine / shade
PAGES = (255, 246, 232)    # cream page block
CROSS = (247, 201, 72)     # gold cross
TAG = (255, 255, 255)      # white "PPT" tag
TAG_TEXT = (198, 58, 66)   # tag text (matches cover)
SHADOW = (120, 40, 45)


def _squircle_mask(size: int, n: float = 4.0) -> Image.Image:
    mask = Image.new("L", (size, size), 0)
    px = mask.load()
    r = size / 2.0
    for y in range(size):
        for x in range(size):
            nx = abs((x + 0.5 - r) / r)
            ny = abs((y + 0.5 - r) / r)
            if nx**n + ny**n <= 1.0:
                px[x, y] = 255
    return mask


def _diagonal_gradient(size: int) -> Image.Image:
    grad = Image.new("RGB", (size, size))
    px = grad.load()
    for y in range(size):
        for x in range(size):
            t = (x + y) / (2 * size)
            px[x, y] = tuple(
                int(GRAD_A[i] + (GRAD_B[i] - GRAD_A[i]) * t) for i in range(3)
            )
    return grad


def _content_layer(size: int) -> Image.Image:
    """A burgundy Bible (gold cross on the cover) + a white "PPT" tag."""
    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    s = size / U

    def R(x0, y0, x1, y1, rad, **kw):
        d.rounded_rectangle([x0 * s, y0 * s, x1 * s, y1 * s], rad * s, **kw)

    # --- Bible (closed book, cover facing us) ---
    # cream page block peeking out along the right & bottom edges
    R(210, 205, 745, 795, 34, fill=PAGES)
    # cover, sitting slightly up-left of the pages
    R(190, 185, 715, 770, 40, fill=COVER)
    # darker spine band down the left edge
    R(190, 185, 268, 770, 40, fill=COVER_DK)
    d.rectangle([248 * s, 185 * s, 268 * s, 770 * s], fill=COVER_DK)

    # --- gold cross on the cover ---
    ccx = 486  # cross centre x (centred on the cover face, right of the spine)
    v_top, v_bot = 300, 660
    h_y = 400
    bar = 62
    R(ccx - bar / 2, v_top, ccx + bar / 2, v_bot, bar / 2, fill=CROSS)  # vertical
    R(ccx - 150, h_y - bar / 2, ccx + 150, h_y + bar / 2, bar / 2, fill=CROSS)  # cross-bar

    # --- white "PPT" tag overlapping the lower-right corner ---
    tx0, ty0, tx1, ty1 = 545, 610, 880, 792
    R(tx0, ty0, tx1, ty1, 30, fill=TAG)
    font = ImageFont.truetype(str(FONT_PATH), int(118 * s))
    cx_tag = (tx0 + tx1) / 2 * s
    cy_tag = (ty0 + ty1) / 2 * s
    d.text((cx_tag, cy_tag), "PPT", font=font, fill=TAG_TEXT, anchor="mm")
    return layer


def draw_icon(size: int) -> Image.Image:
    s = size * SS
    tile = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    grad = _diagonal_gradient(s)
    mask = _squircle_mask(s)
    tile.paste(grad, (0, 0), mask)

    content = _content_layer(s)
    # soft drop shadow beneath the white shapes for depth
    shadow = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    sh = content.split()[3].point(lambda a: int(a * 0.30))
    solid = Image.new("RGBA", (s, s), SHADOW + (255,))
    shadow.paste(solid, (0, int(14 * SS)), sh)
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=10 * SS))
    tile = Image.alpha_composite(tile, shadow)
    tile = Image.alpha_composite(tile, content)

    out = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    out.paste(tile, (0, 0), mask)
    return out.resize((size, size), Image.LANCZOS)


def main() -> None:
    master = draw_icon(SIZE)
    master.save(ROOT / "run_icon.png")
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    master.save(ROOT / "run_icon.ico", sizes=sizes)
    print(f"wrote {ROOT / 'run_icon.ico'} and run_icon.png")


if __name__ == "__main__":
    main()
