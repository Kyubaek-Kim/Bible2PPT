"""Generate ``run_icon.ico`` — a cute, intuitive Bible2PPT app icon.

Concept: "scripture → slides, automatically". A friendly rounded tile holds a
white presentation *slide card*; on it sits a smiling open-book character
(kawaii dot-eyes + smile + rosy cheeks), and a sparkle ✨ hints at automatic
generation. Rendered at 4× and downsampled for crisp edges, then exported as a
multi-resolution ``.ico`` for the Windows executable.

Run: ``python scripts/make_icon.py``.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
SIZE = 512
SS = 4  # supersampling factor
U = 1000  # design space (drawn in a 1000-unit box, then scaled)

# palette — warm, friendly
GRAD_A = (255, 138, 176)   # coral pink
GRAD_B = (255, 175, 120)   # peach
CARD = (255, 255, 255)
CARD_EDGE = (255, 226, 210)
BOOK = (129, 201, 255)     # soft sky blue
BOOK_DK = (94, 173, 234)   # book shade
LINE = (223, 242, 255)
SPINE = (74, 144, 205)
FACE = (60, 72, 96)        # dark slate for eyes/smile
CHEEK = (255, 148, 170)
SPARKLE = (255, 214, 92)   # sunny yellow
SHADOW = (203, 92, 96)


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
            px[x, y] = tuple(int(GRAD_A[i] + (GRAD_B[i] - GRAD_A[i]) * t) for i in range(3))
    return grad


def _sparkle(d: ImageDraw.ImageDraw, cx: float, cy: float, r: float, s: float, color) -> None:
    """A four-point star (diamond with concave sides approximated by a polygon)."""
    k = 0.28 * r
    pts = [
        (cx, cy - r), (cx + k, cy - k), (cx + r, cy), (cx + k, cy + k),
        (cx, cy + r), (cx - k, cy + k), (cx - r, cy), (cx - k, cy - k),
    ]
    d.polygon([(x * s, y * s) for x, y in pts], fill=color)


def _content_layer(size: int) -> Image.Image:
    """Slide card + book character + sparkles, drawn on a transparent layer."""
    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    s = size / U

    def R(x0, y0, x1, y1, rad, **kw):
        d.rounded_rectangle([x0 * s, y0 * s, x1 * s, y1 * s], rad * s, **kw)

    # presentation slide card (landscape 4:3-ish)
    R(150, 235, 850, 720, 55, fill=CARD, outline=CARD_EDGE, width=int(6 * s))
    # little stand under the slide (screen vibe)
    d.rectangle([(492 * s, 720 * s), (508 * s, 772 * s)], fill=CARD_EDGE)
    R(420, 762, 580, 792, 15, fill=CARD_EDGE)

    # --- open-book character on the slide ---
    cx = 500
    top = 350
    bottom = 560
    lift = 40
    left, right = 250, 750

    left_page = [(left, top + lift), (cx, top), (cx, bottom), (left, bottom + lift)]
    right_page = [(right, top + lift), (cx, top), (cx, bottom), (right, bottom + lift)]
    d.polygon([(x * s, y * s) for x, y in left_page], fill=BOOK)
    d.polygon([(x * s, y * s) for x, y in right_page], fill=BOOK_DK)
    d.line([(cx * s, top * s), (cx * s, bottom * s)], fill=SPINE, width=int(9 * s))
    for i in range(2):
        yy = top + 78 + i * 46
        d.line([((left + 55) * s, (yy + lift * 0.55) * s), ((cx - 35) * s, yy * s)], fill=LINE, width=int(11 * s))
        d.line([((cx + 35) * s, yy * s), ((right - 55) * s, (yy + lift * 0.55) * s)], fill=LINE, width=int(11 * s))

    # kawaii face (sits just above the book, centered)
    ey = 315
    for ex in (452, 548):
        d.ellipse([(ex - 15) * s, (ey - 18) * s, (ex + 15) * s, (ey + 18) * s], fill=FACE)
        d.ellipse([(ex - 3) * s, (ey - 14) * s, (ex + 7) * s, (ey - 4) * s], fill=(255, 255, 255))
    # cheeks
    for ex in (417, 583):
        d.ellipse([(ex - 15) * s, (ey + 8) * s, (ex + 15) * s, (ey + 30) * s], fill=CHEEK)
    # smile
    d.arc([475 * s, (ey - 2) * s, 525 * s, (ey + 40) * s], start=15, end=165, fill=FACE, width=int(7 * s))

    # sparkles = "auto / magic"
    _sparkle(d, 735, 300, 74, s, SPARKLE)
    _sparkle(d, 815, 405, 34, s, SPARKLE)
    _sparkle(d, 235, 640, 40, s, SPARKLE)
    return layer


def draw_icon(size: int) -> Image.Image:
    s = size * SS
    tile = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    grad = _diagonal_gradient(s)
    mask = _squircle_mask(s)
    tile.paste(grad, (0, 0), mask)

    content = _content_layer(s)
    # soft drop shadow
    shadow = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    sh = content.split()[3].point(lambda a: int(a * 0.30))
    solid = Image.new("RGBA", (s, s), SHADOW + (255,))
    shadow.paste(solid, (0, int(16 * SS)), sh)
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=11 * SS))
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
