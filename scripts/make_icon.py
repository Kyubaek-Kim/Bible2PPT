"""Generate ``run_icon.ico`` — a simple, legible Bible2PPT app icon.

Concept: "scripture → slides". The icon is deliberately reduced to two bold,
high-contrast shapes so it stays readable even at 16px in the taskbar:

* a large white **open book** (universally reads as scripture / Bible), and
* a small golden **play badge** in the corner (▶ = generate / auto-make),

both sitting on a warm rounded tile. Everything is drawn big and thick — no
thin lines, faces or sparkles that dissolve at small sizes. Rendered at 4× and
downsampled for crisp edges, then exported as a multi-resolution ``.ico``.

Run: ``python scripts/make_icon.py``.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
SIZE = 512
SS = 4  # supersampling factor
U = 1000  # design space (drawn in a 1000-unit box, then scaled)

# palette — warm, friendly, high contrast against a white book
GRAD_A = (255, 122, 89)    # coral
GRAD_B = (255, 176, 74)    # amber
BOOK = (255, 255, 255)     # white pages
BOOK_SHADE = (226, 236, 248)  # faint page shade (right page)
SPINE = (120, 140, 170)    # soft slate spine
LINE = (176, 196, 222)     # text lines on the pages
BADGE = (255, 201, 51)     # golden play badge
BADGE_EDGE = (255, 255, 255)
PLAY = (208, 92, 60)       # play triangle (matches the tile)
SHADOW = (150, 60, 55)


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
    """Big white open book + a golden play badge, on a transparent layer."""
    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    s = size / U

    def P(pts, **kw):
        d.polygon([(x * s, y * s) for x, y in pts], **kw)

    def L(x0, y0, x1, y1, w, fill):
        d.line([(x0 * s, y0 * s), (x1 * s, y1 * s)], fill=fill, width=int(w * s))

    # --- large open book, centred, filling most of the tile ---
    cx = 500
    top = 300        # top of the spine
    crest = 250      # outer top corners sit a touch higher (gentle fan)
    bottom = 690
    drop = 60        # outer bottom corners drop below the spine base
    left, right = 150, 850

    left_page = [(left, crest), (cx, top), (cx, bottom), (left, bottom + drop)]
    right_page = [(right, crest), (cx, top), (cx, bottom), (right, bottom + drop)]
    P(left_page, fill=BOOK)
    P(right_page, fill=BOOK_SHADE)
    # spine
    L(cx, top, cx, bottom, 16, SPINE)

    # a few bold text lines per page (thick enough to survive downscaling)
    for i in range(3):
        yy = top + 70 + i * 74
        sag = 26 - i * 4  # lines follow the page fan
        L(left + 60, yy + sag, cx - 45, yy, 16, LINE)
        L(cx + 45, yy, right - 60, yy + sag, 16, LINE)

    # --- golden play badge (bottom-right): "generate / auto" ---
    bcx, bcy, br = 762, 720, 150
    d.ellipse(
        [(bcx - br) * s, (bcy - br) * s, (bcx + br) * s, (bcy + br) * s],
        fill=BADGE, outline=BADGE_EDGE, width=int(20 * s),
    )
    tri = [(bcx - 46, bcy - 66), (bcx - 46, bcy + 66), (bcx + 70, bcy)]
    P(tri, fill=PLAY)
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
