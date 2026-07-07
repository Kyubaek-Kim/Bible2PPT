"""Generate ``run_icon.ico`` — a modern, minimal Bible2PPT app icon.

Design: a rounded *squircle* tile with a diagonal indigo→violet gradient, a clean
duotone open-book glyph in white with a soft shadow, and a single slim accent
bookmark. Rendered at 4× and downsampled for crisp anti-aliased edges, then
exported as a multi-resolution ``.ico`` for the Windows executable.

Run: ``python scripts/make_icon.py``.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
SIZE = 512
SS = 4  # supersampling factor

# palette
GRAD_A = (99, 102, 241)    # indigo-500
GRAD_B = (139, 92, 246)    # violet-500
GRAD_C = (217, 70, 239)    # fuchsia-500 (corner accent)
PAGE = (255, 255, 255)
PAGE_TINT = (226, 232, 255)  # cool lavender for the duotone shade
ACCENT = (253, 186, 116)   # warm amber bookmark
SHADOW = (49, 46, 129)     # deep indigo shadow


def _squircle_mask(size: int, n: float = 4.0) -> Image.Image:
    """Superellipse (squircle) alpha mask — softer than a rounded rect."""
    mask = Image.new("L", (size, size), 0)
    px = mask.load()
    r = size / 2.0
    for y in range(size):
        for x in range(size):
            nx = abs((x + 0.5 - r) / r)
            ny = abs((y + 0.5 - r) / r)
            if nx ** n + ny ** n <= 1.0:
                px[x, y] = 255
    return mask


def _diagonal_gradient(size: int) -> Image.Image:
    grad = Image.new("RGB", (size, size))
    px = grad.load()
    for y in range(size):
        for x in range(size):
            t = (x + y) / (2 * size)          # main indigo->violet sweep
            c = (x / size) * (y / size)       # subtle fuchsia in one corner
            r = GRAD_A[0] + (GRAD_B[0] - GRAD_A[0]) * t + (GRAD_C[0] - GRAD_B[0]) * c * 0.6
            g = GRAD_A[1] + (GRAD_B[1] - GRAD_A[1]) * t + (GRAD_C[1] - GRAD_B[1]) * c * 0.6
            b = GRAD_A[2] + (GRAD_B[2] - GRAD_A[2]) * t + (GRAD_C[2] - GRAD_B[2]) * c * 0.6
            px[x, y] = (int(r), int(g), int(b))
    return grad


def _book_layer(size: int) -> Image.Image:
    """Open-book glyph on a transparent layer, drawn in a centered 1000-unit box."""
    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    u = size / 1000.0

    def P(x, y):
        return (x * u, y * u)

    cx = 500
    top = 330
    bottom = 690
    lift = 70          # outer edges sit lower than the spine
    left = 210
    right = 790

    # left / right pages as smooth quads
    left_page = [P(left, top + lift), P(cx, top), P(cx, bottom), P(left, bottom + lift)]
    right_page = [P(right, top + lift), P(cx, top), P(cx, bottom), P(right, bottom + lift)]

    # duotone: right page slightly tinted for depth
    d.polygon(left_page, fill=PAGE)
    d.polygon(right_page, fill=PAGE_TINT)

    # spine
    d.line([P(cx, top), P(cx, bottom)], fill=(148, 130, 220), width=int(8 * u))

    # minimal text strokes (rounded)
    for i in range(3):
        y = top + 70 + i * 62
        d.line([P(left + 70, y + lift * 0.5), P(cx - 45, y)], fill=PAGE_TINT, width=int(12 * u))
        d.line([P(cx + 45, y), P(right - 70, y + lift * 0.5)], fill=(200, 208, 235), width=int(12 * u))

    # slim accent bookmark hanging from the right page
    bx = 640
    d.polygon(
        [P(bx - 26, top + 8), P(bx + 26, top + 4), P(bx + 26, bottom + 118),
         P(bx, bottom + 78), P(bx - 26, bottom + 118)],
        fill=ACCENT,
    )
    return layer


def draw_icon(size: int) -> Image.Image:
    s = size * SS
    tile = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    grad = _diagonal_gradient(s)
    mask = _squircle_mask(s)
    tile.paste(grad, (0, 0), mask)

    # soft drop shadow for the book
    book = _book_layer(s)
    shadow = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    sh = book.split()[3].point(lambda a: int(a * 0.35))
    solid = Image.new("RGBA", (s, s), SHADOW + (255,))
    shadow.paste(solid, (0, int(14 * SS)), sh)
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=10 * SS))
    tile = Image.alpha_composite(tile, shadow)
    tile = Image.alpha_composite(tile, book)

    # re-apply squircle mask so the shadow doesn't bleed past the tile
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
