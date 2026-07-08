"""Generate ``run_icon.ico`` / ``run_icon.png`` from the bundled source artwork.

The app icon is the user-supplied artwork at ``assets/icon_source.png`` (a Bible
with a gold cross, a "convert" arrow, and a "PPT" tag). This script normalises it
to a square RGBA master and exports a multi-resolution Windows ``.ico`` plus a
PNG, so the icon builds reproducibly without relying on system fonts.

Run: ``python scripts/make_icon.py``.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "assets" / "icon_source.png"
MASTER_SIZE = 256
ICO_SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def build_master(size: int = MASTER_SIZE) -> Image.Image:
    """Load the source artwork and fit it, centred, onto a transparent square."""
    src = Image.open(SOURCE).convert("RGBA")
    w, h = src.size
    scale = size / max(w, h)
    resized = src.resize((max(1, round(w * scale)), max(1, round(h * scale))), Image.LANCZOS)
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    canvas.paste(resized, ((size - resized.width) // 2, (size - resized.height) // 2), resized)
    return canvas


def main() -> None:
    master = build_master()
    master.save(ROOT / "run_icon.png")
    master.save(ROOT / "run_icon.ico", sizes=ICO_SIZES)
    print(f"wrote {ROOT / 'run_icon.ico'} and run_icon.png from {SOURCE}")


if __name__ == "__main__":
    main()
