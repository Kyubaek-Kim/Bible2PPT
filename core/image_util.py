"""Background image handling (Pillow): aspect-fit crop with a confirm step.

When a background's aspect ratio differs from the chosen slide ratio, the image
is *cover*-cropped so it fills the slide with no letterboxing. Before doing so
the app tells the user exactly how much will be cut — in **pixels and cm** — and
waits for confirmation (the UI shows the message from :func:`plan_crop`).
"""
from __future__ import annotations

import shutil
import time
from dataclasses import dataclass
from itertools import count
from pathlib import Path

from PIL import Image

from . import paths

# EMU per centimetre (python-pptx / OOXML unit).
EMU_PER_CM = 360000


@dataclass(frozen=True)
class CropPlan:
    needs_crop: bool
    axis: str  # "vertical" (top/bottom) | "horizontal" (left/right) | ""
    crop_px: int  # total pixels removed along the cropped axis
    crop_cm: float  # equivalent cm relative to the final slide size
    box: tuple[int, int, int, int]  # left, upper, right, lower (cover box)
    image_size: tuple[int, int]


def plan_crop(
    image_path: str | Path,
    slide_w_cm: float,
    slide_h_cm: float,
) -> CropPlan:
    """Compute the centered cover-crop needed to fit ``slide_w_cm:slide_h_cm``."""
    with Image.open(image_path) as im:
        w, h = im.size
    target_ratio = slide_w_cm / slide_h_cm
    img_ratio = w / h

    if abs(img_ratio - target_ratio) < 1e-6:
        return CropPlan(False, "", 0, 0.0, (0, 0, w, h), (w, h))

    if img_ratio > target_ratio:
        # too wide -> crop left/right
        new_w = int(round(target_ratio * h))
        crop_px = w - new_w
        left = crop_px // 2
        box = (left, 0, left + new_w, h)
        crop_cm = (crop_px / w) * slide_w_cm
        axis = "horizontal"
    else:
        # too tall -> crop top/bottom
        new_h = int(round(w / target_ratio))
        crop_px = h - new_h
        upper = crop_px // 2
        box = (0, upper, w, upper + new_h)
        crop_cm = (crop_px / h) * slide_h_cm
        axis = "vertical"

    return CropPlan(True, axis, crop_px, crop_cm, box, (w, h))


def apply_crop(image_path: str | Path, plan: CropPlan, out_path: str | Path) -> Path:
    """Write the cover-cropped image to ``out_path`` (no-op copy if not needed)."""
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(image_path) as im:
        im = im.convert("RGB") if im.mode not in ("RGB", "RGBA") else im
        if plan.needs_crop:
            im = im.crop(plan.box)
        im.save(out)
    return out


def add_to_history(image_path: str | Path) -> Path:
    """Copy a newly attached background into the history folder (once).

    Called only when the user *attaches* an image — never when selecting an
    already-registered one — so the list can't grow duplicate timestamped
    copies. A numeric suffix guards against same-second collisions.
    """
    src = Path(image_path)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    hist = paths.background_history_dir()
    dst = hist / f"{stamp}_{src.name}"
    for i in count(1):
        if not dst.exists():
            break
        dst = hist / f"{stamp}-{i}_{src.name}"
    shutil.copy2(src, dst)
    return dst


def delete_background(stored_path: str | Path) -> None:
    """Delete a registered background file (and any cached aspect crops of it).

    Only removes files under the app's background folders; a missing file is
    ignored so deletion is idempotent.
    """
    stored = Path(stored_path)
    hist = paths.background_history_dir().resolve()
    try:
        within_hist = stored.resolve().parent == hist
    except OSError:
        within_hist = False
    if within_hist:
        stored.unlink(missing_ok=True)
    # drop cached "<aspect>_<name>" crops derived from this original
    cache = paths.background_cache_dir()
    for f in cache.glob(f"*_{stored.name}"):
        f.unlink(missing_ok=True)


def emu_to_cm(emu: int) -> float:
    return emu / EMU_PER_CM
