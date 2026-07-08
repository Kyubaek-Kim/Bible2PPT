"""Font handling: a small curated font list, bundling, and preview support.

The font dropdown exposes a **fixed, curated** set of families rather than every
system face, so a non-technical user picks from options that are known to look
good and to be available on the Windows target:

* **맑은 고딕** — ships with Windows (not bundled/redistributed here);
* **나눔스퀘어 볼드** — bundled (OFL), the default body font;
* **나눔고딕** — bundled (OFL).

Bundled fonts (NanumSquare / NanumGothic) are freely redistributable, so output
is reproducible. Because a bundled font may not be installed system-wide, the
live preview first tries to register it for the current process (Windows
``AddFontResourceEx``); if that fails it surfaces an install hint (see
:func:`ensure_font_available`).
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from . import paths, platform_util


@dataclass(frozen=True)
class FontChoice:
    label: str  # shown in the dropdown and stored in settings
    typeface: str  # family name used for the PPT run + preview
    bold: bool  # whether the face is inherently bold (e.g. NanumSquare Bold)
    bundled: Path | None = None  # bundled TTF, or None for OS-provided fonts


def _fonts_dir() -> Path:
    return paths.fonts_dir()


def curated_fonts() -> list[FontChoice]:
    """The fixed dropdown list, in display order (default first)."""
    d = _fonts_dir()
    return [
        FontChoice("나눔스퀘어 볼드", "나눔스퀘어", True, d / "NanumSquareB.ttf"),
        FontChoice("맑은 고딕", "맑은 고딕", False, None),
        FontChoice("나눔고딕", "나눔고딕", False, d / "NanumGothic-Regular.ttf"),
    ]


DEFAULT_FONT_LABEL = "나눔스퀘어 볼드"


def resolve(label: str) -> FontChoice:
    """Map a stored label to its :class:`FontChoice` (falls back to default)."""
    fonts = curated_fonts()
    for f in fonts:
        if label in (f.label, f.typeface):
            return f
    return fonts[0]


def default_font_name() -> str:
    return DEFAULT_FONT_LABEL


def register_bundled_fonts() -> list[tuple[str, bool]]:
    """Register bundled fonts for the current GUI session (best effort)."""
    results: list[tuple[str, bool]] = []
    for f in curated_fonts():
        if f.bundled and f.bundled.exists():
            ok = platform_util.register_font(f.bundled)
            results.append((f.label, ok))
    return results


def system_font_families(tk_root) -> list[str]:
    """Sorted, de-duplicated list of families Tkinter can render."""
    import tkinter.font as tkfont

    fams = {f for f in tkfont.families(tk_root) if not f.startswith("@")}
    return sorted(fams)


def font_dropdown_values(tk_root=None) -> list[str]:
    """The curated dropdown labels, in display order."""
    return [f.label for f in curated_fonts()]


def ensure_font_available(label: str, tk_root) -> tuple[bool, str]:
    """Ensure ``label`` renders in the preview.

    Returns ``(available, hint)``. When a bundled font is not yet visible to
    Tkinter, attempts a per-process registration; if that fails, ``hint``
    explains how to install it so the preview matches the final PPT.
    """
    choice = resolve(label)
    families = system_font_families(tk_root)
    for name in (choice.typeface, choice.label):
        if name in families:
            return True, ""
    if choice.bundled and choice.bundled.exists():
        platform_util.register_font(choice.bundled)
        families = system_font_families(tk_root)
        for name in (choice.typeface, choice.label):
            if name in families:
                return True, ""
        return False, platform_util.font_install_hint(choice.bundled)
    # OS-provided font (e.g. 맑은 고딕): available on the Windows target only.
    return choice.typeface in families, ""
