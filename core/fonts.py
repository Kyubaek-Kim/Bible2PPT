"""Font handling: bundled OFL fonts, system enumeration, preview support.

The PPT default font is a **bundled, freely-redistributable** font (NanumGothic,
OFL) so output is reproducible on any machine — replacing the old hard-coded
Windows-only fonts (맑은 고딕 / 넥슨 풋볼고딕 B) with their silent ``try/except pass``.

The font dropdown lists every family Tkinter can see, but defaults to the bundled
font. Because the bundled font may not be installed system-wide, the live preview
first tries to register it for the current process (Windows ``AddFontResourceEx``);
if that fails it surfaces an install hint (see :func:`ensure_font_available`).
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from . import paths, platform_util


@dataclass(frozen=True)
class BundledFont:
    name: str  # family name used for PPT + display (Korean)
    name_en: str  # Latin family name (matches the font's English name record)
    regular: Path
    bold: Path | None = None


def bundled_fonts() -> list[BundledFont]:
    d = paths.fonts_dir()
    fonts: list[BundledFont] = []
    nanum = d / "NanumGothic-Regular.ttf"
    if nanum.exists():
        fonts.append(
            BundledFont(
                name="나눔고딕",
                name_en="NanumGothic",
                regular=nanum,
                bold=(d / "NanumGothic-Bold.ttf")
                if (d / "NanumGothic-Bold.ttf").exists()
                else None,
            )
        )
    return fonts


def default_font() -> BundledFont | None:
    fonts = bundled_fonts()
    return fonts[0] if fonts else None


def default_font_name() -> str:
    f = default_font()
    return f.name if f else "나눔고딕"


def register_bundled_fonts() -> list[tuple[str, bool]]:
    """Register bundled fonts for the current GUI session (best effort)."""
    results: list[tuple[str, bool]] = []
    for f in bundled_fonts():
        ok = platform_util.register_font(f.regular)
        if f.bold:
            platform_util.register_font(f.bold)
        results.append((f.name, ok))
    return results


def system_font_families(tk_root) -> list[str]:
    """Sorted, de-duplicated list of families Tkinter can render."""
    import tkinter.font as tkfont

    fams = {f for f in tkfont.families(tk_root) if not f.startswith("@")}
    return sorted(fams)


def font_dropdown_values(tk_root) -> list[str]:
    """Bundled fonts first (the defaults), then the rest of the system families."""
    bundled = [f.name for f in bundled_fonts()]
    system = system_font_families(tk_root)
    ordered = list(bundled)
    for fam in system:
        if fam not in ordered:
            ordered.append(fam)
    return ordered


def ensure_font_available(name: str, tk_root) -> tuple[bool, str]:
    """Ensure ``name`` renders in the preview.

    Returns ``(available, hint)``. When a bundled font is not yet visible to
    Tkinter, attempts a per-process registration; if that fails, ``hint``
    explains how to install it so the preview matches the final PPT.
    """
    if name in system_font_families(tk_root):
        return True, ""
    for f in bundled_fonts():
        if name in (f.name, f.name_en):
            ok = platform_util.register_font(f.regular)
            if f.bold:
                platform_util.register_font(f.bold)
            if ok and name in system_font_families(tk_root):
                return True, ""
            return False, platform_util.font_install_hint(f.regular)
    return name in system_font_families(tk_root), ""
