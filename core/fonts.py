"""Font handling: a small curated font list, bundling, and preview support.

The font dropdown exposes a **fixed, curated** set of families rather than every
system face, so a non-technical user picks from options that are known to look
good and to be available on the Windows target:

* **나눔스퀘어 Bold / 나눔스퀘어** — bundled (free), the default body font is the Bold;
* **나눔고딕 / 나눔고딕 Bold** — bundled (OFL);
* **맑은 고딕 / 굴림 / 돋움 / 바탕 / 궁서** — common Windows-provided families.

Weight handling is subtle and is what :attr:`FontChoice.needs_bold_bit` encodes:

* **나눔고딕** ships Regular + Bold members under one family (``나눔고딕``), so its
  Bold is reached by *setting the bold bit* on the family;
* **나눔스퀘어 Bold** is a *standalone* family whose name already carries the
  weight (Windows lists each NanumSquare weight as its own family), so it must
  be selected by name **without** an extra bold bit — otherwise the weight is
  doubled (fake-bold).

Bundled fonts are freely redistributable, so output is reproducible. Because a
bundled font may not be installed system-wide, the live preview first tries to
register it for the current process (Windows ``AddFontResourceEx``); if that
fails it surfaces an install hint (see :func:`ensure_font_available`).
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from . import paths, platform_util


@dataclass(frozen=True)
class FontChoice:
    label: str  # shown in the dropdown and stored in settings
    typeface: str  # exact family name used for the PPT run + preview
    bold: bool  # the face renders bold-weight (drives defaults / preview weight)
    needs_bold_bit: bool  # must set the bold attribute to reach this weight
    bundled: Path | None = None  # bundled TTF, or None for OS-provided fonts


def _fonts_dir() -> Path:
    return paths.fonts_dir()


def curated_fonts() -> list[FontChoice]:
    """The fixed dropdown list, in display order (default first).

    Bundled faces (OFL) render reproducibly everywhere; the remaining entries
    are common Windows-provided families (they degrade to a fallback in the
    preview when not installed, e.g. on the build machine)."""
    d = _fonts_dir()
    return [
        # NanumSquare Bold is a standalone family (name carries the weight):
        # selected by name, no bold bit.
        FontChoice("나눔스퀘어 Bold", "나눔스퀘어 Bold", True, False, d / "NanumSquareB.ttf"),
        FontChoice("나눔스퀘어", "나눔스퀘어", False, False, d / "NanumSquareR.ttf"),
        # NanumGothic Regular + Bold share one family; Bold needs the bold bit.
        FontChoice("나눔고딕", "나눔고딕", False, False, d / "NanumGothic-Regular.ttf"),
        FontChoice("나눔고딕 Bold", "나눔고딕", True, True, d / "NanumGothic-Bold.ttf"),
        FontChoice("맑은 고딕", "맑은 고딕", False, False, None),
        FontChoice("굴림", "굴림", False, False, None),
        FontChoice("돋움", "돋움", False, False, None),
        FontChoice("바탕", "바탕", False, False, None),
        FontChoice("궁서", "궁서", False, False, None),
    ]


DEFAULT_FONT_LABEL = "나눔스퀘어 Bold"


def resolve(label: str) -> FontChoice:
    """Map a stored label to its :class:`FontChoice` (falls back to default)."""
    fonts = curated_fonts()
    for f in fonts:
        if label in (f.label, f.typeface):
            return f
    return fonts[0]


def default_font_name() -> str:
    return DEFAULT_FONT_LABEL


def run_bold(choice: FontChoice, user_bold: bool) -> bool:
    """Whether a run using ``choice`` should set the bold attribute.

    * a Bold member of a RIBBI family (나눔고딕 Bold) always needs the bit;
    * a standalone bold-named family (나눔스퀘어 Bold) already renders bold, so
      the bit is left off to avoid fake-bold doubling;
    * a regular face gets the bit only when the user explicitly asks for bold.
    """
    if choice.needs_bold_bit:
        return True
    if choice.bold:
        return False
    return user_bold


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
