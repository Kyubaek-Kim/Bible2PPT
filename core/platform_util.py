"""OS-dependent behaviour, isolated behind a small helper API.

Every place in the app that would otherwise call ``os.startfile`` or shell out
to an OS command routes through here, so porting to macOS/Linux means editing
this one module. Only the Windows branches are implemented for this release;
other platforms have working best-effort or clearly-marked stub branches.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def open_folder(path: os.PathLike | str) -> None:
    """Open ``path`` (a folder, or a file's containing folder) in the OS file browser."""
    p = Path(path)
    target = p if p.is_dir() else p.parent
    if sys.platform.startswith("win"):
        os.startfile(str(target))  # type: ignore[attr-defined]  # Windows only
    elif sys.platform == "darwin":
        # macOS port target.
        subprocess.run(["open", str(target)], check=False)
    else:
        # Linux / other (development convenience).
        subprocess.run(["xdg-open", str(target)], check=False)


def reveal_file(path: os.PathLike | str) -> None:
    """Reveal a specific file, selecting it when the OS supports it."""
    p = Path(path)
    if sys.platform.startswith("win"):
        subprocess.run(["explorer", "/select,", str(p)], check=False)
    elif sys.platform == "darwin":
        subprocess.run(["open", "-R", str(p)], check=False)
    else:
        open_folder(p)


def register_font(font_path: os.PathLike | str) -> bool:
    """Make a font file usable by the current GUI session (best effort).

    Needed so the bundled font renders in the Tkinter preview even when it is
    not installed system-wide. Returns ``True`` if registration was attempted
    successfully.

    * Windows: ``AddFontResourceExW`` (process-private, no admin required).
    * macOS/Linux: no reliable per-process API; returns ``False`` so callers
      fall back to an install prompt.
    """
    p = Path(font_path)
    if not p.exists():
        return False
    if sys.platform.startswith("win"):
        try:  # pragma: no cover - Windows only
            import ctypes

            FR_PRIVATE = 0x10
            added = ctypes.windll.gdi32.AddFontResourceExW(str(p), FR_PRIVATE, 0)
            return bool(added)
        except Exception:
            return False
    # macOS/Linux port target — no per-process registration; prompt to install.
    return False


def font_install_hint(font_path: os.PathLike | str) -> str:
    """Human-readable instruction for installing a bundled font when auto-register fails."""
    name = Path(font_path).name
    if sys.platform.startswith("win"):
        return f"'{name}' 폰트를 더블클릭 후 [설치]를 눌러 시스템에 설치하면 미리보기가 정확해집니다."
    if sys.platform == "darwin":
        return f"'{name}' 폰트를 Font Book으로 열어 설치하면 미리보기가 정확해집니다."
    return f"'{name}' 폰트를 시스템 폰트 폴더에 설치하면 미리보기가 정확해집니다."
