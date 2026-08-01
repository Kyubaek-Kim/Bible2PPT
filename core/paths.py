"""OS-independent path & resource resolution.

Two distinct roots:

* **Bundle resources** (read-only): the app's shipped assets — canonical Bible
  JSON, i18n tables, versification maps, bundled fonts, default background.
  When frozen by PyInstaller these live under ``sys._MEIPASS``; in a source
  checkout they live next to the repository root.
* **User data** (writable): settings, generated output, imported Bibles,
  background history. These live in the per-OS standard application-data
  folder so they survive app upgrades and never touch the read-only bundle.

Only Windows is implemented for this release. macOS branches are present as
stubs (see :func:`_user_data_root` / :func:`default_output_dir`) so the port is
a matter of filling them in rather than restructuring.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "Bible2PPT"


# --------------------------------------------------------------------------- #
# Bundle resources (read-only)
# --------------------------------------------------------------------------- #
def _bundle_root() -> Path:
    """Root under which shipped assets live.

    PyInstaller unpacks data files into ``sys._MEIPASS`` at runtime; from a
    source checkout we fall back to the repository root (parent of ``core``).
    """
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parent.parent


def resource_path(*parts: str) -> Path:
    """Absolute path to a bundled resource, e.g. ``resource_path('data', 'canon.json')``."""
    return _bundle_root().joinpath(*parts)


def data_dir() -> Path:
    return resource_path("data")


def bibles_dir() -> Path:
    """Bundled (shipped) translations."""
    return data_dir() / "bibles"


def i18n_dir() -> Path:
    return data_dir() / "i18n"


def versification_dir() -> Path:
    return data_dir() / "versification"


def fonts_dir() -> Path:
    return data_dir() / "fonts"


def canon_file() -> Path:
    return data_dir() / "canon.json"


def default_background() -> Path:
    return data_dir() / "ppt배경.png"


# --------------------------------------------------------------------------- #
# User data (writable, per-OS standard location)
# --------------------------------------------------------------------------- #
def _user_data_root() -> Path:
    """Per-OS writable application-data directory.

    * Windows: ``%APPDATA%\\Bible2PPT``
    * macOS (stub/port target): ``~/Library/Application Support/Bible2PPT``
    * Other (Linux/dev): ``$XDG_DATA_HOME/Bible2PPT`` or ``~/.local/share/Bible2PPT``
    """
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        return Path(base) / APP_NAME
    if sys.platform == "darwin":
        # macOS port target — matches Apple's Application Support convention.
        return Path.home() / "Library" / "Application Support" / APP_NAME
    # Linux / other (development fallback).
    base = os.environ.get("XDG_DATA_HOME") or str(Path.home() / ".local" / "share")
    return Path(base) / APP_NAME


def user_data_dir() -> Path:
    p = _user_data_root()
    p.mkdir(parents=True, exist_ok=True)
    return p


def settings_file() -> Path:
    return user_data_dir() / "settings.json"


def user_bibles_dir() -> Path:
    """Translations registered by the user (imported at runtime)."""
    p = user_data_dir() / "bibles"
    p.mkdir(parents=True, exist_ok=True)
    return p


def user_originals_dir() -> Path:
    """Archive of the raw files the user imported (kept verbatim)."""
    p = user_data_dir() / "originals"
    p.mkdir(parents=True, exist_ok=True)
    return p


def background_history_dir() -> Path:
    """Copies of custom backgrounds the user has registered (originals)."""
    p = user_data_dir() / "backgrounds"
    p.mkdir(parents=True, exist_ok=True)
    return p


def background_cache_dir() -> Path:
    """Aspect-cropped renders of the selected background (safe to delete)."""
    p = user_data_dir() / "backgrounds" / "cache"
    p.mkdir(parents=True, exist_ok=True)
    return p


def default_output_dir() -> Path:
    """Default folder for generated ``.pptx`` files.

    * Windows: the user's Documents folder (``%USERPROFILE%\\Documents``) +
      ``Bible2PPT``.
    * macOS (stub/port target): ``~/Documents/Bible2PPT``.
    * Other: ``~/Documents/Bible2PPT`` when present, else ``~/Bible2PPT``.
    """
    if sys.platform.startswith("win"):
        docs = _windows_documents_dir()
        return docs / APP_NAME
    if sys.platform == "darwin":
        return Path.home() / "Documents" / APP_NAME
    docs = Path.home() / "Documents"
    base = docs if docs.exists() else Path.home()
    return base / APP_NAME


def _windows_documents_dir() -> Path:
    """Resolve the real Windows Documents folder (honours a relocated folder).

    Falls back to ``~/Documents`` if the shell API is unavailable.
    """
    try:  # pragma: no cover - Windows only
        import ctypes
        import ctypes.wintypes

        CSIDL_PERSONAL = 5  # My Documents
        SHGFP_TYPE_CURRENT = 0
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(
            None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf
        )
        if buf.value:
            return Path(buf.value)
    except Exception:
        pass
    return Path.home() / "Documents"


def ensure_dir(path: os.PathLike | str) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p
