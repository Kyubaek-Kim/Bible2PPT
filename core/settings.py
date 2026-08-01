"""User settings persistence (local JSON in the per-OS user-data folder).

Stores everything needed to restore the UI on next launch: UI language,
selected + favourite translations, base translation, background + history,
aspect ratio, body font / size / bold, output mode and folder, plus the slide
layout-customisation (fractional boxes and per-element title / section
typography). Imported user Bibles live as files under the user-data ``bibles``
folder and are rediscovered by the registry, so only lightweight bookkeeping is
kept here.

Loading is forward/backward tolerant: unknown keys in the JSON are ignored and
missing keys fall back to the dataclass defaults, so an older settings file
keeps working after new fields are added.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path

from . import paths

DEFAULT_ASPECT = "16:9"
DEFAULT_FONT_SIZE = 32


@dataclass
class Settings:
    ui_language: str = "ko"
    default_translation: str = "KRV"
    selected_translations: list[str] = field(default_factory=lambda: ["KRV"])
    # translations pinned above the "show more" fold, in the order the user
    # checked their "자주 사용" (frequently used) boxes.
    favorite_translations: list[str] = field(default_factory=list)
    aspect_ratio: str = DEFAULT_ASPECT
    font: str = ""  # empty -> resolved to the bundled default at load time
    body_font_size: int = DEFAULT_FONT_SIZE
    body_bold: bool = True  # user-toggled bold for the body text
    generate_mode: str = "separate"  # "separate" | "combined"
    output_folder: str = ""  # empty -> paths.default_output_dir()
    # Registered custom backgrounds (stored *original* image paths). The default
    # background is always implicitly the first option and is not listed here.
    background_history: list[str] = field(default_factory=list)
    # The selected background: "" -> the built-in default, else one of the paths
    # in ``background_history``. Persisted so the choice survives a restart.
    selected_background: str = ""
    # Text colours as ``#rrggbb`` (empty -> engine default black).
    title_color: str = ""
    section_color: str = ""
    body_color: str = ""
    # Slide-layout customisation (item 6). Boxes are stored as fractional
    # rectangles [x, y, w, h] of the slide, so they are aspect-independent; an
    # empty dict means "use the engine defaults". The font *face* is always the
    # global 화면 설정 글자체; only size / bold / visibility differ per element.
    layout_boxes: dict[str, list[float]] = field(default_factory=dict)
    title_font_size: int = 40
    title_bold: bool = True
    title_enabled: bool = True
    section_font_size: int = 26
    section_bold: bool = True
    section_enabled: bool = True

    # -- persistence ------------------------------------------------------ #
    @classmethod
    def load(cls) -> Settings:
        fp = paths.settings_file()
        if not fp.exists():
            return cls()
        try:
            raw = json.loads(fp.read_text(encoding="utf-8"))
        except Exception:
            return cls()
        known = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in raw.items() if k in known})

    def save(self) -> None:
        fp = paths.settings_file()
        fp.write_text(
            json.dumps(asdict(self), ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    # -- resolved accessors ---------------------------------------------- #
    def resolved_output_folder(self) -> Path:
        return Path(self.output_folder) if self.output_folder else paths.default_output_dir()

    def resolved_background(self) -> Path:
        """The selected *original* background image (default when none selected)."""
        if self.selected_background and Path(self.selected_background).exists():
            return Path(self.selected_background)
        return paths.default_background()

    def background_options(self) -> list[tuple[str, str]]:
        """(key, display-name) pairs for the 배경 선택 dropdown.

        The first option is always the built-in default (key ``""``); the rest
        are the registered custom backgrounds, newest first.
        """
        opts: list[tuple[str, str]] = [("", "background_default")]
        for p in self.background_history:
            opts.append((p, Path(p).name))
        return opts

    def add_background(self, path: str, limit: int = 20) -> None:
        """Register a newly attached background (newest first, de-duplicated)."""
        if path in self.background_history:
            self.background_history.remove(path)
        self.background_history.insert(0, path)
        del self.background_history[limit:]

    def remove_background(self, path: str) -> None:
        """Unregister a background; reset the selection to default if it was it."""
        if path in self.background_history:
            self.background_history.remove(path)
        if self.selected_background == path:
            self.selected_background = ""

    # -- favourites ------------------------------------------------------- #
    def set_favorite(self, code: str, on: bool) -> None:
        """Pin/unpin a translation; pinning preserves the check order."""
        if on:
            if code not in self.favorite_translations:
                self.favorite_translations.append(code)
        elif code in self.favorite_translations:
            self.favorite_translations.remove(code)
