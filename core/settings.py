"""User settings persistence (local JSON in the per-OS user-data folder).

Stores everything needed to restore the UI on next launch: UI language,
selected translations, default background + background history, aspect ratio,
font / font size, output mode and output folder. Imported user Bibles live as
files under the user-data ``bibles`` folder and are rediscovered by the
registry, so only lightweight bookkeeping is kept here.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path

from . import paths

DEFAULT_ASPECT = "16:9"
DEFAULT_FONT_SIZE = 40


@dataclass
class Settings:
    ui_language: str = "ko"
    default_translation: str = "KRV"
    selected_translations: list[str] = field(default_factory=lambda: ["KRV"])
    aspect_ratio: str = DEFAULT_ASPECT
    font: str = ""  # empty -> resolved to the bundled default at load time
    body_font_size: int = DEFAULT_FONT_SIZE
    generate_mode: str = "separate"  # "separate" | "combined"
    output_folder: str = ""  # empty -> paths.default_output_dir()
    background: str = ""  # empty -> paths.default_background()
    background_history: list[str] = field(default_factory=list)

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
        return Path(self.background) if self.background else paths.default_background()

    def add_background_history(self, path: str, limit: int = 10) -> None:
        if path in self.background_history:
            self.background_history.remove(path)
        self.background_history.insert(0, path)
        del self.background_history[limit:]
