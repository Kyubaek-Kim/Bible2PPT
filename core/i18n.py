"""UI internationalisation.

UI language (what labels/buttons/book-dropdowns are shown in) is deliberately
**separate** from the translation of the Bible text being rendered. This module
only concerns the former.

String tables live in ``data/i18n/<lang>.json`` with shape::

    {
        "lang": "ko",
        "name": "한국어",
        "ui": { "generate": "생성", ... },
        "books": { "Gen": "창세기", ... }   # optional; falls back to canon
    }

Book display names fall back to the canonical table (``data/canon.json``) so a
new language only needs to translate the ``ui`` strings to be usable.
"""
from __future__ import annotations

import json

from . import paths

DEFAULT_LANG = "ko"
FALLBACK_LANG = "en"


class I18n:
    def __init__(self, lang: str = DEFAULT_LANG):
        self._tables: dict[str, dict] = {}
        self._canon_names: dict[str, dict[str, str]] = {}
        self._load_canon_names()
        self.lang = lang if self._available(lang) else DEFAULT_LANG
        self._ensure_loaded(self.lang)
        self._ensure_loaded(FALLBACK_LANG)

    # -- loading ---------------------------------------------------------- #
    def _load_canon_names(self) -> None:
        try:
            canon = json.loads(paths.canon_file().read_text(encoding="utf-8"))
        except Exception:
            canon = []
        for entry in canon:
            for lang_key in ("ko", "en"):
                if lang_key in entry:
                    self._canon_names.setdefault(lang_key, {})[entry["id"]] = entry[
                        lang_key
                    ]

    def _available(self, lang: str) -> bool:
        return (paths.i18n_dir() / f"{lang}.json").exists() or lang in self._canon_names

    def _ensure_loaded(self, lang: str) -> None:
        if lang in self._tables:
            return
        fp = paths.i18n_dir() / f"{lang}.json"
        if fp.exists():
            self._tables[lang] = json.loads(fp.read_text(encoding="utf-8"))
        else:
            self._tables[lang] = {"lang": lang, "ui": {}, "books": {}}

    # -- public API ------------------------------------------------------- #
    def set_lang(self, lang: str) -> None:
        if self._available(lang):
            self.lang = lang
            self._ensure_loaded(lang)

    def available_langs(self) -> list[tuple[str, str]]:
        """List of ``(code, display_name)`` for every installed UI language."""
        out: dict[str, str] = {}
        for fp in sorted(paths.i18n_dir().glob("*.json")):
            try:
                data = json.loads(fp.read_text(encoding="utf-8"))
            except Exception:
                continue
            out[fp.stem] = data.get("name", fp.stem)
        # canon-derived languages that lack an i18n file are still selectable
        for lang in self._canon_names:
            out.setdefault(lang, lang)
        return sorted(out.items())

    def t(self, key: str, /, **fmt: object) -> str:
        """Translate a UI key, formatting with ``**fmt`` when placeholders exist."""
        for lang in (self.lang, FALLBACK_LANG):
            ui = self._tables.get(lang, {}).get("ui", {})
            if key in ui:
                text = ui[key]
                return text.format(**fmt) if fmt else text
        return key

    def language_name(self, lang_code: str) -> str:
        """Localized display name for a Bible language code (e.g. 'ko' -> '한국어')."""
        for lang in (self.lang, FALLBACK_LANG):
            table = self._tables.get(lang, {}).get("languages", {})
            if lang_code in table:
                return table[lang_code]
        return lang_code

    def translation_label(self, name: str, lang_code: str) -> str:
        """Translation name annotated with its language, e.g. '개역한글 (한국어)'."""
        return f"{name} ({self.language_name(lang_code)})"

    def book_name(self, book_id: str) -> str:
        books = self._tables.get(self.lang, {}).get("books", {})
        if book_id in books:
            return books[book_id]
        if book_id in self._canon_names.get(self.lang, {}):
            return self._canon_names[self.lang][book_id]
        # fall back to fallback language / canon en / raw id
        return (
            self._tables.get(FALLBACK_LANG, {}).get("books", {}).get(book_id)
            or self._canon_names.get(FALLBACK_LANG, {}).get(book_id)
            or book_id
        )
