"""Canonical book model, translation registry and text lookup.

Design rules:

* The canonical **book id** (``Gen``…``Rev``) and a canonical **verse axis**
  (from ``data/versification/kjv.json``) are the app's coordinate system.
* A *translation* stores its verses under canonical book ids but keeps its own
  original verse numbering and, crucially, its **text verbatim**. Nothing in
  this module rewrites verse text.
* Translations ship in ``data/bibles/`` (bundled) and/or are imported by the
  user into the writable ``bibles`` folder. Both are discovered here.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

from . import paths
from .parser import Reference


# --------------------------------------------------------------------------- #
# Canon
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Book:
    id: str
    order: int
    testament: str
    ko: str
    ko_abbr: str
    en: str
    en_abbr: str


@lru_cache(maxsize=1)
def load_canon() -> list[Book]:
    data = json.loads(paths.canon_file().read_text(encoding="utf-8"))
    return [Book(**entry) for entry in data]


@lru_cache(maxsize=1)
def canon_index() -> dict[str, Book]:
    return {b.id: b for b in load_canon()}


def book_aliases() -> dict[str, str]:
    """Every recognisable spelling → canonical book id (for the parser).

    Always includes Korean and English names/abbreviations plus the id itself,
    so references are understood regardless of the active UI language.
    """
    aliases: dict[str, str] = {}
    for b in load_canon():
        for name in (b.id, b.ko, b.ko_abbr, b.en, b.en_abbr):
            if name:
                aliases[name] = b.id
    return aliases


@lru_cache(maxsize=1)
def canonical_versification() -> dict[str, dict[int, int]]:
    """Canonical last-verse table: ``{book_id: {chapter: last_verse}}`` (KJV axis).

    Used to enumerate reference ranges independently of any one translation.
    Falls back to an empty table if the file is missing (callers then use a
    translation's own verse keys).
    """
    fp = paths.versification_dir() / "kjv.json"
    if not fp.exists():
        return {}
    raw = json.loads(fp.read_text(encoding="utf-8"))
    out: dict[str, dict[int, int]] = {}
    for book_id, chapters in raw.items():
        out[book_id] = {int(c): int(v) for c, v in chapters.items()}
    return out


# --------------------------------------------------------------------------- #
# Translations
# --------------------------------------------------------------------------- #
@dataclass
class TranslationMeta:
    code: str
    name: str
    language: str
    abbr: str = ""
    versification: str = "kjv"
    source: str = ""
    license: str = ""
    direction: str = "ltr"
    origin: str = "bundled"  # "bundled" | "user"
    path: Path | None = None


class Translation:
    """A single Bible translation. Text is loaded lazily and never mutated."""

    def __init__(self, meta: TranslationMeta):
        self.meta = meta
        self._books: dict[str, dict[str, dict[str, str]]] | None = None

    @property
    def code(self) -> str:
        return self.meta.code

    def _ensure_loaded(self) -> None:
        if self._books is None:
            assert self.meta.path is not None
            data = json.loads(self.meta.path.read_text(encoding="utf-8"))
            self._books = data.get("books", {})

    def has_book(self, book_id: str) -> bool:
        self._ensure_loaded()
        return book_id in self._books  # type: ignore[operator]

    def chapters(self, book_id: str) -> list[int]:
        self._ensure_loaded()
        book = self._books.get(book_id, {})  # type: ignore[union-attr]
        return sorted(int(c) for c in book)

    def verses(self, book_id: str, chapter: int) -> list[int]:
        self._ensure_loaded()
        ch = self._books.get(book_id, {}).get(str(chapter), {})  # type: ignore[union-attr]
        return sorted(int(v) for v in ch)

    def get_verse(self, book_id: str, chapter: int, verse: int) -> str | None:
        """Return the *original* verse text, or ``None`` if absent."""
        self._ensure_loaded()
        return (
            self._books.get(book_id, {})  # type: ignore[union-attr]
            .get(str(chapter), {})
            .get(str(verse))
        )

    def last_verse(self, book_id: str, chapter: int) -> int | None:
        vs = self.verses(book_id, chapter)
        return vs[-1] if vs else None


# --------------------------------------------------------------------------- #
# Registry
# --------------------------------------------------------------------------- #
@dataclass
class Registry:
    translations: dict[str, Translation] = field(default_factory=dict)

    @classmethod
    def load(cls) -> Registry:
        reg = cls()
        reg.reload()
        return reg

    def reload(self) -> None:
        self.translations = {}
        for directory, origin in (
            (paths.bibles_dir(), "bundled"),
            (paths.user_bibles_dir(), "user"),
        ):
            if not directory.exists():
                continue
            for fp in sorted(directory.glob("*.json")):
                meta = self._read_meta(fp, origin)
                if meta is not None:
                    # user translations override bundled ones with the same code
                    self.translations[meta.code] = Translation(meta)

    @staticmethod
    def _read_meta(fp: Path, origin: str) -> TranslationMeta | None:
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
        except Exception:
            return None
        meta = data.get("meta")
        if not meta or "code" not in meta:
            return None
        known = {
            "code",
            "name",
            "language",
            "abbr",
            "versification",
            "source",
            "license",
            "direction",
        }
        kwargs = {k: v for k, v in meta.items() if k in known}
        return TranslationMeta(origin=origin, path=fp, **kwargs)

    # -- queries ---------------------------------------------------------- #
    def codes(self) -> list[str]:
        return list(self.translations.keys())

    def list_meta(self) -> list[TranslationMeta]:
        return [t.meta for t in self.translations.values()]

    def get(self, code: str) -> Translation | None:
        return self.translations.get(code)

    def by_language(self, language: str) -> list[Translation]:
        return [t for t in self.translations.values() if t.meta.language == language]


# --------------------------------------------------------------------------- #
# Reference expansion (canonical axis)
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Coord:
    """A canonical verse coordinate."""

    book_id: str
    chapter: int
    verse: int


def _last_verse(book_id: str, chapter: int, fallback: Translation | None) -> int:
    table = canonical_versification()
    if book_id in table and chapter in table[book_id]:
        return table[book_id][chapter]
    if fallback is not None:
        lv = fallback.last_verse(book_id, chapter)
        if lv is not None:
            return lv
    return 0


def expand_reference(ref: Reference, fallback: Translation | None = None) -> list[Coord]:
    """Enumerate the canonical coordinates covered by ``ref``.

    Supports cross-chapter ranges (``창 1:23-2:5``) and open/whole-chapter
    ranges (``end_verse is None`` -> last verse of ``end_chapter``). Uses the
    canonical versification table, falling back to ``fallback``'s own verse
    keys when the table lacks the book/chapter.
    """
    coords: list[Coord] = []
    start_c, end_c = ref.start_chapter, ref.end_chapter
    for chapter in range(start_c, end_c + 1):
        first_v = ref.start_verse if chapter == start_c else 1
        if chapter == end_c:
            last_v = ref.end_verse if ref.end_verse is not None else _last_verse(
                ref.book_id, chapter, fallback
            )
        else:
            last_v = _last_verse(ref.book_id, chapter, fallback)
        if not last_v:
            continue
        for verse in range(first_v, last_v + 1):
            coords.append(Coord(ref.book_id, chapter, verse))
    return coords
