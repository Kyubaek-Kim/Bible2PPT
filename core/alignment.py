"""Cross-translation verse alignment — a pure *display-layer* mapping.

The app pins a canonical verse axis (KJV, see :mod:`core.bible`). Different
translations number verses slightly differently (Hebrew Psalm titles, merged
verses, verses some editions omit, …). This module maps a canonical coordinate
to the coordinate(s) inside a given translation **without ever editing text**:

* *shift / relocation* → look the text up at the mapped source coordinate;
* *merged verses* (one source verse == several canonical verses) → show the
  text once on the head coordinate with an ``"N-M"`` label, skip the rest;
* *missing verses* → emit a ``missing`` cell rendered as "(verse not present)".

Scheme exceptions live in ``data/versification/<scheme>.map.json``. A missing
file means identity mapping (the common case for KJV-versified translations).
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from functools import cache

from . import paths
from .bible import Coord, Translation


def _key(book_id: str, chapter: int, verse: int) -> str:
    return f"{book_id}/{chapter}/{verse}"


def _parse_key(s: str) -> tuple[str, int, int]:
    book, chap, verse = s.rsplit("/", 2)
    return book, int(chap), int(verse)


@dataclass(frozen=True)
class Versification:
    scheme: str
    missing: frozenset[str]
    relocate: dict[str, str]
    # head canonical key -> ordered list of canonical keys the source merges
    merge_groups: dict[str, list[str]]
    # member canonical key -> head canonical key (for O(1) lookup)
    _member_to_head: dict[str, str]

    def is_missing(self, book_id: str, chapter: int, verse: int) -> bool:
        return _key(book_id, chapter, verse) in self.missing

    def source_coord(self, coord: Coord) -> Coord:
        """Canonical coord -> the coord to read text from in this scheme."""
        k = _key(coord.book_id, coord.chapter, coord.verse)
        if k in self.relocate:
            b, c, v = _parse_key(self.relocate[k])
            return Coord(b, c, v)
        return coord


@cache
def load_versification(scheme: str) -> Versification:
    fp = paths.versification_dir() / f"{scheme}.map.json"
    if not fp.exists():
        return Versification(scheme, frozenset(), {}, {}, {})
    raw = json.loads(fp.read_text(encoding="utf-8"))
    merge_groups: dict[str, list[str]] = raw.get("merge", {})
    member_to_head: dict[str, str] = {}
    for head, members in merge_groups.items():
        for m in members:
            member_to_head[m] = head
    return Versification(
        scheme=scheme,
        missing=frozenset(raw.get("missing", [])),
        relocate=dict(raw.get("relocate", {})),
        merge_groups=merge_groups,
        _member_to_head=member_to_head,
    )


@dataclass(frozen=True)
class Cell:
    """One translation's rendering of one canonical coordinate."""

    status: str  # "ok" | "merged" | "missing" | "skip"
    label: str  # verse-number label, e.g. "1" or "4-5"
    text: str  # original verse text (verbatim), or "" for missing/skip

    @property
    def visible(self) -> bool:
        return self.status in ("ok", "merged", "missing")


def align_cell(
    coord: Coord,
    translation: Translation,
    vsf: Versification,
    *,
    missing_text: str,
) -> Cell:
    """Render one canonical coordinate for one translation.

    Never mutates or reflows text — only selects which verbatim source verse(s)
    to show and how to label them.
    """
    k = _key(coord.book_id, coord.chapter, coord.verse)

    # part of a merged group?
    head = vsf._member_to_head.get(k)
    if head is not None:
        if head != k:
            return Cell(status="skip", label="", text="")
        members = vsf.merge_groups[head]
        verses = [_parse_key(m)[2] for m in members]
        src = vsf.source_coord(coord)
        text = translation.get_verse(src.book_id, src.chapter, src.verse)
        if text is None:
            return Cell(status="missing", label=_merge_label(verses), text=missing_text)
        return Cell(status="merged", label=_merge_label(verses), text=text)

    if vsf.is_missing(coord.book_id, coord.chapter, coord.verse):
        return Cell(status="missing", label=str(coord.verse), text=missing_text)

    src = vsf.source_coord(coord)
    text = translation.get_verse(src.book_id, src.chapter, src.verse)
    if text is None:
        return Cell(status="missing", label=str(coord.verse), text=missing_text)
    return Cell(status="ok", label=str(coord.verse), text=text)


def _merge_label(verses: list[int]) -> str:
    if not verses:
        return ""
    lo, hi = min(verses), max(verses)
    return str(lo) if lo == hi else f"{lo}-{hi}"


@dataclass(frozen=True)
class VerseBundle:
    """All selected translations' cells for one canonical coordinate.

    A bundle is the *atomic* unit of pagination: every visible cell for a given
    canonical verse must stay together on one slide (verse integrity).
    """

    coord: Coord
    cells: list[tuple[str, Cell]]  # (translation code, cell)

    @property
    def any_visible(self) -> bool:
        return any(cell.visible for _, cell in self.cells)


def build_bundles(
    coords: list[Coord],
    translations: list[Translation],
    *,
    missing_text: str,
) -> list[VerseBundle]:
    """Interleave translations per canonical verse (교차 배치).

    For each canonical coordinate, produce one cell per selected translation in
    the given order (translation A verse 1, translation B verse 1, A verse 2 …).
    Merge-continuation cells are dropped so merged text shows exactly once.
    """
    vsfs = {t.code: load_versification(t.meta.versification) for t in translations}
    bundles: list[VerseBundle] = []
    for coord in coords:
        cells: list[tuple[str, Cell]] = []
        for t in translations:
            cell = align_cell(coord, t, vsfs[t.code], missing_text=missing_text)
            if cell.status == "skip":
                continue
            cells.append((t.code, cell))
        bundle = VerseBundle(coord=coord, cells=cells)
        if bundle.any_visible:
            bundles.append(bundle)
    return bundles
