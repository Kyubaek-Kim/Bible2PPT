"""Verse-reference parsing & normalisation.

Turns free-form user input into a structured :class:`Reference`. Handles:

* full names, built-in abbreviations and the *currently selected UI language's*
  book names (all supplied via the alias map);
* separator / spacing / special-character variants — ``창세기 15:1-15``,
  ``창 15 1 15``, ``창15:1~15``, ``Gen 15:1-15``, ``1 Corinthians 13:4-7``;
* cross-chapter ranges — ``창 1:23-2:5``;
* whole-chapter references — ``창 15`` (start verse 1 → end of chapter).

This module never touches Bible *text*; it only computes references.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass


class ParseError(ValueError):
    """Raised when a reference string cannot be understood."""


@dataclass(frozen=True)
class Reference:
    book_id: str
    start_chapter: int
    start_verse: int
    end_chapter: int
    # ``None`` means "to the end of ``end_chapter``" (whole-chapter / open range);
    # the resolver (core.bible) fills the real last verse.
    end_verse: int | None

    @property
    def is_single_verse(self) -> bool:
        return (
            self.start_chapter == self.end_chapter
            and self.end_verse is not None
            and self.start_verse == self.end_verse
        )


# Separators that all mean "range" or "chapter:verse".
_RANGE_CHARS = "-~–—〜­"
_CV_CHARS = ":.．：·"


def normalize_text(text: str) -> str:
    """Normalise width, separators and whitespace without discarding structure."""
    s = unicodedata.normalize("NFKC", text).strip()
    for ch in _RANGE_CHARS:
        s = s.replace(ch, "-")
    for ch in _CV_CHARS:
        s = s.replace(ch, ":")
    # collapse runs of spaces/tabs to a single space
    s = re.sub(r"\s+", " ", s)
    # tidy spacing around separators: "15 : 1 - 15" -> "15:1-15"
    s = re.sub(r"\s*:\s*", ":", s)
    s = re.sub(r"\s*-\s*", "-", s)
    return s.strip()


class ReferenceParser:
    """Parses references given an alias→book_id map.

    ``aliases`` should map every recognisable spelling (full name, abbreviation,
    localized name) to the canonical book id. Matching is case-insensitive and
    space-insensitive.
    """

    def __init__(self, aliases: dict[str, str]):
        self._alias_to_id: dict[str, str] = {}
        for alias, book_id in aliases.items():
            key = self._alias_key(alias)
            if key:
                self._alias_to_id[key] = book_id
        # Longest first so "요한일서" wins over "요", "1 John" over "1".
        names = sorted(aliases.keys(), key=len, reverse=True)
        # Build a book-name regex where each literal space becomes optional.
        parts = []
        for name in names:
            esc = re.escape(name.strip())
            esc = esc.replace(r"\ ", r"\s*")
            parts.append(esc)
        self._book_re = re.compile(r"^(" + "|".join(parts) + r")\s*", re.IGNORECASE)

    @staticmethod
    def _alias_key(alias: str) -> str:
        return re.sub(r"\s+", "", alias).lower()

    def resolve_book(self, name: str) -> str | None:
        return self._alias_to_id.get(self._alias_key(name))

    def match_book_prefix(self, text: str) -> tuple[str, str] | None:
        """Match a leading book name; return ``(book_id, remainder)`` or ``None``.

        Unlike :meth:`parse` this does not normalise separators in the
        remainder, so callers (e.g. the importer) can keep trailing verse text
        verbatim.
        """
        s = text.lstrip()
        m = self._book_re.match(s)
        if not m:
            return None
        book_id = self.resolve_book(m.group(1))
        if book_id is None:
            return None
        return book_id, s[m.end():]

    def parse(self, text: str) -> Reference:
        if not text or not text.strip():
            raise ParseError("empty reference")
        s = normalize_text(text)

        m = self._book_re.match(s)
        if not m:
            raise ParseError(f"unknown book in: {text!r}")
        book_name = m.group(1)
        book_id = self.resolve_book(book_name)
        if book_id is None:
            raise ParseError(f"unknown book: {book_name!r}")

        remainder = s[m.end():].strip()
        if not remainder:
            raise ParseError(f"no chapter/verse in: {text!r}")

        return self._parse_numbers(book_id, remainder, original=text)

    def _parse_numbers(self, book_id: str, rem: str, *, original: str) -> Reference:
        # Case 1: explicit chapter:verse form, optional range.
        #   15:1  |  15:1-15  |  1:23-2:5  |  15  (whole chapter)
        if ":" in rem:
            m = re.fullmatch(
                r"(?P<c1>\d+):(?P<v1>\d+)"
                r"(?:-(?:(?P<c2>\d+):)?(?P<v2>\d+))?",
                rem,
            )
            if not m:
                raise ParseError(f"could not parse reference: {original!r}")
            c1 = int(m.group("c1"))
            v1 = int(m.group("v1"))
            c2 = int(m.group("c2")) if m.group("c2") else c1
            v2 = int(m.group("v2")) if m.group("v2") else v1
            return Reference(book_id, c1, v1, c2, v2)

        # Case 2: space/dash separated numbers, no colon.
        #   15            -> whole chapter 15
        #   15 1          -> 15:1
        #   15 1 15       -> 15:1-15
        #   15 1 16 5     -> 15:1 - 16:5  (cross chapter, 4 numbers)
        nums = [int(n) for n in re.findall(r"\d+", rem)]
        if not nums:
            raise ParseError(f"no numbers in: {original!r}")
        if len(nums) == 1:  # whole chapter
            return Reference(book_id, nums[0], 1, nums[0], None)
        if len(nums) == 2:  # chapter:verse (single verse)
            return Reference(book_id, nums[0], nums[1], nums[0], nums[1])
        if len(nums) == 3:  # chapter:v1-v2 within one chapter
            return Reference(book_id, nums[0], nums[1], nums[0], nums[2])
        if len(nums) == 4:  # c1:v1 - c2:v2
            return Reference(book_id, nums[0], nums[1], nums[2], nums[3])
        raise ParseError(f"too many numbers in: {original!r}")
