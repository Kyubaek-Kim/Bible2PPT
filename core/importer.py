"""User Bible upload: parse → validate/review → register.

Accepts ``.txt`` and ``.json`` files, auto-detects their layout and converts to
the app's canonical ``{book_id: {chapter: {verse: text}}}`` schema. The verse
**text is preserved byte-for-byte** — only the surrounding delimiters between a
reference and its text are stripped.

The flow is deliberately two-phase:

1. :func:`parse_file` → an :class:`ImportReport` with statistics and every
   problem line (malformed / duplicate / unknown book) plus per-book verse-count
   mismatches against the canon. The UI shows this for review.
2. :func:`register` → writes the translation into the user's writable bibles
   folder, archives the original file, and returns the path so the registry can
   pick it up. Registration must be gated on a passing review by the caller.
"""
from __future__ import annotations

import json
import re
import shutil
import time
from dataclasses import dataclass, field
from pathlib import Path

from . import bible, paths
from .parser import ReferenceParser

# book c:v  |  book c v  (tab / multi-space separated), text follows
_CV_COLON = re.compile(r"^\s*(\d+)\s*[:.]\s*(\d+)[\s\t]+(.*)$", re.DOTALL)
_CV_SPACE = re.compile(r"^\s*(\d+)[\s\t]+(\d+)[\s\t]+(.*)$", re.DOTALL)
# flat-json key "창1:1" / "Gen 1:1"
_FLAT_KEY = re.compile(r"^\s*(.+?)\s*(\d+)\s*[:.]\s*(\d+)\s*$")


@dataclass
class Problem:
    line_no: int
    raw: str
    reason: str


@dataclass
class ImportReport:
    # canonical nested text (verbatim)
    books: dict[str, dict[str, dict[str, str]]] = field(default_factory=dict)
    problems: list[Problem] = field(default_factory=list)
    duplicates: list[Problem] = field(default_factory=list)
    count_mismatch: list[str] = field(default_factory=list)  # book_ids
    source_format: str = ""

    @property
    def n_books(self) -> int:
        return len(self.books)

    @property
    def n_chapters(self) -> int:
        return sum(len(ch) for ch in self.books.values())

    @property
    def n_verses(self) -> int:
        return sum(len(vs) for ch in self.books.values() for vs in ch.values())

    @property
    def ok(self) -> bool:
        """Review passes when there is text and no malformed/duplicate lines."""
        return self.n_verses > 0 and not self.problems and not self.duplicates


def _resolver() -> ReferenceParser:
    return ReferenceParser(bible.book_aliases())


def _add_verse(
    report: ImportReport,
    book_id: str,
    chapter: int,
    verse: int,
    text: str,
    *,
    line_no: int,
    raw: str,
) -> None:
    ch = report.books.setdefault(book_id, {}).setdefault(str(chapter), {})
    key = str(verse)
    if key in ch:
        report.duplicates.append(Problem(line_no, raw, f"duplicate {book_id} {chapter}:{verse}"))
        return
    ch[key] = text


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #
def parse_file(path: str | Path) -> ImportReport:
    p = Path(path)
    text = p.read_text(encoding="utf-8-sig")
    if p.suffix.lower() == ".json":
        return _parse_json(text)
    return _parse_txt(text)


def _parse_txt(content: str) -> ImportReport:
    report = ImportReport(source_format="txt")
    resolver = _resolver()
    for i, raw in enumerate(content.splitlines(), start=1):
        line = raw.rstrip("\r\n")
        if not line.strip():
            continue
        matched = resolver.match_book_prefix(line)
        if not matched:
            report.problems.append(Problem(i, raw, "unknown/missing book"))
            continue
        book_id, rest = matched
        m = _CV_COLON.match(rest) or _CV_SPACE.match(rest)
        if not m:
            report.problems.append(Problem(i, raw, "no chapter:verse"))
            continue
        chapter, verse, vtext = int(m.group(1)), int(m.group(2)), m.group(3).strip()
        if not vtext:
            report.problems.append(Problem(i, raw, "empty verse text"))
            continue
        _add_verse(report, book_id, chapter, verse, vtext, line_no=i, raw=raw)
    _flag_count_mismatch(report)
    return report


def _parse_json(content: str) -> ImportReport:
    data = json.loads(content)
    report = ImportReport(source_format="json")
    resolver = _resolver()

    if isinstance(data, dict) and "books" in data and isinstance(data["books"], dict):
        _ingest_nested(report, data["books"], resolver, assume_canonical=True)
    elif isinstance(data, list):
        _ingest_rows(report, data, resolver)
    elif isinstance(data, dict) and _is_flat(data):
        _ingest_flat(report, data, resolver)
    elif isinstance(data, dict):
        _ingest_nested(report, data, resolver, assume_canonical=False)
    else:
        report.problems.append(Problem(0, "", "unrecognized JSON structure"))

    _flag_count_mismatch(report)
    return report


def _is_flat(data: dict) -> bool:
    for k, v in list(data.items())[:20]:
        if isinstance(v, str) and _FLAT_KEY.match(k):
            return True
    return False


def _ingest_flat(report, data: dict, resolver) -> None:
    for i, (key, value) in enumerate(data.items(), start=1):
        if not isinstance(value, str):
            report.problems.append(Problem(i, str(key), "value is not text"))
            continue
        m = _FLAT_KEY.match(key)
        if not m:
            report.problems.append(Problem(i, str(key), "malformed key"))
            continue
        book_id = resolver.resolve_book(m.group(1))
        if not book_id:
            report.problems.append(Problem(i, str(key), f"unknown book {m.group(1)!r}"))
            continue
        _add_verse(report, book_id, int(m.group(2)), int(m.group(3)), value, line_no=i, raw=key)


def _ingest_rows(report, rows: list, resolver) -> None:
    for i, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            report.problems.append(Problem(i, str(row), "row is not an object"))
            continue
        book_raw = row.get("book") or row.get("book_id") or row.get("name")
        text = row.get("text") or row.get("verse_text") or row.get("t")
        chapter = row.get("chapter") or row.get("c")
        verse = row.get("verse") or row.get("v")
        if book_raw is None or text is None or chapter is None or verse is None:
            report.problems.append(Problem(i, str(row)[:60], "missing field"))
            continue
        book_id = resolver.resolve_book(str(book_raw)) or (
            str(book_raw) if str(book_raw) in bible.canon_index() else None
        )
        if not book_id:
            report.problems.append(Problem(i, str(book_raw), "unknown book"))
            continue
        _add_verse(report, book_id, int(chapter), int(verse), str(text), line_no=i, raw=str(row)[:60])


def _ingest_nested(report, books: dict, resolver, *, assume_canonical: bool) -> None:
    for i, (book_key, chapters) in enumerate(books.items(), start=1):
        if assume_canonical and book_key in bible.canon_index():
            book_id = book_key
        else:
            book_id = resolver.resolve_book(str(book_key)) or (
                book_key if book_key in bible.canon_index() else None
            )
        if not book_id:
            report.problems.append(Problem(i, str(book_key), "unknown book"))
            continue
        if not isinstance(chapters, dict):
            report.problems.append(Problem(i, str(book_key), "chapters not an object"))
            continue
        for chap_key, verses in chapters.items():
            if not isinstance(verses, dict):
                report.problems.append(Problem(i, f"{book_key} {chap_key}", "verses not an object"))
                continue
            try:
                chapter = int(chap_key)
            except (TypeError, ValueError):
                report.problems.append(Problem(i, f"{book_key} {chap_key}", "bad chapter number"))
                continue
            for verse_key, vtext in verses.items():
                try:
                    verse = int(verse_key)
                except (TypeError, ValueError):
                    report.problems.append(Problem(i, f"{book_key} {chap_key}:{verse_key}", "bad verse number"))
                    continue
                _add_verse(report, book_id, chapter, verse, str(vtext), line_no=i, raw=f"{book_key} {chap_key}:{verse_key}")


def _flag_count_mismatch(report: ImportReport) -> None:
    """Flag books whose per-chapter verse counts differ from the canon."""
    table = bible.canonical_versification()
    if not table:
        return
    for book_id, chapters in report.books.items():
        canon_chapters = table.get(book_id)
        if not canon_chapters:
            continue
        mismatch = False
        for chap_str, verses in chapters.items():
            chap = int(chap_str)
            expected = canon_chapters.get(chap)
            if expected is not None and len(verses) != expected:
                mismatch = True
                break
        if mismatch:
            report.count_mismatch.append(book_id)


# --------------------------------------------------------------------------- #
# Registration
# --------------------------------------------------------------------------- #
def _sanitize_code(code: str) -> str:
    c = re.sub(r"[^0-9A-Za-z_\-]", "", code).strip("-_")
    return c or "USER"


def register(
    report: ImportReport,
    *,
    code: str,
    name: str,
    language: str,
    abbr: str = "",
    versification: str = "kjv",
    original_path: str | Path | None = None,
) -> Path:
    """Persist a reviewed import as a user translation; return its file path.

    Writes into the writable user bibles folder (never the read-only bundle),
    archives the original upload, and returns the JSON path. The caller should
    reload the registry afterwards so the translation appears in the dropdown.
    """
    safe_code = _sanitize_code(code)
    payload = {
        "meta": {
            "code": safe_code,
            "name": name or safe_code,
            "language": language or "und",
            "abbr": abbr,
            "versification": versification,
            "source": f"user upload ({report.source_format})",
            "license": "user-provided",
        },
        "books": report.books,
    }
    out = paths.user_bibles_dir() / f"{safe_code}.json"
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if original_path:
        src = Path(original_path)
        if src.exists():
            stamp = time.strftime("%Y%m%d-%H%M%S")
            shutil.copy2(src, paths.user_originals_dir() / f"{safe_code}_{stamp}{src.suffix}")
    return out
