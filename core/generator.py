"""High-level orchestration: references + translations → ``.pptx`` files.

This is the seam the UI drives. It ties together parsing, reference expansion,
cross-translation alignment and the slide engine, and enforces the two output
modes (one PPT per passage vs. a single combined PPT). It contains no Tkinter
code so it stays unit-testable and reusable.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from . import alignment, bible, ppt
from .i18n import I18n
from .parser import ParseError, Reference, ReferenceParser


@dataclass
class PassageInput:
    reference_text: str
    title: str = ""


def make_parser() -> ReferenceParser:
    return ReferenceParser(bible.book_aliases())


def format_reference(ref: Reference, i18n: I18n) -> str:
    """Localized human range, e.g. ``창세기 1:1-5`` or ``창세기 1:23-2:5``."""
    name = i18n.book_name(ref.book_id)
    start = f"{ref.start_chapter}:{ref.start_verse}"
    if ref.end_verse is None:
        return f"{name} {ref.start_chapter}장" if ref.start_chapter == ref.end_chapter else f"{name} {ref.start_chapter}-{ref.end_chapter}장"
    if ref.start_chapter == ref.end_chapter:
        if ref.start_verse == ref.end_verse:
            return f"{name} {start}"
        return f"{name} {start}-{ref.end_verse}"
    return f"{name} {start}-{ref.end_chapter}:{ref.end_verse}"


def _safe_filename(text: str) -> str:
    s = re.sub(r'[\\/:*?"<>|]+', "_", text).strip()
    return s or "bible"


def _build_passage(
    passage: PassageInput,
    parser: ReferenceParser,
    translations: list[bible.Translation],
    i18n: I18n,
) -> tuple[ppt.PassageContent, str]:
    ref = parser.parse(passage.reference_text)
    fallback = translations[0] if translations else None
    coords = bible.expand_reference(ref, fallback=fallback)
    bundles = alignment.build_bundles(
        coords, translations, missing_text=i18n.t("verse_missing")
    )
    section_info = format_reference(ref, i18n)
    content = ppt.PassageContent(
        title=passage.title, section_info=section_info, bundles=bundles
    )
    return content, section_info


@dataclass
class GenerationResult:
    output_paths: list[Path]
    errors: list[tuple[PassageInput, str]]


def generate(
    passages: list[PassageInput],
    *,
    registry: bible.Registry,
    translation_codes: list[str],
    style: ppt.SlideStyle,
    background: Path | None,
    output_folder: Path,
    mode: str,  # "separate" | "combined"
    i18n: I18n,
) -> GenerationResult:
    """Generate PPT file(s). Returns produced paths and per-passage errors."""
    parser = make_parser()
    translations = [t for c in translation_codes if (t := registry.get(c)) is not None]
    if not translations:
        raise ValueError("no valid translations selected")

    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    contents: list[tuple[ppt.PassageContent, str]] = []
    errors: list[tuple[PassageInput, str]] = []
    for p in passages:
        try:
            content, section = _build_passage(p, parser, translations, i18n)
            contents.append((content, section))
        except ParseError:
            errors.append((p, i18n.t("parse_failed", text=p.reference_text)))
        except Exception as exc:  # noqa: BLE001 - surface any per-passage failure
            errors.append((p, str(exc)))

    outputs: list[Path] = []
    if not contents:
        return GenerationResult(outputs, errors)

    if mode == "combined":
        prs = ppt.render([c for c, _ in contents], style, background)
        first_section = contents[0][1]
        suffix = f"_외{len(contents) - 1}건" if len(contents) > 1 else ""
        out = output_folder / f"{_safe_filename(first_section)}{suffix}.pptx"
        outputs.append(ppt.save(prs, out))
    else:
        seen: dict[str, int] = {}
        for content, section in contents:
            base = _safe_filename(section)
            seen[base] = seen.get(base, 0) + 1
            name = base if seen[base] == 1 else f"{base}({seen[base]})"
            prs = ppt.render([content], style, background)
            outputs.append(ppt.save(prs, output_folder / f"{name}.pptx"))

    return GenerationResult(outputs, errors)
