"""Core-logic tests (no Tkinter). Run with: pytest -q"""
from __future__ import annotations

import pytest

from core import alignment, bible, generator, importer, ppt
from core.i18n import I18n
from core.parser import ParseError, ReferenceParser, normalize_text


@pytest.fixture(scope="module")
def registry() -> bible.Registry:
    return bible.Registry.load()


@pytest.fixture(scope="module")
def parser() -> ReferenceParser:
    return generator.make_parser()


# --------------------------------------------------------------------------- #
# Parser
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    "text,expected",
    [
        ("창세기 1:1-3", ("Gen", 1, 1, 1, 3)),
        ("창 1:23-2:5", ("Gen", 1, 23, 2, 5)),
        ("창15:1~2", ("Gen", 15, 1, 15, 2)),
        ("Gen 1:1", ("Gen", 1, 1, 1, 1)),
        ("1 Corinthians 13:4-7", ("1Cor", 13, 4, 13, 7)),
        ("요일1:1", ("1John", 1, 1, 1, 1)),
        ("창 1 1 3", ("Gen", 1, 1, 1, 3)),
        ("창 15 1 16 5", ("Gen", 15, 1, 16, 5)),
    ],
)
def test_parse_variants(parser, text, expected):
    r = parser.parse(text)
    assert (r.book_id, r.start_chapter, r.start_verse, r.end_chapter, r.end_verse) == expected


def test_parse_whole_chapter(parser):
    r = parser.parse("창 15")
    assert (r.book_id, r.start_chapter, r.end_chapter, r.end_verse) == ("Gen", 15, 15, None)


def test_parse_failure(parser):
    with pytest.raises(ParseError):
        parser.parse("nonsense 없는책 1:1")


def test_normalize_text():
    assert normalize_text("창 15 : 1 ~ 15") == "창 15:1-15"


# --------------------------------------------------------------------------- #
# Reference expansion (incl. legacy start-verse-inclusion bug)
# --------------------------------------------------------------------------- #
def test_expand_includes_start_verse(parser, registry):
    ref = parser.parse("Gen 1:1-3")
    coords = bible.expand_reference(ref, fallback=registry.get("KRV"))
    assert [(c.chapter, c.verse) for c in coords] == [(1, 1), (1, 2), (1, 3)]


def test_expand_cross_chapter(parser, registry):
    ref = parser.parse("창 1:30-2:2")
    coords = bible.expand_reference(ref, fallback=registry.get("KRV"))
    assert (coords[0].chapter, coords[0].verse) == (1, 30)
    assert (coords[-1].chapter, coords[-1].verse) == (2, 2)
    assert any(c.chapter == 2 for c in coords)


# --------------------------------------------------------------------------- #
# Text integrity + alignment
# --------------------------------------------------------------------------- #
def test_text_verbatim(registry):
    # stored exactly as the source provides it — including the trailing space,
    # proving nothing rewrites/strips verse text.
    krv = registry.get("KRV")
    assert krv.get_verse("Gen", 1, 1) == "태초에 하나님이 천지를 창조하시니라 "


def test_alignment_interleaves(parser, registry):
    ref = parser.parse("Gen 1:1-2")
    coords = bible.expand_reference(ref, fallback=registry.get("KRV"))
    trans = [registry.get("KRV"), registry.get("KJV")]
    bundles = alignment.build_bundles(coords, trans, missing_text="(missing)")
    assert len(bundles) == 2
    # each bundle has one cell per translation, in order
    assert [code for code, _ in bundles[0].cells] == ["KRV", "KJV"]
    assert bundles[0].cells[0][1].text.startswith("태초에")


def test_alignment_missing_marker(registry):
    # a translation without a book yields a "missing" cell, never a crash
    from core.bible import Coord

    trans = registry.get("TR")  # NT only
    vsf = alignment.load_versification(trans.meta.versification)
    cell = alignment.align_cell(Coord("Gen", 1, 1), trans, vsf, missing_text="(none)")
    assert cell.status == "missing" and cell.text == "(none)"


# --------------------------------------------------------------------------- #
# PPT engine
# --------------------------------------------------------------------------- #
def test_meaningful_title():
    assert ppt.meaningful_title("사랑") is True
    assert ppt.meaningful_title("   ") is False
    assert ppt.meaningful_title("!! ?? --") is False


def test_wrap_line_hard_split():
    lines = ppt.wrap_line("x" * 100, 10)
    assert all(len(line) <= 10 for line in lines)
    assert "".join(lines) == "x" * 100


def test_generate_separate(tmp_path, registry):
    i18n = I18n("ko")
    style = ppt.SlideStyle(aspect="16:9", font_name="나눔고딕", body_font_size=32)
    res = generator.generate(
        [generator.PassageInput("Gen 1:1-3", "창조"), generator.PassageInput("시 23:1-6", "")],
        registry=registry,
        translation_codes=["KRV", "KJV"],
        style=style,
        background=None,
        output_folder=tmp_path,
        mode="separate",
        i18n=i18n,
    )
    assert len(res.output_paths) == 2
    assert not res.errors
    for p in res.output_paths:
        assert p.exists() and p.suffix == ".pptx"


def test_generate_combined(tmp_path, registry):
    i18n = I18n("en")
    style = ppt.SlideStyle()
    res = generator.generate(
        [generator.PassageInput("Gen 1:1-3"), generator.PassageInput("John 3:16")],
        registry=registry,
        translation_codes=["KJV"],
        style=style,
        background=None,
        output_folder=tmp_path,
        mode="combined",
        i18n=i18n,
    )
    assert len(res.output_paths) == 1


def test_generate_reports_parse_error(tmp_path, registry):
    i18n = I18n("ko")
    res = generator.generate(
        [generator.PassageInput("완전히없는책 1:1")],
        registry=registry,
        translation_codes=["KRV"],
        style=ppt.SlideStyle(),
        background=None,
        output_folder=tmp_path,
        mode="separate",
        i18n=i18n,
    )
    assert res.output_paths == []
    assert len(res.errors) == 1


# --------------------------------------------------------------------------- #
# Importer
# --------------------------------------------------------------------------- #
def test_import_txt_roundtrip(tmp_path):
    f = tmp_path / "sample.txt"
    f.write_text("창 1:1 태초에 본문\nGenesis 1:2 second verse\n창 1:1 dup\n", encoding="utf-8")
    report = importer.parse_file(f)
    assert report.books["Gen"]["1"]["1"] == "태초에 본문"
    assert report.books["Gen"]["1"]["2"] == "second verse"
    assert len(report.duplicates) == 1
    assert report.ok is False  # duplicate blocks review


def test_import_json_nested(tmp_path):
    f = tmp_path / "sample.json"
    f.write_text('{"창": {"1": {"1": "본문1", "2": "본문2"}}}', encoding="utf-8")
    report = importer.parse_file(f)
    assert report.n_verses == 2
    assert report.books["Gen"]["1"]["2"] == "본문2"


def test_import_json_flat(tmp_path):
    f = tmp_path / "flat.json"
    f.write_text('{"창1:1": "본문", "Gen 1:2": "text2"}', encoding="utf-8")
    report = importer.parse_file(f)
    assert report.books["Gen"]["1"]["1"] == "본문"
    assert report.books["Gen"]["1"]["2"] == "text2"


def test_import_register(tmp_path):
    f = tmp_path / "u.json"
    f.write_text('{"창": {"1": {"1": "나의 번역"}}}', encoding="utf-8")
    report = importer.parse_file(f)
    out = importer.register(
        report, code="MINE", name="내 번역", language="ko", original_path=f
    )
    assert out.exists()
    reg = bible.Registry.load()
    assert "MINE" in reg.codes()
    assert reg.get("MINE").get_verse("Gen", 1, 1) == "나의 번역"
    out.unlink()
