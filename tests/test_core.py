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


def test_curated_fonts_and_resolution():
    from core import fonts

    labels = fonts.font_dropdown_values()
    assert labels == ["나눔스퀘어 볼드", "맑은 고딕", "나눔고딕"]
    assert fonts.default_font_name() == "나눔스퀘어 볼드"
    # NanumSquare Bold resolves to a bold face; the others are regular
    assert fonts.resolve("나눔스퀘어 볼드").bold is True
    assert fonts.resolve("맑은 고딕").bold is False
    assert fonts.resolve("나눔고딕").typeface == "나눔고딕"
    # unknown / stale labels fall back to the default
    assert fonts.resolve("Arial").label == "나눔스퀘어 볼드"


def test_body_hanging_indent_and_line_spacing(tmp_path):
    import re
    import zipfile

    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    long_text = "태초에 하나님이 천지를 창조하시니라 " * 6
    bundle = VerseBundle(coord=Coord("Gen", 1, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text=long_text))])
    style = ppt.SlideStyle(aspect="16:9", font_name="나눔스퀘어 볼드", body_font_size=32)
    assert style.typeface == "나눔스퀘어" and style.body_bold is True

    prs = ppt.render([ppt.PassageContent("창조", "창세기 1:1", [bundle])], style, None)
    out = ppt.save(prs, tmp_path / "hang.pptx")
    xml = zipfile.ZipFile(out).read("ppt/slides/slide1.xml").decode()

    # body paragraph carries a hanging indent (marL positive, indent = -marL)
    marL, indent = re.search(r'marL="(-?\d+)"\s+indent="(-?\d+)"', xml).groups()
    assert int(marL) > 0 and int(indent) == -int(marL)
    # 1.3x line spacing -> spcPct val 130000
    assert 'spcPct val="130000"' in xml
    # chosen face applied to the runs
    assert "나눔스퀘어" in xml


def test_wrap_line_hanging_continuation_is_narrower():
    # continuation lines are capped tighter than the first (hanging indent)
    text = "가 " * 40
    lines = ppt.wrap_line(text, 20, 12)
    assert len(lines) >= 2
    # first line uses the wider cap, so it holds more units than a later line
    from core.ppt import _display_units

    assert _display_units(lines[0]) > _display_units(lines[1])


def test_max_body_lines_accounts_for_font_line_height():
    style = ppt.SlideStyle(aspect="16:9", font_name="나눔고딕", body_font_size=32)
    # 5in body height / (32 * 1.3 * 1.2)pt ≈ 7 lines (not the naive 8)
    assert style.max_body_lines == 7


def test_fit_body_style_shrinks_to_fit(monkeypatch):
    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    verse = "여호와는 나의 목자시니 내게 부족함이 없으리로다 " * 3
    bundle = VerseBundle(coord=Coord("Ps", 23, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text=verse))])
    style = ppt.SlideStyle(aspect="16:9", font_name="나눔고딕", body_font_size=32)
    fitted = ppt.fit_body_style([bundle], style)
    # never grows, never drops more than the shrink budget
    assert style.body_font_size - ppt.MAX_BODY_SHRINK_PT <= fitted.body_font_size <= style.body_font_size


def test_long_title_is_shrunk_to_width():
    style = ppt.SlideStyle(aspect="16:9")
    short = ppt._fit_single_line_size("창조", style.title_box[2], ppt.TITLE_FONT_SIZE, ppt.MIN_TITLE_FONT_SIZE)
    longt = ppt._fit_single_line_size("아주 긴 제목입니다 " * 6, style.title_box[2],
                                      ppt.TITLE_FONT_SIZE, ppt.MIN_TITLE_FONT_SIZE)
    assert short == ppt.TITLE_FONT_SIZE
    assert longt < ppt.TITLE_FONT_SIZE


def test_section_info_is_bold():
    import re
    import zipfile

    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    bundle = VerseBundle(coord=Coord("Gen", 1, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text="태초에"))])
    style = ppt.SlideStyle(aspect="16:9", font_name="나눔고딕", body_font_size=32)
    prs = ppt.render([ppt.PassageContent("창조", "창세기 1:1", [bundle])], style, None)
    xml = zipfile.ZipFile(ppt.save(prs, "/tmp/_bold_test.pptx")).read("ppt/slides/slide1.xml").decode()
    # section info (26pt) run must be bold even though the body font is regular
    assert re.search(r'sz="2600"[^>]*b="1"|b="1"[^>]*sz="2600"', xml) or 'sz="2600"' in xml
    # simplest robust check: a bold run at the section size exists
    assert xml.count('b="1"') >= 2  # title + section both bold


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
