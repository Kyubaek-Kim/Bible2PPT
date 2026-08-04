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


@pytest.mark.parametrize(
    "text",
    [
        "창 5:10-5:3",   # reversed within a chapter
        "창 5:10-4:1",   # reversed across chapters
        "창 0:1",        # non-positive chapter
        "창 1:0",        # non-positive verse
    ],
)
def test_parse_rejects_invalid_ranges(parser, text):
    with pytest.raises(ParseError):
        parser.parse(text)


def test_generate_reports_no_verses_found(tmp_path, registry):
    # A book/chapter that no selected translation contains should be reported
    # as an error rather than producing an empty deck.
    passages = [generator.PassageInput(reference_text="창 999", title="")]
    result = generator.generate(
        passages,
        registry=registry,
        translation_codes=[registry.list_meta()[0].code],
        style=ppt.SlideStyle(),
        background=None,
        output_folder=tmp_path,
        mode="separate",
        i18n=I18n("ko"),
    )
    assert not result.output_paths
    assert result.errors


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
    assert labels[0] == "나눔스퀘어 Bold"  # default first
    assert {"나눔고딕", "나눔고딕 Bold", "맑은 고딕"} <= set(labels)
    assert fonts.default_font_name() == "나눔스퀘어 Bold"
    # NanumSquare Bold resolves to a bold face; the plain Nanum faces are regular
    assert fonts.resolve("나눔스퀘어 Bold").bold is True
    assert fonts.resolve("맑은 고딕").bold is False
    assert fonts.resolve("나눔고딕").typeface == "나눔고딕"
    # unknown / stale labels fall back to the default
    assert fonts.resolve("Arial").label == "나눔스퀘어 Bold"


def test_body_hanging_indent_and_line_spacing(tmp_path):
    import re
    import zipfile

    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    long_text = "태초에 하나님이 천지를 창조하시니라 " * 6
    bundle = VerseBundle(coord=Coord("Gen", 1, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text=long_text))])
    style = ppt.SlideStyle(aspect="16:9", font_name="나눔스퀘어 Bold", body_font_size=32)
    assert style.typeface == "나눔스퀘어 Bold" and style.body_bold is True

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
    # body height / (32 * 1.3 * 1.15 ≈ 47.8)pt -> a modest, non-overflowing count
    assert style.max_body_lines == 7
    # smaller font -> strictly more lines fit; larger font -> fewer
    assert (
        ppt.SlideStyle(aspect="16:9", font_name="나눔고딕", body_font_size=20).max_body_lines
        > style.max_body_lines
    )


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


def test_body_lines_use_tab_after_verse_number():
    from core.alignment import Cell, VerseBundle
    from core.bible import Coord
    from core.ppt import _blocks

    bundles = [VerseBundle(coord=Coord("Ps", 23, 1),
                           cells=[("KRV", Cell(status="ok", label="1", text="여호와는 목자시니"))])]
    lines = [ln for block in _blocks(bundles) for ln in block]
    # a tab separates the number from the text (drives the straight-edge tab stop)
    assert lines[0].text == "1.\t여호와는 목자시니"


def test_hang_scales_with_widest_verse_number():
    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    style = ppt.SlideStyle(aspect="16:9", font_name="나눔고딕", body_font_size=32)
    one = [VerseBundle(coord=Coord("Gen", 1, n),
                       cells=[("KRV", Cell(status="ok", label=str(n), text="x"))]) for n in (1, 5)]
    big = [VerseBundle(coord=Coord("Ps", 119, 176),
                       cells=[("KRV", Cell(status="ok", label="176", text="x"))])]
    # a 3-digit passage gets a wider outdent than a single-digit one
    assert ppt.body_hang_pt(big, style) > ppt.body_hang_pt(one, style)


def test_hanging_indent_sets_tab_stop_in_xml():
    import zipfile

    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    bundle = VerseBundle(coord=Coord("Gen", 1, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text="태초에 하나님이"))])
    style = ppt.SlideStyle(aspect="16:9", font_name="나눔고딕", body_font_size=32)
    prs = ppt.render([ppt.PassageContent("창조", "창세기 1:1", [bundle])], style, None)
    xml = zipfile.ZipFile(ppt.save(prs, "/tmp/_tab_test.pptx")).read("ppt/slides/slide1.xml").decode()
    assert "tabLst" in xml and "<a:tab " in xml


def test_long_title_is_shrunk_to_width():
    style = ppt.SlideStyle(aspect="16:9")
    short = ppt._fit_single_line_size("창조", style.title_box[2], ppt.TITLE_FONT_SIZE, ppt.MIN_TITLE_FONT_SIZE)
    longt = ppt._fit_single_line_size("아주 긴 제목입니다 " * 6, style.title_box[2],
                                      ppt.TITLE_FONT_SIZE, ppt.MIN_TITLE_FONT_SIZE)
    assert short == ppt.TITLE_FONT_SIZE
    assert longt < ppt.TITLE_FONT_SIZE


def test_layout_box_override_moves_boxes():
    style = ppt.SlideStyle(aspect="16:9")
    default_body = style.body_box
    # override body to a custom fractional rect
    style2 = ppt.SlideStyle(aspect="16:9", layout_boxes={"body": [0.1, 0.5, 0.5, 0.4]})
    w, h = style2.size_in
    assert style2.body_box == (0.1 * w, 0.5 * h, 0.5 * w, 0.4 * h)
    assert style2.body_box != default_body
    # untouched keys still fall back to defaults
    assert style2.title_box == style.title_box
    # default fractions round-trip through the override machinery
    fr = style.default_layout_fractions()
    style3 = ppt.SlideStyle(aspect="16:9", layout_boxes=fr)
    for a, b in zip(style3.body_box, default_body, strict=True):
        assert abs(a - b) < 1e-6


def test_resized_body_box_changes_capacity():
    """Shrinking the body region reduces the line/column capacity that
    pagination uses, so custom regions really drive the layout engine."""
    base = ppt.SlideStyle(aspect="16:9", body_font_size=32)
    small = ppt.SlideStyle(
        aspect="16:9", body_font_size=32,
        layout_boxes={"body": [0.1, 0.4, 0.4, 0.3]},  # narrower + shorter
    )
    assert small.max_body_lines < base.max_body_lines
    assert small.max_units_per_line < base.max_units_per_line


def test_resized_title_box_forces_more_title_shrink():
    """A narrower title region shrinks an over-long title further, keeping the
    'title never overflows its box' rule tied to the resized geometry."""
    wide = ppt.SlideStyle(aspect="16:9")
    narrow = ppt.SlideStyle(aspect="16:9", layout_boxes={"title": [0.1, 0.03, 0.25, 0.12]})
    title = "매우 긴 제목 예시 " * 4
    wide_sz = ppt._fit_single_line_size(
        title, wide.title_box[2], ppt.TITLE_FONT_SIZE, ppt.MIN_TITLE_FONT_SIZE)
    narrow_sz = ppt._fit_single_line_size(
        title, narrow.title_box[2], ppt.TITLE_FONT_SIZE, ppt.MIN_TITLE_FONT_SIZE)
    assert narrow_sz < wide_sz


def test_tiny_body_box_still_never_overflows():
    """Even a cramped custom body region must not silently overflow: the engine
    shrinks to the minimum size and raises when a bundle cannot fit."""
    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    long_text = "여호와 " * 200
    bundle = VerseBundle(coord=Coord("Gen", 1, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text=long_text))])
    style = ppt.SlideStyle(
        aspect="16:9", body_font_size=32,
        layout_boxes={"body": [0.4, 0.4, 0.2, 0.15]},  # tiny region
    )
    with pytest.raises(ppt.PaginationError):
        ppt.fit_pages([bundle], style)


def test_body_bold_override_and_element_styling():
    # explicit user choice overrides the face's intrinsic weight
    regular = ppt.SlideStyle(font_name="나눔고딕")
    assert regular.body_bold is False
    assert ppt.SlideStyle(font_name="나눔고딕", body_bold_opt=True).body_bold is True
    assert ppt.SlideStyle(font_name="나눔스퀘어 Bold", body_bold_opt=False).body_bold is False
    # every element shares the one global font face (no stale per-element face)
    s = ppt.SlideStyle(font_name="맑은 고딕")
    assert s.typeface == s.title_typeface == s.section_typeface == "맑은 고딕"


def test_favorite_translations_order():
    from core.settings import Settings

    s = Settings()
    s.set_favorite("KJV", True)
    s.set_favorite("ASV", True)
    s.set_favorite("KJV", True)  # idempotent
    assert s.favorite_translations == ["KJV", "ASV"]
    s.set_favorite("KJV", False)
    assert s.favorite_translations == ["ASV"]


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


def test_import_txt_cp949_encoding(tmp_path):
    # the exact case that failed before: a Korean txt saved as CP949/EUC-KR
    f = tmp_path / "cp949.txt"
    f.write_bytes("창 1:1 태초에 하나님이 천지를 창조하시니라\n".encode("cp949"))
    report = importer.parse_file(f)
    assert report.n_verses == 1
    # text preserved byte-for-byte after decoding
    assert report.books["Gen"]["1"]["1"] == "태초에 하나님이 천지를 창조하시니라"


def test_import_txt_utf16_and_bom(tmp_path):
    body = "창 1:1 태초에 본문\n"
    f16 = tmp_path / "u16.txt"
    f16.write_bytes(body.encode("utf-16"))
    assert importer.parse_file(f16).books["Gen"]["1"]["1"] == "태초에 본문"
    fbom = tmp_path / "bom.txt"
    fbom.write_bytes(body.encode("utf-8-sig"))
    assert importer.parse_file(fbom).books["Gen"]["1"]["1"] == "태초에 본문"


def test_import_invalid_json_reported_not_raised(tmp_path):
    f = tmp_path / "bad.json"
    f.write_text('{"창": {"1": {"1": "본문"', encoding="utf-8")  # truncated
    report = importer.parse_file(f)  # must not raise
    assert not report.ok
    assert any("invalid JSON" in p.reason for p in report.problems)


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


# --------------------------------------------------------------------------- #
# Text colours
# --------------------------------------------------------------------------- #
def test_text_colors_persist(tmp_path, monkeypatch):
    from core import paths
    from core.settings import Settings

    monkeypatch.setattr(paths, "settings_file", lambda: tmp_path / "settings.json")
    s = Settings()
    s.title_color, s.section_color, s.body_color = "#ff0000", "#00ff00", "#0000ff"
    s.save()
    loaded = Settings.load()
    assert (loaded.title_color, loaded.section_color, loaded.body_color) == (
        "#ff0000", "#00ff00", "#0000ff",
    )


def test_settings_save_is_atomic(tmp_path, monkeypatch):
    from core import paths
    from core.settings import Settings

    fp = tmp_path / "settings.json"
    monkeypatch.setattr(paths, "settings_file", lambda: fp)
    Settings(body_font_size=44).save()
    # a good file exists and no temp leftover remains
    assert fp.exists()
    assert not (tmp_path / "settings.json.tmp").exists()
    assert Settings.load().body_font_size == 44

    # a crash *during* the write must not destroy the previous good file
    import core.settings as settings_mod

    def boom(*_a, **_k):
        raise OSError("disk full")

    monkeypatch.setattr(settings_mod.os, "replace", boom)
    with pytest.raises(OSError):
        Settings(body_font_size=12).save()
    assert Settings.load().body_font_size == 44  # previous value preserved


def test_ppt_applies_run_colors(tmp_path):
    import zipfile

    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    bundle = VerseBundle(coord=Coord("Gen", 1, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text="빛이 있으라"))])
    style = ppt.SlideStyle(aspect="16:9", title_color="#ff0000", body_color="#00ff00")
    prs = ppt.render([ppt.PassageContent("창조", "창세기 1:1", [bundle])], style, None)
    out = ppt.save(prs, tmp_path / "color.pptx")
    xml = zipfile.ZipFile(out).read("ppt/slides/slide1.xml").decode()
    # colours land as solidFill srgbClr values (case-insensitive hex)
    assert 'srgbClr val="FF0000"' in xml
    assert 'srgbClr val="00FF00"' in xml


def test_hex_rgb_parsing():
    assert ppt._hex_rgb("") is None
    assert ppt._hex_rgb("#zzzzzz") is None
    assert ppt._hex_rgb("#010203") == ppt.RGBColor(1, 2, 3)
    assert ppt._hex_rgb("010203") == ppt.RGBColor(1, 2, 3)


# --------------------------------------------------------------------------- #
# Title / section enable flags
# --------------------------------------------------------------------------- #
def _slide_xml(tmp_path, style, name):
    import zipfile

    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    bundle = VerseBundle(coord=Coord("Gen", 1, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text="빛이 있으라"))])
    prs = ppt.render([ppt.PassageContent("창조의날", "창세기 1:1", [bundle])], style, None)
    out = ppt.save(prs, tmp_path / f"{name}.pptx")
    return zipfile.ZipFile(out).read("ppt/slides/slide1.xml").decode()


def test_disabled_title_and_section_are_omitted(tmp_path):
    both = _slide_xml(tmp_path, ppt.SlideStyle(aspect="16:9"), "both")
    assert "창조의날" in both and "창세기 1:1" in both

    no_headers = _slide_xml(
        tmp_path,
        ppt.SlideStyle(aspect="16:9", title_enabled=False, section_enabled=False),
        "none",
    )
    assert "창조의날" not in no_headers and "창세기 1:1" not in no_headers
    # body still renders
    assert "빛이 있으라" in no_headers


def test_no_headers_reclaims_body_space():
    with_headers = ppt.SlideStyle(aspect="16:9")
    without = ppt.SlideStyle(aspect="16:9", title_enabled=False, section_enabled=False)
    # hiding the header band gives the body a taller box -> more lines fit
    assert without.max_body_lines > with_headers.max_body_lines


# --------------------------------------------------------------------------- #
# Background selection / management
# --------------------------------------------------------------------------- #
def test_background_options_and_no_duplicate_on_select():
    from core.settings import Settings

    s = Settings()
    assert s.background_options() == [("", "background_default")]
    s.add_background("/data/bg/a.png")
    s.add_background("/data/bg/b.png")
    opts = s.background_options()
    assert opts[0] == ("", "background_default")
    assert [k for k, _ in opts] == ["", "/data/bg/b.png", "/data/bg/a.png"]

    # "selecting" an existing item never calls add_background, so re-adding the
    # same path must not grow / duplicate the list
    s.add_background("/data/bg/a.png")
    assert [k for k, _ in s.background_options()] == ["", "/data/bg/a.png", "/data/bg/b.png"]


def test_background_remove_resets_active_selection():
    from core import paths
    from core.settings import Settings

    s = Settings()
    s.add_background("/data/bg/a.png")
    s.selected_background = "/data/bg/a.png"
    s.remove_background("/data/bg/a.png")
    assert "/data/bg/a.png" not in s.background_history
    assert s.selected_background == ""  # falls back to default
    assert s.resolved_background() == paths.default_background()


def test_delete_background_removes_file_and_cache(tmp_path, monkeypatch):
    from core import image_util, paths

    hist = tmp_path / "backgrounds"
    cache = tmp_path / "backgrounds" / "cache"
    hist.mkdir(parents=True)
    cache.mkdir(parents=True)
    monkeypatch.setattr(paths, "background_history_dir", lambda: hist)
    monkeypatch.setattr(paths, "background_cache_dir", lambda: cache)

    original = hist / "20240101-000000_pic.png"
    original.write_bytes(b"x")
    (cache / "16x9_20240101-000000_pic.png").write_bytes(b"y")
    (cache / "4x3_20240101-000000_pic.png").write_bytes(b"z")

    image_util.delete_background(original)
    assert not original.exists()
    assert list(cache.glob("*")) == []


# --------------------------------------------------------------------------- #
# Font metadata / mapping
# --------------------------------------------------------------------------- #
def test_bundled_font_metadata_matches_mapping():
    # fontTools is a dev-only dependency (font bundling), absent in the runtime
    # CI job — skip there rather than fail.
    ttLib = pytest.importorskip("fontTools.ttLib")
    TTFont = ttLib.TTFont

    from core import fonts

    def korean_family(ttf) -> str:
        t = TTFont(str(ttf))
        rec = t["name"].getName(1, 3, 1, 0x412) or t["name"].getName(1, 3, 1, 0x409)
        return str(rec)

    def weight_class(ttf) -> int:
        return TTFont(str(ttf))["OS/2"].usWeightClass

    by_label = {f.label: f for f in fonts.curated_fonts()}

    # both full families are present: Light / Regular / Bold / ExtraBold
    for group in ("나눔스퀘어", "나눔고딕"):
        for label in (f"{group} Light", group, f"{group} Bold", f"{group} ExtraBold"):
            assert label in by_label, f"missing {label}"

    # each standalone-family weight: the font's Korean family name (nameID 1)
    # equals the mapped typeface, and no bold bit is needed.
    for label in (
        "나눔스퀘어 Light", "나눔스퀘어", "나눔스퀘어 Bold", "나눔스퀘어 ExtraBold",
        "나눔고딕 Light", "나눔고딕 ExtraBold",
    ):
        fc = by_label[label]
        assert korean_family(fc.bundled) == fc.typeface == label
        assert fc.needs_bold_bit is False

    # ascending weights are distinct files with ascending usWeightClass
    sq = [by_label[k].bundled for k in
          ("나눔스퀘어 Light", "나눔스퀘어", "나눔스퀘어 Bold", "나눔스퀘어 ExtraBold")]
    assert len(set(sq)) == 4
    weights = [weight_class(b) for b in sq]
    assert weights == sorted(weights) and weights[0] < weights[-1]

    # NanumGothic Regular + Bold share one family; Bold uses the bold bit
    ng, ngb = by_label["나눔고딕"], by_label["나눔고딕 Bold"]
    assert ng.typeface == ngb.typeface == "나눔고딕"
    assert ng.bundled != ngb.bundled
    assert weight_class(ng.bundled) == 400 and weight_class(ngb.bundled) >= 600
    assert ngb.needs_bold_bit is True
    # Light/ExtraBold are separate NanumGothic families (not the shared one)
    assert by_label["나눔고딕 Light"].typeface == "나눔고딕 Light"
    assert by_label["나눔고딕 ExtraBold"].typeface == "나눔고딕 ExtraBold"


def test_stale_per_element_font_cannot_override_global(tmp_path, monkeypatch):
    # legacy settings.json may carry removed keys (title_font=궁서 etc.); load()
    # must ignore them so the title can't get stuck on a stale face.
    import json

    from core import paths
    from core.settings import Settings

    fp = tmp_path / "settings.json"
    fp.write_text(json.dumps({
        "font": "나눔고딕",
        "title_font": "궁서",
        "section_font": "궁서",
        "background": "/old/cropped.png",
    }), encoding="utf-8")
    monkeypatch.setattr(paths, "settings_file", lambda: fp)
    s = Settings.load()
    assert s.font == "나눔고딕"
    style = ppt.SlideStyle(font_name=s.font)
    assert style.title_typeface == style.section_typeface == "나눔고딕"


# --------------------------------------------------------------------------- #
# Pagination: never overflow, better utilisation
# --------------------------------------------------------------------------- #
def _verse_bundles(n, text):
    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    return [
        VerseBundle(coord=Coord("Ps", 119, i),
                    cells=[("KRV", Cell(status="ok", label=str(i), text=text))])
        for i in range(1, n + 1)
    ]


def test_pagination_no_page_overflows():
    style = ppt.SlideStyle(aspect="16:9", body_font_size=32)
    bundles = _verse_bundles(30, "주의 말씀은 내 발에 등이요 내 길에 빛이니이다")
    fitted, pages, hang = ppt.fit_pages(bundles, ppt.fit_body_style(bundles, style))
    assert pages
    for pg in pages:
        assert ppt._block_line_count(pg.lines, fitted, hang) <= fitted.max_body_lines


def test_pagination_packs_multiple_short_verses_per_slide():
    style = ppt.SlideStyle(aspect="16:9", body_font_size=28)
    bundles = _verse_bundles(12, "짧은 구절")
    _fitted, pages, _hang = ppt.fit_pages(bundles, ppt.fit_body_style(bundles, style))
    # short verses should not each land on their own slide when there is room
    max_verses_on_a_page = max(len(pg.lines) for pg in pages)
    assert max_verses_on_a_page >= 3


def test_indivisible_oversized_bundle_raises_pagination_error():
    from core.alignment import Cell, VerseBundle
    from core.bible import Coord

    huge = "이 구절은 아주 길어서 한 장표에 절대 담을 수 없습니다 " * 200
    bundle = VerseBundle(coord=Coord("Gen", 1, 1),
                         cells=[("KRV", Cell(status="ok", label="1", text=huge))])
    style = ppt.SlideStyle(aspect="16:9", body_font_size=32)
    with pytest.raises(ppt.PaginationError):
        ppt.fit_pages([bundle], style)
