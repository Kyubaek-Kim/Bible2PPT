"""Slide engine: turns aligned verse bundles into a ``.pptx``.

Responsibilities:

* aspect ratio → slide size and text-box geometry (16:9 / 4:3 / A4);
* fonts → title / section-info / body (single family; body size configurable,
  title & section fixed spec) with East-Asian font names set correctly;
* full-slide background image;
* verse formatting → verse-number prefix, one verse per line, ``<N장>`` marker at
  chapter changes, cross-translation interleave;
* pagination → a verse *bundle* (all translations of one canonical verse) never
  splits across slides; lines-per-slide are recomputed from font/size/ratio; a
  bundle longer than a whole slide gets its own slide (overflow allowed).
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field, replace
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from . import fonts
from .alignment import VerseBundle

# Aspect ratio -> (width_in, height_in).
ASPECT_RATIOS: dict[str, tuple[float, float]] = {
    "16:9": (13.333, 7.5),
    "4:3": (10.0, 7.5),
    "A4": (11.69, 8.27),  # landscape A4
}
DEFAULT_ASPECT = "16:9"

TITLE_FONT_SIZE = 40
SECTION_FONT_SIZE = 26
MARGIN_IN = 0.6
IN_TO_PT = 72.0
LINE_SPACING = 1.3
# A font's intrinsic single-line height is larger than its point size; the
# "multiple" line-spacing above is applied on top of it. Fold both into the
# vertical-fit estimate so bundles don't overflow the bottom of a slide.
FONT_LINE_HEIGHT = 1.2
# hanging indent for body verses: the verse number hangs to the left while
# wrapped continuation lines are indented, so verse boundaries read clearly.
BODY_HANG_FACTOR = 1.6
# how far the body font size may be auto-reduced to fit verses tidily
MAX_BODY_SHRINK_PT = 2
MIN_BODY_FONT_SIZE = 12
# title auto-shrink bounds (keep the title on one line within the slide width)
MIN_TITLE_FONT_SIZE = 18


def _display_units(text: str) -> int:
    """Approximate rendered width in half-em units (CJK≈2, latin≈1)."""
    total = 0
    for ch in text:
        total += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return total


@dataclass
class SlideStyle:
    aspect: str = DEFAULT_ASPECT
    font_name: str = "나눔스퀘어 볼드"  # stored dropdown label
    body_font_size: int = 32

    @property
    def typeface(self) -> str:
        """Actual font family to set on runs (resolved from the label)."""
        return fonts.resolve(self.font_name).typeface

    @property
    def body_bold(self) -> bool:
        """Whether the chosen face is inherently bold (e.g. NanumSquare Bold)."""
        return fonts.resolve(self.font_name).bold

    @property
    def size_in(self) -> tuple[float, float]:
        return ASPECT_RATIOS.get(self.aspect, ASPECT_RATIOS[DEFAULT_ASPECT])

    @property
    def body_box(self) -> tuple[float, float, float, float]:
        """left, top, width, height (inches) of the body text area."""
        w, h = self.size_in
        left = MARGIN_IN
        top = 2.0
        width = w - 2 * MARGIN_IN
        height = h - top - 0.5
        return left, top, width, height

    @property
    def title_box(self) -> tuple[float, float, float, float]:
        w, _ = self.size_in
        return MARGIN_IN, 0.25, w - 2 * MARGIN_IN, 1.1

    @property
    def section_box(self) -> tuple[float, float, float, float]:
        w, _ = self.size_in
        return MARGIN_IN, 1.15, w - 2 * MARGIN_IN, 0.8

    @property
    def max_units_per_line(self) -> int:
        _, _, width_in, _ = self.body_box
        width_pt = width_in * IN_TO_PT
        # each half-em unit ~ font_size * 0.5 pt wide
        return max(1, int(width_pt / (self.body_font_size * 0.5)))

    @property
    def line_pitch_pt(self) -> float:
        """Vertical distance between consecutive body lines, in points."""
        return self.body_font_size * LINE_SPACING * FONT_LINE_HEIGHT

    @property
    def max_body_lines(self) -> int:
        _, _, _, height_in = self.body_box
        height_pt = height_in * IN_TO_PT
        return max(1, int(height_pt / self.line_pitch_pt))

    @property
    def body_hang_units(self) -> int:
        """Hanging-indent width in half-em units (reduces wrapped-line width)."""
        return round(BODY_HANG_FACTOR / 0.5)


def wrap_line(text: str, max_units: int, cont_units: int | None = None) -> list[str]:
    """Word-wrap ``text`` to ``max_units`` half-em units, hard-splitting long tokens.

    When ``cont_units`` is given, the first line is capped at ``max_units`` and
    every subsequent (wrapped) line at ``cont_units`` — this models the hanging
    indent, where wrapped lines are narrower than the first.
    """
    if not text:
        return [""]

    def cap(n_done: int) -> int:
        return max_units if (n_done == 0 or cont_units is None) else cont_units

    lines: list[str] = []
    current = ""
    current_units = 0
    for token in re.split(r"(\s+)", text):
        if token == "":
            continue
        tu = _display_units(token)
        if token.isspace():
            if current:
                current += token
                current_units += tu
            continue
        limit = cap(len(lines))
        if current_units + tu <= limit or not current:
            if tu > limit and not current:
                # token itself longer than a line -> hard split
                for chunk in _hard_split(token, limit):
                    lines.append(chunk)
                current, current_units = "", 0
                continue
            current += token
            current_units += tu
        else:
            lines.append(current.rstrip())
            limit = cap(len(lines))
            if tu > limit:
                for chunk in _hard_split(token, limit):
                    lines.append(chunk)
                current, current_units = "", 0
            else:
                current, current_units = token, tu
    if current.strip():
        lines.append(current.rstrip())
    return lines or [""]


def _hard_split(token: str, max_units: int) -> list[str]:
    out: list[str] = []
    buf = ""
    units = 0
    for ch in token:
        cu = _display_units(ch)
        if units + cu > max_units and buf:
            out.append(buf)
            buf, units = ch, cu
        else:
            buf += ch
            units += cu
    if buf:
        out.append(buf)
    return out


def meaningful_title(title: str) -> bool:
    """True if the title contains any letter/number/CJK (not blank/punct-only)."""
    return bool(re.search(r"[0-9A-Za-z\uAC00-\uD7A3\u3040-\u30FF\u4E00-\u9FFF]", title or ""))


# --------------------------------------------------------------------------- #
# Pagination
# --------------------------------------------------------------------------- #
@dataclass
class RenderLine:
    text: str
    kind: str  # "chapter" | "verse"


@dataclass
class SlidePage:
    lines: list[RenderLine] = field(default_factory=list)


def _bundle_block(bundle: VerseBundle, with_marker: bool) -> list[RenderLine]:
    """Flatten one bundle to render lines (optional chapter marker + interleaved cells)."""
    lines: list[RenderLine] = []
    if with_marker:
        lines.append(RenderLine(text=f"<{bundle.coord.chapter}장>", kind="chapter"))
    for _code, cell in bundle.cells:
        if not cell.visible:
            continue
        lines.append(RenderLine(text=f"{cell.label}. {cell.text}", kind="verse"))
    return lines


def _blocks(bundles: list[VerseBundle]) -> list[list[RenderLine]]:
    """One indivisible render block per bundle; chapter markers decided once.

    A ``<N장>`` marker is attached to a bundle iff the passage spans multiple
    chapters and this bundle opens a new chapter (reading order). Markers are
    never duplicated across page breaks — they mark true change points only.
    """
    multi_chapter = len({b.coord.chapter for b in bundles}) > 1
    blocks: list[list[RenderLine]] = []
    prev_chapter: int | None = None
    for bundle in bundles:
        with_marker = multi_chapter and bundle.coord.chapter != prev_chapter
        blocks.append(_bundle_block(bundle, with_marker))
        prev_chapter = bundle.coord.chapter
    return blocks


def _line_count(line: RenderLine, style: SlideStyle) -> int:
    """Wrapped-line count for one render line, honouring the hanging indent."""
    max_units = style.max_units_per_line
    if line.kind == "verse":
        cont = max(1, max_units - style.body_hang_units)
        return len(wrap_line(line.text, max_units, cont))
    return len(wrap_line(line.text, max_units))


def _block_line_count(block: list[RenderLine], style: SlideStyle) -> int:
    return sum(_line_count(ln, style) for ln in block)


def fit_body_style(bundles: list[VerseBundle], style: SlideStyle) -> SlideStyle:
    """Auto-reduce the body font (up to ``MAX_BODY_SHRINK_PT``) so verses fit tidily.

    Picks the largest size within the shrink budget for which no single bundle
    overflows a slide, which both avoids downward overflow and reduces awkward
    single-verse-per-slide splits. If nothing fits, uses the smallest tried.
    """
    base = style.body_font_size
    floor = max(MIN_BODY_FONT_SIZE, base - MAX_BODY_SHRINK_PT)
    blocks = _blocks(bundles)
    smallest = style
    for size in range(base, floor - 1, -1):
        trial = replace(style, body_font_size=size)
        smallest = trial
        if all(_block_line_count(b, trial) <= trial.max_body_lines for b in blocks):
            return trial
    return smallest


def paginate(bundles: list[VerseBundle], style: SlideStyle) -> list[SlidePage]:
    """Greedy pagination that keeps each bundle intact on a single slide."""
    max_lines = style.max_body_lines

    pages: list[SlidePage] = []
    current = SlidePage()
    current_count = 0

    for block in _blocks(bundles):
        wrapped = _block_line_count(block, style)
        if current.lines and current_count + wrapped > max_lines:
            pages.append(current)
            current = SlidePage()
            current_count = 0
        current.lines.extend(block)
        current_count += wrapped
        # a single bundle taller than a whole slide sits alone and overflows
        if current_count > max_lines and current.lines == block:
            pages.append(current)
            current = SlidePage()
            current_count = 0

    if current.lines:
        pages.append(current)
    return pages


# --------------------------------------------------------------------------- #
# Rendering
# --------------------------------------------------------------------------- #
@dataclass
class PassageContent:
    title: str
    section_info: str
    bundles: list[VerseBundle]


def _set_font(run, name: str, size: int, *, bold: bool = False) -> None:
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    # ensure East-Asian text also uses the chosen family
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def _set_hanging_indent(p, hang_pt: float) -> None:
    """Outdent the first line by ``hang_pt`` so the verse number sticks out left
    and wrapped continuation lines align under the verse text."""
    hang = int(Pt(hang_pt))
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(hang))
    pPr.set("indent", str(-hang))


def _add_textbox(
    slide, box, text_lines, name, size, *,
    bold=False, align=PP_ALIGN.LEFT, line_spacing=None, hanging_pt=None,
):
    left, top, width, height = box
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for line in text_lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if line_spacing is not None:
            p.line_spacing = line_spacing
        if hanging_pt:
            _set_hanging_indent(p, hanging_pt)
        run = p.add_run()
        run.text = line
        _set_font(run, name, size, bold=bold)
    return tb


def _fit_single_line_size(text: str, box_width_in: float, base_size: int, min_size: int) -> int:
    """Largest size (≤ ``base_size``) at which ``text`` fits the box width on one line."""
    width_pt = box_width_in * IN_TO_PT
    units = _display_units(text)
    if units == 0:
        return base_size
    for size in range(base_size, min_size - 1, -1):
        if units * (size * 0.5) <= width_pt:
            return size
    return min_size


def _render_page(prs, page: SlidePage, passage: PassageContent, style: SlideStyle, background: Path | None):
    w, h = style.size_in
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank)

    if background is not None and Path(background).exists():
        slide.shapes.add_picture(
            str(background), Inches(0), Inches(0), width=Inches(w), height=Inches(h)
        )

    face = style.typeface
    title_w = style.title_box[2]
    has_title = meaningful_title(passage.title)
    if has_title:
        # shrink an over-long title so it never runs past the slide edge
        title_size = _fit_single_line_size(passage.title, title_w, TITLE_FONT_SIZE, MIN_TITLE_FONT_SIZE)
        _add_textbox(slide, style.title_box, [passage.title], face,
                     title_size, bold=True)
        _add_textbox(slide, style.section_box, [passage.section_info], face,
                     SECTION_FONT_SIZE, bold=True)
    else:
        # blank title -> put section info in the title position (task 13)
        sec_size = _fit_single_line_size(passage.section_info, title_w, TITLE_FONT_SIZE, MIN_TITLE_FONT_SIZE)
        _add_textbox(slide, style.title_box, [passage.section_info], face,
                     sec_size, bold=True)

    body_lines = [ln.text for ln in page.lines]
    _add_textbox(
        slide, style.body_box, body_lines, face, style.body_font_size,
        bold=style.body_bold,
        line_spacing=LINE_SPACING,
        hanging_pt=style.body_font_size * BODY_HANG_FACTOR,
    )


def render(
    passages: list[PassageContent],
    style: SlideStyle,
    background: Path | None,
) -> Presentation:
    """Render one or more passages into a single presentation."""
    prs = Presentation()
    w, h = style.size_in
    prs.slide_width = Inches(w)
    prs.slide_height = Inches(h)
    for passage in passages:
        # auto-fit the body font per passage so verses fill without overflowing
        passage_style = fit_body_style(passage.bundles, style)
        pages = paginate(passage.bundles, passage_style)
        for page in pages:
            _render_page(prs, page, passage, passage_style, background)
    return prs


def save(prs: Presentation, out_path: str | Path) -> Path:
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return out
