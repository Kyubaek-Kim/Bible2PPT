# Bible2PPT — E2E Test Plan (PR #1)

App: Tkinter desktop app, launched with `DISPLAY=:0 /usr/bin/python3 main.py` (system Python ships Tk).
Rendering proof: generated `.pptx` opened via LibreOffice → PDF → PNG.

## What changed (user-visible)
Monolithic script → layered core/ui desktop app. New: multi-translation selection with
**language annotated in parentheses**, cross-chapter references, multi-passage list, aspect/font/size
options + font preview, background crop, separate/combined PPT generation. New cute app icon.

## Code paths grounding the plan
- Translation label with language: `core/i18n.py:101 translation_label()` → used in `ui/app.py:220 _refresh_translation_list`.
- Add passage: `ui/app.py:349 _add_passage` (validates via `generator.make_parser().parse`).
- Chapter marker `<N장>`: `core/ppt.py:174`; verse line `{label}. {text}`: `core/ppt.py:178`.
- Blank title → section info in title box: `core/ppt.py:284`.
- Generate: `ui/app.py:_generate` → `core/generator.generate`, success dialog + open folder.

## Primary flow (record this)

### T1 — Translation options show language in parentheses (user's explicit request)
- Action: launch app; read the translation list.
- PASS: entries read like `개역한글 (한국어)`, `King James Version (영어)`, `Textus Receptus (NT) (헬라어)`.
- FAIL: any entry shows a name with no `(언어)` suffix.
- Would-look-different-if-broken: without the change, labels would be bare names (no parentheses).

### T2 — Select 2 translations + add a cross-chapter passage with a title
- Action: select `개역한글 (한국어)` and `King James Version (영어)` in the list; type `창 1:30-2:3` in 직접 입력, title `창조`; click 담기.
- PASS: list shows one item `창조 — 창세기 1:30-2:3`.
- FAIL: no item added, or reference text unparsed/garbled.

### T3 — Add a second passage with blank title
- Action: type `요 3:16`, leave title empty, click 담기.
- PASS: list shows a second item `요한복음 3:16` (localized ref, no title prefix).

### T4 — Font preview reacts to font change
- Action: change 글자체 dropdown to a different family.
- PASS: preview sample text re-renders in the newly selected family (visible change).

### T5 — Generate (separate mode) + verify output slides
- Action: ensure 생성 방식 = 구절별 개별 PPT; click 생성; on success dialog click to open folder.
- PASS (dialog): success dialog reports a save path; output folder opens.
- PASS (slide content, opened in LibreOffice):
  - 창세기 file: page 1 title `창조`, section `창세기 1:30-2:3`, marker `<1장>`, verse `30. ` for **both** KRV (Korean) then KJV (English) — proving interleave and verse-number prefix.
  - marker `<2장>` appears exactly once where chapter turns from 1→2 (not repeated per page).
  - 요한복음 file: **no** title text; section `요한복음 3:16` sits in the bold title position (blank-title rule).
- FAIL: missing verse numbers, only one translation shown, `<2장>` missing/duplicated, or blank-title page shows an empty title band with section pushed down.

## Icon check (asset, not runtime)
- Show `run_icon.png`: cute smiling open-book on a slide/screen with sparkles (conveys "verses → PPT automatically").

## Out of scope
Windows .exe build, macOS, background-image crop dialog (optional if time permits as regression).
