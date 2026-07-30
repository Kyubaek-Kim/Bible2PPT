"""Tkinter GUI for Bible2PPT.

A phone-sized, vertically-scrolling window. All heavy lifting lives in
:mod:`core`; this module only wires widgets to that logic and persists UI state
via :class:`core.settings.Settings`. Only Tkinter standard widgets are used so
the dialogs stay OS-independent; OS-specific actions (open folder, register
font) go through :mod:`core.platform_util`.
"""
from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core import (
    bible,
    fonts,
    generator,
    image_util,
    importer,
    paths,
    platform_util,
    ppt,
)
from core.i18n import I18n
from core.settings import Settings

# Kept short vertically so the window fits inside a small laptop screen; the
# body scrolls and the generate button is pinned in an always-visible footer.
WINDOW_SIZE = "440x660"
FONT_SIZES = [16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60]
ASPECTS = list(ppt.ASPECT_RATIOS.keys())


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.settings = Settings.load()
        self.registry = bible.Registry.load()
        self.i18n = I18n(self.settings.ui_language)
        fonts.register_bundled_fonts()
        if self.settings.font not in fonts.font_dropdown_values():
            self.settings.font = fonts.default_font_name()

        self.passages: list[generator.PassageInput] = []
        # widgets whose text depends on the UI language: (widget, key)
        self._i18n_widgets: list[tuple[tk.Widget, str]] = []
        self._trans_index_to_code: list[str] = []

        self.title(self.i18n.t("app_title"))
        self.geometry(WINDOW_SIZE)
        self.minsize(400, 480)

        self._build_footer()
        self._build_scroll_container()
        self._build_all_sections()
        self._load_state()
        self._bind_mousewheel()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ #
    # Scaffolding
    # ------------------------------------------------------------------ #
    def _build_footer(self) -> None:
        """A fixed footer, pinned to the bottom of the window (outside the
        scroll area) so the generate button is always reachable regardless of
        scroll position or window height (item 4)."""
        footer = ttk.Frame(self, relief="raised", borderwidth=1)
        footer.pack(side="bottom", fill="x")
        self.generate_btn = ttk.Button(
            footer, text=self.i18n.t("generate"), command=self._generate
        )
        self.generate_btn.pack(fill="x", padx=8, pady=6)
        self._i18n_widgets.append((self.generate_btn, "generate"))

    def _build_scroll_container(self) -> None:
        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0)
        vbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.body = ttk.Frame(canvas)
        self.body.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        self._win = canvas.create_window((0, 0), window=self.body, anchor="nw")
        canvas.bind(
            "<Configure>", lambda e: canvas.itemconfigure(self._win, width=e.width)
        )
        self.canvas = canvas
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

    def _bind_mousewheel(self) -> None:
        """Route wheel scrolling to the outer canvas — but only when the pointer
        is *not* over a widget that scrolls itself (a Listbox or a Text). Binding
        per-widget instead of ``bind_all`` lets those inner widgets (and open
        Combobox popdowns, which are separate toplevels we never bind) keep their
        own wheel behaviour, so the outer view no longer scrolls along with
        them."""
        for widget in self._wheelable_widgets(self.canvas):
            widget.bind("<MouseWheel>", self._on_mousewheel)
            widget.bind("<Button-4>", self._on_mousewheel)
            widget.bind("<Button-5>", self._on_mousewheel)

    def _wheelable_widgets(self, root: tk.Widget) -> list[tk.Widget]:
        collected: list[tk.Widget] = []
        stack: list[tk.Widget] = [root]
        while stack:
            w = stack.pop()
            if isinstance(w, (tk.Listbox, tk.Text)):
                continue  # self-scrolling widgets keep the wheel to themselves
            collected.append(w)
            stack.extend(w.winfo_children())
        return collected

    def _on_mousewheel(self, event) -> str:
        num = getattr(event, "num", 0)
        if num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif num == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _section(self, key: str) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(self.body, text=self.i18n.t(key))
        frame.pack(fill="x", padx=8, pady=5)
        self._i18n_widgets.append((frame, key))
        return frame

    def _label(self, parent, key: str, **grid) -> ttk.Label:
        lbl = ttk.Label(parent, text=self.i18n.t(key))
        if grid:
            lbl.grid(**grid)
        self._i18n_widgets.append((lbl, key))
        return lbl

    def _button(self, parent, key: str, command, **pack) -> ttk.Button:
        btn = ttk.Button(parent, text=self.i18n.t(key), command=command)
        if pack:
            btn.pack(**pack)
        self._i18n_widgets.append((btn, key))
        return btn

    def _build_all_sections(self) -> None:
        self._build_language_section()
        self._build_translations_section()
        self._build_passage_section()
        self._build_list_section()
        self._build_mode_section()  # item 3: right below the passage list
        self._build_options_section()
        self._build_background_section()
        self._build_output_section()

    # ------------------------------------------------------------------ #
    # Language
    # ------------------------------------------------------------------ #
    def _build_language_section(self) -> None:
        sec = self._section("ui_language")
        self._lang_codes = [c for c, _ in self.i18n.available_langs()]
        names = [n for _, n in self.i18n.available_langs()]
        self.lang_var = tk.StringVar()
        self.lang_combo = ttk.Combobox(
            sec, values=names, state="readonly", textvariable=self.lang_var
        )
        self.lang_combo.pack(fill="x", padx=6, pady=4)
        self.lang_combo.bind("<<ComboboxSelected>>", self._on_language_change)

    def _on_language_change(self, _evt=None) -> None:
        idx = self.lang_combo.current()
        if idx < 0:
            return
        self.i18n.set_lang(self._lang_codes[idx])
        self.settings.ui_language = self.i18n.lang
        self._retranslate()

    def _retranslate(self) -> None:
        self.title(self.i18n.t("app_title"))
        for widget, key in self._i18n_widgets:
            try:
                widget.configure(text=self.i18n.t(key))
            except tk.TclError:
                pass
        self._refresh_translation_list()
        self._refresh_book_dropdown()
        self._refresh_passage_list()
        self._update_font_preview()

    # ------------------------------------------------------------------ #
    # Translations (multi-select, language-annotated)
    # ------------------------------------------------------------------ #
    def _build_translations_section(self) -> None:
        sec = self._section("translations")
        head = ttk.Frame(sec)
        head.pack(fill="x", padx=6)
        self._label(head, "select_translations").pack(side="left", anchor="w")
        base_lbl = ttk.Label(head, text=self.i18n.t("base_translation"))
        base_lbl.pack(side="right")
        self._i18n_widgets.append((base_lbl, "base_translation"))
        # Pinned "자주 사용" (favourite) rows stay visible; the rest hide behind
        # a "더보기" toggle so a long translation list isn't overwhelming.
        self.fav_rows = ttk.Frame(sec)
        self.fav_rows.pack(fill="x", padx=6, pady=(4, 0))
        self.more_btn = ttk.Button(sec, text="", command=self._toggle_more)
        self.more_btn.pack(anchor="w", padx=6, pady=2)
        self.more_rows = ttk.Frame(sec)
        self.more_rows.pack(fill="x", padx=6)
        self._more_expanded = False

        self.base_trans_var = tk.StringVar(value=self.settings.default_translation)
        self._trans_check_vars: dict[str, tk.BooleanVar] = {}
        self._trans_fav_vars: dict[str, tk.BooleanVar] = {}
        self._trans_radios: dict[str, ttk.Radiobutton] = {}
        base_hint = ttk.Label(
            sec, text=self.i18n.t("base_translation_hint"),
            foreground="#666", wraplength=380, justify="left",
        )
        base_hint.pack(anchor="w", padx=6)
        self._i18n_widgets.append((base_hint, "base_translation_hint"))
        self._button(sec, "import_bible", self._open_import_dialog).pack(
            anchor="w", padx=6, pady=(0, 4)
        )

    def _toggle_more(self) -> None:
        self._more_expanded = not self._more_expanded
        self._refresh_translation_list()

    def _make_trans_row(self, parent, code: str, name: str, language: str) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x")
        ttk.Checkbutton(
            row,
            text=self.i18n.translation_label(name, language),
            variable=self._trans_check_vars[code],
            command=lambda c=code: self._on_translation_toggle(c),
        ).pack(side="left", anchor="w")
        rb = ttk.Radiobutton(
            row, text=self.i18n.t("base_mark"), value=code,
            variable=self.base_trans_var,
            command=self._on_base_translation_change,
        )
        rb.pack(side="right")
        self._trans_radios[code] = rb
        ttk.Checkbutton(
            row, text=self.i18n.t("favorite"), variable=self._trans_fav_vars[code],
            command=lambda c=code: self._on_favorite_toggle(c),
        ).pack(side="right", padx=(0, 8))

    def _refresh_translation_list(self) -> None:
        for child in self.fav_rows.winfo_children():
            child.destroy()
        for child in self.more_rows.winfo_children():
            child.destroy()
        self._trans_radios = {}

        metas = {m.code: m for m in self.registry.list_meta()}
        self._trans_index_to_code = list(metas.keys())
        # Vars exist for *every* translation (even ones hidden behind "더보기")
        # so selection / favourite state survives collapsing the list.
        self._trans_check_vars = {
            c: tk.BooleanVar(value=c in self.settings.selected_translations)
            for c in metas
        }
        self._trans_fav_vars = {
            c: tk.BooleanVar(value=c in self.settings.favorite_translations)
            for c in metas
        }
        favorites = [c for c in self.settings.favorite_translations if c in metas]
        others = [c for c in metas if c not in favorites]

        if favorites:
            self.fav_rows.pack_propagate(True)
            for code in favorites:
                m = metas[code]
                self._make_trans_row(self.fav_rows, code, m.name, m.language)
        else:
            self.fav_rows.configure(height=1)
            self.fav_rows.pack_propagate(False)
        if self._more_expanded:
            # let the frame grow to fit the rows again after having been pinned
            # to height 1 while collapsed
            self.more_rows.pack_propagate(True)
            for code in others:
                m = metas[code]
                self._make_trans_row(self.more_rows, code, m.name, m.language)
        else:
            # an emptied ttk.Frame keeps its last requested height, so the
            # section wouldn't shrink back on "접기"; force it to collapse.
            self.more_rows.configure(height=1)
            self.more_rows.pack_propagate(False)

        key = "show_less" if self._more_expanded else "show_more"
        self.more_btn.configure(text=self.i18n.t(key, count=len(others)))
        if not others:
            self.more_btn.pack_forget()
        else:
            self.more_btn.pack(anchor="w", padx=6, pady=2, before=self.more_rows)

        self._sync_base_radio_state()
        if hasattr(self, "canvas"):
            self._bind_mousewheel()

    def _on_favorite_toggle(self, code: str) -> None:
        self.settings.set_favorite(code, self._trans_fav_vars[code].get())
        self._refresh_translation_list()

    def _sync_base_radio_state(self) -> None:
        """Only *selected* translations may serve as the base: disable the other
        radios and keep the base pointing at a checked translation."""
        selected = [
            c for c in self._trans_index_to_code if self._trans_check_vars[c].get()
        ]
        for code, rb in self._trans_radios.items():
            rb.configure(state=("normal" if code in selected else "disabled"))
        base = self.base_trans_var.get()
        if base not in selected:
            base = selected[0] if selected else ""
            self.base_trans_var.set(base)
        self.settings.default_translation = base

    def _selected_translation_codes(self) -> list[str]:
        codes = [
            c for c in self._trans_index_to_code if self._trans_check_vars[c].get()
        ]
        # put the base translation first so interleave starts with it
        default = self.settings.default_translation
        if default in codes:
            codes = [default] + [c for c in codes if c != default]
        return codes

    def _apply_base_change(self) -> None:
        self._refresh_book_dropdown()
        if self.book_combo.current() >= 0:
            self._on_book_change()

    def _on_translation_toggle(self, code: str) -> None:
        self.settings.selected_translations = [
            c for c in self._trans_index_to_code if self._trans_check_vars[c].get()
        ]
        prev_base = self.settings.default_translation
        self._sync_base_radio_state()
        if self.settings.default_translation != prev_base:
            self._apply_base_change()

    def _on_base_translation_change(self) -> None:
        self.settings.default_translation = self.base_trans_var.get()
        self._apply_base_change()

    def _default_translation(self) -> bible.Translation | None:
        return self.registry.get(self.settings.default_translation) or (
            self.registry.get(self._trans_index_to_code[0])
            if self._trans_index_to_code
            else None
        )

    # ------------------------------------------------------------------ #
    # Passage input (cascading dropdowns + direct entry + title)
    # ------------------------------------------------------------------ #
    def _build_passage_section(self) -> None:
        sec = self._section("select_book")
        grid = ttk.Frame(sec)
        grid.pack(fill="x", padx=6, pady=4)

        self._label(grid, "book", row=0, column=0, sticky="w")
        self.book_var = tk.StringVar()
        self.book_combo = ttk.Combobox(grid, state="readonly", textvariable=self.book_var, width=16)
        self.book_combo.grid(row=0, column=1, sticky="ew", pady=1)
        self.book_combo.bind("<<ComboboxSelected>>", lambda e: self._on_book_change())

        self._label(grid, "chapter", row=1, column=0, sticky="w")
        self.chapter_var = tk.StringVar()
        self.chapter_combo = ttk.Combobox(grid, state="readonly", textvariable=self.chapter_var, width=16)
        self.chapter_combo.grid(row=1, column=1, sticky="ew", pady=1)
        self.chapter_combo.bind("<<ComboboxSelected>>", lambda e: self._on_chapter_change())

        self._label(grid, "verse_start", row=2, column=0, sticky="w")
        self.vstart_var = tk.StringVar()
        self.vstart_combo = ttk.Combobox(grid, state="readonly", textvariable=self.vstart_var, width=16)
        self.vstart_combo.grid(row=2, column=1, sticky="ew", pady=1)
        self.vstart_combo.bind("<<ComboboxSelected>>", lambda e: self._on_vstart_change())

        self._label(grid, "verse_end", row=3, column=0, sticky="w")
        self.vend_var = tk.StringVar()
        self.vend_combo = ttk.Combobox(grid, state="readonly", textvariable=self.vend_var, width=16)
        self.vend_combo.grid(row=3, column=1, sticky="ew", pady=1)
        grid.columnconfigure(1, weight=1)

        # direct input
        self._label(sec, "direct_input").pack(anchor="w", padx=6)
        self.direct_var = tk.StringVar()
        self.direct_entry = ttk.Entry(sec, textvariable=self.direct_var)
        self.direct_entry.pack(fill="x", padx=6)
        self.direct_hint = ttk.Label(sec, text=self.i18n.t("direct_input_hint"), foreground="#666")
        self.direct_hint.pack(anchor="w", padx=6)
        self._i18n_widgets.append((self.direct_hint, "direct_input_hint"))

        # title
        self._label(sec, "title").pack(anchor="w", padx=6, pady=(4, 0))
        self.title_var = tk.StringVar()
        self.title_entry = ttk.Entry(sec, textvariable=self.title_var)
        self.title_entry.pack(fill="x", padx=6)
        self.title_hint = ttk.Label(sec, text=self.i18n.t("title_hint"), foreground="#666")
        self.title_hint.pack(anchor="w", padx=6)
        self._i18n_widgets.append((self.title_hint, "title_hint"))

        self._button(sec, "add_to_list", self._add_passage).pack(anchor="e", padx=6, pady=4)

    def _refresh_book_dropdown(self) -> None:
        self._book_ids = [b.id for b in bible.load_canon()]
        names = [self.i18n.book_name(bid) for bid in self._book_ids]
        self.book_combo.configure(values=names)

    def _current_book_id(self) -> str | None:
        idx = self.book_combo.current()
        return self._book_ids[idx] if idx >= 0 else None

    def _chapters_for_book(self, book_id: str) -> list[int]:
        trans = self._default_translation()
        chapters = (trans.chapters(book_id) if trans else []) or list(
            bible.canonical_versification().get(book_id, {}).keys()
        )
        return sorted(int(c) for c in chapters)

    def _verses_for_chapter(self, book_id: str, chapter: int) -> list[int]:
        trans = self._default_translation()
        verses = trans.verses(book_id, chapter) if trans else []
        if not verses:
            last = bible.canonical_versification().get(book_id, {}).get(chapter, 0)
            verses = list(range(1, last + 1))
        return verses

    def _on_book_change(self) -> None:
        book_id = self._current_book_id()
        if book_id is None:
            return
        self.chapter_combo.configure(
            values=[str(c) for c in self._chapters_for_book(book_id)]
        )
        self.chapter_var.set("")
        self.vstart_combo.configure(values=[])
        self.vend_combo.configure(values=[])
        self.vstart_var.set("")
        self.vend_var.set("")

    def _current_verses(self) -> list[int]:
        book_id = self._current_book_id()
        if book_id is None or not self.chapter_var.get():
            return []
        return self._verses_for_chapter(book_id, int(self.chapter_var.get()))

    def _on_chapter_change(self) -> None:
        self.vstart_combo.configure(values=[str(v) for v in self._current_verses()])
        self.vend_combo.configure(values=[])
        self.vstart_var.set("")
        self.vend_var.set("")

    def _on_vstart_change(self) -> None:
        book_id = self._current_book_id()
        if book_id is None or not self.chapter_var.get() or not self.vstart_var.get():
            return
        start_ch = int(self.chapter_var.get())
        start_v = int(self.vstart_var.get())
        # The end-verse list runs from the start verse to the end of the start
        # chapter, then continues into the following chapters labelled "장:절",
        # so a cross-chapter range like 창 3:24-4:5 can be picked from dropdowns.
        options = [
            str(v) for v in self._verses_for_chapter(book_id, start_ch) if v >= start_v
        ]
        for ch in self._chapters_for_book(book_id):
            if ch <= start_ch:
                continue
            options.extend(f"{ch}:{v}" for v in self._verses_for_chapter(book_id, ch))
        self.vend_combo.configure(values=options)
        self.vend_var.set("")

    # ------------------------------------------------------------------ #
    # Passage list
    # ------------------------------------------------------------------ #
    def _build_list_section(self) -> None:
        sec = self._section("passage_list")
        self.passage_listbox = tk.Listbox(sec, height=5)
        self.passage_listbox.pack(fill="x", padx=6, pady=4)
        row = ttk.Frame(sec)
        row.pack(fill="x", padx=6, pady=(0, 4))
        self._button(row, "remove_selected", self._remove_passage).pack(side="left")
        self._button(row, "clear_list", self._clear_passages).pack(side="left", padx=4)

    def _build_reference_text(self) -> str | None:
        direct = self.direct_var.get().strip()
        if direct:
            return direct
        book_id = self._current_book_id()
        if book_id is None or not self.chapter_var.get() or not self.vstart_var.get():
            return None
        ref = f"{book_id} {self.chapter_var.get()}:{self.vstart_var.get()}"
        end = self.vend_var.get().strip()
        if end:
            ref += f"-{end}"  # "-25" (same chapter) or "-4:5" (cross-chapter)
        return ref

    def _add_passage(self) -> None:
        ref_text = self._build_reference_text()
        if not ref_text:
            messagebox.showwarning(self.i18n.t("warning"), self.i18n.t("parse_failed", text=""))
            return
        # validate parse before adding
        try:
            generator.make_parser().parse(ref_text)
        except Exception:
            messagebox.showwarning(
                self.i18n.t("warning"), self.i18n.t("parse_failed", text=ref_text)
            )
            return
        title = self.title_var.get()
        self.passages.append(generator.PassageInput(reference_text=ref_text, title=title))
        self._refresh_passage_list()
        self.direct_var.set("")
        self.title_var.set("")

    def _refresh_passage_list(self) -> None:
        self.passage_listbox.delete(0, "end")
        parser = generator.make_parser()
        for p in self.passages:
            try:
                ref = parser.parse(p.reference_text)
                label = generator.format_reference(ref, self.i18n)
            except Exception:
                label = p.reference_text
            if p.title and ppt.meaningful_title(p.title):
                label = f"{p.title} — {label}"
            self.passage_listbox.insert("end", label)

    def _remove_passage(self) -> None:
        sel = self.passage_listbox.curselection()
        if sel:
            del self.passages[sel[0]]
            self._refresh_passage_list()

    def _clear_passages(self) -> None:
        self.passages.clear()
        self._refresh_passage_list()

    # ------------------------------------------------------------------ #
    # Options (aspect, font + preview, body font size)
    # ------------------------------------------------------------------ #
    def _build_options_section(self) -> None:
        sec = self._section("display_settings")
        grid = ttk.Frame(sec)
        grid.pack(fill="x", padx=6, pady=4)

        self._label(grid, "aspect_ratio", row=0, column=0, sticky="w")
        self.aspect_var = tk.StringVar()
        self.aspect_combo = ttk.Combobox(grid, state="readonly", values=ASPECTS, textvariable=self.aspect_var, width=14)
        self.aspect_combo.grid(row=0, column=1, sticky="w", pady=1)
        self.aspect_combo.bind("<<ComboboxSelected>>", lambda e: self._on_option_change())

        self._label(grid, "font", row=1, column=0, sticky="w")
        self.font_var = tk.StringVar()
        self.font_combo = ttk.Combobox(grid, state="readonly", textvariable=self.font_var, width=20)
        self.font_combo.grid(row=1, column=1, sticky="w", pady=1)
        self.font_combo.bind("<<ComboboxSelected>>", lambda e: self._on_font_change())

        self._label(grid, "body_font_size", row=2, column=0, sticky="w")
        self.size_var = tk.StringVar()
        self.size_combo = ttk.Combobox(grid, state="readonly", values=[str(s) for s in FONT_SIZES], textvariable=self.size_var, width=14)
        self.size_combo.grid(row=2, column=1, sticky="w", pady=1)
        self.size_combo.bind("<<ComboboxSelected>>", lambda e: self._on_option_change())

        # body bold (item 5): user-toggled independent of the font's own weight
        self.body_bold_var = tk.BooleanVar(value=self.settings.body_bold)
        self.body_bold_chk = ttk.Checkbutton(
            grid, text=self.i18n.t("body_bold"), variable=self.body_bold_var,
            command=self._on_body_bold_change,
        )
        self.body_bold_chk.grid(row=3, column=1, sticky="w", pady=1)
        self._i18n_widgets.append((self.body_bold_chk, "body_bold"))

        # font preview
        self._label(sec, "font_preview").pack(anchor="w", padx=6, pady=(4, 0))
        self.preview_frame = ttk.Frame(sec, relief="solid", borderwidth=1)
        self.preview_frame.pack(fill="x", padx=6, pady=2)
        self.preview_title = tk.Label(self.preview_frame, anchor="w", justify="left")
        self.preview_title.pack(fill="x", padx=6, pady=(4, 0))
        self.preview_body = tk.Label(self.preview_frame, anchor="w", justify="left", wraplength=360)
        self.preview_body.pack(fill="x", padx=6, pady=(0, 4))
        self.preview_note = ttk.Label(sec, text=self.i18n.t("font_preview_note"), foreground="#888", wraplength=380)
        self.preview_note.pack(anchor="w", padx=6)
        self._i18n_widgets.append((self.preview_note, "font_preview_note"))
        self.font_hint = ttk.Label(sec, text="", foreground="#a33", wraplength=380)
        self.font_hint.pack(anchor="w", padx=6)

        # item 6: open the drag-and-drop layout editor
        self._button(sec, "customize_layout", self._open_layout_dialog).pack(
            fill="x", padx=6, pady=(6, 4)
        )

    def _on_option_change(self) -> None:
        if self.aspect_var.get():
            self.settings.aspect_ratio = self.aspect_var.get()
        if self.size_var.get():
            self.settings.body_font_size = int(self.size_var.get())

    def _on_body_bold_change(self) -> None:
        self.settings.body_bold = self.body_bold_var.get()

    def _build_style(self) -> ppt.SlideStyle:
        """Assemble a :class:`ppt.SlideStyle` from the current settings,
        including any saved layout customisation (item 6)."""
        s = self.settings
        return ppt.SlideStyle(
            aspect=s.aspect_ratio,
            font_name=s.font or fonts.default_font_name(),
            body_font_size=s.body_font_size,
            body_bold_opt=s.body_bold,
            title_font_name=s.title_font,
            title_font_size=s.title_font_size,
            title_bold=s.title_bold,
            section_font_name=s.section_font,
            section_font_size=s.section_font_size,
            section_bold=s.section_bold,
            layout_boxes=dict(s.layout_boxes),
        )

    def _open_layout_dialog(self) -> None:
        LayoutDialog(self)

    def _on_font_change(self) -> None:
        self.settings.font = self.font_var.get()
        self._update_font_preview()

    def _update_font_preview(self) -> None:
        name = self.font_var.get() or fonts.default_font_name()
        available, hint = fonts.ensure_font_available(name, self)
        self.font_hint.configure(text="" if available else hint)
        choice = fonts.resolve(name)
        families = fonts.system_font_families(self)
        family = next((f for f in (choice.typeface, choice.label) if f in families), None)
        body_weight = "bold" if choice.bold else "normal"
        try:
            if family is None:
                raise tk.TclError
            title_font = tkfont.Font(family=family, size=18, weight="bold")
            body_font = tkfont.Font(family=family, size=14, weight=body_weight)
        except tk.TclError:
            title_font = tkfont.Font(size=18, weight="bold")
            body_font = tkfont.Font(size=14, weight=body_weight)
        self.preview_title.configure(text=self.i18n.t("font_preview_sample_title"), font=title_font)
        self.preview_body.configure(
            text="1. 태초에 하나님이 천지를 창조하시니라 / In the beginning God created the heaven and the earth.",
            font=body_font,
        )

    # ------------------------------------------------------------------ #
    # Background
    # ------------------------------------------------------------------ #
    def _build_background_section(self) -> None:
        sec = self._section("background")
        row = ttk.Frame(sec)
        row.pack(fill="x", padx=6, pady=4)
        self._button(row, "background_default", self._use_default_background).pack(side="left")
        self._button(row, "background_custom", self._attach_background).pack(side="left", padx=4)
        self.bg_label = ttk.Label(sec, text="", foreground="#555", wraplength=380)
        self.bg_label.pack(anchor="w", padx=6)

        self._label(sec, "background_history").pack(anchor="w", padx=6, pady=(4, 0))
        self.bg_hist_var = tk.StringVar()
        self.bg_hist_combo = ttk.Combobox(sec, state="readonly", textvariable=self.bg_hist_var)
        self.bg_hist_combo.pack(fill="x", padx=6, pady=(0, 4))
        self.bg_hist_combo.bind("<<ComboboxSelected>>", lambda e: self._on_history_background())

    def _slide_cm(self) -> tuple[float, float]:
        w_in, h_in = ppt.ASPECT_RATIOS[self.settings.aspect_ratio]
        return w_in * 2.54, h_in * 2.54

    def _use_default_background(self) -> None:
        self.settings.background = ""
        self._refresh_background_label()

    def _apply_background_with_confirm(self, source: str) -> None:
        w_cm, h_cm = self._slide_cm()
        plan = image_util.plan_crop(source, w_cm, h_cm)
        if plan.needs_crop:
            axis = "상/하" if plan.axis == "vertical" else "좌/우"
            msg = self.i18n.t("crop_confirm_body", axis=axis, px=plan.crop_px, cm=plan.crop_cm)
            if not messagebox.askokcancel(self.i18n.t("crop_confirm_title"), msg):
                return
        stored = image_util.add_to_history(source)
        cropped = paths.background_history_dir() / f"cropped_{stored.name}"
        image_util.apply_crop(source, plan, cropped)
        self.settings.background = str(cropped)
        self.settings.add_background_history(str(stored))
        self._refresh_background_label()
        self._refresh_background_history()

    def _attach_background(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Image", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All", "*.*")]
        )
        if path:
            self._apply_background_with_confirm(path)

    def _on_history_background(self) -> None:
        idx = self.bg_hist_combo.current()
        if 0 <= idx < len(self.settings.background_history):
            self._apply_background_with_confirm(self.settings.background_history[idx])

    def _refresh_background_history(self) -> None:
        self.bg_hist_combo.configure(
            values=[Path(p).name for p in self.settings.background_history]
        )

    def _refresh_background_label(self) -> None:
        self.bg_label.configure(text=str(self.settings.resolved_background()))

    # ------------------------------------------------------------------ #
    # Output folder
    # ------------------------------------------------------------------ #
    def _build_output_section(self) -> None:
        sec = self._section("output_folder")
        self.output_label = ttk.Label(sec, text="", foreground="#555", wraplength=380)
        self.output_label.pack(anchor="w", padx=6, pady=(4, 0))
        row = ttk.Frame(sec)
        row.pack(fill="x", padx=6, pady=4)
        self._button(row, "change_output_folder", self._change_output_folder).pack(side="left")
        self._button(row, "open_output_folder", self._open_output_folder).pack(side="left", padx=4)

    def _change_output_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.settings.output_folder = folder
            self._refresh_output_label()

    def _open_output_folder(self) -> None:
        folder = self.settings.resolved_output_folder()
        folder.mkdir(parents=True, exist_ok=True)
        platform_util.open_folder(folder)

    def _refresh_output_label(self) -> None:
        self.output_label.configure(text=str(self.settings.resolved_output_folder()))

    # ------------------------------------------------------------------ #
    # Generate
    # ------------------------------------------------------------------ #
    def _build_mode_section(self) -> None:
        sec = self._section("generate_mode")
        self.mode_var = tk.StringVar(value=self.settings.generate_mode)
        self.mode_sep = ttk.Radiobutton(sec, text=self.i18n.t("mode_separate"), value="separate", variable=self.mode_var, command=self._on_mode_change)
        self.mode_sep.pack(anchor="w", padx=6, pady=(4, 0))
        self.mode_comb = ttk.Radiobutton(sec, text=self.i18n.t("mode_combined"), value="combined", variable=self.mode_var, command=self._on_mode_change)
        self.mode_comb.pack(anchor="w", padx=6, pady=(0, 4))
        self._i18n_widgets.append((self.mode_sep, "mode_separate"))
        self._i18n_widgets.append((self.mode_comb, "mode_combined"))

    def _on_mode_change(self) -> None:
        self.settings.generate_mode = self.mode_var.get()

    def _generate(self) -> None:
        if not self.passages:
            messagebox.showwarning(self.i18n.t("warning"), self.i18n.t("no_passages"))
            return
        codes = self._selected_translation_codes()
        if not codes:
            messagebox.showwarning(self.i18n.t("warning"), self.i18n.t("no_translation"))
            return
        style = self._build_style()
        try:
            result = generator.generate(
                self.passages,
                registry=self.registry,
                translation_codes=codes,
                style=style,
                background=self.settings.resolved_background(),
                output_folder=self.settings.resolved_output_folder(),
                mode=self.settings.generate_mode,
                i18n=self.i18n,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(self.i18n.t("error"), str(exc))
            return

        if result.errors:
            detail = "\n".join(msg for _, msg in result.errors)
            messagebox.showwarning(self.i18n.t("warning"), detail)
        if result.output_paths:
            folder = result.output_paths[0].parent
            msg = self.i18n.t("generated_saved", path=str(folder))
            if messagebox.askyesno(self.i18n.t("done"), f"{msg}\n\n{self.i18n.t('open_output_folder')}?"):
                platform_util.open_folder(folder)

    # ------------------------------------------------------------------ #
    # Import dialog
    # ------------------------------------------------------------------ #
    def _open_import_dialog(self) -> None:
        ImportDialog(self)

    # ------------------------------------------------------------------ #
    # State load / persist
    # ------------------------------------------------------------------ #
    def _load_state(self) -> None:
        # language combo
        if self.i18n.lang in self._lang_codes:
            self.lang_combo.current(self._lang_codes.index(self.i18n.lang))
        self._refresh_translation_list()
        self._refresh_book_dropdown()
        self.aspect_var.set(self.settings.aspect_ratio)
        # font list
        self.font_combo.configure(values=fonts.font_dropdown_values(self))
        self.font_var.set(self.settings.font or fonts.default_font_name())
        self.size_var.set(str(self.settings.body_font_size))
        self.mode_var.set(self.settings.generate_mode)
        self._refresh_background_label()
        self._refresh_background_history()
        self._refresh_output_label()
        self._refresh_passage_list()
        self._update_font_preview()

    def _on_close(self) -> None:
        try:
            self.settings.save()
        finally:
            self.destroy()


class ImportDialog(tk.Toplevel):
    """Two-phase import: parse+review, then register (gated on a passing review)."""

    def __init__(self, master: App) -> None:
        super().__init__(master)
        self.app = master
        self.i18n = master.i18n
        self.report: importer.ImportReport | None = None
        self.source_path: str | None = None

        self.title(self.i18n.t("import_bible"))
        self.geometry("520x520")

        ttk.Button(self, text=self.i18n.t("import_bible"), command=self._pick).pack(anchor="w", padx=8, pady=6)
        ttk.Label(self, text=self.i18n.t("import_format_hint"), foreground="#666",
                  justify="left", wraplength=490).pack(anchor="w", padx=8)
        self.stats = ttk.Label(self, text="", justify="left", wraplength=490)
        self.stats.pack(anchor="w", padx=8)
        self.problems = tk.Text(self, height=12, wrap="word")
        self.problems.pack(fill="both", expand=True, padx=8, pady=4)

        form = ttk.Frame(self)
        form.pack(fill="x", padx=8, pady=4)
        self.name_var = tk.StringVar()
        self.lang_var = tk.StringVar(value="ko")
        self.abbr_var = tk.StringVar()
        self.code_var = tk.StringVar()
        for i, (key, var) in enumerate(
            [("import_name", self.name_var), ("import_language", self.lang_var), ("import_abbr", self.abbr_var)]
        ):
            ttk.Label(form, text=self.i18n.t(key)).grid(row=i, column=0, sticky="w")
            ttk.Entry(form, textvariable=var).grid(row=i, column=1, sticky="ew", pady=1)
        ttk.Label(form, text=self.i18n.t("import_code")).grid(row=3, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.code_var).grid(row=3, column=1, sticky="ew", pady=1)
        form.columnconfigure(1, weight=1)

        self.register_btn = ttk.Button(self, text=self.i18n.t("import_register"), command=self._register_translation, state="disabled")
        self.register_btn.pack(anchor="e", padx=8, pady=6)

    def _pick(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Bible", "*.txt *.json"), ("All", "*.*")]
        )
        if not path:
            return
        self.source_path = path
        try:
            self.report = importer.parse_file(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(self.i18n.t("error"), str(exc))
            return
        self._show_report()

    def _show_report(self) -> None:
        r = self.report
        assert r is not None
        stats = self.i18n.t("review_stats", books=r.n_books, chapters=r.n_chapters, verses=r.n_verses)
        stats += "\n" + self.i18n.t("review_problems", count=len(r.problems) + len(r.duplicates))
        if r.count_mismatch:
            names = ", ".join(self.i18n.book_name(b) for b in r.count_mismatch)
            stats += "\n" + self.i18n.t("review_count_mismatch", books=names)
        self.stats.configure(text=stats)

        self.problems.delete("1.0", "end")
        for prob in (r.problems + r.duplicates)[:200]:
            self.problems.insert("end", f"[{prob.line_no}] {prob.reason}: {prob.raw[:70]}\n")
        if not r.ok:
            self.problems.insert("end", "\n" + self.i18n.t("review_pass_required") + "\n")
        self.register_btn.configure(state="normal" if r.ok else "disabled")

    def _register_translation(self) -> None:
        if not self.report or not self.report.ok:
            return
        code = self.code_var.get().strip() or self.abbr_var.get().strip() or "USER"
        importer.register(
            self.report,
            code=code,
            name=self.name_var.get().strip() or code,
            language=self.lang_var.get().strip() or "und",
            abbr=self.abbr_var.get().strip(),
            original_path=self.source_path,
        )
        self.app.registry.reload()
        self.app._refresh_translation_list()
        messagebox.showinfo(self.i18n.t("done"), self.i18n.t("import_register"))
        self.destroy()


# order the boxes are drawn/edited in, with their outline colours
_LAYOUT_KEYS = (("title", "#c0392b"), ("section", "#2980b9"), ("body", "#27ae60"))
_CANVAS_W = 380  # px; height follows the slide aspect


class LayoutDialog(tk.Toplevel):
    """Drag-and-drop editor for the slide layout (item 6).

    Shows the title / reference / body boxes as draggable rectangles laid over
    a slide-shaped canvas for the current aspect ratio, plus font / size / bold
    controls for the title and reference. "저장" persists the arrangement so all
    subsequent PPTs use it; "초기화" restores the original engine layout.
    """

    def __init__(self, master: App) -> None:
        super().__init__(master)
        self.app = master
        self.i18n = master.i18n
        self.settings = master.settings

        self.title(self.i18n.t("customize_layout"))
        self.geometry("640x520")

        self._style = master._build_style()
        w_in, h_in = ppt.ASPECT_RATIOS.get(self.settings.aspect_ratio, ppt.ASPECT_RATIOS["16:9"])
        self._scale = _CANVAS_W / w_in
        self._cw = _CANVAS_W
        self._ch = int(h_in * self._scale)

        wrap = ttk.Frame(self)
        wrap.pack(fill="both", expand=True, padx=8, pady=8)

        left = ttk.Frame(wrap)
        left.pack(side="left", fill="y")
        ttk.Label(left, text=self.i18n.t("customize_hint"), foreground="#666",
                  wraplength=self._cw, justify="left").pack(anchor="w")
        self.canvas = tk.Canvas(left, width=self._cw, height=self._ch,
                                background="#f4f4f4", highlightthickness=1,
                                highlightbackground="#bbb")
        self.canvas.pack(pady=4)

        self._rects: dict[str, dict] = {}
        self._drag: tuple | None = None
        self._draw_boxes(self._current_fractions())

        right = ttk.Frame(wrap)
        right.pack(side="left", fill="both", expand=True, padx=(12, 0))
        s = self.settings
        # explicit per-element typography vars (no dynamic attribute access)
        self._title_vars = self._build_element_controls(
            right, "customize_title", s.title_font, s.title_font_size, s.title_bold
        )
        self._section_vars = self._build_element_controls(
            right, "customize_section", s.section_font, s.section_font_size, s.section_bold
        )

        btns = ttk.Frame(right)
        btns.pack(anchor="w", pady=(12, 0))
        ttk.Button(btns, text=self.i18n.t("save"), command=self._save).pack(side="left")
        ttk.Button(btns, text=self.i18n.t("reset"), command=self._reset).pack(side="left", padx=6)

    # -- fractions / drawing --------------------------------------------- #
    def _current_fractions(self) -> dict[str, list[float]]:
        base = self._style.default_layout_fractions()
        base.update({k: list(v) for k, v in self.settings.layout_boxes.items()})
        return base

    def _draw_boxes(self, fractions: dict[str, list[float]]) -> None:
        self.canvas.delete("all")
        self._rects = {}
        for key, colour in _LAYOUT_KEYS:
            x, y, w, h = fractions[key]
            x0, y0 = x * self._cw, y * self._ch
            x1, y1 = x0 + w * self._cw, y0 + h * self._ch
            rid = self.canvas.create_rectangle(
                x0, y0, x1, y1, outline=colour, width=2, fill=colour, stipple="gray12",
            )
            tid = self.canvas.create_text(
                (x0 + x1) / 2, (y0 + y1) / 2, text=self.i18n.t(f"customize_{key}"),
                fill=colour, font=("TkDefaultFont", 9, "bold"),
            )
            self._rects[key] = {"rect": rid, "text": tid}
            for item in (rid, tid):
                self.canvas.tag_bind(item, "<Button-1>", lambda e, k=key: self._press(e, k))
                self.canvas.tag_bind(item, "<B1-Motion>", self._motion)

    def _press(self, event, key: str) -> None:
        self._drag = (key, event.x, event.y)

    def _motion(self, event) -> None:
        if not self._drag:
            return
        key, px, py = self._drag
        rid = self._rects[key]["rect"]
        x0, y0, x1, y1 = self.canvas.coords(rid)
        dx = max(-x0, min(event.x - px, self._cw - x1))
        dy = max(-y0, min(event.y - py, self._ch - y1))
        self.canvas.move(rid, dx, dy)
        self.canvas.move(self._rects[key]["text"], dx, dy)
        self._drag = (key, event.x, event.y)

    def _fraction_of(self, key: str) -> list[float]:
        x0, y0, x1, y1 = self.canvas.coords(self._rects[key]["rect"])
        return [x0 / self._cw, y0 / self._ch, (x1 - x0) / self._cw, (y1 - y0) / self._ch]

    # -- element font controls ------------------------------------------- #
    def _build_element_controls(
        self, parent, title_key: str, font: str, size: int, bold: bool
    ) -> dict[str, tk.Variable]:
        """Build a font/size/bold control group; return its three Tk vars."""
        frame = ttk.LabelFrame(parent, text=self.i18n.t(title_key))
        frame.pack(fill="x", pady=(0, 8))
        font_var = tk.StringVar(value=font or fonts.default_font_name())
        size_var = tk.StringVar(value=str(size))
        bold_var = tk.BooleanVar(value=bold)

        ttk.Label(frame, text=self.i18n.t("font")).grid(row=0, column=0, sticky="w", padx=4, pady=2)
        ttk.Combobox(frame, state="readonly", width=16, textvariable=font_var,
                     values=fonts.font_dropdown_values()).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Label(frame, text=self.i18n.t("body_font_size")).grid(row=1, column=0, sticky="w", padx=4, pady=2)
        ttk.Combobox(frame, state="readonly", width=8, textvariable=size_var,
                     values=[str(s) for s in FONT_SIZES]).grid(row=1, column=1, sticky="w", pady=2)
        ttk.Checkbutton(frame, text=self.i18n.t("bold"), variable=bold_var).grid(
            row=2, column=1, sticky="w", pady=2)
        return {"font": font_var, "size": size_var, "bold": bold_var}

    @staticmethod
    def _size_of(var: tk.Variable, fallback: int) -> int:
        try:
            return int(str(var.get()))
        except (TypeError, ValueError):
            return fallback

    # -- save / reset ---------------------------------------------------- #
    def _save(self) -> None:
        s = self.settings
        s.layout_boxes = {k: self._fraction_of(k) for k, _ in _LAYOUT_KEYS}
        s.title_font = str(self._title_vars["font"].get())
        s.title_font_size = self._size_of(self._title_vars["size"], s.title_font_size)
        s.title_bold = bool(self._title_vars["bold"].get())
        s.section_font = str(self._section_vars["font"].get())
        s.section_font_size = self._size_of(self._section_vars["size"], s.section_font_size)
        s.section_bold = bool(self._section_vars["bold"].get())
        s.save()
        if messagebox.askyesno(self.i18n.t("done"), self.i18n.t("customize_saved")):
            self.destroy()

    def _reset(self) -> None:
        d = Settings()  # engine defaults
        s = self.settings
        s.layout_boxes = {}
        s.title_font, s.title_font_size, s.title_bold = d.title_font, d.title_font_size, d.title_bold
        s.section_font, s.section_font_size, s.section_bold = (
            d.section_font, d.section_font_size, d.section_bold,
        )
        self._title_vars["font"].set(d.title_font or fonts.default_font_name())
        self._title_vars["size"].set(str(d.title_font_size))
        self._title_vars["bold"].set(d.title_bold)
        self._section_vars["font"].set(d.section_font or fonts.default_font_name())
        self._section_vars["size"].set(str(d.section_font_size))
        self._section_vars["bold"].set(d.section_bold)
        s.save()
        self._draw_boxes(self._current_fractions())


def run() -> None:
    App().mainloop()
