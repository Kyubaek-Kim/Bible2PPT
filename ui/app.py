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
from collections.abc import Callable
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

from PIL import Image, ImageTk

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
        # text-colour swatches keyed by 'title'/'section'/'body'
        self._color_swatches: dict[str, tk.Button] = {}
        # background dropdown option keys, parallel to the combobox values
        self._bg_option_keys: list[str] = []
        # the open layout-customizer window, if any (kept so its faded
        # background can follow live 배경 selection changes)
        self._layout_dialog: LayoutDialog | None = None

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

        # text colours (title / section / body), shown above the preview
        self._label(sec, "font_color").pack(anchor="w", padx=6, pady=(6, 0))
        colors = ttk.Frame(sec)
        colors.pack(fill="x", padx=6, pady=(0, 2))
        self._color_button(colors, "customize_title", "title")
        self._color_button(colors, "customize_section", "section")
        self._color_button(colors, "customize_body", "body")

        # font preview
        self._label(sec, "font_preview").pack(anchor="w", padx=6, pady=(4, 0))
        self.preview_frame = ttk.Frame(sec, relief="solid", borderwidth=1)
        self.preview_frame.pack(fill="x", padx=6, pady=2)
        self.preview_title = tk.Label(self.preview_frame, anchor="w", justify="left")
        self.preview_title.pack(fill="x", padx=6, pady=(4, 0))
        self.preview_section = tk.Label(self.preview_frame, anchor="w", justify="left")
        self.preview_section.pack(fill="x", padx=6)
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
        self._update_font_preview()

    def _on_body_bold_change(self) -> None:
        self.settings.body_bold = self.body_bold_var.get()
        self._update_font_preview()

    # -- text colour controls -------------------------------------------- #
    def _get_color(self, kind: str) -> str:
        """Stored colour for ``kind`` in {'title','section','body'} ('' = default)."""
        if kind == "title":
            return self.settings.title_color
        if kind == "section":
            return self.settings.section_color
        return self.settings.body_color

    def _set_color(self, kind: str, value: str) -> None:
        if kind == "title":
            self.settings.title_color = value
        elif kind == "section":
            self.settings.section_color = value
        else:
            self.settings.body_color = value

    def _color_button(self, parent, label_key: str, kind: str) -> None:
        """A labelled swatch that opens a colour picker for the ``kind`` colour.

        The swatch fill reflects the current colour; picking updates the setting
        and the live preview. Buttons for title / section / body sit in a row."""
        cell = ttk.Frame(parent)
        cell.pack(side="left", padx=(0, 10))
        lbl = ttk.Label(cell, text=self.i18n.t(label_key))
        lbl.pack(side="left", padx=(0, 3))
        swatch = tk.Button(cell, width=3, relief="groove", takefocus=False)
        swatch.pack(side="left")
        swatch.configure(command=lambda: self._pick_color(kind))
        self._color_swatches[kind] = swatch
        self._i18n_widgets.append((lbl, label_key))

    def _pick_color(self, kind: str) -> None:
        _rgb, hexval = colorchooser.askcolor(
            color=self._preview_color(kind), parent=self
        )
        if hexval:
            self._set_color(kind, hexval)
            self._refresh_color_swatch(kind)
            self._update_font_preview()

    def _refresh_color_swatch(self, kind: str) -> None:
        swatch = self._color_swatches.get(kind)
        if swatch is not None:
            swatch.configure(background=self._preview_color(kind))

    def _preview_color(self, kind: str) -> str:
        return self._get_color(kind) or "#000000"

    def _build_style(self) -> ppt.SlideStyle:
        """Assemble a :class:`ppt.SlideStyle` from the current settings,
        including any saved layout customisation (item 6)."""
        s = self.settings
        return ppt.SlideStyle(
            aspect=s.aspect_ratio,
            font_name=s.font or fonts.default_font_name(),
            body_font_size=s.body_font_size,
            body_bold_opt=s.body_bold,
            title_font_size=s.title_font_size,
            title_bold=s.title_bold,
            title_enabled=s.title_enabled,
            section_font_size=s.section_font_size,
            section_bold=s.section_bold,
            section_enabled=s.section_enabled,
            title_color=s.title_color,
            section_color=s.section_color,
            body_color=s.body_color,
            layout_boxes=dict(s.layout_boxes),
        )

    def _open_layout_dialog(self) -> None:
        if self._layout_dialog is not None and self._layout_dialog.winfo_exists():
            self._layout_dialog.lift()
            return
        LayoutDialog(self)

    def _notify_background_changed(self) -> None:
        """Refresh the customizer's faded background if it is open."""
        dlg = self._layout_dialog
        if dlg is not None and dlg.winfo_exists():
            dlg.refresh_background()

    def _on_font_change(self) -> None:
        self.settings.font = self.font_var.get()
        self._update_font_preview()

    # the preview is a scaled-down mock of a slide; real point sizes are
    # multiplied by this so their *relative* sizes (title vs. body) show through.
    PREVIEW_SCALE = 0.45

    def _preview_font(self, family: str | None, size: int, *, bold: bool) -> tkfont.Font:
        px = max(9, round(size * self.PREVIEW_SCALE))
        weight = "bold" if bold else "normal"
        try:
            if family is None:
                raise tk.TclError
            return tkfont.Font(family=family, size=px, weight=weight)
        except tk.TclError:
            return tkfont.Font(size=px, weight=weight)

    def _update_font_preview(self) -> None:
        name = self.font_var.get() or fonts.default_font_name()
        available, hint = fonts.ensure_font_available(name, self)
        self.font_hint.configure(text="" if available else hint)
        choice = fonts.resolve(name)
        families = fonts.system_font_families(self)
        family = next((f for f in (choice.typeface, choice.label) if f in families), None)
        s = self.settings
        body_bold = fonts.run_bold(choice, s.body_bold)
        title_bold = fonts.run_bold(choice, s.title_bold)
        section_bold = fonts.run_bold(choice, s.section_bold)

        # title (uses the title size relative to the body size)
        if s.title_enabled:
            self.preview_title.configure(
                text=self.i18n.t("font_preview_sample_title"),
                font=self._preview_font(family, s.title_font_size, bold=title_bold),
                fg=self._preview_color("title"),
            )
            self.preview_title.pack(fill="x", padx=6, pady=(4, 0), before=self.preview_body)
        else:
            self.preview_title.pack_forget()

        if s.section_enabled:
            self.preview_section.configure(
                text=self.i18n.t("font_preview_sample_section"),
                font=self._preview_font(family, s.section_font_size, bold=section_bold),
                fg=self._preview_color("section"),
            )
            self.preview_section.pack(fill="x", padx=6, before=self.preview_body)
        else:
            self.preview_section.pack_forget()

        self.preview_body.configure(
            text="1. 태초에 하나님이 천지를 창조하시니라 / In the beginning God created the heaven and the earth.",
            font=self._preview_font(family, s.body_font_size, bold=body_bold),
            fg=self._preview_color("body"),
        )

    # ------------------------------------------------------------------ #
    # Background
    # ------------------------------------------------------------------ #
    def _build_background_section(self) -> None:
        sec = self._section("background")
        # attach a new image  +  select from the registered backgrounds  +  manage
        self._button(sec, "background_custom", self._attach_background).pack(
            anchor="w", padx=6, pady=(4, 2)
        )
        self._label(sec, "background_select").pack(anchor="w", padx=6, pady=(2, 0))
        row = ttk.Frame(sec)
        row.pack(fill="x", padx=6, pady=(0, 2))
        self.bg_select_var = tk.StringVar()
        self.bg_select_combo = ttk.Combobox(
            row, state="readonly", textvariable=self.bg_select_var
        )
        self.bg_select_combo.pack(side="left", fill="x", expand=True)
        self.bg_select_combo.bind("<<ComboboxSelected>>", lambda e: self._on_select_background())
        self._button(row, "background_manage", self._open_background_manager).pack(
            side="left", padx=(4, 0)
        )

    def _slide_cm(self) -> tuple[float, float]:
        w_in, h_in = ppt.ASPECT_RATIOS[self.settings.aspect_ratio]
        return w_in * 2.54, h_in * 2.54

    def _bg_display(self, key: str, name_key_or_name: str) -> str:
        """Default option is localized; custom entries show their file name."""
        return self.i18n.t("background_default") if key == "" else name_key_or_name

    def _refresh_background_combo(self) -> None:
        options = self.settings.background_options()
        self._bg_option_keys = [k for k, _ in options]
        self.bg_select_combo.configure(
            values=[self._bg_display(k, n) for k, n in options]
        )
        sel = self.settings.selected_background
        idx = self._bg_option_keys.index(sel) if sel in self._bg_option_keys else 0
        self.bg_select_combo.current(idx)

    def _confirm_crop(self, source: str) -> bool:
        """Warn (확인/취소) how much of ``source`` is cropped for the current
        aspect. True when the user confirms (or no crop needed)."""
        w_cm, h_cm = self._slide_cm()
        plan = image_util.plan_crop(source, w_cm, h_cm)
        if not plan.needs_crop:
            return True
        axis = "상/하" if plan.axis == "vertical" else "좌/우"
        msg = self.i18n.t("crop_confirm_body", axis=axis, px=plan.crop_px, cm=plan.crop_cm)
        return messagebox.askokcancel(self.i18n.t("crop_confirm_title"), msg)

    def _attach_background(self) -> None:
        """Attach (import) a new image: copy into AppData, register it once, and
        select it after a crop confirmation."""
        path = filedialog.askopenfilename(
            filetypes=[("Image", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All", "*.*")]
        )
        if not path:
            return
        if not self._confirm_crop(path):
            return
        stored = image_util.add_to_history(path)
        self.settings.add_background(str(stored))
        self.settings.selected_background = str(stored)
        self._refresh_background_combo()
        self._notify_background_changed()

    def _on_select_background(self) -> None:
        """Select an already-registered background (never re-imports/duplicates)."""
        idx = self.bg_select_combo.current()
        if not (0 <= idx < len(self._bg_option_keys)):
            return
        key = self._bg_option_keys[idx]
        if key == self.settings.selected_background:
            return
        if key and not self._confirm_crop(key):
            self._refresh_background_combo()  # revert combo to the current selection
            return
        self.settings.selected_background = key
        self._refresh_background_combo()
        self._notify_background_changed()

    def _open_background_manager(self) -> None:
        BackgroundManager(self)

    def _render_background(self) -> Path | None:
        """The background to embed: the selected original cropped to the current
        aspect (cached), or None when the file is missing."""
        original = self.settings.resolved_background()
        if not original.exists():
            return None
        w_cm, h_cm = self._slide_cm()
        plan = image_util.plan_crop(str(original), w_cm, h_cm)
        aspect_tag = self.settings.aspect_ratio.replace(":", "x")
        cached = paths.background_cache_dir() / f"{aspect_tag}_{original.name}"
        image_util.apply_crop(str(original), plan, cached)
        return cached

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
                background=self._render_background(),
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
        self._refresh_background_combo()
        for kind in ("title", "section", "body"):
            self._refresh_color_swatch(kind)
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


_MIN_BOX_PX = 26          # a box may not be shrunk below this in either axis
_HANDLE = 7               # half-size of a corner resize handle (px)
# corner -> (which x edge, which y edge) it moves, and the hover cursor
_CORNERS = {
    "nw": (0, 1, "top_left_corner"),
    "ne": (2, 1, "top_right_corner"),
    "sw": (0, 3, "bottom_left_corner"),
    "se": (2, 3, "bottom_right_corner"),
}


class LayoutDialog(tk.Toplevel):
    """Drag-and-drop + resize editor for the slide layout.

    Shows the title / reference / body boxes as rectangles laid over a
    slide-shaped canvas (with the current background faded behind them) for the
    active aspect ratio, plus size / bold / 활성화 controls for the title and
    reference. Boxes can be *moved* (drag the body) and *resized* (drag a corner
    handle — the cursor turns into a diagonal double-arrow). "저장" persists the
    arrangement so all subsequent PPTs use those exact regions; "초기화" restores
    the original engine layout. Live warnings flag overlaps / too-small regions.
    """

    def __init__(self, master: App) -> None:
        super().__init__(master)
        self.app = master
        self.i18n = master.i18n
        self.settings = master.settings

        self.title(self.i18n.t("customize_layout"))
        self.geometry("660x560")

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
        # live layout warnings (overlap / too small); empty when the layout is ok
        self.warn_label = ttk.Label(left, text="", foreground="#c0392b",
                                    wraplength=self._cw, justify="left")
        self.warn_label.pack(anchor="w")

        # box geometry model in canvas px: {key: [x0, y0, x1, y1]}
        self._boxes: dict[str, list[float]] = {}
        self._bg_photo: ImageTk.PhotoImage | None = None
        self._drag: tuple | None = None  # (key, mode, corner|None, last_x, last_y)

        right = ttk.Frame(wrap)
        right.pack(side="left", fill="both", expand=True, padx=(12, 0))
        s = self.settings
        # explicit per-element typography vars (no dynamic attribute access)
        self._title_vars = self._build_element_controls(
            right, "customize_title", s.title_font_size, s.title_bold, s.title_enabled,
            on_change=self._refresh_preview,
        )
        self._section_vars = self._build_element_controls(
            right, "customize_section", s.section_font_size, s.section_bold, s.section_enabled,
            on_change=self._refresh_preview,
        )

        # seed the model from the current fractions, then paint
        fr = self._current_fractions()
        self._boxes = {k: self._fr_to_px(fr[k]) for k, _ in _LAYOUT_KEYS}
        self._draw_background()
        self._redraw_overlay()

        btns = ttk.Frame(right)
        btns.pack(anchor="w", pady=(12, 0))
        ttk.Button(btns, text=self.i18n.t("save"), command=self._save).pack(side="left")
        ttk.Button(btns, text=self.i18n.t("reset"), command=self._reset).pack(side="left", padx=6)

        # let the App refresh our background live when the user changes it
        self.app._layout_dialog = self
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self) -> None:
        self.app._layout_dialog = None
        self.destroy()

    # -- fractions / geometry -------------------------------------------- #
    def _current_fractions(self) -> dict[str, list[float]]:
        base = self._style.default_layout_fractions()
        base.update({k: list(v) for k, v in self.settings.layout_boxes.items()})
        return base

    def _fr_to_px(self, fr: list[float]) -> list[float]:
        x, y, w, h = fr
        return [x * self._cw, y * self._ch, (x + w) * self._cw, (y + h) * self._ch]

    def _fraction_of(self, key: str) -> list[float]:
        x0, y0, x1, y1 = self._boxes[key]
        return [x0 / self._cw, y0 / self._ch, (x1 - x0) / self._cw, (y1 - y0) / self._ch]

    def _element_enabled(self, key: str) -> bool:
        if key == "title":
            return bool(self._title_vars["enabled"].get())
        if key == "section":
            return bool(self._section_vars["enabled"].get())
        return True  # body is always present

    # -- background ------------------------------------------------------ #
    def _draw_background(self) -> None:
        """Paint the currently-selected background, faded to 20% opacity, so the
        boxes stay readable. Follows the main window's background live."""
        self.canvas.delete("bg")
        self._bg_photo = None
        src = self.app._render_background()
        if src is None or not Path(src).exists():
            return
        try:
            img = Image.open(src).convert("RGB").resize((self._cw, self._ch))
        except OSError:
            return
        # blend toward white: 20% of the image + 80% white == 80% transparency
        faded = Image.blend(Image.new("RGB", img.size, "white"), img, 0.2)
        self._bg_photo = ImageTk.PhotoImage(faded)
        self.canvas.create_image(0, 0, anchor="nw", image=self._bg_photo, tags="bg")

    def refresh_background(self) -> None:
        """Called by the App when the selected background changes."""
        self._draw_background()
        self._redraw_overlay()

    # -- overlay (boxes + handles) --------------------------------------- #
    def _redraw_overlay(self) -> None:
        self.canvas.delete("ov")
        for key, colour in _LAYOUT_KEYS:
            enabled = self._element_enabled(key)
            draw_colour = colour if enabled else "#b0b0b0"
            x0, y0, x1, y1 = self._boxes[key]
            rid = self.canvas.create_rectangle(
                x0, y0, x1, y1, outline=draw_colour, width=2,
                fill=draw_colour, stipple="gray12", tags=("ov", f"box:{key}"),
            )
            label = self.i18n.t(f"customize_{key}")
            if not enabled:
                label = f"{label} ({self.i18n.t('customize_disabled')})"
            tid = self.canvas.create_text(
                (x0 + x1) / 2, (y0 + y1) / 2, text=label, fill=draw_colour,
                font=self._canvas_font(*self._element_style(key)), tags=("ov", f"box:{key}"),
            )
            for item in (rid, tid):
                self.canvas.tag_bind(item, "<Button-1>", lambda e, k=key: self._press_move(e, k))
                self.canvas.tag_bind(item, "<B1-Motion>", self._motion)
                self.canvas.tag_bind(item, "<ButtonRelease-1>", self._release)
                self.canvas.tag_bind(item, "<Enter>", lambda e: self.canvas.configure(cursor="fleur"))
                self.canvas.tag_bind(item, "<Leave>", lambda e: self.canvas.configure(cursor=""))
            # corner resize handles
            for corner, (ix, iy, cursor) in _CORNERS.items():
                hx, hy = self._boxes[key][ix], self._boxes[key][iy]
                hid = self.canvas.create_rectangle(
                    hx - _HANDLE, hy - _HANDLE, hx + _HANDLE, hy + _HANDLE,
                    outline=draw_colour, fill="white", width=1, tags=("ov", f"box:{key}"),
                )
                self.canvas.tag_bind(hid, "<Button-1>",
                                     lambda e, k=key, c=corner: self._press_resize(e, k, c))
                self.canvas.tag_bind(hid, "<B1-Motion>", self._motion)
                self.canvas.tag_bind(hid, "<ButtonRelease-1>", self._release)
                self.canvas.tag_bind(hid, "<Enter>",
                                     lambda e, cur=cursor: self.canvas.configure(cursor=cur))
                self.canvas.tag_bind(hid, "<Leave>", lambda e: self.canvas.configure(cursor=""))
        self._validate()

    def _element_style(self, key: str) -> tuple[str, int, bool]:
        """Return (font_name, size_pt, bold) for a preview box.

        Every element uses the 화면 설정 글자체; only size/bold differ per
        element. Title/reference sizes+bold come from this dialog's live
        controls, the body from the main window's 본문 글자크기/굵게."""
        font_name = self.settings.font or fonts.default_font_name()
        if key == "title":
            return (
                font_name,
                self._size_of(self._title_vars["size"], self._style.title_font_size),
                bool(self._title_vars["bold"].get()),
            )
        if key == "section":
            return (
                font_name,
                self._size_of(self._section_vars["size"], self._style.section_font_size),
                bool(self._section_vars["bold"].get()),
            )
        return (font_name, self.settings.body_font_size, self.settings.body_bold)

    def _canvas_font(self, name: str, size_pt: int, bold: bool) -> tkfont.Font:
        """Build a preview font scaled from PPT points to canvas pixels."""
        choice = fonts.resolve(name or fonts.default_font_name())
        families = fonts.system_font_families(self.app)
        family = next((f for f in (choice.typeface, choice.label) if f in families), None)
        weight = "bold" if (bold or choice.bold) else "normal"
        px = max(7, int(size_pt / 72 * self._scale))
        try:
            if family is None:
                raise tk.TclError
            return tkfont.Font(family=family, size=-px, weight=weight)
        except tk.TclError:
            return tkfont.Font(size=-px, weight=weight)

    def _refresh_preview(self) -> None:
        """Re-apply preview fonts and enabled-state live when a control changes."""
        self._redraw_overlay()

    # -- move / resize interaction --------------------------------------- #
    def _press_move(self, event, key: str) -> None:
        self._drag = (key, "move", None, event.x, event.y)

    def _press_resize(self, event, key: str, corner: str) -> None:
        self._drag = (key, "resize", corner, event.x, event.y)

    def _motion(self, event) -> None:
        if not self._drag:
            return
        key, mode, corner, px, py = self._drag
        dx, dy = event.x - px, event.y - py
        box = self._boxes[key]
        if mode == "move":
            dx = max(-box[0], min(dx, self._cw - box[2]))
            dy = max(-box[1], min(dy, self._ch - box[3]))
            self._boxes[key] = [box[0] + dx, box[1] + dy, box[2] + dx, box[3] + dy]
        else:
            ix, iy, _ = _CORNERS[corner]
            nx = min(max(box[ix] + dx, 0), self._cw)
            ny = min(max(box[iy] + dy, 0), self._ch)
            new = list(box)
            new[ix], new[iy] = nx, ny
            # keep min size by clamping the moved edge, not flipping the box
            if abs(new[2] - new[0]) >= _MIN_BOX_PX:
                box[ix] = nx
            if abs(new[3] - new[1]) >= _MIN_BOX_PX:
                box[iy] = ny
        self._drag = (key, mode, corner, event.x, event.y)
        self._redraw_overlay()

    def _release(self, event) -> None:
        self._drag = None

    # -- validation ------------------------------------------------------ #
    def _validate(self) -> list[str]:
        """Collect and display warnings; return them. Advisory only — the engine
        still guarantees no slide overflow regardless of these regions."""
        warns: list[str] = []

        # the critical rule: the body must not intrude into an enabled header.
        # (the title/reference bands may share a little vertical space by design,
        # so only body-vs-header overlaps are flagged.)
        for header in ("title", "section"):
            if self._element_enabled(header) and self._overlap(
                self._boxes["body"], self._boxes[header]
            ):
                warns.append(self.i18n.t(
                    "customize_warn_overlap",
                    a=self.i18n.t("customize_body"), b=self.i18n.t(f"customize_{header}"),
                ))

        # body too small to hold even two lines at the current body size
        trial = ppt.SlideStyle(
            aspect=self.settings.aspect_ratio,
            font_name=self.settings.font or fonts.default_font_name(),
            body_font_size=self.settings.body_font_size,
            layout_boxes={"body": self._fraction_of("body")},
        )
        if trial.max_body_lines < 2 or trial.max_units_per_line < 6:
            warns.append(self.i18n.t("customize_warn_body_small"))

        # a header box shorter than its own font can clip the text
        for key, vars_ in (("title", self._title_vars), ("section", self._section_vars)):
            if not self._element_enabled(key):
                continue
            _, _, _, bh = self._fraction_of(key)
            box_h_pt = bh * self._ch / self._scale * ppt.IN_TO_PT
            font_pt = self._size_of(vars_["size"], 24)
            if box_h_pt < font_pt * 1.1:
                warns.append(self.i18n.t(
                    "customize_warn_box_small", el=self.i18n.t(f"customize_{key}")))

        self.warn_label.configure(text="\n".join(warns))
        return warns

    @staticmethod
    def _overlap(a: list[float], b: list[float]) -> bool:
        pad = 1.0  # ignore hairline touching
        return not (a[2] - pad <= b[0] or b[2] - pad <= a[0]
                    or a[3] - pad <= b[1] or b[3] - pad <= a[1])

    # -- element font controls ------------------------------------------- #
    def _build_element_controls(
        self, parent, title_key: str, size: int, bold: bool, enabled: bool,
        on_change: Callable[[], None],
    ) -> dict[str, tk.Variable]:
        """Build an 활성화 / size / bold control group; return its Tk vars.

        The font face is intentionally *not* selectable here: it is governed by
        the 화면 설정 글자체 for every element. ``on_change`` fires whenever a
        control changes so the preview can update live."""
        frame = ttk.LabelFrame(parent, text=self.i18n.t(title_key))
        frame.pack(fill="x", pady=(0, 8))
        enabled_var = tk.BooleanVar(value=enabled)
        size_var = tk.StringVar(value=str(size))
        bold_var = tk.BooleanVar(value=bold)

        ttk.Checkbutton(frame, text=self.i18n.t("customize_enabled"), variable=enabled_var,
                        command=on_change).grid(row=0, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        ttk.Label(frame, text=self.i18n.t("body_font_size")).grid(row=1, column=0, sticky="w", padx=4, pady=2)
        size_combo = ttk.Combobox(frame, state="readonly", width=8, textvariable=size_var,
                                  values=[str(s) for s in FONT_SIZES])
        size_combo.grid(row=1, column=1, sticky="w", pady=2)
        ttk.Checkbutton(frame, text=self.i18n.t("bold"), variable=bold_var,
                        command=on_change).grid(row=2, column=1, sticky="w", pady=2)
        size_combo.bind("<<ComboboxSelected>>", lambda e: on_change())
        return {"enabled": enabled_var, "size": size_var, "bold": bold_var}

    @staticmethod
    def _size_of(var: tk.Variable, fallback: int) -> int:
        try:
            return int(str(var.get()))
        except (TypeError, ValueError):
            return fallback

    # -- save / reset ---------------------------------------------------- #
    def _save(self) -> None:
        # if the layout is risky, warn once and let the user decide
        warns = self._validate()
        if warns and not messagebox.askyesno(
            self.i18n.t("customize_warn_title"),
            "\n".join(warns) + "\n\n" + self.i18n.t("customize_warn_confirm"),
        ):
            return
        s = self.settings
        s.layout_boxes = {k: self._fraction_of(k) for k, _ in _LAYOUT_KEYS}
        # font face is governed by 화면 설정; only size / bold / visibility differ
        s.title_font_size = self._size_of(self._title_vars["size"], s.title_font_size)
        s.title_bold = bool(self._title_vars["bold"].get())
        s.title_enabled = bool(self._title_vars["enabled"].get())
        s.section_font_size = self._size_of(self._section_vars["size"], s.section_font_size)
        s.section_bold = bool(self._section_vars["bold"].get())
        s.section_enabled = bool(self._section_vars["enabled"].get())
        s.save()
        self.app._update_font_preview()
        if messagebox.askyesno(self.i18n.t("done"), self.i18n.t("customize_saved")):
            self._on_close()

    def _reset(self) -> None:
        d = Settings()  # engine defaults
        s = self.settings
        s.layout_boxes = {}
        s.title_font_size, s.title_bold, s.title_enabled = d.title_font_size, d.title_bold, d.title_enabled
        s.section_font_size, s.section_bold, s.section_enabled = (
            d.section_font_size, d.section_bold, d.section_enabled
        )
        self._title_vars["size"].set(str(d.title_font_size))
        self._title_vars["bold"].set(d.title_bold)
        self._title_vars["enabled"].set(d.title_enabled)
        self._section_vars["size"].set(str(d.section_font_size))
        self._section_vars["bold"].set(d.section_bold)
        self._section_vars["enabled"].set(d.section_enabled)
        s.save()
        fr = self._current_fractions()
        self._boxes = {k: self._fr_to_px(fr[k]) for k, _ in _LAYOUT_KEYS}
        self._redraw_overlay()
        self.app._update_font_preview()


class BackgroundManager(tk.Toplevel):
    """Compact popup to batch-delete registered backgrounds.

    The main window keeps only the space-efficient 배경 선택 dropdown; deletion
    lives here as a checkbox list so several images can be removed at once. The
    built-in 기본 배경 is shown locked (no checkbox) and can never be deleted.
    Deleting removes both the settings entry and the file under AppData.
    """

    def __init__(self, master: App) -> None:
        super().__init__(master)
        self.app = master
        self.i18n = master.i18n
        self.settings = master.settings

        self.title(self.i18n.t("background_manage_title"))
        self.geometry("420x360")
        self.transient(master)

        ttk.Label(self, text=self.i18n.t("background_manage_hint"),
                  foreground="#666", wraplength=390, justify="left").pack(
            anchor="w", padx=10, pady=(10, 4))

        self._body = ttk.Frame(self)
        self._body.pack(fill="both", expand=True, padx=10)
        self._vars: list[tuple[str, tk.BooleanVar]] = []
        self._build_rows()

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=10)
        ttk.Button(btns, text=self.i18n.t("background_delete_selected"),
                   command=self._delete_selected).pack(side="left")
        ttk.Button(btns, text=self.i18n.t("close"), command=self.destroy).pack(side="right")

    def _build_rows(self) -> None:
        for child in self._body.winfo_children():
            child.destroy()
        self._vars = []
        # locked default row
        ttk.Label(self._body, text=f"🔒 {self.i18n.t('background_default')}",
                  foreground="#888").pack(anchor="w", pady=2)
        for path in self.settings.background_history:
            var = tk.BooleanVar(value=False)
            ttk.Checkbutton(self._body, text=Path(path).name, variable=var).pack(anchor="w", pady=2)
            self._vars.append((path, var))
        if not self._vars:
            ttk.Label(self._body, text=self.i18n.t("background_manage_empty"),
                      foreground="#aaa").pack(anchor="w", pady=6)

    def _delete_selected(self) -> None:
        targets = [p for p, v in self._vars if v.get()]
        if not targets:
            return
        if not messagebox.askyesno(
            self.i18n.t("background_manage_title"),
            self.i18n.t("background_delete_confirm", count=len(targets)),
            parent=self,
        ):
            return
        for path in targets:
            image_util.delete_background(path)
            self.settings.remove_background(path)
        self.settings.save()
        self.app._refresh_background_combo()
        self.app._notify_background_changed()
        self._build_rows()


def run() -> None:
    App().mainloop()
