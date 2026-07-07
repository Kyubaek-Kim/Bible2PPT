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

WINDOW_SIZE = "430x820"
FONT_SIZES = [24, 28, 32, 36, 40, 44, 48, 54, 60]
ASPECTS = list(ppt.ASPECT_RATIOS.keys())


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.settings = Settings.load()
        self.registry = bible.Registry.load()
        self.i18n = I18n(self.settings.ui_language)
        fonts.register_bundled_fonts()
        if not self.settings.font:
            self.settings.font = fonts.default_font_name()

        self.passages: list[generator.PassageInput] = []
        # widgets whose text depends on the UI language: (widget, key)
        self._i18n_widgets: list[tuple[tk.Widget, str]] = []
        self._trans_index_to_code: list[str] = []

        self.title(self.i18n.t("app_title"))
        self.geometry(WINDOW_SIZE)
        self.minsize(400, 600)

        self._build_scroll_container()
        self._build_all_sections()
        self._load_state()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ #
    # Scaffolding
    # ------------------------------------------------------------------ #
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
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"),
        )
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

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
        self._build_options_section()
        self._build_background_section()
        self._build_output_section()
        self._build_generate_section()

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
        self._label(sec, "select_translations").pack(anchor="w", padx=6)
        self.trans_list = tk.Listbox(sec, selectmode="multiple", height=6, exportselection=False)
        self.trans_list.pack(fill="x", padx=6, pady=4)
        self.trans_list.bind("<<ListboxSelect>>", lambda e: self._on_translation_select())

        row = ttk.Frame(sec)
        row.pack(fill="x", padx=6, pady=(0, 4))
        self._label(row, "default_translation").pack(side="left")
        self.default_trans_var = tk.StringVar()
        self.default_trans_combo = ttk.Combobox(
            row, state="readonly", textvariable=self.default_trans_var, width=22
        )
        self.default_trans_combo.pack(side="left", padx=4)
        self.default_trans_combo.bind(
            "<<ComboboxSelected>>", lambda e: self._on_default_translation_change()
        )
        self._button(sec, "import_bible", self._open_import_dialog).pack(
            anchor="w", padx=6, pady=(0, 4)
        )

    def _refresh_translation_list(self) -> None:
        self.trans_list.delete(0, "end")
        self._trans_index_to_code = []
        metas = self.registry.list_meta()
        for meta in metas:
            label = self.i18n.translation_label(meta.name, meta.language)
            self.trans_list.insert("end", label)
            self._trans_index_to_code.append(meta.code)
        # restore selection
        for i, code in enumerate(self._trans_index_to_code):
            if code in self.settings.selected_translations:
                self.trans_list.selection_set(i)
        # default translation dropdown
        labels = [
            self.i18n.translation_label(m.name, m.language) for m in metas
        ]
        self.default_trans_combo.configure(values=labels)
        if self.settings.default_translation in self._trans_index_to_code:
            self.default_trans_combo.current(
                self._trans_index_to_code.index(self.settings.default_translation)
            )
        elif labels:
            self.default_trans_combo.current(0)
            self.settings.default_translation = self._trans_index_to_code[0]

    def _selected_translation_codes(self) -> list[str]:
        codes = [self._trans_index_to_code[i] for i in self.trans_list.curselection()]
        # put the default translation first so interleave starts with it
        default = self.settings.default_translation
        if default in codes:
            codes = [default] + [c for c in codes if c != default]
        return codes

    def _on_translation_select(self) -> None:
        self.settings.selected_translations = [
            self._trans_index_to_code[i] for i in self.trans_list.curselection()
        ]

    def _on_default_translation_change(self) -> None:
        idx = self.default_trans_combo.current()
        if 0 <= idx < len(self._trans_index_to_code):
            self.settings.default_translation = self._trans_index_to_code[idx]
            self._refresh_book_dropdown()

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

    def _on_book_change(self) -> None:
        trans = self._default_translation()
        idx = self.book_combo.current()
        if idx < 0 or trans is None:
            return
        book_id = self._book_ids[idx]
        chapters = trans.chapters(book_id) or list(
            bible.canonical_versification().get(book_id, {}).keys()
        )
        self.chapter_combo.configure(values=[str(c) for c in chapters])
        self.chapter_var.set("")
        self.vstart_combo.configure(values=[])
        self.vend_combo.configure(values=[])

    def _current_verses(self) -> list[int]:
        trans = self._default_translation()
        idx = self.book_combo.current()
        if idx < 0 or trans is None or not self.chapter_var.get():
            return []
        book_id = self._book_ids[idx]
        chapter = int(self.chapter_var.get())
        verses = trans.verses(book_id, chapter)
        if not verses:
            last = bible.canonical_versification().get(book_id, {}).get(chapter, 0)
            verses = list(range(1, last + 1))
        return verses

    def _on_chapter_change(self) -> None:
        verses = [str(v) for v in self._current_verses()]
        self.vstart_combo.configure(values=verses)
        self.vend_combo.configure(values=verses)
        self.vstart_var.set("")
        self.vend_var.set("")

    def _on_vstart_change(self) -> None:
        # end-verse list starts at (and includes) the start verse — fixes the
        # legacy off-by-one that excluded the start verse.
        verses = self._current_verses()
        if not self.vstart_var.get():
            return
        start = int(self.vstart_var.get())
        self.vend_combo.configure(values=[str(v) for v in verses if v >= start])

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
        idx = self.book_combo.current()
        if idx < 0 or not self.chapter_var.get() or not self.vstart_var.get():
            return None
        book_id = self._book_ids[idx]
        ref = f"{book_id} {self.chapter_var.get()}:{self.vstart_var.get()}"
        if self.vend_var.get():
            ref += f"-{self.vend_var.get()}"
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
        sec = self._section("aspect_ratio")
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

    def _on_option_change(self) -> None:
        if self.aspect_var.get():
            self.settings.aspect_ratio = self.aspect_var.get()
        if self.size_var.get():
            self.settings.body_font_size = int(self.size_var.get())

    def _on_font_change(self) -> None:
        self.settings.font = self.font_var.get()
        self._update_font_preview()

    def _update_font_preview(self) -> None:
        name = self.font_var.get() or fonts.default_font_name()
        available, hint = fonts.ensure_font_available(name, self)
        self.font_hint.configure(text="" if available else hint)
        try:
            title_font = tkfont.Font(family=name, size=18, weight="bold")
            body_font = tkfont.Font(family=name, size=14)
        except tk.TclError:
            title_font = tkfont.Font(size=18, weight="bold")
            body_font = tkfont.Font(size=14)
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
    def _build_generate_section(self) -> None:
        sec = self._section("generate_mode")
        self.mode_var = tk.StringVar(value=self.settings.generate_mode)
        self.mode_sep = ttk.Radiobutton(sec, text=self.i18n.t("mode_separate"), value="separate", variable=self.mode_var, command=self._on_mode_change)
        self.mode_sep.pack(anchor="w", padx=6)
        self.mode_comb = ttk.Radiobutton(sec, text=self.i18n.t("mode_combined"), value="combined", variable=self.mode_var, command=self._on_mode_change)
        self.mode_comb.pack(anchor="w", padx=6)
        self._i18n_widgets.append((self.mode_sep, "mode_separate"))
        self._i18n_widgets.append((self.mode_comb, "mode_combined"))
        self._button(sec, "generate", self._generate).pack(fill="x", padx=6, pady=6)

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
        style = ppt.SlideStyle(
            aspect=self.settings.aspect_ratio,
            font_name=self.settings.font or fonts.default_font_name(),
            body_font_size=self.settings.body_font_size,
        )
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
        ttk.Label(form, text="code").grid(row=3, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.code_var).grid(row=3, column=1, sticky="ew", pady=1)
        form.columnconfigure(1, weight=1)

        self.register_btn = ttk.Button(self, text=self.i18n.t("import_register"), command=self._register, state="disabled")
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

    def _register(self) -> None:
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


def run() -> None:
    App().mainloop()
