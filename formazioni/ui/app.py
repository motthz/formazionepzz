"""Finestra principale di Formazioni PZZ."""

from __future__ import annotations

import os
import queue
import tempfile
import threading
import time
import traceback
from dataclasses import replace as dc_replace
from datetime import date, datetime
from pathlib import Path
from typing import Any

from ..config import (
    DEFAULT_LANG,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_TEMPLATE_DIR,
    DEFAULT_THEME,
    ICON_PNG,
    LOGO_HEADER_PNG,
    PREVIEW_PREFIX,
    SYSTEM_THEME_POLL_MS,
    available_languages,
    load_language,
    load_settings,
    resolve_theme,
    save_settings,
)
from ..icons import BUTTON_ICONS, strip_leading_symbol
from ..logs import setup_error_log
from ..office import warm_up_office
from ..pdf import build_pdf
from ..system import (
    _build_in_memory_icon,
    _enable_dpi_awareness,
    _png_from_rgba,
    _set_app_icon,
    _set_titlebar_dark,
    cleanup_old_previews,
    notify_windows,
    open_folder,
)
from ..templates import (
    TemplateFile,
    apply_order,
    classify_template_hashes,
    compute_template_hash,
    department_options,
    discover_templates,
    load_module_settings,
    load_saved_hashes,
    safe_file_part,
    save_hashes,
    templates_for_department,
    templates_for_departments,
)
from ..tkcompat import BOTH, END, LEFT, RIGHT, BooleanVar, StringVar, X, filedialog, messagebox, tk, ttk
from ..ui.widgets import DatePickerFrame, Tooltip
from .batch import BatchMixin
from .dialogs import DialogsMixin
from .feedback import FeedbackMixin
from .history import HistoryMixin
from .ordering import OrderingMixin
from .settings import SettingsMixin

try:
    from .kit import RoundedPanel, UiKit
except ImportError:  # senza Pillow resta lo stile base
    RoundedPanel = UiKit = None  # type: ignore[assignment]


class FormazioniApp(OrderingMixin, BatchMixin, HistoryMixin, SettingsMixin, DialogsMixin, FeedbackMixin):
    def __init__(self, root: tk.Tk) -> None:
        if tk is None or ttk is None:
            raise RuntimeError(
                "Questa installazione di Python non include Tkinter. "
                "Installa Python con il supporto Tk e riavvia Formazioni PZZ."
            )
        self.root = root

        self.settings = load_settings()
        self.language_code = self.settings.get("language", DEFAULT_LANG)
        self.language = load_language(self.language_code)
        self.saved_hashes = load_saved_hashes()
        self.hash_status: dict[str, str] = {}

        self.template_dir = StringVar(value=self.settings.get("last_template_dir", str(DEFAULT_TEMPLATE_DIR)))
        self.output_dir = StringVar(value=self.settings.get("last_output_dir", str(DEFAULT_OUTPUT_DIR)))
        self.employee_name = StringVar()
        self.department = StringVar()
        self.auto_open = BooleanVar(value=True)
        self.multi_dept_mode = BooleanVar(value=False)
        # theme_pref e' la scelta dell'utente (light/dark/system); self.theme il tema
        # effettivamente applicato, sempre "light" o "dark".
        pref = self.settings.get("theme", DEFAULT_THEME)
        self.theme_pref = pref if pref in ("light", "dark", "system") else DEFAULT_THEME
        self.theme = StringVar(value=resolve_theme(self.theme_pref))
        self.language_var = StringVar(value=self.language_code)
        self.status = StringVar(value=self.tr("status_initial"))
        self.count_label = StringVar(value=self.tr("count_none"))
        self.progress_label = StringVar(value=self.tr("progress_ready"))
        self.hash_stat_label = StringVar(value="")
        self.templates: list[TemplateFile] = []
        self.ignored: list[Path] = []
        self.template_inclusion: dict[Path, bool] = {}
        # Nomi in app e ordini salvati (file nella cartella template); doc_order e'
        # l'ordine attuale della lista, active_order il nome dell'ordine applicato.
        self.module_settings: dict[str, Any] = {"labels": {}, "orders": {}}
        self.doc_order: list[str] = []
        self.active_order = StringVar()
        self.multi_dept_values: dict[str, BooleanVar] = {}
        self._tooltips: list[Tooltip] = []
        self._queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._worker_active = False
        self._current_width = 1120
        self._current_breakpoint = "wide"
        self._progressbar: ttk.Progressbar | None = None
        self._history_lock = threading.Lock()
        self._update_check_running = False

        self.root.title(self.tr("app_title"))
        try:
            _set_app_icon(self.root)
        except Exception:
            pass

        self._configure_style()
        self._place_window()
        self._apply_theme_root()
        self._build_scaffold()
        self._build_header()
        self._build_body()
        self._install_drain_loop()
        self.root.bind("<Configure>", self._on_root_resize)
        self.refresh_templates()
        self._office_warmed_at = 0.0
        self.employee_name.trace_add("write", lambda *_a: self.warm_up_office())
        self._bind_shortcuts()
        self._install_file_drop()
        self.root.after(SYSTEM_THEME_POLL_MS, self._follow_system_theme)
        threading.Thread(target=cleanup_old_previews, daemon=True).start()
        if self.settings.get("check_updates") and \
                self.settings.get("last_update_check") != date.today().isoformat():
            self.root.after(2500, lambda: self.check_for_updates(manual=False))

    # ----------------------------- Utilities ------------------------------
    def tr(self, key: str, **kwargs: Any) -> str:
        raw = self.language.get(key, key)
        try:
            return raw.format(**kwargs) if kwargs else raw
        except (KeyError, IndexError):
            return raw

    def _follow_system_theme(self) -> None:
        """Con il tema "Sistema" segue i cambi di Windows chiaro/scuro mentre l'app e' aperta."""
        try:
            if self.theme_pref == "system" and not self._worker_active:
                wanted = resolve_theme("system")
                if wanted != self.theme.get():
                    self.theme.set(wanted)
                    self._rebuild_ui()
        finally:
            self.root.after(SYSTEM_THEME_POLL_MS, self._follow_system_theme)

    def _place_window(self) -> None:
        """Finestra proporzionata a schermo e DPI, centrata (prima era fissa a 1120x820)."""
        kit = getattr(self, "_kit", None)
        scale = kit.scale if kit else 1.0
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w = max(min(round(1260 * scale), sw - 40), min(980, sw))
        h = max(min(round(900 * scale), sh - 90), min(680, sh))
        self.root.minsize(min(980, sw), min(680, sh))
        x, y = max(0, (sw - w) // 2), max(0, (sh - 60 - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self._current_width = w

    def _persist_settings(self) -> None:
        self.settings["theme"] = self.theme_pref
        self.settings["language"] = self.language_code
        self.settings["last_template_dir"] = self.template_dir.get()
        self.settings["last_output_dir"] = self.output_dir.get()
        save_settings(self.settings)

    # ----------------------------- Style ----------------------------------
    def _apply_theme_root(self):
        theme = self.theme.get()
        if theme == "dark":
            self.root.configure(bg="#0b141b")
        else:
            self.root.configure(bg="#f0f4f6")
        _set_titlebar_dark(self.root, theme == "dark")
        if not getattr(self, "_toplevel_titlebar_bound", False):
            # Anche le finestre secondarie (impostazioni, storico, ...) seguono il tema
            self.root.bind_class(
                "Toplevel", "<Map>",
                lambda e: _set_titlebar_dark(e.widget, self.theme.get() == "dark")
                if e.widget.winfo_toplevel() is e.widget else None, add="+")
            self._toplevel_titlebar_bound = True

    def _configure_style(self):
        dark = self.theme.get() == "dark"
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        if dark:
            app_bg = "#0b141b"
            card_bg = "#13202a"
            card_body_bg = "#13202a"
            sh1 = "#060c10"
            sh2 = "#1f303b"
            text = "#e8eff2"
            text_muted = "#8fa7b1"
            gold = "#e0a640"
            accent_bg = "#08121a"
            accent_fg = "#e8eff2"
            title_bg = "#08121a"
            title_fg = "#f4f8fa"
            subtitle_fg = "#8fa7b1"
            section_bg = card_bg
            section_fg = text
            accent_label_bg = card_bg
            accent_label_fg = gold
            muted_bg = card_bg
            muted_fg = text_muted
            appmuted_bg = app_bg
            appmuted_fg = text_muted
            count_bg = "#10393a"
            count_fg = "#9fe3d8"
            secure_bg = count_bg
            secure_fg = count_fg
            primary_bg = "#2aa19c"
            primary_fg = "#03201f"
            primary_hover = "#36b8b2"
            primary_press = "#238a86"
            secondary_bg = "#1f303b"
            secondary_fg = text
            secondary_hover = "#2a3f4c"
            secondary_press = "#35505f"
            accent_btn_bg = gold
            accent_btn_fg = "#2a1a00"
            field_bg = "#0f1a22"
            field_fg = text
            border = "#2a3d48"
            focus = "#2aa19c"
            tree_bg = card_bg
            tree_field = card_bg
            tree_fg = text
            tree_head_bg = "#1a2a34"
            tree_head_fg = "#a9c3cc"
            tree_head_hover = "#223744"
            tree_sel_bg = "#1d4d4c"
            tree_sel_fg = text
            scroll_bg = "#2a3d48"
            scroll_trough = card_bg
            check_bg = card_bg
            check_fg = text
            gold_bg = "#3a2a08"
            gold_fg = gold
            row_even = "#172631"
            row_tutti = "#33280c"
            row_modified = "#4a3c05"
            row_new = "#12361f"
        else:
            app_bg = "#f0f4f6"
            card_bg = "#ffffff"
            card_body_bg = "#ffffff"
            sh1 = "#d5dfe4"
            sh2 = "#e6edf0"
            text = "#11293a"
            text_muted = "#5b7280"
            gold = "#e0a640"
            accent_bg = "#0b2a3d"
            accent_fg = "#ffffff"
            title_bg = "#0b2a3d"
            title_fg = "#ffffff"
            subtitle_fg = "#bcd5dd"
            section_bg = card_bg
            section_fg = text
            accent_label_bg = card_bg
            accent_label_fg = "#c08a26"
            muted_bg = card_bg
            muted_fg = text_muted
            appmuted_bg = app_bg
            appmuted_fg = text_muted
            count_bg = "#e2f3f0"
            count_fg = "#17706b"
            secure_bg = count_bg
            secure_fg = count_fg
            primary_bg = "#1f8a87"
            primary_fg = "#ffffff"
            primary_hover = "#187370"
            primary_press = "#125a58"
            secondary_bg = "#edf2f4"
            secondary_fg = text
            secondary_hover = "#dfe8eb"
            secondary_press = "#cfdce0"
            accent_btn_bg = gold
            accent_btn_fg = "#2a1a00"
            field_bg = "#f8fafb"
            field_fg = text
            border = "#d3dee3"
            focus = "#1f8a87"
            tree_bg = "#ffffff"
            tree_field = "#ffffff"
            tree_fg = text
            tree_head_bg = "#eef4f5"
            tree_head_fg = "#33505c"
            tree_head_hover = "#e1ebed"
            tree_sel_bg = "#cdeae6"
            tree_sel_fg = text
            scroll_bg = "#c9d6db"
            scroll_trough = "#f3f6f8"
            check_bg = card_bg
            check_fg = text
            gold_bg = "#fcf1dd"
            gold_fg = "#9a6716"
            row_even = "#f6fafa"
            row_tutti = "#fff8ea"
            row_modified = "#fff5c2"
            row_new = "#dcf5e3"

        style.configure("App.TFrame", background=app_bg)
        style.configure("Card.TFrame", background=card_bg)
        style.configure("CardBody.TFrame", background=card_body_bg)
        style.configure("Shadow1.TFrame", background=sh1)
        style.configure("Shadow2.TFrame", background=sh2)

        style.configure("Title.TLabel", background=title_bg, foreground=title_fg,
                        font=("Segoe UI Semibold", 26, "bold"))
        style.configure("Subtitle.TLabel", background=title_bg, foreground=subtitle_fg,
                        font=("Segoe UI", 10))
        style.configure("Eyebrow.TLabel", background=title_bg, foreground=gold,
                        font=("Segoe UI", 8, "bold"))

        style.configure("Section.TLabel", background=section_bg, foreground=section_fg,
                        font=("Segoe UI Semibold", 13, "bold"))
        style.configure("SectionAccent.TLabel", background=accent_label_bg, foreground=accent_label_fg,
                        font=("Segoe UI", 9, "bold"))
        style.configure("Muted.TLabel", background=muted_bg, foreground=muted_fg,
                        font=("Segoe UI", 9))
        style.configure("AppMuted.TLabel", background=appmuted_bg, foreground=appmuted_fg,
                        font=("Segoe UI", 9))
        style.configure("FieldLabel.TLabel", background=card_bg, foreground=text,
                        font=("Segoe UI Semibold", 9, "bold"))

        style.configure("Count.TLabel", background=count_bg, foreground=count_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(12, 6))
        style.configure("Gold.TLabel", background=gold_bg, foreground=gold_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(12, 6))
        style.configure("Secure.TLabel", background=secure_bg, foreground=secure_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(12, 6))

        style.configure("Primary.TButton", background=primary_bg, foreground=primary_fg,
                        font=("Segoe UI Semibold", 10, "bold"), padding=(18, 11),
                        borderwidth=0, focusthickness=0)
        style.map("Primary.TButton",
                  background=[("pressed", primary_press), ("active", primary_hover)])

        style.configure("Secondary.TButton", background=secondary_bg, foreground=secondary_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(14, 9),
                        borderwidth=0, focusthickness=0)
        style.map("Secondary.TButton",
                  background=[("pressed", secondary_press), ("active", secondary_hover)])

        style.configure("Accent.TButton", background=accent_btn_bg, foreground=accent_btn_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(14, 9),
                        borderwidth=0, focusthickness=0)
        style.map("Accent.TButton",
                  background=[("active", "#c08c2d"), ("pressed", "#a4741d")])

        style.configure("TEntry", fieldbackground=field_bg, foreground=field_fg,
                        bordercolor=border, lightcolor=border, darkcolor=border,
                        padding=9, focusthickness=2, focuscolor=focus,
                        font=("Segoe UI", 10), borderwidth=1, relief="solid")
        style.map("TEntry",
                  bordercolor=[("focus", focus), ("!focus", border)],
                  lightcolor=[("focus", focus), ("!focus", border)],
                  darkcolor=[("focus", focus), ("!focus", border)])

        style.configure("TCombobox", fieldbackground=field_bg, foreground=field_fg,
                        background=field_bg, arrowcolor=tree_head_fg,
                        bordercolor=border, lightcolor=border, darkcolor=border,
                        padding=8, focusthickness=2, focuscolor=focus,
                        font=("Segoe UI", 10), borderwidth=1, relief="solid")
        style.map("TCombobox",
                  bordercolor=[("focus", focus), ("!focus", border)],
                  lightcolor=[("focus", focus), ("!focus", border)],
                  darkcolor=[("focus", focus), ("!focus", border)],
                  background=[("readonly", field_bg), ("active", field_bg)],
                  fieldbackground=[("readonly", field_bg)])

        style.configure("Treeview", background=tree_bg, fieldbackground=tree_field,
                        foreground=tree_fg, rowheight=30,
                        bordercolor=border, borderwidth=1, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background=tree_head_bg, foreground=tree_head_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(9, 8),
                        relief="flat", borderwidth=0)
        style.map("Treeview",
                  background=[("selected", tree_sel_bg)],
                  foreground=[("selected", tree_sel_fg)])
        style.map("Treeview.Heading",
                  background=[("active", tree_head_hover)])

        style.configure("Vertical.TScrollbar", background=scroll_bg, troughcolor=scroll_trough,
                        bordercolor=scroll_bg, arrowcolor=tree_head_fg, arrowsize=14,
                        relief="flat", borderwidth=0, gripcount=0, width=12)
        style.map("Vertical.TScrollbar",
                  background=[("active", focus), ("disabled", scroll_trough)])

        style.configure("Horizontal.TProgressbar",
                        troughcolor=scroll_trough,
                        background=focus,
                        bordercolor=border,
                        lightcolor=focus,
                        darkcolor=focus,
                        thickness=8)

        style.configure("TCheckbutton", background=check_bg, foreground=check_fg,
                        font=("Segoe UI Semibold", 9), focusthickness=0)
        style.map("TCheckbutton",
                  background=[("active", check_bg)],
                  foreground=[("active", check_fg)])

        self._style_colors = {
            "app_bg": app_bg, "card_bg": card_bg, "card_body_bg": card_body_bg,
            "sh1": sh1, "sh2": sh2, "text": text, "text_muted": text_muted,
            "gold": gold, "accent_bg": accent_bg, "accent_fg": accent_fg,
            "title_bg": title_bg, "title_fg": title_fg, "subtitle_fg": subtitle_fg,
            "accent_label_bg": accent_label_bg, "accent_label_fg": accent_label_fg,
            "section_bg": section_bg, "section_fg": section_fg,
            "muted_bg": muted_bg, "muted_fg": muted_fg,
            "gold_bg": gold_bg, "count_bg": count_bg, "secure_bg": secure_bg,
            "field_bg": field_bg, "border": border, "focus": focus,
            "count_fg": count_fg, "primary_bg": primary_bg,
            "row_even": row_even, "row_tutti": row_tutti,
            "row_modified": row_modified, "row_new": row_new,
            "title_fg_strong": title_fg, "sh_shadow": "#000000" if dark else "#1b3a4b",
        }

        # Elementi arrotondati disegnati con ui_kit (pulsanti, campi, card...)
        if getattr(self, "_kit", None) is None:
            self._kit = UiKit(self.root)
        self._kit.install(style, {
            "surface": card_body_bg, "text": text, "muted": text_muted, "gold": gold,
            "primary": primary_bg, "primary_hover": primary_hover, "primary_press": primary_press,
            "on_primary": primary_fg,
            "accent": accent_btn_bg,
            "accent_hover": "#ecb655" if dark else "#d39a33",
            "accent_press": "#c98f2c" if dark else "#bf8628",
            "on_accent": accent_btn_fg,
            "secondary": secondary_bg, "secondary_hover": secondary_hover,
            "secondary_press": secondary_press, "on_secondary": secondary_fg,
            "secondary_border": "#2c4250" if dark else "#d9e3e7",
            "field": field_bg, "field_border": border,
            "field_hover": "#3d5664" if dark else "#a9bcc4",
            "focus": focus, "check_border": "#4a6472" if dark else "#9fb3bb",
            "disabled_bg": "#1a2832" if dark else "#eef2f4",
            "disabled_fg": "#5d7480" if dark else "#9aabb2",
            "select_bg": "#1d4d4c" if dark else "#cdeae6", "select_fg": text,
            "trough": "#1f303b" if dark else "#e4ecef",
            "thumb": "#34505e" if dark else "#c3d1d7",
            "thumb_hover": "#4a6b7a" if dark else "#9fb4bc",
            "count_bg": count_bg, "gold_bg": gold_bg, "secure_bg": secure_bg,
            "header_bg": title_bg, "header_field": "#12222c" if dark else "#14374d",
            "header_border": "#22394a" if dark else "#255069",
            "header_hover": "#35576a" if dark else "#3a6a85", "header_fg": "#f1f6f8",
        })

    # ---------------------------- Scaffold --------------------------------
    def _build_scaffold(self):
        self._tooltips.clear()
        for child in self.root.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass

    def _set_progress(self, step: int, total: int):
        total = max(1, total)
        step = max(0, min(step, total))
        if self._progressbar is not None:
            self._progressbar["maximum"] = total
            self._progressbar["value"] = step
        self.progress_label.set(
            self.tr("progress_label", step=step, total=total,
                    pct=round(step * 100.0 / total, 1))
        )

    def _post(self, kind: str, payload: Any):
        self._queue.put((kind, payload))

    # ---------------------------- Header ----------------------------------
    def _build_header(self):
        dark = self.theme.get() == "dark"
        px = self._kit.px
        # Font dei tre testi: l'altezza dell'intestazione si ricava dalle loro righe,
        # cosi' resta compatta e non taglia nulla a 125-200% di scala.
        f_eyebrow = ("Segoe UI Variable Text Semibold", 8, "bold")
        f_title = ("Segoe UI Variable Display Semib", 21, "bold")
        f_sub = ("Segoe UI Variable Text", 10)

        def linespace(font) -> int:
            return int(self.root.tk.call("font", "metrics", font, "-linespace"))

        pad_y = px(14)
        text_h = linespace(f_eyebrow) + linespace(f_title) + linespace(f_sub)
        logo_h = 0
        header_h = max(text_h, px(64)) + 2 * pad_y + px(3)

        header_outer = tk.Frame(self.root, bg=self._style_colors["app_bg"])
        header_outer.pack(fill=X, side="top", pady=(0, px(10)))

        title_bg = self._style_colors["title_bg"]
        c = tk.Canvas(header_outer, height=header_h, highlightthickness=0, bd=0, bg=title_bg)
        c.pack(fill=X, side="top")

        try:
            if LOGO_HEADER_PNG.exists():
                logo_img = tk.PhotoImage(file=str(LOGO_HEADER_PNG))
            elif ICON_PNG.exists():
                logo_img = tk.PhotoImage(file=str(ICON_PNG))
                logo_img = logo_img.subsample(max(1, logo_img.width() // 64))
            else:
                rgba, ww, hh = _build_in_memory_icon(64, 64)
                logo_img = tk.PhotoImage(data=_png_from_rgba(rgba, ww, hh))
            self._header_logo = logo_img
            logo_h = logo_img.height()
        except Exception:
            self._header_logo = None
        if logo_h > header_h - 2 * pad_y - px(3):
            header_h = logo_h + 2 * pad_y + px(3)
            c.configure(height=header_h)
        mid_y = (header_h - px(3)) // 2

        colors = self._style_colors
        # Logo e testi disegnati direttamente sul canvas: niente riquadri di
        # sfondo pieno, cosi' le decorazioni restano visibili dietro al testo.
        glow = "#1f8a87" if not dark else "#1a6f6c"
        header_state = {"w": 0, "img": None, "job": None}

        def paint_header(_evt=None):
            header_state["job"] = None
            w = max(c.winfo_width(), 1)
            if w == header_state["w"] and _evt is not None:
                return
            header_state["w"] = w
            c.delete("all")
            # Sfondo renderizzato: sfumatura, bagliori sfocati, puntinatura e filetto oro
            header_state["img"] = self._kit.header_image(
                w, header_h, title_bg, glow, colors["gold"], colors["primary_bg"])
            c.create_image(0, 0, image=header_state["img"], anchor="nw")
            x = px(32)
            if self._header_logo is not None:
                c.create_image(x, mid_y, image=self._header_logo, anchor="w")
                x += self._header_logo.width() + px(18)
            item = c.create_text(x, mid_y - text_h // 2, anchor="nw", text=self.tr("eyebrow"),
                                 fill=colors["gold"], font=f_eyebrow)
            item = c.create_text(x - 1, c.bbox(item)[3], anchor="nw",
                                 text=self.tr("header_title"), fill=colors["title_fg"],
                                 font=f_title)
            c.create_text(x, c.bbox(item)[3] - 1, anchor="nw",
                          text=self.tr("header_subtitle"), fill=colors["subtitle_fg"],
                          font=f_sub)

        def schedule_paint(_evt=None):
            # Si ridisegna solo quando il ridimensionamento si ferma: nel frattempo
            # resta l'immagine precedente, la cui parte destra ha gia' il colore di fondo.
            if header_state["job"] is not None:
                c.after_cancel(header_state["job"])
            header_state["job"] = c.after(120, paint_header, _evt)

        c.bind("<Configure>", schedule_paint)
        c.after(1, paint_header)

        controls = tk.Frame(header_outer, bg=title_bg)
        controls.place(relx=1.0, x=-px(32), y=mid_y, anchor="e")
        row = tk.Frame(controls, bg=title_bg)
        row.pack(anchor="e")

        tk.Label(row, text=self.tr("lbl_theme") + "  ", bg=title_bg,
                 fg=self._style_colors["subtitle_fg"],
                 font=("Segoe UI Semibold", 9, "bold")).pack(side=LEFT)
        theme_choices = ("light", "dark", "system")
        theme_switch = ttk.Combobox(
            row, values=[self.tr("theme_light"), self.tr("theme_dark"), self.tr("theme_system")],
            state="readonly", width=9, style="Header.TCombobox"
        )
        theme_switch.current(theme_choices.index(self.theme_pref))
        theme_switch.pack(side=LEFT, padx=(0, 18))
        tk.Label(row, text=self.tr("lbl_language") + "  ", bg=title_bg,
                 fg=self._style_colors["subtitle_fg"],
                 font=("Segoe UI Semibold", 9, "bold")).pack(side=LEFT)
        langs = available_languages()
        labels = []
        current_idx = 0
        for i, code in enumerate(langs):
            labels.append(code.upper())
            if code == self.language_var.get():
                current_idx = i
        lang_combo = ttk.Combobox(row, values=labels, state="readonly", width=6,
                                  style="Header.TCombobox")
        lang_combo.current(current_idx)
        lang_combo.pack(side=LEFT)
        def on_theme(_e=None):
            chosen = theme_switch.current()
            if not 0 <= chosen < len(theme_choices):
                return
            self.theme_pref = theme_choices[chosen]
            self._persist_settings()
            new_theme = resolve_theme(self.theme_pref)
            if new_theme != self.theme.get():
                self.theme.set(new_theme)
                self._rebuild_ui()

        def on_lang(_e=None):
            idx = lang_combo.current()
            if idx < 0 or idx >= len(langs):
                return
            code = langs[idx]
            if code == self.language_code:
                return
            self.language_code = code
            self.language_var.set(code)
            self.language = load_language(code)
            self._persist_settings()
            self._rebuild_ui()

        theme_switch.bind("<<ComboboxSelected>>", on_theme)
        lang_combo.bind("<<ComboboxSelected>>", on_lang)

    def _draw_hgradient(self, canvas, w, y, h, color1, color2):
        """Banda orizzontale alta h con sfumatura da sinistra (color1) a destra (color2)."""
        r1, g1, b1 = canvas.winfo_rgb(color1)
        r2, g2, b2 = canvas.winfo_rgb(color2)
        steps = 64
        seg = w / steps
        for i in range(steps):
            t = i / (steps - 1)
            color = "#%04x%04x%04x" % (int(r1 + (r2 - r1) * t), int(g1 + (g2 - g1) * t),
                                       int(b1 + (b2 - b1) * t))
            canvas.create_rectangle(int(i * seg), y, int((i + 1) * seg) + 1, y + h,
                                    fill=color, outline="")

    def _draw_gradient(self, canvas, w, h, color1, color2):
        steps = max(1, h)
        r1, g1, b1 = canvas.winfo_rgb(color1)
        r2, g2, b2 = canvas.winfo_rgb(color2)
        rr = (r2 - r1) / steps
        rg = (g2 - g1) / steps
        rb = (b2 - b1) / steps
        for i in range(steps):
            nr = int(r1 + rr * i)
            ng = int(g1 + rg * i)
            nb = int(b1 + rb * i)
            color = f"#{nr:04x}{ng:04x}{nb:04x}"
            canvas.create_line(0, i, w, i, fill=color)

    # ----------------------------- Body -----------------------------------
    def _build_body(self):
        outer_wrap = ttk.Frame(self.root, style="App.TFrame")
        outer_wrap.pack(fill=BOTH, expand=True)
        outer_wrap.columnconfigure(0, weight=1)
        outer_wrap.rowconfigure(0, weight=1)

        canvas_wrap = tk.Canvas(outer_wrap, background=self._style_colors["app_bg"],
                                highlightthickness=0, borderwidth=0)
        canvas_wrap.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(outer_wrap, orient="vertical", command=canvas_wrap.yview)
        sb.grid(row=0, column=1, sticky="ns")
        canvas_wrap.configure(yscrollcommand=sb.set)

        scrolled = ttk.Frame(canvas_wrap, style="App.TFrame")
        scrolled_id = canvas_wrap.create_window((0, 0), window=scrolled, anchor="nw")
        self._body_canvas = canvas_wrap
        self._body_scrolled_id = scrolled_id
        self._body_scrolled = scrolled

        def _on_scroll_config(_evt=None):
            canvas_wrap.configure(scrollregion=canvas_wrap.bbox("all"))

        def _on_canvas_config(evt):
            canvas_wrap.itemconfigure(scrolled_id, width=evt.width)

        scrolled.bind("<Configure>", _on_scroll_config)
        canvas_wrap.bind("<Configure>", _on_canvas_config)

        # Scorrimento a passi di pochi pixel: con i "units" predefiniti (1/10 della
        # finestra) ogni scatto saltava di molto e i touchpad, che mandano delta
        # piccoli, non scorrevano affatto o andavano a scatti.
        canvas_wrap.configure(yscrollincrement=self._kit.px(10) if self._kit else 10)
        # Widget che gestiscono da soli la rotellina (liste, menu a tendina...).
        # Si azzera a ogni ricostruzione dell'interfaccia (cambio tema o lingua).
        self._wheel_local_widgets = set()
        self._wheel_remainder = 0.0

        def _register_local_wheel(w):
            self._wheel_local_widgets.add(str(w))

        self._register_local_wheel = _register_local_wheel

        # I binding globali si installano una volta sola: prima venivano aggiunti a
        # ogni cambio di tema/lingua e ogni scatto della rotellina li eseguiva tutti.
        if not getattr(self, "_wheel_bound", False):
            self.root.bind_all("<MouseWheel>", lambda e: self._on_body_wheel(e, -e.delta / 120), add="+")
            self.root.bind_all("<Button-4>", lambda e: self._on_body_wheel(e, -1), add="+")
            self.root.bind_all("<Button-5>", lambda e: self._on_body_wheel(e, 1), add="+")
            self._wheel_bound = True

        self._body = body = ttk.Frame(scrolled, style="App.TFrame", padding=(16, 6, 16, 4))
        body.pack(fill=BOTH, expand=True)
        body.columnconfigure(0, weight=5)
        body.columnconfigure(1, weight=4)
        body.rowconfigure(1, weight=1)
        self._body_layout_root = body

        self._build_card1(body)
        self._build_card2(body)
        self._build_card3(body)
        self._build_footer(body)
        # Le card appena create sono in due colonne; le misure richieste dai
        # contenuti sono pronte solo dopo il primo calcolo del layout.
        self._current_breakpoint = "wide"
        self.root.after_idle(self._apply_breakpoint_when_measured)

    def _card(self, parent, title=None, subtitle=None, accent=None, **kwargs):
        colors = self._style_colors
        # Card arrotondata con bordo sottile e ombra morbida (disegnata da ui_kit)
        panel = RoundedPanel(parent, self._kit, outer_bg=colors["app_bg"],
                             fill=colors["card_body_bg"], border=colors["sh2"],
                             shadow=colors["sh_shadow"],
                             shadow_alpha=150 if self.theme.get() == "dark" else 38,
                             padx=22, pady=16)
        shadow1 = panel
        inner = panel.inner

        head = tk.Frame(inner, bg=colors["card_body_bg"])
        head.pack(fill=X)
        if title or accent:
            # "01  ·  CONFIGURAZIONE" -> pillola col numero + etichetta
            number, label = "", accent or ""
            if accent and "·" in accent:
                number, label = (part.strip() for part in accent.split("·", 1))
            if number:
                kit = self._kit
                badge = kit.pill(kit.px(38), kit.px(38), colors["count_bg"], r=kit.px(11))
                tk.Label(head, text=number, image=badge, compound="center", bd=0,
                         bg=colors["card_body_bg"], fg=colors["count_fg"],
                         font=("Segoe UI Variable Display Semib", 12, "bold"),
                         ).pack(side=LEFT, padx=(0, 14), anchor="center")
            left = tk.Frame(head, bg=colors["card_body_bg"])
            left.pack(side=LEFT, fill=X, expand=True)
            if label:
                ttk.Label(left, text=label, style="SectionAccent.TLabel").pack(anchor="w")
            if title:
                ttk.Label(left, text=title, style="Section.TLabel").pack(anchor="w", pady=(1, 0))
            # Divisore: breve tratto oro su linea sottile
            divider = tk.Frame(inner, bg=colors["card_body_bg"], height=3)
            divider.pack(fill=X, pady=(12, 14))
            tk.Frame(divider, bg=colors["sh2"]).place(x=0, y=1, relwidth=1.0, height=1)
            accent_line = self._kit.pill(self._kit.px(48), 3, colors["gold"])
            tk.Label(divider, image=accent_line, bd=0, bg=colors["card_body_bg"]
                     ).place(x=0, y=0, height=3)

        if subtitle:
            ttk.Label(inner, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(0, 12))

        body = tk.Frame(inner, bg=colors["card_body_bg"])
        body.pack(fill=BOTH, expand=True)
        return shadow1, body

    @staticmethod
    def _fluid_wrap(label, inset: int = 0) -> None:
        """Testo che va a capo sulla larghezza assegnata all'etichetta.

        Con width=1 l'etichetta non impone la propria larghezza alla card: cosi'
        una frase lunga non decide da sola se le card stanno affiancate.
        """
        label.configure(width=1)

        def fit(evt):
            wrap = max(120, evt.width - inset - 2)
            if str(label.cget("wraplength")) != str(wrap):
                label.configure(wraplength=wrap)

        label.bind("<Configure>", fit, add="+")

    def _field_label(self, parent, text, row, col=0, span=2, label_col_width=160):
        lbl = tk.Label(
            parent, text=text, bg=self._style_colors["card_body_bg"],
            fg=self._style_colors["text"],
            font=("Segoe UI Semibold", 9, "bold"), anchor="w",
        )
        lbl.grid(row=row, column=col, sticky="we", padx=(0, 16), pady=(0, 6))
        parent.grid_columnconfigure(col, minsize=label_col_width)
        return lbl

    def _wrap_field(self, parent, row, widget, col=0, span=2, pad_bottom=14):
        container = tk.Frame(parent, bg=self._style_colors["card_body_bg"])
        container.grid(row=row, column=col + 1, sticky="nsew", columnspan=span - 1, pady=(0, pad_bottom))
        container.grid_columnconfigure(0, weight=1)
        widget_master = widget.master
        widget.reparent = None  # noop
        if widget_master is not container:
            widget.pack_forget() if hasattr(widget, "pack_forget") else None
            widget.grid_forget() if hasattr(widget, "grid_forget") else None
        widget.grid(in_=container, row=0, column=0, sticky="ew")
        return container

    # ------------------------- Card 1 (Setup) -----------------------------
    def _build_card1(self, body):
        src_shadow, src_body = self._card(
            body,
            title=self.tr("card01_title"),
            accent=self.tr("card01_accent"),
            subtitle=self.tr("card01_subtitle"),
        )
        src_shadow.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 2))
        src_body.columnconfigure(1, weight=1)
        src_body.columnconfigure(2, weight=0)

        def add_row(r, label_text, var, btn_cmd):
            tk.Label(src_body, text=label_text,
                     bg=self._style_colors["card_body_bg"], fg=self._style_colors["text"],
                     font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                     ).grid(row=r, column=0, sticky="we", padx=(0, 18), pady=(0, 4))
            ent = ttk.Entry(src_body, textvariable=var)
            ent.grid(row=r, column=1, sticky="ew", padx=(0, 10), pady=(0, 14))
            btn = self._button(src_body, "btn_choose_folder",
                             style="Secondary.TButton", command=btn_cmd)
            btn.grid(row=r, column=2, sticky="ew", pady=(0, 14))
            return ent, btn

        tpl_entry, tpl_btn = add_row(0, self.tr("lbl_templates"), self.template_dir,
                                     self.choose_template_dir)
        out_entry, out_btn = add_row(1, self.tr("lbl_output"), self.output_dir,
                                     self.choose_output_dir)
        self._add_tooltip(tpl_entry, lambda: self.tr("tt_choose_tpl"))
        self._add_tooltip(tpl_btn, lambda: self.tr("tt_choose_tpl"))
        self._add_tooltip(out_entry, lambda: self.tr("tt_choose_out"))
        self._add_tooltip(out_btn, lambda: self.tr("tt_choose_out"))

        action_row = tk.Frame(src_body, bg=self._style_colors["card_body_bg"])
        action_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        refresh_btn = self._button(action_row, "btn_refresh",
                                 style="Accent.TButton", command=self.refresh_templates)
        refresh_btn.pack(side=LEFT)
        self._add_tooltip(refresh_btn, lambda: self.tr("tt_refresh") + "  (F5)")
        dept_btn = self._button(action_row, "btn_depts",
                              style="Secondary.TButton", command=self.open_department_editor)
        dept_btn.pack(side=LEFT, padx=(10, 0))
        self._add_tooltip(dept_btn, lambda: self.tr("tt_depts"))
        modules_btn = self._button(action_row, "btn_modules",
                                 style="Secondary.TButton", command=self.open_template_manager)
        modules_btn.pack(side=LEFT, padx=(10, 0))
        self._add_tooltip(modules_btn, lambda: self.tr("tt_modules"))
        settings_btn = self._button(action_row, "btn_settings",
                                  style="Secondary.TButton", command=self.open_settings)
        settings_btn.pack(side=RIGHT)
        self._add_tooltip(settings_btn, lambda: self.tr("tt_settings") + "  (Ctrl+,)")
        history_btn = self._button(action_row, "btn_history",
                                 style="Secondary.TButton", command=self.open_history)
        history_btn.pack(side=RIGHT, padx=(0, 10))
        self._add_tooltip(history_btn, lambda: self.tr("tt_history") + "  (Ctrl+H)")

    # ------------------------- Card 2 (Employee) --------------------------
    def _build_card2(self, body):
        form_shadow, form_body = self._card(
            body,
            title=self.tr("card02_title"),
            accent=self.tr("card02_accent"),
            subtitle=self.tr("card02_subtitle"),
        )
        self._card2_shadow = form_shadow
        form_shadow.grid(row=1, column=0, sticky="nsew", padx=(0, 2))
        form_body.columnconfigure(1, weight=1)

        def add_row(r, label_text, widget_or_constructor):
            tk.Label(form_body, text=label_text,
                     bg=self._style_colors["card_body_bg"], fg=self._style_colors["text"],
                     font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                     ).grid(row=r, column=0, sticky="we", padx=(0, 18), pady=(0, 4))
            w = widget_or_constructor(form_body)
            w.grid(row=r, column=1, sticky="ew", pady=(0, 14))
            return w

        name_entry = add_row(0, self.tr("lbl_name"),
            lambda parent: ttk.Entry(parent, textvariable=self.employee_name))
        self._add_tooltip(name_entry, lambda: self.tr("tt_name") + "  " + self.tr("sc_enter_hint"))
        name_entry.bind("<Return>", lambda _e: self._inline_batch_add_current())
        self._register_local_wheel(name_entry)
        dp = add_row(1, self.tr("lbl_date"),
            lambda parent: DatePickerFrame(parent, self.language,
                                           bg=self._style_colors["card_body_bg"]))
        self.date_picker = dp
        self._add_tooltip(self.date_picker, lambda: self.tr("tt_date"))
        self._register_local_wheel(self.date_picker)

        for w in (self.date_picker.day_cb, self.date_picker.month_cb, self.date_picker.year_cb):
            self._add_tooltip(w, lambda: self.tr("tt_date"))

        # --- Section: Batch inline (add people + run directly) ---
        batch_box = tk.Frame(form_body, bg=self._style_colors["card_body_bg"])
        batch_box.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 8))
        batch_box.columnconfigure(0, weight=0)
        batch_box.columnconfigure(1, weight=1)
        batch_box.rowconfigure(0, weight=1)

        list_icon = self._icon(BUTTON_ICONS.get("bat_inline_title"), "Count.TLabel")
        bt = tk.Label(
            batch_box,
            text=("  " + strip_leading_symbol(self.tr("bat_inline_title"))) if list_icon
            else "   " + self.tr("bat_inline_title"),
            image=list_icon or "", compound="left", padx=10,
            bg=self._style_colors["count_bg"], fg=self._style_colors["count_fg"],
            font=("Segoe UI Semibold", 10, "bold"), anchor="w", pady=7,
        )
        bt.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        batch_hint = ttk.Label(batch_box, text=self.tr("bat_inline_subtitle"), style="Muted.TLabel")
        batch_hint.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        self._fluid_wrap(batch_hint)

        people_tree_wrap = tk.Frame(batch_box, bg=self._style_colors["card_body_bg"])
        people_tree_wrap.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        people_tree_wrap.rowconfigure(0, weight=1)
        people_tree_wrap.columnconfigure(0, weight=1)
        batch_body_columns = ("nome", "data", "reparto")
        people_tree = ttk.Treeview(people_tree_wrap, columns=batch_body_columns,
                                   show="headings", height=8)
        for col_key, heading, width in (
            ("nome", self.tr("col_name"), 180),
            ("data", self.tr("col_date"), 110),
            ("reparto", self.tr("col_dept"), 150),
        ):
            people_tree.heading(col_key, text=heading)
            people_tree.column(col_key, width=width, anchor="w")
        people_tree.grid(row=0, column=0, sticky="nsew")
        p_sb = ttk.Scrollbar(people_tree_wrap, orient="vertical", command=people_tree.yview)
        p_sb.grid(row=0, column=1, sticky="ns")
        people_tree.configure(yscrollcommand=p_sb.set)
        self._inline_batch_tree = people_tree
        self._inline_batch_rows: list[dict[str, str]] = []
        self._register_local_wheel(self._inline_batch_tree)
        # Messaggio al centro dell'elenco vuoto (nascosto dalla prima persona aggiunta)
        self._inline_batch_empty = tk.Label(
            people_tree, text=self.tr("bat_empty_hint"), justify="center",
            bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
            font=("Segoe UI", 9), wraplength=360)
        actions = tk.Frame(batch_box, bg=self._style_colors["card_body_bg"])
        actions.grid(row=3, column=0, columnspan=2, sticky="ew")
        # Due righe: modifica dell'elenco sopra, import/modello da file sotto
        list_actions = tk.Frame(actions, bg=self._style_colors["card_body_bg"])
        list_actions.pack(fill=X)
        file_actions = tk.Frame(actions, bg=self._style_colors["card_body_bg"])
        file_actions.pack(fill=X, pady=(6, 0))
        btn_add = self._button(list_actions, "bat_add_person",
                             style="Secondary.TButton",
                             command=self._inline_batch_add_current)
        btn_add.pack(side=LEFT)
        btn_del = self._button(list_actions, "bat_del_person",
                             style="Secondary.TButton",
                             command=self._inline_batch_remove_selected)
        btn_del.pack(side=LEFT, padx=(6, 0))
        btn_clear = self._button(list_actions, "bat_clear",
                               style="Secondary.TButton",
                               command=self._inline_batch_clear)
        btn_clear.pack(side=LEFT, padx=(6, 0))
        btn_load = self._button(file_actions, "bat_import_file",
                              style="Secondary.TButton",
                              command=self._inline_batch_import_file)
        btn_load.pack(side=LEFT)
        btn_model = self._button(file_actions, "bat_template_download",
                               style="Secondary.TButton",
                               command=self.download_batch_template)
        btn_model.pack(side=LEFT, padx=(6, 0))
        self._add_tooltip(btn_model, lambda: self.tr("tt_batch_template"))

        batch_run_wrap = tk.Frame(batch_box, bg=self._style_colors["card_body_bg"])
        batch_run_wrap.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        tk.Frame(batch_run_wrap, bg=self._style_colors["gold"],
                 height=2).pack(fill=X, side="top", pady=(0, 10))
        run_count_lbl = ttk.Label(batch_run_wrap, textvariable=self._inline_batch_count(),
                                  style="Count.TLabel")
        run_count_lbl.pack(side=LEFT)
        btn_run = self._button(batch_run_wrap, "bat_run_all", style="Primary.TButton",
                             command=self._inline_batch_run)
        btn_run.pack(side=RIGHT, ipadx=8, ipady=5)
        self._inline_batch_run_btn = btn_run
        self._refresh_inline_batch_count()
        self._add_tooltip(btn_run, lambda: self.tr("tt_batch_run") + "  (Ctrl+B)")
        self._add_tooltip(btn_add, lambda: self.tr("tt_batch_add"))
        self._add_tooltip(btn_del, lambda: self.tr("tt_batch_del"))
        self._add_tooltip(btn_load, lambda: self.tr("tt_batch_import"))
        self._add_tooltip(btn_clear, lambda: self.tr("tt_batch_clear"))

        form_body.rowconfigure(2, weight=1)
        batch_box.rowconfigure(2, weight=1)

        bottom_form = tk.Frame(form_body, bg=self._style_colors["card_body_bg"])
        bottom_form.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        auto_cb = ttk.Checkbutton(bottom_form, text=self.tr("cb_autoopen"),
                                  variable=self.auto_open)
        auto_cb.pack(side=LEFT)
        self._add_tooltip(auto_cb, lambda: self.tr("tt_autoopen"))

    # ----------------------- Card 3 (Summary) -----------------------------
    def _build_card3(self, body):
        prev_shadow, prev_body = self._card(
            body,
            title=self.tr("card03_title"),
            accent=self.tr("card03_accent"),
            subtitle=self.tr("card03_subtitle"),
        )
        self._card3_shadow = prev_shadow
        prev_shadow.grid(row=1, column=1, sticky="nsew")
        prev_body.columnconfigure(0, weight=1)

        # --- Reparto selector ---
        dept_frame = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        dept_frame.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        dept_frame.columnconfigure(1, weight=1)

        multi_cb = ttk.Checkbutton(dept_frame, text=self.tr("cb_multidept"),
                                   variable=self.multi_dept_mode,
                                   command=self._toggle_multi_dept)
        multi_cb.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self._add_tooltip(multi_cb, lambda: self.tr("tt_multidept"))

        tk.Label(dept_frame, text=self.tr("lbl_department"),
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text"],
                 font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                 ).grid(row=1, column=0, sticky="w", padx=(0, 14), pady=(0, 6))

        single_wrap = tk.Frame(dept_frame, bg=self._style_colors["card_body_bg"])
        single_wrap.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        single_wrap.columnconfigure(0, weight=1)
        self.department_combo = ttk.Combobox(
            single_wrap, state="readonly", textvariable=self.department,
            font=("Segoe UI Semibold", 10, "bold"), height=14,
        )
        self.department_combo.grid(row=0, column=0, sticky="ew")
        self.department_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_department_changed())
        self._single_dept_wrap = single_wrap
        self._register_local_wheel(self.department_combo)
        multi_wrap = tk.Frame(dept_frame, bg=self._style_colors["card_body_bg"])
        multi_wrap.grid(row=2, column=0, columnspan=2, sticky="ew")
        multi_wrap.grid_remove()
        self._multi_dept_wrap = multi_wrap
        self._render_multi_dept_list(multi_wrap)

        tk.Label(dept_frame, text="  " + self.tr("lbl_department_hint").lstrip(),
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
                 font=("Segoe UI", 8), anchor="w",
                 ).grid(row=3, column=0, columnspan=2, sticky="w")

        self._add_tooltip(self.department_combo, lambda: self.tr("tt_dept"))

        self._build_order_row(prev_body, row=1)

        # --- Badge count + select all/none ---
        badge_row = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        badge_row.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(badge_row, textvariable=self.count_label, style="Count.TLabel").pack(side=LEFT)
        right_badges = tk.Frame(badge_row, bg=self._style_colors["card_body_bg"])
        right_badges.pack(side=RIGHT)
        btn_all = self._button(right_badges, "btn_selectall",
                             style="Secondary.TButton",
                             command=lambda: self._set_all_inclusion(True))
        btn_none = self._button(right_badges, "btn_selectnone",
                              style="Secondary.TButton",
                              command=lambda: self._set_all_inclusion(False))
        btn_all.pack(side=LEFT, padx=(0, 6))
        btn_none.pack(side=LEFT)
        self._add_tooltip(btn_all, lambda: self.tr("tt_include"))
        self._add_tooltip(btn_none, lambda: self.tr("tt_include"))

        # --- Treeview with checkboxes ---
        tree_wrap = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        tree_wrap.grid(row=3, column=0, sticky="nsew", pady=(0, 4))
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)
        prev_body.rowconfigure(3, weight=1)

        self.tree = ttk.Treeview(
            tree_wrap,
            columns=("include", "documento", "copie", "stato"),
            show="headings",
            height=10,
            selectmode="browse",
        )
        self.tree.heading("include", text=self.tr("col_include"))
        self.tree.heading("documento", text=self.tr("col_document"))
        self.tree.heading("copie", text=self.tr("col_copies"))
        self.tree.heading("stato", text=self.tr("col_status"))
        self.tree.column("include", width=46, anchor="center", stretch=False)
        self.tree.column("documento", width=240, anchor="w", stretch=True)
        self.tree.column("copie", width=70, anchor="center", stretch=False)
        self.tree.column("stato", width=96, anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_sb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        tree_sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=tree_sb.set)
        self.tree.bind("<Button-1>", self._on_tree_click)
        self._register_local_wheel(self.tree)

        self.tree.bind("<space>", lambda _e: self._toggle_focused_row())
        self._bind_tree_ordering()
        self._add_tooltip(self.tree, lambda: self.tr("tt_include"))
        self._build_order_tools(prev_body, row=4)

        # --- Colors hash status tags ---
        colors = self._style_colors
        self.tree.tag_configure("odd", background=colors["card_body_bg"])
        self.tree.tag_configure("even", background=colors["row_even"])
        self.tree.tag_configure("tutti", background=colors["row_tutti"])
        self.tree.tag_configure("modified", background=colors["row_modified"])
        self.tree.tag_configure("new", background=colors["row_new"])

        # --- Progress + buttons ---
        prog_wrap = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        prog_wrap.grid(row=5, column=0, sticky="ew", pady=(4, 0))
        # Con persone nell'elenco batch: ricorda che le spunte valgono per tutte
        self._batch_scope_label = tk.Label(
            prog_wrap, text="", bg=self._style_colors["card_body_bg"],
            fg=self._style_colors["focus"], font=("Segoe UI Semibold", 9),
            anchor="w", justify="left")
        self._fluid_wrap(self._batch_scope_label)
        self._progressbar = ttk.Progressbar(prog_wrap, orient="horizontal",
                                             mode="determinate", maximum=100, value=0)
        self._progressbar.pack(fill=X, side="top")
        ttk.Label(prog_wrap, textvariable=self.progress_label, style="Muted.TLabel"
                  ).pack(anchor="w", side="top", pady=(4, 0))

        gen_frame = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        gen_frame.grid(row=6, column=0, sticky="ew", pady=(14, 0))
        gen_buttons = tk.Frame(gen_frame, bg=self._style_colors["card_body_bg"])
        gen_buttons.pack(fill=X)
        gen_btn = self._button(gen_buttons, "btn_generate",
                             style="Primary.TButton", command=self.generate)
        gen_btn.pack(side=RIGHT, ipadx=8, ipady=5)
        preview_btn = self._button(gen_buttons, "btn_preview",
                                 style="Secondary.TButton", command=self.preview)
        preview_btn.pack(side=RIGHT, padx=(0, 10), ipady=5)
        self._add_tooltip(preview_btn, lambda: self.tr("tt_preview") + "  (Ctrl+P)")
        tk.Label(gen_frame, text=self.tr("lbl_generate_hint"),
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
                 font=("Segoe UI", 8), anchor="e",
                 ).pack(side=RIGHT, pady=(4, 0))
        self._add_tooltip(gen_btn, lambda: self.tr("tt_generate") + "  (Ctrl+G)")

        help_frame = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        help_frame.grid(row=7, column=0, sticky="ew", pady=(14, 0))
        tk.Frame(help_frame, bg=self._style_colors["count_bg"], width=4).pack(side=LEFT, fill="y")
        tip = tk.Label(
            help_frame, text=self.tr("help_tip_body"),
            bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
            font=("Segoe UI", 9), justify="left", anchor="w",
            padx=12, pady=6,
        )
        self._help_tip_label = tip
        tip.pack(side=LEFT, fill=X, expand=True)
        self._fluid_wrap(tip, inset=2 * 12)

        self._toggle_multi_dept()

    # ---------------------- Footer ----------------------------------------
    def _build_footer(self, body_container):
        footer_outer = tk.Frame(self.root, bg=self._style_colors["app_bg"])
        footer_outer.pack(fill=X, side="bottom")
        colors = self._style_colors
        footer_panel = RoundedPanel(footer_outer, self._kit, outer_bg=colors["app_bg"],
                                    fill=colors["card_body_bg"], border=colors["sh2"],
                                    shadow=colors["sh_shadow"],
                                    shadow_alpha=150 if self.theme.get() == "dark" else 38,
                                    radius=12, padx=16, pady=4)
        footer_panel.pack(fill=X, padx=16, pady=(0, 8))
        footer = footer_panel.inner

        status_wrap = tk.Frame(footer, bg=self._style_colors["card_body_bg"])
        status_wrap.pack(side=LEFT, fill=X, expand=True)
        tk.Label(status_wrap, text="●", bg=self._style_colors["card_body_bg"],
                 fg=self._style_colors["focus"], font=("Segoe UI", 10, "bold"),
                 ).pack(side=LEFT)
        tk.Label(status_wrap, textvariable=self.status,
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text"],
                 font=("Segoe UI Semibold", 9), anchor="w", padx=8,
                 ).pack(side=LEFT, fill=X, expand=True)
        tk.Label(status_wrap, textvariable=self.hash_stat_label,
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
                 font=("Segoe UI", 8), anchor="w", padx=14,
                 ).pack(side=LEFT)

        # Icona disegnata al posto dell'emoji: cercare un font con l'emoji costava
        # quasi mezzo secondo all'avvio, e la resa cambiava da un PC all'altro.
        secure_icon = self._icon(BUTTON_ICONS.get("secure_label"), "Secure.TLabel")
        secure_text = self.tr("secure_label")
        if secure_icon is not None:
            secure_text = " " + strip_leading_symbol(secure_text)
        secure = ttk.Label(footer, text=secure_text, style="Secure.TLabel",
                           image=secure_icon or "", compound="left")
        secure.pack(side=RIGHT)
        self._footer = footer_outer

    def _add_tooltip(self, widget, getter):
        self._tooltips.append(Tooltip(widget, getter))

    # ---------------------- Multi reparto UI toggle ----------------------
    def _render_multi_dept_list(self, parent):
        for child in parent.winfo_children():
            child.destroy()
        self.multi_dept_values.clear()
        options = department_options(self.templates)
        if not options:
            tk.Label(parent, text="—", bg=self._style_colors["card_body_bg"],
                     fg=self._style_colors["text_muted"]).pack(anchor="w")
            return
        cols = min(3, max(1, (len(options) + 3) // 4))
        for i, dept in enumerate(options):
            var = BooleanVar(value=False)
            self.multi_dept_values[dept] = var
            cb = ttk.Checkbutton(parent, text=dept, variable=var,
                                 command=self.update_document_list)
            cb.grid(row=i // cols, column=i % cols, sticky="w", padx=(0, 10), pady=2)
            self._add_tooltip(cb, lambda d=dept: self.tr("tt_dept") + f" [{d}]")

    def _toggle_multi_dept(self):
        multi = self.multi_dept_mode.get()
        if multi:
            self._single_dept_wrap.grid_remove()
            self._multi_dept_wrap.grid()
            self._render_multi_dept_list(self._multi_dept_wrap)
            self.department.set("")
        else:
            self._multi_dept_wrap.grid_remove()
            self._single_dept_wrap.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        self.update_document_list()

    def _current_departments(self) -> list[str]:
        if self.multi_dept_mode.get():
            return [d for d, v in self.multi_dept_values.items() if v.get()]
        d = self.department.get().strip().upper()
        return [d] if d else []

    # ---------------------- Treeview inclusion logic ----------------------
    def _current_selected_templates(self) -> list[TemplateFile]:
        depts = self._current_departments()
        if not depts:
            return []
        return [t for t in self._listed_templates()
                if self.template_inclusion.get(t.path, True)]

    def _listed_templates(self) -> list[TemplateFile]:
        """Moduli dei reparti scelti, nell'ordine della lista (anche quelli esclusi)."""
        depts = self._current_departments()
        if not depts:
            return []
        if len(depts) == 1:
            chosen = templates_for_department(self.templates, depts[0])
        else:
            chosen = templates_for_departments(self.templates, depts)
        return apply_order(chosen, self.doc_order, self._template_folder())

    def _set_all_inclusion(self, value: bool):
        for path in list(self.template_inclusion.keys()):
            self.template_inclusion[path] = value
        self._mark_order_changed()
        self.update_document_list()

    def _on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item:
            return
        if region == "cell" and col == "#1":
            self._toggle_row(item)
            return
        if region == "heading" and col == "#1":
            current_all = all(
                self.template_inclusion.get(self._row_path.get(item), True)
                for item in self.tree.get_children()
            ) if self.tree.get_children() else True
            self._set_all_inclusion(not current_all)

    def _toggle_focused_row(self):
        item = self.tree.focus()
        if item:
            self._toggle_row(item)

    def _toggle_row(self, item):
        path = self._row_path.get(item)
        if path is None:
            return
        self.template_inclusion[path] = not self.template_inclusion.get(path, True)
        self._mark_order_changed()
        self.update_document_list()

    # ---------------------- Rotellina del mouse --------------------------
    def _wheel_over_local_widget(self, evt) -> bool:
        """True se il puntatore e' su un widget che scorre per conto suo."""
        local = getattr(self, "_wheel_local_widgets", set())
        if not local:
            return False
        try:
            widget = self.root.winfo_containing(evt.x_root, evt.y_root)
        except (tk.TclError, KeyError):
            return False
        path = str(widget) if widget is not None else ""
        # Basta risalire il percorso Tk (".!frame.!canvas...") fino alla radice
        while path:
            if path in local:
                return True
            path = path.rpartition(".")[0]
        return False

    def _on_body_wheel(self, evt, notches: float):
        canvas = getattr(self, "_body_canvas", None)
        if canvas is None or self._wheel_over_local_widget(evt):
            return "break"
        try:
            if not canvas.winfo_exists() or str(canvas.winfo_toplevel()) != str(evt.widget.winfo_toplevel()):
                return None
        except (tk.TclError, AttributeError, KeyError):
            return None
        # 6 passi da 10 px per scatto; le frazioni dei touchpad si accumulano
        self._wheel_remainder += notches * 6
        steps = int(self._wheel_remainder)
        if steps:
            self._wheel_remainder -= steps
            canvas.yview_scroll(steps, "units")
        return "break"

    # ---------------------- Breakpoint responsive ------------------------
    def _on_root_resize(self, evt):
        if evt.widget is not self.root:
            return
        self._current_width = evt.width
        try:
            self._apply_breakpoint()
        except Exception:
            pass

    def _two_columns_fit(self) -> bool:
        """True se le card 02 e 03 stanno affiancate senza tagliare pulsanti o testi."""
        cards = (getattr(self, "_card2_shadow", None), getattr(self, "_card3_shadow", None))
        if not self._current_width or any(c is None or not c.winfo_exists() for c in cards):
            return True
        kit = getattr(self, "_kit", None)
        chrome = 2 * 16 + 2 + (kit.px(12) if kit else 12) + 8  # margini, spazio, barra
        return self._current_width >= sum(c.required_width() for c in cards) + chrome

    def _apply_breakpoint_when_measured(self):
        try:
            self.root.update_idletasks()  # completa il calcolo delle misure richieste
            self._apply_breakpoint()
        except tk.TclError:
            pass

    def _apply_breakpoint(self):
        want = "wide" if self._two_columns_fit() else "narrow"
        if want == self._current_breakpoint:
            return
        self._current_breakpoint = want
        layout = getattr(self, "_body_layout_root", None)
        if layout is None:
            return
        if want == "narrow":
            if self._card2_shadow is not None:
                self._card2_shadow.grid_configure(row=1, column=0, columnspan=2,
                                                   sticky="nsew", padx=(0, 0), pady=(0, 2))
            if self._card3_shadow is not None:
                self._card3_shadow.grid_configure(row=2, column=0, columnspan=2, sticky="nsew")
            layout.rowconfigure(1, weight=1)
            layout.rowconfigure(2, weight=1)
        else:
            if self._card2_shadow is not None:
                self._card2_shadow.grid_configure(row=1, column=0, columnspan=1,
                                                   sticky="nsew", padx=(0, 2), pady=(0, 0))
            if self._card3_shadow is not None:
                self._card3_shadow.grid_configure(row=1, column=1, columnspan=1, sticky="nsew")
            layout.rowconfigure(2, weight=0)

    # ---------------------- Rebuild on theme/lang change ------------------
    def _rebuild_ui(self):
        self._configure_style()
        self._apply_theme_root()
        self.root.title(self.tr("app_title"))
        saved_date = None
        try:
            if hasattr(self, "date_picker") and self.date_picker is not None:
                try:
                    if self.date_picker.winfo_exists():
                        saved_date = self.date_picker.get_date()
                except Exception:
                    saved_date = None
        except Exception:
            saved_date = None
        saved_batch: list[dict[str, str]] = []
        saved_batch_tree_rows: list[tuple] = []
        try:
            if hasattr(self, "_inline_batch_rows") and isinstance(self._inline_batch_rows, list):
                saved_batch = [dict(r) for r in self._inline_batch_rows]
            if hasattr(self, "_inline_batch_tree") and self._inline_batch_tree is not None:
                try:
                    tree = self._inline_batch_tree
                    if tree.winfo_exists():
                        for c in tree.get_children():
                            try:
                                saved_batch_tree_rows.append(tuple(tree.item(c, "values")))
                            except Exception:
                                continue
                except Exception:
                    pass
        except Exception:
            pass
        # Status/count/progress StringVar: reapply keys that are language-dependent
        self.status.set(self.tr("status_initial"))
        self.count_label.set(self.tr("count_none"))
        self.progress_label.set(self.tr("progress_ready"))
        for child in list(self.root.winfo_children()):
            try:
                child.destroy()
            except Exception:
                pass
        self._build_header()
        self._build_body()
        self.refresh_templates()
        try:
            if saved_date is not None and hasattr(self, "date_picker") and self.date_picker is not None:
                try:
                    if self.date_picker.winfo_exists():
                        self.date_picker.set_date(saved_date)
                        self.date_picker.configure_language(self.language)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            if saved_batch and hasattr(self, "_inline_batch_rows"):
                self._inline_batch_rows = [dict(r) for r in saved_batch]
                tree = getattr(self, "_inline_batch_tree", None)
                if tree is not None and tree.winfo_exists():
                    for c in tree.get_children():
                        tree.delete(c)
                    rows_to_use = (saved_batch_tree_rows
                                   if len(saved_batch_tree_rows) == len(saved_batch)
                                   else [(r.get("Nome", ""), r.get("Data", ""),
                                          r.get("Reparto", "")) for r in saved_batch])
                    for vals in rows_to_use:
                        tree.insert("", END, values=tuple(vals))
                self._refresh_inline_batch_count()
        except Exception:
            pass

    # ---------------------- Dir choices -----------------------------------
    def choose_template_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.template_dir.get(),
                                         title=self.tr("btn_choose_folder"))
        if chosen:
            self.template_dir.set(chosen)
            self._persist_settings()
            self.refresh_templates()

    def choose_output_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.output_dir.get(),
                                         title=self.tr("btn_choose_folder"))
        if chosen:
            self.output_dir.set(chosen)
            self._persist_settings()

    def warm_up_office(self) -> None:
        """Avvia Word/Excel in background mentre si compila il modulo: alla
        generazione sono gia' aperti. Al piu' una richiesta ogni 30 secondi."""
        now = time.monotonic()
        if now - self._office_warmed_at < 30 or not self.templates:
            return
        self._office_warmed_at = now
        try:
            warm_up_office({t.path.suffix for t in self.templates})
        except Exception:  # noqa: BLE001 - solo un'ottimizzazione
            traceback.print_exc()

    # ---------------------- Refresh templates -----------------------------
    def refresh_templates(self) -> None:
        folder = self._template_folder()
        self.templates, self.ignored = discover_templates(folder)
        if folder != getattr(self, "_module_settings_folder", None):
            # Altra cartella: altri nomi e ordini, si riparte dall'ordine predefinito
            self.module_settings = load_module_settings(folder)
            self._module_settings_folder = folder
            self.doc_order = []
            self.active_order.set("")
        departments = department_options(self.templates)
        self.department_combo["values"] = departments
        if departments and self.department.get().upper() not in departments:
            self.department.set(departments[0])
            self._apply_department_order()
        elif not departments:
            self.department.set("")
        if hasattr(self, "_multi_dept_wrap") and self._multi_dept_wrap is not None:
            self._render_multi_dept_list(self._multi_dept_wrap)
        for tpl in self.templates:
            if tpl.path not in self.template_inclusion:
                self.template_inclusion[tpl.path] = True
        current_paths = {t.path for t in self.templates}
        stale = [p for p in self.template_inclusion if p not in current_paths]
        for s in stale:
            del self.template_inclusion[s]
        self.hash_status = classify_template_hashes(self.templates, self.saved_hashes)
        self._refresh_order_choices()
        self.update_document_list()
        ok = sum(1 for v in self.hash_status.values() if v == "ok")
        mod = sum(1 for v in self.hash_status.values() if v == "modified")
        new = sum(1 for v in self.hash_status.values() if v == "new")
        self.hash_stat_label.set(self.tr("hash_status", ok=ok, mod=mod, new=new))
        if not folder.exists():
            self.status.set(self.tr("status_no_folder"))
        elif not self.templates:
            self.status.set(self.tr("status_no_templates"))
        elif self.ignored:
            self.status.set(self.tr("status_templates_ignored", n=len(self.templates),
                                    ignored=len(self.ignored)))
        else:
            self.status.set(self.tr("status_templates_ready", n=len(self.templates)))

    def update_document_list(self) -> None:
        if not hasattr(self, "tree") or self.tree is None:
            return
        keep = self._selected_tree_path()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._row_path: dict[str, Path] = {}
        selected = self._listed_templates()
        total_copies = sum(template.copies for template in selected
                           if self.template_inclusion.get(template.path, True))
        total_templates = sum(1 for template in selected
                              if self.template_inclusion.get(template.path, True))
        self.count_label.set(self.tr("count_label", total=total_copies, templates=total_templates))
        for idx, template in enumerate(selected):
            scope = "Tutti" if template.is_for_every_department else template.department.upper()
            included = self.template_inclusion.get(template.path, True)
            mark = "☑" if included else "☐"
            status_key = self.hash_status.get(str(template.path), "ok")
            if status_key == "modified":
                status_text = self.tr("status_modified")
            elif status_key == "new":
                status_text = self.tr("status_new")
            else:
                status_text = self.tr("status_ok")
            tags: list[str] = []
            if template.is_for_every_department:
                tags.append("tutti")
            elif idx % 2 == 0:
                tags.append("even")
            else:
                tags.append("odd")
            if status_key in {"modified", "new"}:
                tags.append(status_key)
            item = self.tree.insert(
                "", END,
                values=(mark, f"  {scope} · {self._template_label(template)}",
                        template.copies, status_text),
                tags=tuple(tags),
            )
            self._row_path[item] = template.path
            if template.path == keep:
                # Dopo spunte e spostamenti resta selezionato lo stesso documento
                self.tree.selection_set(item)
                self.tree.focus(item)
                self.tree.see(item)

    # ---------------------- Generate single -------------------------------
    def _collect_form(self, require_name: bool = True):
        """Valida il modulo; ritorna (nome, data, reparti, template) oppure None."""
        name = self.employee_name.get().strip()
        entry_date = self.date_picker.get_string()
        departments = self._current_departments()
        if not name and require_name:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_name"))
            return None
        if not entry_date:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_date"))
            return None
        if not departments:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_dept"))
            return None
        selected = self._current_selected_templates()
        if not selected:
            messagebox.showwarning(self.tr("mb_no_docs_title"), self.tr("mb_no_docs_body"))
            return None
        return name, entry_date, departments, selected

    def preview(self) -> None:
        """Dossier di prova in una cartella temporanea: una copia per modulo e filigrana."""
        if self._worker_active:
            return
        form = self._collect_form(require_name=False)
        if form is None:
            return
        name, entry_date, departments, selected = form
        name = name or self.tr("pv_sample_name")
        single_copies = [dc_replace(t, copies=1) for t in selected]
        output_path = Path(tempfile.gettempdir()) / (
            f"{PREVIEW_PREFIX}{os.getpid()}_{datetime.now():%H%M%S%f}.pdf")
        watermark = self.tr("pv_watermark")
        self._worker_active = True
        self.status.set(self.tr("pv_status_running"))
        self._post("progress_ready_label", None)

        def progress_cb(step: int, total: int):
            self._post("progress", {"step": step, "total": total})

        def work():
            try:
                build_pdf(output_path, name, entry_date, "+".join(departments), "", "",
                          single_copies, progress_cb=progress_cb, watermark=watermark)
                self._post("done_preview", {"ok": True, "path": str(output_path)})
            except Exception as error:  # noqa: BLE001
                traceback.print_exc()
                self._post("done_preview", {"ok": False, "error": str(error)})
            finally:
                self._post("worker_done", None)

        threading.Thread(target=work, daemon=True).start()

    def generate(self) -> None:
        if self._worker_active:
            return
        form = self._collect_form()
        if form is None:
            return
        name, entry_date, departments, selected = form
        output_dir_path = Path(self.output_dir.get()).expanduser()
        dept_tag = safe_file_part("+".join(departments))
        base = output_dir_path / f"dossier_{safe_file_part(name)}_{dept_tag}.pdf"
        output_path = base
        counter = 2
        while output_path.exists():
            output_path = output_dir_path / (
                f"dossier_{safe_file_part(name)}_{dept_tag}_{counter}.pdf"
            )
            counter += 1
        dept_str = "+".join(departments)
        self._worker_active = True
        self.status.set(self.tr("mb_progress_status"))
        self._post("progress_ready_label", None)
        self._persist_settings()
        # Ruolo e note disabilitati: valori statici vuoti (rimossi da UI in v2)
        ruolo = ""
        note = ""
        auto_open_value = bool(self.auto_open.get())
        saved_snapshot = dict(self.saved_hashes) if self.saved_hashes else {}
        tr_done = self.tr("mb_status_done", name=output_path.name)
        tr_error = self.tr("mb_status_error")
        pdf_options = self._pdf_options()

        def progress_cb(step: int, total: int):
            self._post("progress", {"step": step, "total": total})

        def work():
            try:
                total = build_pdf(output_path, name, entry_date, dept_str,
                                  ruolo, note, selected,
                                  progress_cb=progress_cb, **pdf_options)
                self._save_history(output_path, name, dept_str, total, entry_date)
                hashes = dict(saved_snapshot)
                for t in selected:
                    try:
                        hashes[str(t.path)] = compute_template_hash(t.path)
                    except Exception:
                        pass
                save_hashes(hashes)
                self._post("status", tr_done)
                self._post("done_single", {
                    "ok": True, "path": str(output_path), "total": total,
                    "auto_open": auto_open_value,
                })
            except Exception as error:  # noqa: BLE001
                traceback.print_exc()
                self._post("status", tr_error)
                self._post("done_single", {"ok": False, "error": str(error)})
            finally:
                self._post("worker_done", None)

        threading.Thread(target=work, daemon=True).start()

    # ---------------------- Worker -> mainloop routing --------------------
    def _drain_queue(self):
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "progress":
                    self._set_progress(payload["step"], payload["total"])
                elif kind == "status":
                    self.status.set(str(payload))
                elif kind == "progress_ready_label":
                    self.progress_label.set(self.tr("progress_ready"))
                    if self._progressbar is not None:
                        self._progressbar["value"] = 0
                elif kind == "done_single":
                    if payload.get("ok"):
                        pdf = Path(str(payload.get("path", "")))
                        message = self.tr("ts_done_single", name=pdf.name,
                                          total=payload.get("total", 0))
                        self._toast(message, "success", (self.tr("ts_open_pdf"),
                                                         lambda: open_folder(pdf)), 9000)
                        notify_windows(self.root, self.tr("mb_done_title"), message)
                        if payload.get("auto_open"):
                            open_folder(Path(str(payload["path"])).parent)
                        self.refresh_templates()
                    else:
                        self._show_error(str(payload.get("error", "")))
                elif kind == "done_preview":
                    if payload.get("ok"):
                        self.status.set(self.tr("pv_status_done"))
                        open_folder(Path(str(payload["path"])))
                    else:
                        self.status.set(self.tr("mb_status_error"))
                        self._show_error(str(payload.get("error", "")))
                elif kind == "update_result":
                    self._handle_update_result(payload)
                elif kind == "update_downloaded":
                    self._run_setup(Path(str(payload)))
                elif kind == "update_failed":
                    self._update_failed(payload)
                elif kind == "done_batch":
                    counts = (
                        f"{payload.get('ok', 0)} " + self.tr("bat_summary_ok") +
                        "   ·   " + f"{payload.get('skip', 0)} " + self.tr("bat_summary_skip") +
                        "   ·   " + f"{payload.get('fail', 0)} " + self.tr("bat_summary_fail")
                    )
                    summary = counts + "\n\n" + str(payload.get("detail", ""))
                    if payload.get("auto_open") and payload.get("out_dir"):
                        open_folder(Path(str(payload["out_dir"])))
                    problems = payload.get("skip", 0) or payload.get("fail", 0)
                    self._toast(
                        self.tr("bat_summary_title") + "  ·  " + counts,
                        "warning" if problems else "success",
                        (self.tr("ts_details"),
                         lambda text=summary: messagebox.showinfo(self.tr("bat_summary_title"), text)),
                        12000,
                    )
                    notify_windows(self.root, self.tr("bat_summary_title"), counts)
                    self.refresh_templates()
                elif kind == "worker_done":
                    self._worker_active = False
                    self.saved_hashes = load_saved_hashes()
                    self.hash_status = classify_template_hashes(self.templates, self.saved_hashes)
                    ok = sum(1 for v in self.hash_status.values() if v == "ok")
                    mod = sum(1 for v in self.hash_status.values() if v == "modified")
                    new = sum(1 for v in self.hash_status.values() if v == "new")
                    self.hash_stat_label.set(self.tr("hash_status", ok=ok, mod=mod, new=new))
        except queue.Empty:
            pass
        finally:
            self.root.after(80, self._drain_queue)

    def _install_drain_loop(self):
        self.root.after(80, self._drain_queue)


def main() -> None:
    if tk is None:
        raise SystemExit(
            "Tkinter non è disponibile in Python. Installa una versione di Python "
            "con il supporto Tk per avviare l'interfaccia desktop."
        )
    setup_error_log()
    _enable_dpi_awareness()
    root = tk.Tk()
    FormazioniApp(root)
    root.mainloop()
