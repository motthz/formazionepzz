"""Pulsanti con icona, avvisi, scorciatoie e trascinamento file."""

from __future__ import annotations

import base64
import os
import traceback
from pathlib import Path

from ..config import SUPPORTED_EXTENSIONS
from ..icons import BUTTON_ICONS, draw_icon_png, strip_leading_symbol
from ..logs import log_file
from ..system import FileDropTarget
from ..tkcompat import LEFT, messagebox, tk, ttk


class FeedbackMixin:
    """Parte di FormazioniApp: pulsanti con icona, avvisi, scorciatoie e trascinamento file."""

    # ---------------------- Tooltip helper --------------------------------
    def _icon(self, name: str | None, style: str = "TButton"):
        """PhotoImage dell'icona nel colore del testo dello stile ttk (None se assente)."""
        if not name:
            return None
        try:
            color = str(ttk.Style().lookup(style, "foreground") or self._style_colors["text"])
            if not color.startswith("#"):
                color = "#%02x%02x%02x" % tuple(c // 257 for c in self.root.winfo_rgb(color))
            size = max(14, round(16 * self.root.winfo_fpixels("1i") / 96))
            cache = self.__dict__.setdefault("_icon_cache", {})
            key = (name, color.lower(), size)
            if key not in cache:
                png = draw_icon_png(name, color, size)
                cache[key] = tk.PhotoImage(data=base64.b64encode(png)) if png else None
            return cache[key]
        except Exception:  # noqa: BLE001 - senza PIL si resta con il testo
            return None

    def _button(self, parent, key: str, **kwargs):
        """ttk.Button con testo tradotto e icona a sinistra al posto dell'emoji."""
        text = self.tr(key)
        image = self._icon(BUTTON_ICONS.get(key), kwargs.get("style", "TButton"))
        if image is not None:
            kwargs.update(image=image, compound="left")
            text = " " + strip_leading_symbol(text)
        return ttk.Button(parent, text=text, **kwargs)

    def _show_error(self, message: str) -> None:
        """Errore bloccante con l'indicazione del registro, da allegare alle segnalazioni."""
        messagebox.showerror(self.tr("mb_error_title"),
                             message + "\n\n" + self.tr("log_hint", path=log_file()))

    # ---------------------- Notifiche non bloccanti ----------------------
    def _toast(self, message: str, kind: str = "info", action=None, duration: int = 6000):
        """Avviso in basso nella finestra che sparisce da solo; action=(testo, funzione)."""
        colors = self._style_colors
        self._close_toast()
        bg = colors["title_bg"]
        accent = {"success": "#3fb27f", "warning": colors["gold"],
                  "error": "#e5534b"}.get(kind, colors["primary_bg"])
        frame = tk.Frame(self.root, bg=bg, padx=14, pady=10)
        tk.Frame(frame, bg=accent, width=4).pack(side=LEFT, fill="y", padx=(0, 12))
        tk.Label(frame, text=message, bg=bg, fg="#ffffff", font=("Segoe UI Semibold", 9),
                 justify="left", anchor="w", wraplength=560).pack(side=LEFT)
        if action is not None:
            label, callback = action
            link = tk.Label(frame, text=label, bg=bg, fg=colors["gold"], cursor="hand2",
                            font=("Segoe UI Semibold", 9, "bold"), padx=14)
            link.pack(side=LEFT)
            link.bind("<Button-1>", lambda _e: (self._close_toast(), callback()))
        close = tk.Label(frame, text="✕", bg=bg, fg="#bcd5dd", cursor="hand2",
                         font=("Segoe UI", 9), padx=4)
        close.pack(side=LEFT, padx=(6, 0))
        close.bind("<Button-1>", lambda _e: self._close_toast())
        frame.place(relx=0.5, rely=1.0, y=-96, anchor="s")
        frame.lift()
        self._toast_frame = frame
        self._toast_after = self.root.after(duration, self._close_toast)

    def _close_toast(self):
        after_id = getattr(self, "_toast_after", None)
        if after_id:
            try:
                self.root.after_cancel(after_id)
            except tk.TclError:
                pass
        frame = getattr(self, "_toast_frame", None)
        if frame is not None:
            try:
                frame.destroy()
            except tk.TclError:
                pass
        self._toast_frame = self._toast_after = None

    # ---------------------- Scorciatoie da tastiera ----------------------
    SHORTCUTS = (
        ("Ctrl+G", "sc_generate"), ("Ctrl+P", "sc_preview"), ("Ctrl+B", "sc_batch"),
        ("Ctrl+H", "sc_history"), ("Ctrl+,", "sc_settings"), ("F5", "sc_refresh"),
        ("sc_key_enter", "sc_enter"), ("F1", "sc_help"),
    )

    def _bind_shortcuts(self):
        def run(action):
            # Le scorciatoie valgono solo nella finestra principale e mai a generazione in corso
            if self.root.grab_current() is None:
                action()
            return "break"

        for sequence, action in (
            ("g", self.generate), ("p", self.preview), ("b", self._inline_batch_run),
            ("h", self.open_history), ("comma", self.open_settings),
        ):
            # Le lettere anche maiuscole (Bloc Maiusc attivo); "comma" e' un keysym fisso
            variants = {sequence, sequence.upper()} if len(sequence) == 1 else {sequence}
            for variant in variants:
                self.root.bind(f"<Control-{variant}>", lambda _e, a=action: run(a))
        # Nei campi di testo Tk usa Ctrl+H come Backspace (stile Emacs): si disattiva,
        # cosi' Ctrl+H apre lo storico senza cancellare l'ultima lettera.
        for entry_class in ("TEntry", "Entry", "TCombobox"):
            for variant in ("<Control-h>", "<Control-H>"):
                self.root.bind_class(entry_class, variant, lambda _e: None)
        self.root.bind("<F5>", lambda _e: run(self.refresh_templates))
        self.root.bind("<F1>", lambda _e: run(self.show_shortcuts))

    def show_shortcuts(self):
        lines = [f"{self.tr(keys) if keys.startswith('sc_') else keys}   →   {self.tr(key)}"
                 for keys, key in self.SHORTCUTS]
        messagebox.showinfo(self.tr("sc_title"), "\n".join(lines))

    # ---------------------- Trascinamento file ---------------------------
    def _install_file_drop(self):
        if os.name != "nt":
            return
        try:
            self._drop_target = FileDropTarget(self.root, self._on_files_dropped)
        except Exception:  # noqa: BLE001 - funzione accessoria
            traceback.print_exc()

    def _on_files_dropped(self, paths: list[Path], x: int, y: int) -> None:
        """CSV, o Excel lasciato sull'elenco dipendenti -> import batch;
        Word/Excel/PDF altrove -> nuovi moduli (template)."""
        if self._worker_active:
            return
        widget = self.root.winfo_containing(x, y)
        card2 = getattr(self, "_card2_shadow", None)
        on_batch = widget is not None and card2 is not None and str(widget).startswith(str(card2))
        batch_files = [f for f in paths if f.suffix.lower() == ".csv"
                       or (on_batch and f.suffix.lower() == ".xlsx")]
        modules = [f for f in paths if f not in batch_files
                   and f.suffix.lower() in SUPPORTED_EXTENSIONS and f.is_file()]
        for batch_file in batch_files:
            self._inline_batch_import_file(batch_file)
        if modules:
            self.open_template_manager(initial_files=modules)
        if not batch_files and not modules:
            self._toast(self.tr("dd_unsupported"), "warning")
