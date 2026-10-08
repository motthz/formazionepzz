"""Finestra Impostazioni e controllo aggiornamenti."""

from __future__ import annotations

import os
import subprocess
import tempfile
import threading
import traceback
import webbrowser
from datetime import date
from pathlib import Path
from typing import Any

from ..config import APP_DIR, APP_VERSION, DEFAULT_UPDATE_SOURCE
from ..logs import log_file
from ..system import _set_app_icon, open_folder
from ..tkcompat import BOTH, LEFT, RIGHT, BooleanVar, StringVar, messagebox, tk, ttk
from ..updates import download_setup, fetch_update_manifest, is_installed_copy, parse_version


class SettingsMixin:
    """Parte di FormazioniApp: finestra Impostazioni e controllo aggiornamenti."""

    # ---------------------- Impostazioni ----------------------------------
    def open_settings(self):
        colors = self._style_colors
        win = tk.Toplevel(self.root)
        win.title(self.tr("st_title"))
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        win.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(win)
        except Exception:
            pass
        body = tk.Frame(win, bg=colors["card_body_bg"], padx=22, pady=20)
        body.pack(fill=BOTH, expand=True, padx=14, pady=14)
        body.columnconfigure(0, weight=1)

        def section(text, row):
            tk.Label(body, text=text, bg=colors["card_body_bg"], fg=colors["accent_label_fg"],
                     font=("Segoe UI", 9, "bold"), anchor="w"
                     ).grid(row=row, column=0, sticky="w", pady=(0, 8))

        def hint(text, row):
            tk.Label(body, text=text, bg=colors["card_body_bg"], fg=colors["text_muted"],
                     font=("Segoe UI", 8), anchor="w", justify="left", wraplength=460
                     ).grid(row=row, column=0, sticky="w", pady=(0, 14))

        watermark_var = StringVar(value=str(self.settings.get("pdf_watermark") or ""))
        protect_var = BooleanVar(value=bool(self.settings.get("pdf_protect")))
        updates_var = BooleanVar(value=bool(self.settings.get("check_updates")))
        source_var = StringVar(value=str(self.settings.get("update_source") or ""))

        section(self.tr("st_pdf_section"), 0)
        tk.Label(body, text=self.tr("st_watermark"), bg=colors["card_body_bg"], fg=colors["text"],
                 font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                 ).grid(row=1, column=0, sticky="w", pady=(0, 4))
        wm_entry = ttk.Entry(body, textvariable=watermark_var, width=52)
        wm_entry.grid(row=2, column=0, sticky="ew", pady=(0, 4))
        hint(self.tr("st_watermark_hint"), 3)
        ttk.Checkbutton(body, text=self.tr("st_protect"), variable=protect_var
                        ).grid(row=4, column=0, sticky="w", pady=(0, 4))
        hint(self.tr("st_protect_hint"), 5)

        section(self.tr("st_update_section"), 6)
        ttk.Checkbutton(body, text=self.tr("st_check_updates"), variable=updates_var
                        ).grid(row=7, column=0, sticky="w", pady=(0, 8))
        tk.Label(body, text=self.tr("st_update_source"), bg=colors["card_body_bg"],
                 fg=colors["text"], font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                 ).grid(row=8, column=0, sticky="w", pady=(0, 4))
        ttk.Entry(body, textvariable=source_var, width=52
                  ).grid(row=9, column=0, sticky="ew", pady=(0, 4))
        hint(self.tr("st_update_hint", version=APP_VERSION), 10)

        def store():
            self.settings["pdf_watermark"] = watermark_var.get().strip()
            self.settings["pdf_protect"] = bool(protect_var.get())
            self.settings["check_updates"] = bool(updates_var.get())
            self.settings["update_source"] = source_var.get().strip()
            self._persist_settings()

        def check_now():
            store()
            self.check_for_updates(manual=True)

        def save_and_close(_evt=None):
            store()
            win.destroy()

        btns = tk.Frame(body, bg=colors["card_body_bg"])
        btns.grid(row=11, column=0, sticky="ew", pady=(6, 0))
        self._button(btns, "st_check_now", style="Secondary.TButton",
                   command=check_now).pack(side=LEFT)
        # Registro errori: da allegare alle segnalazioni su GitHub
        self._button(btns, "st_open_log", style="Secondary.TButton",
                     command=lambda: open_folder(log_file() if log_file().exists() else log_file().parent)
                     ).pack(side=LEFT, padx=(8, 0))
        self._button(btns, "de_save", style="Primary.TButton",
                   command=save_and_close).pack(side=RIGHT)
        self._button(btns, "de_cancel", style="Secondary.TButton",
                   command=win.destroy).pack(side=RIGHT, padx=(0, 10))
        win.bind("<Escape>", lambda _e: win.destroy())
        wm_entry.focus_set()

    # ---------------------- Aggiornamenti ---------------------------------
    def check_for_updates(self, manual: bool = False) -> None:
        # Campo vuoto = release pubblicate su GitHub
        source = str(self.settings.get("update_source") or "").strip() or DEFAULT_UPDATE_SOURCE
        if self._update_check_running:
            return
        self._update_check_running = True

        def work():
            try:
                manifest = fetch_update_manifest(source)
                self._post("update_result", {"manual": manual, "manifest": manifest})
            except FileNotFoundError:
                # Cartella di rete o version.json inesistente: messaggio comprensibile
                self._post("update_result", {"manual": manual,
                                             "error": self.tr("up_err_missing", path=source)})
            except Exception as error:  # noqa: BLE001 - rete/percorso non raggiungibile
                self._post("update_result", {"manual": manual, "error": str(error)})

        threading.Thread(target=work, daemon=True).start()

    def _handle_update_result(self, payload: dict[str, Any]) -> None:
        self._update_check_running = False
        manual = bool(payload.get("manual"))
        if payload.get("error"):
            # Controllo automatico: offline o sorgente irraggiungibile non deve disturbare
            if manual:
                messagebox.showerror(self.tr("up_title"),
                                     self.tr("up_error", e=payload["error"]))
            return
        self.settings["last_update_check"] = date.today().isoformat()
        self._persist_settings()
        manifest = payload.get("manifest") or {}
        remote = manifest.get("version", "")
        if not remote or parse_version(remote) <= parse_version(APP_VERSION):
            if manual:
                messagebox.showinfo(self.tr("up_title"),
                                    self.tr("up_latest", version=APP_VERSION))
            return
        notes = manifest.get("notes", "")
        body = self.tr("up_available", new=remote, current=APP_VERSION)
        if notes:
            body += "\n\n" + notes
        url = manifest.get("url", "")

        one_click = is_installed_copy(APP_DIR) and url.lower().endswith(".exe")

        def open_update():
            if one_click and url.lower().startswith(("http://", "https://")):
                self._install_update(url, remote)
            elif one_click and Path(url).is_file():
                self._run_setup(Path(url))  # setup su cartella di rete: niente download
            elif url.lower().startswith(("http://", "https://")):
                webbrowser.open(url)
            else:
                # Percorso di rete: si apre la cartella, l'utente avvia il setup
                setup = Path(url)
                open_folder(setup.parent if setup.suffix else setup)

        if manual:
            # Dalla finestra Impostazioni (modale) serve una risposta esplicita
            if not url:
                messagebox.showinfo(self.tr("up_title"), body)
            elif messagebox.askyesno(self.tr("up_title"), body + "\n\n" + self.tr("up_open")):
                open_update()
            return
        # Controllo automatico all'avvio: avviso discreto che non interrompe il lavoro
        label = self.tr("up_install") if one_click else self.tr("up_get")
        self._toast(body, "info", (label, open_update) if url else None, 15000)

    def _install_update(self, url: str, version: str) -> None:
        """Scarica il setup in background; a download finito _run_setup aggiorna e riavvia."""
        self._toast(self.tr("up_downloading", version=version), "info", duration=300000)

        def work():
            try:
                target_dir = Path(tempfile.gettempdir()) / "FormazioniPZZ-update"
                self._post("update_downloaded", str(download_setup(url, target_dir, version)))
            except Exception as exc:  # noqa: BLE001 - rete assente, file non valido...
                traceback.print_exc()
                self._post("update_failed", {"error": str(exc), "url": url})

        threading.Thread(target=work, daemon=True).start()

    def _run_setup(self, setup: Path) -> None:
        """Avvia l'installer in modalita' silenziosa e chiude l'app: l'installer aspetta
        la chiusura, sostituisce i file e riapre Formazioni PZZ."""
        if self._worker_active:
            # Mai interrompere una generazione: si riprova appena finisce
            self.root.after(2000, lambda: self._run_setup(setup))
            return
        try:
            subprocess.Popen([str(setup), "--silent", "--dir", str(APP_DIR), "--no-shortcuts",
                              "--wait-pid", str(os.getpid()), "--relaunch"], close_fds=True)
        except OSError as exc:
            self._update_failed({"error": str(exc), "url": str(setup)})
            return
        self._persist_settings()
        self.root.destroy()

    def _update_failed(self, payload: dict[str, Any]) -> None:
        url = str(payload.get("url", ""))
        action = (self.tr("up_get"), lambda: webbrowser.open(url)) if url.startswith("http") else None
        self._toast(self.tr("up_failed", e=payload.get("error", "")), "error", action, 15000)
