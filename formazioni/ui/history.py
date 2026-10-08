"""Storico dei dossier generati."""

from __future__ import annotations

import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

from ..config import HISTORY_FILE, HISTORY_LIMIT
from ..system import _set_app_icon, open_folder
from ..templates import department_options
from ..tkcompat import BOTH, END, LEFT, RIGHT, StringVar, X, messagebox, tk, ttk


class HistoryMixin:
    """Parte di FormazioniApp: storico dei dossier generati."""

    # ---------------------- History ---------------------------------------
    @staticmethod
    def _load_history() -> list[dict[str, Any]]:
        if not HISTORY_FILE.exists():
            return []
        try:
            history = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return []
        return [h for h in history if isinstance(h, dict)] if isinstance(history, list) else []

    def _write_history(self, history: list[dict[str, Any]]) -> None:
        try:
            HISTORY_FILE.write_text(
                json.dumps(history[:HISTORY_LIMIT], ensure_ascii=False, indent=2),
                encoding="utf-8")
        except OSError:
            # La cronologia e' accessoria: un errore qui non deve far fallire il dossier
            traceback.print_exc()

    def _save_history(self, output_path: Path, name: str, department: str, total: int,
                      entry_date: str = "") -> None:
        with self._history_lock:
            history = self._load_history()
            history.insert(0, {
                "path": str(output_path),
                "name": name,
                "entry_date": entry_date,
                "department": department,
                "documents": total,
                "ts": datetime.now().isoformat(timespec="seconds"),
            })
            self._write_history(history)

    def open_history(self):
        colors = self._style_colors
        win = tk.Toplevel(self.root)
        win.title(self.tr("hi_title"))
        win.transient(self.root)
        win.geometry("980x560")
        win.minsize(720, 400)
        win.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(win)
        except Exception:
            pass
        body = tk.Frame(win, bg=colors["card_body_bg"], padx=18, pady=18)
        body.pack(fill=BOTH, expand=True, padx=14, pady=14)

        top = tk.Frame(body, bg=colors["card_body_bg"])
        top.pack(fill=X, pady=(0, 12))
        tk.Label(top, text=self.tr("hi_search"), bg=colors["card_body_bg"], fg=colors["text"],
                 font=("Segoe UI Semibold", 9, "bold")).pack(side=LEFT, padx=(0, 10))
        query = StringVar()
        search = ttk.Entry(top, textvariable=query)
        search.pack(side=LEFT, fill=X, expand=True)
        count_var = StringVar()
        ttk.Label(top, textvariable=count_var, style="Count.TLabel").pack(side=RIGHT, padx=(12, 0))

        toolbar = tk.Frame(body, bg=colors["card_body_bg"])
        toolbar.pack(fill=X, side="bottom")

        list_wrap = tk.Frame(body, bg=colors["card_body_bg"])
        list_wrap.pack(fill=BOTH, expand=True, pady=(0, 12))
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)
        columns = ("ts", "name", "entry", "dept", "docs", "file")
        tree = ttk.Treeview(list_wrap, columns=columns, show="headings", selectmode="browse")
        widths = {"ts": 130, "name": 180, "entry": 100, "dept": 190, "docs": 60, "file": 300}
        for col in columns:
            tree.heading(col, text=self.tr(f"hi_col_{col}"))
            tree.column(col, width=widths[col], anchor="center" if col == "docs" else "w",
                        stretch=col in ("name", "dept", "file"))
        tree.tag_configure("missing", foreground=colors["text_muted"])
        tree.tag_configure("even", background=colors["row_even"])
        tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=sb.set)
        rows: dict[str, dict[str, Any]] = {}

        def fmt_ts(value: str) -> str:
            try:
                return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
            except (TypeError, ValueError):
                return str(value or "")

        def reload(*_):
            tree.delete(*tree.get_children())
            rows.clear()
            needle = query.get().strip().lower()
            history = self._load_history()
            shown = 0
            for entry in history:
                haystack = " ".join(str(entry.get(k, "")) for k in
                                    ("name", "department", "entry_date", "path")).lower()
                if needle and needle not in haystack:
                    continue
                path = Path(str(entry.get("path", "")))
                exists = path.exists()
                tags = [] if exists else ["missing"]
                if shown % 2:
                    tags.append("even")
                iid = tree.insert("", END, tags=tuple(tags), values=(
                    fmt_ts(entry.get("ts", "")), entry.get("name", ""),
                    entry.get("entry_date", "") or "—", entry.get("department", ""),
                    entry.get("documents", ""),
                    path.name if exists else f"⚠ {path.name} · {self.tr('hi_missing')}",
                ))
                rows[iid] = entry
                shown += 1
            count_var.set(self.tr("hi_count", shown=shown, total=len(history)))

        def current() -> dict[str, Any] | None:
            sel = tree.selection()
            if not sel or sel[0] not in rows:
                messagebox.showwarning(self.tr("de_err_none_title"),
                                       self.tr("hi_select"), parent=win)
                return None
            return rows[sel[0]]

        def open_pdf(_evt=None):
            entry = current()
            if entry is None:
                return
            path = Path(str(entry.get("path", "")))
            if not path.exists():
                messagebox.showwarning(self.tr("hi_title"), self.tr("hi_missing_body"), parent=win)
                return
            open_folder(path)

        def open_dir():
            entry = current()
            if entry is None:
                return
            folder = Path(str(entry.get("path", ""))).parent
            open_folder(folder if folder.exists() else Path(self.output_dir.get()))

        def reuse():
            entry = current()
            if entry is None:
                return
            self._apply_history_entry(entry)
            win.destroy()

        def remove():
            entry = current()
            if entry is None:
                return
            if not messagebox.askyesno(self.tr("hi_title"),
                                       self.tr("hi_confirm_remove", name=entry.get("name", "")),
                                       parent=win):
                return
            with self._history_lock:
                history = [h for h in self._load_history() if h != entry]
                self._write_history(history)
            reload()

        for text_key, cmd, style in (
            ("hi_open_pdf", open_pdf, "Accent.TButton"),
            ("hi_open_dir", open_dir, "Secondary.TButton"),
            ("hi_reuse", reuse, "Secondary.TButton"),
            ("hi_remove", remove, "Secondary.TButton"),
        ):
            self._button(toolbar, text_key, style=style, command=cmd
                       ).pack(side=LEFT, padx=(0, 8))
        self._button(toolbar, "tm_close", style="Primary.TButton",
                   command=win.destroy).pack(side=RIGHT)

        query.trace_add("write", reload)
        tree.bind("<Double-1>", open_pdf)
        tree.bind("<Return>", open_pdf)
        tree.bind("<Delete>", lambda _e: remove())
        win.bind("<Escape>", lambda _e: win.destroy())
        reload()
        search.focus_set()

    def _apply_history_entry(self, entry: dict[str, Any]) -> None:
        """Ricompila il modulo con i dati di un dossier dello storico."""
        self.employee_name.set(str(entry.get("name", "")))
        parsed = self._parse_date_str(str(entry.get("entry_date", "")))
        if parsed is not None:
            self.date_picker.set_date(parsed)
        options = department_options(self.templates)
        wanted = [d.strip().upper() for d in str(entry.get("department", "")).split("+")
                  if d.strip()]
        known = [d for d in wanted if d in options]
        if len(known) > 1:
            self.multi_dept_mode.set(True)
            self._toggle_multi_dept()
            for dept in known:
                if dept in self.multi_dept_values:
                    self.multi_dept_values[dept].set(True)
        elif known:
            if self.multi_dept_mode.get():
                self.multi_dept_mode.set(False)
                self._toggle_multi_dept()
            self.department.set(known[0])
        self.update_document_list()
        if len(known) != len(wanted):
            missing = ", ".join(d for d in wanted if d not in known)
            messagebox.showwarning(self.tr("hi_title"), self.tr("hi_dept_missing", d=missing))
        self.status.set(self.tr("hi_reused", name=entry.get("name", "")))
        self._toast(self.tr("hi_reused", name=entry.get("name", "")))
