"""Gestione reparti e gestione moduli."""

from __future__ import annotations

import os
import shutil
import traceback
from pathlib import Path
from typing import Any

from ..config import DEPARTMENTS_FILE, SUPPORTED_EXTENSIONS
from ..pdf import clear_caches
from ..system import FileDropTarget, _set_app_icon, open_folder
from ..templates import (
    TRASH_DIR_NAME,
    build_template_filename,
    department_options,
    load_departments_from_file,
    move_template_to_trash,
    parse_template,
    rename_template_key,
    save_hashes,
    suggest_template_code,
    template_key,
    validate_template_fields,
)
from ..tkcompat import BOTH, END, LEFT, RIGHT, StringVar, X, filedialog, messagebox, tk, ttk


class DialogsMixin:
    """Parte di FormazioniApp: gestione reparti e gestione moduli."""

    # ---------------------- Department Editor -----------------------------
    def open_department_editor(self):
        win = tk.Toplevel(self.root)
        win.title(self.tr("de_title"))
        win.transient(self.root)
        win.grab_set()
        win.geometry("520x480")
        win.configure(bg=self._style_colors["app_bg"])
        try:
            _set_app_icon(win)
        except Exception:
            pass
        body = tk.Frame(win, bg=self._style_colors["card_body_bg"], padx=18, pady=18)
        body.pack(fill=BOTH, expand=True, padx=14, pady=14)
        ttk.Label(body, text=self.tr("de_hint"), style="Muted.TLabel").pack(anchor="w", pady=(0, 12))
        list_wrap = tk.Frame(body, bg=self._style_colors["card_body_bg"])
        list_wrap.pack(fill=BOTH, expand=True, pady=(0, 12))
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)
        lb = tk.Listbox(list_wrap, activestyle="dotbox",
                        bg=self._style_colors["field_bg"], fg=self._style_colors["text"],
                        font=("Segoe UI", 10), selectmode="single",
                        highlightthickness=2,
                        highlightbackground=self._style_colors["border"],
                        highlightcolor=self._style_colors["focus"],
                        relief="flat", borderwidth=0)
        lb.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=lb.yview)
        sb.grid(row=0, column=1, sticky="ns")
        lb.configure(yscrollcommand=sb.set)
        current = load_departments_from_file()
        for item in current:
            lb.insert(END, item)
        row = tk.Frame(body, bg=self._style_colors["card_body_bg"])
        row.pack(fill=X, pady=(0, 12))
        tk.Label(row, text=self.tr("de_entry"), bg=self._style_colors["card_body_bg"],
                 fg=self._style_colors["text_muted"],
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 4))
        entry = ttk.Entry(row)
        entry.pack(fill=X, pady=(0, 10))
        btns = tk.Frame(row, bg=self._style_colors["card_body_bg"])
        btns.pack(fill=X)

        def add():
            val = entry.get().strip().upper()
            if not val:
                messagebox.showwarning(self.tr("de_err_empty_title"), self.tr("de_err_empty_body"))
                return
            if val in lb.get(0, END):
                messagebox.showwarning(self.tr("de_err_exists_title"), self.tr("de_err_exists_body"))
                return
            lb.insert(END, val)
            entry.delete(0, END)

        def rename():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning(self.tr("de_err_none_title"), self.tr("de_err_none_body"))
                return
            new_val = entry.get().strip().upper()
            if not new_val:
                messagebox.showwarning(self.tr("de_err_empty_title"), self.tr("de_err_empty_body"))
                return
            if new_val in lb.get(0, END):
                messagebox.showwarning(self.tr("de_err_exists_title"), self.tr("de_err_exists_body"))
                return
            lb.delete(sel[0])
            lb.insert(sel[0], new_val)

        def delete():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning(self.tr("de_err_none_title"), self.tr("de_err_none_body"))
                return
            lb.delete(sel[0])

        self._button(btns, "de_add", style="Secondary.TButton", command=add
                   ).pack(side=LEFT, padx=(0, 8))
        self._button(btns, "de_rename", style="Secondary.TButton", command=rename
                   ).pack(side=LEFT, padx=(0, 8))
        self._button(btns, "de_delete", style="Secondary.TButton", command=delete
                   ).pack(side=LEFT)

        footer = tk.Frame(body, bg=self._style_colors["card_body_bg"])
        footer.pack(fill=X, side="bottom")

        def save_and_close():
            new_list = list(dict.fromkeys(lb.get(0, END)))  # preserve order, unique
            try:
                DEPARTMENTS_FILE.parent.mkdir(parents=True, exist_ok=True)
                content = "\n".join(new_list) + ("\n" if new_list else "")
                DEPARTMENTS_FILE.write_text(content, encoding="utf-8")
            except OSError as exc:
                messagebox.showerror(self.tr("mb_error_title"), str(exc))
                return
            clear_caches()
            self.refresh_templates()
            win.destroy()

        self._button(footer, "de_save", style="Primary.TButton",
                   command=save_and_close).pack(side=RIGHT)
        self._button(footer, "de_cancel", style="Secondary.TButton",
                   command=win.destroy).pack(side=RIGHT, padx=(0, 10))

    # ---------------------- Gestione moduli ------------------------------
    def _template_form(self, parent, title: str, source_name: str,
                       initial: dict[str, Any]) -> dict[str, Any] | None:
        """Finestra modale per reparto / copie / codice. Ritorna i valori o None."""
        colors = self._style_colors
        dlg = tk.Toplevel(parent)
        dlg.title(title)
        dlg.transient(parent)
        dlg.resizable(False, False)
        dlg.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(dlg)
        except Exception:
            pass
        body = tk.Frame(dlg, bg=colors["card_body_bg"], padx=20, pady=18)
        body.pack(fill=BOTH, expand=True, padx=12, pady=12)
        body.columnconfigure(1, weight=1)

        def label(text, row):
            tk.Label(body, text=text, bg=colors["card_body_bg"], fg=colors["text"],
                     font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                     ).grid(row=row, column=0, sticky="w", padx=(0, 14), pady=(0, 10))

        tk.Label(body, text=self.tr("tm_form_file", name=source_name),
                 bg=colors["card_body_bg"], fg=colors["text_muted"],
                 font=("Segoe UI", 9), anchor="w", wraplength=420, justify="left"
                 ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))

        suffix = Path(source_name).suffix
        dept_var = StringVar(value=initial.get("department", ""))
        copies_var = StringVar(value=str(initial.get("copies", 1)))
        code_var = StringVar(value=initial.get("code", ""))
        code_touched = {"value": bool(initial.get("code"))}

        label(self.tr("tm_col_dept"), 1)
        dept_values = ["TUTTI"] + [d for d in department_options(self.templates) if d != "TUTTI"]
        dept_cb = ttk.Combobox(body, textvariable=dept_var, values=dept_values, width=32)
        dept_cb.grid(row=1, column=1, sticky="ew", pady=(0, 10))

        label(self.tr("tm_col_copies"), 2)
        copies_sp = ttk.Spinbox(body, from_=1, to=50, textvariable=copies_var, width=6)
        copies_sp.grid(row=2, column=1, sticky="w", pady=(0, 10))

        label(self.tr("tm_col_code"), 3)
        code_ent = ttk.Entry(body, textvariable=code_var, width=10)
        code_ent.grid(row=3, column=1, sticky="w", pady=(0, 10))
        code_ent.bind("<Key>", lambda _e: code_touched.__setitem__("value", True))

        label(self.tr("tm_col_label"), 4)
        label_var = StringVar(value=initial.get("label", ""))
        ttk.Entry(body, textvariable=label_var, width=34
                  ).grid(row=4, column=1, sticky="ew", pady=(0, 10))

        tk.Label(body, text=self.tr("tm_form_hint"), bg=colors["card_body_bg"],
                 fg=colors["text_muted"], font=("Segoe UI", 8), anchor="w",
                 wraplength=420, justify="left"
                 ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 10))

        preview_var = StringVar()
        tk.Label(body, textvariable=preview_var, bg=colors["card_body_bg"],
                 fg=colors["text"], font=("Consolas", 10, "bold"), anchor="w"
                 ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 14))

        def refresh_preview(*_):
            if not code_touched["value"]:
                code_var.set(suggest_template_code(dept_var.get()))
            err = validate_template_fields(dept_var.get(), copies_var.get(), code_var.get())
            if err:
                preview_var.set("⚠ " + self.tr(err))
            else:
                preview_var.set(self.tr("tm_form_preview", name=build_template_filename(
                    dept_var.get(), int(copies_var.get()), code_var.get(), suffix)))

        dept_var.trace_add("write", refresh_preview)
        copies_var.trace_add("write", refresh_preview)
        code_var.trace_add("write", refresh_preview)
        refresh_preview()

        result: dict[str, Any] = {}

        def confirm(_evt=None):
            err = validate_template_fields(dept_var.get(), copies_var.get(), code_var.get())
            if err:
                messagebox.showwarning(self.tr("de_err_empty_title"), self.tr(err), parent=dlg)
                return
            result.update(department=dept_var.get().strip().upper(),
                          copies=int(copies_var.get()),
                          code=code_var.get().strip().upper(),
                          label=" ".join(label_var.get().split()))
            dlg.destroy()

        btns = tk.Frame(body, bg=colors["card_body_bg"])
        btns.grid(row=7, column=0, columnspan=2, sticky="e")
        self._button(btns, "de_cancel", style="Secondary.TButton",
                   command=dlg.destroy).pack(side=LEFT, padx=(0, 10))
        self._button(btns, "tm_form_ok", style="Primary.TButton",
                   command=confirm).pack(side=LEFT)
        dlg.bind("<Return>", confirm)
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        dept_cb.focus_set()
        dlg.grab_set()
        parent.wait_window(dlg)
        try:
            parent.grab_set()
        except tk.TclError:
            pass
        return result or None

    def open_template_manager(self, initial_files: list[Path] | None = None):
        colors = self._style_colors
        win = tk.Toplevel(self.root)
        win.title(self.tr("tm_title"))
        win.transient(self.root)
        win.grab_set()
        win.geometry("900x560")
        win.minsize(700, 420)
        win.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(win)
        except Exception:
            pass
        body = tk.Frame(win, bg=colors["card_body_bg"], padx=18, pady=18)
        body.pack(fill=BOTH, expand=True, padx=14, pady=14)
        ttk.Label(body, text=self.tr("tm_hint"), style="Muted.TLabel",
                  wraplength=820, justify="left").pack(anchor="w", pady=(0, 12))

        toolbar = tk.Frame(body, bg=colors["card_body_bg"])
        toolbar.pack(fill=X, side="bottom")

        list_wrap = tk.Frame(body, bg=colors["card_body_bg"])
        list_wrap.pack(fill=BOTH, expand=True, pady=(0, 12))
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)
        columns = ("file", "label", "dept", "copies", "code", "status")
        tree = ttk.Treeview(list_wrap, columns=columns, show="headings", selectmode="extended")
        widths = {"file": 230, "label": 220, "dept": 150, "copies": 60, "code": 70,
                  "status": 170}
        for col in columns:
            tree.heading(col, text=self.tr(f"tm_col_{col}"))
            tree.column(col, width=widths[col],
                        anchor="center" if col in ("copies", "code") else "w")
        tree.tag_configure("invalid", foreground="#c0392b" if self.theme.get() == "light" else "#ff8a7a")
        tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=sb.set)
        rows: dict[str, Path] = {}

        def folder() -> Path:
            return Path(self.template_dir.get()).expanduser()

        def reload():
            self.refresh_templates()
            tree.delete(*tree.get_children())
            rows.clear()
            ordered = sorted(self.templates, key=lambda t: (not t.is_for_every_department,
                                                            t.department.upper(),
                                                            t.path.name.lower()))
            for tpl in ordered:
                dept = (self.tr("tm_all_depts") if tpl.is_for_every_department
                        else tpl.department.upper())
                label = self.module_settings.get("labels", {}).get(
                    template_key(tpl.path, folder()), "")
                iid = tree.insert("", END, values=(tpl.display_name, label, dept, tpl.copies,
                                                   tpl.code, "✓ " + self.tr("tm_status_ok")))
                rows[iid] = tpl.path
            for path in self.ignored:
                iid = tree.insert("", END, values=(path.name, "", "—", "—", "—",
                                                   "⚠ " + self.tr("tm_status_invalid")),
                                  tags=("invalid",))
                rows[iid] = path

        def selected_paths() -> list[Path]:
            return [rows[i] for i in tree.selection() if i in rows]

        def require_one() -> Path | None:
            paths = selected_paths()
            if len(paths) != 1:
                messagebox.showwarning(self.tr("de_err_none_title"),
                                       self.tr("tm_err_select_one"), parent=win)
                return None
            return paths[0]

        def initial_for(path: Path) -> dict[str, Any]:
            label = self.module_settings.get("labels", {}).get(template_key(path, folder()), "")
            tpl = parse_template(path)
            if tpl:
                return {"department": tpl.department.upper(), "copies": tpl.copies,
                        "code": tpl.code, "label": label}
            return {"department": self.department.get() or "", "copies": 1, "code": "",
                    "label": label}

        def add_files():
            sources = filedialog.askopenfilenames(
                parent=win, title=self.tr("tm_add"),
                filetypes=[(self.tr("tm_filetypes"), "*.docx *.doc *.xlsx *.xls *.pdf")])
            add_paths(sources)

        def add_paths(sources):
            if not sources:
                return
            dest_dir = folder()
            try:
                dest_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                messagebox.showerror(self.tr("tm_title"), str(exc), parent=win)
                return
            added = 0
            for src in map(Path, sources):
                if src.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue
                initial = initial_for(src)
                while True:
                    values = self._template_form(win, self.tr("tm_form_title_add"),
                                                 src.name, initial)
                    if values is None:
                        break
                    target = dest_dir / build_template_filename(
                        values["department"], values["copies"], values["code"], src.suffix)
                    if target.exists():
                        messagebox.showwarning(self.tr("de_err_exists_title"),
                                               self.tr("tm_err_exists", name=target.name),
                                               parent=win)
                        initial = values
                        continue
                    try:
                        shutil.copy2(src, target)
                        added += 1
                        self.set_template_label(target, values["label"])
                    except OSError as exc:
                        messagebox.showerror(self.tr("tm_title"), str(exc), parent=win)
                    break
            if added:
                self._save_module_settings(parent=win)
                reload()
                self.status.set(self.tr("tm_added", n=added))

        def edit_selected(_evt=None):
            path = require_one()
            if path is None:
                return
            initial = initial_for(path)
            while True:
                values = self._template_form(win, self.tr("tm_form_title_edit"),
                                             path.name, initial)
                if values is None:
                    return
                target = path.with_name(build_template_filename(
                    values["department"], values["copies"], values["code"], path.suffix))
                if target == path:
                    if values["label"] != initial.get("label", ""):
                        self.set_template_label(path, values["label"])
                        self._save_module_settings(parent=win)
                        reload()
                    return
                if target.exists():
                    messagebox.showwarning(self.tr("de_err_exists_title"),
                                           self.tr("tm_err_exists", name=target.name),
                                           parent=win)
                    initial = values
                    continue
                try:
                    path.rename(target)
                except OSError as exc:
                    messagebox.showerror(self.tr("tm_title"), str(exc), parent=win)
                    return
                if str(path) in self.saved_hashes:
                    self.saved_hashes[str(target)] = self.saved_hashes.pop(str(path))
                    save_hashes(self.saved_hashes)
                if path in self.template_inclusion:
                    self.template_inclusion[target] = self.template_inclusion.pop(path)
                # Nome in app e posizione negli ordini seguono il file rinominato
                old_key, new_key = template_key(path, folder()), template_key(target, folder())
                rename_template_key(self.module_settings, old_key, new_key)
                self.doc_order = [new_key if k == old_key else k for k in self.doc_order]
                self.set_template_label(target, values["label"])
                self._save_module_settings(parent=win)
                reload()
                self.status.set(self.tr("tm_renamed", name=target.name))
                return

        def delete_selected():
            paths = selected_paths()
            if not paths:
                messagebox.showwarning(self.tr("de_err_none_title"),
                                       self.tr("tm_err_select_any"), parent=win)
                return
            names = "\n".join("• " + p.name for p in paths[:10])
            if len(paths) > 10:
                names += "\n…"
            if not messagebox.askyesno(self.tr("tm_delete"),
                                       self.tr("tm_confirm_delete", n=len(paths), names=names,
                                               trash=TRASH_DIR_NAME), parent=win):
                return
            for path in paths:
                try:
                    move_template_to_trash(path, folder())
                except OSError as exc:
                    messagebox.showerror(self.tr("tm_title"), str(exc), parent=win)
            reload()
            self.status.set(self.tr("tm_deleted", n=len(paths)))

        def open_selected():
            path = require_one()
            if path is not None:
                open_folder(path)

        for text_key, cmd, style in (
            ("tm_add", add_files, "Accent.TButton"),
            ("tm_edit", edit_selected, "Secondary.TButton"),
            ("tm_open", open_selected, "Secondary.TButton"),
            ("tm_delete", delete_selected, "Secondary.TButton"),
        ):
            self._button(toolbar, text_key, style=style, command=cmd
                       ).pack(side=LEFT, padx=(0, 8))
        self._button(toolbar, "tm_close", style="Primary.TButton",
                   command=win.destroy).pack(side=RIGHT)
        self._button(toolbar, "tm_open_folder", style="Secondary.TButton",
                   command=lambda: open_folder(folder())).pack(side=RIGHT, padx=(0, 8))

        tree.bind("<Double-1>", edit_selected)
        tree.bind("<Return>", edit_selected)
        tree.bind("<Delete>", lambda _e: delete_selected())
        reload()
        if os.name == "nt":
            try:
                win._drop_target = FileDropTarget(win, lambda paths, _x, _y: add_paths(paths))
            except Exception:  # noqa: BLE001
                traceback.print_exc()
        if initial_files:
            win.after(150, lambda: add_paths(initial_files))
