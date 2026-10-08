"""Ordine dei documenti nel dossier, ordini salvati come modelli e nomi dei moduli
visibili solo nell'app."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..system import _set_app_icon
from ..templates import (
    TemplateFile,
    apply_order,
    department_options,
    order_for_department,
    save_module_settings,
    template_key,
)
from ..tkcompat import BOTH, LEFT, StringVar, X, messagebox, tk, ttk


class OrderingMixin:
    """Parte di FormazioniApp: ordine dei documenti, ordini salvati e nomi in app."""

    # ---------------------- Dati ------------------------------------------
    def _template_folder(self) -> Path:
        return Path(self.template_dir.get()).expanduser()

    def _template_key(self, template: TemplateFile) -> str:
        return template_key(template.path, self._template_folder())

    def _template_label(self, template: TemplateFile) -> str:
        """Nome mostrato nell'app: quello scelto dall'utente, altrimenti il nome del file."""
        label = self.module_settings.get("labels", {}).get(self._template_key(template))
        return label or template.path.name

    def _save_module_settings(self, parent=None) -> bool:
        try:
            save_module_settings(self._template_folder(), self.module_settings)
            return True
        except OSError as exc:
            messagebox.showerror(self.tr("ord_title"), self.tr("ord_err_save", e=exc),
                                 parent=parent or self.root)
            return False

    def _full_order(self) -> list[str]:
        """Chiavi di tutti i moduli nell'ordine attuale."""
        folder = self._template_folder()
        return [template_key(t.path, folder)
                for t in apply_order(self.templates, self.doc_order, folder)]

    # ---------------------- Widget ----------------------------------------
    def _build_order_row(self, parent, row: int) -> None:
        colors = self._style_colors
        frame = tk.Frame(parent, bg=colors["card_body_bg"])
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        frame.columnconfigure(1, weight=1)
        tk.Label(frame, text=self.tr("ord_label"), bg=colors["card_body_bg"], fg=colors["text"],
                 font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                 ).grid(row=0, column=0, sticky="w", padx=(0, 14))
        self._order_choice = StringVar()
        combo = ttk.Combobox(frame, state="readonly", textvariable=self._order_choice,
                             font=("Segoe UI", 10), height=12, width=12)
        combo.grid(row=0, column=1, sticky="ew")
        combo.bind("<<ComboboxSelected>>", lambda _e: self._on_order_selected())
        self._order_combo = combo
        self._register_local_wheel(combo)
        self._add_tooltip(combo, lambda: self.tr("tt_order"))
        # Pulsanti sotto il menu: affiancati lo stringerebbero troppo nella card
        buttons = tk.Frame(frame, bg=colors["card_body_bg"])
        buttons.grid(row=1, column=1, sticky="w", pady=(6, 0))
        save_btn = self._button(buttons, "ord_save", style="Secondary.TButton",
                                command=self.save_current_order)
        save_btn.pack(side=LEFT)
        self._add_tooltip(save_btn, lambda: self.tr("tt_order_save"))
        self._order_delete_btn = self._button(buttons, "ord_delete", style="Secondary.TButton",
                                              command=self.delete_current_order)
        self._order_delete_btn.pack(side=LEFT, padx=(6, 0))
        self._order_hint = tk.Label(frame, text="", bg=colors["card_body_bg"],
                                    fg=colors["text_muted"], font=("Segoe UI", 8),
                                    anchor="w", justify="left")
        self._order_hint.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        self._fluid_wrap(self._order_hint)
        self._refresh_order_choices()

    def _build_order_tools(self, parent, row: int) -> None:
        colors = self._style_colors
        frame = tk.Frame(parent, bg=colors["card_body_bg"])
        frame.grid(row=row, column=0, sticky="ew", pady=(2, 8))
        buttons = tk.Frame(frame, bg=colors["card_body_bg"])
        buttons.pack(side="top", anchor="w")
        for key, cmd, tip, width in (
                ("ord_up", lambda: self._move_selected(-1), "tt_order_move", 3),
                ("ord_down", lambda: self._move_selected(1), "tt_order_move", 3),
                ("ord_rename", self.rename_selected_template, "tt_label", None)):
            options = {"width": width} if width else {}
            btn = self._button(buttons, key, style="Secondary.TButton", command=cmd, **options)
            btn.pack(side=LEFT, padx=(0, 6))
            self._add_tooltip(btn, lambda k=tip: self.tr(k))
        hint = tk.Label(frame, text=self.tr("ord_tools_hint"), bg=colors["card_body_bg"],
                        fg=colors["text_muted"], font=("Segoe UI", 8),
                        anchor="w", justify="left")
        hint.pack(side="top", fill=X, pady=(6, 0))
        self._fluid_wrap(hint)

    def _refresh_order_choices(self) -> None:
        combo = getattr(self, "_order_combo", None)
        if combo is None or not combo.winfo_exists():
            return
        orders = self.module_settings.get("orders", {})
        name = self.active_order.get()
        if name and name not in orders:
            self.active_order.set("")
            name = ""
        combo["values"] = [self.tr("ord_default")] + sorted(orders, key=str.lower)
        shown = name or self.tr("ord_default")
        if getattr(self, "_order_changed", False):
            shown += self.tr("ord_changed_suffix")
        self._order_choice.set(shown)
        dept = orders.get(name, {}).get("department") if name else ""
        if dept:
            hint = self.tr("ord_hint_linked", d=dept)
        elif name:
            hint = self.tr("ord_hint_manual")
        else:
            hint = self.tr("ord_hint_default")
        self._order_hint.configure(text=hint)
        self._order_delete_btn.state(["!disabled"] if name else ["disabled"])

    def _mark_order_changed(self) -> None:
        if not getattr(self, "_order_changed", False):
            self._order_changed = True
            self._refresh_order_choices()

    # ---------------------- Applicare un ordine ---------------------------
    def apply_saved_order(self, name: str | None) -> None:
        """Applica un ordine salvato (ordine e documenti esclusi), o quello predefinito."""
        entry = self.module_settings.get("orders", {}).get(name or "")
        if entry is None:
            self.doc_order = []
            excluded: set[str] = set()
            self.active_order.set("")
        else:
            self.doc_order = list(entry.get("order", []))
            excluded = set(entry.get("excluded", []))
            self.active_order.set(name or "")
        for template in self.templates:
            self.template_inclusion[template.path] = self._template_key(template) not in excluded
        self._order_changed = False
        self._refresh_order_choices()

    def _on_order_selected(self) -> None:
        choice = self._order_choice.get()
        orders = self.module_settings.get("orders", {})
        self.apply_saved_order(choice if choice in orders else None)
        self.update_document_list()

    def _apply_department_order(self) -> None:
        """Scelto un reparto, si applica l'ordine collegato a quel reparto; se era
        applicato l'ordine di un altro reparto si torna a quello predefinito."""
        if self.multi_dept_mode.get():
            return
        dept = self.department.get().strip().upper()
        linked = order_for_department(self.module_settings, dept)
        active = self.active_order.get()
        if linked:
            if linked != active:
                self.apply_saved_order(linked)
            return
        active_dept = self.module_settings.get("orders", {}).get(active, {}).get("department")
        if active_dept and active_dept != dept:
            self.apply_saved_order(None)

    def _on_department_changed(self) -> None:
        self._apply_department_order()
        self.update_document_list()

    # ---------------------- Riordinare la lista ---------------------------
    def _selected_tree_path(self) -> Path | None:
        tree = getattr(self, "tree", None)
        if tree is None or not tree.winfo_exists():
            return None
        selection = tree.selection()
        return self._row_path.get(selection[0]) if selection else None

    def _commit_visible_order(self, visible: list[Path]) -> None:
        """La lista mostra solo i moduli dei reparti scelti: questi vanno in testa
        nell'ordine nuovo, gli altri mantengono la loro posizione relativa."""
        folder = self._template_folder()
        keys = [template_key(p, folder) for p in visible]
        shown = set(keys)
        self.doc_order = keys + [k for k in self._full_order() if k not in shown]
        self._mark_order_changed()
        self.update_document_list()

    def _move_selected(self, delta: int) -> None:
        path = self._selected_tree_path()
        if path is None:
            messagebox.showinfo(self.tr("ord_title"), self.tr("ord_select_row"))
            return
        paths = [self._row_path[i] for i in self.tree.get_children()]
        index = paths.index(path)
        target = index + delta
        if not 0 <= target < len(paths):
            return
        paths.insert(target, paths.pop(index))
        self._commit_visible_order(paths)

    def _bind_tree_ordering(self) -> None:
        tree = self.tree
        self._drag: dict[str, Any] | None = None

        def press(evt):
            item = tree.identify_row(evt.y)
            if item and tree.identify_column(evt.x) != "#1" \
                    and tree.identify("region", evt.x, evt.y) == "cell":
                self._drag = {"item": item, "moved": False}
            else:
                self._drag = None

        def motion(evt):
            drag = self._drag
            if not drag:
                return
            target = tree.identify_row(evt.y)
            if target and target != drag["item"]:
                tree.move(drag["item"], "", tree.index(target))
                drag["moved"] = True

        def release(_evt):
            drag, self._drag = self._drag, None
            if drag and drag["moved"]:
                tree.selection_set(drag["item"])
                self._commit_visible_order([self._row_path[i] for i in tree.get_children()])

        def double(evt):
            if tree.identify_row(evt.y) and tree.identify_column(evt.x) != "#1":
                self.rename_selected_template()

        tree.bind("<ButtonPress-1>", press, add="+")
        tree.bind("<B1-Motion>", motion)
        tree.bind("<ButtonRelease-1>", release)
        tree.bind("<Double-1>", double)
        tree.bind("<Alt-Up>", lambda _e: (self._move_selected(-1), "break")[1])
        tree.bind("<Alt-Down>", lambda _e: (self._move_selected(1), "break")[1])
        tree.bind("<F2>", lambda _e: self.rename_selected_template())

    # ---------------------- Dialoghi --------------------------------------
    def _small_dialog(self, parent, title: str):
        colors = self._style_colors
        dlg = tk.Toplevel(parent)
        dlg.title(title)
        dlg.transient(parent)
        dlg.resizable(False, False)
        dlg.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(dlg)
        except Exception:  # noqa: BLE001
            pass
        body = tk.Frame(dlg, bg=colors["card_body_bg"], padx=20, pady=18)
        body.pack(fill=BOTH, expand=True, padx=12, pady=12)
        body.columnconfigure(1, weight=1)
        return dlg, body

    def _run_dialog(self, parent, dlg, body, row: int, confirm, focus) -> None:
        btns = tk.Frame(body, bg=self._style_colors["card_body_bg"])
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(4, 0))
        self._button(btns, "de_cancel", style="Secondary.TButton",
                     command=dlg.destroy).pack(side=LEFT, padx=(0, 10))
        self._button(btns, "tm_form_ok", style="Primary.TButton",
                     command=confirm).pack(side=LEFT)
        dlg.bind("<Return>", lambda _e: confirm())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        focus.focus_set()
        dlg.grab_set()
        parent.wait_window(dlg)
        try:
            if parent is not self.root:
                parent.grab_set()
        except tk.TclError:
            pass

    def _muted(self, body, text: str, row: int) -> None:
        colors = self._style_colors
        tk.Label(body, text=text, bg=colors["card_body_bg"], fg=colors["text_muted"],
                 font=("Segoe UI", 9), anchor="w", wraplength=440, justify="left"
                 ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 12))

    def _field_title(self, body, text: str, row: int) -> None:
        colors = self._style_colors
        tk.Label(body, text=text, bg=colors["card_body_bg"], fg=colors["text"],
                 font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                 ).grid(row=row, column=0, sticky="w", padx=(0, 14), pady=(0, 12))

    def ask_template_label(self, parent, file_name: str, current: str) -> str | None:
        """Chiede il nome in app di un modulo; None se annullato, "" per toglierlo."""
        dlg, body = self._small_dialog(parent, self.tr("lbl_title"))
        self._muted(body, self.tr("lbl_prompt", file=file_name), 0)
        self._field_title(body, self.tr("lbl_title"), 1)
        var = StringVar(value=current)
        entry = ttk.Entry(body, textvariable=var, width=40)
        entry.grid(row=1, column=1, sticky="ew", pady=(0, 12))
        entry.select_range(0, "end")
        result: dict[str, str] = {}

        def confirm():
            result["value"] = " ".join(var.get().split())
            dlg.destroy()

        self._run_dialog(parent, dlg, body, 2, confirm, entry)
        return result.get("value")

    def set_template_label(self, path: Path, label: str) -> None:
        labels = self.module_settings.setdefault("labels", {})
        key = template_key(path, self._template_folder())
        if label:
            labels[key] = label
        else:
            labels.pop(key, None)

    def rename_selected_template(self) -> None:
        path = self._selected_tree_path()
        if path is None:
            messagebox.showinfo(self.tr("lbl_title"), self.tr("ord_select_row"))
            return
        current = self.module_settings.get("labels", {}).get(
            template_key(path, self._template_folder()), "")
        value = self.ask_template_label(self.root, path.name, current)
        if value is None or value == current:
            return
        self.set_template_label(path, value)
        if self._save_module_settings():
            self.status.set(self.tr("lbl_saved"))
        self.update_document_list()

    def save_current_order(self) -> None:
        if not self.tree.get_children():
            messagebox.showinfo(self.tr("ord_title"), self.tr("ord_empty"))
            return
        orders = self.module_settings.setdefault("orders", {})
        active = self.active_order.get()
        no_dept = self.tr("ord_no_dept")
        depts = self._current_departments()
        initial_dept = (orders.get(active, {}).get("department")
                        or (depts[0] if len(depts) == 1 else ""))

        dlg, body = self._small_dialog(self.root, self.tr("ord_save_title"))
        self._muted(body, self.tr("ord_save_hint"), 0)
        self._field_title(body, self.tr("ord_save_name"), 1)
        name_var = StringVar(value=active or (depts[0].title() if len(depts) == 1 else ""))
        name_cb = ttk.Combobox(body, textvariable=name_var, width=36,
                               values=sorted(orders, key=str.lower))
        name_cb.grid(row=1, column=1, sticky="ew", pady=(0, 12))
        self._field_title(body, self.tr("ord_save_dept"), 2)
        dept_var = StringVar(value=initial_dept or no_dept)
        ttk.Combobox(body, textvariable=dept_var, state="readonly", width=36,
                     values=[no_dept] + department_options(self.templates)
                     ).grid(row=2, column=1, sticky="ew", pady=(0, 12))
        result: dict[str, str] = {}

        def confirm():
            name = " ".join(name_var.get().split())
            if not name:
                messagebox.showwarning(self.tr("ord_title"), self.tr("ord_err_name"), parent=dlg)
                return
            if name in orders and name != active and not messagebox.askyesno(
                    self.tr("ord_title"), self.tr("ord_overwrite", name=name), parent=dlg):
                return
            dept = dept_var.get()
            result.update(name=name, department="" if dept == no_dept else dept)
            dlg.destroy()

        self._run_dialog(self.root, dlg, body, 3, confirm, name_cb)
        if not result:
            return
        name, dept = result["name"], result["department"]
        visible = [self._row_path[i] for i in self.tree.get_children()]
        folder = self._template_folder()
        if dept:
            # Un reparto ha un solo ordine collegato
            for other in orders.values():
                if other.get("department") == dept:
                    other["department"] = ""
        orders[name] = {
            "order": [template_key(p, folder) for p in visible],
            "excluded": [template_key(p, folder) for p in visible
                         if not self.template_inclusion.get(p, True)],
            "department": dept,
        }
        if not self._save_module_settings():
            return
        self.active_order.set(name)
        self._order_changed = False
        self._refresh_order_choices()
        self.status.set(self.tr("ord_saved", name=name))

    def delete_current_order(self) -> None:
        name = self.active_order.get()
        orders = self.module_settings.get("orders", {})
        if not name or name not in orders:
            return
        if not messagebox.askyesno(self.tr("ord_title"),
                                   self.tr("ord_confirm_delete", name=name)):
            return
        removed = orders.pop(name)
        if not self._save_module_settings():
            orders[name] = removed
            return
        self.apply_saved_order(None)
        self.update_document_list()
        self.status.set(self.tr("ord_deleted", name=name))
