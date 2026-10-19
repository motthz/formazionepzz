"""Elenco batch: aggiunta persone, import CSV/Excel, generazione multipla."""

from __future__ import annotations

import threading
from pathlib import Path
from typing import Any

from ..batch_excel import write_batch_template
from ..batch_import import (
    BatchPlan,
    UnsupportedBatchFile,
    coerce_date_value,
    normalize_column_header,
    parse_batch_file,
    parse_date,
    people_from_records,
    plan_batch,
)
from ..config import ALL_DEPARTMENT_NAMES
from ..dossier import remember_template_hashes
from ..icons import strip_leading_symbol
from ..pdf import build_pdfs
from ..system import open_folder
from ..templates import department_options
from ..tkcompat import END, StringVar, filedialog, messagebox


class BatchMixin:
    """Parte di FormazioniApp: elenco batch: aggiunta persone, import CSV/Excel, generazione multipla."""

    def _inline_batch_count(self) -> StringVar:
        self._inline_batch_count_var = var = StringVar()

        def _fmt(n):
            if n == 0:
                return self.tr("bat_count_none")
            return self.tr("bat_count_one") if n == 1 else self.tr("bat_count", n=n)

        self._inline_batch_count_fmt = _fmt
        self._refresh_inline_batch_count()
        return var

    def _refresh_inline_batch_count(self):
        fmt = getattr(self, "_inline_batch_count_fmt", None)
        var = getattr(self, "_inline_batch_count_var", None)
        if var is None or fmt is None:
            return
        n = len(getattr(self, "_inline_batch_rows", []))
        var.set(fmt(n))
        empty = getattr(self, "_inline_batch_empty", None)
        if empty is not None and empty.winfo_exists():
            if n:
                empty.place_forget()
            else:
                empty.place(relx=0.5, rely=0.55, anchor="center")
        self._refresh_batch_scope_hint()
        run_btn = getattr(self, "_inline_batch_run_btn", None)
        # Dopo un cambio di tema/lingua il riferimento puo' puntare al pulsante distrutto
        if run_btn is not None and run_btn.winfo_exists():
            text = self.tr("bat_run_all", n=n)
            if run_btn.cget("image"):
                text = " " + strip_leading_symbol(text)
            run_btn.configure(text=text)

    def _batch_common_department(self) -> str | None:
        """Reparto comune a tutte le persone in elenco, oppure None se sono diversi."""
        depts = {(r.get("Reparto") or "").strip().upper()
                 for r in getattr(self, "_inline_batch_rows", [])}
        return next(iter(depts)) if len(depts) == 1 else None

    def _refresh_batch_scope_hint(self):
        """Le spunte dei documenti valgono per tutte le persone dell'elenco: lo si dice
        sotto la lista dei documenti, con il reparto se e' lo stesso per tutti."""
        label = getattr(self, "_batch_scope_label", None)
        if label is None or not label.winfo_exists():
            return
        n = len(getattr(self, "_inline_batch_rows", []))
        if not n:
            label.pack_forget()
            return
        dept = self._batch_common_department()
        if dept:
            text = (self.tr("bat_scope_same_one", d=dept) if n == 1
                    else self.tr("bat_scope_same", n=n, d=dept))
        else:
            text = self.tr("bat_scope_mixed", n=n)
        label.configure(text=text)
        if not label.winfo_manager():
            label.pack(fill="x", side="top", pady=(0, 6), before=self._progressbar)

    def _show_batch_department(self):
        """Se tutte le persone sono di un solo reparto, la lista mostra quel reparto:
        cosi' si spuntano proprio i documenti che riceveranno."""
        dept = self._batch_common_department()
        if not dept or "+" in dept or self.multi_dept_mode.get():
            return
        if dept in tuple(self.department_combo["values"]) and self.department.get().upper() != dept:
            self.department.set(dept)
            self._on_department_changed()

    def _current_dept_for_batch(self) -> str:
        depts = self._current_departments()
        return depts[0] if len(depts) == 1 else "+".join(depts)

    def _inline_batch_add_current(self):
        name = self.employee_name.get().strip()
        if not name:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_name"))
            return
        entry_date = self.date_picker.get_string()
        if not entry_date:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_date"))
            return
        depts = self._current_departments()
        if not depts:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_dept"))
            return
        if not self._current_selected_templates():
            messagebox.showwarning(self.tr("mb_no_docs_title"), self.tr("mb_no_docs_body"))
            return
        dept_str = "+".join(depts)
        common = self._batch_common_department()
        if self._inline_batch_rows and common != dept_str:
            # Le spunte sono uniche per tutto l'elenco: reparti diversi ricevono
            # documenti diversi, quindi si chiede conferma.
            if not messagebox.askyesno(
                    self.tr("bat_mixed_title"),
                    self.tr("bat_mixed_body", name=name, new=dept_str,
                            d=common or self.tr("bat_mixed_many"))):
                return
        row = {"Nome": name, "Data": entry_date, "Reparto": dept_str}
        self._inline_batch_rows.append(row)
        tree = self._inline_batch_tree
        tree.insert("", END, values=(row["Nome"], row["Data"], row["Reparto"]))
        self._refresh_inline_batch_count()
        self.employee_name.set("")

    def _inline_batch_remove_selected(self):
        tree = self._inline_batch_tree
        sel = tree.selection()
        if not sel:
            return
        # Map tree children index to rows (insert order = _inline_batch_rows)
        children = tree.get_children()
        idxs_to_remove = set()
        for s in sel:
            try:
                idxs_to_remove.add(children.index(s))
            except ValueError:
                continue
        new_rows = [r for i, r in enumerate(self._inline_batch_rows)
                    if i not in idxs_to_remove]
        self._inline_batch_rows = new_rows
        for s in sel:
            tree.delete(s)
        self._refresh_inline_batch_count()

    def _inline_batch_clear(self):
        self._inline_batch_rows.clear()
        tree = self._inline_batch_tree
        for c in tree.get_children():
            tree.delete(c)
        self._refresh_inline_batch_count()

    def _inline_batch_import_file(self, p: Path | None = None):
        if p is None:
            p = filedialog.askopenfilename(
                title=self.tr("bat_choose"),
                initialdir=self.template_dir.get(),
                filetypes=[("CSV / Excel", "*.csv *.xlsx"), ("All", "*.*")],
            )
        if not p:
            return
        before = len(self._inline_batch_rows)
        try:
            loaded = self._parse_batch_file(Path(p))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(self.tr("mb_error_title"), str(exc))
            return
        tree = self._inline_batch_tree
        people, ignored = people_from_records(loaded)
        for row in people:
            self._inline_batch_rows.append(row)
            tree.insert("", END, values=(row["Nome"], row["Data"], row["Reparto"]))
        self._refresh_inline_batch_count()
        self._show_batch_department()
        self.warm_up_office()
        added = len(self._inline_batch_rows) - before
        message = (self.tr("dd_imported_one", name=Path(p).name) if added == 1
                   else self.tr("dd_imported", n=added, name=Path(p).name))
        if ignored:
            # Righe senza nome o data: non entrano nell'elenco, ma lo si dice
            message += " " + self.tr("dd_import_ignored", n=ignored)
        self._toast(message, "warning" if ignored else "success")

    def _pdf_options(self) -> dict[str, Any]:
        return {
            "watermark": str(self.settings.get("pdf_watermark") or ""),
            "protect": bool(self.settings.get("pdf_protect")),
        }

    def _inline_batch_run(self):
        rows = list(getattr(self, "_inline_batch_rows", []))
        if not rows:
            messagebox.showwarning(self.tr("bat_summary_title"), self.tr("bat_no_rows"))
            return
        if self._worker_active:
            return
        self._worker_active = True
        self._post("progress_ready_label", None)
        auto_open_batch = bool(self.auto_open.get())
        lang_snap = dict(self.language) if self.language else {}

        def s_tr(key: str, **kw) -> str:
            raw = lang_snap.get(key, key)
            try:
                return raw.format(**kw) if kw else raw
            except Exception:
                return raw

        tpl_snap = list(self.templates)
        plan_options = {
            "folder": self._template_folder(),
            "output_dir": Path(self.output_dir.get()).expanduser(),
            "doc_order": list(self.doc_order),
            "inclusion": dict(self.template_inclusion),
            "module_settings": {"orders": {k: dict(v) for k, v in
                                           self.module_settings.get("orders", {}).items()}},
            "shown_departments": self._current_departments(),
        }
        saved_snap = dict(self.saved_hashes) if self.saved_hashes else {}
        out_dir_path = plan_options["output_dir"]
        pdf_options = self._pdf_options()

        def work2():
            ok = fail = 0
            plan = BatchPlan()
            details: list[str] = []
            try:
                # Prima si validano tutte le righe, poi si generano tutti i dossier insieme
                # (una sola sessione Office, oppure in parallelo senza Office).
                plan = plan_batch(rows, tpl_snap, **plan_options)
                details = [s_tr(key, **kw) for key, kw in plan.notes]

                def progress_cb(step: int, total: int):
                    self._post("progress", {"step": step, "total": total})

                results = build_pdfs(plan.jobs, progress_cb, **pdf_options)
                for idx, job, error in zip(plan.rows, plan.jobs, results):
                    if error is None:
                        self._save_history(job.output_path, job.employee_name, job.department,
                                           len(job.templates), job.entry_date)
                        ok += 1
                        details.append(f"#{idx} OK · {job.output_path.name}")
                    else:
                        fail += 1
                        details.append(s_tr("bat_fail_generic", i=idx, e=str(error)))
                remember_template_hashes(tpl_snap, saved_snap)
            finally:
                self._post("done_batch", {
                    "ok": ok, "skip": plan.skipped, "fail": fail,
                    "detail": "\n".join(details[-30:]),
                    "out_dir": str(out_dir_path),
                    "auto_open": bool(auto_open_batch),
                })
                self._post("worker_done", None)

        threading.Thread(target=work2, daemon=True).start()

    # Lettura degli elenchi in formazioni.batch_import; qui restano i nomi usati dall'app.
    def _parse_date_str(self, s: str):
        return parse_date(s)

    def _normalize_column_header(self, c: str) -> str:
        return normalize_column_header(c)

    def _coerce_date_value(self, val) -> str:
        return coerce_date_value(val)

    def _parse_batch_file(self, p: Path) -> list[dict[str, str]]:
        try:
            return parse_batch_file(p)
        except UnsupportedBatchFile:
            raise ValueError(self.tr("bat_xls_unsupported")) from None

    # ---------------------- Modello Excel batch ---------------------------
    def download_batch_template(self):
        target = filedialog.asksaveasfilename(
            title=self.tr("bat_template_download"),
            initialdir=self.output_dir.get(),
            initialfile=self.tr("bat_template_filename"),
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not target:
            return
        departments = [d for d in department_options(self.templates)
                       if d.upper() not in ALL_DEPARTMENT_NAMES]
        try:
            write_batch_template(Path(target), departments,
                                 italian=self.language_code == "it")
        except OSError as exc:
            messagebox.showerror(self.tr("mb_error_title"), str(exc))
            return
        self.status.set(self.tr("bat_template_saved", name=Path(target).name))
        self._toast(self.tr("bat_template_saved", name=Path(target).name), "success",
                    (self.tr("ts_open"), lambda: open_folder(Path(target))))
