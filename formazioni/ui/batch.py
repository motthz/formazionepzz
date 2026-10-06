"""Elenco batch: aggiunta persone, import CSV/Excel, generazione multipla."""

from __future__ import annotations

import csv
import threading
from datetime import date, datetime
from pathlib import Path
from typing import Any

from ..batch_excel import write_batch_template
from ..config import ALL_DEPARTMENT_NAMES
from ..documents import load_workbook
from ..icons import strip_leading_symbol
from ..pdf import DossierJob, build_pdfs
from ..system import open_folder
from ..templates import (
    compute_template_hash,
    department_options,
    safe_file_part,
    save_hashes,
    templates_for_departments,
)
from ..tkcompat import END, StringVar, filedialog, messagebox


class BatchMixin:
    """Parte di FormazioniApp: elenco batch: aggiunta persone, import CSV/Excel, generazione multipla."""

    def _inline_batch_count(self) -> StringVar:
        self._inline_batch_count_var = var = StringVar()

        def _fmt(n):
            if n == 0:
                return self.tr("bat_count_none")
            return self.tr("bat_count", n=n)

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
        text = (self.tr("bat_scope_same", n=n, d=dept) if dept
                else self.tr("bat_scope_mixed", n=n))
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
            self.update_document_list()

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
        for r in loaded:
            nome = (r.get("Nome") or r.get("name") or "").strip()
            data = (r.get("Data") or r.get("date") or "").strip()
            reparto = (r.get("Reparto") or r.get("department") or "").strip().upper()
            if not nome or not data:
                continue
            row = {"Nome": nome, "Data": data, "Reparto": reparto}
            self._inline_batch_rows.append(row)
            tree.insert("", END, values=(nome, data, reparto))
        self._refresh_inline_batch_count()
        self._show_batch_department()
        self.warm_up_office()
        self._toast(self.tr("dd_imported", n=len(self._inline_batch_rows) - before,
                            name=Path(p).name), "success")

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
        inc_snap = dict(self.template_inclusion)
        saved_snap = dict(self.saved_hashes) if self.saved_hashes else {}
        dept_opts = {d.upper() for d in department_options(tpl_snap)}
        out_dir_path = Path(self.output_dir.get()).expanduser()
        pdf_options = self._pdf_options()

        def work2():
            ok = skip = fail = 0
            details: list[str] = []
            # Prima si validano tutte le righe, poi si generano tutti i dossier insieme
            # (una sola sessione Office, oppure in parallelo senza Office).
            jobs: list[DossierJob] = []
            job_rows: list[int] = []
            reserved: set[Path] = set()
            try:
                for idx, row in enumerate(rows, start=1):
                    nome = (row.get("Nome") or "").strip()
                    data = (row.get("Data") or "").strip()
                    reparto_raw = (row.get("Reparto") or "").strip()
                    reparto_list = [d.strip().upper() for d in reparto_raw.split("+") if d.strip()]
                    if not reparto_list:
                        if len(dept_opts) == 1:
                            reparto_list = [next(iter(dept_opts))]
                        else:
                            skip += 1
                            details.append(s_tr("bat_skip_dept", i=idx, d=reparto_raw or ""))
                            continue
                    note = ""
                    ruolo = ""
                    if not nome:
                        skip += 1
                        details.append(s_tr("bat_skip_empty", i=idx))
                        continue
                    parsed = self._parse_date_str(data)
                    if parsed is None:
                        skip += 1
                        details.append(s_tr("bat_skip_date", i=idx))
                        continue
                    invalid_depts = [d for d in reparto_list if d not in dept_opts]
                    if invalid_depts:
                        skip += 1
                        details.append(s_tr("bat_skip_dept", i=idx, d=", ".join(invalid_depts)))
                        continue
                    sel = templates_for_departments(tpl_snap, reparto_list)
                    filt = [t for t in sel if inc_snap.get(t.path, True)]
                    if not filt:
                        skip += 1
                        for d in reparto_list:
                            details.append(s_tr("bat_skip_dept", i=idx, d=d))
                        continue
                    dept_tag = safe_file_part("+".join(reparto_list))
                    out_file = out_dir_path / (
                        f"dossier_{safe_file_part(nome)}_{dept_tag}.pdf"
                    )
                    c = 2
                    # I file del lotto non esistono ancora: si riservano i nomi gia' assegnati
                    while out_file.exists() or out_file in reserved:
                        out_file = out_dir_path / (
                            f"dossier_{safe_file_part(nome)}_{dept_tag}_{c}.pdf"
                        )
                        c += 1
                    reserved.add(out_file)
                    jobs.append(DossierJob(out_file, nome, parsed.strftime("%d/%m/%Y"),
                                           "+".join(reparto_list), filt, ruolo, note))
                    job_rows.append(idx)

                def progress_cb(step: int, total: int):
                    self._post("progress", {"step": step, "total": total})

                results = build_pdfs(jobs, progress_cb, **pdf_options)
                for idx, job, error in zip(job_rows, jobs, results):
                    if error is None:
                        self._save_history(job.output_path, job.employee_name, job.department,
                                           len(job.templates), job.entry_date)
                        ok += 1
                        details.append(f"#{idx} OK · {job.output_path.name}")
                    else:
                        fail += 1
                        details.append(s_tr("bat_fail_generic", i=idx, e=str(error)))
                hashes = dict(saved_snap)
                for t in tpl_snap:
                    try:
                        hashes[str(t.path)] = compute_template_hash(t.path)
                    except Exception:
                        pass
                save_hashes(hashes)
            finally:
                self._post("done_batch", {
                    "ok": ok, "skip": skip, "fail": fail,
                    "detail": "\n".join(details[-30:]),
                    "out_dir": str(out_dir_path),
                    "auto_open": bool(auto_open_batch),
                })
                self._post("worker_done", None)

        threading.Thread(target=work2, daemon=True).start()

    def _parse_date_str(self, s: str):
        if not s:
            return None
        if isinstance(s, datetime):
            return s.date()
        if isinstance(s, date):
            return s
        s_norm = str(s).strip().replace("-", "/").replace(".", "/")
        for fmt in ("%d/%m/%Y", "%Y/%m/%d", "%d/%m/%y"):
            try:
                return datetime.strptime(s_norm, fmt).date()
            except ValueError:
                continue
        return None

    def _normalize_column_header(self, c: str) -> str:
        c_norm = str(c).strip().lower().replace("  ", " ")
        if any(k in c_norm for k in ("nome", "cognome", "name", "full name", "dipendente", "employee", "persona")):
            return "Nome"
        if any(k in c_norm for k in ("data", "date", "ingresso", "entrata", "entry")):
            return "Data"
        if any(k in c_norm for k in ("reparto", "department", "dipartimento", "settore", "area", "ufficio")):
            return "Reparto"
        if any(k in c_norm for k in ("ruolo", "role", "mansione", "qualifica", "position", "job")):
            return "Ruolo"
        if any(k in c_norm for k in ("note", "notes", "commento", "commenti", "comment", "osservazioni")):
            return "Note"
        return c

    def _coerce_date_value(self, val) -> str:
        if val is None:
            return ""
        if isinstance(val, datetime):
            return val.strftime("%d/%m/%Y")
        if isinstance(val, date):
            return val.strftime("%d/%m/%Y")
        s = str(val).strip()
        if not s:
            return ""
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S",
                    "%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        return s

    def _parse_batch_file(self, p: Path) -> list[dict[str, str]]:
        suffix = p.suffix.lower()
        if suffix == ".xls":
            raise ValueError(self.tr("bat_xls_unsupported"))
        if suffix == ".xlsx":
            wb = load_workbook(p, data_only=True, read_only=True)
            try:
                ws = wb.worksheets[0]
                rows_iter = ws.iter_rows(values_only=True)
                raw_rows = [list(r) for r in rows_iter]
            finally:
                wb.close()
            if not raw_rows:
                return []
            header = [str(c).strip().lower() if c is not None else "" for c in raw_rows[0]]
            norm = [self._normalize_column_header(c) for c in header]
            out: list[dict[str, str]] = []
            for r in raw_rows[1:]:
                if all(v is None or str(v).strip() == "" for v in r):
                    continue
                obj: dict[str, str] = {}
                for key, val in zip(norm, r):
                    if key == "Data":
                        obj[key] = self._coerce_date_value(val)
                    else:
                        obj[key] = "" if val is None else str(val).strip()
                out.append(obj)
            return out
        # CSV
        text = p.read_text(encoding="utf-8-sig", errors="replace")
        reader = csv.reader(text.splitlines())
        rows_list = list(reader)
        if not rows_list:
            return []
        header = [c.strip().lower() for c in rows_list[0]]
        norm = [self._normalize_column_header(c) for c in header]
        out = []
        for r in rows_list[1:]:
            if all(v.strip() == "" for v in r):
                continue
            while len(r) < len(norm):
                r.append("")
            obj = {}
            for k, v in zip(norm, r):
                if k == "Data":
                    obj[k] = self._coerce_date_value(v)
                else:
                    obj[k] = v.strip()
            out.append(obj)
        return out

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
