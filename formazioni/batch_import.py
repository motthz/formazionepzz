"""Lotti di dossier, senza interfaccia: lettura di un elenco CSV/Excel di persone e
preparazione dei dossier da generare, con il motivo di ogni riga scartata."""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .documents import load_workbook
from .dossier import output_path_for
from .pdf import DossierJob
from .templates import (
    TemplateFile,
    apply_order,
    department_options,
    order_for_department,
    templates_for_departments,
    templates_from_saved_order,
)


class UnsupportedBatchFile(ValueError):
    """Formato dell'elenco non leggibile (per esempio il vecchio .xls)."""


def parse_date(value: Any) -> date | None:
    """Data da testo (gg/mm/aaaa, aaaa-mm-gg, gg.mm.aa...) o da date/datetime."""
    if not value:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s_norm = str(value).strip().replace("-", "/").replace(".", "/")
    for fmt in ("%d/%m/%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s_norm, fmt).date()
        except ValueError:
            continue
    return None


def normalize_column_header(c: str) -> str:
    """Intestazione di colonna riconosciuta (Nome, Data, Reparto, Ruolo, Note) in piu' lingue."""
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


def coerce_date_value(val: Any) -> str:
    """Cella data di Excel/CSV come testo gg/mm/aaaa; il testo non riconosciuto resta com'e'."""
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
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


def _records(header: list[Any], rows: list[list[Any]]) -> list[dict[str, str]]:
    norm = [normalize_column_header(str(c).strip().lower() if c is not None else "") for c in header]
    out: list[dict[str, str]] = []
    for r in rows:
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        out.append({key: coerce_date_value(val) if key == "Data"
                    else ("" if val is None else str(val).strip())
                    for key, val in zip(norm, r)})
    return out


def parse_batch_file(path: Path) -> list[dict[str, str]]:
    """Righe di un elenco .xlsx o .csv, con le colonne riconosciute rinominate."""
    suffix = path.suffix.lower()
    if suffix == ".xls":
        raise UnsupportedBatchFile(path.name)
    if suffix == ".xlsx":
        wb = load_workbook(path, data_only=True, read_only=True)
        try:
            raw_rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
        finally:
            wb.close()
    else:
        text = path.read_text(encoding="utf-8-sig", errors="replace")
        raw_rows = list(csv.reader(text.splitlines()))
        if raw_rows:
            # Righe CSV piu' corte dell'intestazione: le colonne mancanti restano vuote
            width = len(raw_rows[0])
            raw_rows = [raw_rows[0]] + [r + [""] * (width - len(r)) for r in raw_rows[1:]]
    if not raw_rows:
        return []
    return _records(raw_rows[0], raw_rows[1:])


def people_from_records(records: list[dict[str, str]]) -> tuple[list[dict[str, str]], int]:
    """Persone da aggiungere al lotto (Nome, Data, Reparto) e quante righe sono
    state ignorate perche' senza nome o data."""
    people: list[dict[str, str]] = []
    ignored = 0
    for r in records:
        nome = (r.get("Nome") or r.get("name") or "").strip()
        data = (r.get("Data") or r.get("date") or "").strip()
        reparto = (r.get("Reparto") or r.get("department") or "").strip().upper()
        if not nome or not data:
            ignored += 1
            continue
        people.append({"Nome": nome, "Data": data, "Reparto": reparto})
    return people, ignored


@dataclass
class BatchPlan:
    """Dossier da generare per un lotto. `rows` e' il numero di riga (da 1) di ogni
    dossier; `notes` sono i motivi degli scarti come (chiave di lingua, argomenti)."""
    jobs: list[DossierJob] = field(default_factory=list)
    rows: list[int] = field(default_factory=list)
    skipped: int = 0
    notes: list[tuple[str, dict[str, Any]]] = field(default_factory=list)


def plan_batch(people: list[dict[str, str]], templates: list[TemplateFile], *,
               folder: Path, output_dir: Path, doc_order: list[str],
               inclusion: dict[Path, bool], module_settings: dict[str, Any],
               shown_departments: list[str]) -> BatchPlan:
    """Valida le righe e prepara un DossierJob per ciascuna. L'ordine e le spunte
    correnti (`doc_order`, `inclusion`) valgono per chi e' dei reparti mostrati
    (`shown_departments`); gli altri ricevono l'ordine salvato per il loro reparto,
    se c'e', altrimenti le stesse spunte."""
    plan = BatchPlan()
    dept_opts = {d.upper() for d in department_options(templates)}
    reserved: set[Path] = set()

    def skip(key: str, **kwargs: Any) -> None:
        plan.notes.append((key, kwargs))

    for idx, row in enumerate(people, start=1):
        nome = (row.get("Nome") or "").strip()
        data = (row.get("Data") or "").strip()
        reparto_raw = (row.get("Reparto") or "").strip()
        reparto_list = [d.strip().upper() for d in reparto_raw.split("+") if d.strip()]
        if not reparto_list:
            if len(dept_opts) == 1:
                reparto_list = [next(iter(dept_opts))]
            else:
                plan.skipped += 1
                skip("bat_skip_nodept", i=idx)
                continue
        if not nome:
            plan.skipped += 1
            skip("bat_skip_empty", i=idx)
            continue
        parsed = parse_date(data)
        if parsed is None:
            plan.skipped += 1
            skip("bat_skip_date", i=idx)
            continue
        invalid_depts = [d for d in reparto_list if d not in dept_opts]
        if invalid_depts:
            plan.skipped += 1
            skip("bat_skip_dept", i=idx, d=", ".join(invalid_depts))
            continue
        sel = templates_for_departments(templates, reparto_list)
        linked = (order_for_department(module_settings, reparto_list[0])
                  if len(reparto_list) == 1 and reparto_list != shown_departments else None)
        if linked:
            chosen = templates_from_saved_order(sel, module_settings["orders"][linked], folder)
        else:
            chosen = [t for t in apply_order(sel, doc_order, folder) if inclusion.get(t.path, True)]
        if not chosen:
            plan.skipped += 1
            for d in reparto_list:
                skip("bat_skip_dept", i=idx, d=d)
            continue
        out_file = output_path_for(output_dir, nome, reparto_list, reserved)
        plan.jobs.append(DossierJob(out_file, nome, parsed.strftime("%d/%m/%Y"),
                                    "+".join(reparto_list), chosen, "", ""))
        plan.rows.append(idx)
    return plan
