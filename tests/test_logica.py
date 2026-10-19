"""Logica senza interfaccia (lotti, dossier, storico, ordini): deve funzionare e
importarsi senza Tkinter, cosi' la puo' usare anche un backend web."""

from __future__ import annotations

import subprocess
import sys
from datetime import date
from pathlib import Path

from formazioni import history
from formazioni.batch_import import (
    UnsupportedBatchFile,
    parse_batch_file,
    parse_date,
    people_from_records,
    plan_batch,
)
from formazioni.dossier import check_request, hash_counts, output_path_for
from formazioni.ordering import (
    excluded_keys,
    merge_visible_order,
    save_order,
    set_template_label,
    template_label,
)
from formazioni.templates import discover_templates, order_for_department

ROOT = Path(__file__).resolve().parent.parent
CORE_MODULES = ["batch_import", "dossier", "history", "ordering", "templates", "pdf",
                "documents", "office", "config"]


def test_core_modules_import_without_tkinter() -> None:
    code = ("import sys; sys.modules['tkinter'] = None\n"
            + "".join(f"import formazioni.{m}\n" for m in CORE_MODULES)
            + "assert 'formazioni.tkcompat' not in sys.modules, 'tkcompat caricato'\n")
    result = subprocess.run([sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr


def _folder(tmp: Path) -> Path:
    for name in ("TUTTI_1_GEN.docx", "VENDITE_1_AAA.docx", "VENDITE_2_BBB.xlsx",
                 "MAGAZZINO_1_MAG.docx"):
        (tmp / name).write_bytes(b"x")
    return tmp


def test_parse_date_formats() -> None:
    assert parse_date("01/10/2026") == date(2026, 10, 1)
    assert parse_date("2026-10-01") == date(2026, 10, 1)
    assert parse_date("1.10.26") == date(2026, 10, 1)
    assert parse_date("ieri") is None
    assert parse_date("") is None


def test_csv_import_and_people(tmp: Path) -> None:
    csv_file = tmp / "elenco.csv"
    csv_file.write_text("Nome,Data di ingresso,Reparto\nMario Rossi,2026-10-01 00:00:00,vendite\n"
                        ",01/10/2026,VENDITE\nAnna\n\n", encoding="utf-8")
    records = parse_batch_file(csv_file)
    assert records[0] == {"Nome": "Mario Rossi", "Data": "01/10/2026", "Reparto": "vendite"}
    people, ignored = people_from_records(records)
    assert people == [{"Nome": "Mario Rossi", "Data": "01/10/2026", "Reparto": "VENDITE"}]
    assert ignored == 2
    xls = tmp / "vecchio.xls"
    xls.write_bytes(b"x")
    try:
        parse_batch_file(xls)
    except UnsupportedBatchFile:
        pass
    else:
        raise AssertionError("un .xls deve essere rifiutato")


def test_output_path_skips_existing_and_reserved(tmp: Path) -> None:
    first = output_path_for(tmp, "Mario Rossi", ["VENDITE"])
    assert first.name == "dossier_Mario-Rossi_VENDITE.pdf"
    first.write_bytes(b"x")
    reserved: set[Path] = set()
    second = output_path_for(tmp, "Mario Rossi", ["VENDITE"], reserved)
    third = output_path_for(tmp, "Mario Rossi", ["VENDITE"], reserved)
    assert (second.name, third.name) == ("dossier_Mario-Rossi_VENDITE_2.pdf",
                                         "dossier_Mario-Rossi_VENDITE_3.pdf")


def test_check_request_and_hash_counts() -> None:
    assert check_request("", "01/10/2026", ["VENDITE"], []) == ("mb_missing_title", "mb_missing_name")
    assert check_request("", "", ["VENDITE"], [], require_name=False) == \
        ("mb_missing_title", "mb_missing_date")
    assert check_request("Anna", "01/10/2026", ["VENDITE"], []) == ("mb_no_docs_title", "mb_no_docs_body")
    assert hash_counts({"a": "ok", "b": "modified", "c": "new", "d": "ok"}) == \
        {"ok": 2, "mod": 1, "new": 1}


def test_plan_batch_uses_saved_order_and_reports_skips(tmp: Path) -> None:
    (tmp / "tpl").mkdir()
    folder = _folder(tmp / "tpl")
    templates, _ = discover_templates(folder)
    settings: dict = {"labels": {}, "orders": {}}
    vendite = [t.path for t in templates if t.department in ("TUTTI", "VENDITE")]
    save_order(settings, "Vendite", list(reversed(vendite)),
               {folder / "TUTTI_1_GEN.docx": False}, folder, "VENDITE")
    people = [
        {"Nome": "Mario", "Data": "01/10/2026", "Reparto": "VENDITE"},
        {"Nome": "Anna", "Data": "boh", "Reparto": "VENDITE"},
        {"Nome": "Luca", "Data": "01/10/2026", "Reparto": "ACQUISTI"},
        {"Nome": "Sara", "Data": "01/10/2026", "Reparto": ""},
    ]
    out = tmp / "out"
    plan = plan_batch(people, templates, folder=folder, output_dir=out, doc_order=[],
                      inclusion={}, module_settings=settings, shown_departments=["MAGAZZINO"])
    assert plan.rows == [1]
    assert [t.path.name for t in plan.jobs[0].templates] == ["VENDITE_2_BBB.xlsx", "VENDITE_1_AAA.docx"]
    assert plan.jobs[0].output_path == out / "dossier_Mario_VENDITE.pdf"
    assert plan.skipped == 3
    assert [key for key, _ in plan.notes] == ["bat_skip_date", "bat_skip_dept", "bat_skip_nodept"]


def test_orders_and_labels(tmp: Path) -> None:
    folder = _folder(tmp)
    templates, _ = discover_templates(folder)
    data: dict = {"labels": {}, "orders": {}}
    gen = next(t for t in templates if t.path.name == "TUTTI_1_GEN.docx")
    set_template_label(data, gen.path, folder, "Benvenuto")
    assert template_label(data, gen, folder) == "Benvenuto"
    set_template_label(data, gen.path, folder, "")
    assert template_label(data, gen, folder) == "TUTTI_1_GEN.docx"

    save_order(data, "A", [gen.path], {gen.path: False}, folder, "VENDITE")
    save_order(data, "B", [gen.path], {}, folder, "VENDITE")
    # Un reparto ha un solo ordine collegato
    assert order_for_department(data, "VENDITE") == "B"
    assert data["orders"]["A"]["department"] == ""
    assert excluded_keys(data, "A") == {"TUTTI_1_GEN.docx"}
    assert excluded_keys(data, "manca") is None

    mag = folder / "MAGAZZINO_1_MAG.docx"
    order = merge_visible_order([mag], templates, [], folder)
    assert order[0] == "MAGAZZINO_1_MAG.docx" and len(order) == len(templates)


def test_history_roundtrip(tmp: Path) -> None:
    target = tmp / "storico.json"
    history.record_dossier(tmp / "a.pdf", "Mario", "VENDITE", 3, "01/10/2026", path=target)
    history.record_dossier(tmp / "b.pdf", "Anna", "MAGAZZINO", 2, path=target)
    entries = history.load_history(target)
    assert [e["name"] for e in entries] == ["Anna", "Mario"]
    assert history.matches_query(entries[1], "vend") and not history.matches_query(entries[1], "anna")
    history.remove_entry(entries[0], path=target)
    assert [e["name"] for e in history.load_history(target)] == ["Mario"]
