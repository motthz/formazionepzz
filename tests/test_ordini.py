"""Ordine dei documenti, ordini salvati per reparto e nomi dei moduli in app."""

from __future__ import annotations

from pathlib import Path

from formazioni.templates import (
    MODULE_SETTINGS_NAME,
    apply_order,
    discover_templates,
    load_module_settings,
    order_for_department,
    rename_template_key,
    save_module_settings,
    templates_for_department,
    templates_from_saved_order,
)


def _folder(tmp: Path) -> Path:
    for name in ("TUTTI_1_GEN.docx", "VENDITE_1_AAA.docx", "VENDITE_2_BBB.xlsx",
                 "MAGAZZINO_1_MAG.docx"):
        (tmp / name).write_bytes(b"x")
    return tmp


def _names(templates) -> list[str]:
    return [t.path.name for t in templates]


def test_apply_order_puts_unknown_templates_last(tmp: Path) -> None:
    folder = _folder(tmp)
    templates, _ = discover_templates(folder)
    vendite = templates_for_department(templates, "VENDITE")
    assert _names(vendite) == ["TUTTI_1_GEN.docx", "VENDITE_1_AAA.docx", "VENDITE_2_BBB.xlsx"]
    ordered = apply_order(vendite, ["VENDITE_2_BBB.xlsx", "TUTTI_1_GEN.docx"], folder)
    # Il modulo non presente nell'ordine salvato (aggiunto dopo) va in fondo
    assert _names(ordered) == ["VENDITE_2_BBB.xlsx", "TUTTI_1_GEN.docx", "VENDITE_1_AAA.docx"]
    assert apply_order(vendite, [], folder) == vendite


def test_saved_order_round_trip_and_department_link(tmp: Path) -> None:
    folder = _folder(tmp)
    data = {
        "labels": {"TUTTI_1_GEN.docx": "Regolamento aziendale"},
        "orders": {"Vendite": {"order": ["VENDITE_2_BBB.xlsx", "TUTTI_1_GEN.docx"],
                               "excluded": ["VENDITE_1_AAA.docx"],
                               "department": "VENDITE"}},
    }
    save_module_settings(folder, data)
    loaded = load_module_settings(folder)
    assert loaded == data
    assert order_for_department(loaded, "vendite") == "Vendite"
    assert order_for_department(loaded, "MAGAZZINO") is None

    templates, ignored = discover_templates(folder)
    assert not ignored, "il file di nomi e ordini non e' un modulo"
    chosen = templates_from_saved_order(templates_for_department(templates, "VENDITE"),
                                        loaded["orders"]["Vendite"], folder)
    assert _names(chosen) == ["VENDITE_2_BBB.xlsx", "TUTTI_1_GEN.docx"]


def test_load_ignores_broken_file(tmp: Path) -> None:
    (tmp / MODULE_SETTINGS_NAME).write_text("{non json", encoding="utf-8")
    assert load_module_settings(tmp) == {"labels": {}, "orders": {}}
    assert load_module_settings(tmp / "manca") == {"labels": {}, "orders": {}}


def test_rename_keeps_label_and_position() -> None:
    data = {"labels": {"A_1_AA.docx": "Uno"},
            "orders": {"X": {"order": ["B_1_BB.docx", "A_1_AA.docx"],
                             "excluded": ["A_1_AA.docx"], "department": ""}}}
    rename_template_key(data, "A_1_AA.docx", "A_2_AA.docx")
    assert data["labels"] == {"A_2_AA.docx": "Uno"}
    assert data["orders"]["X"]["order"] == ["B_1_BB.docx", "A_2_AA.docx"]
    assert data["orders"]["X"]["excluded"] == ["A_2_AA.docx"]
