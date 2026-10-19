"""Ordini dei documenti salvati come modelli e nomi dei moduli, senza interfaccia.
Lavorano sul dizionario di load_module_settings / save_module_settings."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .templates import TemplateFile, apply_order, template_key


def template_label(data: dict[str, Any], template: TemplateFile, folder: Path) -> str:
    """Nome mostrato di un modulo: quello scelto dall'utente, altrimenti il nome del file."""
    return data.get("labels", {}).get(template_key(template.path, folder)) or template.path.name


def set_template_label(data: dict[str, Any], path: Path, folder: Path, label: str) -> None:
    """Imposta il nome mostrato di un modulo; un nome vuoto torna al nome del file."""
    labels = data.setdefault("labels", {})
    key = template_key(path, folder)
    if label:
        labels[key] = label
    else:
        labels.pop(key, None)


def full_order(templates: list[TemplateFile], order: list[str], folder: Path) -> list[str]:
    """Chiavi di tutti i moduli nell'ordine indicato."""
    return [template_key(t.path, folder) for t in apply_order(templates, order, folder)]


def merge_visible_order(visible: list[Path], templates: list[TemplateFile], order: list[str],
                        folder: Path) -> list[str]:
    """Nuovo ordine dopo aver riordinato la lista: i moduli mostrati (`visible`) vanno
    in testa nell'ordine dato, gli altri mantengono la loro posizione relativa."""
    keys = [template_key(p, folder) for p in visible]
    shown = set(keys)
    return keys + [k for k in full_order(templates, order, folder) if k not in shown]


def excluded_keys(data: dict[str, Any], name: str | None) -> set[str] | None:
    """Moduli esclusi da un ordine salvato; None se l'ordine non esiste."""
    entry = data.get("orders", {}).get(name or "")
    return None if entry is None else set(entry.get("excluded", []))


def save_order(data: dict[str, Any], name: str, visible: list[Path], included: dict[Path, bool],
               folder: Path, department: str = "") -> None:
    """Salva (o sovrascrive) un ordine con i moduli nell'ordine mostrato e quelli non
    spuntati come esclusi. Un reparto ha un solo ordine collegato: chi lo aveva lo perde."""
    orders = data.setdefault("orders", {})
    if department:
        for other in orders.values():
            if other.get("department") == department:
                other["department"] = ""
    orders[name] = {
        "order": [template_key(p, folder) for p in visible],
        "excluded": [template_key(p, folder) for p in visible if not included.get(p, True)],
        "department": department,
    }


def linked_department(data: dict[str, Any], name: str | None) -> str:
    """Reparto collegato a un ordine salvato, o stringa vuota."""
    return data.get("orders", {}).get(name or "", {}).get("department") or ""
