"""Storico dei dossier generati: lettura e scrittura del file JSON, senza interfaccia."""

from __future__ import annotations

import json
import threading
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import HISTORY_FILE, HISTORY_LIMIT

_lock = threading.Lock()


def _history_path(path: Path | None) -> Path:
    # Letto a ogni chiamata: i test possono dirottare HISTORY_FILE su una cartella temporanea
    return path if path is not None else HISTORY_FILE


def load_history(path: Path | None = None) -> list[dict[str, Any]]:
    target = _history_path(path)
    if not target.exists():
        return []
    try:
        history = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return [h for h in history if isinstance(h, dict)] if isinstance(history, list) else []


def write_history(history: list[dict[str, Any]], path: Path | None = None) -> None:
    try:
        _history_path(path).write_text(
            json.dumps(history[:HISTORY_LIMIT], ensure_ascii=False, indent=2),
            encoding="utf-8")
    except OSError:
        # La cronologia e' accessoria: un errore qui non deve far fallire il dossier
        traceback.print_exc()


def record_dossier(output_path: Path, name: str, department: str, total: int,
                   entry_date: str = "", path: Path | None = None) -> None:
    """Aggiunge in cima allo storico un dossier appena generato."""
    with _lock:
        history = load_history(path)
        history.insert(0, {
            "path": str(output_path),
            "name": name,
            "entry_date": entry_date,
            "department": department,
            "documents": total,
            "ts": datetime.now().isoformat(timespec="seconds"),
        })
        write_history(history, path)


def remove_entry(entry: dict[str, Any], path: Path | None = None) -> None:
    with _lock:
        write_history([h for h in load_history(path) if h != entry], path)


def matches_query(entry: dict[str, Any], query: str) -> bool:
    """Ricerca nello storico: per nome, reparto, data di ingresso o percorso."""
    needle = query.strip().lower()
    if not needle:
        return True
    haystack = " ".join(str(entry.get(k, "")) for k in
                        ("name", "department", "entry_date", "path")).lower()
    return needle in haystack
