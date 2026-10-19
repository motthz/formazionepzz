"""Regole della generazione di un dossier, senza interfaccia: controllo dei dati,
nome del file in uscita e impronte dei moduli usati."""

from __future__ import annotations

import os
import tempfile
from datetime import datetime
from pathlib import Path

from .config import PREVIEW_PREFIX
from .templates import TemplateFile, compute_template_hash, safe_file_part, save_hashes


def check_request(name: str, entry_date: str, departments: list[str],
                  templates: list[TemplateFile], require_name: bool = True) -> tuple[str, str] | None:
    """Primo dato mancante per generare un dossier, come chiavi di lingua
    (titolo, messaggio); None se si puo' generare."""
    if not name.strip() and require_name:
        return "mb_missing_title", "mb_missing_name"
    if not entry_date:
        return "mb_missing_title", "mb_missing_date"
    if not departments:
        return "mb_missing_title", "mb_missing_dept"
    if not templates:
        return "mb_no_docs_title", "mb_no_docs_body"
    return None


def output_path_for(output_dir: Path, name: str, departments: list[str],
                    reserved: set[Path] | None = None) -> Path:
    """dossier_<nome>_<reparti>.pdf, con _2, _3... se il file esiste gia'. I nomi in
    `reserved` (file di un lotto non ancora scritti) contano come esistenti e il nome
    scelto vi si aggiunge."""
    taken = reserved if reserved is not None else set()
    stem = f"dossier_{safe_file_part(name)}_{safe_file_part('+'.join(departments))}"
    path = output_dir / f"{stem}.pdf"
    counter = 2
    while path.exists() or path in taken:
        path = output_dir / f"{stem}_{counter}.pdf"
        counter += 1
    if reserved is not None:
        reserved.add(path)
    return path


def preview_output_path() -> Path:
    """File temporaneo per un'anteprima (lo ripulisce cleanup_old_previews)."""
    return Path(tempfile.gettempdir()) / (
        f"{PREVIEW_PREFIX}{os.getpid()}_{datetime.now():%H%M%S%f}.pdf")


def remember_template_hashes(templates: list[TemplateFile], saved: dict[str, str]) -> None:
    """Dopo una generazione i moduli usati non risultano piu' modificati."""
    hashes = dict(saved)
    for template in templates:
        try:
            hashes[str(template.path)] = compute_template_hash(template.path)
        except Exception:  # noqa: BLE001 - un modulo illeggibile resta "modificato"
            pass
    save_hashes(hashes)


def hash_counts(hash_status: dict[str, str]) -> dict[str, int]:
    """Quanti moduli sono invariati, modificati o nuovi rispetto all'ultima generazione."""
    values = list(hash_status.values())
    return {"ok": values.count("ok"), "mod": values.count("modified"), "new": values.count("new")}
