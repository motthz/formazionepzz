"""Registro degli errori in %LOCALAPPDATA%\\FormazioniPZZ\\formazioni.log.

Nell'exe senza console sys.stderr e' None: le tracce degli errori (traceback.print_exc,
eccezioni non gestite, errori nei callback di Tk) andrebbero perse, o farebbero fallire
proprio il codice che le stampa. Qui stderr viene indirizzato al file di registro;
avviando da sorgente si scrive sia sul terminale sia sul file.
"""

from __future__ import annotations

import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from .config import APP_VERSION

LOG_LIMIT_BYTES = 1_000_000


def log_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    return (Path(base) if base else Path(tempfile.gettempdir())) / "FormazioniPZZ"


def log_file() -> Path:
    return log_dir() / "formazioni.log"


class _Tee:
    """Scrive su piu' flussi; un errore di uno non blocca gli altri."""

    def __init__(self, *streams) -> None:
        self.streams = [s for s in streams if s is not None]

    def write(self, text: str) -> int:
        for stream in self.streams:
            try:
                stream.write(text)
                stream.flush()
            except Exception:  # noqa: BLE001 - il registro non deve mai far fallire l'app
                pass
        return len(text)

    def flush(self) -> None:
        for stream in self.streams:
            try:
                stream.flush()
            except Exception:  # noqa: BLE001
                pass


def setup_error_log() -> Path | None:
    """Attiva il registro; ritorna il percorso del file, o None se non scrivibile."""
    path = log_file()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists() and path.stat().st_size > LOG_LIMIT_BYTES:
            # Una sola copia precedente: il registro resta piccolo
            os.replace(path, path.with_suffix(".log.1"))
        handle = open(path, "a", encoding="utf-8", buffering=1)  # noqa: SIM115 - resta aperto
    except OSError:
        return None
    handle.write(f"\n=== {datetime.now():%Y-%m-%d %H:%M:%S} · Formazioni PZZ {APP_VERSION} "
                 f"· Python {sys.version.split()[0]} · {sys.platform} ===\n")
    console = sys.__stderr__
    sys.stderr = _Tee(console, handle)
    if sys.stdout is None:  # exe senza console
        sys.stdout = _Tee(handle)
    return path
