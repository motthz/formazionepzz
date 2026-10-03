"""Controllo degli aggiornamenti (release GitHub, cartella di rete o URL)."""

from __future__ import annotations

import json
import re
import urllib.request
from pathlib import Path

from .config import APP_VERSION


def parse_version(value: str) -> tuple[int, ...]:
    parts = re.findall(r"\d+", str(value))
    return tuple(int(p) for p in parts[:4]) or (0,)


def fetch_update_manifest(source: str, timeout: float = 4.0) -> dict[str, str] | None:
    """Legge il manifest {version, url, notes} da un URL http(s) o da un percorso
    locale / di rete (file version.json o cartella che lo contiene)."""
    source = source.strip()
    if not source:
        return None
    if source.lower().startswith(("http://", "https://")):
        request = urllib.request.Request(
            source, headers={"User-Agent": f"FormazioniPZZ/{APP_VERSION}"})
        with urllib.request.urlopen(request, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8-sig"))
        default_url = ""
        if isinstance(data, dict) and "tag_name" in data:
            # API GitHub "releases/latest": tag vX.Y.Z, link diretto al setup se allegato
            setup = next((a.get("browser_download_url") for a in data.get("assets") or []
                          if str(a.get("name", "")).lower() == "formazionipzz_setup.exe"), None)
            data = {"version": str(data["tag_name"]).lstrip("vV"),
                    "url": setup or data.get("html_url") or "",
                    "notes": data.get("name") or ""}
    else:
        manifest = Path(source).expanduser()
        if manifest.is_dir():
            manifest = manifest / "version.json"
        data = json.loads(manifest.read_text(encoding="utf-8-sig"))
        default_url = str(manifest.parent / "FormazioniPZZ_Setup.exe")
    if not isinstance(data, dict) or not data.get("version"):
        raise ValueError("Manifest aggiornamenti non valido: manca 'version'.")
    return {
        "version": str(data["version"]),
        "url": str(data.get("url") or default_url),
        "notes": str(data.get("notes") or ""),
    }


SETUP_NAME = "FormazioniPZZ_Setup.exe"


def is_installed_copy(app_dir: Path) -> bool:
    """True se l'app gira dalla cartella creata dall'installer (si puo' aggiornare da sola).
    L'exe portatile e l'avvio da sorgente si aggiornano invece a mano."""
    import sys

    return (getattr(sys, "frozen", False) and (app_dir / "_internal").is_dir()
            and (app_dir / "Disinstalla.exe").exists())


def download_setup(url: str, target_dir: Path, version: str, timeout: float = 30.0) -> Path:
    """Scarica il setup in target_dir; il file definitivo compare solo a download completo."""
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"FormazioniPZZ_Setup-{version}.exe"
    partial = target.with_suffix(".part")
    request = urllib.request.Request(url, headers={"User-Agent": f"FormazioniPZZ/{APP_VERSION}"})
    with urllib.request.urlopen(request, timeout=timeout) as response, partial.open("wb") as handle:
        while chunk := response.read(256 * 1024):
            handle.write(chunk)
    if partial.stat().st_size < 1_000_000:  # una pagina di errore, non un installer
        partial.unlink(missing_ok=True)
        raise ValueError("Il file scaricato non e' un programma di installazione valido.")
    partial.replace(target)
    return target
