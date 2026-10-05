"""Percorsi, versione, impostazioni, lingue e tema."""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Any


def _resolve_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "release":
            return exe_dir.parent
        return exe_dir
    # Da sorgente: la cartella del progetto, sopra il pacchetto formazioni/
    return Path(__file__).resolve().parent.parent


APP_VERSION = "2.3.1"
DEFAULT_UPDATE_SOURCE = "https://api.github.com/repos/motthz/formazionepzz/releases/latest"
APP_DIR = _resolve_app_dir()
DEFAULT_TEMPLATE_DIR = APP_DIR / "templates"
DEFAULT_OUTPUT_DIR = APP_DIR / "output"
HISTORY_FILE = APP_DIR / ".formazioni_history.json"
HISTORY_LIMIT = 500
PREVIEW_PREFIX = "formazioni_anteprima_"
DEPARTMENTS_FILE = APP_DIR / "reparti.txt"
SUPPORTED_EXTENSIONS = {".doc", ".docx", ".xls", ".xlsx", ".pdf"}
ALL_DEPARTMENT_NAMES = {"TUTTI", "TUTTE", "ALL"}
FILENAME_PATTERN = re.compile(
    r"^(?P<department>.+)_(?P<count>\d+)_(?P<code>[A-Za-z]{2,5})$",
    re.IGNORECASE,
)


SETTINGS_FILE = APP_DIR / "settings.json"
HASHES_FILE = APP_DIR / ".template_hashes.json"


# File di sola lettura (lingue, icone): nell'exe PyInstaller stanno in _MEIPASS,
# non accanto all'exe; da sorgente coincidono con APP_DIR.
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
LANG_DIR = RESOURCE_DIR / "lang" if (RESOURCE_DIR / "lang").is_dir() else APP_DIR / "lang"
ICON_DIR = RESOURCE_DIR / "assets" if (RESOURCE_DIR / "assets").is_dir() else APP_DIR / "assets"
ICON_ICO = ICON_DIR / "app_icon.ico"
ICON_PNG = ICON_DIR / "app_icon.png"
LOGO_HEADER_PNG = ICON_DIR / "logo_header.png"
DEFAULT_LANG = "en"
DEFAULT_THEME = "light"
MONTH_KEYS = (
    "dp_jan", "dp_feb", "dp_mar", "dp_apr", "dp_may", "dp_jun",
    "dp_jul", "dp_aug", "dp_sep", "dp_oct", "dp_nov", "dp_dec",
)


SYSTEM_THEME_POLL_MS = 4000


def windows_prefers_dark() -> bool:
    """Legge "Modalita' app" di Windows (Impostazioni > Personalizzazione > Colori)."""
    if os.name != "nt":
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def resolve_theme(preference: str) -> str:
    if preference == "system":
        return "dark" if windows_prefers_dark() else "light"
    return "dark" if preference == "dark" else "light"


# --------------------------- SETTINGS -------------------------------------

def _default_settings() -> dict[str, Any]:
    return {
        "theme": DEFAULT_THEME,
        "language": DEFAULT_LANG,
        "last_template_dir": str(DEFAULT_TEMPLATE_DIR),
        "last_output_dir": str(DEFAULT_OUTPUT_DIR),
        "pdf_watermark": "",
        "pdf_protect": False,
        "check_updates": True,
        "update_source": "",
        "last_update_check": "",
    }


def load_settings() -> dict[str, Any]:
    data = _default_settings()
    if SETTINGS_FILE.exists():
        try:
            raw = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                data.update({k: v for k, v in raw.items() if k in data})
        except (OSError, json.JSONDecodeError):
            pass
    return data


def save_settings(data: dict[str, Any]) -> None:
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass


# --------------------------- I18N -----------------------------------------

def available_languages() -> list[str]:
    if not LANG_DIR.exists():
        return [DEFAULT_LANG]
    return sorted(p.stem for p in LANG_DIR.glob("*.json"))


def load_language(code: str) -> dict[str, str]:
    fallback: dict[str, str] = {}
    if (LANG_DIR / f"{DEFAULT_LANG}.json").exists():
        try:
            fallback = json.loads(
                (LANG_DIR / f"{DEFAULT_LANG}.json").read_text(encoding="utf-8")
            )
        except (OSError, json.JSONDecodeError):
            fallback = {}
    chosen_file = LANG_DIR / f"{code}.json"
    chosen: dict[str, str] = {}
    if chosen_file.exists():
        try:
            chosen = json.loads(chosen_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            chosen = {}
    merged = dict(fallback)
    merged.update(chosen)
    return merged
