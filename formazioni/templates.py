"""Moduli (template): nomi, reparti, impronte MD5 e gestione dei file."""

from __future__ import annotations

import hashlib
import json
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import (
    ALL_DEPARTMENT_NAMES,
    DEPARTMENTS_FILE,
    FILENAME_PATTERN,
    HASHES_FILE,
    SUPPORTED_EXTENSIONS,
)

# --------------------------- HASHING --------------------------------------

_HASH_CACHE: dict[str, tuple[int, int, str]] = {}


def compute_template_hash(path: Path) -> str:
    # Rileggere ogni template a ogni aggiornamento della lista e' lento con molti
    # file o cartelle di rete: si ricalcola solo se cambiano data o dimensione.
    try:
        stat = path.stat()
    except OSError:
        return ""
    key = str(path)
    cached = _HASH_CACHE.get(key)
    if cached and cached[0] == stat.st_mtime_ns and cached[1] == stat.st_size:
        return cached[2]
    hasher = hashlib.md5()
    try:
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(65536), b""):
                hasher.update(chunk)
    except OSError:
        return ""
    digest = hasher.hexdigest()
    _HASH_CACHE[key] = (stat.st_mtime_ns, stat.st_size, digest)
    return digest


def load_saved_hashes() -> dict[str, str]:
    if not HASHES_FILE.exists():
        return {}
    try:
        data = json.loads(HASHES_FILE.read_text(encoding="utf-8"))
        return {str(k): str(v) for k, v in data.items()} if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_hashes(data: dict[str, str]) -> None:
    try:
        HASHES_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass


def classify_template_hashes(
    templates: list[Any], saved: dict[str, str]
) -> dict[str, str]:
    result: dict[str, str] = {}
    for template in templates:
        key = str(template.path)
        current = compute_template_hash(template.path)
        previous = saved.get(key)
        if previous is None:
            result[key] = "new"
        elif previous == current:
            result[key] = "ok"
        else:
            result[key] = "modified"
    return result


# --------------------------- BUSINESS -------------------------------------

def load_departments_from_file() -> list[str]:
    if not DEPARTMENTS_FILE.exists():
        return []
    try:
        raw = DEPARTMENTS_FILE.read_text(encoding="utf-8")
    except OSError:
        return []
    departments: list[str] = []
    seen: set[str] = set()
    for line in raw.splitlines():
        name = line.strip()
        if not name or name.startswith("#"):
            continue
        key = name.upper()
        if key not in seen:
            seen.add(key)
            departments.append(key)
    return departments


@dataclass(frozen=True)
class TemplateFile:
    path: Path
    department: str
    copies: int
    code: str

    @property
    def display_name(self) -> str:
        return self.path.name

    @property
    def is_for_every_department(self) -> bool:
        return self.department.upper() in ALL_DEPARTMENT_NAMES


def parse_template(path: Path) -> TemplateFile | None:
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return None
    match = FILENAME_PATTERN.fullmatch(path.stem)
    if not match:
        return None
    return TemplateFile(
        path=path,
        department=match.group("department"),
        copies=int(match.group("count")),
        code=match.group("code").upper(),
    )


def discover_templates(folder: Path) -> tuple[list[TemplateFile], list[Path]]:
    valid: list[TemplateFile] = []
    ignored: list[Path] = []
    if not folder.exists():
        return valid, ignored
    for path in sorted(folder.rglob("*")):
        if TRASH_DIR_NAME in path.relative_to(folder).parts:
            continue
        if path.is_file() and path.name.lower() != "readme.md":
            template = parse_template(path)
            if template and template.copies > 0:
                valid.append(template)
            elif path.suffix.lower() in SUPPORTED_EXTENSIONS:
                ignored.append(path)
    return valid, ignored


def department_options(templates: list[TemplateFile]) -> list[str]:
    from_file = load_departments_from_file()
    from_templates = {
        template.department.upper()
        for template in templates
        if not template.is_for_every_department
    }
    merged: list[str] = []
    seen: set[str] = set()
    for name in from_file:
        if name not in seen:
            seen.add(name)
            merged.append(name)
    for name in sorted(from_templates):
        if name not in seen:
            seen.add(name)
            merged.append(name)
    return merged


def templates_for_department(
    templates: list[TemplateFile], department: str
) -> list[TemplateFile]:
    chosen = department.strip().upper()
    matching = [
        template
        for template in templates
        if template.is_for_every_department
        or template.department.upper() == chosen
    ]
    return sorted(
        matching,
        key=lambda item: (
            item.is_for_every_department is False,
            item.department.upper(),
            item.path.name.lower(),
        ),
    )


def templates_for_departments(
    templates: list[TemplateFile], departments: list[str]
) -> list[TemplateFile]:
    chosen = {d.strip().upper() for d in departments if d and d.strip()}
    seen_paths: set[Path] = set()
    result: list[TemplateFile] = []
    for template in templates:
        key = template.path
        if key in seen_paths:
            continue
        if template.is_for_every_department or (template.department.upper() in chosen):
            seen_paths.add(key)
            result.append(template)
    return sorted(
        result,
        key=lambda item: (
            item.is_for_every_department is False,
            item.department.upper(),
            item.path.name.lower(),
        ),
    )


# --------------------------- GESTIONE MODULI ------------------------------

TRASH_DIR_NAME = "_moduli_eliminati"
_INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def suggest_template_code(department: str) -> str:
    letters = re.sub(r"[^A-Za-z]", "", department).upper()
    return letters[:3] if len(letters) >= 2 else "GEN"


def build_template_filename(department: str, copies: int, code: str, suffix: str) -> str:
    return f"{department.strip().upper()}_{int(copies)}_{code.strip().upper()}{suffix.lower()}"


def validate_template_fields(department: str, copies: str, code: str) -> str | None:
    """Restituisce la chiave di traduzione dell'errore, oppure None se i campi sono validi."""
    dept = department.strip()
    if not dept:
        return "tm_err_dept"
    if _INVALID_FILENAME_CHARS.search(dept):
        return "tm_err_dept_chars"
    try:
        if int(copies) < 1:
            return "tm_err_copies"
    except (TypeError, ValueError):
        return "tm_err_copies"
    if not re.fullmatch(r"[A-Za-z]{2,5}", code.strip()):
        return "tm_err_code"
    return None


def move_template_to_trash(path: Path, folder: Path) -> Path:
    trash = folder / TRASH_DIR_NAME
    trash.mkdir(parents=True, exist_ok=True)
    target = trash / path.name
    if target.exists():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        target = trash / f"{path.stem}__{stamp}{path.suffix}"
    shutil.move(str(path), str(target))
    return target


def safe_file_part(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9À-ÿ_-]+", "-", value.strip())
    return cleaned.strip("-_") or "persona"


# --------------------------- NOMI E ORDINI --------------------------------

# Nomi dei moduli mostrati solo nell'app e ordini dei documenti salvati come
# modelli: stanno nella cartella template, cosi' chi la condivide in rete li vede.
MODULE_SETTINGS_NAME = "_nomi_e_ordini.json"


def template_key(path: Path, folder: Path) -> str:
    """Chiave stabile di un modulo: percorso relativo alla cartella template."""
    try:
        return path.relative_to(folder).as_posix()
    except ValueError:
        return path.name


def load_module_settings(folder: Path) -> dict[str, Any]:
    data: dict[str, Any] = {"labels": {}, "orders": {}}
    try:
        raw = json.loads((folder / MODULE_SETTINGS_NAME).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return data
    if not isinstance(raw, dict):
        return data
    labels = raw.get("labels")
    if isinstance(labels, dict):
        data["labels"] = {str(k): str(v).strip() for k, v in labels.items() if str(v).strip()}
    orders = raw.get("orders")
    if isinstance(orders, dict):
        for name, entry in orders.items():
            if not isinstance(entry, dict):
                continue
            data["orders"][str(name)] = {
                "order": [str(k) for k in entry.get("order") or [] if isinstance(k, str)],
                "excluded": [str(k) for k in entry.get("excluded") or [] if isinstance(k, str)],
                "department": str(entry.get("department") or "").strip().upper(),
            }
    return data


def save_module_settings(folder: Path, data: dict[str, Any]) -> None:
    """Scrive nomi e ordini; solleva OSError se la cartella non e' scrivibile."""
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / MODULE_SETTINGS_NAME
    temp = target.with_suffix(".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(target)


def apply_order(templates: list[TemplateFile], order: list[str],
                folder: Path) -> list[TemplateFile]:
    """Mette prima i moduli nell'ordine indicato; gli altri restano in coda
    nell'ordine di partenza (per esempio i moduli aggiunti dopo aver salvato)."""
    if not order:
        return list(templates)
    position = {key: i for i, key in enumerate(order)}
    tail = len(position)
    return sorted(templates,
                  key=lambda t: position.get(template_key(t.path, folder), tail))


def order_for_department(data: dict[str, Any], department: str) -> str | None:
    """Nome dell'ordine salvato collegato al reparto, se c'e'."""
    wanted = department.strip().upper()
    if not wanted:
        return None
    for name, entry in data.get("orders", {}).items():
        if entry.get("department") == wanted:
            return name
    return None


def rename_template_key(data: dict[str, Any], old: str, new: str) -> None:
    """Dopo aver rinominato un modulo, nome e posizione negli ordini lo seguono."""
    labels = data.setdefault("labels", {})
    if old in labels:
        labels[new] = labels.pop(old)
    for entry in data.get("orders", {}).values():
        for field in ("order", "excluded"):
            entry[field] = [new if key == old else key for key in entry.get(field, [])]


def templates_from_saved_order(templates: list[TemplateFile], entry: dict[str, Any],
                               folder: Path) -> list[TemplateFile]:
    """Moduli di un dossier secondo un ordine salvato: in quell'ordine e senza gli esclusi."""
    excluded = set(entry.get("excluded", []))
    return [t for t in apply_order(templates, entry.get("order", []), folder)
            if template_key(t.path, folder) not in excluded]
