"""Conversione con Microsoft Office / LibreOffice e operazioni sui PDF."""

from __future__ import annotations

import contextlib
import functools
import hashlib
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from .documents import fill_office_placeholders


@functools.lru_cache(maxsize=1)
def _office_command() -> str | None:
    for command in ("libreoffice", "soffice"):
        if shutil.which(command):
            return command
    if os.name == "nt":
        try:
            import win32com.client  # type: ignore[import-not-found]  # noqa: F401
        except ImportError:
            return None
        return "microsoft-office"
    return None


@contextlib.contextmanager
def _com_apartment():
    """COM va inizializzato in ogni thread che usa Word/Excel (la generazione gira in un worker)."""
    import pythoncom  # type: ignore[import-not-found]

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def _start_office_app(prog_id: str):
    """Avvia una nuova istanza di Word/Excel. Se un'istanza precedente si sta ancora
    chiudendo, Windows puo' rifiutare l'avvio ("esecuzione del server non riuscita"):
    si riprova dopo una breve pausa prima di arrendersi."""
    import time

    import win32com.client  # type: ignore[import-not-found]

    for attempt in range(3):
        try:
            return win32com.client.DispatchEx(prog_id)
        except Exception:  # noqa: BLE001 - pywintypes.com_error
            if attempt == 2:
                raise
            time.sleep(1.5 * (attempt + 1))


def _convert_with_ms_office(source: Path, converted: Path, extension: str) -> None:

    if source.suffix.lower() in {".doc", ".docx"}:
        word = _start_office_app("Word.Application")
        document = word.Documents.Open(str(source.resolve()))
        try:
            if extension == "pdf":
                document.ExportAsFixedFormat(str(converted.resolve()), 17)
            else:
                document.SaveAs2(str(converted.resolve()), FileFormat=16)
        finally:
            document.Close(False)
            word.Quit()
    else:
        excel = _start_office_app("Excel.Application")
        workbook = excel.Workbooks.Open(str(source.resolve()))
        try:
            if extension == "pdf":
                workbook.ExportAsFixedFormat(0, str(converted.resolve()))
            else:
                workbook.SaveAs(str(converted.resolve()), FileFormat=51)
        finally:
            workbook.Close(False)
            excel.Quit()


def _convert_with_office(source: Path, output_dir: Path, extension: str) -> Path:
    command = _office_command()
    if command is None:
        raise RuntimeError(
            "Per mantenere identico il layout dei template Word/Excel serve "
            "LibreOffice installato e disponibile nel PATH."
        )
    output_dir.mkdir(parents=True, exist_ok=True)
    converted = output_dir / f"{source.stem}.{extension}"
    if command == "microsoft-office":
        with _com_apartment():
            _convert_with_ms_office(source, converted, extension)
        if not converted.exists():
            raise RuntimeError(f"Conversione Microsoft Office fallita per {source.name}")
        return converted

    profile = output_dir / "office-profile"
    profile_uri = profile.resolve().as_uri()
    result = subprocess.run(
        [
            command,
            "--headless",
            "--convert-to",
            extension,
            "--outdir",
            str(output_dir),
            f"-env:UserInstallation={profile_uri}",
            str(source),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0 or not converted.exists():
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Conversione Office fallita per {source.name}: {detail}")
    return converted


def _convert_with_ms_office_batch(
    sources: list[Path], output_dir: Path, extension: str, on_each=None
) -> dict[Path, Path]:

    applications = {}
    converted: dict[Path, Path] = {}
    try:
        for source in sources:
            kind = "word" if source.suffix.lower() == ".docx" else "excel"
            if kind not in applications:
                applications[kind] = (
                    _start_office_app("Word.Application")
                    if kind == "word"
                    else _start_office_app("Excel.Application")
                )
            application = applications[kind]
            target = output_dir / f"{source.stem}.{extension}"
            document = (
                application.Documents.Open(str(source.resolve()))
                if kind == "word"
                else application.Workbooks.Open(str(source.resolve()))
            )
            try:
                if extension == "pdf":
                    if kind == "word":
                        document.ExportAsFixedFormat(str(target.resolve()), 17)
                    else:
                        document.ExportAsFixedFormat(0, str(target.resolve()))
            finally:
                document.Close(False)
            converted[source] = target
            if on_each:
                on_each()
    finally:
        for application in applications.values():
            application.Quit()
    return converted


def _convert_with_office_batch(
    sources: list[Path], output_dir: Path, extension: str, on_each=None
) -> dict[Path, Path]:
    if not sources:
        return {}
    command = _office_command()
    if command is None:
        raise RuntimeError("Motore Office non disponibile per la conversione dei template.")
    output_dir.mkdir(parents=True, exist_ok=True)
    if command == "microsoft-office":
        with _com_apartment():
            return _convert_with_ms_office_batch(sources, output_dir, extension, on_each)

    profile = output_dir / "office-profile"
    result = subprocess.run(
        [
            command,
            "--headless",
            "--convert-to",
            extension,
            "--outdir",
            str(output_dir),
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            *[str(source) for source in sources],
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    converted = {source: output_dir / f"{source.stem}.{extension}" for source in sources}
    if result.returncode != 0 or any(not path.exists() for path in converted.values()):
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Conversione Office fallita: {detail}")
    if on_each:
        for _source in sources:
            on_each()
    return converted


def _prepare_office_template(
    template: Path,
    work_dir: Path,
    employee_name: str,
    entry_date: str,
) -> Path:
    suffix = template.suffix.lower()
    work_dir.mkdir(parents=True, exist_ok=True)
    if suffix in {".doc", ".xls"}:
        modern_extension = "docx" if suffix == ".doc" else "xlsx"
        source = _convert_with_office(template, work_dir, modern_extension)
    else:
        source = work_dir / template.name
        shutil.copy2(template, source)

    if source.suffix.lower() in {".docx", ".xlsx"}:
        # Solo il testo dei segnaposto cambia: immagini e intestazioni restano intatte
        fill_office_placeholders(source, employee_name, entry_date)
    return source


def _merge_pdfs(output_path: Path, pdf_paths: list[Path]) -> None:
    from pypdf import PdfWriter

    writer = PdfWriter()
    for pdf_path in pdf_paths:
        writer.append(str(pdf_path))
    with output_path.open("wb") as handle:
        writer.write(handle)


def _remove_trailing_blank_pages(pdf_path: Path, dest_dir: Path | None = None) -> Path:
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(str(pdf_path))
    last_page = len(reader.pages)
    while last_page > 1:
        page = reader.pages[last_page - 1]
        text = page.extract_text() or ""
        resources = page.get("/Resources")
        if resources is not None:
            resources = resources.get_object()
        has_images = bool(resources and resources.get("/XObject"))
        content = page.get_contents()
        has_drawing_commands = bool(content and content.get_data().strip())
        if text.strip() or has_images or has_drawing_commands:
            break
        last_page -= 1
    pages = reader.pages[:last_page]
    if len(pages) == len(reader.pages):
        return pdf_path

    target_dir = dest_dir if dest_dir is not None else Path(tempfile.gettempdir())
    trimmed = target_dir / f"{pdf_path.stem}-trimmed-{os.getpid()}-{hashlib.md5(str(pdf_path).encode()).hexdigest()[:8]}.pdf"
    writer = PdfWriter()
    for page in pages:
        writer.add_page(page)
    with trimmed.open("wb") as handle:
        writer.write(handle)
    return trimmed


def _pdf_is_landscape(pdf_path: Path) -> bool:
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    return bool(reader.pages and reader.pages[0].mediabox.width > reader.pages[0].mediabox.height)
