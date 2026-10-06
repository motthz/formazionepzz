"""Conversione con Microsoft Office / LibreOffice e operazioni sui PDF."""

from __future__ import annotations

import atexit
import contextlib
import functools
import hashlib
import os
import queue
import shutil
import subprocess
import tempfile
import threading
import time
from concurrent.futures import Future, as_completed
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


def _start_office_app(prog_id: str):
    """Avvia una nuova istanza di Word/Excel. Se un'istanza precedente si sta ancora
    chiudendo, Windows puo' rifiutare l'avvio ("esecuzione del server non riuscita"):
    si riprova dopo una breve pausa prima di arrendersi."""
    import win32com.client  # type: ignore[import-not-found]

    for attempt in range(3):
        try:
            return win32com.client.DispatchEx(prog_id)
        except Exception:  # noqa: BLE001 - pywintypes.com_error
            if attempt == 2:
                raise
            time.sleep(1.5 * (attempt + 1))


def _quiet(app, **settings) -> None:
    """Imposta le proprieta' che velocizzano l'automazione; quelle non supportate
    dalla versione di Office installata si ignorano."""
    for name, value in settings.items():
        with contextlib.suppress(Exception):
            setattr(app, name, value)


def _start_word():
    word = _start_office_app("Word.Application")
    _quiet(word, Visible=False, DisplayAlerts=0, ScreenUpdating=False,
           AutomationSecurity=3)  # 3 = macro sempre disattivate
    _quiet(word.Options, CheckSpellingAsYouType=False, CheckGrammarAsYouType=False,
           BackgroundSave=False, SaveNormalPrompt=False, UpdateLinksAtOpen=False)
    return word


def _start_excel():
    excel = _start_office_app("Excel.Application")
    _quiet(excel, Visible=False, DisplayAlerts=False, ScreenUpdating=False, EnableEvents=False,
           AskToUpdateLinks=False, AutomationSecurity=3)
    return excel


# Se Office non parte, per un po' non si riprova: ogni tentativo costa secondi
# e il ripiego su ReportLab arriva subito.
OFFICE_RETRY_AFTER_FAILURE = 60.0
# L'istanza di Word/Excel resta aperta tra una generazione e l'altra: l'avvio
# costa piu' della conversione di un modulo. Si chiude dopo questa inattivita'.
OFFICE_IDLE_SECONDS = 600.0


class _OfficeSession:
    """Una istanza di Word o Excel in un thread dedicato (un oggetto COM si usa solo
    dal thread che lo ha creato). I lavori arrivano in coda e tornano come Future."""

    def __init__(self, starter) -> None:
        self._starter = starter
        self._queue: queue.Queue = queue.Queue()
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._failure: tuple[float, BaseException] | None = None

    def submit(self, work) -> Future:
        """work(app) gira nel thread di Office; il Future ne restituisce il risultato."""
        future: Future = Future()
        with self._lock:
            if self._thread is None:
                self._thread = threading.Thread(target=self._run, daemon=True,
                                                name="FormazioniPZZ-office")
                self._thread.start()
            self._queue.put((work, future))
        return future

    def warm_up(self) -> None:
        self.submit(lambda _app: None)

    def shutdown(self, timeout: float = 15.0) -> None:
        with self._lock:
            thread = self._thread
            if thread is not None:
                self._queue.put(None)
        if thread is not None:
            thread.join(timeout)

    def _application(self, app):
        if app is not None:
            try:
                app.Name  # noqa: B018 - un'istanza chiusa o bloccata solleva un errore
                return app
            except Exception:  # noqa: BLE001
                pass
        if self._failure and time.monotonic() - self._failure[0] < OFFICE_RETRY_AFTER_FAILURE:
            raise self._failure[1]
        try:
            app = self._starter()
        except Exception as exc:
            self._failure = (time.monotonic(), exc)
            raise
        self._failure = None
        return app

    def _run(self) -> None:
        import pythoncom  # type: ignore[import-not-found]

        pythoncom.CoInitialize()
        app = None
        try:
            while True:
                try:
                    item = self._queue.get(timeout=OFFICE_IDLE_SECONDS)
                except queue.Empty:
                    with self._lock:
                        if not self._queue.empty():
                            continue
                        self._thread = None
                    return
                if item is None:
                    with self._lock:
                        self._thread = None
                    return
                work, future = item
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    app = self._application(app)
                    future.set_result(work(app))
                except BaseException as exc:  # noqa: BLE001 - finisce nel Future
                    future.set_exception(exc)
        finally:
            if app is not None:
                with contextlib.suppress(Exception):
                    app.Quit()
            app = None
            pythoncom.CoUninitialize()


_SESSIONS: dict[str, _OfficeSession] = {}
_SESSIONS_LOCK = threading.Lock()


def _session(kind: str) -> _OfficeSession:
    with _SESSIONS_LOCK:
        if kind not in _SESSIONS:
            _SESSIONS[kind] = _OfficeSession(_start_word if kind == "word" else _start_excel)
            if len(_SESSIONS) == 1:
                atexit.register(shutdown_office)
        return _SESSIONS[kind]


def _office_kind(path: Path) -> str:
    return "word" if path.suffix.lower() in {".doc", ".docx"} else "excel"


def warm_up_office(suffixes: set[str] | None = None) -> None:
    """Avvia Word/Excel in background prima che servano (es. mentre si compila il
    modulo), cosi' la generazione non aspetta l'apertura del programma."""
    if _office_command() != "microsoft-office":
        return
    suffixes = {s.lower() for s in suffixes} if suffixes is not None else {".docx", ".xlsx"}
    kinds = {"word" if s in {".doc", ".docx"} else "excel"
             for s in suffixes if s in {".doc", ".docx", ".xls", ".xlsx"}}
    for kind in kinds:
        _session(kind).warm_up()


def shutdown_office() -> None:
    with _SESSIONS_LOCK:
        sessions = list(_SESSIONS.values())
    for session in sessions:
        session.shutdown()


def _export_word(word, source: Path, target: Path, extension: str) -> None:
    # Sola lettura e niente elenco "file recenti": Word non crea il file di blocco ~$
    document = word.Documents.Open(str(source.resolve()), False, True, False)
    try:
        if extension == "pdf":
            # Argomenti: file, PDF (17), non aprire, qualita' stampa (0), tutto il
            # documento (0, da 1 a 1 ignorati), solo contenuto (0), senza proprieta',
            # IRM, senza segnalibri (0) ne' tag di struttura (export piu' rapido),
            # font mancanti come immagini, senza PDF/A
            document.ExportAsFixedFormat(str(target.resolve()), 17, False, 0, 0, 1, 1, 0,
                                         False, True, 0, False, True, False)
        else:
            document.SaveAs2(str(target.resolve()), 16)
    finally:
        document.Close(False)


def _export_excel(excel, source: Path, target: Path, extension: str) -> None:
    workbook = excel.Workbooks.Open(str(source.resolve()), 0, True)  # UpdateLinks=0, ReadOnly
    try:
        if extension == "pdf":
            workbook.ExportAsFixedFormat(0, str(target.resolve()), 0, False, False)
        else:
            workbook.SaveAs(str(target.resolve()), 51)
    finally:
        workbook.Close(False)


def _submit_conversion(source: Path, target: Path, extension: str) -> Future:
    kind = _office_kind(source)
    export = _export_word if kind == "word" else _export_excel
    return _session(kind).submit(lambda app: export(app, source, target, extension))


def _convert_with_ms_office(source: Path, converted: Path, extension: str) -> None:
    _submit_conversion(source, converted, extension).result()


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
    """Word ed Excel lavorano in parallelo, ognuno sulla propria coda; on_each
    si chiama da questo thread, non da quelli di Office."""
    futures = {
        _submit_conversion(source, output_dir / f"{source.stem}.{extension}", extension): source
        for source in sources
    }
    converted: dict[Path, Path] = {}
    try:
        for future in as_completed(futures):
            future.result()
            source = futures[future]
            converted[source] = output_dir / f"{source.stem}.{extension}"
            if on_each:
                on_each()
    except BaseException:
        for future in futures:
            future.cancel()
        raise
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
