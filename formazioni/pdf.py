"""Motore di generazione dei dossier PDF (Office o ReportLab), cache e batch."""

from __future__ import annotations

import copy as _copy_mod
import functools
import os
import re
import secrets
import shutil
import tempfile
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm

if TYPE_CHECKING:
    from reportlab.platypus import Paragraph, Table
from .config import APP_VERSION
from .documents import Document, load_workbook, replace_docx_placeholders, replace_placeholders
from .office import (
    _convert_with_office_batch,
    _merge_pdfs,
    _office_command,
    _pdf_is_landscape,
    _prepare_office_template,
    _remove_trailing_blank_pages,
)
from .templates import TemplateFile, compute_template_hash

_STYLE_CACHE: dict[str, object] | None = None
_TEMPLATE_STORY_CACHE: dict[tuple, list[object]] = {}


def clear_caches() -> None:
    global _STYLE_CACHE
    _STYLE_CACHE = None
    _TEMPLATE_STORY_CACHE.clear()
    _office_command.cache_clear()


@dataclass
class DossierJob:
    output_path: Path
    employee_name: str
    entry_date: str
    department: str
    templates: list[TemplateFile]
    role: str = ""
    notes: str = ""

    def expanded(self) -> list[tuple[TemplateFile, int]]:
        return [(t, n) for t in self.templates for n in range(1, t.copies + 1)]


_PLACEHOLDER_RE = re.compile(r"\*(nome|data)\*", re.IGNORECASE)
_XML_TAG_RE = re.compile(r"<[^>]+>")
_PLACEHOLDER_SCAN: dict[tuple[str, int, int], bool] = {}
PDF_CACHE_LIMIT = 200


def template_has_placeholders(path: Path) -> bool:
    """True se il modulo contiene *nome* o *data*, o se non si puo' stabilirlo
    (.doc/.xls binari): in quel caso va preparato per ogni persona."""
    if path.suffix.lower() not in {".docx", ".xlsx"}:
        return True
    try:
        stat = path.stat()
    except OSError:
        return True
    key = (str(path), stat.st_mtime_ns, stat.st_size)
    if key in _PLACEHOLDER_SCAN:
        return _PLACEHOLDER_SCAN[key]
    import zipfile

    try:
        with zipfile.ZipFile(path) as archive:
            # Senza i tag XML un segnaposto spezzato in piu' "run" torna contiguo
            found = any(
                _PLACEHOLDER_RE.search(_XML_TAG_RE.sub("", archive.read(name).decode("utf-8", "ignore")))
                for name in archive.namelist() if name.endswith(".xml")
            )
    except (OSError, zipfile.BadZipFile):
        found = True
    _PLACEHOLDER_SCAN[key] = found
    return found


def _pdf_cache_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    root = Path(base) / "FormazioniPZZ" if base else Path(tempfile.gettempdir()) / "FormazioniPZZ"
    return root / "pdf-cache"


def _cached_pdf_path(template: Path) -> Path | None:
    digest = compute_template_hash(template)
    if not digest:
        return None
    engine = re.sub(r"[^a-z]", "", (_office_command() or "none").lower())
    return _pdf_cache_dir() / f"{digest}-{engine}.pdf"


def _store_cached_pdf(template: Path, pdf: Path) -> Path:
    target = _cached_pdf_path(template)
    if target is None:
        return pdf
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(pdf, target)
        cached = sorted(target.parent.glob("*.pdf"), key=lambda f: f.stat().st_mtime)
        for old in cached[:-PDF_CACHE_LIMIT]:
            old.unlink(missing_ok=True)
        return target
    except OSError:
        return pdf


def _build_native_jobs(jobs: list[DossierJob], progress_cb=None) -> list[Exception | None]:
    """Prepara i moduli di tutti i dossier e li converte con un'unica sessione di
    Word/Excel (o una sola chiamata a LibreOffice). I moduli senza segnaposto si
    convertono una volta sola e restano in cache tra una generazione e l'altra.
    Ritorna, per ogni dossier, None se riuscito oppure l'errore."""
    results: list[Exception | None] = [None] * len(jobs)
    # Word/Excel possono tenere un file bloccato per qualche istante dopo la chiusura:
    # un residuo nella cartella temporanea non deve far fallire un dossier gia' pronto.
    with tempfile.TemporaryDirectory(prefix="formazioni-pdf-", ignore_cleanup_errors=True) as temp_name:
        temp_dir = Path(temp_name)
        names = iter(range(1, 1_000_000))
        ready: dict[Path, Path] = {}            # moduli fissi gia' convertiti (cache)
        static_sources: dict[Path, Path] = {}   # moduli fissi da convertire una volta
        personal: dict[tuple[int, Path], Path] = {}
        office_paths = [
            list(dict.fromkeys(t.path for t in job.templates if t.path.suffix.lower() != ".pdf"))
            for job in jobs
        ]
        total = max(1, sum(len(paths) for paths in office_paths) * 2 + len(jobs))
        step = 0

        def tick():
            nonlocal step
            step = min(step + 1, total)
            if progress_cb:
                progress_cb(step, total)

        for index, job in enumerate(jobs):
            try:
                for template_path in office_paths[index]:
                    if not template_has_placeholders(template_path):
                        if template_path not in ready and template_path not in static_sources:
                            cached = _cached_pdf_path(template_path)
                            if cached is not None and cached.exists():
                                ready[template_path] = cached
                            else:
                                copy = temp_dir / f"{next(names)}-{template_path.name}"
                                shutil.copy2(template_path, copy)
                                static_sources[template_path] = copy
                    else:
                        prepared = _prepare_office_template(
                            template_path, temp_dir / f"job-{index}-{next(names)}",
                            job.employee_name, job.entry_date)
                        # Nome univoco: i convertitori chiamano il PDF come il sorgente
                        unique = prepared.with_name(f"{next(names)}-{prepared.name}")
                        prepared.rename(unique)
                        personal[(index, template_path)] = unique
                    tick()
            except Exception as exc:  # noqa: BLE001 - l'errore resta legato a quel dossier
                results[index] = exc

        sources = list(static_sources.values()) + [
            source for (index, _), source in personal.items() if results[index] is None
        ]
        converted = _convert_with_office_batch(
            sources, temp_dir / "converted", "pdf", on_each=tick) if sources else {}
        for template_path, source in static_sources.items():
            trimmed = _remove_trailing_blank_pages(converted[source], temp_dir)
            ready[template_path] = _store_cached_pdf(template_path, trimmed)

        trimmed_cache: dict[Path, Path] = {}

        def trimmed(pdf: Path) -> Path:
            if pdf not in trimmed_cache:
                trimmed_cache[pdf] = _remove_trailing_blank_pages(pdf, temp_dir)
            return trimmed_cache[pdf]

        for index, job in enumerate(jobs):
            if results[index] is not None:
                tick()
                continue
            try:
                expanded = job.expanded()
                # Stessa regola del motore ReportLab: niente copertina se ci sono PDF orizzontali
                has_landscape_pdf = any(
                    template.path.suffix.lower() == ".pdf" and _pdf_is_landscape(template.path)
                    for template, _ in expanded
                )
                cover = [] if has_landscape_pdf else [_build_cover_pdf(
                    temp_dir / f"cover-{index}.pdf", job.employee_name, job.entry_date,
                    job.department, job.role, job.notes,
                )]
                pages = cover + [
                    trimmed(template.path) if template.path.suffix.lower() == ".pdf"
                    else ready.get(template.path)
                    or trimmed(converted[personal[(index, template.path)]])
                    for template, _copy_number in expanded
                ]
                job.output_path.parent.mkdir(parents=True, exist_ok=True)
                _merge_pdfs(job.output_path, pages)
            except Exception as exc:  # noqa: BLE001
                results[index] = exc
            tick()
    return results


def paragraph_text(text: str, style: ParagraphStyle) -> Paragraph:
    from reportlab.platypus import Paragraph

    safe = escape(text).replace("\n", "<br/>")
    return Paragraph(safe or " ", style)


TABLE_WIDTH = 174 * mm


def _table_widths(values: list[list[str]], available_width: float = TABLE_WIDTH) -> list[float]:
    column_count = max((len(row) for row in values), default=0)
    if not column_count:
        return []
    lengths = []
    for column in range(column_count):
        longest = max(
            (len(str(row[column]).replace("\n", " ").strip()) for row in values if column < len(row)),
            default=4,
        )
        lengths.append(max(4, min(longest, 36)))
    minimum = available_width / column_count
    weights = [max(length, 8) for length in lengths]
    total = sum(weights)
    widths = [available_width * weight / total for weight in weights]
    if any(width < minimum for width in widths):
        widths = [minimum] * column_count
    return widths


def _table_flowable(rows: list[list[object]], raw_values: list[list[str]]) -> Table:
    from reportlab.platypus import Table, TableStyle

    column_count = max((len(row) for row in rows), default=0)
    rows = [row + [""] * (column_count - len(row)) for row in rows]
    raw_values = [row + [""] * (column_count - len(row)) for row in raw_values]
    return Table(
        rows,
        colWidths=_table_widths(raw_values),
        repeatRows=1,
        splitByRow=1,
        splitInRow=0,
        hAlign="LEFT",
        style=TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e7eef1")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#173642")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#b9c9cf")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        ),
    )


def docx_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
) -> list[object]:
    from reportlab.platypus import Spacer

    document = Document(path)
    replace_docx_placeholders(document, employee_name, entry_date)
    story: list[object] = []
    style_name_cache: dict[int, str] = {}
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        heading = ""
        try:
            if paragraph.style is not None:
                pid = id(paragraph.style)
                if pid not in style_name_cache:
                    style_name_cache[pid] = (paragraph.style.name or "").lower()
                heading = style_name_cache[pid]
        except Exception:
            heading = ""
        style = styles["subheading"] if "heading" in heading else styles["body"]
        story.append(paragraph_text(text, style))
        story.append(Spacer(1, 2.4 * mm))
    for table in document.tables:
        rows: list[list[object]] = []
        raw_rows: list[list[str]] = []
        for row in table.rows:
            values = [cell.text.strip() for cell in row.cells]
            if any(values):
                while values and not values[-1]:
                    values.pop()
                raw_rows.append(values)
                rows.append([paragraph_text(value, styles["table"]) for value in values])
        if rows:
            story.append(Spacer(1, 2 * mm))
            story.append(_table_flowable(rows, raw_rows))
            story.append(Spacer(1, 4 * mm))
    return story or [paragraph_text("Documento senza contenuto testuale.", styles["muted"])]


def xlsx_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
) -> list[object]:
    from reportlab.platypus import Paragraph, Spacer

    workbook = load_workbook(path, data_only=False, read_only=False)
    story: list[object] = []
    try:
        for sheet in workbook.worksheets:
            rows: list[list[object]] = []
            raw_rows: list[list[str]] = []
            for row in sheet.iter_rows(values_only=True):
                values = [
                    replace_placeholders(value, employee_name, entry_date)
                    for value in row
                ]
                if any(value not in (None, "") for value in values):
                    values = [str(value) if value is not None else "" for value in values]
                    while values and not values[-1]:
                        values.pop()
                    raw_rows.append(values)
                    rows.append(
                        [paragraph_text(value, styles["table"]) for value in values]
                    )
            if rows:
                story.append(Paragraph(escape(sheet.title), styles["subheading"]))
                story.append(Spacer(1, 2 * mm))
                story.append(_table_flowable(rows, raw_rows))
                story.append(Spacer(1, 5 * mm))
    finally:
        workbook.close()
    return story or [paragraph_text("Foglio senza contenuto.", styles["muted"])]


def make_styles() -> dict[str, ParagraphStyle]:
    global _STYLE_CACHE
    if _STYLE_CACHE is not None:
        return _STYLE_CACHE  # type: ignore[return-value]
    base = getSampleStyleSheet()
    result = {
        "cover_title": ParagraphStyle("CoverTitle", parent=base["Title"], fontName="Helvetica-Bold",
                                     fontSize=25, leading=30, textColor=colors.HexColor("#173642"),
                                     alignment=TA_CENTER, spaceAfter=8 * mm),
        "cover_subtitle": ParagraphStyle("CoverSubtitle", parent=base["Normal"], fontName="Helvetica",
                                         fontSize=12, leading=17, textColor=colors.HexColor("#48636d"),
                                         alignment=TA_CENTER),
        "meta": ParagraphStyle("Meta", parent=base["Normal"], fontName="Helvetica",
                               fontSize=10.5, leading=16, textColor=colors.HexColor("#173642")),
        "heading": ParagraphStyle("Heading", parent=base["Heading1"], fontName="Helvetica-Bold",
                                  fontSize=15, leading=19, textColor=colors.HexColor("#173642"),
                                  spaceAfter=4 * mm),
        "subheading": ParagraphStyle("Subheading", parent=base["Heading2"], fontName="Helvetica-Bold",
                                     fontSize=11.5, leading=15, textColor=colors.HexColor("#247b7b"),
                                     spaceBefore=2 * mm, spaceAfter=2 * mm),
        "body": ParagraphStyle("Body", parent=base["BodyText"], fontName="Helvetica",
                               fontSize=10, leading=14, textColor=colors.HexColor("#263f47")),
        "table": ParagraphStyle("Table", parent=base["BodyText"], fontName="Helvetica",
                                fontSize=8.2, leading=10, textColor=colors.HexColor("#263f47")),
        "muted": ParagraphStyle("Muted", parent=base["BodyText"], fontName="Helvetica-Oblique",
                                fontSize=9.5, leading=13, textColor=colors.HexColor("#71838a")),
        "small": ParagraphStyle("Small", parent=base["BodyText"], fontName="Helvetica",
                                fontSize=8.5, leading=11, textColor=colors.HexColor("#71838a")),
        "pdf_page": ParagraphStyle("PdfPage", parent=base["BodyText"], fontName="Helvetica",
                                   fontSize=8.5, leading=11.5, textColor=colors.HexColor("#263f47")),
    }
    _STYLE_CACHE = result  # type: ignore[assignment]
    return result


def pdf_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
    frame_width: float = 174 * mm,
    frame_height: float = 263 * mm,
) -> list[object]:
    from reportlab.platypus import KeepInFrame, PageBreak, Paragraph, Spacer

    story: list[object] = []
    pdf_style = styles["pdf_page"]
    is_landscape = frame_width > frame_height
    if is_landscape:
        pdf_style = ParagraphStyle("PdfPageLandscape", parent=styles["pdf_page"],
                                   fontSize=7.5, leading=10)
    try:
        from pypdf import PdfReader
    except Exception:
        story.append(paragraph_text(
            "Libreria pypdf non disponibile per leggere il template PDF.", styles["muted"]))
        return story
    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        story.append(paragraph_text(f"Impossibile aprire il PDF: {exc}", styles["muted"]))
        return story
    if not reader.pages:
        story.append(paragraph_text("PDF senza pagine.", styles["muted"]))
        return story
    for idx, page in enumerate(reader.pages, start=1):
        page_story: list[object] = []
        if len(reader.pages) > 1:
            page_story.append(Paragraph(f"Pagina {idx} del PDF", styles["subheading"]))
        raw = ""
        try:
            raw = page.extract_text() or ""
        except Exception:
            raw = ""
        raw = replace_placeholders(raw, employee_name, entry_date)
        lines = [ln.rstrip() for ln in raw.splitlines()]
        if not any(ln.strip() for ln in lines):
            page_story.append(paragraph_text(
                "(Nessun testo estraibile dal PDF — file incluso come riferimento.)",
                styles["muted"]))
            page_story.append(Spacer(1, 3 * mm))
        else:
            for line in lines:
                if not line.strip():
                    page_story.append(Spacer(1, 1.5 * mm))
                    continue
                page_story.append(Paragraph(escape(line), pdf_style))
            page_story.append(Spacer(1, 4 * mm))
        story.append(KeepInFrame(frame_width, frame_height, page_story, mode="shrink"))
        if idx < len(reader.pages):
            story.append(PageBreak())
    return story or [paragraph_text("Documento PDF senza contenuto testuale.", styles["muted"])]


def _cover_story(
    styles: dict[str, ParagraphStyle],
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
) -> list[object]:
    from reportlab.platypus import Spacer

    cover_rows_raw = [
        ["Campo", "Valore"],
        ["Nome e Cognome", employee_name],
        ["Data Ingresso / Corso", entry_date],
        ["Reparto/i", department],
    ]
    if role and role.strip():
        cover_rows_raw.append(["Mansione / Ruolo", role.strip()])
    if notes and notes.strip():
        cover_rows_raw.append(["Note aggiuntive", notes.strip()])
    cover_rows = [
        [paragraph_text(str(c), styles["body"] if i else styles["meta"])
         for i, c in enumerate(rr)]
        for rr in cover_rows_raw
    ]
    return [
        Spacer(1, 12 * mm),
        paragraph_text("Dossier Formazione", styles["cover_title"]),
        paragraph_text("Documento di accompagnamento per la formazione individuale",
                       styles["cover_subtitle"]),
        Spacer(1, 8 * mm),
        _table_flowable(cover_rows, cover_rows_raw),
        Spacer(1, 8 * mm),
        paragraph_text(
            "Generato da Formazioni PZZ il: " + datetime.now().strftime("%d/%m/%Y %H:%M"),
            styles["small"],
        ),
    ]


def _build_cover_pdf(
    output_path: Path,
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
) -> Path:
    from reportlab.platypus import SimpleDocTemplate

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=14 * mm, leftMargin=14 * mm, topMargin=14 * mm, bottomMargin=14 * mm,
        title=f"Dossier formazione - {employee_name}",
        author="Formazioni PZZ",
    )
    doc.build(_cover_story(make_styles(), employee_name, entry_date, department, role, notes))
    return output_path


def _watermark_page(text: str, width: float, height: float):
    """Pagina PDF trasparente con la filigrana in diagonale, centrata."""
    import io
    import math

    from pypdf import PdfReader
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas as rl_canvas

    font = "Helvetica-Bold"
    diagonal = math.hypot(width, height)
    size = min(120.0, 0.7 * diagonal / max(stringWidth(text, font, 1), 1))
    buffer = io.BytesIO()
    canvas = rl_canvas.Canvas(buffer, pagesize=(width, height))
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#7f939b"))
    canvas.setFillAlpha(0.16)
    canvas.translate(width / 2, height / 2)
    canvas.rotate(math.degrees(math.atan2(height, width)))
    canvas.setFont(font, size)
    canvas.drawCentredString(0, -size / 3, text)
    canvas.restoreState()
    canvas.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]


def finalize_pdf(
    path: Path,
    title: str,
    subject: str,
    watermark: str = "",
    protect: bool = False,
) -> None:
    """Metadati, filigrana opzionale e blocco modifiche opzionale sul PDF finale."""
    from pypdf import PdfReader, PdfWriter, Transformation
    from pypdf.constants import UserAccessPermissions as Perm

    writer = PdfWriter(clone_from=PdfReader(str(path)))
    watermark = watermark.strip()
    if watermark:
        overlays: dict[tuple[int, int], Any] = {}
        for page in writer.pages:
            box = page.mediabox
            key = (round(float(box.width)), round(float(box.height)))
            if key not in overlays:
                overlays[key] = _watermark_page(watermark, float(box.width), float(box.height))
            page.merge_transformed_page(
                overlays[key], Transformation().translate(float(box.left), float(box.bottom))
            )
    writer.add_metadata({
        "/Title": title,
        "/Author": "Formazioni PZZ",
        "/Subject": subject,
        "/Creator": f"Formazioni PZZ {APP_VERSION}",
        "/Producer": f"Formazioni PZZ {APP_VERSION}",
    })
    if protect:
        everything = functools.reduce(lambda a, b: a | b, Perm)
        blocked = Perm.MODIFY | Perm.ADD_OR_MODIFY | Perm.ASSEMBLE_DOC | Perm.FILL_FORM_FIELDS
        # Si apre senza password; la password proprietario casuale impedisce di
        # togliere il blocco. RC4 non richiede librerie crittografiche esterne.
        writer.encrypt(user_password="", owner_password=secrets.token_hex(16),
                       permissions_flag=everything & ~blocked, algorithm="RC4-128")
    temp_path = path.with_name(path.stem + ".finalizing.pdf")
    with temp_path.open("wb") as handle:
        writer.write(handle)
    os.replace(temp_path, path)


def build_pdf(
    output_path: Path,
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
    templates: list[TemplateFile],
    progress_cb=None,
    *,
    watermark: str = "",
    protect: bool = False,
) -> int:
    total = _build_pdf_content(output_path, employee_name, entry_date, department,
                               role, notes, templates, progress_cb)
    _finalize_dossier(output_path, employee_name, entry_date, department, watermark, protect)
    return total


def _finalize_dossier(output_path: Path, employee_name: str, entry_date: str,
                      department: str, watermark: str, protect: bool) -> None:
    try:
        finalize_pdf(output_path, f"Dossier formazione - {employee_name}",
                     f"Reparto: {department} - Ingresso: {entry_date}",
                     watermark=watermark, protect=protect)
    except Exception:
        # Filigrana e protezione richieste esplicitamente: l'errore va mostrato.
        # I soli metadati sono accessori e non devono far fallire il dossier.
        if watermark.strip() or protect:
            raise
        traceback.print_exc()


def _build_job(job: DossierJob, watermark: str, protect: bool) -> None:
    """Un dossier completo; funzione di modulo per poterla eseguire in un altro processo."""
    build_pdf(job.output_path, job.employee_name, job.entry_date, job.department,
              job.role, job.notes, job.templates, watermark=watermark, protect=protect)


def build_pdfs(
    jobs: list[DossierJob],
    progress_cb=None,
    *,
    watermark: str = "",
    protect: bool = False,
) -> list[Exception | None]:
    """Genera piu' dossier. Con Office: una sola sessione per tutto il lotto.
    Senza Office (motore ReportLab): piu' dossier in parallelo su processi separati.
    Ritorna, per ogni dossier, None se riuscito oppure l'errore."""
    results: list[Exception | None] = [None] * len(jobs)
    pending = list(range(len(jobs)))
    if not jobs:
        return results
    if _office_command():
        try:
            native = _build_native_jobs(jobs, progress_cb)
            for index, error in enumerate(native):
                if error is None:
                    job = jobs[index]
                    try:
                        _finalize_dossier(job.output_path, job.employee_name, job.entry_date,
                                          job.department, watermark, protect)
                    except Exception as exc:  # noqa: BLE001
                        results[index] = exc
            pending = [index for index, error in enumerate(native) if error is not None]
        except Exception:  # noqa: BLE001 - conversione di gruppo fallita: uno alla volta
            traceback.print_exc()
        # I falliti si riprovano singolarmente (con il ripiego su ReportLab)
        for done, index in enumerate(pending, start=1):
            try:
                _build_job(jobs[index], watermark, protect)
            except Exception as exc:  # noqa: BLE001
                results[index] = exc
            if progress_cb:
                progress_cb(done, len(pending))
        return results

    finished: set[int] = set()

    def record(index: int, error: Exception | None) -> None:
        results[index] = error
        finished.add(index)
        if progress_cb:
            progress_cb(len(finished), len(jobs))

    workers = min(len(jobs), 4, max(1, (os.cpu_count() or 2) - 1))
    if workers > 1 and len(jobs) >= 3:
        from concurrent.futures import ProcessPoolExecutor, as_completed

        try:
            with ProcessPoolExecutor(max_workers=workers) as pool:
                futures = {pool.submit(_build_job, jobs[i], watermark, protect): i for i in pending}
                for future in as_completed(futures):
                    try:
                        future.result()
                        record(futures[future], None)
                    except Exception as exc:  # noqa: BLE001
                        record(futures[future], exc)
        except Exception:  # noqa: BLE001 - processi non disponibili: si prosegue in sequenza
            traceback.print_exc()
    for index in pending:
        if index in finished:
            continue
        try:
            _build_job(jobs[index], watermark, protect)
            record(index, None)
        except Exception as exc:  # noqa: BLE001
            record(index, exc)
    return results


def _build_pdf_content(
    output_path: Path,
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
    templates: list[TemplateFile],
    progress_cb=None,
) -> int:
    from reportlab.platypus import PageBreak, SimpleDocTemplate, Table

    output_path.parent.mkdir(parents=True, exist_ok=True)
    styles = make_styles()
    expanded = [
        (template, copy_number)
        for template in templates
        for copy_number in range(1, template.copies + 1)
    ]
    needs_legacy_office = any(
        template.path.suffix.lower() in {".doc", ".xls"}
        for template, _ in expanded
    )
    if _office_command():
        # Il motore Office conserva il layout originale dei template (loghi, tabelle, immagini)
        try:
            error = _build_native_jobs([DossierJob(
                output_path, employee_name, entry_date, department, templates, role, notes,
            )], progress_cb)[0]
            if error is None:
                return len(expanded)
            raise error
        except Exception:
            if needs_legacy_office:
                raise
            # .doc/.xls richiedono Office; per gli altri formati si ripiega su ReportLab
            traceback.print_exc()
    has_landscape_pdf = any(
        template.path.suffix.lower() == ".pdf" and _pdf_is_landscape(template.path)
        for template, _ in expanded
    )
    page_size = landscape(A4) if has_landscape_pdf else A4
    page_width, page_height = page_size
    horizontal_frame = min(page_width - 28 * mm, 210 * mm)
    vertical_frame = min(page_height - 28 * mm, 139 * mm)
    h_margin = 10 * mm if has_landscape_pdf else 14 * mm
    v_margin = 10 * mm if has_landscape_pdf else 14 * mm

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=page_size,
        rightMargin=h_margin, leftMargin=h_margin, topMargin=v_margin, bottomMargin=v_margin,
        title=f"Dossier formazione - {employee_name}",
        author="Formazioni PZZ",
    )
    story: list[object] = []
    # Copertina solo in portrait: con PDF landscape si conserva l'impaginazione originale
    if not has_landscape_pdf:
        story.extend(_cover_story(styles, employee_name, entry_date, department, role, notes))
        story.append(PageBreak())
    story_cache_local: dict[Path, list[object]] = {}
    total_steps = max(1, len(expanded) + 1)

    def _fetch_story(template: TemplateFile, step_idx: int) -> list[object]:
        if progress_cb:
            progress_cb(step_idx, total_steps)
        if template.path in story_cache_local:
            return story_cache_local[template.path]
        suffix = template.path.suffix.lower()
        try:
            mtime = template.path.stat().st_mtime_ns
        except OSError:
            mtime = 0
        key = (
            hash(template.path.resolve().as_posix()),
            mtime, employee_name, entry_date, suffix,
            int(horizontal_frame), int(vertical_frame),
        )
        cached_global = _TEMPLATE_STORY_CACHE.get(key)
        if cached_global is not None:
            story_cache_local[template.path] = cached_global
            return cached_global
        if suffix == ".docx":
            generated = docx_story(template.path, employee_name, entry_date, styles)
        elif suffix == ".xlsx":
            generated = xlsx_story(template.path, employee_name, entry_date, styles)
        elif suffix == ".pdf":
            generated = pdf_story(template.path, employee_name, entry_date, styles,
                                  horizontal_frame, vertical_frame)
        else:
            generated = []
        story_cache_local[template.path] = generated
        _TEMPLATE_STORY_CACHE[key] = generated
        return generated

    for index, (template, copy_number) in enumerate(expanded):
        if index > 0:
            story.append(PageBreak())
        cached = _fetch_story(template, index + 1)
        if len(cached) > 3 or isinstance(cached[0] if cached else None, Table):
            story.extend(_copy_mod.copy(cached))
        else:
            story.extend(cached)
    doc.build(story)
    if progress_cb:
        progress_cb(total_steps, total_steps)
    return len(expanded)
