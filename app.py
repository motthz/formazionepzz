"""Formazioni PZZ: crea dossier PDF locali a partire da template Word ed Excel."""

from __future__ import annotations

import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
try:
    import tkinter as tk
    from tkinter import BOTH, END, LEFT, RIGHT, X, BooleanVar, StringVar, filedialog, messagebox, ttk
except ImportError:
    # The PDF engine can still be imported and used on headless machines.
    # The launcher will show a clear message when the desktop GUI is started
    # without the Tk support supplied by the local Python installation.
    tk = None  # type: ignore[assignment]
    BOTH = END = LEFT = RIGHT = X = None  # type: ignore[assignment]
    BooleanVar = StringVar = filedialog = messagebox = ttk = None  # type: ignore[assignment]
from xml.sax.saxutils import escape

from docx import Document
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


# In a PyInstaller one-file build, __file__ points inside a temporary
# extraction directory. Keep user templates, output and history beside the
# executable instead. If the executable lives inside a ./release/ subfolder
# we intentionally use the PARENT folder as the shared workspace so the
# release EXE, the source launcher and the batch file all share the same
# reparti.txt, templates/ and output/ with zero duplication.
def _resolve_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "release":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


APP_DIR = _resolve_app_dir()
DEFAULT_TEMPLATE_DIR = APP_DIR / "templates"
DEFAULT_OUTPUT_DIR = APP_DIR / "output"
HISTORY_FILE = APP_DIR / ".formazioni_history.json"
DEPARTMENTS_FILE = APP_DIR / "reparti.txt"
SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pdf"}
ALL_DEPARTMENT_NAMES = {"TUTTI", "TUTTE", "ALL"}
FILENAME_PATTERN = re.compile(
    r"^(?P<department>.+)_(?P<count>\d+)_(?P<code>[A-Za-z]{2,5})$",
    re.IGNORECASE,
)


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
    """Interpreta REPARTO_NUMERO_CODICE dal nome senza estensione."""
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


def replace_placeholders(value: object, employee_name: str, entry_date: str) -> object:
    if not isinstance(value, str):
        return value
    value = re.sub(r"\*nome\*", employee_name, value, flags=re.IGNORECASE)
    return re.sub(r"\*data\*", entry_date, value, flags=re.IGNORECASE)


def replace_paragraph(paragraph, employee_name: str, entry_date: str) -> None:
    original = paragraph.text
    replaced = replace_placeholders(original, employee_name, entry_date)
    if original == replaced:
        return
    if paragraph.runs:
        paragraph.runs[0].text = str(replaced)
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(str(replaced))


def replace_docx_placeholders(document: Document, employee_name: str, entry_date: str) -> None:
    for paragraph in document.paragraphs:
        replace_paragraph(paragraph, employee_name, entry_date)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph(paragraph, employee_name, entry_date)
    for section in document.sections:
        for container in (section.header, section.footer):
            for paragraph in container.paragraphs:
                replace_paragraph(paragraph, employee_name, entry_date)
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            replace_paragraph(paragraph, employee_name, entry_date)


def paragraph_text(text: str, style: ParagraphStyle) -> Paragraph:
    safe = escape(text).replace("\n", "<br/>")
    return Paragraph(safe or " ", style)


def docx_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
) -> list[object]:
    document = Document(path)
    replace_docx_placeholders(document, employee_name, entry_date)
    story: list[object] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            heading = paragraph.style.name.lower() if paragraph.style else ""
            style = styles["subheading"] if "heading" in heading else styles["body"]
            story.append(paragraph_text(text, style))
            story.append(Spacer(1, 2.4 * mm))
    for table_index, table in enumerate(document.tables, start=1):
        rows: list[list[object]] = []
        for row in table.rows:
            values = [cell.text.strip() for cell in row.cells]
            if any(values):
                rows.append([paragraph_text(value, styles["table"]) for value in values])
        if rows:
            story.append(Spacer(1, 2 * mm))
            story.append(
                Table(
                    rows,
                    repeatRows=1,
                    hAlign="LEFT",
                    style=TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e7eef1")),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#173642")),
                            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#b9c9cf")),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 5),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                            ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                        ]
                    ),
                )
            )
            story.append(Spacer(1, 4 * mm))
    return story or [paragraph_text("Documento senza contenuto testuale.", styles["muted"])]


def xlsx_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
) -> list[object]:
    workbook = load_workbook(path, data_only=False, read_only=False)
    story: list[object] = []
    try:
        for sheet in workbook.worksheets:
            rows: list[list[object]] = []
            for row in sheet.iter_rows():
                values = [
                    replace_placeholders(cell.value, employee_name, entry_date)
                    for cell in row
                ]
                if any(value not in (None, "") for value in values):
                    rows.append(
                        [
                            paragraph_text(str(value) if value is not None else "", styles["table"])
                            for value in values
                        ]
                    )
            if rows:
                story.append(Paragraph(escape(sheet.title), styles["subheading"]))
                story.append(Spacer(1, 2 * mm))
                story.append(
                    Table(
                        rows,
                        repeatRows=1,
                        hAlign="LEFT",
                        style=TableStyle(
                            [
                                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e7eef1")),
                                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#b9c9cf")),
                                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                                ("TOPPADDING", (0, 0), (-1, -1), 4),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                            ]
                        ),
                    )
                )
                story.append(Spacer(1, 5 * mm))
    finally:
        workbook.close()
    return story or [paragraph_text("Foglio senza contenuto.", styles["muted"])]


def make_styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "cover_title": ParagraphStyle(
            "CoverTitle",
            parent=base["Title"],
            fontName="Helvetica-Bold",
            fontSize=25,
            leading=30,
            textColor=colors.HexColor("#173642"),
            alignment=TA_CENTER,
            spaceAfter=8 * mm,
        ),
        "cover_subtitle": ParagraphStyle(
            "CoverSubtitle",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=12,
            leading=17,
            textColor=colors.HexColor("#48636d"),
            alignment=TA_CENTER,
        ),
        "meta": ParagraphStyle(
            "Meta",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=10.5,
            leading=16,
            textColor=colors.HexColor("#173642"),
        ),
        "heading": ParagraphStyle(
            "Heading",
            parent=base["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=15,
            leading=19,
            textColor=colors.HexColor("#173642"),
            spaceAfter=4 * mm,
        ),
        "subheading": ParagraphStyle(
            "Subheading",
            parent=base["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=11.5,
            leading=15,
            textColor=colors.HexColor("#247b7b"),
            spaceBefore=2 * mm,
            spaceAfter=2 * mm,
        ),
        "body": ParagraphStyle(
            "Body",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#263f47"),
        ),
        "table": ParagraphStyle(
            "Table",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=8.2,
            leading=10,
            textColor=colors.HexColor("#263f47"),
        ),
        "muted": ParagraphStyle(
            "Muted",
            parent=base["BodyText"],
            fontName="Helvetica-Oblique",
            fontSize=9.5,
            leading=13,
            textColor=colors.HexColor("#71838a"),
        ),
        "small": ParagraphStyle(
            "Small",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#71838a"),
        ),
        "pdf_page": ParagraphStyle(
            "PdfPage",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=11.5,
            textColor=colors.HexColor("#263f47"),
        ),
    }


def pdf_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
) -> list[object]:
    """Include a PDF template as text pages in the dossier.

    We intentionally do NOT try to binary-merge the source PDF into the
    output because ReportLab's Platypus pipeline builds a single,
    consistent PDF from scratch (cover page + documents + page numbers are
    preserved). Instead we:
      1. Extract any text layer from the source PDF with pypdf
      2. Substitute *nome* / *data* case-insensitively
      3. Emit one Platypus subheading per PDF page followed by the text
    If the source PDF has no extractable text layer we still include a
    page that references the file name so the dossier is complete.
    """
    story: list[object] = []
    try:
        from pypdf import PdfReader
    except Exception:
        story.append(
            paragraph_text(
                "Libreria pypdf non disponibile per leggere il template PDF. Installare con: pip install pypdf",
                styles["muted"],
            )
        )
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
        if len(reader.pages) > 1:
            story.append(Paragraph(f"Pagina {idx} del PDF", styles["subheading"]))
        raw = ""
        try:
            raw = page.extract_text() or ""
        except Exception:
            raw = ""
        raw = replace_placeholders(raw, employee_name, entry_date)
        lines = [ln.rstrip() for ln in raw.splitlines()]
        if not any(ln.strip() for ln in lines):
            story.append(
                paragraph_text(
                    "(Nessun testo estraibile dal PDF. Il file è stato comunque incluso nel dossier come riferimento.)",
                    styles["muted"],
                )
            )
            story.append(Spacer(1, 3 * mm))
            continue
        for line in lines:
            if not line.strip():
                story.append(Spacer(1, 1.5 * mm))
                continue
            story.append(Paragraph(escape(line), styles["pdf_page"]))
        story.append(Spacer(1, 4 * mm))
    return story or [paragraph_text("Documento PDF senza contenuto testuale.", styles["muted"])]


def build_pdf(
    output_path: Path,
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
    templates: list[TemplateFile],
) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    styles = make_styles()
    expanded = [
        (template, copy_number)
        for template in templates
        for copy_number in range(1, template.copies + 1)
    ]
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=17 * mm,
        bottomMargin=17 * mm,
        title=f"Dossier formazione - {employee_name}",
        author="Formazioni PZZ",
    )
    story: list[object] = [
        Spacer(1, 17 * mm),
        Paragraph("Formazioni PZZ", styles["cover_title"]),
        Paragraph("Dossier di inserimento del personale", styles["cover_subtitle"]),
        Spacer(1, 13 * mm),
        HRFlowable(width="70%", thickness=1, color=colors.HexColor("#e4a24c"), hAlign="CENTER"),
        Spacer(1, 13 * mm),
        Table(
            [
                [Paragraph("<b>Nome</b>", styles["meta"]), paragraph_text(employee_name, styles["meta"])],
                [Paragraph("<b>Data di ingresso</b>", styles["meta"]), paragraph_text(entry_date, styles["meta"])],
                [Paragraph("<b>Reparto</b>", styles["meta"]), paragraph_text(department, styles["meta"])],
                [Paragraph("<b>Ruolo</b>", styles["meta"]), paragraph_text(role or "—", styles["meta"])],
                [Paragraph("<b>Documenti</b>", styles["meta"]), paragraph_text(str(len(expanded)), styles["meta"])],
            ],
            colWidths=[44 * mm, 100 * mm],
            hAlign="CENTER",
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f1f6f5")),
                    ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#b9c9cf")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#d6e1e3")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 7),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ]
            ),
        ),
        Spacer(1, 11 * mm),
        Paragraph(
            "Il dossier contiene i documenti applicabili al reparto selezionato "
            "secondo quanto indicato nei file presenti nella cartella dei template.",
            styles["small"],
        ),
    ]
    if notes.strip():
        story.extend([Spacer(1, 7 * mm), Paragraph("<b>Note</b>", styles["subheading"]), paragraph_text(notes, styles["body"])])

    for index, (template, copy_number) in enumerate(expanded):
        if index > 0 or notes.strip() or len(card_story) > 0:
            story.append(PageBreak())
        if template.path.suffix.lower() == ".docx":
            story.extend(docx_story(template.path, employee_name, entry_date, styles))
        elif template.path.suffix.lower() == ".xlsx":
            story.extend(xlsx_story(template.path, employee_name, entry_date, styles))
        elif template.path.suffix.lower() == ".pdf":
            story.extend(pdf_story(template.path, employee_name, entry_date, styles))
    doc.build(story)
    return len(expanded)


def safe_file_part(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9À-ÿ_-]+", "-", value.strip())
    return cleaned.strip("-_") or "persona"


def open_folder(path: Path) -> None:
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except OSError:
        pass


class FormazioniApp:
    def __init__(self, root: tk.Tk) -> None:
        if tk is None or ttk is None:
            raise RuntimeError(
                "Questa installazione di Python non include Tkinter. "
                "Installa Python con il supporto Tk e riavvia Formazioni PZZ."
            )
        self.root = root
        self.root.title("Formazioni PZZ — Dossier Formazione")
        self.root.geometry("1120x820")
        self.root.minsize(980, 720)
        self.root.configure(bg="#eef3f3")

        self.template_dir = StringVar(value=str(DEFAULT_TEMPLATE_DIR))
        self.output_dir = StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.employee_name = StringVar()
        self.entry_date = StringVar()
        self.department = StringVar()
        self.role = StringVar()
        self.notes = StringVar()
        self.auto_open = BooleanVar(value=True)
        self.status = StringVar(value="Scegli una cartella template per iniziare.")
        self.count_label = StringVar(value="Nessun documento rilevato")
        self.templates: list[TemplateFile] = []
        self.ignored: list[Path] = []

        self._configure_style()
        self._build_header()
        self._build_body()
        self.refresh_templates()

    def _configure_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("App.TFrame", background="#eef3f3")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("CardBody.TFrame", background="#ffffff")
        style.configure("Shadow1.TFrame", background="#dce4e5")
        style.configure("Shadow2.TFrame", background="#e6ecec")

        style.configure(
            "Title.TLabel",
            background="#0a2235",
            foreground="#ffffff",
            font=("Segoe UI Semibold", 26, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background="#0a2235",
            foreground="#b7d0d8",
            font=("Segoe UI", 10),
        )
        style.configure(
            "Eyebrow.TLabel",
            background="#0a2235",
            foreground="#d9a13f",
            font=("Segoe UI", 8, "bold"),
        )

        style.configure(
            "Section.TLabel",
            background="#ffffff",
            foreground="#0f2a36",
            font=("Segoe UI Semibold", 13, "bold"),
        )
        style.configure(
            "SectionAccent.TLabel",
            background="#ffffff",
            foreground="#d9a13f",
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "Muted.TLabel",
            background="#ffffff",
            foreground="#56707a",
            font=("Segoe UI", 9),
        )
        style.configure(
            "AppMuted.TLabel",
            background="#eef3f3",
            foreground="#56707a",
            font=("Segoe UI", 9),
        )

        style.configure(
            "FieldLabel.TLabel",
            background="#ffffff",
            foreground="#0f2a36",
            font=("Segoe UI Semibold", 9, "bold"),
        )

        style.configure(
            "Count.TLabel",
            background="#e3f1ee",
            foreground="#1c6262",
            font=("Segoe UI Semibold", 9, "bold"),
            padding=(12, 6),
        )
        style.configure(
            "Gold.TLabel",
            background="#fbf0dc",
            foreground="#a8721c",
            font=("Segoe UI Semibold", 9, "bold"),
            padding=(12, 6),
        )
        style.configure(
            "Secure.TLabel",
            background="#e3f1ee",
            foreground="#1c6262",
            font=("Segoe UI Semibold", 9, "bold"),
            padding=(12, 6),
        )

        style.configure(
            "Primary.TButton",
            background="#2a8e8e",
            foreground="#ffffff",
            font=("Segoe UI Semibold", 10, "bold"),
            padding=(18, 11),
            borderwidth=0,
            focusthickness=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1d6a6a"), ("pressed", "#154d4d")],
        )

        style.configure(
            "Secondary.TButton",
            background="#eaf0f0",
            foreground="#0f2a36",
            font=("Segoe UI Semibold", 9, "bold"),
            padding=(14, 9),
            borderwidth=0,
            focusthickness=0,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#d5e0e1"), ("pressed", "#c2d2d3")],
        )

        style.configure(
            "Accent.TButton",
            background="#d9a13f",
            foreground="#2a1a00",
            font=("Segoe UI Semibold", 9, "bold"),
            padding=(14, 9),
            borderwidth=0,
            focusthickness=0,
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#c08c2d"), ("pressed", "#a4741d")],
        )

        style.configure(
            "TEntry",
            fieldbackground="#ffffff",
            foreground="#0f2a36",
            bordercolor="#cfe1e4",
            lightcolor="#cfe1e4",
            darkcolor="#cfe1e4",
            padding=9,
            focusthickness=2,
            focuscolor="#2a8e8e",
            font=("Segoe UI", 10),
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", "#2a8e8e"), ("!focus", "#cfe1e4")],
            lightcolor=[("focus", "#2a8e8e"), ("!focus", "#cfe1e4")],
            darkcolor=[("focus", "#2a8e8e"), ("!focus", "#cfe1e4")],
        )

        style.configure(
            "TCombobox",
            fieldbackground="#ffffff",
            foreground="#0f2a36",
            background="#ffffff",
            arrowcolor="#1c6262",
            bordercolor="#cfe1e4",
            lightcolor="#cfe1e4",
            darkcolor="#cfe1e4",
            padding=8,
            focusthickness=2,
            focuscolor="#2a8e8e",
            font=("Segoe UI", 10),
        )
        style.map(
            "TCombobox",
            bordercolor=[("focus", "#2a8e8e"), ("!focus", "#cfe1e4")],
            lightcolor=[("focus", "#2a8e8e"), ("!focus", "#cfe1e4")],
            darkcolor=[("focus", "#2a8e8e"), ("!focus", "#cfe1e4")],
            background=[("readonly", "#ffffff"), ("active", "#ffffff")],
            fieldbackground=[("readonly", "#ffffff")],
        )

        style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#0f2a36",
            rowheight=28,
            bordercolor="#cfe1e4",
            borderwidth=1,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Treeview.Heading",
            background="#1c6262",
            foreground="#ffffff",
            font=("Segoe UI Semibold", 9, "bold"),
            padding=9,
            relief="flat",
            borderwidth=0,
        )
        style.map(
            "Treeview",
            background=[
                ("selected", "#2a8e8e"),
            ],
            foreground=[
                ("selected", "#ffffff"),
            ],
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#1d6a6a")],
        )

        style.configure(
            "Vertical.TScrollbar",
            background="#cfe1e4",
            troughcolor="#f2f7f7",
            bordercolor="#cfe1e4",
            arrowcolor="#1c6262",
            arrowsize=14,
            relief="flat",
            borderwidth=0,
            gripcount=0,
            width=14,
        )
        style.map(
            "Vertical.TScrollbar",
            background=[
                ("active", "#2a8e8e"),
                ("disabled", "#e6ecec"),
            ],
        )

        style.configure(
            "TCheckbutton",
            background="#ffffff",
            foreground="#0f2a36",
            font=("Segoe UI Semibold", 9),
            focusthickness=0,
        )
        style.map(
            "TCheckbutton",
            background=[("active", "#ffffff")],
            foreground=[("active", "#0f2a36")],
        )

    def _draw_gradient(self, canvas, w, h, color1, color2):
        steps = max(1, h)
        r1, g1, b1 = canvas.winfo_rgb(color1)
        r2, g2, b2 = canvas.winfo_rgb(color2)
        ratio_r = (r2 - r1) / steps
        ratio_g = (g2 - g1) / steps
        ratio_b = (b2 - b1) / steps
        for i in range(steps):
            nr = int(r1 + ratio_r * i)
            ng = int(g1 + ratio_g * i)
            nb = int(b1 + ratio_b * i)
            color = f"#{nr:04x}{ng:04x}{nb:04x}"
            canvas.create_line(0, i, w, i, fill=color)

    def _build_header(self) -> None:
        header_outer = tk.Frame(self.root, bg="#eef3f3", height=160)
        header_outer.pack(fill=X, side="top")
        header_outer.pack_propagate(False)

        c = tk.Canvas(header_outer, height=148, highlightthickness=0, bd=0, bg="#0a2235")
        c.pack(fill=X, side="top")

        def paint_header(_evt=None):
            w = max(c.winfo_width(), 1)
            c.delete("all")
            self._draw_gradient(c, w, 148, "#0a2235", "#153a52")
            c.create_rectangle(0, 146, w, 148, fill="#d9a13f", outline="")

        c.bind("<Configure>", paint_header)
        c.after(1, paint_header)

        content = tk.Frame(header_outer, bg="#0a2235")
        content.place(x=42, y=28, relwidth=1.0, width=-84)

        ttk.Label(content, text="FORMAZIONI  ·  PZZ", style="Eyebrow.TLabel").pack(anchor="w")
        ttk.Label(content, text="Dossier di formazione", style="Title.TLabel").pack(anchor="w", pady=(4, 0))
        ttk.Label(
            content,
            text="Crea dossier PDF pronti per la stampa, partendo dai template Word / Excel / PDF dei singoli reparti.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(6, 0))

    def _card(self, parent, title=None, subtitle=None, accent=None, **kwargs):
        shadow1 = tk.Frame(parent, bg="#dce4e5", highlightthickness=0)
        shadow2 = tk.Frame(shadow1, bg="#e6ecec", padx=1, pady=1)
        shadow2.pack(fill=BOTH, expand=True, padx=2, pady=2)
        card = tk.Frame(shadow2, bg="#ffffff", padx=2, pady=2)
        card.pack(fill=BOTH, expand=True, padx=1, pady=1)
        inner = tk.Frame(card, bg="#ffffff", padx=26, pady=24)
        inner.pack(fill=BOTH, expand=True)

        head = tk.Frame(inner, bg="#ffffff")
        head.pack(fill=X)
        if title or accent:
            left = tk.Frame(head, bg="#ffffff")
            left.pack(side=LEFT, fill=X, expand=True)
            if accent:
                ttk.Label(left, text=accent, style="SectionAccent.TLabel").pack(anchor="w")
            if title:
                ttk.Label(left, text=title, style="Section.TLabel").pack(anchor="w", pady=(2, 0))
            tk.Frame(inner, bg="#d9a13f", height=2).pack(fill=X, pady=(16, 22))

        if subtitle:
            ttk.Label(inner, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(0, 16))

        body = tk.Frame(inner, bg="#ffffff")
        body.pack(fill=BOTH, expand=True)
        return shadow1, body

    def _labeled_field(self, parent, label_text, widget_factory, row, col=0, span=2, label_col_width=160, **w_kwargs):
        lbl = tk.Label(
            parent,
            text=label_text,
            bg="#ffffff",
            fg="#0f2a36",
            font=("Segoe UI Semibold", 9, "bold"),
            anchor="w",
        )
        lbl.grid(row=row, column=col, sticky="we", padx=(0, 16), pady=(0, 6))
        parent.grid_columnconfigure(col, minsize=label_col_width)
        container = tk.Frame(parent, bg="#ffffff")
        container.grid(row=row, column=col + 1, sticky="nsew", columnspan=span - 1, pady=(0, 14))
        container.grid_columnconfigure(0, weight=1)
        w_kwargs.setdefault("master", container)
        widget = widget_factory(**w_kwargs)
        widget.grid(row=0, column=0, sticky="ew")
        return widget

    def _build_body(self) -> None:
        outer_wrap = ttk.Frame(self.root, style="App.TFrame")
        outer_wrap.pack(fill=BOTH, expand=True)
        outer_wrap.columnconfigure(0, weight=1)
        outer_wrap.rowconfigure(0, weight=1)

        canvas_wrap = tk.Canvas(outer_wrap, background="#eef3f3", highlightthickness=0, borderwidth=0)
        canvas_wrap.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(outer_wrap, orient="vertical", command=canvas_wrap.yview)
        sb.grid(row=0, column=1, sticky="ns")
        canvas_wrap.configure(yscrollcommand=sb.set)

        scrolled = ttk.Frame(canvas_wrap, style="App.TFrame")
        scrolled_id = canvas_wrap.create_window((0, 0), window=scrolled, anchor="nw")

        def _on_scroll_config(_evt=None):
            canvas_wrap.configure(scrollregion=canvas_wrap.bbox("all"))

        def _on_canvas_config(evt):
            canvas_wrap.itemconfigure(scrolled_id, width=evt.width)

        scrolled.bind("<Configure>", _on_scroll_config)
        canvas_wrap.bind("<Configure>", _on_canvas_config)
        canvas_wrap.bind_all(
            "<MouseWheel>",
            lambda evt: canvas_wrap.yview_scroll(int(-1 * (evt.delta / 120)), "units"),
            add="+",
        )

        body = ttk.Frame(scrolled, style="App.TFrame", padding=(30, 6, 30, 10))
        body.pack(fill=BOTH, expand=True)
        body.columnconfigure(0, weight=5)
        body.columnconfigure(1, weight=4)
        body.rowconfigure(1, weight=1)

        # ==== CARD 1: SORGENTI ====
        src_shadow, src_body = self._card(
            body,
            title="Cartelle locali",
            accent="01  ·  CONFIGURAZIONE",
            subtitle="Indica dove si trovano i template Word / Excel / PDF e dove salvare i PDF generati.",
        )
        src_shadow.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 18))
        src_body.columnconfigure(1, weight=1)
        src_body.columnconfigure(2, weight=0)

        def add_row(r, label_text, var, btn_cmd):
            tk.Label(
                src_body, text=label_text, bg="#ffffff", fg="#0f2a36",
                font=("Segoe UI Semibold", 9, "bold"), anchor="w",
            ).grid(row=r, column=0, sticky="we", padx=(0, 18), pady=(0, 4))
            ent = ttk.Entry(src_body, textvariable=var)
            ent.grid(row=r, column=1, sticky="ew", padx=(0, 10), pady=(0, 14))
            ttk.Button(src_body, text="Scegli cartella", style="Secondary.TButton", command=btn_cmd).grid(
                row=r, column=2, sticky="ew", pady=(0, 14)
            )

        add_row(0, "Template Word / Excel / PDF (.docx, .xlsx, .pdf)", self.template_dir, self.choose_template_dir)
        add_row(1, "Destinazione PDF generati", self.output_dir, self.choose_output_dir)

        action_row = tk.Frame(src_body, bg="#ffffff")
        action_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Button(
            action_row,
            text="⟳  Aggiorna documenti",
            style="Accent.TButton",
            command=self.refresh_templates,
        ).pack(side=LEFT)

        # ==== CARD 2: FORM ====
        form_shadow, form_body = self._card(
            body,
            title="Nuovo dossier",
            accent="02  ·  ANAGRAFICA",
            subtitle="Inserisci i dati della persona. I campi con * sono obbligatori.",
        )
        form_shadow.grid(row=1, column=0, sticky="nsew", padx=(0, 18))
        form_body.columnconfigure(1, weight=1)

        self._labeled_field(
            form_body,
            "Nome e cognome *",
            lambda **kw: ttk.Entry(**kw),
            row=0,
            textvariable=self.employee_name,
        )
        self._labeled_field(
            form_body,
            "Data di ingresso *",
            lambda **kw: ttk.Entry(**kw),
            row=1,
            textvariable=self.entry_date,
        )
        self._labeled_field(
            form_body,
            "Ruolo / Mansione (facoltativo)",
            lambda **kw: ttk.Entry(**kw),
            row=2,
            textvariable=self.role,
        )

        tk.Label(
            form_body,
            text="Note (facoltative)",
            bg="#ffffff",
            fg="#0f2a36",
            font=("Segoe UI Semibold", 9, "bold"),
            anchor="w",
        ).grid(row=3, column=0, sticky="nwe", padx=(0, 16), pady=(0, 6))
        notes_wrap = tk.Frame(form_body, bg="#ffffff")
        notes_wrap.grid(row=3, column=1, sticky="nsew", pady=(0, 14))
        notes_wrap.grid_columnconfigure(0, weight=1)
        notes_entry = tk.Text(
            notes_wrap,
            height=7,
            wrap="word",
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#0f2a36",
            relief="solid",
            borderwidth=1,
            highlightthickness=2,
            highlightbackground="#cfe1e4",
            highlightcolor="#2a8e8e",
            padx=10,
            pady=9,
            insertbackground="#2a8e8e",
        )
        notes_entry.grid(row=0, column=0, sticky="nsew")
        self.notes_widget = notes_entry
        sb_notes = ttk.Scrollbar(notes_wrap, orient="vertical", command=notes_entry.yview)
        sb_notes.grid(row=0, column=1, sticky="ns")
        notes_entry.configure(yscrollcommand=sb_notes.set)

        form_body.rowconfigure(3, weight=1)

        bottom_form = tk.Frame(form_body, bg="#ffffff")
        bottom_form.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Checkbutton(
            bottom_form,
            text="Apri la cartella di destinazione al termine",
            variable=self.auto_open,
        ).pack(side=LEFT)

        # ==== CARD 3: PREVIEW DOCUMENTI ====
        prev_shadow, prev_body = self._card(
            body,
            title="Documenti inclusi",
            accent="03  ·  RIEPILOGO",
            subtitle="Scegli il reparto → vedi l'elenco dei documenti → genera il PDF.",
        )
        prev_shadow.grid(row=1, column=1, sticky="nsew")
        prev_body.columnconfigure(0, weight=1)

        # --- Reparto selector (PRIMO, sempre visibile) ---
        dept_frame = tk.Frame(prev_body, bg="#ffffff")
        dept_frame.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        dept_frame.columnconfigure(1, weight=1)
        tk.Label(
            dept_frame,
            text="Reparto *",
            bg="#ffffff",
            fg="#0a2235",
            font=("Segoe UI Semibold", 9, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 14), pady=(0, 6))
        self.department_combo = ttk.Combobox(
            dept_frame,
            state="readonly",
            textvariable=self.department,
            font=("Segoe UI Semibold", 10, "bold"),
            height=14,
        )
        self.department_combo.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        self.department_combo.bind("<<ComboboxSelected>>", lambda _e: self.update_document_list())
        tk.Label(
            dept_frame,
            text="  La selezione aggiorna l'elenco qui sotto in tempo reale.",
            bg="#ffffff",
            fg="#56707a",
            font=("Segoe UI", 8),
            anchor="w",
        ).grid(row=1, column=0, columnspan=2, sticky="w")

        badge_row = tk.Frame(prev_body, bg="#ffffff")
        badge_row.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(badge_row, textvariable=self.count_label, style="Count.TLabel").pack(side=LEFT)

        tree_wrap = tk.Frame(prev_body, bg="#ffffff")
        tree_wrap.grid(row=2, column=0, sticky="nsew", pady=(0, 4))
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)
        prev_body.rowconfigure(2, weight=1)

        self.tree = ttk.Treeview(
            tree_wrap,
            columns=("documento", "copie"),
            show="headings",
            height=10,
        )
        self.tree.heading("documento", text="  Template · Reparto")
        self.tree.heading("copie", text="Copie")
        self.tree.column("documento", width=320, anchor="w")
        self.tree.column("copie", width=80, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.tag_configure("odd", background="#ffffff")
        self.tree.tag_configure("even", background="#f3faf8")
        self.tree.tag_configure("tutti", background="#fff8ea")

        # --- Generate button (PRIMARY, bottom-right card 3, SEMPRE visibile) ---
        gen_frame = tk.Frame(prev_body, bg="#ffffff")
        gen_frame.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        ttk.Button(
            gen_frame,
            text="⬇  Genera PDF unico",
            style="Primary.TButton",
            command=self.generate,
        ).pack(side=RIGHT, fill=X, expand=True, ipadx=18, ipady=5)
        tk.Label(
            gen_frame,
            text="1 clic → dossier PDF",
            bg="#ffffff",
            fg="#56707a",
            font=("Segoe UI", 8),
            anchor="w",
        ).pack(side=LEFT, padx=(4, 0))

        help_frame = tk.Frame(prev_body, bg="#ffffff")
        help_frame.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        tk.Frame(help_frame, bg="#e3f1ee", width=4, height=56).pack(side=LEFT)
        tip = tk.Label(
            help_frame,
            text=(
                "Suggerimento: il nome di ogni template determina reparto e numero di copie.\n"
                "Formato standard:  REPARTO_NUMERO_CODICE.docx  (o .xlsx / .pdf)\n"
                "Esempio:  SICUREZZA_2_SIC.docx   (2 copie per ogni dipendente)"
            ),
            bg="#ffffff",
            fg="#56707a",
            font=("Segoe UI", 9),
            justify="left",
            anchor="w",
            padx=12,
            pady=8,
            wraplength=380,
        )
        tip.pack(side=LEFT, fill=X, expand=True)

        # ==== FOOTER ====
        footer_outer = tk.Frame(self.root, bg="#eef3f3")
        footer_outer.pack(fill=X, side="bottom")
        footer = tk.Frame(footer_outer, bg="#ffffff", padx=24, pady=12)
        footer.pack(fill=X, padx=30, pady=(0, 16))

        status_wrap = tk.Frame(footer, bg="#ffffff")
        status_wrap.pack(side=LEFT, fill=X, expand=True)
        tk.Label(
            status_wrap,
            text="●",
            bg="#ffffff",
            fg="#2a8e8e",
            font=("Segoe UI", 10, "bold"),
        ).pack(side=LEFT)
        tk.Label(
            status_wrap,
            textvariable=self.status,
            bg="#ffffff",
            fg="#0f2a36",
            font=("Segoe UI Semibold", 9),
            anchor="w",
            padx=8,
        ).pack(side=LEFT, fill=X, expand=True)

        ttk.Label(footer, text="🔒  Tutto resta sul computer · nessun dato inviato in rete", style="Secure.TLabel").pack(side=RIGHT)

    def choose_template_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.template_dir.get(), title="Scegli la cartella dei template")
        if chosen:
            self.template_dir.set(chosen)
            self.refresh_templates()

    def choose_output_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.output_dir.get(), title="Scegli la cartella dei PDF")
        if chosen:
            self.output_dir.set(chosen)

    def refresh_templates(self) -> None:
        folder = Path(self.template_dir.get()).expanduser()
        self.templates, self.ignored = discover_templates(folder)
        departments = department_options(self.templates)
        self.department_combo["values"] = departments
        if departments and self.department.get().upper() not in departments:
            self.department.set(departments[0])
        elif not departments:
            self.department.set("")
        self.update_document_list()
        if not folder.exists():
            self.status.set("La cartella template non esiste ancora: puoi crearla o sceglierne un'altra.")
        elif not self.templates:
            self.status.set("Nessun template valido. Usa il formato REPARTO_NUMERO_CODICE.")
        elif self.ignored:
            self.status.set(f"{len(self.templates)} template trovati; {len(self.ignored)} file ignorati.")
        else:
            self.status.set(f"{len(self.templates)} template pronti.")

    def update_document_list(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        selected = templates_for_department(self.templates, self.department.get())
        total = sum(template.copies for template in selected)
        self.count_label.set(f"  {total} DOCUMENTI  ·  {len(selected)} TEMPLATE  ")
        for idx, template in enumerate(selected):
            scope = "Tutti i reparti" if template.is_for_every_department else template.department.upper()
            if template.is_for_every_department:
                tag = "tutti"
            elif idx % 2 == 0:
                tag = "even"
            else:
                tag = "odd"
            self.tree.insert("", END, values=(f"  {scope} · {template.path.name}", template.copies), tags=(tag,))

    def generate(self) -> None:
        name = self.employee_name.get().strip()
        entry_date = self.entry_date.get().strip()
        department = self.department.get().strip().upper()
        if not name:
            messagebox.showwarning("Dati mancanti", "Inserisci nome e cognome.")
            return
        if not entry_date:
            messagebox.showwarning("Dati mancanti", "Inserisci la data di ingresso.")
            return
        if not department:
            messagebox.showwarning("Dati mancanti", "Scegli un reparto.")
            return
        selected = templates_for_department(self.templates, department)
        if not selected:
            messagebox.showwarning("Nessun documento", "Non ci sono template per il reparto selezionato.")
            return
        output_dir = Path(self.output_dir.get()).expanduser()
        base = output_dir / f"dossier_{safe_file_part(name)}.pdf"
        output_path = base
        counter = 2
        while output_path.exists():
            output_path = output_dir / f"dossier_{safe_file_part(name)}_{counter}.pdf"
            counter += 1
        notes = self.notes_widget.get("1.0", END).strip()
        try:
            self.status.set("Creo il PDF…")
            self.root.update_idletasks()
            total = build_pdf(output_path, name, entry_date, department, self.role.get().strip(), notes, selected)
            self._save_history(output_path, name, department, total)
            self.status.set(f"PDF creato: {output_path.name}")
            messagebox.showinfo("PDF pronto", f"Il dossier è stato creato con {total} documenti.\n\n{output_path}")
            if self.auto_open.get():
                open_folder(output_path.parent)
        except Exception as error:
            self.status.set("Errore durante la generazione.")
            traceback.print_exc()
            messagebox.showerror("Impossibile creare il PDF", str(error))

    def _save_history(self, output_path: Path, name: str, department: str, total: int) -> None:
        history: list[dict[str, object]] = []
        if HISTORY_FILE.exists():
            try:
                history = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                history = []
        history.insert(
            0,
            {
                "path": str(output_path),
                "name": name,
                "department": department,
                "documents": total,
            },
        )
        HISTORY_FILE.write_text(json.dumps(history[:20], ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> None:
    if tk is None:
        raise SystemExit(
            "Tkinter non è disponibile in Python. Installa una versione di Python "
            "con il supporto Tk per avviare l'interfaccia desktop."
        )
    root = tk.Tk()
    FormazioniApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()