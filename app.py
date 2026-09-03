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
# executable instead.
APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
DEFAULT_TEMPLATE_DIR = APP_DIR / "templates"
DEFAULT_OUTPUT_DIR = APP_DIR / "output"
HISTORY_FILE = APP_DIR / ".formazioni_history.json"
SUPPORTED_EXTENSIONS = {".docx", ".xlsx"}
ALL_DEPARTMENT_NAMES = {"TUTTI", "TUTTE", "ALL"}
FILENAME_PATTERN = re.compile(
    r"^(?P<department>.+)_(?P<count>\d+)_(?P<code>[A-Za-z]{3})$",
    re.IGNORECASE,
)


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
    names = {
        template.department.upper()
        for template in templates
        if not template.is_for_every_department
    }
    return sorted(names)


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
    }


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
            f"Generato il {datetime.now().strftime('%d/%m/%Y alle %H:%M')}. "
            "Il dossier contiene i documenti applicabili al reparto selezionato.",
            styles["small"],
        ),
    ]
    if notes.strip():
        story.extend([Spacer(1, 7 * mm), Paragraph("<b>Note</b>", styles["subheading"]), paragraph_text(notes, styles["body"])])

    for index, (template, copy_number) in enumerate(expanded):
        story.extend(
            [
                PageBreak(),
                Paragraph(escape(template.path.stem.replace("_", " ")), styles["heading"]),
                Paragraph(
                    f"File: {escape(template.display_name)} · Copia {copy_number} di {template.copies}",
                    styles["small"],
                ),
                Spacer(1, 5 * mm),
            ]
        )
        if template.path.suffix.lower() == ".docx":
            story.extend(docx_story(template.path, employee_name, entry_date, styles))
        elif template.path.suffix.lower() == ".xlsx":
            story.extend(xlsx_story(template.path, employee_name, entry_date, styles))
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
        self.root.title("Formazioni PZZ")
        self.root.geometry("1040x760")
        self.root.minsize(820, 620)
        self.root.configure(bg="#f4f7f6")

        self.template_dir = StringVar(value=str(DEFAULT_TEMPLATE_DIR))
        self.output_dir = StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.employee_name = StringVar()
        self.entry_date = StringVar(value=date.today().strftime("%d/%m/%Y"))
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
        style.configure("App.TFrame", background="#f4f7f6")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#173642", foreground="#ffffff", font=("Helvetica", 21, "bold"))
        style.configure("Subtitle.TLabel", background="#173642", foreground="#cfe0de", font=("Helvetica", 10))
        style.configure("Section.TLabel", background="#ffffff", foreground="#173642", font=("Helvetica", 12, "bold"))
        style.configure("Muted.TLabel", background="#ffffff", foreground="#70858b", font=("Helvetica", 9))
        style.configure("Count.TLabel", background="#eaf4f1", foreground="#247b7b", font=("Helvetica", 9, "bold"))
        style.configure("Primary.TButton", background="#247b7b", foreground="#ffffff", font=("Helvetica", 10, "bold"), padding=(13, 8))
        style.map("Primary.TButton", background=[("active", "#1c6262")])
        style.configure("Secondary.TButton", background="#edf3f2", foreground="#173642", padding=(10, 7))
        style.configure("TEntry", padding=7)
        style.configure("TCombobox", padding=6)

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg="#173642", padx=28, pady=22)
        header.pack(fill=X)
        ttk.Label(header, text="Formazioni PZZ", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Dossier di formazione pronti da stampare, partendo dai tuoi template locali.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

    def _card(self, parent, **kwargs):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=18, **kwargs)
        frame.configure()
        return frame

    def _build_body(self) -> None:
        body = ttk.Frame(self.root, style="App.TFrame", padding=22)
        body.pack(fill=BOTH, expand=True)
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(1, weight=1)

        source = self._card(body)
        source.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        source.columnconfigure(1, weight=1)
        ttk.Label(source, text="Cartelle locali", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(source, text="Template Word / Excel", style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(source, textvariable=self.template_dir).grid(row=1, column=1, sticky="ew", padx=10, pady=(12, 0))
        ttk.Button(source, text="Scegli", style="Secondary.TButton", command=self.choose_template_dir).grid(row=1, column=2, pady=(12, 0))
        ttk.Label(source, text="Destinazione PDF", style="Muted.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(source, textvariable=self.output_dir).grid(row=2, column=1, sticky="ew", padx=10, pady=(8, 0))
        ttk.Button(source, text="Scegli", style="Secondary.TButton", command=self.choose_output_dir).grid(row=2, column=2, pady=(8, 0))
        ttk.Button(source, text="Aggiorna documenti", style="Secondary.TButton", command=self.refresh_templates).grid(row=3, column=1, sticky="w", pady=(12, 0))

        form = self._card(body)
        form.grid(row=1, column=0, sticky="nsew", padx=(0, 14))
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text="Nuovo dossier", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(form, text="Nome e cognome *", style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(18, 0))
        ttk.Entry(form, textvariable=self.employee_name).grid(row=1, column=1, sticky="ew", padx=(18, 0), pady=(18, 0))
        ttk.Label(form, text="Data di ingresso *", style="Muted.TLabel").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(form, textvariable=self.entry_date).grid(row=2, column=1, sticky="ew", padx=(18, 0), pady=(12, 0))
        ttk.Label(form, text="Reparto *", style="Muted.TLabel").grid(row=3, column=0, sticky="w", pady=(12, 0))
        self.department_combo = ttk.Combobox(form, textvariable=self.department, state="readonly")
        self.department_combo.grid(row=3, column=1, sticky="ew", padx=(18, 0), pady=(12, 0))
        self.department_combo.bind("<<ComboboxSelected>>", lambda _event: self.update_document_list())
        ttk.Label(form, text="Ruolo (facoltativo)", style="Muted.TLabel").grid(row=4, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(form, textvariable=self.role).grid(row=4, column=1, sticky="ew", padx=(18, 0), pady=(12, 0))
        ttk.Label(form, text="Note (facoltative)", style="Muted.TLabel").grid(row=5, column=0, sticky="nw", pady=(12, 0))
        notes_entry = tk.Text(form, height=4, width=30, font=("Helvetica", 10), relief="solid", borderwidth=1, highlightthickness=0)
        notes_entry.grid(row=5, column=1, sticky="ew", padx=(18, 0), pady=(12, 0))
        self.notes_widget = notes_entry
        ttk.Checkbutton(form, text="Apri la cartella al termine", variable=self.auto_open).grid(row=6, column=1, sticky="w", pady=(13, 0))
        ttk.Button(form, text="Genera PDF unico", style="Primary.TButton", command=self.generate).grid(row=7, column=1, sticky="ew", pady=(20, 0))

        preview = self._card(body)
        preview.grid(row=1, column=1, sticky="nsew")
        preview.rowconfigure(2, weight=1)
        preview.columnconfigure(0, weight=1)
        ttk.Label(preview, text="Documenti inclusi", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(preview, textvariable=self.count_label, style="Count.TLabel", padding=(8, 4)).grid(row=1, column=0, sticky="w", pady=(10, 10))
        tree_frame = ttk.Frame(preview, style="Card.TFrame")
        tree_frame.grid(row=2, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_frame, columns=("documento", "copie"), show="headings", height=14)
        self.tree.heading("documento", text="Template")
        self.tree.heading("copie", text="Copie")
        self.tree.column("documento", width=220, anchor="w")
        self.tree.column("copie", width=60, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        ttk.Label(preview, text="Il numero nel nome file determina quante copie vengono inserite.", style="Muted.TLabel", wraplength=290).grid(row=3, column=0, sticky="w", pady=(12, 0))

        footer = ttk.Frame(self.root, style="App.TFrame", padding=(22, 0, 22, 14))
        footer.pack(fill=X)
        ttk.Label(footer, textvariable=self.status, style="Muted.TLabel").pack(side=LEFT)
        ttk.Label(footer, text="Tutto resta sul computer", style="Count.TLabel", padding=(8, 4)).pack(side=RIGHT)

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
        self.count_label.set(f"{total} documenti · {len(selected)} template")
        for template in selected:
            scope = "Tutti i reparti" if template.is_for_every_department else template.department.upper()
            self.tree.insert("", END, values=(f"{scope} · {template.path.name}", template.copies))

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
        output_path = output_dir / f"dossier_{safe_file_part(name)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
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
                "created_at": datetime.now().isoformat(timespec="seconds"),
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