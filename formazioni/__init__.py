"""Formazioni PZZ: crea dossier PDF di formazione a partire da template Word, Excel e PDF.

Struttura del pacchetto:
    config       percorsi, versione, impostazioni, lingue, tema
    templates    moduli (template): nomi, reparti, impronte MD5
    documents    Word/Excel e segnaposto *nome* / *data*
    office       conversione con Microsoft Office o LibreOffice
    pdf          motore dei dossier (copertina, unione, filigrana, batch)
    batch_excel  modello Excel per l'import di piu' dipendenti
    updates      controllo e scaricamento degli aggiornamenti
    logs         registro degli errori
    icons, system, tkcompat   supporto all'interfaccia e a Windows
    ui/          interfaccia grafica Tkinter (ui/kit.py: grafica arrotondata)
"""

import importlib

# Le funzioni principali si possono importare da qui (from formazioni import build_pdf).
# Il caricamento e' ritardato: l'installer usa solo formazioni.ui.kit e non deve
# trascinarsi dietro ReportLab, Word ed Excel.
_EXPORTS = {
    "APP_DIR": "config", "APP_VERSION": "config", "HISTORY_FILE": "config",
    "load_language": "config", "load_settings": "config",
    "replace_docx_placeholders": "documents", "replace_placeholders": "documents",
    "DossierJob": "pdf", "build_pdf": "pdf", "build_pdfs": "pdf", "docx_story": "pdf",
    "make_styles": "pdf", "xlsx_story": "pdf",
    "TemplateFile": "templates", "department_options": "templates",
    "discover_templates": "templates", "load_departments_from_file": "templates",
    "parse_template": "templates", "safe_file_part": "templates",
    "templates_for_department": "templates", "templates_for_departments": "templates",
    "FormazioniApp": "ui.app",
}
__all__ = sorted(_EXPORTS)


def __getattr__(name: str):
    if name not in _EXPORTS:
        raise AttributeError(f"module 'formazioni' has no attribute {name!r}")
    return getattr(importlib.import_module(f".{_EXPORTS[name]}", __name__), name)
