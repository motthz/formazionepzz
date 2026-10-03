"""Accesso a Word/Excel (import ritardati) e sostituzione dei segnaposto."""

from __future__ import annotations

import re


# python-docx, openpyxl e reportlab.platypus pesano quasi un secondo all'avvio:
# si importano alla prima generazione, non all'apertura della finestra.
def Document(*args, **kwargs):  # noqa: N802 - stesso nome dell'API di python-docx
    from docx import Document as _document

    return _document(*args, **kwargs)


def load_workbook(*args, **kwargs):
    from openpyxl import load_workbook as _load_workbook

    return _load_workbook(*args, **kwargs)


def replace_placeholders(value: object, employee_name: str, entry_date: str) -> object:
    if not isinstance(value, str):
        return value
    value = re.sub(r"\*nome\*", employee_name, value, flags=re.IGNORECASE)
    return re.sub(r"\*data\*", entry_date, value, flags=re.IGNORECASE)


def replace_paragraph(paragraph, employee_name: str, entry_date: str) -> None:
    original = paragraph.text
    for run in paragraph.runs:
        run.text = str(replace_placeholders(run.text, employee_name, entry_date))
    if paragraph.text != original:
        return
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


def replace_xlsx_placeholders(workbook, employee_name: str, entry_date: str) -> None:
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                try:
                    cell.value = replace_placeholders(cell.value, employee_name, entry_date)
                except AttributeError:
                    pass
