"""Crea un template PDF di esempio (RISORSE UMANE) con text-layer
in modo che l'app possa estrarre *nome* e *data* e sostituirli nel dossier."""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
TEMPL_DIR = ROOT / "templates"
OUT = TEMPL_DIR / "RISORSE UMANE_1_HR.pdf"


def draw(c: canvas.Canvas) -> None:
    w, h = A4
    c.setFont("Helvetica-Bold", 18)
    c.setFillColor(colors.HexColor("#0a2235"))
    c.drawString(22 * mm, h - 28 * mm, "MODULO DI ASSUNZIONE")

    c.setLineWidth(0.6)
    c.setStrokeColor(colors.HexColor("#d9a13f"))
    c.line(22 * mm, h - 31 * mm, w - 22 * mm, h - 31 * mm)

    c.setFont("Helvetica", 10.5)
    c.setFillColor(colors.HexColor("#0a2235"))
    c.drawString(22 * mm, h - 48 * mm, "Dipendente:  *nome*")
    c.drawString(22 * mm, h - 56 * mm, "Data di assunzione:  *data*")
    c.drawString(22 * mm, h - 64 * mm, "Reparto:  RISORSE UMANE")
    c.drawString(22 * mm, h - 72 * mm, "Tipo contratto:  Indeterminato full-time")

    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(colors.HexColor("#1c6262"))
    c.drawString(22 * mm, h - 95 * mm, "Dati retributivi")
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.HexColor("#0a2235"))
    c.drawString(26 * mm, h - 106 * mm, "Livello retributivo:  V livello")
    c.drawString(26 * mm, h - 114 * mm, "RAL mensile:  € __________")
    c.drawString(26 * mm, h - 122 * mm, "Maturazione TFR:  Legale")
    c.drawString(26 * mm, h - 130 * mm, "Giorni ferie annui:  22 gg + 4 ROL")

    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(colors.HexColor("#1c6262"))
    c.drawString(22 * mm, h - 155 * mm, "Firme")
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.HexColor("#0a2235"))
    c.drawString(26 * mm, h - 170 * mm, "Il lavoratore *nome*")
    c.drawString(26 * mm, h - 178 * mm, "Data firma *data*")
    c.drawString(w / 2, h - 170 * mm, "Il Responsabile HR")
    c.drawString(w / 2, h - 178 * mm, "Data firma *data*")

    c.setFont("Helvetica-Oblique", 8.5)
    c.setFillColor(colors.HexColor("#70858b"))
    c.drawCentredString(w / 2, 20 * mm,
                        "Modulo generato in template — i campi *nome* e *data* saranno sostituiti nel dossier finale.")
    c.showPage()


canv = canvas.Canvas(str(OUT), pagesize=A4)
draw(canv)
canv.save()
print(f"Template PDF creato: {OUT} ({OUT.stat().st_size} bytes)")
