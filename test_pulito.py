import app as a
from pathlib import Path
import tempfile
import re
import os
from pypdf import PdfReader
import tkinter as tk
from tkinter import messagebox

root_dir = Path(a.DEPARTMENTS_FILE).parent

os.startfile = lambda *x, **k: None
for m in ("showinfo", "showwarning", "showerror"):
    setattr(messagebox, m, lambda *a, **k: None)

with tempfile.TemporaryDirectory(dir=root_dir.as_posix()) as td:
    td_p = Path(td)
    out_dir = td_p / "out"
    out_dir.mkdir()

    root = tk.Tk()
    root.withdraw()
    g = a.FormazioniApp(root)
    assert g.entry_date.get() == "", f"entry_date NON vuota: {repr(g.entry_date.get())}"
    print("[OK] entry_date = VUOTA (nessuna data odierna precompilata)")

    sf = a.safe_file_part("Mario Rossi")
    assert sf == "Mario-Rossi", f"safe_file_part sbagliato: {sf}"
    print(f"[OK] safe_file_part = {sf}")
    base = out_dir / f"dossier_{sf}.pdf"
    assert base.name == "dossier_Mario-Rossi.pdf"
    print(f"[OK] nome file base (no timestamp) = {base.name}")
    base.write_bytes(b"%PDF-1.4 fake")
    candidate = base
    c = 2
    while candidate.exists():
        candidate = out_dir / f"dossier_{sf}_{c}.pdf"
        c += 1
    assert candidate.name == "dossier_Mario-Rossi_2.pdf", f"Conflitto: {candidate.name}"
    print(f"[OK] In caso di conflitto -> {candidate.name}")

    tmpl_list, _ignored = a.discover_templates(root_dir / "templates")
    sel = a.templates_for_department(tmpl_list, "AMMINISTRAZIONE")
    pdf_p = out_dir / "test-pulito.pdf"
    total = a.build_pdf(pdf_p, "Mario Rossi", "15/07/2025", "AMMINISTRAZIONE", "Impiegato", "Nota: inserimento manuale.", sel)
    assert pdf_p.exists() and pdf_p.stat().st_size > 1000, "PDF non generato"
    print(f"[OK] PDF generato: {pdf_p.stat().st_size} bytes, documenti {total}")

    r = PdfReader(str(pdf_p))
    full = " ".join(p.extract_text() or "" for p in r.pages)
    pattern_today = re.compile(r"04/09/2026|04-09-2026|20260904|Generato il")
    m_found = pattern_today.search(full)
    assert not m_found, f"TROVATO dato generato nel PDF: {repr(m_found.group())}"
    print("[OK] NESSUNA data odierna o testo 'Generato il' nel PDF.")

    assert "Mario Rossi" in full
    assert "15/07/2025" in full
    assert "AMMINISTRAZIONE" in full.upper()
    assert "Impiegato" in full
    assert "Nota" in full and "inserimento manuale" in full
    print("[OK] Solo dati forniti dall'utente presenti nel PDF.")

    has_richiesta = "Richiesta" in full
    has_regolamento = "Regolamento" in full or "Generale" in full or "Aziendal" in full
    assert has_richiesta, "Contenuto AMM docx mancante"
    assert has_regolamento, "Contenuto TUTTI docx mancante"
    print("[OK] Contenuto originale dei template (solo files .docx/xlsx/pdf) presente.")

    root.destroy()

print()
print("== TUTTI I TEST DATI PULITI SUPERATI ==")
