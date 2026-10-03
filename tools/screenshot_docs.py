"""Rigenera gli screenshot della documentazione (docs/img/).

Usa impostazioni, storico e percorsi finti in una cartella temporanea, cosi' nelle
immagini non compaiono dati o percorsi del PC su cui si lancia lo script.
Solo Windows (cattura dello schermo con Pillow). Uso: python tools/screenshot_docs.py
"""

from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PIL import ImageGrab  # noqa: E402

import formazioni.ui.app as ui_app  # noqa: E402

OUT = ROOT / "docs" / "img"
FAKE = Path(tempfile.mkdtemp(prefix="fpzz-docs-"))
SHOWN_TEMPLATES = r"C:\FormazioniPZZ\templates"
SHOWN_OUTPUT = r"C:\FormazioniPZZ\output"


def patch(name: str, value) -> None:
    """Sostituisce un percorso in tutti i moduli del pacchetto che lo usano."""
    for module_name, module in list(sys.modules.items()):
        if module_name.startswith("formazioni") and name in vars(module):
            setattr(module, name, value)


def fake_history() -> list[dict]:
    people = [("Giulia Ferri", "SICUREZZA", 3), ("Marco Bassi", "MAGAZZINO+LOGISTICA", 5),
              ("Sara Conti", "AMMINISTRAZIONE", 2), ("Luca Moretti", "PRODUZIONE", 4)]
    now = datetime(2026, 9, 28, 9, 30)
    history = []
    for i, (name, dept, docs) in enumerate(people):
        pdf = FAKE / "output" / f"dossier_{name.replace(' ', '-')}_{dept}.pdf"
        pdf.parent.mkdir(exist_ok=True)
        pdf.write_bytes(b"%PDF-1.4\n")
        history.append({"path": str(pdf), "name": name, "department": dept, "documents": docs,
                        "entry_date": (now - timedelta(days=i * 3)).strftime("%d/%m/%Y"),
                        "ts": (now - timedelta(days=i * 3, hours=i)).isoformat(timespec="seconds")})
    return history


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    (FAKE / "settings.json").write_text(json.dumps({
        "language": "it", "theme": "light", "check_updates": False,
        "last_template_dir": str(ROOT / "templates"), "last_output_dir": str(FAKE / "output"),
    }), encoding="utf-8")
    (FAKE / "history.json").write_text(json.dumps(fake_history()), encoding="utf-8")
    patch("SETTINGS_FILE", FAKE / "settings.json")
    patch("HISTORY_FILE", FAKE / "history.json")
    patch("HASHES_FILE", FAKE / "hashes.json")

    ui_app._enable_dpi_awareness()
    root = ui_app.tk.Tk()
    app = ui_app.FormazioniApp(root)
    root.geometry("1280x900+0+0")
    root.attributes("-topmost", True)

    def neutral_paths():
        app.template_dir.set(SHOWN_TEMPLATES)
        app.output_dir.set(SHOWN_OUTPUT)

    def grab(widget, name):
        widget.update()
        x, y = widget.winfo_rootx(), widget.winfo_rooty()
        image = ImageGrab.grab(bbox=(x, y, x + widget.winfo_width(), y + widget.winfo_height()),
                               all_screens=True)
        image.convert("RGB").save(OUT / name, optimize=True)
        print("salvato", OUT / name)

    def toplevel():
        return [w for w in root.winfo_children() if isinstance(w, ui_app.tk.Toplevel)][-1]

    steps = []

    def step(func):
        steps.append(func)
        return func

    @step
    def principale():
        neutral_paths()
        app.employee_name.set("Giulia Ferri")
        app.department.set("SICUREZZA")
        app.update_document_list()
        for name, day, dept in (("Marco Bassi", "01/10/2026", "MAGAZZINO"),
                                ("Sara Conti", "02/10/2026", "AMMINISTRAZIONE")):
            row = {"Nome": name, "Data": day, "Reparto": dept}
            app._inline_batch_rows.append(row)
            app._inline_batch_tree.insert("", "end", values=(name, day, dept))
        app._refresh_inline_batch_count()
        grab(root, "principale.png")

    @step
    def generazione():
        app._body_canvas.yview_moveto(1.0)
        app._toast("Dossier pronto: dossier_Giulia-Ferri_SICUREZZA.pdf · 3 documenti", "success",
                   ("Apri PDF", lambda: None), 60000)
        grab(root, "generazione.png")
        app._close_toast()
        app._body_canvas.yview_moveto(0.0)

    @step
    def storico():
        app.open_history()
        grab(toplevel(), "storico.png")
        toplevel().destroy()

    @step
    def impostazioni():
        app.open_settings()
        grab(toplevel(), "impostazioni.png")
        toplevel().destroy()

    @step
    def scuro():
        app.theme_pref = "dark"
        app.theme.set("dark")
        app.template_dir.set(str(ROOT / "templates"))  # la ricostruzione rilegge i moduli
        app._rebuild_ui()
        neutral_paths()
        grab(root, "tema-scuro.png")

    def run(index=0):
        if index >= len(steps):
            root.destroy()
            return
        steps[index]()
        root.after(700, lambda: run(index + 1))

    root.after(2000, run)
    root.mainloop()


if __name__ == "__main__":
    main()
