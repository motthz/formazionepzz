# Contribuire a Formazioni PZZ

## Preparare l'ambiente

```bash
git clone https://github.com/motthz/formazionepzz.git
cd formazionepzz
python -m pip install -r requirements-dev.txt
python app.py
```

Serve Windows per l'interfaccia completa e per la conversione con Microsoft Office. Senza Office l'app usa LibreOffice, se installato, oppure il motore ReportLab.

## Come è organizzato il codice

| Percorso | Contenuto |
|---|---|
| `app.py` | avvio dell'app |
| `formazioni/config.py` | versione, percorsi, impostazioni, lingue |
| `formazioni/templates.py` | nomi dei moduli, reparti, impronte MD5 |
| `formazioni/office.py` | conversione con Word/Excel o LibreOffice |
| `formazioni/pdf.py` | motore dei dossier: copertina, unione, filigrana, batch, cache |
| `formazioni/ui/` | interfaccia: `app.py` finestra principale, `batch.py`, `history.py`, `settings.py`, `dialogs.py`, `feedback.py`, `kit.py` (grafica) |
| `lang/*.json` | testi nelle 5 lingue |
| `installer/` | programma di installazione |
| `tests/` | test (pytest) |
| `tools/` | logo e screenshot della documentazione |

## Flusso di lavoro

`main` è protetto: ogni modifica entra con una **pull request** e solo se i controlli automatici passano.

1. Crea un ramo: `git switch -c nome-della-modifica`.
2. Lavora e verifica in locale:
   ```bash
   python -m ruff check .
   python -m pytest
   ```
3. Fai push del ramo e apri la pull request. Il workflow "Test e build Windows" lancia ruff e i test, compila gli exe, li avvia e prova l'installazione.
4. Quando è tutto verde, unisci la pull request.

Ogni nuovo testo dell'interfaccia va aggiunto in tutti e cinque i file `lang/*.json`.

## Pubblicare una versione

1. Aggiorna `APP_VERSION` in `formazioni/config.py`, `release/version.json` e `CHANGELOG.md`; i test verificano che le versioni coincidano.
2. Unisci la pull request su `main`.
3. Crea il tag e invialo: `git tag v2.4.0` e poi `git push origin v2.4.0`.
4. Il workflow compila gli exe e li allega alla release del tag. Le app installate propongono l'aggiornamento al successivo avvio.
