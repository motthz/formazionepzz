# Formazioni PZZ

[![Test e build](https://github.com/motthz/formazionepzz/actions/workflows/build-windows-exe.yml/badge.svg)](https://github.com/motthz/formazionepzz/actions/workflows/build-windows-exe.yml)
[![Ultima versione](https://img.shields.io/github/v/release/motthz/formazionepzz?label=versione)](https://github.com/motthz/formazionepzz/releases/latest)
[![Licenza MIT](https://img.shields.io/badge/licenza-MIT-blue)](LICENSE)

🇬🇧 [English version](README.en.md)

Programma per Windows che crea in pochi secondi il **dossier di formazione in PDF** per chi entra in azienda. Scegli nome, data di ingresso e reparto: Formazioni PZZ prende i moduli Word, Excel e PDF del reparto, ci scrive `*nome*` e `*data*` e li unisce in un unico file pronto da stampare. Tutto resta sul computer.

![Finestra principale](docs/img/principale.png)

📖 **[Guida all'uso](https://motthz.github.io/formazionepzz/)** · 🆕 **[Novità](CHANGELOG.md)** · 🐞 **[Segnala un problema](https://github.com/motthz/formazionepzz/issues/new/choose)**

## Installazione

Scarica **`FormazioniPZZ_Setup.exe`** dall'[ultima versione](https://github.com/motthz/formazionepzz/releases/latest) ed eseguilo.

- Installa per l'utente corrente in `%LOCALAPPDATA%\Programs\FormazioniPZZ`, senza permessi di amministratore.
- Crea il collegamento sul desktop e nel menu Start, e registra il programma in *Impostazioni → App installate*.
- Gli aggiornamenti arrivano da soli: "Aggiorna ora" scarica la nuova versione, chiude l'app e la riapre aggiornata, mantenendo modelli, reparti e impostazioni.
- Installazione automatica: `FormazioniPZZ_Setup.exe --silent` (opzioni `--dir PERCORSO`, `--no-shortcuts`).

In alternativa, `FormazioniPZZ.exe` nella stessa release è la versione **portatile**: si avvia senza installazione, ma si apre più lentamente e si aggiorna a mano.

## Funzioni principali

- **Modelli per reparto**: il nome del file indica reparto e copie (`SICUREZZA_2_SIC.docx`, `TUTTI_1_GEN.docx`); i dettagli sono in [templates/README.md](templates/README.md).
- **Una persona o tante**: elenco batch, importazione da CSV/Excel con modello già pronto, conversione in un'unica sessione di Word.
- **Anteprima**, **storico** dei dossier con "Usa questi dati", **filigrana** e **blocco modifica** del PDF.
- Cinque lingue (IT, EN, DE, ES, FR), tema chiaro, scuro o "Sistema", scorciatoie da tastiera, trascinamento dei file.
- I moduli senza segnaposto restano in cache; senza Office il batch lavora in parallelo.

Per conservare il layout originale dei modelli serve **Microsoft Office** oppure **LibreOffice** (con `soffice` nel PATH). Senza nessuno dei due, l'app ricostruisce i documenti con ReportLab.

## Aggiornamenti e dati inviati

All'avvio, al massimo una volta al giorno, l'app legge il numero dell'ultima versione pubblicata su GitHub: non invia nessun dato. Si può disattivare in **Impostazioni**, dove si può anche indicare un'altra origine: una cartella di rete con `version.json` e `FormazioniPZZ_Setup.exe` (esempio in [release/version.json](release/version.json)) oppure l'indirizzo web di un `version.json`.

## Sviluppo

```bash
python -m pip install -r requirements-dev.txt
python app.py              # avvia l'app
python -m pytest           # test
python -m ruff check .     # controllo del codice
```

Su Windows si può anche fare doppio clic su `avvia_formazioni.bat`: crea da solo l'ambiente `.venv` e installa le dipendenze. Struttura del codice, flusso con pull request e pubblicazione delle versioni sono in **[CONTRIBUTING.md](CONTRIBUTING.md)**; la compilazione degli exe è in [BUILD_EXE.md](BUILD_EXE.md).

## Licenza

[MIT](LICENSE): puoi usare, modificare e ridistribuire il programma, anche a scopo commerciale, mantenendo l'indicazione dell'autore.
