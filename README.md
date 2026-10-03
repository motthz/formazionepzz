# Formazioni PZZ

Programma Python locale per creare un unico PDF di formazione per una persona appena entrata in azienda.

Inserisci nome, data di ingresso e reparto: l'app legge i template dalla cartella `templates`, sostituisce `*nome*` e `*data*`, applica i documenti del reparto e quelli nominati `TUTTI`, quindi prepara un dossier PDF stampabile.

## Installazione (consigliata)

Scarica ed esegui:

```text
release/FormazioniPZZ_Setup.exe
```

Il programma di installazione:

- installa Formazioni PZZ per l'utente corrente in `%LOCALAPPDATA%\Programs\FormazioniPZZ` (non servono permessi di amministratore);
- crea il collegamento **sul desktop** e nel **menu Start**;
- registra il programma in *Impostazioni → App installate*, da cui si può disinstallare.

Rieseguendo il setup su un PC dove è già installato, il programma viene aggiornato mantenendo modelli, reparti e impostazioni.
Per installazioni automatiche: `FormazioniPZZ_Setup.exe --silent` (opzioni: `--dir PERCORSO`, `--no-shortcuts`).

## Avvio senza installare dipendenze

Per il PC aziendale usa il file:

```text
release/FormazioniPZZ.exe
```

È un eseguibile Windows autonomo: contiene già Python e tutte le librerie necessarie. Non richiede installazioni, permessi amministrativi o connessione internet per l'utilizzo.

## Avvio da sorgente

- Windows: doppio clic su `avvia_formazioni.bat`
- macOS / Linux: esegui `./avvia_formazioni.sh`
- In alternativa: `python launcher.py`

L'avvio da sorgente crea automaticamente un ambiente `.venv` locale e installa le dipendenze da `requirements.txt`.

## Template

Inserisci i file Word `.doc` / `.docx` o Excel `.xls` / `.xlsx` nella cartella `templates`.

Il nome deve essere:

```text
REPARTO_NUMERO_CODICE.doc (o .docx)
REPARTO_NUMERO_CODICE.xls (o .xlsx)
```

Esempi:

```text
SD_1_AAA.docx
TUTTI_2_AAA.xlsx
```

`NUMERO` indica quante copie del documento vengono inserite. `CODICE` è composto da tre lettere e serve solo per distinguere file simili. I dettagli completi sono in `templates/README.md`.

## Nota sui formati

I formati `.doc`, `.docx`, `.xls` e `.xlsx` sono supportati. Per conservare nel PDF il layout originale, installa LibreOffice e rendi `libreoffice` o `soffice` disponibile nel PATH. L'app modifica solo `*nome*` e `*data*`, converte ogni modulo con il motore Office e accoda le pagine originali senza ridisegnarle.
