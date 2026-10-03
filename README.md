# Formazioni PZZ

Programma Python locale per creare un unico PDF di formazione per una persona appena entrata in azienda.

Inserisci nome, data di ingresso e reparto: l'app legge i template dalla cartella `templates`, sostituisce `*nome*` e `*data*`, applica i documenti del reparto e quelli nominati `TUTTI`, quindi prepara un dossier PDF stampabile.

## Installazione (consigliata)

Scarica **`FormazioniPZZ_Setup.exe`** dall'ultima release: <https://github.com/motthz/formazionepzz/releases/latest>, poi eseguilo.

Il programma di installazione:

- installa Formazioni PZZ per l'utente corrente in `%LOCALAPPDATA%\Programs\FormazioniPZZ` (non servono permessi di amministratore);
- crea il collegamento **sul desktop** e nel **menu Start**;
- registra il programma in *Impostazioni → App installate*, da cui si può disinstallare.

Rieseguendo il setup su un PC dove è già installato, il programma viene aggiornato mantenendo modelli, reparti e impostazioni.
Per installazioni automatiche: `FormazioniPZZ_Setup.exe --silent` (opzioni: `--dir PERCORSO`, `--no-shortcuts`).

## Avvio senza installare dipendenze

Per il PC aziendale usa `FormazioniPZZ.exe`, allegato alla stessa [release](https://github.com/motthz/formazionepzz/releases/latest).

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

## Storico, anteprima e opzioni PDF

- **Storico**: elenca i dossier generati (fino a 500), con ricerca, apertura del PDF o della cartella e *Usa questi dati* per ricompilare il modulo.
- **Anteprima**: crea un dossier di prova in una cartella temporanea, con una sola copia per modulo e la filigrana ANTEPRIMA.
- **Modello Excel**: nell'elenco batch, scarica un `.xlsx` con le colonne Nome, Data e Reparto già pronte per l'import.
- **Impostazioni → Dossier PDF**: filigrana facoltativa su ogni pagina (es. `COPIA CONTROLLATA`) e blocco della modifica del PDF, che resta apribile e stampabile senza password.

## Uso rapido

- **Trascina i file sulla finestra**: un CSV, oppure un Excel lasciato sull'elenco dipendenti, viene importato nel batch. Word, Excel e PDF lasciati altrove si aggiungono come nuovi moduli.
- **Scorciatoie** (F1 per l'elenco completo): Ctrl+G genera, Ctrl+P anteprima, Ctrl+B genera tutto l'elenco, Ctrl+H storico, Ctrl+, impostazioni, F5 aggiorna i documenti, Invio nel campo nome aggiunge la persona all'elenco.
- **Tema "Sistema"**: segue la modalità chiara/scura di Windows, anche se cambia mentre l'app è aperta.
- A fine generazione compare un avviso in basso nella finestra, con "Apri PDF" o "Dettagli". Se si sta usando un'altra finestra arriva anche la notifica di Windows.

## Prestazioni

- Il batch converte i moduli di tutte le persone in **un'unica sessione** di Word/Excel: con 5 dossier circa la metà del tempo rispetto a generarli uno alla volta.
- I moduli **senza** `*nome*` e `*data*` (regolamenti, procedure) si convertono una sola volta: il PDF resta in cache in `%LOCALAPPDATA%\FormazioniPZZ\pdf-cache` e viene riusato finché il file non cambia.
- Senza Office installato il batch genera più dossier **in parallelo**, su più processori.
- L'installer installa la versione "cartella" dell'app, che si apre più in fretta dell'exe singolo. Le librerie di Word, Excel e PDF si caricano solo alla prima generazione.

## Aggiornamenti

All'avvio (al massimo una volta al giorno) l'app controlla in background l'ultima release pubblicata su GitHub. Se c'è una versione più recente compare un avviso con il pulsante **Scarica**, che apre il download di `FormazioniPZZ_Setup.exe`: eseguendolo il programma si aggiorna e mantiene modelli e impostazioni. Il controllo scarica solo il numero di versione e non invia dati. Se la rete non è raggiungibile, l'app non mostra nessun messaggio. Si può disattivare in **Impostazioni**.

In **Impostazioni → Origine aggiornamenti** si può indicare, al posto di GitHub:

- una cartella di rete (es. `\\server\FormazioniPZZ`) che contiene `version.json` e `FormazioniPZZ_Setup.exe`;
- oppure l'indirizzo web di un `version.json`.

Formato di `version.json` (un esempio è in `release/version.json`):

```json
{ "version": "2.2.0", "url": "facoltativo: link al setup", "notes": "novità della versione" }
```

## Pubblicare una nuova versione

1. Aggiorna `APP_VERSION` in `app.py` e `release/version.json` (`test_nuove_funzioni.py` verifica che coincidano).
2. Fai commit e push su `main`, poi crea e invia il tag: `git tag v2.3.0` e `git push origin v2.3.0`.
3. Il workflow GitHub compila `FormazioniPZZ_Setup.exe` e `FormazioniPZZ.exe` e li allega alla release di quel tag, creandola se non esiste. Le app installate vedranno l'aggiornamento al successivo avvio.
