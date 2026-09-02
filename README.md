# Formazioni PZZ

Programma Python locale per creare un unico PDF di formazione per una persona appena entrata in azienda.

Inserisci nome, data di ingresso e reparto: l'app legge i template dalla cartella `templates`, sostituisce `*nome*` e `*data*`, applica i documenti del reparto e quelli nominati `TUTTI`, quindi prepara un dossier PDF stampabile.

## Avvio

- Windows: doppio clic su `avvia_formazioni.bat`
- macOS / Linux: esegui `./avvia_formazioni.sh`
- In alternativa: `python launcher.py`

Al primo avvio il programma crea automaticamente un ambiente `.venv` locale e installa le dipendenze da `requirements.txt`. I file e i dati restano sul computer: non è necessario alcun account o collegamento internet per usare l'app dopo l'installazione.

## Template

Inserisci i file Word `.docx` o Excel `.xlsx` nella cartella `templates`.

Il nome deve essere:

```text
REPARTO_NUMERO_CODICE.docx
REPARTO_NUMERO_CODICE.xlsx
```

Esempi:

```text
SD_1_AAA.docx
TUTTI_2_AAA.xlsx
```

`NUMERO` indica quante copie del documento vengono inserite. `CODICE` è composto da tre lettere e serve solo per distinguere file simili. I dettagli completi sono in `templates/README.md`.

## Nota sui formati

La prima versione supporta i formati moderni `.docx` e `.xlsx`. Se hai file `.doc` o `.xls`, salvali prima nei rispettivi formati moderni da Word o Excel.