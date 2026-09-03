# Cartella template

Inserisci qui i documenti Word o Excel che vuoi includere nei dossier.

## Nome dei file

Il nome deve seguire questo formato:

```text
REPARTO_NUMERO_CODICE.docx
REPARTO_NUMERO_CODICE.xlsx
```

Esempi:

```text
SD_1_AAA.docx
TUTTI_2_AAA.xlsx
```

- `REPARTO` è il codice del reparto, per esempio `SD`.
- `TUTTI` include il documento in ogni dossier.
- `NUMERO` indica quante copie del documento inserire nel PDF.
- `CODICE` è un codice di tre lettere usato solo per distinguere file con lo stesso reparto e numero.
- Il programma riconosce anche reparti con più parole separate da `_`.

Nel contenuto dei file puoi usare:

- `*nome*` per il nome completo della persona;
- `*data*` per la data di ingresso.

Il formato consigliato è `.docx` per Word e `.xlsx` per Excel.