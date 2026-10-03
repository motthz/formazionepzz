# Novità

Tutte le modifiche importanti di Formazioni PZZ, dalla più recente. Gli eseguibili di ogni versione si scaricano dalle [Release](https://github.com/motthz/formazionepzz/releases).

## 2.3.0

**Per chi usa l'app**
- **Aggiornamento in un clic**: dall'avviso "Aggiorna ora" l'app scarica il nuovo setup, si chiude e si riapre aggiornata. La versione installata è visibile in *Impostazioni di Windows → App installate*.
- **Registro errori** in `%LOCALAPPDATA%\FormazioniPZZ\formazioni.log`, con il pulsante "Apri registro errori" nelle Impostazioni. I messaggi di errore indicano dove trovarlo.
- La disinstallazione rimuove anche la cache dei moduli convertiti.
- [Guida all'uso](https://motthz.github.io/formazionepzz/) con screenshot.

**Correzioni**
- Nell'exe senza console la stampa degli errori interni poteva fallire a sua volta e bloccare il ripiego sul motore ReportLab.
- Un file ancora bloccato da Word dopo la conversione non fa più scartare un dossier già pronto.

**Progetto**
- Licenza MIT.
- Codice diviso nel pacchetto `formazioni/` (motore PDF, Office, modelli, interfaccia…) al posto di un unico `app.py` da oltre 5.000 righe.
- Test in `tests/` eseguiti con pytest, controllo del codice con ruff, entrambi automatici su ogni pull request. Prima della pubblicazione l'exe viene avviato e il setup installato in prova.
- Versioni delle librerie bloccate e aggiornate da Dependabot.
- Cronologia del repository ripulita dagli eseguibili (da 410 MB a meno di 1 MB).
- README in inglese, moduli guidati per le segnalazioni, istruzioni per contribuire.

## 2.2.0

- Controllo aggiornamenti dalle release di GitHub.
- Gli eseguibili si pubblicano come allegati delle release e non sono più salvati nel repository.

## 2.1.0

- Storico dei dossier, anteprima, modello Excel per il batch, filigrana e blocco modifica del PDF.
- Batch più veloce: un'unica sessione di Word/Excel, cache dei moduli senza segnaposto, generazione in parallelo senza Office.
- Avvio in circa 1 secondo con l'installer (versione "cartella").
- Tema "Sistema", icone uniformi, avvisi non bloccanti e notifiche di Windows, scorciatoie da tastiera, trascinamento dei file.
- Nuova grafica con pulsanti, campi e schede arrotondati.
- Corretto: l'app installata non trovava traduzioni e icone.
