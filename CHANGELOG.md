# Novità

Tutte le modifiche importanti di Formazioni PZZ, dalla più recente. Gli eseguibili di ogni versione si scaricano dalle [Release](https://github.com/motthz/formazionepzz/releases).

## 2.5.0

**Per chi usa l'app**
- **Niente più pagina introduttiva**: il dossier contiene solo i moduli, senza la copertina "Dossier Formazione" con nome, data e reparto.
- **Niente nome del file o del foglio in fondo alla pagina**: i campi "Nome file" dei modelli Word (che stampavano il nome del file temporaneo, per esempio `1-MAGAZZINO_1_MAG.docx`) e i codici nome file/nome foglio delle intestazioni Excel vengono tolti; il resto dell'intestazione e del piè di pagina (loghi, numeri di pagina, testi) resta com'è.
- **Documenti scelti per più persone**: con persone nell'elenco batch, sotto la lista dei documenti compare che le spunte valgono per tutte (e di quale reparto sono). Aggiungendo una persona di un altro reparto l'app chiede conferma; dopo l'importazione di un file con persone tutte dello stesso reparto, la lista mostra quel reparto, così si spuntano proprio i documenti che riceveranno.

## 2.4.0

**Per chi usa l'app**
- **Finestra più compatta**: intestazione più bassa e card più strette. Nome, data e reparto sono visibili appena si apre il programma, senza scorrere.
- La finestra si adatta allo schermo e alla scala di Windows e si apre centrata.
- Le card "Anagrafica" e "Riepilogo" si affiancano solo quando c'è spazio: niente più pulsanti tagliati nelle finestre strette. I testi lunghi vanno a capo sulla larghezza disponibile.
- L'elenco batch vuoto spiega come aggiungere le persone.
- **Avvio più rapido**: la finestra si apre in circa un terzo di tempo in meno. Il lucchetto nel piè di pagina è un'icona disegnata invece di un'emoji, che costava quasi mezzo secondo all'avvio.

**Correzioni**
- I messaggi di stato dei template ("11 template pronti" e simili) erano sempre in italiano: ora sono tradotti in tutte le lingue.

## 2.3.1

**Correzioni**
- I modelli Word ed Excel con `*nome*` o `*data*` non perdono più immagini e loghi delle intestazioni: si modifica solo il testo dei segnaposto e il resto del file resta identico all'originale (prima le immagini nella stessa riga del segnaposto venivano cancellate, e in Excel si perdevano immagini d'intestazione, forme e caselle di testo).
- Un segnaposto spezzato in più parti (per esempio dopo una correzione in Word) viene sostituito anche se nello stesso paragrafo ce n'è un altro.
- I segnaposto si sostituiscono anche nelle caselle di testo e nelle intestazioni/piè di pagina di prima pagina e pagine pari, e nelle intestazioni di stampa di Excel.
- Interfaccia molto più fluida durante scorrimento, ridimensionamento e passaggio a schermo intero: pulsanti e campi arrotondati si disegnano con poche operazioni invece di centinaia, le card non ricreano più un'immagine grande quanto la finestra e l'intestazione si ridisegna solo a ridimensionamento finito.
- Scorrimento con rotellina e touchpad a passi regolari; non scorre più la finestra principale quando si usa la rotellina in un'altra finestra.

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
