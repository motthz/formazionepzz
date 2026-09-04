# Formazioni PZZ — Estensione 11 funzionalità
## Product Requirements Document

## Overview
- **Summary**: Estensione dell'app desktop Formazioni PZZ (Tkinter) con 11 migliorie selezionate: tema scuro persistente, date picker grafico, progress bar percentuale reale, multilingua IT/EN, tooltip contestuali, layout DPI-aware responsive, generazione massiva batch, checklist template selezionabili, selezione multi-reparto, editor reparti integrato, versioning hash MD5 dei template.
- **Purpose**: Migliorare drasticamente l'usabilità, la produttività e la tracciabilità dell'app, in ottica operatore HR/back-office che crea dossier formazione per assunzioni.
- **Target Users**: Operatori amministrativi / HR / formazione aziendale, tipicamente su PC Windows aziendale senza diritti admin.

## Goals
- G1: Rendere l'app gradevole e leggibile in ambienti con luce ridotta (dark mode).
- G2: Eliminare errori di battitura data in formato non valido (date picker).
- G3: Fornire feedback affidabile sul progresso generazione operazioni lunghe (progress bar + etichettate e percentuali.
- G4: Supportare utenti non italofoni nella stessa installazione (multilingua).
- G5: Guidare l'utente con hint contestuali riducendo errore operatore (tooltip).
- G6: Garantire leggibilità e usabilità a risoluzioni e DPI diversi (responsive/DPI-aware).
- G7: Accelerare onboarding multiplo contemporaneo (batch mode).
- G8: Permettere all'operatore di decidere al volo cosa includere (checklist).
- G9: Gestire persone afferenti a più reparti con unico dossier.
- G10: Evitare l'editing manuale di `reparti.txt`.
- G11: Tracciare le modifiche dei template per audit/formazione.

## Non-Goals
- Non si cambia architettura di generazione PDF (ReportLab + Office Converter).
- Non si reingezzano i template esistenti (compatibilità al 100%).
- Non si introducono dipendenze esterne oltre a quelle già presenti in requirements.txt (+ python-dateutil se indispensabile).
- Non si implementa cloud / sincronizzazione remota.
- Non si implementa firma digitale o gestione utenti con autenticazione.

## Background & Context
L'app è scritta Python con Tkinter (ttk.Style a tema clam), localizzata in [app.py](file:///c:/Users/stemo/OneDrive/Desktop/formazionepzz/app.py). La GUI ha 3 card in layout Canvas scrollabile verticale. Genera dossier PDF a partire da template Word/Excel/PDF sostituendo `*nome*` e `*data*`. I reparti sono letti da `reparti.txt` e inferiti dai nomi file.
Stato attuale analizzato: nessuna delle 11 feature in oggetto è presente ad oggi.

## Functional Requirements

### UX / UI
- **FR-1**: L'app espone un interruttore "Tema" (Chiaro/Scuro) nell'intestazione o nella card 01; lo stato viene salvato e ripristinato all'avvio successivo (`settings.json`).
- **FR-2**: Il campo data di ingresso è sostituito da un widget calendario date-picker Tk nativo (senza dipendenze pip aggiuntive, puro Tk canvas o tkcalendar se presente, altrimenti implementazione minimale in puro Tk con liste combinate).
- **FR-3**: Durante la generazione viene mostrata una progress bar (0-100%) con etichetta di stato testuale "Template N/TOT", senza bloccare totalmente il rendering; al termine batch viene mostrata la percentuale 100%.
- **FR-4**: L'app mette a disposizione un selettore lingua "Italiano / English" e carica le stringhe da file JSON; la scelta persiste tra le sessioni.
- **FR-5**: Ogni campo (nome, data, reparto, ruolo, note, pulsanti, scelta cartelle) ha tooltip che appare al passaggio del mouse con testo istruttivo.
- **FR-6**: L'app gestisce DPI-awareness e layout fluido adattando elementi in proporzione: le colonne e le card si riorganizzano/fluiscono a risoluzioni <1280x720 e >1920x1080, barre di scorrimento sempre presenti.
- **FR-7**: Pulsante "Batch multiplo" → apre finestra per importare file CSV o Excel con colonne Nome,Data,Reparto → genera un PDF per riga e report finale.
- **FR-8**: La lista di riepilogo documenti (Card 3) espone checkbox per ogni riga; l'utente può deselezionare template da escludere dalla generazione senza modificare file.
- **FR-9**: Il selettore reparto supporta modalità singola oppure multipla: una checkbox "Multi-reparto" → diventa lista check; se abilitata, i template di TUTTI i reparti selezionati vengono uniti + TUTTI.
- **FR-10**: Pulsante "Gestisci reparti" → finestra modale: aggiungi / rinomina / elimina reparti con salvataggio diretto su `reparti.txt`, senza editare il file manualmente.
- **FR-11**: L'app calcola MD5 di ogni template al discovery; salva il digest in `.template_hashes.json`; mostra indicatore visivo "nuovo / modificato" nell'elenco; avvisa l'utente se qualche template è cambiato dall'ultima generazione per audit.

## Non-Functional Requirements
- **NFR-1**: Backward compatibility totale con workflow esistente a tema chiaro singolo reparto.
- **NFR-2**: Prestazioni: nessun rallentamento percepibile (>200ms) rispetto alla versione corrente.
- **NFR-3**: Zero regressioni: build PyInstaller senza errori.
- **NFR-4**: Robustezza: ogni nuova feature gestisce fallback se dato mancante / formato errato senza traceback brutale.
- **NFR-5**: Zero nuove dipendenze obbligatorie; se `tkcalendar` non è installato, il date picker usa 3 combobox (anno/mese/giorno) in puro Tk.
- **NFR-6**: File aggiuntivi creati al primo avvio: `settings.json`, `template_hashes.json`, `lang/it.json`, `lang/en.json` tutti nella APP_DIR.

## Constraints
- **Tecniche**: Python 3.x + Tkinter ttk theme 'clam' + stack esistente; nessuna riscrittura GUI.
- **Business**: File sempre 100% offline; nessuna trasmissione di dati fuori dal PC.
- **Dipendenze**: `openpyxl>=3.1.5, python-docx>=1.1.2, reportlab>=4.2.5, pypdf>=5.0.0, pyinstaller>=6.0.0, pywin32; opzionale `tkcalendar` e `babel` se disponibili.

## Assumptions
- Windows è il target principale (99% casi d'uso).
- LibreOffice / Microsoft Office restano disponibili come motori di conversione come da comportamento esistente.
- CSV batch: colonne esatte "Nome", "Data", "Reparto" (intestazioni case-insensitive); se una riga non valida viene saltata e segnalata a fine batch.

## Acceptance Criteria

### AC-1: Dark mode persistente
- **Type**: rule
- **Given**: App avviata per la prima volta
- **When**: utente clicca l'interruttore tema
- **Then**: colori si invertono secondo palette scura coerente; riavvio dell'applicazione ricorda la scelta
- **Pass Condition**: switch tema + riavvio → tema selezionato ripristinato
- **Evidence**: screenshot / settings.json contiene `"theme": "dark"` o `"light"`; ispezione colori del frame root bg

### AC-2: Date picker grafico
- **Type**: rule
- **Given**: card 2 "Nuovo dossier"
- **When**: utente clicca il widget data
- **Then**: si apre un calendario minimale e la data selezionata viene inserita nel campo in formato `DD/MM/YYYY`
- **Pass Condition**: data valida, nessun errore su 31/02/2026 (non valida), nessun crash
- **Evidence**: selezione 04/09/2026 → campo valorizzato correttamente; prova data non valida bloccata o messaggio

### AC-3: Progress bar percentuale
- **Type**: rule
- **Given**: 5+ template e utente avvia generazione singolo o batch
- **When**: generazione in corso
- **Then**: barra incrementa per step proporzionali ai template convertiti (0%→100%; label `1/5`, ecc.); non freeze GUI
- **Pass Condition**: 100% al termine; nessun blocco dell'interfaccia
- **Evidence**: osservazione visuale con 5 template → 5 step distinti nella progress

### AC-4: Multilingua IT/EN
- **Type**: rule
- **Given**: app avviata
- **When**: cambio lingua da IT a EN
- **Then**: tutti i testi visibili (label, pulsanti, tooltip, messaggi) commutano senza refresh; riavvio ricorda lingua
- **Pass Condition**: "Genera PDF unico → Generate single PDF; tooltip inglesi corretti
- **Evidence**: confronto screenshot prima/dopo; settings.json

### AC-5: Tooltip contestuali
- **Type**: rubric
- **Dimension**: completezza e accuratezza dei tooltip
- **Scale**: 1-5
- **Anchors**: 1 = nessun tooltip visibile; 3 = tooltip campi obbligatori; 5 = tooltip tutti i campi + pulsanti + checkboxes + label; tooltip che scompaiono dopo ~8s o al click fuori; tooltip con delay ~600ms
- **Pass Threshold**: >= 4
- **Evidence**: screenshot hover campo nome, hover pulsante "Genera PDF unico"

### AC-6: Layout responsive DPI-aware
- **Type**: rule
- **Given**: finestre ridotta a 980x720 (minsize attuale)
- **When**: si ridimensiona a 800x600 (sotto minsize)
- **Then**: barre di scrollbar compaiono; nessun elemento essenziale nascosto o fuori frame
- **Pass Condition minsize mantenuto; widget scalano. Nessun testo troncato visibile
- **Evidence**: screenshot 980x720 e 1600x1200; testo "⬇ Genera PDF" completamente visibili in entrambe

### AC-7: Batch CSV batch
- **Type**: rule
- **Given**: CSV 3 righe valide (Mario Rossi 01/09/2026 Produzione, ...
- **When**: si clicca genera batch
- **Then**: 3 PDF distinti generati con nomi file separati; report summary con OK/FAIL per ogni riga
- **Pass Condition**: 3 file esistenti; 0 FAIL per righe valide + 1 riga non valida saltata
- **Evidence**: elenco file generati + log console o messagebox summary

### AC-8: Checklist template selezionabili
- **Type**: rule
- **Given**: reparto Produzione con 4 template
- **When**: deseleziono 2 template
- **Then**: PDF generato include solo i 2 selezionati (controllo pagine = 2
- **Pass Condition**: conteggio pagine attese; riepilogo conteggio corretto
- **Evidence**: Treeview con checkboxes; PDF prodotto dopo deselezione

### AC-9: Selezione multi-reparto
- **Type**: rule
- **Given**: checkbox "Multi-reparto" abilitata e selezionati Produzione + Logistica (+ template TUTTI
- **When**: genera PDF
- **Then**: template di Produzione + Logistica + TUTTI sono inclusi deduplicati
- **Pass Condition**: conteggio template sommato e dedup; nessun duplicato
- **Evidence**: albero documenti e conteggio badge

### AC-10: Editor reparti integrato
- **Type**: rule
- **Given**: finestra modale "Gestisci reparti"
- **When**: aggiungo nuovo reparto "R&D" e salvo, poi elimino "TEST"
- **Then**: `reparti.txt` contiene R&D presente; TEST assenza di "TEST" in combobox reparto
- **Pass Condition**: file `reparti.txt` rispecchia cambiamenti dopo refresh
- **Evidence**: contenuto reparti.txt ispezionato + values della Combobox dopo refresh

### AC-11: Versioning MD5 template
- **Type**: rule
- **Given**: 3 template
- **When**: modifico contenuto MAGAZZINO_1_MAG.docx (edit 1 byte; rieseguo discover template
- **Then**: tag "MODIFICATO" evidenziato in giallo nell'elenco; .template_hashes.json contiene hash nuovo; messaggio stato all'utente "2 template invariati + 1 modificato
- **Pass Condition**: hash diversi prima / dopo la modifica; evidenza visiva; allarme; avviso messagebox opzionale
- **Evidence**: ispezione file JSON + tag giallo in Treeview

## Open Questions
- [ ] L'utente desidera che il progress bar sia indeterminato come alternativa grafico a 0%→100% (già stepwise stabilito → stepwise fisso) Risposta: stepwise su N step come da FR-3).
- [ ] Lingua aggiuntive oltre IT/EN Risposta: no
