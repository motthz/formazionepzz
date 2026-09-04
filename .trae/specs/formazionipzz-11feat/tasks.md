# Formazioni PZZ 11 feat — Implementation Plan

## Task 1: Infrastructure layer — settings + i18n + template hash utility
- **Status**: `pending`
- **Priority**: high
- **Depends On**: None
- **Description**:
  - Creare `settings.json` in APP_DIR (persistenza tema, lingua, ultimi valori)
  - Creare cartella `lang/` con file `it.json` e `en.json` (tutti i testi hardcoded estratti in modo minimale, focus UI)
  - Creare modulo utility `compute_template_hash(path)` → MD5, persistenza `.template_hashes.json` + funzione `detect_changed_templates(templates, saved_hashes) → set[Path]`
  - Aggiungere attributi `self.settings`, `self.lang`, `self.tr(key)` in `FormazioniApp.__init__`
- **Acceptance Criteria Addressed**: AC-1, AC-4, AC-11
- **Test Requirements**:
  - `rule` TR-1.1: settings.json creato al primo avvio; load/save roundtrip OK; evidenza JSON scritto con `theme`, `language`
  - `rule` TR-1.2: file lang/it.json e lang/en.json creati con almeno 20 chiavi ciascuno (label e pulsanti principali); evidenza confronto key-sync
  - `rule` TR-1.3: `compute_template_hash` produce hash identico per file identico, diverso se modifico 1 byte; evidenza `assert hash1 == hash2_before_edit != hash3_after_edit`
  - `rubric` TR-1.4: Pulizia codice nuova infrastruttura; scala 1-5; anchors: 1 = spaghetti; 3 = funzioni isolate; 5 = modulo incapsulato riutilizzabile; threshold >=4; evidenza lettura codice e separazione responsabilità
- **Notes**: Questo task è il fondamento per tutti gli altri; deve essere fatto per primo.

## Task 2: Dark mode — palette ttk.Style e switch persistente
- **Status**: `pending`
- **Priority**: high
- **Depends On**: Task 1
- **Description**:
  - Introdurre seconda palette di stili "Dark.*" per TFrame, Card, Label, TEntry, TButton, Treeview, Scrollbar, Checkbutton, Combobox
  - Aggiungere pulsante toggle tema in header (stile switch in alto a destra nel canvas header)
  - Callback: ricostruire stili + ricolorare tutti i widget tk.Frame/TLabel nativi che non usano ttk
  - Persistere `theme` in settings.json; caricare all'avvio
- **Acceptance Criteria Addressed**: AC-1
- **Test Requirements**:
  - `rule` TR-2.1: Switch chiaro → scuro: `self.root.cget("bg")` e `style.lookup("Card.TFrame","background")` cambiano in "#0e1a22" (dark); evidenza cambio tema 2 clic
  - `rule` TR-2.2: Riavvio app dopo switch → tema ripristinato; evidenza settings.json
  - `rubric` TR-2.3: Qualità estetica dark mode; scala 1-5; anchors 1=incompleto/sbiadito; 3=leggibile; 5=palette coerente WCAG AA contrasto >= 4.5:1 su testo normale; threshold >=4; evidenza ispezione visiva colori e confronto

## Task 3: Date picker widget (puro Tk, con fallback tkcalendar se disponibile)
- **Status**: `pending`
- **Priority**: high
- **Depends On**: Task 1 (per `tr` key i18n nomi mesi/giorni)
- **Description**:
  - Creare `DatePickerFrame` custom (Tk Frame + 3 Combobox Anno/Mese/Giorno + opzionalmente popup calendario a griglia 7x6 se si vuole render minimale; bastano 3 combobox sincronizzati)
  - Validazione automatica: numero giorni mese → bisestile incluso; blocca combinazioni non valide
  - Formato output stringa DD/MM/YYYY + metodo `get_date()` → `date`
  - Sostituire Entry campo data in card 2 con questo widget
- **Acceptance Criteria Addressed**: AC-2
- **Test Requirements**:
  - `rule` TR-3.1: Impostare 29 febbraio 2024 (valido) → campo mostra "29/02/2024" e `get_date().isoformat()=="2024-02-29"`; evidenza
  - `rule` TR-3.2: Provare a impostare 31/04 → bloccato o auto corretto a 30/04; nessun traceback; evidenza
  - `rule` TR-3.3: Inizializzare con data odierna all'apertura form nuovo dossier; evidenza valore iniziale

## Task 4: Progress bar reale + generazione non bloccante (threading + after)
- **Status**: `pending`
- **Priority**: high
- **Depends On**: Task 1, Task 3
- **Description**:
  - Aggiungere `ttk.Progressbar` determinate 0-100 in card 3 sopra il pulsante genera + label `self.progress_label`
  - Refattorizzare `generate` in 2 fasi: calcola step totali → esegue in thread di lavoro con coda / `root.after` per aggiornare GUI
  - Per ogni template applica step: (1/n * 100)% con n = numero template totali (include copie? no, n template unici)
  - Batch mode userà lo stesso componente
- **Acceptance Criteria Addressed**: AC-3
- **Test Requirements**:
  - `rule` TR-4.1: Con 5 template, progress va da 0 a 20→40→60→80→100; label mostra 1/5 → 5/5; evidenza print nei log
  - `rule` TR-4.2: GUI non freeze — durante generazione window drag resize possibile; evidenza
  - `rule` TR-4.3: Errori durante thread → catturati e mostrati come messagebox senza hang; evidenza errore simulato e catch
- **Notes**: usare `queue.Queue` + `self.root.after(50, self._drain_queue)` pattern standard per thread-safe Tk update.

## Task 5: Multilingua IT/EN completo + selettore
- **Status**: `pending`
- **Priority**: medium
- **Depends On**: Task 1
- **Description**:
  - Sostituire tutte le stringhe hardcoded UI in `_build_*` con `self.tr("CHIAVE")`
  - Aggiungere Combobox "Lingua / Language" nella card 01 (configurazione) o nell'header
  - Callback cambio lingua: distruggi e ricostruisci body (header, card 1, 2, 3, footer) oppure rebind label testi (ricostruire il body è più pulito)
  - Tradurre anche messagebox, status, tooltip
- **Acceptance Criteria Addressed**: AC-4
- **Test Requirements**:
  - `rule` TR-5.1: Switch IT → EN → "⬇ Genera PDF unico" diventa "⬇ Generate single PDF" (o corrispettiva chiave EN); evidenza
  - `rule` TR-5.2: Riavvio dopo cambio lingua → ricordata; settings.json "language":"en"; evidenza
  - `rubric` TR-5.3: Completezza traduzioni (minimo 40 chiavi) senza hardcoding residuo evidente; scala 1-5; anchors 1 = metà; 3 = 90%; 5 = 100% testi UI; threshold >=4; evidenza grep "Inserisci nome|Data di ingresso|Reparto|Genera|Scegli" in app.py deve dare ZERO dopo Task 5

## Task 6: Tooltip contestuali (puro Tk — classe Tooltip con delay + auto dismiss)
- **Status**: `pending`
- **Priority**: medium
- **Depends On**: Task 1, Task 5 (per testi tooltip i18n)
- **Description**:
  - Creare `Tooltip` class in app.py (o modulo helper): widget tk.Toplevel piccolo giallo pallido con testo, attiva bind `<Enter>` con after 600ms; `<Leave>` e `<Motion>` distruggono; auto-dismiss dopo 8 sec
  - Applicare tooltip a: entry nome, entry data, ruolo, note, reparto combo, pulsanti "Scegli cartella" x2, "Aggiorna documenti", "Genera PDF unico", "Batch multiplo", "Gestisci reparti", checkbox Multi-reparto, Checkbox auto-open
- **Acceptance Criteria Addressed**: AC-5
- **Test Requirements**:
  - `rule` TR-6.1: Hover sul campo nome per >600ms → tooltip appare con testo i18n corretto; evidenza screenshot
  - `rule` TR-6.2: Mouse via dopo apparizione → tooltip scompare entro 200ms; evidenza
  - `rubric` TR-6.3: Completezza (tutti i widget elencati hanno tooltip); scala 1-5; anchors 1 = <3; 3 = 8/11; 5 = 11/11; threshold >=4; evidenza lista binding verificata codice

## Task 7: Layout responsive + DPI-awareness (Windows) + barre di scroll sempre visibili
- **Status**: `pending`
- **Priority**: medium
- **Depends On**: Task 2 (colori adattivi)
- **Description**:
  - Su Windows: chiamata `ctypes.windll.shcore.SetProcessDpiAwareness(1)` prima di `Tk()` se disponibile (try/except)
  - Introdurre layout breakpoints: se finestra `<1100 px` larghezza → card 2 e 3 non più affiancate ma in colonna (stacco verticale)
  - Aggiungere `minsize` consistente; scrollbar sempre disponibili; wraplength proporzionale alle label di help
  - I testi non devono mai essere troncati: wraplength dinamico sulle Label nei card
- **Acceptance Criteria Addressed**: AC-6
- **Test Requirements**:
  - `rule` TR-7.1: Su Windows con DPI 125% — app non sfocata; SetProcessDpiAwareness chiamata e non genera eccezione; evidenza
  - `rule` TR-7.2: Larghezza finestra ridotta a 980px (minsize) → card 3 va sotto la card 2 (layout 1-colonna); nessun elemento nascosto; barre di scorrimento raggiungibili
  - `rule` TR-7.3: Allargamento a 1920 → card affiancate proporzionalmente; label d'aiuto con wrap dinamico: testo non troncato

## Task 8: Checklist template in Treeview (checkbox per riga)
- **Status**: `pending`
- **Priority**: high
- **Depends On**: Task 1 (per `self.tr`)
- **Description**:
  - Sostituire Treeview esistente con Treeview + colonna "Includi" con icona ☑/☐ gestite via tags + evento `<Button-1>` click toggle + spazio tastiera
  - tenere `self.template_inclusion: dict[Path, bool]`
  - Aggiungere 2 shortcut "Seleziona tutti" / "Deseleziona tutti" in alto
  - Aggiornare `update_document_list` per popolare i check allineando default True
  - In `generate`: filtrare `selected_template = [t for t in selected if self.template_inclusion.get(t.path, True)]`
  - Badge conteggio deve riflettere i soli selezionati
- **Acceptance Criteria Addressed**: AC-8
- **Test Requirements**:
  - `rule` TR-8.1: 4 template visualizzati → clic sulla seconda riga → checkbox toggla a off; badge conteggio scende di (copie_riga_2)
  - `rule` TR-8.2: Deseleziono 2 → genera PDF → solo 2 template presenti; evidenza conteggio pagine PDF o somma copie
  - `rule` TR-8.3: Pulsanti "Tutti" e "Nessuno" → stato istantaneo all check on/off; evidenza

## Task 9: Multi-reparto (checkbox + Lista check)
- **Status**: `pending`
- **Priority**: high
- **Depends On**: Task 8 (condivisione struttura Treeview inclusioni)
- **Description**:
  - Introdurre checkbox "Modalità multi-reparto" sopra l'attuale Combobox singolo reparto
  - Quando False → UI come prima
  - Quando True → nascondi Combobox singolo, mostro Listbox multi-selezione (o Treeview con check) dei reparti
  - Implementare `templates_for_departments(templates, departments: list[str])` → unione dei template di ciascun reparto + TUTTI, dedup per path, ordinamento stabile
  - Aggiornare funzione `generate` per supportare lista reparti (unire in metadati, nome reparto nel PDF come "Produzione + Logistica")
- **Acceptance Criteria Addressed**: AC-9
- **Test Requirements**:
  - `rule` TR-9.1: checkbox multi-reparto on → UI mostra Listbox / check-list reparti; off → torna combobox singolo
  - `rule` TR-9.2: seleziono Produzione e Logistica (entrambi con almeno 1 template) + 2 template TUTTI → albero include template Produzione ∪ Logistica ∪ TUTTI dedup; conteggio corretto
  - `rule` TR-9.3: generazione PDF non crasha; reparto multi in metadati visibile nel nome file o messagebox

## Task 10: Editor reparti integrato (finestra modale + persistenza su `reparti.txt`)
- **Status**: `pending`
- **Priority**: medium
- **Depends On**: Task 1, Task 5 (i18n messaggi)
- **Description**:
  - Aggiungere pulsante "Gestisci reparti" in Card 01 o vicino alla combobox reparto
  - Finestra modale Toplevel con Listbox + Entry + Aggiungi / Rinomina selezionato / Elimina selezionato + 2 pulsanti "Salva e chiudi" / "Annulla"
  - Salvataggio scrive `reparti.txt` con uppercase, commenti `#` preservati se si vuole (minimo: righe non commento prese in carico)
  - Dopo salvataggio esegui `refresh_templates()` per aggiornare immediatamente la combobox
- **Acceptance Criteria Addressed**: AC-10
- **Test Requirements**:
  - `rule` TR-10.1: Aggiungi reparto "R&D" → salva → reparti.txt contiene R&D; combobox dopo refresh lo include; evidenza file + GUI
  - `rule` TR-10.2: Elimino "R&D" → salva → file non più contiene; combobox refresh lo rimuove
  - `rule` TR-10.3: Annulla → nessuna modifica su disco; evidenza confronto hash md5 prima/apri/annulla

## Task 11: Generazione batch (CSV/Excel import)
- **Status**: `pending`
- **Priority**: high
- **Depends On**: Task 4 (progress), Task 8 (inclusioni template), Task 9 (reparti)
- **Description**:
  - Pulsante "Batch multiplo" nel card 3 o accanto a "Genera PDF unico"
  - Finestra modale: scegli file (.csv / .xlsx / .xls) → mostra anteprima tabella prime N righe → pulsante "Esegui batch"
  - Header previsti: `Nome,Data,Reparto[,Ruolo,Note]` case-insensitive
  - Per ogni riga: convalida → se invalida salta con messaggio "SKIP riga N: motivo"; se valida esegue generate-like
  - Riutilizza funzione di generazione + progress bar condivisa
  - Finisce con summary messagebox "N OK / M SKIP / K FAIL" + lista percorsi PDF
- **Acceptance Criteria Addressed**: AC-7
- **Test Requirements**:
  - `rule` TR-11.1: CSV 3 righe valide, 1 riga Data non valida → 3 PDF creati; report "3 OK / 1 SKIP / 0 FAIL"; evidenza file creati + messagebox
  - `rule` TR-11.2: XLSX 2 righe valide → 2 PDF; evidenza; nome file in output come dossier_NOME.pdf
  - `rule` TR-11.3: Riga con reparto inesistente → SKIP; nessun crash

## Task 12: Template hashing MD5 — indicatore visivo e avvisi
- **Status**: `pending`
- **Priority**: medium
- **Depends On**: Task 1 (infrastruttura hash), Task 8 (Treeview tag)
- **Description**:
  - In `refresh_templates`: calcola MD5 ogni template valido; confronta con `.template_hashes.json`; set tag `modified` o `new` nel Treeview
  - Colore tag `modified` = sfondo giallo chiaro (light mode) / arancio (dark); `new` = verde chiaro
  - Aggiornare `.template_hashes.json` DOPO la generazione andata a buon fine (quindi l'utente vede "MODIFICATO" finché non genera almeno una volta il PDF con quel template; comportamento simile a cache validation)
  - Aggiungere tooltip/stat nel footer: "3 template invariati · 1 modificato · 0 nuovi"
- **Acceptance Criteria Addressed**: AC-11
- **Test Requirements**:
  - `rule` TR-12.1: Modifico byte in un template → hash diverso; Treeview riga evidenziata tag modified; footer indica modificato
  - `rule` TR-12.2: Dopo una generazione andata a buon fine → hash salvato nel JSON; evidenza JSON contains file → dopo refresh tag sparisce (invariato)
  - `rule` TR-12.3: Nuovo file non nel JSON → evidenziato come nuovo tag "new"; dopo generazione → passa a invariato

## Task 13: Smoke test end-to-end simulando operatore reale + build EXE validation
- **Status**: `pending`
- **Priority**: high
- **Depends On**: Task 1 → Task 12 tutti
- **Description**:
  - Eseguire lo script da sorgente `python app.py` con scenario completo operatore:
    1. Avvio → switch dark mode
    2. Apro "Gestisci reparti" → aggiungo "TESTBATCH", salvo
    3. Switch a EN → torno IT
    4. Inserisco nome "Mario Rossi", scelgo data dal date picker
    5. Abilito multi-reparto e seleziono PRODUZIONE + LOGISTICA
    6. Deseleziono 1 template nella checklist
    7. Genero singolo PDF (controllo progress 0→100, nessun freeze)
    8. Avvio batch 3 righe CSV e controllo report
    9. Modifico 1 template docx con testo diverso → refresh vede tag modified
    10. Controllo messaggio footer MD5 stato
  - Build EXE: `pyinstaller FormazioniPZZ.spec` e check output esistente (no crash)
  - Run test_completo.py o test_pulito.py se esistenti
- **Acceptance Criteria Addressed**: AC-1..AC-11 finali
- **Test Requirements**:
  - `rule` TR-13.1: I 10 step dello scenario completano senza eccezione; al termine 1 singolo PDF + 3 batch PDF esistono in output/
  - `rule` TR-13.2: Build PyInstaller exit code 0; EXE generato (se build non è disponibile per tempo, almeno validazione `py_compile app.py` e `import app` in subprocesso Python OK)
  - `rubric` TR-13.3: Esperienza operatore reale; scala 1-5; anchors 1=bug/UX confusa; 3=funzionante ma sbavature; 5=fluida, chiara, nessun messaggio ambiguo; threshold >=4; evidenza note di marcia dello scenario stepwise
- **Notes**: Questo task è la verifica finale richiesta dall'utente "simulando di essere un vero operatore".
