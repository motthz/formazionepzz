# Ottimizzazione Velocità Generazione Dossier — Implementation Plan

## Repository Research (Baseline e Colli di Bottiglia MISURATI con cProfile (5 run, reparto PRODUZIONE = 4 documenti totali):

| Misurazione:
  - Baseline TEMPO MEDIO 142.8 ms, MEDIANA 132.8 ms (min 130.0 / max 186.3 ms)

Categorie di costo (tempo cumulativo su 0.721s di profiling):
  1. [docx (Document() I/O + parsing stili DOCX    75.0%   (534.9 ms)
     - python-docx Document(path) = 96 ms per open (×10 = 96ms)
     - lookup stili xmlchemy/xpath = ~205ms
  2. [shutil.which] LibreOffice/Mot presente): 27.2%   (196 ms) — 5× _office_command() — CHIAMATO OGNI VOLTA ANCHE SE LIBREOFFICE NON C'E INSTALLATO
  3. [reportlab build] Layout+ platypus build = 17.0% (122 ms — 24.4 ms/run
  4. [deepcopy(story_cache] Cache deep-copiate 23.6% (168 ms — 33ms/run
  5. [placeholder sostituzione  1.2%  (8.5 ms — NON È UN PROBLEMA)

Ottimizzazioni proposte, ordinate per GUADAGNO atteso (% sulla baseline 132 ms:

### 🔹 1. Caching _office_command risultato (LOW-HANGING FRUIT — CACHING DELL'ESITO
  - **Problema**: ad ogni build_pdf() viene rieseguo `shutil.which("libreoffice")` e `soffice` (2×4940 chiamate a nt._path_exists
  - **Soluzione**: caching globale con `functools.lru_cache(maxsize=1 o variabile modulo-level booleana calcolata UNA SOLA ALL'AVVIO e rieseguita se cambia PATH
  - **Risparmio atteso**: -27 ms / 132 = **-20.5% ↓20%** (da 133 → 106 (85 ms
  - **Rischio**: BASSO — idempotente, senza side-effect.

### 🔹 2. Story caching DOCX/XLSX/PDF story PRE-caricato UNA SOLA VOLTA per template + caching persistente cross-build (MEMORY CACHE GLOBALE
  - **Problema**: attuale cache `story_cache` è locale a build_pdf() — svuotata OGNI build. Inoltre le storie vengono RICARICANO da disco OGNI chiamata build_pdf; eseguita ogni template viene DE-SERIALIZZATO da zero
  - **Soluzione**: spostare story_cache → MODULE-level cache globale persistente (con invalidazione su template mtime(path) + caching lazy: solo se il file è cambiato sul disco
  - **Risparmio atteso**: -40~50 ms (da 133 → 83-93 → **-34.9% ↓**
  - **Rischio**: BASSO-MEDIO — invalidazione per mtime garantisce correttezza; invalidazione manuale disponibile
      (click refresh

### 🔹 3. Evitare deepcopy() inutili sulle story
  - **Problema**: deepcopy 28'820 chiamate deepcopy per run. Condizione attuale in [app.py:874-L877:
    ```python
    if len(cached) > 3 or isinstance(cached[0] if cached else None, Table):
        story.extend(deepcopy(cached))
    else:
        story.extend(cached)
    ```
  - **Soluzione**: invece di deepcopiare per template × N copie → generare story UNA volta e riusare direttamente. OPPURE usare copy.copy() (shallow) per le liste semplici. Il vero motivo originale era "modifiche post-cache? No — le story non sono mutate dopo la creazione. Valutare se è possibile rimuovere completamente il deepcopy in modo sicuro.
  - **Risparmio atteso**: -20~24 ms (  → **-16%**
  - **Rischio**: MEDIO — Bisogna assicurarsi che ReportLab build() NON muti flowables che poi le story non vengono riusate. Safe.

### 🔹 4. Python-docx stile lookup: evitare `paragraph.style.name.lower() su OGNI paragrafo (FA-146ms di lookup stili
  - **Problema**: [docx_story (app.py:564)
    `heading = paragraph.style.name.lower() if paragraph.style else ""`
    → 160 chiamate style lookup xpath.
  - **Soluzione**: caching per documento: memoizzare stili in un dict {style_id: is_heading_bool} calcolato 1 volta per documento. Oppure sostituire con pattern/stringa e usare hasattr/settatr.
  - **Risparmio atteso**: -10~14 ms (circa **-8~10%**
  - **Rischio**: BASSO.

### 🔹 5. Make_styles caching (caching stili ReportLab — invece di rigenerare 5 stili OGNI build
  - **Problema**: make_styles() crea 10 ParagraphStyle OGNI build_pdf()
  - **Soluzione**: caching @lru_cache o calcolo UNA volta sola (gli stili non dipendono da input utente)
  - **Risparmio atteso**: -3~5 ms ( **-3%
  - **Rischio**: BASSO.

### 🔹 6. (OPZIONALE AVANZATO) ReportLab `invariant=True + build lazy PDF ReportLab parametri di tuning come `@*  - **Risparmio atteso**: -5~8 ms ( — -4-6% meno impattante (

### Guadagno complessivo (cumulato ottimizzazioni 1-5:
  ** -32% (dallo scenario peggiore) **fino a 55-60% (migliore-60% — da 133 → **55-90 ms ( -43 ms (

## Files and Modules

- `app.py`:
  - Aggiunta cache `_office_command` lru_cache
  - Spostamento story_cache → modulo-global + mtime invalidation
  - Rimozione deepcopy sicura delle story (Verificare ReportLab side-effect) → o shallow copy
  - Memoize paragraph.style.name lookup per documento DOCX
  - Memoize make_styles() @lru_cache o caching globale

## Implementation Steps (dependency order)

1.  Caching `_office_command()` (indipendente, zero side-effect, testabile da solo)
2.  Memoize `make_styles()` → caching globale stili
3.  Spostare `story_cache` → MODULE-LEVEL dict persistente con: `key=(path, mtime_ns, employee_name_entry_date hash? → No: le storie hanno *nome*/*data* placeholder già sostituiti ATTENZIONE: la cache attuale sostituisce i placeholder! Quindi la cache per story dipende *nome* e *data*. **Due livelli:
    a. Caching LIVELLO 1: parsing template RAW DOCX/XLSX → flowables pre-placeholders. Poi applico placeholders
    b. Oppure KEY = (path, mtime, employee_name, entry_date) → per scenario singolo utente ripetuto → stesso nome + stesso dipendente batch.
    → Soluzione migliore per uso reale: `cache key = (template.path.resolve(), template.path.stat().st_mtime_ns)` per il DOCX story SENZA placeholders; POI applico sostituzione placeholders a story prebuildate. Questo evita di re-parse tutto il documento per OGNI build! Questa KEY.
4.  Ottimizzazione DOCX style lookup per documento.
5.  (Dopo testata) Eliminare deepcopy l deep copy inutile se il ReportLab.

## Dependencies and Considerations

- **Nessuna dipendenza. Tutte le modifiche sono standard library + librerie già installate.
- `functools.lru_cache / @lru_cache standard library (built-in.
- `pathlib.Path.stat().st_mtime_ns (standard.
- Bisogna assicurarsi: se ReportLab `SimpleDocTemplate.build()` NON muta i flowables in `story`. Se si → allora copy.copy() (shallow) basta).
- Cache invalidation: se l'utente modifica templates/ mentre l'app è aperta → click Aggiorna documenti svuotare la cache modulo-level. Aggiornare `refresh_templates()` in FormazioniApp per pulire la cache.

## Validation

1.  Rieseguire `_profile_build.py` ( 5 run, reparto PRODUZIONE):
    - Confrontare media/mediana PRE vs POST
    - Confermare risparmio percentuale ≥ 30-60%
2.  Regressione: rieseguire test_completo.py (67/67 PASS invariato
3.  Regressione test E2E: tutti gli 8 reparti generano PDF corretti (nome+data presente)
4.  Smoke test EXE finale: ricompilato e avvio OK

## Risks

- **Deepcopy rischio**: Rimuovere deepcopy() puà portare a corruzione di flowables riusati. **Mitigazione**: invece di rimuovere del tutto, fare PRIMA verificare: 1 run → 3 documenti in entry_date ( prima implementazione: mantenere deepcopy sui copy() (shallow + testare; se OK passano → rimozione totale
- **story caching invalida: se template modificato disco → ricaricare. Mitigazione: usare `mtime_ns` + reset della cache dentro
- **_office_command caching cambio PATH a runtime → impossibile nella media quasi MA se l'utente installa LibreOffice mentre l'app è aperta non vede non vede. Mitigazione: cache ttl (10s) oppure clear cache al click Aggiorna.
- **Compatibilit EXE ricompilato dopo modifica per avere EXE e sorgente allineati.
