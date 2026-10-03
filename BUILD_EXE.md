# Rigenerare l'eseguibile Windows

## Prerequisiti

- **Windows** (obbligatorio per compilare con PyInstaller)
- Python 3.10+ (compilato con `--enable-shared`)
- LibreOffice (per conversioni Office durante il build)

## Istruzioni

### 1. Preparazione

```bash
# Clonare il repository
git clone https://github.com/motthz/formazionepzz.git
cd formazionepzz

# Installare dipendenze
pip install -r requirements.txt
pip install pyinstaller
```

### Logo e icona (opzionale)

Il logo (`assets/app_icon.ico`, `assets/app_icon.png`, `assets/logo_header.png`) è già incluso.
Per rigenerarlo dopo aver modificato colori/forme in `tools/genera_logo.py` (richiede Pillow):

```bash
python tools/genera_logo.py
```

L'icona `.ico` viene incorporata automaticamente nell'eseguibile dal file `.spec`.

### 2. Build

```bash
# Rigenerare l'eseguibile
python -m PyInstaller FormazioniPZZ.spec
```

### 2b. Programma di installazione

Dopo aver compilato l'eseguibile, crea il setup (installa per l'utente corrente e mette il collegamento sul desktop):

```bash
python -m PyInstaller --noconfirm installer/installer.spec
copy dist\FormazioniPZZ_Setup.exe release\
```

### 3. Copia

L'eseguibile verrà creato in `dist/FormazioniPZZ/`. Copia il file `FormazioniPZZ.exe` nella cartella `release/`:

```bash
copy dist\FormazioniPZZ\FormazioniPZZ.exe release\
```

---

## Note sulle Ottimizzazioni

Questa versione include:

✅ **Riduzione frame landscape**: -20% di dimensioni (263mm → 210mm, 174mm → 139mm)  
✅ **Margini adattivi**: Landscape usa margini ridotti (14mm → 10mm)  
✅ **Batch processing**: Conversioni Office più veloci  
✅ **Deepcopy intelligente**: Solo quando necessario (tabelle/contenuti grandi)  
✅ **Font ottimizzato**: Landscape usa font 7.5pt (da 8pt)

**Impatto**: ~30-40% più veloce per file orizzontali
