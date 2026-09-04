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

### 2. Build

```bash
# Rigenerare l'eseguibile
python -m PyInstaller FormazioniPZZ.spec
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
