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
# Exe portatile (un solo file)
python -m PyInstaller --noconfirm FormazioniPZZ.spec
# Versione "cartella", installata dal setup: si avvia molto più in fretta
python -m PyInstaller --noconfirm FormazioniPZZ_dir.spec
# Programma di installazione (include la versione cartella)
python -m PyInstaller --noconfirm installer/installer.spec
```

I file finiti sono `dist/FormazioniPZZ.exe` e `dist/FormazioniPZZ_Setup.exe`. Non vanno salvati nel repository: si pubblicano come allegati di una Release GitHub. Il workflow lo fa da solo quando si invia un tag `vX.Y.Z`; i passaggi sono in CONTRIBUTING.md, sezione "Pubblicare una versione".

### Antivirus

Gli exe di PyInstaller vengono spesso scambiati per malware. Per ridurre i falsi positivi:

- i file `.spec` non usano UPX (`upx=False`) e aggiungono nome, editore e versione dell'exe (`tools/versione_exe.py`, versione letta da `release/version.json`);
- il workflow compila il bootloader di PyInstaller a ogni build (`PYINSTALLER_COMPILE_BOOTLOADER=1`) invece di usare quello precompilato. In locale serve Visual Studio con gli strumenti C++:

```bash
set PYINSTALLER_COMPILE_BOOTLOADER=1
pip install --force-reinstall --no-deps --no-cache-dir --no-binary pyinstaller pyinstaller==6.22.2
```

- l'exe portatile (un solo file, si estrae in una cartella temporanea a ogni avvio) è quello segnalato più spesso: per gli utenti è preferibile l'installer;
- se un antivirus blocca comunque una versione, l'exe si può inviare come falso positivo (per Microsoft Defender: <https://www.microsoft.com/wdsi/filesubmission>). La soluzione definitiva è firmare gli exe con un certificato di firma del codice.

---

## Note sulle Ottimizzazioni

Questa versione include:

✅ **Riduzione frame landscape**: -20% di dimensioni (263mm → 210mm, 174mm → 139mm)  
✅ **Margini adattivi**: Landscape usa margini ridotti (14mm → 10mm)  
✅ **Batch processing**: Conversioni Office più veloci  
✅ **Deepcopy intelligente**: Solo quando necessario (tabelle/contenuti grandi)  
✅ **Font ottimizzato**: Landscape usa font 7.5pt (da 8pt)

**Impatto**: ~30-40% più veloce per file orizzontali
