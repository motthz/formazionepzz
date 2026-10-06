# -*- mode: python ; coding: utf-8 -*-
# Versione "cartella" (onedir) usata dall'installer: si avvia in meno tempo
# dell'exe singolo, che a ogni apertura deve prima estrarsi in una cartella temporanea.
#   python -m PyInstaller --noconfirm FormazioniPZZ_dir.spec
# Risultato: dist/FormazioniPZZ_app/FormazioniPZZ.exe (+ _internal)

import os
import sys

# Niente compressione UPX e informazioni di versione nell'exe: gli exe compressi
# e anonimi sono quelli che gli antivirus segnalano piu' spesso come sospetti.
sys.path.insert(0, os.path.join(SPECPATH, 'tools'))
from versione_exe import version_resource  # noqa: E402
_lang_files = [
    (os.path.join('lang', f), os.path.join('lang'))
    for f in os.listdir('lang') if f.lower().endswith('.json')
] if os.path.isdir('lang') else []

_asset_files = []
if os.path.isdir('assets'):
    for _f in os.listdir('assets'):
        if _f.lower().endswith(('.ico', '.png')):
            _asset_files.append((os.path.join('assets', _f), os.path.join('assets')))

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=_lang_files + _asset_files,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='FormazioniPZZ',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=version_resource('Formazioni PZZ', 'FormazioniPZZ.exe'),
    icon=os.path.join('assets', 'app_icon.ico'),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='FormazioniPZZ_app',
)
