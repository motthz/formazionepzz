# -*- mode: python ; coding: utf-8 -*-

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

_datas = _lang_files + _asset_files

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=_datas,
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
    a.binaries,
    a.datas,
    [],
    name='FormazioniPZZ',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=version_resource('Formazioni PZZ', 'FormazioniPZZ.exe'),
    icon=os.path.join('assets', 'app_icon.ico'),
)
