# -*- mode: python ; coding: utf-8 -*-

import os
_lang_files = [
    (os.path.join('lang', f), os.path.join('lang'))
    for f in os.listdir('lang') if f.lower().endswith('.json')
] if os.path.isdir('lang') else []


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=_lang_files,
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
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
