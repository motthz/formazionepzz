# -*- mode: python ; coding: utf-8 -*-
# Installer di Formazioni PZZ. Va compilato DOPO FormazioniPZZ.spec:
#   python -m PyInstaller --noconfirm FormazioniPZZ.spec
#   python -m PyInstaller --noconfirm installer/installer.spec
# Risultato: dist/FormazioniPZZ_Setup.exe

import os

ROOT = os.path.abspath(os.path.join(SPECPATH, '..'))

_app_dir = os.path.join(ROOT, 'dist', 'FormazioniPZZ_app')
_app_exe = os.path.join(ROOT, 'dist', 'FormazioniPZZ.exe')
if not os.path.exists(_app_exe):
    _app_exe = os.path.join(ROOT, 'release', 'FormazioniPZZ.exe')

# Preferita la versione "cartella" (FormazioniPZZ_dir.spec): avvio piu' rapido.
# In mancanza si installa l'exe singolo, come nelle versioni precedenti.
_program = ((_app_dir, os.path.join('payload', 'app')) if os.path.isdir(_app_dir)
            else (_app_exe, 'payload'))

_datas = [
    _program,
    (os.path.join(ROOT, 'release', 'reparti.txt'), 'payload'),
    (os.path.join(ROOT, 'release', 'version.json'), 'payload'),  # DisplayVersion
    (os.path.join(ROOT, 'release', 'templates'), os.path.join('payload', 'templates')),
    (os.path.join(ROOT, 'assets', 'app_icon.ico'), 'assets'),
    (os.path.join(ROOT, 'assets', 'logo_header.png'), 'assets'),
]

a = Analysis(
    [os.path.join(SPECPATH, 'installer.py')],
    pathex=[ROOT],
    binaries=[],
    datas=_datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['openpyxl', 'docx', 'reportlab', 'pypdf', 'numpy'],
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
    name='FormazioniPZZ_Setup',
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
    icon=os.path.join(ROOT, 'assets', 'app_icon.ico'),
)
