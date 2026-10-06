"""Risorsa "versione" di Windows per gli exe (Proprieta' > Dettagli): nome del
prodotto, editore, descrizione e numero di versione. Un exe senza queste
informazioni sembra anonimo e gli antivirus lo giudicano piu' sospetto.
Usato dai file .spec di PyInstaller."""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def app_version() -> str:
    data = json.loads((ROOT / "release" / "version.json").read_text(encoding="utf-8-sig"))
    return str(data["version"])


def version_resource(description: str, filename: str):
    from PyInstaller.utils.win32.versioninfo import (
        FixedFileInfo,
        StringFileInfo,
        StringStruct,
        StringTable,
        VarFileInfo,
        VarStruct,
        VSVersionInfo,
    )

    numbers = tuple(([int(p) for p in app_version().split(".") if p.isdigit()] + [0, 0, 0, 0])[:4])
    text = ".".join(str(n) for n in numbers)
    return VSVersionInfo(
        ffi=FixedFileInfo(filevers=numbers, prodvers=numbers),
        kids=[
            StringFileInfo([StringTable("041004B0", [  # italiano, Unicode
                StringStruct("CompanyName", "PZZ"),
                StringStruct("FileDescription", description),
                StringStruct("FileVersion", text),
                StringStruct("InternalName", Path(filename).stem),
                StringStruct("LegalCopyright", "Copyright (c) 2026 motthz - licenza MIT"),
                StringStruct("OriginalFilename", filename),
                StringStruct("ProductName", "Formazioni PZZ"),
                StringStruct("ProductVersion", text),
            ])]),
            VarFileInfo([VarStruct("Translation", [0x0410, 1200])]),
        ],
    )
