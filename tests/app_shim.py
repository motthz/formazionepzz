"""Vista "a modulo unico" del pacchetto formazioni, per gli script di test storici.

Leggere un nome lo cerca in tutti i moduli del pacchetto. Assegnarlo lo sostituisce in
ogni modulo che lo usa: i test cosi' possono dirottare percorsi come SETTINGS_FILE o
HISTORY_FILE su una cartella temporanea, senza toccare i dati veri.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import formazioni.ui.app  # noqa: E402,F401 - carica tutti i moduli del pacchetto


def _modules():
    return [module for name, module in sorted(sys.modules.items())
            if module is not None and (name == "formazioni" or name.startswith("formazioni."))]


class _AppShim:
    def __getattr__(self, name):
        for module in _modules():
            if name in vars(module):
                return vars(module)[name]
        raise AttributeError(name)

    def __setattr__(self, name, value):
        targets = [module for module in _modules() if name in vars(module)]
        if not targets:
            raise AttributeError(name)
        for module in targets:
            setattr(module, name, value)


app = _AppShim()
