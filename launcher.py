"""Avvia Formazioni PZZ e prepara automaticamente le dipendenze locali."""

from __future__ import annotations

import os
import subprocess
import sys
import venv
from pathlib import Path


ROOT = Path(__file__).resolve().parent
VENV = ROOT / ".venv"
REQUIREMENTS = ROOT / "requirements.txt"


def venv_python() -> Path:
    if os.name == "nt":
        return VENV / "Scripts" / "python.exe"
    return VENV / "bin" / "python"


def running_in_project_venv() -> bool:
    return Path(sys.prefix).resolve() == VENV.resolve()


def install_dependencies() -> None:
    if running_in_project_venv():
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS)]
        )
        return

    if not VENV.exists():
        print("Creo l'ambiente locale di Formazioni PZZ...")
        venv.EnvBuilder(with_pip=True).create(VENV)

    python = venv_python()
    print("Controllo le dipendenze necessarie...")
    subprocess.check_call([str(python), "-m", "pip", "install", "-r", str(REQUIREMENTS)])
    os.execv(str(python), [str(python), str(ROOT / "app.py"), *sys.argv[1:]])


if __name__ == "__main__":
    install_dependencies()