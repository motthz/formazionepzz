"""Esegue gli script di test storici (controlli scritti come script) come test pytest."""

import os
import subprocess
import sys
from pathlib import Path

import pytest

TESTS_DIR = Path(__file__).resolve().parent


@pytest.mark.parametrize("script", ["legacy_completo.py", "legacy_interfaccia.py"])
def test_script_storico(script: str) -> None:
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    result = subprocess.run([sys.executable, str(TESTS_DIR / script)], capture_output=True,
                            text=True, encoding="utf-8", errors="replace", env=env, timeout=1200)
    summary = [line for line in result.stdout.splitlines() if "RIEPILOGO" in line or "[FAIL]" in line]
    assert result.returncode == 0, "\n".join(summary) or result.stderr[-3000:]
