"""Configurazione comune dei test (pytest)."""

import sys
from pathlib import Path

import pytest

TESTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TESTS_DIR))
sys.path.insert(0, str(TESTS_DIR.parent))


@pytest.fixture
def tmp(tmp_path: Path) -> Path:
    """Cartella temporanea (nome usato dai test del motore)."""
    return tmp_path
