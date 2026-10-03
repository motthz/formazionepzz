"""Aggiornamento in un clic: scaricamento del setup e passaggio di consegne all'installer."""

from __future__ import annotations

import io
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

from formazioni import updates
from formazioni.ui import settings as settings_ui


class _FakeResponse(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def test_download_setup_scrive_il_file_solo_a_download_completo(tmp_path, monkeypatch):
    payload = b"MZ" + b"\0" * 2_000_000
    monkeypatch.setattr(updates.urllib.request, "urlopen", lambda *a, **k: _FakeResponse(payload))
    setup = updates.download_setup("https://example/setup.exe", tmp_path, "9.9.9")
    assert setup.name == "FormazioniPZZ_Setup-9.9.9.exe"
    assert setup.read_bytes() == payload
    assert not list(tmp_path.glob("*.part"))


def test_download_setup_rifiuta_una_pagina_di_errore(tmp_path, monkeypatch):
    monkeypatch.setattr(updates.urllib.request, "urlopen",
                        lambda *a, **k: _FakeResponse(b"<html>Not Found</html>"))
    with pytest.raises(ValueError):
        updates.download_setup("https://example/setup.exe", tmp_path, "9.9.9")
    assert not list(tmp_path.iterdir())


def test_copia_installata_riconosciuta(tmp_path, monkeypatch):
    assert not updates.is_installed_copy(tmp_path)  # da sorgente: mai aggiornamento automatico
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    assert not updates.is_installed_copy(tmp_path)  # exe portatile
    (tmp_path / "_internal").mkdir()
    (tmp_path / "Disinstalla.exe").write_bytes(b"")
    assert updates.is_installed_copy(tmp_path)


def _fake_app(worker_active=False):
    calls = SimpleNamespace(destroyed=False, persisted=False, scheduled=[])
    root = SimpleNamespace(destroy=lambda: setattr(calls, "destroyed", True),
                           after=lambda ms, fn: calls.scheduled.append(ms))
    app = SimpleNamespace(root=root, _worker_active=worker_active,
                          _persist_settings=lambda: setattr(calls, "persisted", True))
    return app, calls


def test_run_setup_avvia_installer_e_chiude_app(monkeypatch):
    launched = []
    monkeypatch.setattr(settings_ui.subprocess, "Popen", lambda args, **k: launched.append(args))
    app, calls = _fake_app()
    settings_ui.SettingsMixin._run_setup(app, Path("C:/tmp/FormazioniPZZ_Setup-9.9.9.exe"))
    args = launched[0]
    assert args[1:3] == ["--silent", "--dir"] and args[3] == str(settings_ui.APP_DIR)
    assert "--relaunch" in args and "--no-shortcuts" in args
    assert args[args.index("--wait-pid") + 1].isdigit()
    assert calls.persisted and calls.destroyed


def test_run_setup_aspetta_la_fine_della_generazione(monkeypatch):
    launched = []
    monkeypatch.setattr(settings_ui.subprocess, "Popen", lambda args, **k: launched.append(args))
    app, calls = _fake_app(worker_active=True)
    settings_ui.SettingsMixin._run_setup(app, Path("setup.exe"))
    assert not launched and not calls.destroyed and calls.scheduled
