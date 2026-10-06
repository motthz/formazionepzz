"""Sessioni di Word/Excel: istanza riusata tra le conversioni, Word ed Excel in
parallelo, chiusura per inattivita' e niente tentativi ripetuti se Office non parte.
Word ed Excel sono simulati: i test girano anche senza Office installato."""

from __future__ import annotations

import threading
import time
from pathlib import Path

import pytest

pytest.importorskip("pythoncom")

from formazioni import office  # noqa: E402


class _FakeDocument:
    def __init__(self, app: _FakeApp, path: str) -> None:
        self.app, self.path = app, path

    def ExportAsFixedFormat(self, *args) -> None:  # noqa: N802 - API di Office
        target = args[0] if self.app.kind == "word" else args[1]
        self.app.threads.add(threading.get_ident())
        time.sleep(self.app.delay)
        Path(target).write_bytes(b"%PDF-1.4 finto")

    def Close(self, _save) -> None:  # noqa: N802
        pass


class _FakeCollection:
    def __init__(self, app: _FakeApp) -> None:
        self.app = app

    def Open(self, path, *_args) -> _FakeDocument:  # noqa: N802
        return _FakeDocument(self.app, path)


class _FakeApp:
    def __init__(self, kind: str, delay: float = 0.0) -> None:
        self.kind, self.delay = kind, delay
        self.Name = kind
        self.quit_called = False
        self.threads: set[int] = set()
        self.Documents = self.Workbooks = _FakeCollection(self)

    def Quit(self) -> None:  # noqa: N802
        self.quit_called = True


@pytest.fixture
def fake_office(monkeypatch):
    started: list[_FakeApp] = []

    def starter(kind):
        def start():
            app = _FakeApp(kind, delay=0.3)
            started.append(app)
            return app
        return start

    monkeypatch.setattr(office, "_start_word", starter("word"))
    monkeypatch.setattr(office, "_start_excel", starter("excel"))
    office.shutdown_office()
    monkeypatch.setattr(office, "_SESSIONS", {})
    yield started
    office.shutdown_office()


def _sources(tmp: Path, names: list[str]) -> list[Path]:
    paths = []
    for name in names:
        path = tmp / name
        path.write_bytes(b"x")
        paths.append(path)
    return paths


def test_word_and_excel_run_in_parallel_and_stay_open(tmp: Path, fake_office) -> None:
    sources = _sources(tmp, ["a.docx", "b.xlsx"])
    ticks = []
    start = time.perf_counter()
    converted = office._convert_with_ms_office_batch(sources, tmp, "pdf", lambda: ticks.append(1))
    elapsed = time.perf_counter() - start
    assert set(converted) == set(sources)
    assert all(path.exists() for path in converted.values())
    assert len(ticks) == 2
    assert elapsed < 0.55, f"Word ed Excel devono lavorare in parallelo ({elapsed:.2f}s)"

    # Seconda generazione: stesse istanze, nessun nuovo avvio
    office._convert_with_ms_office_batch(_sources(tmp, ["c.docx"]), tmp, "pdf")
    assert [app.kind for app in fake_office] == ["word", "excel"]
    assert not any(app.quit_called for app in fake_office)

    office.shutdown_office()
    assert all(app.quit_called for app in fake_office)


def test_idle_session_closes_and_restarts(tmp: Path, fake_office, monkeypatch) -> None:
    monkeypatch.setattr(office, "OFFICE_IDLE_SECONDS", 0.2)
    office._convert_with_ms_office_batch(_sources(tmp, ["a.docx"]), tmp, "pdf")
    deadline = time.monotonic() + 3
    while not fake_office[0].quit_called and time.monotonic() < deadline:
        time.sleep(0.05)
    assert fake_office[0].quit_called, "Word va chiuso dopo l'inattivita'"
    office._convert_with_ms_office_batch(_sources(tmp, ["b.docx"]), tmp, "pdf")
    assert len(fake_office) == 2


def test_failed_start_is_not_retried_immediately(tmp: Path, monkeypatch) -> None:
    attempts = []

    def broken():
        attempts.append(1)
        raise OSError("Esecuzione del server non riuscita")

    monkeypatch.setattr(office, "_start_word", broken)
    office.shutdown_office()
    monkeypatch.setattr(office, "_SESSIONS", {})
    try:
        for name in ("a.docx", "b.docx"):
            with pytest.raises(OSError):
                office._convert_with_ms_office_batch(_sources(tmp, [name]), tmp, "pdf")
        assert len(attempts) == 1
    finally:
        office.shutdown_office()
