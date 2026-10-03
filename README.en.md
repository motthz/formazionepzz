# Formazioni PZZ

[![Tests and build](https://github.com/motthz/formazionepzz/actions/workflows/build-windows-exe.yml/badge.svg)](https://github.com/motthz/formazionepzz/actions/workflows/build-windows-exe.yml)
[![Latest version](https://img.shields.io/github/v/release/motthz/formazionepzz?label=version)](https://github.com/motthz/formazionepzz/releases/latest)
[![MIT License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

🇮🇹 [Versione italiana](README.md)

A Windows program that builds a **training dossier as a PDF** for new employees in a few seconds. Pick a name, start date and department: Formazioni PZZ takes that department's Word, Excel and PDF templates, fills in `*nome*` (name) and `*data*` (date) and merges them into a single print-ready file. Nothing leaves the computer.

![Main window](docs/img/principale.png)

📖 **[User guide (Italian)](https://motthz.github.io/formazionepzz/)** · 🆕 **[Changelog (Italian)](CHANGELOG.md)** · 🐞 **[Report a problem](https://github.com/motthz/formazionepzz/issues/new/choose)**

## Installation

Download **`FormazioniPZZ_Setup.exe`** from the [latest release](https://github.com/motthz/formazionepzz/releases/latest) and run it.

- Installs for the current user in `%LOCALAPPDATA%\Programs\FormazioniPZZ`, with no administrator rights needed.
- Creates desktop and Start menu shortcuts and registers the program under *Settings → Installed apps*.
- Updates arrive on their own: "Update now" downloads the new version, closes the app and reopens it updated, keeping templates, departments and settings.
- Unattended install: `FormazioniPZZ_Setup.exe --silent` (options `--dir PATH`, `--no-shortcuts`).

`FormazioniPZZ.exe` in the same release is the **portable** version: it runs without installing, but starts more slowly and is updated by hand.

The interface is available in Italian, English, German, Spanish and French.

## Main features

- **Templates per department**: the file name sets the department and number of copies (`SICUREZZA_2_SIC.docx`, `TUTTI_1_GEN.docx` for every department). Details are in [templates/README.md](templates/README.md) (Italian).
- **One person or many**: a batch list, import from CSV/Excel with a ready-made template, and conversion in a single Word session.
- **Preview**, dossier **history** with "Use this data", PDF **watermark** and **edit lock**.
- Light, dark or "System" theme, keyboard shortcuts, drag and drop of files.
- Templates without placeholders are cached; without Office, batches run in parallel.

To keep the templates' original layout you need **Microsoft Office** or **LibreOffice** (with `soffice` on the PATH). Without either, the app rebuilds the documents with ReportLab.

## Updates and data sent

At startup, at most once a day, the app reads the latest version number published on GitHub; no data is sent. This can be turned off in **Settings**, where you can also point to another source: a network folder with `version.json` and `FormazioniPZZ_Setup.exe`, or the web address of a `version.json`.

## Development

```bash
python -m pip install -r requirements-dev.txt
python app.py              # run the app
python -m pytest           # tests
python -m ruff check .     # lint
```

The code layout, the pull-request workflow and the release process are described in [CONTRIBUTING.md](CONTRIBUTING.md) (Italian).

## License

[MIT](LICENSE).
