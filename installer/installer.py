"""Installer grafico di Formazioni PZZ.

Installa il programma per l'utente corrente (nessun permesso di amministratore)
in %LOCALAPPDATA%\\Programs\\FormazioniPZZ, crea il collegamento sul desktop e nel
menu Start e registra il disinstallatore in "App installate" di Windows.

Compilazione: python -m PyInstaller --noconfirm installer/installer.spec
Opzioni da riga di comando (per installazioni automatiche):
    --silent           installa senza interfaccia
    --dir PERCORSO     cartella di installazione
    --no-shortcuts     non crea i collegamenti (desktop e menu Start)
    --no-register      non registra il disinstallatore in Windows
    --uninstall        disinstalla (usato da "App installate")
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

if not getattr(sys, "frozen", False):
    # Avvio da sorgente: ui_kit.py e' nella cartella del progetto
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
try:
    from ui_kit import UiKit
except Exception:  # senza Pillow l'installer funziona con lo stile base
    UiKit = None

APP_NAME = "Formazioni PZZ"
APP_ID = "FormazioniPZZ"
APP_EXE = "FormazioniPZZ.exe"
UNINSTALLER_EXE = "Disinstalla.exe"
INTERNAL_DIR = "_internal"  # librerie della versione "cartella" di PyInstaller
PUBLISHER = "PZZ"
UNINSTALL_KEY = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_ID}"
DEFAULT_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "Programs" / APP_ID
# Dati dell'utente: copiati solo se mancanti, mai sovrascritti in un aggiornamento
USER_DATA = ("reparti.txt", "templates")

NAVY = "#0b2a3d"
TEAL = "#1f8a87"
TEAL_HOVER = "#187370"
GOLD = "#e0a640"
BG = "#f0f4f6"
TEXT = "#11293a"
MUTED = "#5b7280"
NO_WINDOW = 0x08000000  # CREATE_NO_WINDOW


def payload_dir() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / "payload" if (base / "payload").is_dir() else base / "release"


def resource(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / "assets" / name


# ------------------------------ Windows helpers ------------------------------

def _powershell(script: str) -> str:
    out = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True, text=True, creationflags=NO_WINDOW,
    )
    if out.returncode != 0:
        raise RuntimeError(out.stderr.strip() or "PowerShell error")
    return out.stdout.strip()


def special_folder(name: str) -> Path:
    """'Desktop' o 'Programs' (menu Start); gestisce il desktop spostato su OneDrive."""
    return Path(_powershell(f"[Environment]::GetFolderPath('{name}')"))


def create_shortcut(lnk: Path, target: Path, icon: Path) -> None:
    def q(p: Path) -> str:
        return str(p).replace("'", "''")
    _powershell(
        "$s = (New-Object -ComObject WScript.Shell).CreateShortcut('%s');"
        "$s.TargetPath = '%s'; $s.WorkingDirectory = '%s';"
        "$s.IconLocation = '%s,0'; $s.Description = '%s'; $s.Save()"
        % (q(lnk), q(target), q(target.parent), q(icon), APP_NAME)
    )


def shortcut_paths() -> list[Path]:
    paths = []
    for folder in ("Desktop", "Programs"):
        try:
            paths.append(special_folder(folder) / f"{APP_NAME}.lnk")
        except Exception:
            pass
    return paths


def register_uninstaller(install_dir: Path) -> None:
    import winreg
    size_kb = sum(f.stat().st_size for f in install_dir.rglob("*") if f.is_file()) // 1024
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY) as k:
        values = {
            "DisplayName": APP_NAME,
            "DisplayIcon": str(install_dir / APP_EXE),
            "Publisher": PUBLISHER,
            "InstallLocation": str(install_dir),
            "UninstallString": f'"{install_dir / UNINSTALLER_EXE}" --uninstall',
        }
        for name, value in values.items():
            winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)
        winreg.SetValueEx(k, "EstimatedSize", 0, winreg.REG_DWORD, size_kb)
        winreg.SetValueEx(k, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(k, "NoRepair", 0, winreg.REG_DWORD, 1)


def registered_install_dir() -> Path | None:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY) as k:
            return Path(winreg.QueryValueEx(k, "InstallLocation")[0])
    except OSError:
        return None


# --------------------------------- Install ---------------------------------

def install(install_dir: Path, shortcuts: bool, register: bool, log=lambda msg: None) -> None:
    src = payload_dir()
    # Versione "cartella" (exe + _internal, avvio rapido) o, nei pacchetti vecchi, exe singolo
    program = src / "app" if (src / "app" / APP_EXE).exists() else src
    if not (program / APP_EXE).exists():
        raise RuntimeError(f"File {APP_EXE} non trovato nel pacchetto di installazione.")
    install_dir.mkdir(parents=True, exist_ok=True)

    log("Copia del programma...")
    try:
        # Prima l'exe: se il programma è aperto fallisce subito, senza lasciare
        # l'installazione a metà.
        shutil.copy2(program / APP_EXE, install_dir / APP_EXE)
        internal = install_dir / INTERNAL_DIR
        shutil.rmtree(internal, ignore_errors=True)
        if (program / INTERNAL_DIR).is_dir():
            shutil.copytree(program / INTERNAL_DIR, internal)
    except PermissionError:
        raise RuntimeError(f"{APP_NAME} è in esecuzione: chiudilo e riprova.") from None

    log("Copia dei file di esempio...")
    for name in USER_DATA:
        s, d = src / name, install_dir / name
        if d.exists() or not s.exists():
            continue
        if s.is_dir():
            shutil.copytree(s, d)
        else:
            shutil.copy2(s, d)
    (install_dir / "output").mkdir(exist_ok=True)

    if getattr(sys, "frozen", False):
        # L'installer stesso fa da disinstallatore
        try:
            shutil.copy2(sys.executable, install_dir / UNINSTALLER_EXE)
        except (PermissionError, shutil.SameFileError):
            pass

    if shortcuts:
        log("Creazione dei collegamenti...")
        for lnk in shortcut_paths():
            create_shortcut(lnk, install_dir / APP_EXE, install_dir / APP_EXE)

    if register:
        log("Registrazione in Windows...")
        register_uninstaller(install_dir)
    log("Installazione completata.")


def uninstall(install_dir: Path, remove_data: bool) -> None:
    for lnk in shortcut_paths():
        lnk.unlink(missing_ok=True)
    try:
        import winreg
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY)
    except OSError:
        pass
    (install_dir / APP_EXE).unlink(missing_ok=True)
    shutil.rmtree(install_dir / INTERNAL_DIR, ignore_errors=True)
    if remove_data:
        for name in ("templates", "output"):
            shutil.rmtree(install_dir / name, ignore_errors=True)
        for name in ("reparti.txt", "settings.json", ".formazioni_history.json", ".template_hashes.json"):
            (install_dir / name).unlink(missing_ok=True)
    # Il disinstallatore in uso non può cancellarsi da solo: lo fa cmd dopo l'uscita
    script = f'ping -n 3 127.0.0.1 >nul & del /f /q "{install_dir / UNINSTALLER_EXE}"'
    if remove_data:
        script += f' & rmdir /s /q "{install_dir}"'
    else:
        script += f' & rmdir "{install_dir}"'
    subprocess.Popen(f"cmd /c {script}", creationflags=NO_WINDOW, cwd=tempfile.gettempdir())


# ----------------------------------- GUI -----------------------------------

class InstallerUI:
    def __init__(self, root: tk.Tk, uninstall_mode: bool) -> None:
        self.root = root
        self.uninstall_mode = uninstall_mode
        root.title(f"Disinstalla {APP_NAME}" if uninstall_mode else f"Installazione di {APP_NAME}")
        root.configure(bg=BG)
        root.resizable(False, False)
        try:
            root.iconbitmap(default=str(resource("app_icon.ico")))
        except Exception:
            pass

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", foreground=MUTED, font=("Segoe UI", 9))
        style.configure("TCheckbutton", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        style.map("TCheckbutton", background=[("active", BG)])
        style.configure("Primary.TButton", background=TEAL, foreground="white", borderwidth=0,
                        font=("Segoe UI Semibold", 10, "bold"), padding=(22, 9))
        style.map("Primary.TButton", background=[("disabled", "#9fbcbc"), ("active", TEAL_HOVER)])
        style.configure("Secondary.TButton", background="#e2eaed", foreground=TEXT, borderwidth=0,
                        font=("Segoe UI Semibold", 9), padding=(14, 7))
        style.map("Secondary.TButton", background=[("active", "#d3dee3")])
        style.configure("TEntry", fieldbackground="white", bordercolor="#d3dee3", padding=6)
        style.configure("Horizontal.TProgressbar", troughcolor="#e2eaed", background=TEAL,
                        bordercolor="#e2eaed", lightcolor=TEAL, darkcolor=TEAL, thickness=8)

        # Stessa grafica dell'app: pulsanti, campi e caselle arrotondati
        self.kit = None
        if UiKit is not None:
            try:
                self.kit = UiKit(root)
                self.kit.install(style, {
                    "surface": BG, "text": TEXT, "muted": MUTED, "gold": GOLD,
                    "primary": TEAL, "primary_hover": TEAL_HOVER, "primary_press": "#125a58",
                    "on_primary": "#ffffff", "accent": GOLD, "accent_hover": "#d39a33",
                    "accent_press": "#bf8628", "on_accent": "#2a1a00",
                    "secondary": "#ffffff", "secondary_hover": "#f3f7f8", "secondary_press": "#e6eef1",
                    "on_secondary": TEXT, "secondary_border": "#d3dee3",
                    "field": "#ffffff", "field_border": "#d3dee3", "field_hover": "#a9bcc4",
                    "focus": TEAL, "check_border": "#9fb3bb",
                    "disabled_bg": "#e6edf0", "disabled_fg": "#9aabb2",
                    "select_bg": "#cdeae6", "select_fg": TEXT,
                    "trough": "#e2eaed", "thumb": "#c3d1d7", "thumb_hover": "#9fb4bc",
                    "count_bg": "#e2f3f0", "gold_bg": "#fcf1dd", "secure_bg": "#e2f3f0",
                    "header_bg": NAVY, "header_field": "#14374d", "header_border": "#255069",
                    "header_hover": "#3a6a85", "header_fg": "#f1f6f8",
                })
            except Exception:
                self.kit = None

        self._build_header()
        self.body = ttk.Frame(root, padding=(32, 22, 32, 24))
        self.body.pack(fill="both", expand=True)

        existing = registered_install_dir()
        self.dir_var = tk.StringVar(value=str(existing or DEFAULT_DIR))
        self.shortcuts_var = tk.BooleanVar(value=True)
        self.launch_var = tk.BooleanVar(value=True)
        self.remove_data_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar()

        if uninstall_mode:
            self._build_uninstall_page()
        else:
            self._build_install_page(update=existing is not None)
        root.update_idletasks()
        w, h = 560, root.winfo_reqheight()
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 3
        root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_header(self) -> None:
        c = tk.Canvas(self.root, height=96, bg=NAVY, highlightthickness=0)
        c.pack(fill="x")
        if self.kit is not None:
            self._header_bg = self.kit.header_image(1400, 96, NAVY, TEAL, GOLD, TEAL, flat_from=0.9)
            c.create_image(0, 0, image=self._header_bg, anchor="nw")
        try:
            self._logo = tk.PhotoImage(file=str(resource("logo_header.png")))
            c.create_image(32, 48, image=self._logo, anchor="w")
            x = 32 + self._logo.width() + 18
        except Exception:
            x = 32
        c.create_text(x, 30, anchor="nw", text="INSTALLAZIONE" if not self.uninstall_mode else "DISINSTALLAZIONE",
                      fill=GOLD, font=("Segoe UI", 8, "bold"))
        c.create_text(x, 46, anchor="nw", text=APP_NAME, fill="white",
                      font=("Segoe UI Semibold", 18, "bold"))
        if self.kit is None:
            c.create_rectangle(0, 93, 600, 96, fill=GOLD, outline="")

    def _build_install_page(self, update: bool) -> None:
        b = self.body
        intro = ("È già installata una versione: verrà aggiornata mantenendo modelli e impostazioni."
                 if update else
                 "Il programma verrà installato per l'utente corrente. Non servono permessi di amministratore.")
        ttk.Label(b, text=intro, wraplength=490, justify="left").pack(anchor="w")

        ttk.Label(b, text="Cartella di installazione", font=("Segoe UI Semibold", 9, "bold")
                  ).pack(anchor="w", pady=(18, 4))
        row = ttk.Frame(b)
        row.pack(fill="x")
        self.dir_entry = ttk.Entry(row, textvariable=self.dir_var, font=("Segoe UI", 9))
        self.dir_entry.pack(side="left", fill="x", expand=True)
        self.browse_btn = ttk.Button(row, text="Sfoglia...", style="Secondary.TButton", command=self._browse)
        self.browse_btn.pack(side="left", padx=(8, 0))

        ttk.Checkbutton(b, text="Crea collegamento sul desktop e nel menu Start",
                        variable=self.shortcuts_var).pack(anchor="w", pady=(16, 2))
        ttk.Checkbutton(b, text=f"Avvia {APP_NAME} al termine", variable=self.launch_var).pack(anchor="w")

        self.progress = ttk.Progressbar(b, mode="determinate", maximum=1, value=0)
        self.progress.pack(fill="x", pady=(20, 4))
        ttk.Label(b, textvariable=self.status_var, style="Muted.TLabel").pack(anchor="w")

        btns = ttk.Frame(b)
        btns.pack(fill="x", pady=(14, 0))
        self.action_btn = ttk.Button(btns, text="Aggiorna" if update else "Installa",
                                     style="Primary.TButton", command=self._start_install)
        self.action_btn.pack(side="right")
        self.cancel_btn = ttk.Button(btns, text="Annulla", style="Secondary.TButton", command=self.root.destroy)
        self.cancel_btn.pack(side="right", padx=(0, 8))

    def _build_uninstall_page(self) -> None:
        b = self.body
        ttk.Label(b, text=f"{APP_NAME} verrà rimosso da:\n{self.dir_var.get()}",
                  wraplength=490, justify="left").pack(anchor="w")
        ttk.Checkbutton(b, text="Elimina anche modelli, PDF generati e impostazioni",
                        variable=self.remove_data_var).pack(anchor="w", pady=(16, 0))
        btns = ttk.Frame(b)
        btns.pack(fill="x", pady=(22, 0))
        ttk.Button(btns, text="Disinstalla", style="Primary.TButton", command=self._do_uninstall).pack(side="right")
        ttk.Button(btns, text="Annulla", style="Secondary.TButton", command=self.root.destroy
                   ).pack(side="right", padx=(0, 8))

    def _browse(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.dir_var.get())
        if chosen:
            path = Path(chosen)
            if path.name.lower() != APP_ID.lower():
                path = path / APP_ID
            self.dir_var.set(str(path))

    def _log(self, msg: str) -> None:
        self.root.after(0, self.status_var.set, msg)

    def _start_install(self) -> None:
        for w in (self.action_btn, self.cancel_btn, self.browse_btn, self.dir_entry):
            w.state(["disabled"])
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        target = Path(self.dir_var.get()).expanduser()

        def work() -> None:
            try:
                install(target, self.shortcuts_var.get(), register=True, log=self._log)
                self.root.after(0, self._finished, target, None)
            except Exception as exc:  # mostrato all'utente
                self.root.after(0, self._finished, target, exc)

        threading.Thread(target=work, daemon=True).start()

    def _finished(self, target: Path, error: Exception | None) -> None:
        self.progress.stop()
        if error is not None:
            self.status_var.set("Installazione non riuscita.")
            messagebox.showerror(APP_NAME, f"Installazione non riuscita:\n{error}")
            for w in (self.action_btn, self.cancel_btn, self.browse_btn, self.dir_entry):
                w.state(["!disabled"])
            return
        self.progress.configure(mode="determinate", maximum=1, value=1)
        msg = f"{APP_NAME} è stato installato."
        if self.shortcuts_var.get():
            msg += "\nTrovi il collegamento sul desktop."
        messagebox.showinfo(APP_NAME, msg)
        if self.launch_var.get():
            subprocess.Popen([str(target / APP_EXE)], cwd=str(target))
        self.root.destroy()

    def _do_uninstall(self) -> None:
        if not messagebox.askyesno(APP_NAME, f"Vuoi davvero disinstallare {APP_NAME}?"):
            return
        try:
            uninstall(Path(self.dir_var.get()), self.remove_data_var.get())
        except PermissionError:
            messagebox.showerror(APP_NAME, f"{APP_NAME} è in esecuzione: chiudilo e riprova.")
            return
        messagebox.showinfo(APP_NAME, f"{APP_NAME} è stato disinstallato.")
        self.root.destroy()


def _arg_value(args: list[str], flag: str) -> str | None:
    if flag in args:
        i = args.index(flag)
        if i + 1 < len(args):
            return args[i + 1]
    return None


def main() -> None:
    args = sys.argv[1:]
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    if "--silent" in args:
        target = Path(_arg_value(args, "--dir") or registered_install_dir() or DEFAULT_DIR)
        if "--uninstall" in args:
            uninstall(target, remove_data=False)
        else:
            install(target, shortcuts="--no-shortcuts" not in args,
                    register="--no-register" not in args, log=print)
        return
    root = tk.Tk()
    InstallerUI(root, uninstall_mode="--uninstall" in args)
    root.mainloop()


if __name__ == "__main__":
    main()
