"""Formazioni PZZ: crea dossier PDF locali a partire da template Word ed Excel."""

from __future__ import annotations

import base64
import calendar
import contextlib
import copy as _copy_mod
import csv
import ctypes
import functools
import hashlib
import json
import os
import platform
import queue
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import urllib.request
import webbrowser
from dataclasses import dataclass, replace as dc_replace
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any
import struct
import zlib

try:
    import tkinter as tk
    from tkinter import BOTH, END, LEFT, RIGHT, X, BooleanVar, StringVar, filedialog, messagebox, ttk
    from ui_kit import RoundedPanel, UiKit
except ImportError:
    tk = None  # type: ignore[assignment]
    RoundedPanel = UiKit = None  # type: ignore[assignment]
    BOTH = END = LEFT = RIGHT = X = None  # type: ignore[assignment]
    BooleanVar = StringVar = filedialog = messagebox = ttk = None  # type: ignore[assignment]
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm

if TYPE_CHECKING:
    from reportlab.platypus import Paragraph, Table


# python-docx, openpyxl e reportlab.platypus pesano quasi un secondo all'avvio:
# si importano alla prima generazione, non all'apertura della finestra.
def Document(*args, **kwargs):  # noqa: N802 - stesso nome dell'API di python-docx
    from docx import Document as _document

    return _document(*args, **kwargs)


def load_workbook(*args, **kwargs):
    from openpyxl import load_workbook as _load_workbook

    return _load_workbook(*args, **kwargs)


def _resolve_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "release":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


APP_VERSION = "2.2.0"
DEFAULT_UPDATE_SOURCE = "https://api.github.com/repos/motthz/formazionepzz/releases/latest"
APP_DIR = _resolve_app_dir()
DEFAULT_TEMPLATE_DIR = APP_DIR / "templates"
DEFAULT_OUTPUT_DIR = APP_DIR / "output"
HISTORY_FILE = APP_DIR / ".formazioni_history.json"
HISTORY_LIMIT = 500
PREVIEW_PREFIX = "formazioni_anteprima_"
DEPARTMENTS_FILE = APP_DIR / "reparti.txt"
SUPPORTED_EXTENSIONS = {".doc", ".docx", ".xls", ".xlsx", ".pdf"}
ALL_DEPARTMENT_NAMES = {"TUTTI", "TUTTE", "ALL"}
FILENAME_PATTERN = re.compile(
    r"^(?P<department>.+)_(?P<count>\d+)_(?P<code>[A-Za-z]{2,5})$",
    re.IGNORECASE,
)
SETTINGS_FILE = APP_DIR / "settings.json"
HASHES_FILE = APP_DIR / ".template_hashes.json"
# File di sola lettura (lingue, icone): nell'exe PyInstaller stanno in _MEIPASS,
# non accanto all'exe; da sorgente coincidono con APP_DIR.
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
LANG_DIR = RESOURCE_DIR / "lang" if (RESOURCE_DIR / "lang").is_dir() else APP_DIR / "lang"
ICON_DIR = RESOURCE_DIR / "assets" if (RESOURCE_DIR / "assets").is_dir() else APP_DIR / "assets"
ICON_ICO = ICON_DIR / "app_icon.ico"
ICON_PNG = ICON_DIR / "app_icon.png"
LOGO_HEADER_PNG = ICON_DIR / "logo_header.png"


def _build_in_memory_icon(width=32, height=32, bg1=(12, 34, 53), bg2=(21, 58, 82),
                          gold=(217, 161, 63)):
    """Build a pure-Python in-memory icon (RGBA) so no external deps needed."""
    pixels = bytearray()
    for y in range(height):
        for x in range(width):
            # card corners (outer/inner rounded)
            def in_rect(px, py, ox, oy, ex, ey, r):
                if not (ox <= px <= ex and oy <= py <= ey):
                    return False
                dx = dy = 0
                if px < ox + r and py < oy + r:
                    dx = ox + r - px
                    dy = oy + r - py
                    if dx * dx + dy * dy > r * r:
                        return False
                if px > ex - r and py < oy + r:
                    dx = px - (ex - r)
                    dy = oy + r - py
                    if dx * dx + dy * dy > r * r:
                        return False
                if px < ox + r and py > ey - r:
                    dx = ox + r - px
                    dy = py - (ey - r)
                    if dx * dx + dy * dy > r * r:
                        return False
                if px > ex - r and py > ey - r:
                    dx = px - (ex - r)
                    dy = py - (ey - r)
                    if dx * dx + dy * dy > r * r:
                        return False
                return True
            outer = in_rect(x, y, 1, 1, width - 2, height - 2, width // 8)
            inner = in_rect(x, y, 3, 3, width - 4, height - 4, max(2, width // 9))
            # gold band bottom
            gold_y = y >= height - (height // 3)
            if not outer:
                r = g = b = 0
                a = 0
            elif not inner:
                r, g, b = bg1
                a = 255
            elif gold_y:
                r, g, b = gold
                a = 255
            else:
                r, g, b = bg2
                a = 255
            pixels.extend((r, g, b, a))
    return bytes(pixels), width, height


def _png_from_rgba(rgba: bytes, w: int, h: int) -> bytes:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)
    raw = bytearray()
    for y in range(h):
        raw.append(0)
        row_start = y * w * 4
        raw.extend(rgba[row_start:row_start + w * 4])
    idat = zlib.compress(bytes(raw), 9)
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


def _set_app_icon(root) -> None:
    """Try to set root window icon using multiple strategies."""
    # Su Windows un AppUserModelID dedicato fa usare alla taskbar l'icona
    # dell'app invece di quella di python.exe (avvio da sorgente).
    if os.name == "nt":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("PZZ.FormazioniPZZ")
        except Exception:
            pass
    # 1) Try .ico file on disk (multi-risoluzione: nitida a ogni dimensione)
    try:
        if ICON_ICO.exists():
            root.iconbitmap(default=str(ICON_ICO))
            if os.name == "nt":
                return
    except Exception:
        pass
    # 2) Try PNG file (preferred fallback) via PhotoImage
    try:
        if ICON_PNG.exists():
            img = tk.PhotoImage(file=str(ICON_PNG))
            root.iconphoto(True, img)
            root._pzz_icon_photo = img  # keep reference
            return
    except Exception:
        pass
    # 3) Fallback: build a simple in-memory icon and feed it as PNG data
    try:
        rgba, w, h = _build_in_memory_icon()
        png_data = _png_from_rgba(rgba, w, h)
        img = tk.PhotoImage(data=png_data)
        root.iconphoto(True, img)
        root._pzz_icon_photo = img
        # Also make larger version (64x64)
        try:
            rgba2, w2, h2 = _build_in_memory_icon(64, 64)
            img2 = tk.PhotoImage(data=_png_from_rgba(rgba2, w2, h2))
            root.iconphoto(False, img2)
            root._pzz_icon_photo_big = img2
        except Exception:
            pass
    except Exception:
        pass


def _set_titlebar_dark(root, dark: bool) -> None:
    """Allinea la barra del titolo di Windows 10/11 al tema scelto."""
    if os.name != "nt":
        return
    try:
        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (nuovo / pre-20H1)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)) == 0:
                break
        # SWP_FRAMECHANGED|NOMOVE|NOSIZE|NOZORDER|NOACTIVATE: ridisegna subito la cornice
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0020 | 0x0002 | 0x0001 | 0x0004 | 0x0010)
    except Exception:
        pass

DEFAULT_LANG = "en"
DEFAULT_THEME = "light"
MONTH_KEYS = (
    "dp_jan", "dp_feb", "dp_mar", "dp_apr", "dp_may", "dp_jun",
    "dp_jul", "dp_aug", "dp_sep", "dp_oct", "dp_nov", "dp_dec",
)

_STYLE_CACHE: dict[str, object] | None = None
_TEMPLATE_STORY_CACHE: dict[tuple, list[object]] = {}


SYSTEM_THEME_POLL_MS = 4000


def windows_prefers_dark() -> bool:
    """Legge "Modalita' app" di Windows (Impostazioni > Personalizzazione > Colori)."""
    if os.name != "nt":
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def resolve_theme(preference: str) -> str:
    if preference == "system":
        return "dark" if windows_prefers_dark() else "light"
    return "dark" if preference == "dark" else "light"


# --------------------------- ICONE ----------------------------------------
# Icone a linea disegnate al volo nel colore del testo del pulsante: niente emoji,
# che Windows rende con font e colori diversi da un PC all'altro.

def draw_icon_png(name: str, color: str, size: int) -> bytes | None:
    """PNG trasparente dell'icona `name`, oppure None se l'icona non esiste."""
    import io
    import math

    from PIL import Image, ImageDraw

    big = size * 4
    s = big / 24.0  # le coordinate sono su una griglia 24x24
    w = max(2, round(2.1 * s))
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    rgb = tuple(int(color.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    ink = rgb + (255,)

    def P(*pts):  # noqa: N802 - punti sulla griglia 24x24
        return [(x * s, y * s) for x, y in pts]

    def line(*pts):
        d.line(P(*pts), fill=ink, width=w, joint="curve")
        for x, y in P(pts[0], pts[-1]):
            d.ellipse((x - w / 2, y - w / 2, x + w / 2, y + w / 2), fill=ink)

    def box(x0, y0, x1, y1, r=2.0):
        d.rounded_rectangle(P((x0, y0), (x1, y1)), radius=r * s, outline=ink, width=w)

    def circle(cx, cy, r, fill=False):
        bbox = P((cx - r, cy - r), (cx + r, cy + r))
        if fill:
            d.ellipse(bbox, fill=ink)
        else:
            d.ellipse(bbox, outline=ink, width=w)

    def arc(cx, cy, r, start, end):
        d.arc(P((cx - r, cy - r), (cx + r, cy + r)), start, end, fill=ink, width=w)

    if name == "refresh":
        arc(12, 12, 8, 40, 330)
        line((20, 4), (20, 9.5), (14.5, 9.5))
    elif name == "download":
        line((12, 3.5), (12, 15))
        line((7, 10.5), (12, 15.5), (17, 10.5))
        line((4.5, 20), (19.5, 20))
    elif name == "upload":
        line((12, 15.5), (12, 4))
        line((7, 8.5), (12, 3.5), (17, 8.5))
        line((4.5, 20), (19.5, 20))
    elif name == "plus":
        line((12, 5), (12, 19))
        line((5, 12), (19, 12))
    elif name == "minus":
        line((5, 12), (19, 12))
    elif name == "check":
        line((5, 12.5), (10, 17.5), (19, 7))
    elif name == "close":
        line((6, 6), (18, 18))
        line((18, 6), (6, 18))
    elif name == "trash":
        line((4, 6.5), (20, 6.5))
        line((9.5, 6.5), (9.5, 3.5), (14.5, 3.5), (14.5, 6.5))
        box(6, 6.5, 18, 20.5, 1.5)
        line((10, 10.5), (10, 16.5))
        line((14, 10.5), (14, 16.5))
    elif name == "file":
        line((14, 3), (6, 3), (6, 21), (18, 21), (18, 7), (14, 3), (14, 7), (18, 7))
        line((9, 12), (15, 12))
        line((9, 16), (15, 16))
    elif name == "folder":
        line((3, 19), (3, 5), (9.5, 5), (11.5, 7.5), (21, 7.5), (21, 19), (3, 19))
    elif name == "edit":
        line((5, 19), (5.8, 15), (16, 4.8), (19.2, 8), (9, 18.2), (5, 19))
        line((13.8, 7), (17, 10.2))
    elif name == "gear":
        for k in range(8):
            a = math.radians(k * 45)
            line((12 + 6.2 * math.cos(a), 12 + 6.2 * math.sin(a)),
                 (12 + 9 * math.cos(a), 12 + 9 * math.sin(a)))
        circle(12, 12, 6.2)
        circle(12, 12, 2.4)
    elif name == "clock":
        circle(12, 12, 8.5)
        line((12, 7), (12, 12), (15.5, 14))
    elif name == "eye":
        arc(12, 24, 15, 233, 307)
        arc(12, 0, 15, 53, 127)
        circle(12, 12, 3.2)
    elif name == "users":
        circle(9, 8, 3.2)
        arc(9, 20, 6.5, 200, 340)
        circle(17, 9, 2.5)
        arc(17, 19.5, 5, 230, 340)
    elif name == "list":
        for y in (6, 12, 18):
            circle(4.5, y, 1.4, fill=True)
            line((9, y), (20, y))
    elif name == "lock":
        box(5, 11, 19, 20.5, 2)
        arc(12, 8, 4, 180, 360)
        line((8, 11), (8, 8))
        line((16, 11), (16, 8))
    elif name == "table":
        box(3.5, 4.5, 20.5, 19.5, 2)
        line((3.5, 9.5), (20.5, 9.5))
        line((10, 9.5), (10, 19.5))
    else:
        return None
    img = img.resize((size, size), Image.LANCZOS)
    out = io.BytesIO()
    img.save(out, "PNG")
    return out.getvalue()


# Icona per ogni testo di pulsante (chiave di traduzione)
BUTTON_ICONS = {
    "btn_refresh": "refresh", "btn_depts": "users", "btn_modules": "file",
    "btn_settings": "gear", "btn_history": "clock", "btn_choose_folder": "folder",
    "btn_selectall": "check", "btn_selectnone": "close", "btn_generate": "download",
    "btn_preview": "eye", "bat_add_person": "plus", "bat_del_person": "minus",
    "bat_import_file": "upload", "bat_clear": "trash", "bat_run_all": "download",
    "bat_template_download": "table", "tm_add": "plus", "tm_edit": "edit",
    "tm_open": "file", "tm_delete": "trash", "tm_open_folder": "folder",
    "hi_open_pdf": "file", "hi_open_dir": "folder", "hi_reuse": "refresh",
    "hi_remove": "trash", "de_add": "plus", "de_rename": "edit", "de_delete": "trash",
    "st_check_now": "refresh", "secure_label": "lock", "bat_inline_title": "list",
}
_LEADING_SYMBOLS_RE = re.compile(r"^[^\w(«\"'*]+", re.UNICODE)


def strip_leading_symbol(text: str) -> str:
    """Toglie l'emoji iniziale dai testi tradotti ("⬇  Genera" -> "Genera")."""
    return _LEADING_SYMBOLS_RE.sub("", text)


# --------------------------- NOTIFICHE / DRAG & DROP (Windows) --------------

def notify_windows(root, title: str, message: str, seconds: int = 8) -> bool:
    """Notifica di Windows (area notifiche) se l'utente sta usando un'altra
    finestra, piu' il lampeggio dell'icona nella barra delle applicazioni."""
    if os.name != "nt":
        return False
    try:
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        if ctypes.windll.user32.GetForegroundWindow() == hwnd:
            return False

        class FLASHWINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("hwnd", ctypes.c_void_p),
                        ("dwFlags", ctypes.c_uint), ("uCount", ctypes.c_uint),
                        ("dwTimeout", ctypes.c_uint)]

        flash = FLASHWINFO(ctypes.sizeof(FLASHWINFO), hwnd, 0x3 | 0xC, 0, 0)  # ALL|TIMERNOFG
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(flash))
        import win32con  # type: ignore[import-not-found]
        import win32gui  # type: ignore[import-not-found]

        try:
            hicon = win32gui.LoadImage(0, str(ICON_ICO), win32con.IMAGE_ICON, 0, 0,
                                       win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE)
        except Exception:  # noqa: BLE001
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
        flags = win32gui.NIF_ICON | win32gui.NIF_TIP | win32gui.NIF_INFO
        data = (hwnd, 7301, flags, 0, hicon, "Formazioni PZZ",
                message[:255], seconds * 1000, title[:63], win32gui.NIIF_INFO)
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, data)

        def remove():
            try:
                win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (hwnd, 7301))
            except Exception:  # noqa: BLE001
                pass

        root.after(seconds * 1000 + 2000, remove)
        return True
    except Exception:  # noqa: BLE001 - notifica accessoria
        return False


class FileDropTarget:
    """Accetta i file trascinati da Esplora risorse su una finestra Tk (solo Windows),
    senza librerie esterne: intercetta WM_DROPFILES sulla finestra di primo livello.
    callback(paths, x_root, y_root) viene chiamata nel thread dell'interfaccia."""

    WM_DROPFILES = 0x0233
    GWLP_WNDPROC = -4

    def __init__(self, toplevel, callback) -> None:
        from ctypes import wintypes

        self.toplevel = toplevel
        self.callback = callback
        user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
        self._lresult = ctypes.c_ssize_t
        proc_type = ctypes.WINFUNCTYPE(self._lresult, wintypes.HWND, wintypes.UINT,
                                       wintypes.WPARAM, wintypes.LPARAM)
        set_long = getattr(user32, "SetWindowLongPtrW", None) or user32.SetWindowLongW
        set_long.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        set_long.restype = ctypes.c_void_p
        self._call = user32.CallWindowProcW
        self._call.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                               wintypes.WPARAM, wintypes.LPARAM]
        self._call.restype = self._lresult
        self._query = shell32.DragQueryFileW
        self._query.argtypes = [ctypes.c_void_p, wintypes.UINT, ctypes.c_wchar_p, wintypes.UINT]
        self._query.restype = wintypes.UINT
        self._point = shell32.DragQueryPoint
        self._point.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.POINT)]
        self._finish = shell32.DragFinish
        self._finish.argtypes = [ctypes.c_void_p]
        self._to_screen = user32.ClientToScreen
        self._to_screen.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.POINT)]
        self._POINT = wintypes.POINT

        self.hwnd = user32.GetParent(toplevel.winfo_id())
        self._dropped: queue.Queue[tuple[list[Path], int, int]] = queue.Queue()
        shell32.DragAcceptFiles(self.hwnd, True)
        self._proc = proc_type(self._wndproc)  # riferimento tenuto vivo dall'oggetto
        self._old = set_long(self.hwnd, self.GWLP_WNDPROC,
                             ctypes.cast(self._proc, ctypes.c_void_p).value)
        self.toplevel.after(150, self._poll)

    def _poll(self) -> None:
        while not self._dropped.empty():
            self.callback(*self._dropped.get_nowait())
        try:
            self.toplevel.after(150, self._poll)
        except tk.TclError:
            pass  # finestra chiusa

    def _wndproc(self, hwnd, msg, wparam, lparam):
        if msg != self.WM_DROPFILES:
            return self._call(self._old, hwnd, msg, wparam, lparam)
        try:
            count = self._query(wparam, 0xFFFFFFFF, None, 0)
            paths = []
            for index in range(count):
                length = self._query(wparam, index, None, 0)
                buffer = ctypes.create_unicode_buffer(length + 1)
                self._query(wparam, index, buffer, length + 1)
                paths.append(Path(buffer.value))
            point = self._POINT()
            self._point(wparam, ctypes.byref(point))
            self._to_screen(hwnd, ctypes.byref(point))
            x, y = point.x, point.y
        finally:
            self._finish(wparam)
        # Qui dentro non si puo' chiamare Tkinter (la window procedure gira dentro il
        # ciclo eventi di Tcl e ne corromperebbe lo stato): i file vanno in coda e
        # li consegna _poll dal normale ciclo di Tk.
        self._dropped.put((paths, x, y))
        return 0


# --------------------------- SETTINGS -------------------------------------

def _default_settings() -> dict[str, Any]:
    return {
        "theme": DEFAULT_THEME,
        "language": DEFAULT_LANG,
        "last_template_dir": str(DEFAULT_TEMPLATE_DIR),
        "last_output_dir": str(DEFAULT_OUTPUT_DIR),
        "pdf_watermark": "",
        "pdf_protect": False,
        "check_updates": True,
        "update_source": "",
        "last_update_check": "",
    }


def load_settings() -> dict[str, Any]:
    data = _default_settings()
    if SETTINGS_FILE.exists():
        try:
            raw = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                data.update({k: v for k, v in raw.items() if k in data})
        except (OSError, json.JSONDecodeError):
            pass
    return data


def save_settings(data: dict[str, Any]) -> None:
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass


# --------------------------- I18N -----------------------------------------

def available_languages() -> list[str]:
    if not LANG_DIR.exists():
        return [DEFAULT_LANG]
    return sorted(p.stem for p in LANG_DIR.glob("*.json"))


def load_language(code: str) -> dict[str, str]:
    fallback: dict[str, str] = {}
    if (LANG_DIR / f"{DEFAULT_LANG}.json").exists():
        try:
            fallback = json.loads(
                (LANG_DIR / f"{DEFAULT_LANG}.json").read_text(encoding="utf-8")
            )
        except (OSError, json.JSONDecodeError):
            fallback = {}
    chosen_file = LANG_DIR / f"{code}.json"
    chosen: dict[str, str] = {}
    if chosen_file.exists():
        try:
            chosen = json.loads(chosen_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            chosen = {}
    merged = dict(fallback)
    merged.update(chosen)
    return merged


# --------------------------- HASHING --------------------------------------

_HASH_CACHE: dict[str, tuple[int, int, str]] = {}


def compute_template_hash(path: Path) -> str:
    # Rileggere ogni template a ogni aggiornamento della lista e' lento con molti
    # file o cartelle di rete: si ricalcola solo se cambiano data o dimensione.
    try:
        stat = path.stat()
    except OSError:
        return ""
    key = str(path)
    cached = _HASH_CACHE.get(key)
    if cached and cached[0] == stat.st_mtime_ns and cached[1] == stat.st_size:
        return cached[2]
    hasher = hashlib.md5()
    try:
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(65536), b""):
                hasher.update(chunk)
    except OSError:
        return ""
    digest = hasher.hexdigest()
    _HASH_CACHE[key] = (stat.st_mtime_ns, stat.st_size, digest)
    return digest


def load_saved_hashes() -> dict[str, str]:
    if not HASHES_FILE.exists():
        return {}
    try:
        data = json.loads(HASHES_FILE.read_text(encoding="utf-8"))
        return {str(k): str(v) for k, v in data.items()} if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_hashes(data: dict[str, str]) -> None:
    try:
        HASHES_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass


def classify_template_hashes(
    templates: list[Any], saved: dict[str, str]
) -> dict[str, str]:
    result: dict[str, str] = {}
    for template in templates:
        key = str(template.path)
        current = compute_template_hash(template.path)
        previous = saved.get(key)
        if previous is None:
            result[key] = "new"
        elif previous == current:
            result[key] = "ok"
        else:
            result[key] = "modified"
    return result


def clear_caches() -> None:
    global _STYLE_CACHE
    _STYLE_CACHE = None
    _TEMPLATE_STORY_CACHE.clear()
    _office_command.cache_clear()


# --------------------------- BUSINESS -------------------------------------

def load_departments_from_file() -> list[str]:
    if not DEPARTMENTS_FILE.exists():
        return []
    try:
        raw = DEPARTMENTS_FILE.read_text(encoding="utf-8")
    except OSError:
        return []
    departments: list[str] = []
    seen: set[str] = set()
    for line in raw.splitlines():
        name = line.strip()
        if not name or name.startswith("#"):
            continue
        key = name.upper()
        if key not in seen:
            seen.add(key)
            departments.append(key)
    return departments


@dataclass(frozen=True)
class TemplateFile:
    path: Path
    department: str
    copies: int
    code: str

    @property
    def display_name(self) -> str:
        return self.path.name

    @property
    def is_for_every_department(self) -> bool:
        return self.department.upper() in ALL_DEPARTMENT_NAMES


def parse_template(path: Path) -> TemplateFile | None:
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return None
    match = FILENAME_PATTERN.fullmatch(path.stem)
    if not match:
        return None
    return TemplateFile(
        path=path,
        department=match.group("department"),
        copies=int(match.group("count")),
        code=match.group("code").upper(),
    )


def discover_templates(folder: Path) -> tuple[list[TemplateFile], list[Path]]:
    valid: list[TemplateFile] = []
    ignored: list[Path] = []
    if not folder.exists():
        return valid, ignored
    for path in sorted(folder.rglob("*")):
        if TRASH_DIR_NAME in path.relative_to(folder).parts:
            continue
        if path.is_file() and path.name.lower() != "readme.md":
            template = parse_template(path)
            if template and template.copies > 0:
                valid.append(template)
            elif path.suffix.lower() in SUPPORTED_EXTENSIONS:
                ignored.append(path)
    return valid, ignored


def department_options(templates: list[TemplateFile]) -> list[str]:
    from_file = load_departments_from_file()
    from_templates = {
        template.department.upper()
        for template in templates
        if not template.is_for_every_department
    }
    merged: list[str] = []
    seen: set[str] = set()
    for name in from_file:
        if name not in seen:
            seen.add(name)
            merged.append(name)
    for name in sorted(from_templates):
        if name not in seen:
            seen.add(name)
            merged.append(name)
    return merged


def templates_for_department(
    templates: list[TemplateFile], department: str
) -> list[TemplateFile]:
    chosen = department.strip().upper()
    matching = [
        template
        for template in templates
        if template.is_for_every_department
        or template.department.upper() == chosen
    ]
    return sorted(
        matching,
        key=lambda item: (
            item.is_for_every_department is False,
            item.department.upper(),
            item.path.name.lower(),
        ),
    )


def templates_for_departments(
    templates: list[TemplateFile], departments: list[str]
) -> list[TemplateFile]:
    chosen = {d.strip().upper() for d in departments if d and d.strip()}
    seen_paths: set[Path] = set()
    result: list[TemplateFile] = []
    for template in templates:
        key = template.path
        if key in seen_paths:
            continue
        if template.is_for_every_department or (template.department.upper() in chosen):
            seen_paths.add(key)
            result.append(template)
    return sorted(
        result,
        key=lambda item: (
            item.is_for_every_department is False,
            item.department.upper(),
            item.path.name.lower(),
        ),
    )


# --------------------------- GESTIONE MODULI ------------------------------

TRASH_DIR_NAME = "_moduli_eliminati"
_INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def suggest_template_code(department: str) -> str:
    letters = re.sub(r"[^A-Za-z]", "", department).upper()
    return letters[:3] if len(letters) >= 2 else "GEN"


def build_template_filename(department: str, copies: int, code: str, suffix: str) -> str:
    return f"{department.strip().upper()}_{int(copies)}_{code.strip().upper()}{suffix.lower()}"


def validate_template_fields(department: str, copies: str, code: str) -> str | None:
    """Restituisce la chiave di traduzione dell'errore, oppure None se i campi sono validi."""
    dept = department.strip()
    if not dept:
        return "tm_err_dept"
    if _INVALID_FILENAME_CHARS.search(dept):
        return "tm_err_dept_chars"
    try:
        if int(copies) < 1:
            return "tm_err_copies"
    except (TypeError, ValueError):
        return "tm_err_copies"
    if not re.fullmatch(r"[A-Za-z]{2,5}", code.strip()):
        return "tm_err_code"
    return None


def move_template_to_trash(path: Path, folder: Path) -> Path:
    trash = folder / TRASH_DIR_NAME
    trash.mkdir(parents=True, exist_ok=True)
    target = trash / path.name
    if target.exists():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        target = trash / f"{path.stem}__{stamp}{path.suffix}"
    shutil.move(str(path), str(target))
    return target


def replace_placeholders(value: object, employee_name: str, entry_date: str) -> object:
    if not isinstance(value, str):
        return value
    value = re.sub(r"\*nome\*", employee_name, value, flags=re.IGNORECASE)
    return re.sub(r"\*data\*", entry_date, value, flags=re.IGNORECASE)


def replace_paragraph(paragraph, employee_name: str, entry_date: str) -> None:
    original = paragraph.text
    for run in paragraph.runs:
        run.text = str(replace_placeholders(run.text, employee_name, entry_date))
    if paragraph.text != original:
        return
    replaced = replace_placeholders(original, employee_name, entry_date)
    if original == replaced:
        return
    if paragraph.runs:
        paragraph.runs[0].text = str(replaced)
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(str(replaced))


def replace_docx_placeholders(document: Document, employee_name: str, entry_date: str) -> None:
    for paragraph in document.paragraphs:
        replace_paragraph(paragraph, employee_name, entry_date)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph(paragraph, employee_name, entry_date)
    for section in document.sections:
        for container in (section.header, section.footer):
            for paragraph in container.paragraphs:
                replace_paragraph(paragraph, employee_name, entry_date)
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            replace_paragraph(paragraph, employee_name, entry_date)


def replace_xlsx_placeholders(workbook, employee_name: str, entry_date: str) -> None:
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                try:
                    cell.value = replace_placeholders(cell.value, employee_name, entry_date)
                except AttributeError:
                    pass


@functools.lru_cache(maxsize=1)
def _office_command() -> str | None:
    for command in ("libreoffice", "soffice"):
        if shutil.which(command):
            return command
    if os.name == "nt":
        try:
            import win32com.client  # type: ignore[import-not-found]  # noqa: F401
        except ImportError:
            return None
        return "microsoft-office"
    return None


@contextlib.contextmanager
def _com_apartment():
    """COM va inizializzato in ogni thread che usa Word/Excel (la generazione gira in un worker)."""
    import pythoncom  # type: ignore[import-not-found]

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def _convert_with_ms_office(source: Path, converted: Path, extension: str) -> None:
    import win32com.client  # type: ignore[import-not-found]

    if source.suffix.lower() in {".doc", ".docx"}:
        word = win32com.client.DispatchEx("Word.Application")
        document = word.Documents.Open(str(source.resolve()))
        try:
            if extension == "pdf":
                document.ExportAsFixedFormat(str(converted.resolve()), 17)
            else:
                document.SaveAs2(str(converted.resolve()), FileFormat=16)
        finally:
            document.Close(False)
            word.Quit()
    else:
        excel = win32com.client.DispatchEx("Excel.Application")
        workbook = excel.Workbooks.Open(str(source.resolve()))
        try:
            if extension == "pdf":
                workbook.ExportAsFixedFormat(0, str(converted.resolve()))
            else:
                workbook.SaveAs(str(converted.resolve()), FileFormat=51)
        finally:
            workbook.Close(False)
            excel.Quit()


def _convert_with_office(source: Path, output_dir: Path, extension: str) -> Path:
    command = _office_command()
    if command is None:
        raise RuntimeError(
            "Per mantenere identico il layout dei template Word/Excel serve "
            "LibreOffice installato e disponibile nel PATH."
        )
    output_dir.mkdir(parents=True, exist_ok=True)
    converted = output_dir / f"{source.stem}.{extension}"
    if command == "microsoft-office":
        with _com_apartment():
            _convert_with_ms_office(source, converted, extension)
        if not converted.exists():
            raise RuntimeError(f"Conversione Microsoft Office fallita per {source.name}")
        return converted

    profile = output_dir / "office-profile"
    profile_uri = profile.resolve().as_uri()
    result = subprocess.run(
        [
            command,
            "--headless",
            "--convert-to",
            extension,
            "--outdir",
            str(output_dir),
            f"-env:UserInstallation={profile_uri}",
            str(source),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0 or not converted.exists():
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Conversione Office fallita per {source.name}: {detail}")
    return converted


def _convert_with_ms_office_batch(
    sources: list[Path], output_dir: Path, extension: str, on_each=None
) -> dict[Path, Path]:
    import win32com.client  # type: ignore[import-not-found]

    applications = {}
    converted: dict[Path, Path] = {}
    try:
        for source in sources:
            kind = "word" if source.suffix.lower() == ".docx" else "excel"
            if kind not in applications:
                applications[kind] = (
                    win32com.client.DispatchEx("Word.Application")
                    if kind == "word"
                    else win32com.client.DispatchEx("Excel.Application")
                )
            application = applications[kind]
            target = output_dir / f"{source.stem}.{extension}"
            document = (
                application.Documents.Open(str(source.resolve()))
                if kind == "word"
                else application.Workbooks.Open(str(source.resolve()))
            )
            try:
                if extension == "pdf":
                    if kind == "word":
                        document.ExportAsFixedFormat(str(target.resolve()), 17)
                    else:
                        document.ExportAsFixedFormat(0, str(target.resolve()))
            finally:
                document.Close(False)
            converted[source] = target
            if on_each:
                on_each()
    finally:
        for application in applications.values():
            application.Quit()
    return converted


def _convert_with_office_batch(
    sources: list[Path], output_dir: Path, extension: str, on_each=None
) -> dict[Path, Path]:
    if not sources:
        return {}
    command = _office_command()
    if command is None:
        raise RuntimeError("Motore Office non disponibile per la conversione dei template.")
    output_dir.mkdir(parents=True, exist_ok=True)
    if command == "microsoft-office":
        with _com_apartment():
            return _convert_with_ms_office_batch(sources, output_dir, extension, on_each)

    profile = output_dir / "office-profile"
    result = subprocess.run(
        [
            command,
            "--headless",
            "--convert-to",
            extension,
            "--outdir",
            str(output_dir),
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            *[str(source) for source in sources],
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    converted = {source: output_dir / f"{source.stem}.{extension}" for source in sources}
    if result.returncode != 0 or any(not path.exists() for path in converted.values()):
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Conversione Office fallita: {detail}")
    if on_each:
        for _source in sources:
            on_each()
    return converted


def _prepare_office_template(
    template: Path,
    work_dir: Path,
    employee_name: str,
    entry_date: str,
) -> Path:
    suffix = template.suffix.lower()
    work_dir.mkdir(parents=True, exist_ok=True)
    if suffix in {".doc", ".xls"}:
        modern_extension = "docx" if suffix == ".doc" else "xlsx"
        source = _convert_with_office(template, work_dir, modern_extension)
    else:
        source = work_dir / template.name
        shutil.copy2(template, source)

    if source.suffix.lower() == ".docx":
        document = Document(str(source))
        replace_docx_placeholders(document, employee_name, entry_date)
        document.save(str(source))
    elif source.suffix.lower() == ".xlsx":
        workbook = load_workbook(source, data_only=False, read_only=False)
        try:
            replace_xlsx_placeholders(workbook, employee_name, entry_date)
            workbook.save(source)
        finally:
            workbook.close()
    return source


def _merge_pdfs(output_path: Path, pdf_paths: list[Path]) -> None:
    from pypdf import PdfWriter

    writer = PdfWriter()
    for pdf_path in pdf_paths:
        writer.append(str(pdf_path))
    with output_path.open("wb") as handle:
        writer.write(handle)


def _remove_trailing_blank_pages(pdf_path: Path, dest_dir: Path | None = None) -> Path:
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(str(pdf_path))
    last_page = len(reader.pages)
    while last_page > 1:
        page = reader.pages[last_page - 1]
        text = page.extract_text() or ""
        resources = page.get("/Resources")
        if resources is not None:
            resources = resources.get_object()
        has_images = bool(resources and resources.get("/XObject"))
        content = page.get_contents()
        has_drawing_commands = bool(content and content.get_data().strip())
        if text.strip() or has_images or has_drawing_commands:
            break
        last_page -= 1
    pages = reader.pages[:last_page]
    if len(pages) == len(reader.pages):
        return pdf_path

    target_dir = dest_dir if dest_dir is not None else Path(tempfile.gettempdir())
    trimmed = target_dir / f"{pdf_path.stem}-trimmed-{os.getpid()}-{hashlib.md5(str(pdf_path).encode()).hexdigest()[:8]}.pdf"
    writer = PdfWriter()
    for page in pages:
        writer.add_page(page)
    with trimmed.open("wb") as handle:
        writer.write(handle)
    return trimmed


def _pdf_is_landscape(pdf_path: Path) -> bool:
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    return bool(reader.pages and reader.pages[0].mediabox.width > reader.pages[0].mediabox.height)


@dataclass
class DossierJob:
    output_path: Path
    employee_name: str
    entry_date: str
    department: str
    templates: list[TemplateFile]
    role: str = ""
    notes: str = ""

    def expanded(self) -> list[tuple[TemplateFile, int]]:
        return [(t, n) for t in self.templates for n in range(1, t.copies + 1)]


_PLACEHOLDER_RE = re.compile(r"\*(nome|data)\*", re.IGNORECASE)
_XML_TAG_RE = re.compile(r"<[^>]+>")
_PLACEHOLDER_SCAN: dict[tuple[str, int, int], bool] = {}
PDF_CACHE_LIMIT = 200


def template_has_placeholders(path: Path) -> bool:
    """True se il modulo contiene *nome* o *data*, o se non si puo' stabilirlo
    (.doc/.xls binari): in quel caso va preparato per ogni persona."""
    if path.suffix.lower() not in {".docx", ".xlsx"}:
        return True
    try:
        stat = path.stat()
    except OSError:
        return True
    key = (str(path), stat.st_mtime_ns, stat.st_size)
    if key in _PLACEHOLDER_SCAN:
        return _PLACEHOLDER_SCAN[key]
    import zipfile

    try:
        with zipfile.ZipFile(path) as archive:
            # Senza i tag XML un segnaposto spezzato in piu' "run" torna contiguo
            found = any(
                _PLACEHOLDER_RE.search(_XML_TAG_RE.sub("", archive.read(name).decode("utf-8", "ignore")))
                for name in archive.namelist() if name.endswith(".xml")
            )
    except (OSError, zipfile.BadZipFile):
        found = True
    _PLACEHOLDER_SCAN[key] = found
    return found


def _pdf_cache_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    root = Path(base) / "FormazioniPZZ" if base else Path(tempfile.gettempdir()) / "FormazioniPZZ"
    return root / "pdf-cache"


def _cached_pdf_path(template: Path) -> Path | None:
    digest = compute_template_hash(template)
    if not digest:
        return None
    engine = re.sub(r"[^a-z]", "", (_office_command() or "none").lower())
    return _pdf_cache_dir() / f"{digest}-{engine}.pdf"


def _store_cached_pdf(template: Path, pdf: Path) -> Path:
    target = _cached_pdf_path(template)
    if target is None:
        return pdf
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(pdf, target)
        cached = sorted(target.parent.glob("*.pdf"), key=lambda f: f.stat().st_mtime)
        for old in cached[:-PDF_CACHE_LIMIT]:
            old.unlink(missing_ok=True)
        return target
    except OSError:
        return pdf


def _build_native_jobs(jobs: list[DossierJob], progress_cb=None) -> list[Exception | None]:
    """Prepara i moduli di tutti i dossier e li converte con un'unica sessione di
    Word/Excel (o una sola chiamata a LibreOffice). I moduli senza segnaposto si
    convertono una volta sola e restano in cache tra una generazione e l'altra.
    Ritorna, per ogni dossier, None se riuscito oppure l'errore."""
    results: list[Exception | None] = [None] * len(jobs)
    with tempfile.TemporaryDirectory(prefix="formazioni-pdf-") as temp_name:
        temp_dir = Path(temp_name)
        names = iter(range(1, 1_000_000))
        ready: dict[Path, Path] = {}            # moduli fissi gia' convertiti (cache)
        static_sources: dict[Path, Path] = {}   # moduli fissi da convertire una volta
        personal: dict[tuple[int, Path], Path] = {}
        office_paths = [
            list(dict.fromkeys(t.path for t in job.templates if t.path.suffix.lower() != ".pdf"))
            for job in jobs
        ]
        total = max(1, sum(len(paths) for paths in office_paths) * 2 + len(jobs))
        step = 0

        def tick():
            nonlocal step
            step = min(step + 1, total)
            if progress_cb:
                progress_cb(step, total)

        for index, job in enumerate(jobs):
            try:
                for template_path in office_paths[index]:
                    if not template_has_placeholders(template_path):
                        if template_path not in ready and template_path not in static_sources:
                            cached = _cached_pdf_path(template_path)
                            if cached is not None and cached.exists():
                                ready[template_path] = cached
                            else:
                                copy = temp_dir / f"{next(names)}-{template_path.name}"
                                shutil.copy2(template_path, copy)
                                static_sources[template_path] = copy
                    else:
                        prepared = _prepare_office_template(
                            template_path, temp_dir / f"job-{index}-{next(names)}",
                            job.employee_name, job.entry_date)
                        # Nome univoco: i convertitori chiamano il PDF come il sorgente
                        unique = prepared.with_name(f"{next(names)}-{prepared.name}")
                        prepared.rename(unique)
                        personal[(index, template_path)] = unique
                    tick()
            except Exception as exc:  # noqa: BLE001 - l'errore resta legato a quel dossier
                results[index] = exc

        sources = list(static_sources.values()) + [
            source for (index, _), source in personal.items() if results[index] is None
        ]
        converted = _convert_with_office_batch(
            sources, temp_dir / "converted", "pdf", on_each=tick) if sources else {}
        for template_path, source in static_sources.items():
            trimmed = _remove_trailing_blank_pages(converted[source], temp_dir)
            ready[template_path] = _store_cached_pdf(template_path, trimmed)

        trimmed_cache: dict[Path, Path] = {}

        def trimmed(pdf: Path) -> Path:
            if pdf not in trimmed_cache:
                trimmed_cache[pdf] = _remove_trailing_blank_pages(pdf, temp_dir)
            return trimmed_cache[pdf]

        for index, job in enumerate(jobs):
            if results[index] is not None:
                tick()
                continue
            try:
                expanded = job.expanded()
                # Stessa regola del motore ReportLab: niente copertina se ci sono PDF orizzontali
                has_landscape_pdf = any(
                    template.path.suffix.lower() == ".pdf" and _pdf_is_landscape(template.path)
                    for template, _ in expanded
                )
                cover = [] if has_landscape_pdf else [_build_cover_pdf(
                    temp_dir / f"cover-{index}.pdf", job.employee_name, job.entry_date,
                    job.department, job.role, job.notes,
                )]
                pages = cover + [
                    trimmed(template.path) if template.path.suffix.lower() == ".pdf"
                    else ready.get(template.path)
                    or trimmed(converted[personal[(index, template.path)]])
                    for template, _copy_number in expanded
                ]
                job.output_path.parent.mkdir(parents=True, exist_ok=True)
                _merge_pdfs(job.output_path, pages)
            except Exception as exc:  # noqa: BLE001
                results[index] = exc
            tick()
    return results


def paragraph_text(text: str, style: ParagraphStyle) -> Paragraph:
    from reportlab.platypus import Paragraph

    safe = escape(text).replace("\n", "<br/>")
    return Paragraph(safe or " ", style)


TABLE_WIDTH = 174 * mm


def _table_widths(values: list[list[str]], available_width: float = TABLE_WIDTH) -> list[float]:
    column_count = max((len(row) for row in values), default=0)
    if not column_count:
        return []
    lengths = []
    for column in range(column_count):
        longest = max(
            (len(str(row[column]).replace("\n", " ").strip()) for row in values if column < len(row)),
            default=4,
        )
        lengths.append(max(4, min(longest, 36)))
    minimum = available_width / column_count
    weights = [max(length, 8) for length in lengths]
    total = sum(weights)
    widths = [available_width * weight / total for weight in weights]
    if any(width < minimum for width in widths):
        widths = [minimum] * column_count
    return widths


def _table_flowable(rows: list[list[object]], raw_values: list[list[str]]) -> Table:
    from reportlab.platypus import Table, TableStyle

    column_count = max((len(row) for row in rows), default=0)
    rows = [row + [""] * (column_count - len(row)) for row in rows]
    raw_values = [row + [""] * (column_count - len(row)) for row in raw_values]
    return Table(
        rows,
        colWidths=_table_widths(raw_values),
        repeatRows=1,
        splitByRow=1,
        splitInRow=0,
        hAlign="LEFT",
        style=TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e7eef1")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#173642")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#b9c9cf")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        ),
    )


def docx_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
) -> list[object]:
    from reportlab.platypus import Spacer

    document = Document(path)
    replace_docx_placeholders(document, employee_name, entry_date)
    story: list[object] = []
    style_name_cache: dict[int, str] = {}
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        heading = ""
        try:
            if paragraph.style is not None:
                pid = id(paragraph.style)
                if pid not in style_name_cache:
                    style_name_cache[pid] = (paragraph.style.name or "").lower()
                heading = style_name_cache[pid]
        except Exception:
            heading = ""
        style = styles["subheading"] if "heading" in heading else styles["body"]
        story.append(paragraph_text(text, style))
        story.append(Spacer(1, 2.4 * mm))
    for table in document.tables:
        rows: list[list[object]] = []
        raw_rows: list[list[str]] = []
        for row in table.rows:
            values = [cell.text.strip() for cell in row.cells]
            if any(values):
                while values and not values[-1]:
                    values.pop()
                raw_rows.append(values)
                rows.append([paragraph_text(value, styles["table"]) for value in values])
        if rows:
            story.append(Spacer(1, 2 * mm))
            story.append(_table_flowable(rows, raw_rows))
            story.append(Spacer(1, 4 * mm))
    return story or [paragraph_text("Documento senza contenuto testuale.", styles["muted"])]


def xlsx_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
) -> list[object]:
    from reportlab.platypus import Paragraph, Spacer

    workbook = load_workbook(path, data_only=False, read_only=False)
    story: list[object] = []
    try:
        for sheet in workbook.worksheets:
            rows: list[list[object]] = []
            raw_rows: list[list[str]] = []
            for row in sheet.iter_rows(values_only=True):
                values = [
                    replace_placeholders(value, employee_name, entry_date)
                    for value in row
                ]
                if any(value not in (None, "") for value in values):
                    values = [str(value) if value is not None else "" for value in values]
                    while values and not values[-1]:
                        values.pop()
                    raw_rows.append(values)
                    rows.append(
                        [paragraph_text(value, styles["table"]) for value in values]
                    )
            if rows:
                story.append(Paragraph(escape(sheet.title), styles["subheading"]))
                story.append(Spacer(1, 2 * mm))
                story.append(_table_flowable(rows, raw_rows))
                story.append(Spacer(1, 5 * mm))
    finally:
        workbook.close()
    return story or [paragraph_text("Foglio senza contenuto.", styles["muted"])]


def make_styles() -> dict[str, ParagraphStyle]:
    global _STYLE_CACHE
    if _STYLE_CACHE is not None:
        return _STYLE_CACHE  # type: ignore[return-value]
    base = getSampleStyleSheet()
    result = {
        "cover_title": ParagraphStyle("CoverTitle", parent=base["Title"], fontName="Helvetica-Bold",
                                     fontSize=25, leading=30, textColor=colors.HexColor("#173642"),
                                     alignment=TA_CENTER, spaceAfter=8 * mm),
        "cover_subtitle": ParagraphStyle("CoverSubtitle", parent=base["Normal"], fontName="Helvetica",
                                         fontSize=12, leading=17, textColor=colors.HexColor("#48636d"),
                                         alignment=TA_CENTER),
        "meta": ParagraphStyle("Meta", parent=base["Normal"], fontName="Helvetica",
                               fontSize=10.5, leading=16, textColor=colors.HexColor("#173642")),
        "heading": ParagraphStyle("Heading", parent=base["Heading1"], fontName="Helvetica-Bold",
                                  fontSize=15, leading=19, textColor=colors.HexColor("#173642"),
                                  spaceAfter=4 * mm),
        "subheading": ParagraphStyle("Subheading", parent=base["Heading2"], fontName="Helvetica-Bold",
                                     fontSize=11.5, leading=15, textColor=colors.HexColor("#247b7b"),
                                     spaceBefore=2 * mm, spaceAfter=2 * mm),
        "body": ParagraphStyle("Body", parent=base["BodyText"], fontName="Helvetica",
                               fontSize=10, leading=14, textColor=colors.HexColor("#263f47")),
        "table": ParagraphStyle("Table", parent=base["BodyText"], fontName="Helvetica",
                                fontSize=8.2, leading=10, textColor=colors.HexColor("#263f47")),
        "muted": ParagraphStyle("Muted", parent=base["BodyText"], fontName="Helvetica-Oblique",
                                fontSize=9.5, leading=13, textColor=colors.HexColor("#71838a")),
        "small": ParagraphStyle("Small", parent=base["BodyText"], fontName="Helvetica",
                                fontSize=8.5, leading=11, textColor=colors.HexColor("#71838a")),
        "pdf_page": ParagraphStyle("PdfPage", parent=base["BodyText"], fontName="Helvetica",
                                   fontSize=8.5, leading=11.5, textColor=colors.HexColor("#263f47")),
    }
    _STYLE_CACHE = result  # type: ignore[assignment]
    return result


def pdf_story(
    path: Path,
    employee_name: str,
    entry_date: str,
    styles: dict[str, ParagraphStyle],
    frame_width: float = 174 * mm,
    frame_height: float = 263 * mm,
) -> list[object]:
    from reportlab.platypus import KeepInFrame, PageBreak, Paragraph, Spacer

    story: list[object] = []
    pdf_style = styles["pdf_page"]
    is_landscape = frame_width > frame_height
    if is_landscape:
        pdf_style = ParagraphStyle("PdfPageLandscape", parent=styles["pdf_page"],
                                   fontSize=7.5, leading=10)
    try:
        from pypdf import PdfReader
    except Exception:
        story.append(paragraph_text(
            "Libreria pypdf non disponibile per leggere il template PDF.", styles["muted"]))
        return story
    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        story.append(paragraph_text(f"Impossibile aprire il PDF: {exc}", styles["muted"]))
        return story
    if not reader.pages:
        story.append(paragraph_text("PDF senza pagine.", styles["muted"]))
        return story
    for idx, page in enumerate(reader.pages, start=1):
        page_story: list[object] = []
        if len(reader.pages) > 1:
            page_story.append(Paragraph(f"Pagina {idx} del PDF", styles["subheading"]))
        raw = ""
        try:
            raw = page.extract_text() or ""
        except Exception:
            raw = ""
        raw = replace_placeholders(raw, employee_name, entry_date)
        lines = [ln.rstrip() for ln in raw.splitlines()]
        if not any(ln.strip() for ln in lines):
            page_story.append(paragraph_text(
                "(Nessun testo estraibile dal PDF — file incluso come riferimento.)",
                styles["muted"]))
            page_story.append(Spacer(1, 3 * mm))
        else:
            for line in lines:
                if not line.strip():
                    page_story.append(Spacer(1, 1.5 * mm))
                    continue
                page_story.append(Paragraph(escape(line), pdf_style))
            page_story.append(Spacer(1, 4 * mm))
        story.append(KeepInFrame(frame_width, frame_height, page_story, mode="shrink"))
        if idx < len(reader.pages):
            story.append(PageBreak())
    return story or [paragraph_text("Documento PDF senza contenuto testuale.", styles["muted"])]


def _cover_story(
    styles: dict[str, ParagraphStyle],
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
) -> list[object]:
    from reportlab.platypus import Spacer

    cover_rows_raw = [
        ["Campo", "Valore"],
        ["Nome e Cognome", employee_name],
        ["Data Ingresso / Corso", entry_date],
        ["Reparto/i", department],
    ]
    if role and role.strip():
        cover_rows_raw.append(["Mansione / Ruolo", role.strip()])
    if notes and notes.strip():
        cover_rows_raw.append(["Note aggiuntive", notes.strip()])
    cover_rows = [
        [paragraph_text(str(c), styles["body"] if i else styles["meta"])
         for i, c in enumerate(rr)]
        for rr in cover_rows_raw
    ]
    return [
        Spacer(1, 12 * mm),
        paragraph_text("Dossier Formazione", styles["cover_title"]),
        paragraph_text("Documento di accompagnamento per la formazione individuale",
                       styles["cover_subtitle"]),
        Spacer(1, 8 * mm),
        _table_flowable(cover_rows, cover_rows_raw),
        Spacer(1, 8 * mm),
        paragraph_text(
            "Generato da Formazioni PZZ il: " + datetime.now().strftime("%d/%m/%Y %H:%M"),
            styles["small"],
        ),
    ]


def _build_cover_pdf(
    output_path: Path,
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
) -> Path:
    from reportlab.platypus import SimpleDocTemplate

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=14 * mm, leftMargin=14 * mm, topMargin=14 * mm, bottomMargin=14 * mm,
        title=f"Dossier formazione - {employee_name}",
        author="Formazioni PZZ",
    )
    doc.build(_cover_story(make_styles(), employee_name, entry_date, department, role, notes))
    return output_path


def _watermark_page(text: str, width: float, height: float):
    """Pagina PDF trasparente con la filigrana in diagonale, centrata."""
    import io
    import math

    from pypdf import PdfReader
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas as rl_canvas

    font = "Helvetica-Bold"
    diagonal = math.hypot(width, height)
    size = min(120.0, 0.7 * diagonal / max(stringWidth(text, font, 1), 1))
    buffer = io.BytesIO()
    canvas = rl_canvas.Canvas(buffer, pagesize=(width, height))
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#7f939b"))
    canvas.setFillAlpha(0.16)
    canvas.translate(width / 2, height / 2)
    canvas.rotate(math.degrees(math.atan2(height, width)))
    canvas.setFont(font, size)
    canvas.drawCentredString(0, -size / 3, text)
    canvas.restoreState()
    canvas.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]


def finalize_pdf(
    path: Path,
    title: str,
    subject: str,
    watermark: str = "",
    protect: bool = False,
) -> None:
    """Metadati, filigrana opzionale e blocco modifiche opzionale sul PDF finale."""
    from pypdf import PdfReader, PdfWriter, Transformation
    from pypdf.constants import UserAccessPermissions as Perm

    writer = PdfWriter(clone_from=PdfReader(str(path)))
    watermark = watermark.strip()
    if watermark:
        overlays: dict[tuple[int, int], Any] = {}
        for page in writer.pages:
            box = page.mediabox
            key = (round(float(box.width)), round(float(box.height)))
            if key not in overlays:
                overlays[key] = _watermark_page(watermark, float(box.width), float(box.height))
            page.merge_transformed_page(
                overlays[key], Transformation().translate(float(box.left), float(box.bottom))
            )
    writer.add_metadata({
        "/Title": title,
        "/Author": "Formazioni PZZ",
        "/Subject": subject,
        "/Creator": f"Formazioni PZZ {APP_VERSION}",
        "/Producer": f"Formazioni PZZ {APP_VERSION}",
    })
    if protect:
        everything = functools.reduce(lambda a, b: a | b, Perm)
        blocked = Perm.MODIFY | Perm.ADD_OR_MODIFY | Perm.ASSEMBLE_DOC | Perm.FILL_FORM_FIELDS
        # Si apre senza password; la password proprietario casuale impedisce di
        # togliere il blocco. RC4 non richiede librerie crittografiche esterne.
        writer.encrypt(user_password="", owner_password=secrets.token_hex(16),
                       permissions_flag=everything & ~blocked, algorithm="RC4-128")
    temp_path = path.with_name(path.stem + ".finalizing.pdf")
    with temp_path.open("wb") as handle:
        writer.write(handle)
    os.replace(temp_path, path)


def build_pdf(
    output_path: Path,
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
    templates: list[TemplateFile],
    progress_cb=None,
    *,
    watermark: str = "",
    protect: bool = False,
) -> int:
    total = _build_pdf_content(output_path, employee_name, entry_date, department,
                               role, notes, templates, progress_cb)
    _finalize_dossier(output_path, employee_name, entry_date, department, watermark, protect)
    return total


def _finalize_dossier(output_path: Path, employee_name: str, entry_date: str,
                      department: str, watermark: str, protect: bool) -> None:
    try:
        finalize_pdf(output_path, f"Dossier formazione - {employee_name}",
                     f"Reparto: {department} - Ingresso: {entry_date}",
                     watermark=watermark, protect=protect)
    except Exception:
        # Filigrana e protezione richieste esplicitamente: l'errore va mostrato.
        # I soli metadati sono accessori e non devono far fallire il dossier.
        if watermark.strip() or protect:
            raise
        traceback.print_exc()


def _build_job(job: DossierJob, watermark: str, protect: bool) -> None:
    """Un dossier completo; funzione di modulo per poterla eseguire in un altro processo."""
    build_pdf(job.output_path, job.employee_name, job.entry_date, job.department,
              job.role, job.notes, job.templates, watermark=watermark, protect=protect)


def build_pdfs(
    jobs: list[DossierJob],
    progress_cb=None,
    *,
    watermark: str = "",
    protect: bool = False,
) -> list[Exception | None]:
    """Genera piu' dossier. Con Office: una sola sessione per tutto il lotto.
    Senza Office (motore ReportLab): piu' dossier in parallelo su processi separati.
    Ritorna, per ogni dossier, None se riuscito oppure l'errore."""
    results: list[Exception | None] = [None] * len(jobs)
    pending = list(range(len(jobs)))
    if not jobs:
        return results
    if _office_command():
        try:
            native = _build_native_jobs(jobs, progress_cb)
            for index, error in enumerate(native):
                if error is None:
                    job = jobs[index]
                    try:
                        _finalize_dossier(job.output_path, job.employee_name, job.entry_date,
                                          job.department, watermark, protect)
                    except Exception as exc:  # noqa: BLE001
                        results[index] = exc
            pending = [index for index, error in enumerate(native) if error is not None]
        except Exception:  # noqa: BLE001 - conversione di gruppo fallita: uno alla volta
            traceback.print_exc()
        # I falliti si riprovano singolarmente (con il ripiego su ReportLab)
        for done, index in enumerate(pending, start=1):
            try:
                _build_job(jobs[index], watermark, protect)
            except Exception as exc:  # noqa: BLE001
                results[index] = exc
            if progress_cb:
                progress_cb(done, len(pending))
        return results

    finished: set[int] = set()

    def record(index: int, error: Exception | None) -> None:
        results[index] = error
        finished.add(index)
        if progress_cb:
            progress_cb(len(finished), len(jobs))

    workers = min(len(jobs), 4, max(1, (os.cpu_count() or 2) - 1))
    if workers > 1 and len(jobs) >= 3:
        from concurrent.futures import ProcessPoolExecutor, as_completed

        try:
            with ProcessPoolExecutor(max_workers=workers) as pool:
                futures = {pool.submit(_build_job, jobs[i], watermark, protect): i for i in pending}
                for future in as_completed(futures):
                    try:
                        future.result()
                        record(futures[future], None)
                    except Exception as exc:  # noqa: BLE001
                        record(futures[future], exc)
        except Exception:  # noqa: BLE001 - processi non disponibili: si prosegue in sequenza
            traceback.print_exc()
    for index in pending:
        if index in finished:
            continue
        try:
            _build_job(jobs[index], watermark, protect)
            record(index, None)
        except Exception as exc:  # noqa: BLE001
            record(index, exc)
    return results


def _build_pdf_content(
    output_path: Path,
    employee_name: str,
    entry_date: str,
    department: str,
    role: str,
    notes: str,
    templates: list[TemplateFile],
    progress_cb=None,
) -> int:
    from reportlab.platypus import PageBreak, SimpleDocTemplate, Table

    output_path.parent.mkdir(parents=True, exist_ok=True)
    styles = make_styles()
    expanded = [
        (template, copy_number)
        for template in templates
        for copy_number in range(1, template.copies + 1)
    ]
    needs_legacy_office = any(
        template.path.suffix.lower() in {".doc", ".xls"}
        for template, _ in expanded
    )
    if _office_command():
        # Il motore Office conserva il layout originale dei template (loghi, tabelle, immagini)
        try:
            error = _build_native_jobs([DossierJob(
                output_path, employee_name, entry_date, department, templates, role, notes,
            )], progress_cb)[0]
            if error is None:
                return len(expanded)
            raise error
        except Exception:
            if needs_legacy_office:
                raise
            # .doc/.xls richiedono Office; per gli altri formati si ripiega su ReportLab
            traceback.print_exc()
    has_landscape_pdf = any(
        template.path.suffix.lower() == ".pdf" and _pdf_is_landscape(template.path)
        for template, _ in expanded
    )
    page_size = landscape(A4) if has_landscape_pdf else A4
    page_width, page_height = page_size
    horizontal_frame = min(page_width - 28 * mm, 210 * mm)
    vertical_frame = min(page_height - 28 * mm, 139 * mm)
    h_margin = 10 * mm if has_landscape_pdf else 14 * mm
    v_margin = 10 * mm if has_landscape_pdf else 14 * mm

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=page_size,
        rightMargin=h_margin, leftMargin=h_margin, topMargin=v_margin, bottomMargin=v_margin,
        title=f"Dossier formazione - {employee_name}",
        author="Formazioni PZZ",
    )
    story: list[object] = []
    # Copertina solo in portrait: con PDF landscape si conserva l'impaginazione originale
    if not has_landscape_pdf:
        story.extend(_cover_story(styles, employee_name, entry_date, department, role, notes))
        story.append(PageBreak())
    story_cache_local: dict[Path, list[object]] = {}
    total_steps = max(1, len(expanded) + 1)

    def _fetch_story(template: TemplateFile, step_idx: int) -> list[object]:
        if progress_cb:
            progress_cb(step_idx, total_steps)
        if template.path in story_cache_local:
            return story_cache_local[template.path]
        suffix = template.path.suffix.lower()
        try:
            mtime = template.path.stat().st_mtime_ns
        except OSError:
            mtime = 0
        key = (
            hash(template.path.resolve().as_posix()),
            mtime, employee_name, entry_date, suffix,
            int(horizontal_frame), int(vertical_frame),
        )
        cached_global = _TEMPLATE_STORY_CACHE.get(key)
        if cached_global is not None:
            story_cache_local[template.path] = cached_global
            return cached_global
        if suffix == ".docx":
            generated = docx_story(template.path, employee_name, entry_date, styles)
        elif suffix == ".xlsx":
            generated = xlsx_story(template.path, employee_name, entry_date, styles)
        elif suffix == ".pdf":
            generated = pdf_story(template.path, employee_name, entry_date, styles,
                                  horizontal_frame, vertical_frame)
        else:
            generated = []
        story_cache_local[template.path] = generated
        _TEMPLATE_STORY_CACHE[key] = generated
        return generated

    for index, (template, copy_number) in enumerate(expanded):
        if index > 0:
            story.append(PageBreak())
        cached = _fetch_story(template, index + 1)
        if len(cached) > 3 or isinstance(cached[0] if cached else None, Table):
            story.extend(_copy_mod.copy(cached))
        else:
            story.extend(cached)
    doc.build(story)
    if progress_cb:
        progress_cb(total_steps, total_steps)
    return len(expanded)


def safe_file_part(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9À-ÿ_-]+", "-", value.strip())
    return cleaned.strip("-_") or "persona"


def open_folder(path: Path) -> None:
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except OSError:
        pass


def write_batch_template(path: Path, departments: list[str], italian: bool = True) -> None:
    """Crea un file Excel pronto da compilare per la generazione multipla."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    headers = ["Nome", "Data", "Reparto"] if italian else ["Name", "Date", "Department"]
    rows_with_rules = 500
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Dipendenti" if italian else "Employees"
    sheet.append(headers)
    example_dept = next((d for d in departments if d.upper() not in ALL_DEPARTMENT_NAMES),
                        departments[0] if departments else "TUTTI")
    sheet.append(["Mario Rossi", date.today(), example_dept])
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0B2A3D")
        cell.alignment = Alignment(vertical="center")
    for row in range(2, rows_with_rules + 2):
        sheet.cell(row=row, column=2).number_format = "DD/MM/YYYY"
    for column, width in zip("ABC", (34, 16, 28)):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"

    date_rule = DataValidation(type="date", operator="greaterThan", formula1="DATE(1990,1,1)",
                               allow_blank=True, errorStyle="warning")
    date_rule.error = "Data non valida (gg/mm/aaaa)" if italian else "Invalid date (dd/mm/yyyy)"
    sheet.add_data_validation(date_rule)
    date_rule.add(f"B2:B{rows_with_rules + 1}")
    if departments:
        # Elenco reparti su un foglio nascosto: niente limite di 255 caratteri.
        lists = workbook.create_sheet("Reparti" if italian else "Departments")
        for index, dept in enumerate(departments, start=1):
            lists.cell(row=index, column=1, value=dept)
        lists.sheet_state = "hidden"
        dept_rule = DataValidation(type="list", allow_blank=True, errorStyle="warning",
                                   formula1=f"='{lists.title}'!$A$1:$A${len(departments)}")
        dept_rule.error = ("Reparto non in elenco. Per più reparti usa A+B."
                           if italian else "Department not listed. For several use A+B.")
        sheet.add_data_validation(dept_rule)
        dept_rule.add(f"C2:C{rows_with_rules + 1}")
    workbook.save(path)


def parse_version(value: str) -> tuple[int, ...]:
    parts = re.findall(r"\d+", str(value))
    return tuple(int(p) for p in parts[:4]) or (0,)


def fetch_update_manifest(source: str, timeout: float = 4.0) -> dict[str, str] | None:
    """Legge il manifest {version, url, notes} da un URL http(s) o da un percorso
    locale / di rete (file version.json o cartella che lo contiene)."""
    source = source.strip()
    if not source:
        return None
    if source.lower().startswith(("http://", "https://")):
        request = urllib.request.Request(
            source, headers={"User-Agent": f"FormazioniPZZ/{APP_VERSION}"})
        with urllib.request.urlopen(request, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8-sig"))
        default_url = ""
        if isinstance(data, dict) and "tag_name" in data:
            # API GitHub "releases/latest": tag vX.Y.Z, link diretto al setup se allegato
            setup = next((a.get("browser_download_url") for a in data.get("assets") or []
                          if str(a.get("name", "")).lower() == "formazionipzz_setup.exe"), None)
            data = {"version": str(data["tag_name"]).lstrip("vV"),
                    "url": setup or data.get("html_url") or "",
                    "notes": data.get("name") or ""}
    else:
        manifest = Path(source).expanduser()
        if manifest.is_dir():
            manifest = manifest / "version.json"
        data = json.loads(manifest.read_text(encoding="utf-8-sig"))
        default_url = str(manifest.parent / "FormazioniPZZ_Setup.exe")
    if not isinstance(data, dict) or not data.get("version"):
        raise ValueError("Manifest aggiornamenti non valido: manca 'version'.")
    return {
        "version": str(data["version"]),
        "url": str(data.get("url") or default_url),
        "notes": str(data.get("notes") or ""),
    }


def cleanup_old_previews() -> None:
    for old in Path(tempfile.gettempdir()).glob(f"{PREVIEW_PREFIX}*.pdf"):
        try:
            old.unlink()
        except OSError:
            pass  # ancora aperta nel visualizzatore: sara' rimossa al prossimo avvio


# --------------------------- UI HELPERS -----------------------------------

class Tooltip:
    """Tooltip minimale con delay e auto-dismiss."""

    def __init__(self, widget: tk.Misc, text_getter, delay_ms: int = 600, auto_dismiss_ms: int = 8000):
        self.widget = widget
        self.text_getter = text_getter if callable(text_getter) else lambda: text_getter
        self.delay_ms = delay_ms
        self.auto_dismiss_ms = auto_dismiss_ms
        self._after_id: str | None = None
        self._dismiss_after_id: str | None = None
        self.tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")
        widget.bind("<Motion>", self._on_leave, add="+")
        widget.bind("<ButtonPress>", self._on_leave, add="+")
        widget.bind("<Destroy>", self._on_leave, add="+")

    def _on_enter(self, _evt=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _on_leave(self, _evt=None):
        self._cancel()
        self._hide()

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        if self._dismiss_after_id is not None:
            try:
                self.widget.after_cancel(self._dismiss_after_id)
            except Exception:
                pass
            self._dismiss_after_id = None

    def _show(self):
        self._after_id = None
        text = str(self.text_getter() or "").strip()
        if not text:
            return
        self.tip = t = tk.Toplevel(self.widget)
        t.wm_overrideredirect(True)
        try:
            t.wm_attributes("-topmost", True)
        except Exception:
            pass
        try:
            t.configure(bg="#0b2a3d")
        except Exception:
            pass
        frame = tk.Frame(t, bg="#0b2a3d", highlightthickness=1,
                         highlightbackground="#1f8a87", padx=11, pady=7)
        frame.pack(fill=BOTH, expand=True)
        label = tk.Label(frame, text=text, justify="left",
                         bg="#0b2a3d", fg="#f1f6f8",
                         font=("Segoe UI", 9), wraplength=360)
        label.pack()
        self.widget.update_idletasks()
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            t.wm_geometry(f"+{x}+{y}")
        except Exception:
            pass
        if self.auto_dismiss_ms > 0:
            self._dismiss_after_id = self.widget.after(self.auto_dismiss_ms, self._hide)

    def _hide(self):
        if self.tip is not None:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


class DatePickerFrame(tk.Frame):
    """Date picker puro Tk: 3 Combobox (giorno/mese/anno) con validazione.

    Nota: intercetta <MouseWheel> sul DatePicker e lo traduce in avanti/indietro
    sul Combobox *sotto il mouse* (giorno/mese/anno) SENZA propagare l'evento
    al resto dell'app — evita che la rotella scrolli insieme tutte le scrollbar.
    """

    def __init__(self, master, language: dict[str, str], initial: date | None = None, **kwargs):
        super().__init__(master, **kwargs)
        self.language = language
        self._build()
        if initial is None:
            initial = date.today()
        self.set_date(initial)
        self._bind_change()
        self._bind_wheel_block()

    def _build(self):
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=0)
        self.columnconfigure(2, weight=0)
        today = date.today()
        end_year = max(today.year + 6, 2051)
        year_values = [str(y) for y in range(today.year - 5, end_year)]
        self.day_cb = ttk.Combobox(self, values=[str(d) for d in range(1, 32)],
                                   width=4, state="readonly")
        self.month_cb = ttk.Combobox(self, values=[self.language.get(k, k) for k in MONTH_KEYS],
                                     width=12, state="readonly")
        self.year_cb = ttk.Combobox(self, values=year_values, width=6, state="readonly")
        self.day_cb.grid(row=0, column=0, padx=(0, 6), sticky="w")
        self.month_cb.grid(row=0, column=1, padx=(0, 6), sticky="w")
        self.year_cb.grid(row=0, column=2, sticky="w")

    def _bind_wheel_block(self):
        def step_cb(cb, direction, _evt):
            vals = list(cb["values"])
            if not vals:
                return "break"
            current = cb.current()
            if current < 0:
                current = 0
            new = current + direction
            if new < 0:
                new = 0
            elif new >= len(vals):
                new = len(vals) - 1
            if new != current:
                cb.current(new)
                cb.event_generate("<<ComboboxSelected>>")
            return "break"

        def on_window(evt):
            # Target the Combobox under the mouse if any; otherwise first focused
            target = None
            x_root, y_root = evt.x_root, evt.y_root
            for cb in (self.day_cb, self.month_cb, self.year_cb):
                try:
                    x0 = cb.winfo_rootx()
                    y0 = cb.winfo_rooty()
                    x1 = x0 + cb.winfo_width()
                    y1 = y0 + cb.winfo_height()
                    if x0 <= x_root <= x1 and y0 <= y_root <= y1:
                        target = cb
                        break
                except Exception:
                    continue
            if target is None:
                try:
                    focus_w = self.focus_get()
                    if focus_w in (self.day_cb, self.month_cb, self.year_cb):
                        target = focus_w
                except Exception:
                    target = None
            # direction: delta > 0 means scroll up / forward => previous value
            direction = 1 if evt.delta < 0 else -1
            if target is not None:
                return step_cb(target, direction, evt)
            return "break"

        for w in (self, self.day_cb, self.month_cb, self.year_cb):
            w.bind("<MouseWheel>", on_window)
            # Unix: helper inline per ritornare "break" correttamente (non tupla)
            def _unix_up(e, _d=-1):
                cb = self._cb_under_mouse(e) or self.day_cb
                step_cb(cb, _d, e)
                return "break"
            def _unix_down(e, _d=+1):
                cb = self._cb_under_mouse(e) or self.day_cb
                step_cb(cb, _d, e)
                return "break"
            w.bind("<Button-4>", _unix_up)
            w.bind("<Button-5>", _unix_down)

    def _cb_under_mouse(self, evt):
        try:
            x_root = evt.x_root if hasattr(evt, "x_root") else self.winfo_rootx() + evt.x
            y_root = evt.y_root if hasattr(evt, "y_root") else self.winfo_rooty() + evt.y
        except Exception:
            return None
        for cb in (self.day_cb, self.month_cb, self.year_cb):
            try:
                x0 = cb.winfo_rootx()
                y0 = cb.winfo_rooty()
                x1 = x0 + cb.winfo_width()
                y1 = y0 + cb.winfo_height()
                if x0 <= x_root <= x1 and y0 <= y_root <= y1:
                    return cb
            except Exception:
                continue
        return None

    def _bind_change(self):
        def sync(_e=None):
            days_max = self._compute_days()
            current = self.day_cb.current() + 1
            day_vals = [str(d) for d in range(1, days_max + 1)]
            self.day_cb.configure(values=day_vals)
            if current > days_max:
                self.day_cb.current(days_max - 1)
            self._on_change()

        self.day_cb.bind("<<ComboboxSelected>>", sync)
        self.month_cb.bind("<<ComboboxSelected>>", sync)
        self.year_cb.bind("<<ComboboxSelected>>", sync)

    def _compute_days(self) -> int:
        m = self.month_cb.current() + 1
        try:
            y = int(self.year_cb.get()) if self.year_cb.get() else date.today().year
        except ValueError:
            y = date.today().year
        if m < 1 or m > 12:
            return 31
        return calendar.monthrange(y, m)[1]

    def _on_change(self):
        pass

    def get_date(self) -> date | None:
        try:
            d = int(self.day_cb.get())
            m = self.month_cb.current() + 1
            y = int(self.year_cb.get())
            if not (1 <= m <= 12):
                return None
            dm = calendar.monthrange(y, m)[1]
            if not (1 <= d <= dm):
                d = dm
            return date(y, m, d)
        except (ValueError, tk.TclError):
            return None

    def get_string(self) -> str:
        d = self.get_date()
        return d.strftime("%d/%m/%Y") if d else ""

    def set_date(self, d: date) -> None:
        try:
            self.day_cb.current(d.day - 1)
        except Exception:
            pass
        try:
            self.month_cb.current(d.month - 1)
        except Exception:
            pass
        vals = list(self.year_cb["values"])
        target = str(d.year)
        if target in vals:
            self.year_cb.current(vals.index(target))
        else:
            self.year_cb.set(target)

    def configure_language(self, language: dict[str, str]) -> None:
        self.language = language
        current_month_index = self.month_cb.current()
        self.month_cb.configure(values=[language.get(k, k) for k in MONTH_KEYS])
        if 0 <= current_month_index < 12:
            self.month_cb.current(current_month_index)


# --------------------------- APP ------------------------------------------

class FormazioniApp:
    def __init__(self, root: tk.Tk) -> None:
        if tk is None or ttk is None:
            raise RuntimeError(
                "Questa installazione di Python non include Tkinter. "
                "Installa Python con il supporto Tk e riavvia Formazioni PZZ."
            )
        self.root = root

        self.settings = load_settings()
        self.language_code = self.settings.get("language", DEFAULT_LANG)
        self.language = load_language(self.language_code)
        self.saved_hashes = load_saved_hashes()
        self.hash_status: dict[str, str] = {}

        self.template_dir = StringVar(value=self.settings.get("last_template_dir", str(DEFAULT_TEMPLATE_DIR)))
        self.output_dir = StringVar(value=self.settings.get("last_output_dir", str(DEFAULT_OUTPUT_DIR)))
        self.employee_name = StringVar()
        self.department = StringVar()
        self.auto_open = BooleanVar(value=True)
        self.multi_dept_mode = BooleanVar(value=False)
        # theme_pref e' la scelta dell'utente (light/dark/system); self.theme il tema
        # effettivamente applicato, sempre "light" o "dark".
        pref = self.settings.get("theme", DEFAULT_THEME)
        self.theme_pref = pref if pref in ("light", "dark", "system") else DEFAULT_THEME
        self.theme = StringVar(value=resolve_theme(self.theme_pref))
        self.language_var = StringVar(value=self.language_code)
        self.status = StringVar(value=self.tr("status_initial"))
        self.count_label = StringVar(value=self.tr("count_none"))
        self.progress_label = StringVar(value=self.tr("progress_ready"))
        self.hash_stat_label = StringVar(value="")
        self.templates: list[TemplateFile] = []
        self.ignored: list[Path] = []
        self.template_inclusion: dict[Path, bool] = {}
        self.multi_dept_values: dict[str, BooleanVar] = {}
        self._tooltips: list[Tooltip] = []
        self._queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._worker_active = False
        self._current_width = 1120
        self._current_breakpoint = "wide"
        self._progressbar: ttk.Progressbar | None = None
        self._history_lock = threading.Lock()
        self._update_check_running = False

        self.root.title(self.tr("app_title"))
        self.root.geometry("1120x820")
        self.root.minsize(980, 720)
        try:
            _set_app_icon(self.root)
        except Exception:
            pass

        self._configure_style()
        self._apply_theme_root()
        self._build_scaffold()
        self._build_header()
        self._build_body()
        self._install_drain_loop()
        self.root.bind("<Configure>", self._on_root_resize)
        self.refresh_templates()
        self._bind_shortcuts()
        self._install_file_drop()
        self.root.after(SYSTEM_THEME_POLL_MS, self._follow_system_theme)
        threading.Thread(target=cleanup_old_previews, daemon=True).start()
        if self.settings.get("check_updates") and \
                self.settings.get("last_update_check") != date.today().isoformat():
            self.root.after(2500, lambda: self.check_for_updates(manual=False))

    # ----------------------------- Utilities ------------------------------
    def tr(self, key: str, **kwargs: Any) -> str:
        raw = self.language.get(key, key)
        try:
            return raw.format(**kwargs) if kwargs else raw
        except (KeyError, IndexError):
            return raw

    def _follow_system_theme(self) -> None:
        """Con il tema "Sistema" segue i cambi di Windows chiaro/scuro mentre l'app e' aperta."""
        try:
            if self.theme_pref == "system" and not self._worker_active:
                wanted = resolve_theme("system")
                if wanted != self.theme.get():
                    self.theme.set(wanted)
                    self._rebuild_ui()
        finally:
            self.root.after(SYSTEM_THEME_POLL_MS, self._follow_system_theme)

    def _persist_settings(self) -> None:
        self.settings["theme"] = self.theme_pref
        self.settings["language"] = self.language_code
        self.settings["last_template_dir"] = self.template_dir.get()
        self.settings["last_output_dir"] = self.output_dir.get()
        save_settings(self.settings)

    # ----------------------------- Style ----------------------------------
    def _apply_theme_root(self):
        theme = self.theme.get()
        if theme == "dark":
            self.root.configure(bg="#0b141b")
        else:
            self.root.configure(bg="#f0f4f6")
        _set_titlebar_dark(self.root, theme == "dark")
        if not getattr(self, "_toplevel_titlebar_bound", False):
            # Anche le finestre secondarie (impostazioni, storico, ...) seguono il tema
            self.root.bind_class(
                "Toplevel", "<Map>",
                lambda e: _set_titlebar_dark(e.widget, self.theme.get() == "dark")
                if e.widget.winfo_toplevel() is e.widget else None, add="+")
            self._toplevel_titlebar_bound = True

    def _configure_style(self):
        dark = self.theme.get() == "dark"
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        if dark:
            app_bg = "#0b141b"
            card_bg = "#13202a"
            card_body_bg = "#13202a"
            sh1 = "#060c10"
            sh2 = "#1f303b"
            text = "#e8eff2"
            text_muted = "#8fa7b1"
            gold = "#e0a640"
            accent_bg = "#08121a"
            accent_fg = "#e8eff2"
            title_bg = "#08121a"
            title_fg = "#f4f8fa"
            subtitle_fg = "#8fa7b1"
            section_bg = card_bg
            section_fg = text
            accent_label_bg = card_bg
            accent_label_fg = gold
            muted_bg = card_bg
            muted_fg = text_muted
            appmuted_bg = app_bg
            appmuted_fg = text_muted
            count_bg = "#10393a"
            count_fg = "#9fe3d8"
            secure_bg = count_bg
            secure_fg = count_fg
            primary_bg = "#2aa19c"
            primary_fg = "#03201f"
            primary_hover = "#36b8b2"
            primary_press = "#238a86"
            secondary_bg = "#1f303b"
            secondary_fg = text
            secondary_hover = "#2a3f4c"
            secondary_press = "#35505f"
            accent_btn_bg = gold
            accent_btn_fg = "#2a1a00"
            field_bg = "#0f1a22"
            field_fg = text
            border = "#2a3d48"
            focus = "#2aa19c"
            tree_bg = card_bg
            tree_field = card_bg
            tree_fg = text
            tree_head_bg = "#1a2a34"
            tree_head_fg = "#a9c3cc"
            tree_head_hover = "#223744"
            tree_sel_bg = "#1d4d4c"
            tree_sel_fg = text
            scroll_bg = "#2a3d48"
            scroll_trough = card_bg
            check_bg = card_bg
            check_fg = text
            gold_bg = "#3a2a08"
            gold_fg = gold
            row_even = "#172631"
            row_tutti = "#33280c"
            row_modified = "#4a3c05"
            row_new = "#12361f"
        else:
            app_bg = "#f0f4f6"
            card_bg = "#ffffff"
            card_body_bg = "#ffffff"
            sh1 = "#d5dfe4"
            sh2 = "#e6edf0"
            text = "#11293a"
            text_muted = "#5b7280"
            gold = "#e0a640"
            accent_bg = "#0b2a3d"
            accent_fg = "#ffffff"
            title_bg = "#0b2a3d"
            title_fg = "#ffffff"
            subtitle_fg = "#bcd5dd"
            section_bg = card_bg
            section_fg = text
            accent_label_bg = card_bg
            accent_label_fg = "#c08a26"
            muted_bg = card_bg
            muted_fg = text_muted
            appmuted_bg = app_bg
            appmuted_fg = text_muted
            count_bg = "#e2f3f0"
            count_fg = "#17706b"
            secure_bg = count_bg
            secure_fg = count_fg
            primary_bg = "#1f8a87"
            primary_fg = "#ffffff"
            primary_hover = "#187370"
            primary_press = "#125a58"
            secondary_bg = "#edf2f4"
            secondary_fg = text
            secondary_hover = "#dfe8eb"
            secondary_press = "#cfdce0"
            accent_btn_bg = gold
            accent_btn_fg = "#2a1a00"
            field_bg = "#f8fafb"
            field_fg = text
            border = "#d3dee3"
            focus = "#1f8a87"
            tree_bg = "#ffffff"
            tree_field = "#ffffff"
            tree_fg = text
            tree_head_bg = "#eef4f5"
            tree_head_fg = "#33505c"
            tree_head_hover = "#e1ebed"
            tree_sel_bg = "#cdeae6"
            tree_sel_fg = text
            scroll_bg = "#c9d6db"
            scroll_trough = "#f3f6f8"
            check_bg = card_bg
            check_fg = text
            gold_bg = "#fcf1dd"
            gold_fg = "#9a6716"
            row_even = "#f6fafa"
            row_tutti = "#fff8ea"
            row_modified = "#fff5c2"
            row_new = "#dcf5e3"

        style.configure("App.TFrame", background=app_bg)
        style.configure("Card.TFrame", background=card_bg)
        style.configure("CardBody.TFrame", background=card_body_bg)
        style.configure("Shadow1.TFrame", background=sh1)
        style.configure("Shadow2.TFrame", background=sh2)

        style.configure("Title.TLabel", background=title_bg, foreground=title_fg,
                        font=("Segoe UI Semibold", 26, "bold"))
        style.configure("Subtitle.TLabel", background=title_bg, foreground=subtitle_fg,
                        font=("Segoe UI", 10))
        style.configure("Eyebrow.TLabel", background=title_bg, foreground=gold,
                        font=("Segoe UI", 8, "bold"))

        style.configure("Section.TLabel", background=section_bg, foreground=section_fg,
                        font=("Segoe UI Semibold", 13, "bold"))
        style.configure("SectionAccent.TLabel", background=accent_label_bg, foreground=accent_label_fg,
                        font=("Segoe UI", 9, "bold"))
        style.configure("Muted.TLabel", background=muted_bg, foreground=muted_fg,
                        font=("Segoe UI", 9))
        style.configure("AppMuted.TLabel", background=appmuted_bg, foreground=appmuted_fg,
                        font=("Segoe UI", 9))
        style.configure("FieldLabel.TLabel", background=card_bg, foreground=text,
                        font=("Segoe UI Semibold", 9, "bold"))

        style.configure("Count.TLabel", background=count_bg, foreground=count_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(12, 6))
        style.configure("Gold.TLabel", background=gold_bg, foreground=gold_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(12, 6))
        style.configure("Secure.TLabel", background=secure_bg, foreground=secure_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(12, 6))

        style.configure("Primary.TButton", background=primary_bg, foreground=primary_fg,
                        font=("Segoe UI Semibold", 10, "bold"), padding=(18, 11),
                        borderwidth=0, focusthickness=0)
        style.map("Primary.TButton",
                  background=[("pressed", primary_press), ("active", primary_hover)])

        style.configure("Secondary.TButton", background=secondary_bg, foreground=secondary_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(14, 9),
                        borderwidth=0, focusthickness=0)
        style.map("Secondary.TButton",
                  background=[("pressed", secondary_press), ("active", secondary_hover)])

        style.configure("Accent.TButton", background=accent_btn_bg, foreground=accent_btn_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(14, 9),
                        borderwidth=0, focusthickness=0)
        style.map("Accent.TButton",
                  background=[("active", "#c08c2d"), ("pressed", "#a4741d")])

        style.configure("TEntry", fieldbackground=field_bg, foreground=field_fg,
                        bordercolor=border, lightcolor=border, darkcolor=border,
                        padding=9, focusthickness=2, focuscolor=focus,
                        font=("Segoe UI", 10), borderwidth=1, relief="solid")
        style.map("TEntry",
                  bordercolor=[("focus", focus), ("!focus", border)],
                  lightcolor=[("focus", focus), ("!focus", border)],
                  darkcolor=[("focus", focus), ("!focus", border)])

        style.configure("TCombobox", fieldbackground=field_bg, foreground=field_fg,
                        background=field_bg, arrowcolor=tree_head_fg,
                        bordercolor=border, lightcolor=border, darkcolor=border,
                        padding=8, focusthickness=2, focuscolor=focus,
                        font=("Segoe UI", 10), borderwidth=1, relief="solid")
        style.map("TCombobox",
                  bordercolor=[("focus", focus), ("!focus", border)],
                  lightcolor=[("focus", focus), ("!focus", border)],
                  darkcolor=[("focus", focus), ("!focus", border)],
                  background=[("readonly", field_bg), ("active", field_bg)],
                  fieldbackground=[("readonly", field_bg)])

        style.configure("Treeview", background=tree_bg, fieldbackground=tree_field,
                        foreground=tree_fg, rowheight=30,
                        bordercolor=border, borderwidth=1, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background=tree_head_bg, foreground=tree_head_fg,
                        font=("Segoe UI Semibold", 9, "bold"), padding=(9, 8),
                        relief="flat", borderwidth=0)
        style.map("Treeview",
                  background=[("selected", tree_sel_bg)],
                  foreground=[("selected", tree_sel_fg)])
        style.map("Treeview.Heading",
                  background=[("active", tree_head_hover)])

        style.configure("Vertical.TScrollbar", background=scroll_bg, troughcolor=scroll_trough,
                        bordercolor=scroll_bg, arrowcolor=tree_head_fg, arrowsize=14,
                        relief="flat", borderwidth=0, gripcount=0, width=12)
        style.map("Vertical.TScrollbar",
                  background=[("active", focus), ("disabled", scroll_trough)])

        style.configure("Horizontal.TProgressbar",
                        troughcolor=scroll_trough,
                        background=focus,
                        bordercolor=border,
                        lightcolor=focus,
                        darkcolor=focus,
                        thickness=8)

        style.configure("TCheckbutton", background=check_bg, foreground=check_fg,
                        font=("Segoe UI Semibold", 9), focusthickness=0)
        style.map("TCheckbutton",
                  background=[("active", check_bg)],
                  foreground=[("active", check_fg)])

        self._style_colors = {
            "app_bg": app_bg, "card_bg": card_bg, "card_body_bg": card_body_bg,
            "sh1": sh1, "sh2": sh2, "text": text, "text_muted": text_muted,
            "gold": gold, "accent_bg": accent_bg, "accent_fg": accent_fg,
            "title_bg": title_bg, "title_fg": title_fg, "subtitle_fg": subtitle_fg,
            "accent_label_bg": accent_label_bg, "accent_label_fg": accent_label_fg,
            "section_bg": section_bg, "section_fg": section_fg,
            "muted_bg": muted_bg, "muted_fg": muted_fg,
            "gold_bg": gold_bg, "count_bg": count_bg, "secure_bg": secure_bg,
            "field_bg": field_bg, "border": border, "focus": focus,
            "count_fg": count_fg, "primary_bg": primary_bg,
            "row_even": row_even, "row_tutti": row_tutti,
            "row_modified": row_modified, "row_new": row_new,
            "title_fg_strong": title_fg, "sh_shadow": "#000000" if dark else "#1b3a4b",
        }

        # Elementi arrotondati disegnati con ui_kit (pulsanti, campi, card...)
        if getattr(self, "_kit", None) is None:
            self._kit = UiKit(self.root)
        self._kit.install(style, {
            "surface": card_body_bg, "text": text, "muted": text_muted, "gold": gold,
            "primary": primary_bg, "primary_hover": primary_hover, "primary_press": primary_press,
            "on_primary": primary_fg,
            "accent": accent_btn_bg,
            "accent_hover": "#ecb655" if dark else "#d39a33",
            "accent_press": "#c98f2c" if dark else "#bf8628",
            "on_accent": accent_btn_fg,
            "secondary": secondary_bg, "secondary_hover": secondary_hover,
            "secondary_press": secondary_press, "on_secondary": secondary_fg,
            "secondary_border": "#2c4250" if dark else "#d9e3e7",
            "field": field_bg, "field_border": border,
            "field_hover": "#3d5664" if dark else "#a9bcc4",
            "focus": focus, "check_border": "#4a6472" if dark else "#9fb3bb",
            "disabled_bg": "#1a2832" if dark else "#eef2f4",
            "disabled_fg": "#5d7480" if dark else "#9aabb2",
            "select_bg": "#1d4d4c" if dark else "#cdeae6", "select_fg": text,
            "trough": "#1f303b" if dark else "#e4ecef",
            "thumb": "#34505e" if dark else "#c3d1d7",
            "thumb_hover": "#4a6b7a" if dark else "#9fb4bc",
            "count_bg": count_bg, "gold_bg": gold_bg, "secure_bg": secure_bg,
            "header_bg": title_bg, "header_field": "#12222c" if dark else "#14374d",
            "header_border": "#22394a" if dark else "#255069",
            "header_hover": "#35576a" if dark else "#3a6a85", "header_fg": "#f1f6f8",
        })

    # ---------------------------- Scaffold --------------------------------
    def _build_scaffold(self):
        self._tooltips.clear()
        for child in self.root.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass

    def _set_progress(self, step: int, total: int):
        total = max(1, total)
        step = max(0, min(step, total))
        if self._progressbar is not None:
            self._progressbar["maximum"] = total
            self._progressbar["value"] = step
        self.progress_label.set(
            self.tr("progress_label", step=step, total=total,
                    pct=round(step * 100.0 / total, 1))
        )

    def _post(self, kind: str, payload: Any):
        self._queue.put((kind, payload))

    # ---------------------------- Header ----------------------------------
    def _build_header(self):
        dark = self.theme.get() == "dark"
        header_outer = tk.Frame(self.root, bg=self._style_colors["app_bg"], height=160)
        header_outer.pack(fill=X, side="top")
        header_outer.pack_propagate(False)

        title_bg = self._style_colors["title_bg"]
        c = tk.Canvas(header_outer, height=148, highlightthickness=0, bd=0, bg=title_bg)
        c.pack(fill=X, side="top")

        try:
            if LOGO_HEADER_PNG.exists():
                logo_img = tk.PhotoImage(file=str(LOGO_HEADER_PNG))
            elif ICON_PNG.exists():
                logo_img = tk.PhotoImage(file=str(ICON_PNG))
                logo_img = logo_img.subsample(max(1, logo_img.width() // 64))
            else:
                rgba, ww, hh = _build_in_memory_icon(64, 64)
                logo_img = tk.PhotoImage(data=_png_from_rgba(rgba, ww, hh))
            self._header_logo = logo_img
        except Exception:
            self._header_logo = None

        colors = self._style_colors
        # Logo e testi disegnati direttamente sul canvas: niente riquadri di
        # sfondo pieno, cosi' le decorazioni restano visibili dietro al testo.
        glow = "#1f8a87" if not dark else "#1a6f6c"
        header_state = {"w": 0, "img": None, "job": None}

        def paint_header(_evt=None):
            header_state["job"] = None
            w = max(c.winfo_width(), 1)
            if w == header_state["w"] and _evt is not None:
                return
            header_state["w"] = w
            c.delete("all")
            # Sfondo renderizzato: sfumatura, bagliori sfocati, puntinatura e filetto oro
            header_state["img"] = self._kit.header_image(
                w, 148, title_bg, glow, colors["gold"], colors["primary_bg"])
            c.create_image(0, 0, image=header_state["img"], anchor="nw")
            x = 42
            if self._header_logo is not None:
                c.create_image(x, 72, image=self._header_logo, anchor="w")
                x += self._header_logo.width() + 22
            item = c.create_text(x, 24, anchor="nw", text=self.tr("eyebrow"),
                                 fill=colors["gold"], font=("Segoe UI Variable Text Semibold", 8, "bold"))
            item = c.create_text(x - 2, c.bbox(item)[3] + 1, anchor="nw",
                                 text=self.tr("header_title"), fill=colors["title_fg"],
                                 font=("Segoe UI Variable Display Semib", 27, "bold"))
            c.create_text(x, c.bbox(item)[3] + 1, anchor="nw",
                          text=self.tr("header_subtitle"), fill=colors["subtitle_fg"],
                          font=("Segoe UI Variable Text", 10))

        def schedule_paint(_evt=None):
            if header_state["job"] is None:
                header_state["job"] = c.after(30, paint_header, _evt)

        c.bind("<Configure>", schedule_paint)
        c.after(1, paint_header)

        controls = tk.Frame(header_outer, bg=title_bg)
        controls.place(relx=1.0, x=-42, y=28, anchor="ne")
        row = tk.Frame(controls, bg=title_bg)
        row.pack(anchor="e")

        tk.Label(row, text=self.tr("lbl_theme") + "  ", bg=title_bg,
                 fg=self._style_colors["subtitle_fg"],
                 font=("Segoe UI Semibold", 9, "bold")).pack(side=LEFT)
        theme_choices = ("light", "dark", "system")
        theme_switch = ttk.Combobox(
            row, values=[self.tr("theme_light"), self.tr("theme_dark"), self.tr("theme_system")],
            state="readonly", width=9, style="Header.TCombobox"
        )
        theme_switch.current(theme_choices.index(self.theme_pref))
        theme_switch.pack(side=LEFT, padx=(0, 18))
        try:
            _reg = getattr(self, "_register_local_wheel", None)
            if callable(_reg):
                _reg(theme_switch)
        except Exception:
            pass


        tk.Label(row, text=self.tr("lbl_language") + "  ", bg=title_bg,
                 fg=self._style_colors["subtitle_fg"],
                 font=("Segoe UI Semibold", 9, "bold")).pack(side=LEFT)
        langs = available_languages()
        labels = []
        current_idx = 0
        for i, code in enumerate(langs):
            labels.append(code.upper())
            if code == self.language_var.get():
                current_idx = i
        lang_combo = ttk.Combobox(row, values=labels, state="readonly", width=6,
                                  style="Header.TCombobox")
        lang_combo.current(current_idx)
        lang_combo.pack(side=LEFT)
        try:
            _reg = getattr(self, "_register_local_wheel", None)
            if callable(_reg):
                _reg(lang_combo)
        except Exception:
            pass


        def on_theme(_e=None):
            chosen = theme_switch.current()
            if not 0 <= chosen < len(theme_choices):
                return
            self.theme_pref = theme_choices[chosen]
            self._persist_settings()
            new_theme = resolve_theme(self.theme_pref)
            if new_theme != self.theme.get():
                self.theme.set(new_theme)
                self._rebuild_ui()

        def on_lang(_e=None):
            idx = lang_combo.current()
            if idx < 0 or idx >= len(langs):
                return
            code = langs[idx]
            if code == self.language_code:
                return
            self.language_code = code
            self.language_var.set(code)
            self.language = load_language(code)
            self._persist_settings()
            self._rebuild_ui()

        theme_switch.bind("<<ComboboxSelected>>", on_theme)
        lang_combo.bind("<<ComboboxSelected>>", on_lang)

    def _draw_hgradient(self, canvas, w, y, h, color1, color2):
        """Banda orizzontale alta h con sfumatura da sinistra (color1) a destra (color2)."""
        r1, g1, b1 = canvas.winfo_rgb(color1)
        r2, g2, b2 = canvas.winfo_rgb(color2)
        steps = 64
        seg = w / steps
        for i in range(steps):
            t = i / (steps - 1)
            color = "#%04x%04x%04x" % (int(r1 + (r2 - r1) * t), int(g1 + (g2 - g1) * t),
                                       int(b1 + (b2 - b1) * t))
            canvas.create_rectangle(int(i * seg), y, int((i + 1) * seg) + 1, y + h,
                                    fill=color, outline="")

    def _draw_gradient(self, canvas, w, h, color1, color2):
        steps = max(1, h)
        r1, g1, b1 = canvas.winfo_rgb(color1)
        r2, g2, b2 = canvas.winfo_rgb(color2)
        rr = (r2 - r1) / steps
        rg = (g2 - g1) / steps
        rb = (b2 - b1) / steps
        for i in range(steps):
            nr = int(r1 + rr * i)
            ng = int(g1 + rg * i)
            nb = int(b1 + rb * i)
            color = f"#{nr:04x}{ng:04x}{nb:04x}"
            canvas.create_line(0, i, w, i, fill=color)

    # ----------------------------- Body -----------------------------------
    def _build_body(self):
        outer_wrap = ttk.Frame(self.root, style="App.TFrame")
        outer_wrap.pack(fill=BOTH, expand=True)
        outer_wrap.columnconfigure(0, weight=1)
        outer_wrap.rowconfigure(0, weight=1)

        canvas_wrap = tk.Canvas(outer_wrap, background=self._style_colors["app_bg"],
                                highlightthickness=0, borderwidth=0)
        canvas_wrap.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(outer_wrap, orient="vertical", command=canvas_wrap.yview)
        sb.grid(row=0, column=1, sticky="ns")
        canvas_wrap.configure(yscrollcommand=sb.set)

        scrolled = ttk.Frame(canvas_wrap, style="App.TFrame")
        scrolled_id = canvas_wrap.create_window((0, 0), window=scrolled, anchor="nw")
        self._body_canvas = canvas_wrap
        self._body_scrolled_id = scrolled_id
        self._body_scrolled = scrolled

        def _on_scroll_config(_evt=None):
            canvas_wrap.configure(scrollregion=canvas_wrap.bbox("all"))

        def _on_canvas_config(evt):
            canvas_wrap.itemconfigure(scrolled_id, width=evt.width)

        scrolled.bind("<Configure>", _on_scroll_config)
        canvas_wrap.bind("<Configure>", _on_canvas_config)

        scrollable_widgets_wheel_local: list[Any] = []

        def _register_local_wheel(w):
            scrollable_widgets_wheel_local.append(w)

        self._register_local_wheel = _register_local_wheel

        def _is_in_local_widget(evt):
            try:
                x_root, y_root = evt.x_root, evt.y_root
            except Exception:
                try:
                    x_root = self.root.winfo_pointerx()
                    y_root = self.root.winfo_pointery()
                except Exception:
                    return False
            for w in list(scrollable_widgets_wheel_local):
                try:
                    if not w.winfo_exists():
                        continue
                except Exception:
                    continue
                stack = [w]
                while stack:
                    cur = stack.pop()
                    try:
                        if not cur.winfo_exists():
                            continue
                    except Exception:
                        continue
                    try:
                        x0 = cur.winfo_rootx()
                        y0 = cur.winfo_rooty()
                        x1 = x0 + cur.winfo_width()
                        y1 = y0 + cur.winfo_height()
                    except Exception:
                        stack.extend(list(getattr(cur, "winfo_children", lambda: [])()))
                        continue
                    if x0 <= x_root <= x1 and y0 <= y_root <= y1:
                        return True
                    stack.extend(list(getattr(cur, "winfo_children", lambda: [])()))
            return False

        def _on_wheel(evt):
            if _is_in_local_widget(evt):
                return "break"
            try:
                if canvas_wrap.winfo_exists():
                    canvas_wrap.yview_scroll(int(-1 * (evt.delta / 120)), "units")
            except Exception:
                pass
            return "break"

        def _on_wheel_up(_evt):
            if _is_in_local_widget(_evt):
                return "break"
            try:
                if canvas_wrap.winfo_exists():
                    canvas_wrap.yview_scroll(-3, "units")
            except Exception:
                pass
            return "break"

        def _on_wheel_down(_evt):
            if _is_in_local_widget(_evt):
                return "break"
            try:
                if canvas_wrap.winfo_exists():
                    canvas_wrap.yview_scroll(3, "units")
            except Exception:
                pass
            return "break"

        canvas_wrap.bind_all("<MouseWheel>", _on_wheel, add="+")
        canvas_wrap.bind_all("<Button-4>", _on_wheel_up, add="+")
        canvas_wrap.bind_all("<Button-5>", _on_wheel_down, add="+")

        self._body = body = ttk.Frame(scrolled, style="App.TFrame", padding=(16, 6, 16, 4))
        body.pack(fill=BOTH, expand=True)
        body.columnconfigure(0, weight=5)
        body.columnconfigure(1, weight=4)
        body.rowconfigure(1, weight=1)
        self._body_layout_root = body

        self._build_card1(body)
        self._build_card2(body)
        self._build_card3(body)
        self._build_footer(body)
        self._apply_breakpoint()

    def _card(self, parent, title=None, subtitle=None, accent=None, **kwargs):
        colors = self._style_colors
        # Card arrotondata con bordo sottile e ombra morbida (disegnata da ui_kit)
        panel = RoundedPanel(parent, self._kit, outer_bg=colors["app_bg"],
                             fill=colors["card_body_bg"], border=colors["sh2"],
                             shadow=colors["sh_shadow"],
                             shadow_alpha=150 if self.theme.get() == "dark" else 38,
                             padx=22, pady=18)
        shadow1 = panel
        inner = panel.inner

        head = tk.Frame(inner, bg=colors["card_body_bg"])
        head.pack(fill=X)
        if title or accent:
            # "01  ·  CONFIGURAZIONE" -> pillola col numero + etichetta
            number, label = "", accent or ""
            if accent and "·" in accent:
                number, label = (part.strip() for part in accent.split("·", 1))
            if number:
                kit = self._kit
                badge = kit.pill(kit.px(42), kit.px(42), colors["count_bg"], r=kit.px(12))
                tk.Label(head, text=number, image=badge, compound="center", bd=0,
                         bg=colors["card_body_bg"], fg=colors["count_fg"],
                         font=("Segoe UI Variable Display Semib", 13, "bold"),
                         ).pack(side=LEFT, padx=(0, 14), anchor="center")
            left = tk.Frame(head, bg=colors["card_body_bg"])
            left.pack(side=LEFT, fill=X, expand=True)
            if label:
                ttk.Label(left, text=label, style="SectionAccent.TLabel").pack(anchor="w")
            if title:
                ttk.Label(left, text=title, style="Section.TLabel").pack(anchor="w", pady=(1, 0))
            # Divisore: breve tratto oro su linea sottile
            divider = tk.Frame(inner, bg=colors["card_body_bg"], height=3)
            divider.pack(fill=X, pady=(16, 18))
            tk.Frame(divider, bg=colors["sh2"]).place(x=0, y=1, relwidth=1.0, height=1)
            accent_line = self._kit.pill(self._kit.px(48), 3, colors["gold"])
            tk.Label(divider, image=accent_line, bd=0, bg=colors["card_body_bg"]
                     ).place(x=0, y=0, height=3)

        if subtitle:
            ttk.Label(inner, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(0, 16))

        body = tk.Frame(inner, bg=colors["card_body_bg"])
        body.pack(fill=BOTH, expand=True)
        return shadow1, body

    def _field_label(self, parent, text, row, col=0, span=2, label_col_width=160):
        lbl = tk.Label(
            parent, text=text, bg=self._style_colors["card_body_bg"],
            fg=self._style_colors["text"],
            font=("Segoe UI Semibold", 9, "bold"), anchor="w",
        )
        lbl.grid(row=row, column=col, sticky="we", padx=(0, 16), pady=(0, 6))
        parent.grid_columnconfigure(col, minsize=label_col_width)
        return lbl

    def _wrap_field(self, parent, row, widget, col=0, span=2, pad_bottom=14):
        container = tk.Frame(parent, bg=self._style_colors["card_body_bg"])
        container.grid(row=row, column=col + 1, sticky="nsew", columnspan=span - 1, pady=(0, pad_bottom))
        container.grid_columnconfigure(0, weight=1)
        widget_master = widget.master
        widget.reparent = None  # noop
        if widget_master is not container:
            widget.pack_forget() if hasattr(widget, "pack_forget") else None
            widget.grid_forget() if hasattr(widget, "grid_forget") else None
        widget.grid(in_=container, row=0, column=0, sticky="ew")
        return container

    # ------------------------- Card 1 (Setup) -----------------------------
    def _build_card1(self, body):
        src_shadow, src_body = self._card(
            body,
            title=self.tr("card01_title"),
            accent=self.tr("card01_accent"),
            subtitle=self.tr("card01_subtitle"),
        )
        src_shadow.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 2))
        src_body.columnconfigure(1, weight=1)
        src_body.columnconfigure(2, weight=0)

        def add_row(r, label_text, var, btn_cmd):
            tk.Label(src_body, text=label_text,
                     bg=self._style_colors["card_body_bg"], fg=self._style_colors["text"],
                     font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                     ).grid(row=r, column=0, sticky="we", padx=(0, 18), pady=(0, 4))
            ent = ttk.Entry(src_body, textvariable=var)
            ent.grid(row=r, column=1, sticky="ew", padx=(0, 10), pady=(0, 14))
            btn = self._button(src_body, "btn_choose_folder",
                             style="Secondary.TButton", command=btn_cmd)
            btn.grid(row=r, column=2, sticky="ew", pady=(0, 14))
            return ent, btn

        tpl_entry, tpl_btn = add_row(0, self.tr("lbl_templates"), self.template_dir,
                                     self.choose_template_dir)
        out_entry, out_btn = add_row(1, self.tr("lbl_output"), self.output_dir,
                                     self.choose_output_dir)
        self._add_tooltip(tpl_entry, lambda: self.tr("tt_choose_tpl"))
        self._add_tooltip(tpl_btn, lambda: self.tr("tt_choose_tpl"))
        self._add_tooltip(out_entry, lambda: self.tr("tt_choose_out"))
        self._add_tooltip(out_btn, lambda: self.tr("tt_choose_out"))

        action_row = tk.Frame(src_body, bg=self._style_colors["card_body_bg"])
        action_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        refresh_btn = self._button(action_row, "btn_refresh",
                                 style="Accent.TButton", command=self.refresh_templates)
        refresh_btn.pack(side=LEFT)
        self._add_tooltip(refresh_btn, lambda: self.tr("tt_refresh") + "  (F5)")
        dept_btn = self._button(action_row, "btn_depts",
                              style="Secondary.TButton", command=self.open_department_editor)
        dept_btn.pack(side=LEFT, padx=(10, 0))
        self._add_tooltip(dept_btn, lambda: self.tr("tt_depts"))
        modules_btn = self._button(action_row, "btn_modules",
                                 style="Secondary.TButton", command=self.open_template_manager)
        modules_btn.pack(side=LEFT, padx=(10, 0))
        self._add_tooltip(modules_btn, lambda: self.tr("tt_modules"))
        settings_btn = self._button(action_row, "btn_settings",
                                  style="Secondary.TButton", command=self.open_settings)
        settings_btn.pack(side=RIGHT)
        self._add_tooltip(settings_btn, lambda: self.tr("tt_settings") + "  (Ctrl+,)")
        history_btn = self._button(action_row, "btn_history",
                                 style="Secondary.TButton", command=self.open_history)
        history_btn.pack(side=RIGHT, padx=(0, 10))
        self._add_tooltip(history_btn, lambda: self.tr("tt_history") + "  (Ctrl+H)")

    # ------------------------- Card 2 (Employee) --------------------------
    def _build_card2(self, body):
        form_shadow, form_body = self._card(
            body,
            title=self.tr("card02_title"),
            accent=self.tr("card02_accent"),
            subtitle=self.tr("card02_subtitle"),
        )
        self._card2_shadow = form_shadow
        form_shadow.grid(row=1, column=0, sticky="nsew", padx=(0, 2))
        form_body.columnconfigure(1, weight=1)

        def add_row(r, label_text, widget_or_constructor):
            tk.Label(form_body, text=label_text,
                     bg=self._style_colors["card_body_bg"], fg=self._style_colors["text"],
                     font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                     ).grid(row=r, column=0, sticky="we", padx=(0, 18), pady=(0, 4))
            w = widget_or_constructor(form_body)
            w.grid(row=r, column=1, sticky="ew", pady=(0, 14))
            return w

        name_entry = add_row(0, self.tr("lbl_name"),
            lambda parent: ttk.Entry(parent, textvariable=self.employee_name))
        self._add_tooltip(name_entry, lambda: self.tr("tt_name") + "  " + self.tr("sc_enter_hint"))
        name_entry.bind("<Return>", lambda _e: self._inline_batch_add_current())
        try:
            _reg = getattr(self, "_register_local_wheel", None)
            if callable(_reg):
                _reg(name_entry)
        except Exception:
            pass


        dp = add_row(1, self.tr("lbl_date"),
            lambda parent: DatePickerFrame(parent, self.language,
                                           bg=self._style_colors["card_body_bg"]))
        self.date_picker = dp
        self._add_tooltip(self.date_picker, lambda: self.tr("tt_date"))
        try:
            _reg = getattr(self, "_register_local_wheel", None)
            if callable(_reg):
                _reg(self.date_picker)
        except Exception:
            pass

        for w in (self.date_picker.day_cb, self.date_picker.month_cb, self.date_picker.year_cb):
            self._add_tooltip(w, lambda: self.tr("tt_date"))

        # --- Section: Batch inline (add people + run directly) ---
        batch_box = tk.Frame(form_body, bg=self._style_colors["card_body_bg"])
        batch_box.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 8))
        batch_box.columnconfigure(0, weight=0)
        batch_box.columnconfigure(1, weight=1)
        batch_box.rowconfigure(0, weight=1)

        bt = tk.Label(
            batch_box,
            text="   " + self.tr("bat_inline_title"),
            bg=self._style_colors["count_bg"], fg=self._style_colors["count_fg"],
            font=("Segoe UI Semibold", 10, "bold"), anchor="w", pady=7,
        )
        bt.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(batch_box, text=self.tr("bat_inline_subtitle"),
                  style="Muted.TLabel"
                  ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        people_tree_wrap = tk.Frame(batch_box, bg=self._style_colors["card_body_bg"])
        people_tree_wrap.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        people_tree_wrap.rowconfigure(0, weight=1)
        people_tree_wrap.columnconfigure(0, weight=1)
        batch_body_columns = ("nome", "data", "reparto")
        people_tree = ttk.Treeview(people_tree_wrap, columns=batch_body_columns,
                                   show="headings", height=8)
        for col_key, heading, width in (
            ("nome", self.tr("col_name"), 200),
            ("data", self.tr("col_date"), 150),
            ("reparto", self.tr("col_dept"), 160),
        ):
            people_tree.heading(col_key, text=heading)
            people_tree.column(col_key, width=width, anchor="w")
        people_tree.grid(row=0, column=0, sticky="nsew")
        p_sb = ttk.Scrollbar(people_tree_wrap, orient="vertical", command=people_tree.yview)
        p_sb.grid(row=0, column=1, sticky="ns")
        people_tree.configure(yscrollcommand=p_sb.set)
        self._inline_batch_tree = people_tree
        self._inline_batch_rows: list[dict[str, str]] = []
        try:
            _reg = getattr(self, "_register_local_wheel", None)
            if callable(_reg):
                _reg(self._inline_batch_tree)
        except Exception:
            pass


        actions = tk.Frame(batch_box, bg=self._style_colors["card_body_bg"])
        actions.grid(row=3, column=0, columnspan=2, sticky="ew")
        # Due righe: modifica dell'elenco sopra, import/modello da file sotto
        list_actions = tk.Frame(actions, bg=self._style_colors["card_body_bg"])
        list_actions.pack(fill=X)
        file_actions = tk.Frame(actions, bg=self._style_colors["card_body_bg"])
        file_actions.pack(fill=X, pady=(6, 0))
        btn_add = self._button(list_actions, "bat_add_person",
                             style="Secondary.TButton",
                             command=self._inline_batch_add_current)
        btn_add.pack(side=LEFT)
        btn_del = self._button(list_actions, "bat_del_person",
                             style="Secondary.TButton",
                             command=self._inline_batch_remove_selected)
        btn_del.pack(side=LEFT, padx=(6, 0))
        btn_clear = self._button(list_actions, "bat_clear",
                               style="Secondary.TButton",
                               command=self._inline_batch_clear)
        btn_clear.pack(side=LEFT, padx=(6, 0))
        btn_load = self._button(file_actions, "bat_import_file",
                              style="Secondary.TButton",
                              command=self._inline_batch_import_file)
        btn_load.pack(side=LEFT)
        btn_model = self._button(file_actions, "bat_template_download",
                               style="Secondary.TButton",
                               command=self.download_batch_template)
        btn_model.pack(side=LEFT, padx=(6, 0))
        self._add_tooltip(btn_model, lambda: self.tr("tt_batch_template"))

        batch_run_wrap = tk.Frame(batch_box, bg=self._style_colors["card_body_bg"])
        batch_run_wrap.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        tk.Frame(batch_run_wrap, bg=self._style_colors["gold"],
                 height=2).pack(fill=X, side="top", pady=(0, 10))
        run_count_lbl = ttk.Label(batch_run_wrap, textvariable=self._inline_batch_count(),
                                  style="Count.TLabel")
        run_count_lbl.pack(side=LEFT)
        btn_run = self._button(batch_run_wrap, "bat_run_all", style="Primary.TButton",
                             command=self._inline_batch_run)
        btn_run.pack(side=RIGHT, ipadx=14, ipady=5)
        self._inline_batch_run_btn = btn_run
        self._refresh_inline_batch_count()
        self._add_tooltip(btn_run, lambda: self.tr("tt_batch_run") + "  (Ctrl+B)")
        self._add_tooltip(btn_add, lambda: self.tr("tt_batch_add"))
        self._add_tooltip(btn_del, lambda: self.tr("tt_batch_del"))
        self._add_tooltip(btn_load, lambda: self.tr("tt_batch_import"))
        self._add_tooltip(btn_clear, lambda: self.tr("tt_batch_clear"))

        form_body.rowconfigure(2, weight=1)
        batch_box.rowconfigure(2, weight=1)

        bottom_form = tk.Frame(form_body, bg=self._style_colors["card_body_bg"])
        bottom_form.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        auto_cb = ttk.Checkbutton(bottom_form, text=self.tr("cb_autoopen"),
                                  variable=self.auto_open)
        auto_cb.pack(side=LEFT)
        self._add_tooltip(auto_cb, lambda: self.tr("tt_autoopen"))

    def _inline_batch_count(self) -> StringVar:
        self._inline_batch_count_var = var = StringVar()

        def _fmt(n):
            if n == 0:
                return self.tr("bat_count_none")
            return self.tr("bat_count", n=n)

        self._inline_batch_count_fmt = _fmt
        self._refresh_inline_batch_count()
        return var

    def _refresh_inline_batch_count(self):
        fmt = getattr(self, "_inline_batch_count_fmt", None)
        var = getattr(self, "_inline_batch_count_var", None)
        if var is None or fmt is None:
            return
        n = len(getattr(self, "_inline_batch_rows", []))
        var.set(fmt(n))
        run_btn = getattr(self, "_inline_batch_run_btn", None)
        # Dopo un cambio di tema/lingua il riferimento puo' puntare al pulsante distrutto
        if run_btn is not None and run_btn.winfo_exists():
            text = self.tr("bat_run_all", n=n)
            if run_btn.cget("image"):
                text = " " + strip_leading_symbol(text)
            run_btn.configure(text=text)

    def _current_dept_for_batch(self) -> str:
        depts = self._current_departments()
        return depts[0] if len(depts) == 1 else "+".join(depts)

    def _inline_batch_add_current(self):
        name = self.employee_name.get().strip()
        if not name:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_name"))
            return
        entry_date = self.date_picker.get_string()
        if not entry_date:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_date"))
            return
        depts = self._current_departments()
        if not depts:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_dept"))
            return
        if not self._current_selected_templates():
            messagebox.showwarning(self.tr("mb_no_docs_title"), self.tr("mb_no_docs_body"))
            return
        dept_str = "+".join(depts)
        row = {"Nome": name, "Data": entry_date, "Reparto": dept_str}
        self._inline_batch_rows.append(row)
        tree = self._inline_batch_tree
        tree.insert("", END, values=(row["Nome"], row["Data"], row["Reparto"]))
        self._refresh_inline_batch_count()
        self.employee_name.set("")

    def _inline_batch_remove_selected(self):
        tree = self._inline_batch_tree
        sel = tree.selection()
        if not sel:
            return
        # Map tree children index to rows (insert order = _inline_batch_rows)
        children = tree.get_children()
        idxs_to_remove = set()
        for s in sel:
            try:
                idxs_to_remove.add(children.index(s))
            except ValueError:
                continue
        new_rows = [r for i, r in enumerate(self._inline_batch_rows)
                    if i not in idxs_to_remove]
        self._inline_batch_rows = new_rows
        for s in sel:
            tree.delete(s)
        self._refresh_inline_batch_count()

    def _inline_batch_clear(self):
        self._inline_batch_rows.clear()
        tree = self._inline_batch_tree
        for c in tree.get_children():
            tree.delete(c)
        self._refresh_inline_batch_count()

    def _inline_batch_import_file(self, p: Path | None = None):
        if p is None:
            p = filedialog.askopenfilename(
                title=self.tr("bat_choose"),
                initialdir=self.template_dir.get(),
                filetypes=[("CSV / Excel", "*.csv *.xlsx"), ("All", "*.*")],
            )
        if not p:
            return
        before = len(self._inline_batch_rows)
        try:
            loaded = self._parse_batch_file(Path(p))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(self.tr("mb_error_title"), str(exc))
            return
        tree = self._inline_batch_tree
        for r in loaded:
            nome = (r.get("Nome") or r.get("name") or "").strip()
            data = (r.get("Data") or r.get("date") or "").strip()
            reparto = (r.get("Reparto") or r.get("department") or "").strip().upper()
            if not nome or not data:
                continue
            row = {"Nome": nome, "Data": data, "Reparto": reparto}
            self._inline_batch_rows.append(row)
            tree.insert("", END, values=(nome, data, reparto))
        self._refresh_inline_batch_count()
        self._toast(self.tr("dd_imported", n=len(self._inline_batch_rows) - before,
                            name=Path(p).name), "success")

    def _pdf_options(self) -> dict[str, Any]:
        return {
            "watermark": str(self.settings.get("pdf_watermark") or ""),
            "protect": bool(self.settings.get("pdf_protect")),
        }

    def _inline_batch_run(self):
        rows = list(getattr(self, "_inline_batch_rows", []))
        if not rows:
            messagebox.showwarning(self.tr("bat_summary_title"), self.tr("bat_no_rows"))
            return
        if self._worker_active:
            return
        self._worker_active = True
        self._post("progress_ready_label", None)
        auto_open_batch = bool(self.auto_open.get())
        lang_snap = dict(self.language) if self.language else {}

        def s_tr(key: str, **kw) -> str:
            raw = lang_snap.get(key, key)
            try:
                return raw.format(**kw) if kw else raw
            except Exception:
                return raw

        tpl_snap = list(self.templates)
        inc_snap = dict(self.template_inclusion)
        saved_snap = dict(self.saved_hashes) if self.saved_hashes else {}
        dept_opts = {d.upper() for d in department_options(tpl_snap)}
        out_dir_path = Path(self.output_dir.get()).expanduser()
        pdf_options = self._pdf_options()

        def work2():
            ok = skip = fail = 0
            details: list[str] = []
            # Prima si validano tutte le righe, poi si generano tutti i dossier insieme
            # (una sola sessione Office, oppure in parallelo senza Office).
            jobs: list[DossierJob] = []
            job_rows: list[int] = []
            reserved: set[Path] = set()
            try:
                for idx, row in enumerate(rows, start=1):
                    nome = (row.get("Nome") or "").strip()
                    data = (row.get("Data") or "").strip()
                    reparto_raw = (row.get("Reparto") or "").strip()
                    reparto_list = [d.strip().upper() for d in reparto_raw.split("+") if d.strip()]
                    if not reparto_list:
                        if len(dept_opts) == 1:
                            reparto_list = [next(iter(dept_opts))]
                        else:
                            skip += 1
                            details.append(s_tr("bat_skip_dept", i=idx, d=reparto_raw or ""))
                            continue
                    note = ""
                    ruolo = ""
                    if not nome:
                        skip += 1
                        details.append(s_tr("bat_skip_empty", i=idx))
                        continue
                    parsed = self._parse_date_str(data)
                    if parsed is None:
                        skip += 1
                        details.append(s_tr("bat_skip_date", i=idx))
                        continue
                    invalid_depts = [d for d in reparto_list if d not in dept_opts]
                    if invalid_depts:
                        skip += 1
                        details.append(s_tr("bat_skip_dept", i=idx, d=", ".join(invalid_depts)))
                        continue
                    sel = templates_for_departments(tpl_snap, reparto_list)
                    filt = [t for t in sel if inc_snap.get(t.path, True)]
                    if not filt:
                        skip += 1
                        for d in reparto_list:
                            details.append(s_tr("bat_skip_dept", i=idx, d=d))
                        continue
                    dept_tag = safe_file_part("+".join(reparto_list))
                    out_file = out_dir_path / (
                        f"dossier_{safe_file_part(nome)}_{dept_tag}.pdf"
                    )
                    c = 2
                    # I file del lotto non esistono ancora: si riservano i nomi gia' assegnati
                    while out_file.exists() or out_file in reserved:
                        out_file = out_dir_path / (
                            f"dossier_{safe_file_part(nome)}_{dept_tag}_{c}.pdf"
                        )
                        c += 1
                    reserved.add(out_file)
                    jobs.append(DossierJob(out_file, nome, parsed.strftime("%d/%m/%Y"),
                                           "+".join(reparto_list), filt, ruolo, note))
                    job_rows.append(idx)

                def progress_cb(step: int, total: int):
                    self._post("progress", {"step": step, "total": total})

                results = build_pdfs(jobs, progress_cb, **pdf_options)
                for idx, job, error in zip(job_rows, jobs, results):
                    if error is None:
                        self._save_history(job.output_path, job.employee_name, job.department,
                                           len(job.templates), job.entry_date)
                        ok += 1
                        details.append(f"#{idx} OK · {job.output_path.name}")
                    else:
                        fail += 1
                        details.append(s_tr("bat_fail_generic", i=idx, e=str(error)))
                hashes = dict(saved_snap)
                for t in tpl_snap:
                    try:
                        hashes[str(t.path)] = compute_template_hash(t.path)
                    except Exception:
                        pass
                save_hashes(hashes)
            finally:
                self._post("done_batch", {
                    "ok": ok, "skip": skip, "fail": fail,
                    "detail": "\n".join(details[-30:]),
                    "out_dir": str(out_dir_path),
                    "auto_open": bool(auto_open_batch),
                })
                self._post("worker_done", None)

        threading.Thread(target=work2, daemon=True).start()

    # ----------------------- Card 3 (Summary) -----------------------------
    def _build_card3(self, body):
        prev_shadow, prev_body = self._card(
            body,
            title=self.tr("card03_title"),
            accent=self.tr("card03_accent"),
            subtitle=self.tr("card03_subtitle"),
        )
        self._card3_shadow = prev_shadow
        prev_shadow.grid(row=1, column=1, sticky="nsew")
        prev_body.columnconfigure(0, weight=1)

        # --- Reparto selector ---
        dept_frame = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        dept_frame.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        dept_frame.columnconfigure(1, weight=1)

        multi_cb = ttk.Checkbutton(dept_frame, text=self.tr("cb_multidept"),
                                   variable=self.multi_dept_mode,
                                   command=self._toggle_multi_dept)
        multi_cb.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self._add_tooltip(multi_cb, lambda: self.tr("tt_multidept"))

        tk.Label(dept_frame, text=self.tr("lbl_department"),
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text"],
                 font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                 ).grid(row=1, column=0, sticky="w", padx=(0, 14), pady=(0, 6))

        single_wrap = tk.Frame(dept_frame, bg=self._style_colors["card_body_bg"])
        single_wrap.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        single_wrap.columnconfigure(0, weight=1)
        self.department_combo = ttk.Combobox(
            single_wrap, state="readonly", textvariable=self.department,
            font=("Segoe UI Semibold", 10, "bold"), height=14,
        )
        self.department_combo.grid(row=0, column=0, sticky="ew")
        self.department_combo.bind("<<ComboboxSelected>>", lambda _e: self.update_document_list())
        self._single_dept_wrap = single_wrap
        try:
            _reg = getattr(self, "_register_local_wheel", None)
            if callable(_reg):
                _reg(self.department_combo)
        except Exception:
            pass


        multi_wrap = tk.Frame(dept_frame, bg=self._style_colors["card_body_bg"])
        multi_wrap.grid(row=2, column=0, columnspan=2, sticky="ew")
        multi_wrap.grid_remove()
        self._multi_dept_wrap = multi_wrap
        self._render_multi_dept_list(multi_wrap)

        tk.Label(dept_frame, text="  " + self.tr("lbl_department_hint").lstrip(),
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
                 font=("Segoe UI", 8), anchor="w",
                 ).grid(row=3, column=0, columnspan=2, sticky="w")

        self._add_tooltip(self.department_combo, lambda: self.tr("tt_dept"))

        # --- Badge count + select all/none ---
        badge_row = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        badge_row.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(badge_row, textvariable=self.count_label, style="Count.TLabel").pack(side=LEFT)
        right_badges = tk.Frame(badge_row, bg=self._style_colors["card_body_bg"])
        right_badges.pack(side=RIGHT)
        btn_all = self._button(right_badges, "btn_selectall",
                             style="Secondary.TButton",
                             command=lambda: self._set_all_inclusion(True))
        btn_none = self._button(right_badges, "btn_selectnone",
                              style="Secondary.TButton",
                              command=lambda: self._set_all_inclusion(False))
        btn_all.pack(side=LEFT, padx=(0, 6))
        btn_none.pack(side=LEFT)
        self._add_tooltip(btn_all, lambda: self.tr("tt_include"))
        self._add_tooltip(btn_none, lambda: self.tr("tt_include"))

        # --- Treeview with checkboxes ---
        tree_wrap = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        tree_wrap.grid(row=2, column=0, sticky="nsew", pady=(0, 4))
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)
        prev_body.rowconfigure(2, weight=1)

        self.tree = ttk.Treeview(
            tree_wrap,
            columns=("include", "documento", "copie", "stato"),
            show="headings",
            height=10,
            selectmode="none",
        )
        self.tree.heading("include", text=self.tr("col_include"))
        self.tree.heading("documento", text=self.tr("col_document"))
        self.tree.heading("copie", text=self.tr("col_copies"))
        self.tree.heading("stato", text=self.tr("col_status"))
        self.tree.column("include", width=46, anchor="center", stretch=False)
        self.tree.column("documento", width=280, anchor="w", stretch=True)
        self.tree.column("copie", width=70, anchor="center", stretch=False)
        self.tree.column("stato", width=96, anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_sb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        tree_sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=tree_sb.set)
        self.tree.bind("<Button-1>", self._on_tree_click)
        try:
            _reg = getattr(self, "_register_local_wheel", None)
            if callable(_reg):
                _reg(self.tree)
        except Exception:
            pass

        self.tree.bind("<space>", lambda _e: self._toggle_focused_row())
        self._add_tooltip(self.tree, lambda: self.tr("tt_include"))

        # --- Colors hash status tags ---
        colors = self._style_colors
        self.tree.tag_configure("odd", background=colors["card_body_bg"])
        self.tree.tag_configure("even", background=colors["row_even"])
        self.tree.tag_configure("tutti", background=colors["row_tutti"])
        self.tree.tag_configure("modified", background=colors["row_modified"])
        self.tree.tag_configure("new", background=colors["row_new"])

        # --- Progress + buttons ---
        prog_wrap = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        prog_wrap.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        self._progressbar = ttk.Progressbar(prog_wrap, orient="horizontal",
                                             mode="determinate", maximum=100, value=0)
        self._progressbar.pack(fill=X, side="top")
        ttk.Label(prog_wrap, textvariable=self.progress_label, style="Muted.TLabel"
                  ).pack(anchor="w", side="top", pady=(4, 0))

        gen_frame = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        gen_frame.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        gen_buttons = tk.Frame(gen_frame, bg=self._style_colors["card_body_bg"])
        gen_buttons.pack(fill=X)
        gen_btn = self._button(gen_buttons, "btn_generate",
                             style="Primary.TButton", command=self.generate)
        gen_btn.pack(side=RIGHT, ipadx=18, ipady=5)
        preview_btn = self._button(gen_buttons, "btn_preview",
                                 style="Secondary.TButton", command=self.preview)
        preview_btn.pack(side=RIGHT, padx=(0, 10), ipady=5)
        self._add_tooltip(preview_btn, lambda: self.tr("tt_preview") + "  (Ctrl+P)")
        tk.Label(gen_frame, text=self.tr("lbl_generate_hint"),
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
                 font=("Segoe UI", 8), anchor="e",
                 ).pack(side=RIGHT, pady=(4, 0))
        self._add_tooltip(gen_btn, lambda: self.tr("tt_generate") + "  (Ctrl+G)")

        help_frame = tk.Frame(prev_body, bg=self._style_colors["card_body_bg"])
        help_frame.grid(row=5, column=0, sticky="ew", pady=(14, 0))
        tk.Frame(help_frame, bg=self._style_colors["count_bg"], width=4, height=56).pack(side=LEFT)
        help_wrap_length = 380
        tip = tk.Label(
            help_frame, text=self.tr("help_tip_body"),
            bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
            font=("Segoe UI", 9), justify="left", anchor="w",
            padx=12, pady=8, wraplength=help_wrap_length,
        )
        self._help_tip_label = tip
        tip.pack(side=LEFT, fill=X, expand=True)

        self._toggle_multi_dept()

    # ---------------------- Footer ----------------------------------------
    def _build_footer(self, body_container):
        footer_outer = tk.Frame(self.root, bg=self._style_colors["app_bg"])
        footer_outer.pack(fill=X, side="bottom")
        colors = self._style_colors
        footer_panel = RoundedPanel(footer_outer, self._kit, outer_bg=colors["app_bg"],
                                    fill=colors["card_body_bg"], border=colors["sh2"],
                                    shadow=colors["sh_shadow"],
                                    shadow_alpha=150 if self.theme.get() == "dark" else 38,
                                    radius=12, padx=16, pady=4)
        footer_panel.pack(fill=X, padx=16, pady=(0, 8))
        footer = footer_panel.inner

        status_wrap = tk.Frame(footer, bg=self._style_colors["card_body_bg"])
        status_wrap.pack(side=LEFT, fill=X, expand=True)
        tk.Label(status_wrap, text="●", bg=self._style_colors["card_body_bg"],
                 fg=self._style_colors["focus"], font=("Segoe UI", 10, "bold"),
                 ).pack(side=LEFT)
        tk.Label(status_wrap, textvariable=self.status,
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text"],
                 font=("Segoe UI Semibold", 9), anchor="w", padx=8,
                 ).pack(side=LEFT, fill=X, expand=True)
        tk.Label(status_wrap, textvariable=self.hash_stat_label,
                 bg=self._style_colors["card_body_bg"], fg=self._style_colors["text_muted"],
                 font=("Segoe UI", 8), anchor="w", padx=14,
                 ).pack(side=LEFT)

        secure = ttk.Label(footer, text=self.tr("secure_label"), style="Secure.TLabel")
        secure.pack(side=RIGHT)
        self._footer = footer_outer

    # ---------------------- Tooltip helper --------------------------------
    def _icon(self, name: str | None, style: str = "TButton"):
        """PhotoImage dell'icona nel colore del testo dello stile ttk (None se assente)."""
        if not name:
            return None
        try:
            color = str(ttk.Style().lookup(style, "foreground") or self._style_colors["text"])
            if not color.startswith("#"):
                color = "#%02x%02x%02x" % tuple(c // 257 for c in self.root.winfo_rgb(color))
            size = max(14, round(16 * self.root.winfo_fpixels("1i") / 96))
            cache = self.__dict__.setdefault("_icon_cache", {})
            key = (name, color.lower(), size)
            if key not in cache:
                png = draw_icon_png(name, color, size)
                cache[key] = tk.PhotoImage(data=base64.b64encode(png)) if png else None
            return cache[key]
        except Exception:  # noqa: BLE001 - senza PIL si resta con il testo
            return None

    def _button(self, parent, key: str, **kwargs):
        """ttk.Button con testo tradotto e icona a sinistra al posto dell'emoji."""
        text = self.tr(key)
        image = self._icon(BUTTON_ICONS.get(key), kwargs.get("style", "TButton"))
        if image is not None:
            kwargs.update(image=image, compound="left")
            text = " " + strip_leading_symbol(text)
        return ttk.Button(parent, text=text, **kwargs)

    # ---------------------- Notifiche non bloccanti ----------------------
    def _toast(self, message: str, kind: str = "info", action=None, duration: int = 6000):
        """Avviso in basso nella finestra che sparisce da solo; action=(testo, funzione)."""
        colors = self._style_colors
        self._close_toast()
        bg = colors["title_bg"]
        accent = {"success": "#3fb27f", "warning": colors["gold"],
                  "error": "#e5534b"}.get(kind, colors["primary_bg"])
        frame = tk.Frame(self.root, bg=bg, padx=14, pady=10)
        tk.Frame(frame, bg=accent, width=4).pack(side=LEFT, fill="y", padx=(0, 12))
        tk.Label(frame, text=message, bg=bg, fg="#ffffff", font=("Segoe UI Semibold", 9),
                 justify="left", anchor="w", wraplength=560).pack(side=LEFT)
        if action is not None:
            label, callback = action
            link = tk.Label(frame, text=label, bg=bg, fg=colors["gold"], cursor="hand2",
                            font=("Segoe UI Semibold", 9, "bold"), padx=14)
            link.pack(side=LEFT)
            link.bind("<Button-1>", lambda _e: (self._close_toast(), callback()))
        close = tk.Label(frame, text="✕", bg=bg, fg="#bcd5dd", cursor="hand2",
                         font=("Segoe UI", 9), padx=4)
        close.pack(side=LEFT, padx=(6, 0))
        close.bind("<Button-1>", lambda _e: self._close_toast())
        frame.place(relx=0.5, rely=1.0, y=-96, anchor="s")
        frame.lift()
        self._toast_frame = frame
        self._toast_after = self.root.after(duration, self._close_toast)

    def _close_toast(self):
        after_id = getattr(self, "_toast_after", None)
        if after_id:
            try:
                self.root.after_cancel(after_id)
            except tk.TclError:
                pass
        frame = getattr(self, "_toast_frame", None)
        if frame is not None:
            try:
                frame.destroy()
            except tk.TclError:
                pass
        self._toast_frame = self._toast_after = None

    # ---------------------- Scorciatoie da tastiera ----------------------
    SHORTCUTS = (
        ("Ctrl+G", "sc_generate"), ("Ctrl+P", "sc_preview"), ("Ctrl+B", "sc_batch"),
        ("Ctrl+H", "sc_history"), ("Ctrl+,", "sc_settings"), ("F5", "sc_refresh"),
        ("sc_key_enter", "sc_enter"), ("F1", "sc_help"),
    )

    def _bind_shortcuts(self):
        def run(action):
            # Le scorciatoie valgono solo nella finestra principale e mai a generazione in corso
            if self.root.grab_current() is None:
                action()
            return "break"

        for sequence, action in (
            ("g", self.generate), ("p", self.preview), ("b", self._inline_batch_run),
            ("h", self.open_history), ("comma", self.open_settings),
        ):
            # Le lettere anche maiuscole (Bloc Maiusc attivo); "comma" e' un keysym fisso
            variants = {sequence, sequence.upper()} if len(sequence) == 1 else {sequence}
            for variant in variants:
                self.root.bind(f"<Control-{variant}>", lambda _e, a=action: run(a))
        # Nei campi di testo Tk usa Ctrl+H come Backspace (stile Emacs): si disattiva,
        # cosi' Ctrl+H apre lo storico senza cancellare l'ultima lettera.
        for entry_class in ("TEntry", "Entry", "TCombobox"):
            for variant in ("<Control-h>", "<Control-H>"):
                self.root.bind_class(entry_class, variant, lambda _e: None)
        self.root.bind("<F5>", lambda _e: run(self.refresh_templates))
        self.root.bind("<F1>", lambda _e: run(self.show_shortcuts))

    def show_shortcuts(self):
        lines = [f"{self.tr(keys) if keys.startswith('sc_') else keys}   →   {self.tr(key)}"
                 for keys, key in self.SHORTCUTS]
        messagebox.showinfo(self.tr("sc_title"), "\n".join(lines))

    # ---------------------- Trascinamento file ---------------------------
    def _install_file_drop(self):
        if os.name != "nt":
            return
        try:
            self._drop_target = FileDropTarget(self.root, self._on_files_dropped)
        except Exception:  # noqa: BLE001 - funzione accessoria
            traceback.print_exc()

    def _on_files_dropped(self, paths: list[Path], x: int, y: int) -> None:
        """CSV, o Excel lasciato sull'elenco dipendenti -> import batch;
        Word/Excel/PDF altrove -> nuovi moduli (template)."""
        if self._worker_active:
            return
        widget = self.root.winfo_containing(x, y)
        card2 = getattr(self, "_card2_shadow", None)
        on_batch = widget is not None and card2 is not None and str(widget).startswith(str(card2))
        batch_files = [f for f in paths if f.suffix.lower() == ".csv"
                       or (on_batch and f.suffix.lower() == ".xlsx")]
        modules = [f for f in paths if f not in batch_files
                   and f.suffix.lower() in SUPPORTED_EXTENSIONS and f.is_file()]
        for batch_file in batch_files:
            self._inline_batch_import_file(batch_file)
        if modules:
            self.open_template_manager(initial_files=modules)
        if not batch_files and not modules:
            self._toast(self.tr("dd_unsupported"), "warning")

    def _add_tooltip(self, widget, getter):
        self._tooltips.append(Tooltip(widget, getter))

    # ---------------------- Multi reparto UI toggle ----------------------
    def _render_multi_dept_list(self, parent):
        for child in parent.winfo_children():
            child.destroy()
        self.multi_dept_values.clear()
        options = department_options(self.templates)
        if not options:
            tk.Label(parent, text="—", bg=self._style_colors["card_body_bg"],
                     fg=self._style_colors["text_muted"]).pack(anchor="w")
            return
        cols = min(3, max(1, (len(options) + 3) // 4))
        for i, dept in enumerate(options):
            var = BooleanVar(value=False)
            self.multi_dept_values[dept] = var
            cb = ttk.Checkbutton(parent, text=dept, variable=var,
                                 command=self.update_document_list)
            cb.grid(row=i // cols, column=i % cols, sticky="w", padx=(0, 10), pady=2)
            self._add_tooltip(cb, lambda d=dept: self.tr("tt_dept") + f" [{d}]")

    def _toggle_multi_dept(self):
        multi = self.multi_dept_mode.get()
        if multi:
            self._single_dept_wrap.grid_remove()
            self._multi_dept_wrap.grid()
            self._render_multi_dept_list(self._multi_dept_wrap)
            self.department.set("")
        else:
            self._multi_dept_wrap.grid_remove()
            self._single_dept_wrap.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        self.update_document_list()

    def _current_departments(self) -> list[str]:
        if self.multi_dept_mode.get():
            return [d for d, v in self.multi_dept_values.items() if v.get()]
        d = self.department.get().strip().upper()
        return [d] if d else []

    # ---------------------- Treeview inclusion logic ----------------------
    def _current_selected_templates(self) -> list[TemplateFile]:
        depts = self._current_departments()
        if not depts:
            return []
        if len(depts) == 1:
            chosen = templates_for_department(self.templates, depts[0])
        else:
            chosen = templates_for_departments(self.templates, depts)
        return [t for t in chosen if self.template_inclusion.get(t.path, True)]

    def _set_all_inclusion(self, value: bool):
        for path in list(self.template_inclusion.keys()):
            self.template_inclusion[path] = value
        self.update_document_list()

    def _on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item:
            return
        if region == "cell" and col == "#1":
            self._toggle_row(item)
            return
        if region == "heading" and col == "#1":
            current_all = all(
                self.template_inclusion.get(self._row_path.get(item), True)
                for item in self.tree.get_children()
            ) if self.tree.get_children() else True
            self._set_all_inclusion(not current_all)

    def _toggle_focused_row(self):
        item = self.tree.focus()
        if item:
            self._toggle_row(item)

    def _toggle_row(self, item):
        path = self._row_path.get(item)
        if path is None:
            return
        self.template_inclusion[path] = not self.template_inclusion.get(path, True)
        self.update_document_list()

    # ---------------------- Breakpoint responsive ------------------------
    def _on_root_resize(self, evt):
        if evt.widget is not self.root:
            return
        self._current_width = evt.width
        try:
            self._apply_breakpoint()
        except Exception:
            pass

    def _apply_breakpoint(self):
        want = "narrow" if self._current_width and self._current_width < 1100 else "wide"
        if want == self._current_breakpoint:
            return
        self._current_breakpoint = want
        layout = getattr(self, "_body_layout_root", None)
        if layout is None:
            return
        if want == "narrow":
            if self._card2_shadow is not None:
                self._card2_shadow.grid_configure(row=1, column=0, columnspan=2,
                                                   sticky="nsew", padx=(0, 0), pady=(0, 2))
            if self._card3_shadow is not None:
                self._card3_shadow.grid_configure(row=2, column=0, columnspan=2, sticky="nsew")
            layout.rowconfigure(1, weight=1)
            layout.rowconfigure(2, weight=1)
        else:
            if self._card2_shadow is not None:
                self._card2_shadow.grid_configure(row=1, column=0, columnspan=1,
                                                   sticky="nsew", padx=(0, 2), pady=(0, 0))
            if self._card3_shadow is not None:
                self._card3_shadow.grid_configure(row=1, column=1, columnspan=1, sticky="nsew")
            layout.rowconfigure(2, weight=0)

    # ---------------------- Rebuild on theme/lang change ------------------
    def _rebuild_ui(self):
        self._configure_style()
        self._apply_theme_root()
        self.root.title(self.tr("app_title"))
        saved_date = None
        try:
            if hasattr(self, "date_picker") and self.date_picker is not None:
                try:
                    if self.date_picker.winfo_exists():
                        saved_date = self.date_picker.get_date()
                except Exception:
                    saved_date = None
        except Exception:
            saved_date = None
        saved_batch: list[dict[str, str]] = []
        saved_batch_tree_rows: list[tuple] = []
        try:
            if hasattr(self, "_inline_batch_rows") and isinstance(self._inline_batch_rows, list):
                saved_batch = [dict(r) for r in self._inline_batch_rows]
            if hasattr(self, "_inline_batch_tree") and self._inline_batch_tree is not None:
                try:
                    tree = self._inline_batch_tree
                    if tree.winfo_exists():
                        for c in tree.get_children():
                            try:
                                saved_batch_tree_rows.append(tuple(tree.item(c, "values")))
                            except Exception:
                                continue
                except Exception:
                    pass
        except Exception:
            pass
        # Status/count/progress StringVar: reapply keys that are language-dependent
        self.status.set(self.tr("status_initial"))
        self.count_label.set(self.tr("count_none"))
        self.progress_label.set(self.tr("progress_ready"))
        for child in list(self.root.winfo_children()):
            try:
                child.destroy()
            except Exception:
                pass
        self._build_header()
        self._build_body()
        self.refresh_templates()
        try:
            if saved_date is not None and hasattr(self, "date_picker") and self.date_picker is not None:
                try:
                    if self.date_picker.winfo_exists():
                        self.date_picker.set_date(saved_date)
                        self.date_picker.configure_language(self.language)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            if saved_batch and hasattr(self, "_inline_batch_rows"):
                self._inline_batch_rows = [dict(r) for r in saved_batch]
                tree = getattr(self, "_inline_batch_tree", None)
                if tree is not None and tree.winfo_exists():
                    for c in tree.get_children():
                        tree.delete(c)
                    rows_to_use = (saved_batch_tree_rows
                                   if len(saved_batch_tree_rows) == len(saved_batch)
                                   else [(r.get("Nome", ""), r.get("Data", ""),
                                          r.get("Reparto", "")) for r in saved_batch])
                    for vals in rows_to_use:
                        tree.insert("", END, values=tuple(vals))
                self._refresh_inline_batch_count()
        except Exception:
            pass

    # ---------------------- Dir choices -----------------------------------
    def choose_template_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.template_dir.get(),
                                         title=self.tr("btn_choose_folder"))
        if chosen:
            self.template_dir.set(chosen)
            self._persist_settings()
            self.refresh_templates()

    def choose_output_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.output_dir.get(),
                                         title=self.tr("btn_choose_folder"))
        if chosen:
            self.output_dir.set(chosen)
            self._persist_settings()

    # ---------------------- Refresh templates -----------------------------
    def refresh_templates(self) -> None:
        folder = Path(self.template_dir.get()).expanduser()
        self.templates, self.ignored = discover_templates(folder)
        departments = department_options(self.templates)
        self.department_combo["values"] = departments
        if departments and self.department.get().upper() not in departments:
            self.department.set(departments[0])
        elif not departments:
            self.department.set("")
        if hasattr(self, "_multi_dept_wrap") and self._multi_dept_wrap is not None:
            self._render_multi_dept_list(self._multi_dept_wrap)
        for tpl in self.templates:
            if tpl.path not in self.template_inclusion:
                self.template_inclusion[tpl.path] = True
        stale = [p for p in list(self.template_inclusion.keys())
                 if not any(t.path == p for t in self.templates)]
        for s in stale:
            del self.template_inclusion[s]
        self.hash_status = classify_template_hashes(self.templates, self.saved_hashes)
        self.update_document_list()
        ok = sum(1 for v in self.hash_status.values() if v == "ok")
        mod = sum(1 for v in self.hash_status.values() if v == "modified")
        new = sum(1 for v in self.hash_status.values() if v == "new")
        self.hash_stat_label.set(self.tr("hash_status", ok=ok, mod=mod, new=new))
        if not folder.exists():
            self.status.set("La cartella template non esiste ancora.")
        elif not self.templates:
            self.status.set("Nessun template valido. Usa REPARTO_NUMERO_CODICE.")
        elif self.ignored:
            self.status.set(f"{len(self.templates)} template; {len(self.ignored)} ignorati.")
        else:
            self.status.set(f"{len(self.templates)} template pronti.")

    def update_document_list(self) -> None:
        if not hasattr(self, "tree") or self.tree is None:
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._row_path: dict[str, Path] = {}
        depts = self._current_departments()
        if len(depts) == 1:
            selected = templates_for_department(self.templates, depts[0])
        elif depts:
            selected = templates_for_departments(self.templates, depts)
        else:
            selected = []
        total_copies = sum(template.copies for template in selected
                           if self.template_inclusion.get(template.path, True))
        total_templates = sum(1 for template in selected
                              if self.template_inclusion.get(template.path, True))
        self.count_label.set(self.tr("count_label", total=total_copies, templates=total_templates))
        for idx, template in enumerate(selected):
            scope = "Tutti" if template.is_for_every_department else template.department.upper()
            included = self.template_inclusion.get(template.path, True)
            mark = "☑" if included else "☐"
            status_key = self.hash_status.get(str(template.path), "ok")
            if status_key == "modified":
                status_text = self.tr("status_modified")
            elif status_key == "new":
                status_text = self.tr("status_new")
            else:
                status_text = self.tr("status_ok")
            tags: list[str] = []
            if template.is_for_every_department:
                tags.append("tutti")
            elif idx % 2 == 0:
                tags.append("even")
            else:
                tags.append("odd")
            if status_key in {"modified", "new"}:
                tags.append(status_key)
            item = self.tree.insert(
                "", END,
                values=(mark, f"  {scope} · {template.path.name}",
                        template.copies, status_text),
                tags=tuple(tags),
            )
            self._row_path[item] = template.path

    # ---------------------- Generate single -------------------------------
    def _collect_form(self, require_name: bool = True):
        """Valida il modulo; ritorna (nome, data, reparti, template) oppure None."""
        name = self.employee_name.get().strip()
        entry_date = self.date_picker.get_string()
        departments = self._current_departments()
        if not name and require_name:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_name"))
            return None
        if not entry_date:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_date"))
            return None
        if not departments:
            messagebox.showwarning(self.tr("mb_missing_title"), self.tr("mb_missing_dept"))
            return None
        selected = self._current_selected_templates()
        if not selected:
            messagebox.showwarning(self.tr("mb_no_docs_title"), self.tr("mb_no_docs_body"))
            return None
        return name, entry_date, departments, selected

    def preview(self) -> None:
        """Dossier di prova in una cartella temporanea: una copia per modulo e filigrana."""
        if self._worker_active:
            return
        form = self._collect_form(require_name=False)
        if form is None:
            return
        name, entry_date, departments, selected = form
        name = name or self.tr("pv_sample_name")
        single_copies = [dc_replace(t, copies=1) for t in selected]
        output_path = Path(tempfile.gettempdir()) / (
            f"{PREVIEW_PREFIX}{os.getpid()}_{datetime.now():%H%M%S%f}.pdf")
        watermark = self.tr("pv_watermark")
        self._worker_active = True
        self.status.set(self.tr("pv_status_running"))
        self._post("progress_ready_label", None)

        def progress_cb(step: int, total: int):
            self._post("progress", {"step": step, "total": total})

        def work():
            try:
                build_pdf(output_path, name, entry_date, "+".join(departments), "", "",
                          single_copies, progress_cb=progress_cb, watermark=watermark)
                self._post("done_preview", {"ok": True, "path": str(output_path)})
            except Exception as error:  # noqa: BLE001
                traceback.print_exc()
                self._post("done_preview", {"ok": False, "error": str(error)})
            finally:
                self._post("worker_done", None)

        threading.Thread(target=work, daemon=True).start()

    def generate(self) -> None:
        if self._worker_active:
            return
        form = self._collect_form()
        if form is None:
            return
        name, entry_date, departments, selected = form
        output_dir_path = Path(self.output_dir.get()).expanduser()
        dept_tag = safe_file_part("+".join(departments))
        base = output_dir_path / f"dossier_{safe_file_part(name)}_{dept_tag}.pdf"
        output_path = base
        counter = 2
        while output_path.exists():
            output_path = output_dir_path / (
                f"dossier_{safe_file_part(name)}_{dept_tag}_{counter}.pdf"
            )
            counter += 1
        dept_str = "+".join(departments)
        self._worker_active = True
        self.status.set(self.tr("mb_progress_status"))
        self._post("progress_ready_label", None)
        self._persist_settings()
        # Ruolo e note disabilitati: valori statici vuoti (rimossi da UI in v2)
        ruolo = ""
        note = ""
        auto_open_value = bool(self.auto_open.get())
        saved_snapshot = dict(self.saved_hashes) if self.saved_hashes else {}
        tr_done = self.tr("mb_status_done", name=output_path.name)
        tr_error = self.tr("mb_status_error")
        pdf_options = self._pdf_options()

        def progress_cb(step: int, total: int):
            self._post("progress", {"step": step, "total": total})

        def work():
            try:
                total = build_pdf(output_path, name, entry_date, dept_str,
                                  ruolo, note, selected,
                                  progress_cb=progress_cb, **pdf_options)
                self._save_history(output_path, name, dept_str, total, entry_date)
                hashes = dict(saved_snapshot)
                for t in selected:
                    try:
                        hashes[str(t.path)] = compute_template_hash(t.path)
                    except Exception:
                        pass
                save_hashes(hashes)
                self._post("status", tr_done)
                self._post("done_single", {
                    "ok": True, "path": str(output_path), "total": total,
                    "auto_open": auto_open_value,
                })
            except Exception as error:  # noqa: BLE001
                traceback.print_exc()
                self._post("status", tr_error)
                self._post("done_single", {"ok": False, "error": str(error)})
            finally:
                self._post("worker_done", None)

        threading.Thread(target=work, daemon=True).start()

    # ---------------------- Worker -> mainloop routing --------------------
    def _drain_queue(self):
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "progress":
                    self._set_progress(payload["step"], payload["total"])
                elif kind == "status":
                    self.status.set(str(payload))
                elif kind == "progress_ready_label":
                    self.progress_label.set(self.tr("progress_ready"))
                    if self._progressbar is not None:
                        self._progressbar["value"] = 0
                elif kind == "done_single":
                    if payload.get("ok"):
                        pdf = Path(str(payload.get("path", "")))
                        message = self.tr("ts_done_single", name=pdf.name,
                                          total=payload.get("total", 0))
                        self._toast(message, "success", (self.tr("ts_open_pdf"),
                                                         lambda: open_folder(pdf)), 9000)
                        notify_windows(self.root, self.tr("mb_done_title"), message)
                        if payload.get("auto_open"):
                            open_folder(Path(str(payload["path"])).parent)
                        self.refresh_templates()
                    else:
                        messagebox.showerror(self.tr("mb_error_title"), str(payload.get("error", "")))
                elif kind == "done_preview":
                    if payload.get("ok"):
                        self.status.set(self.tr("pv_status_done"))
                        open_folder(Path(str(payload["path"])))
                    else:
                        self.status.set(self.tr("mb_status_error"))
                        messagebox.showerror(self.tr("mb_error_title"), str(payload.get("error", "")))
                elif kind == "update_result":
                    self._handle_update_result(payload)
                elif kind == "done_batch":
                    counts = (
                        f"{payload.get('ok', 0)} " + self.tr("bat_summary_ok") +
                        "   ·   " + f"{payload.get('skip', 0)} " + self.tr("bat_summary_skip") +
                        "   ·   " + f"{payload.get('fail', 0)} " + self.tr("bat_summary_fail")
                    )
                    summary = counts + "\n\n" + str(payload.get("detail", ""))
                    if payload.get("auto_open") and payload.get("out_dir"):
                        open_folder(Path(str(payload["out_dir"])))
                    problems = payload.get("skip", 0) or payload.get("fail", 0)
                    self._toast(
                        self.tr("bat_summary_title") + "  ·  " + counts,
                        "warning" if problems else "success",
                        (self.tr("ts_details"),
                         lambda text=summary: messagebox.showinfo(self.tr("bat_summary_title"), text)),
                        12000,
                    )
                    notify_windows(self.root, self.tr("bat_summary_title"), counts)
                    self.refresh_templates()
                elif kind == "worker_done":
                    self._worker_active = False
                    self.saved_hashes = load_saved_hashes()
                    self.hash_status = classify_template_hashes(self.templates, self.saved_hashes)
                    ok = sum(1 for v in self.hash_status.values() if v == "ok")
                    mod = sum(1 for v in self.hash_status.values() if v == "modified")
                    new = sum(1 for v in self.hash_status.values() if v == "new")
                    self.hash_stat_label.set(self.tr("hash_status", ok=ok, mod=mod, new=new))
        except queue.Empty:
            pass
        finally:
            self.root.after(80, self._drain_queue)

    def _install_drain_loop(self):
        self.root.after(80, self._drain_queue)

    # ---------------------- Department Editor -----------------------------
    def open_department_editor(self):
        win = tk.Toplevel(self.root)
        win.title(self.tr("de_title"))
        win.transient(self.root)
        win.grab_set()
        win.geometry("520x480")
        win.configure(bg=self._style_colors["app_bg"])
        try:
            _set_app_icon(win)
        except Exception:
            pass
        body = tk.Frame(win, bg=self._style_colors["card_body_bg"], padx=18, pady=18)
        body.pack(fill=BOTH, expand=True, padx=14, pady=14)
        ttk.Label(body, text=self.tr("de_hint"), style="Muted.TLabel").pack(anchor="w", pady=(0, 12))
        list_wrap = tk.Frame(body, bg=self._style_colors["card_body_bg"])
        list_wrap.pack(fill=BOTH, expand=True, pady=(0, 12))
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)
        lb = tk.Listbox(list_wrap, activestyle="dotbox",
                        bg=self._style_colors["field_bg"], fg=self._style_colors["text"],
                        font=("Segoe UI", 10), selectmode="single",
                        highlightthickness=2,
                        highlightbackground=self._style_colors["border"],
                        highlightcolor=self._style_colors["focus"],
                        relief="flat", borderwidth=0)
        lb.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=lb.yview)
        sb.grid(row=0, column=1, sticky="ns")
        lb.configure(yscrollcommand=sb.set)
        current = load_departments_from_file()
        for item in current:
            lb.insert(END, item)
        row = tk.Frame(body, bg=self._style_colors["card_body_bg"])
        row.pack(fill=X, pady=(0, 12))
        tk.Label(row, text=self.tr("de_entry"), bg=self._style_colors["card_body_bg"],
                 fg=self._style_colors["text_muted"],
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 4))
        entry = ttk.Entry(row)
        entry.pack(fill=X, pady=(0, 10))
        btns = tk.Frame(row, bg=self._style_colors["card_body_bg"])
        btns.pack(fill=X)

        def add():
            val = entry.get().strip().upper()
            if not val:
                messagebox.showwarning(self.tr("de_err_empty_title"), self.tr("de_err_empty_body"))
                return
            if val in lb.get(0, END):
                messagebox.showwarning(self.tr("de_err_exists_title"), self.tr("de_err_exists_body"))
                return
            lb.insert(END, val)
            entry.delete(0, END)

        def rename():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning(self.tr("de_err_none_title"), self.tr("de_err_none_body"))
                return
            new_val = entry.get().strip().upper()
            if not new_val:
                messagebox.showwarning(self.tr("de_err_empty_title"), self.tr("de_err_empty_body"))
                return
            if new_val in lb.get(0, END):
                messagebox.showwarning(self.tr("de_err_exists_title"), self.tr("de_err_exists_body"))
                return
            lb.delete(sel[0])
            lb.insert(sel[0], new_val)

        def delete():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning(self.tr("de_err_none_title"), self.tr("de_err_none_body"))
                return
            lb.delete(sel[0])

        self._button(btns, "de_add", style="Secondary.TButton", command=add
                   ).pack(side=LEFT, padx=(0, 8))
        self._button(btns, "de_rename", style="Secondary.TButton", command=rename
                   ).pack(side=LEFT, padx=(0, 8))
        self._button(btns, "de_delete", style="Secondary.TButton", command=delete
                   ).pack(side=LEFT)

        footer = tk.Frame(body, bg=self._style_colors["card_body_bg"])
        footer.pack(fill=X, side="bottom")

        def save_and_close():
            new_list = list(dict.fromkeys(lb.get(0, END)))  # preserve order, unique
            try:
                DEPARTMENTS_FILE.parent.mkdir(parents=True, exist_ok=True)
                content = "\n".join(new_list) + ("\n" if new_list else "")
                DEPARTMENTS_FILE.write_text(content, encoding="utf-8")
            except OSError as exc:
                messagebox.showerror(self.tr("mb_error_title"), str(exc))
                return
            clear_caches()
            self.refresh_templates()
            win.destroy()

        self._button(footer, "de_save", style="Primary.TButton",
                   command=save_and_close).pack(side=RIGHT)
        self._button(footer, "de_cancel", style="Secondary.TButton",
                   command=win.destroy).pack(side=RIGHT, padx=(0, 10))

    # ---------------------- Gestione moduli ------------------------------
    def _template_form(self, parent, title: str, source_name: str,
                       initial: dict[str, Any]) -> dict[str, Any] | None:
        """Finestra modale per reparto / copie / codice. Ritorna i valori o None."""
        colors = self._style_colors
        dlg = tk.Toplevel(parent)
        dlg.title(title)
        dlg.transient(parent)
        dlg.resizable(False, False)
        dlg.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(dlg)
        except Exception:
            pass
        body = tk.Frame(dlg, bg=colors["card_body_bg"], padx=20, pady=18)
        body.pack(fill=BOTH, expand=True, padx=12, pady=12)
        body.columnconfigure(1, weight=1)

        def label(text, row):
            tk.Label(body, text=text, bg=colors["card_body_bg"], fg=colors["text"],
                     font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                     ).grid(row=row, column=0, sticky="w", padx=(0, 14), pady=(0, 10))

        tk.Label(body, text=self.tr("tm_form_file", name=source_name),
                 bg=colors["card_body_bg"], fg=colors["text_muted"],
                 font=("Segoe UI", 9), anchor="w", wraplength=420, justify="left"
                 ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))

        suffix = Path(source_name).suffix
        dept_var = StringVar(value=initial.get("department", ""))
        copies_var = StringVar(value=str(initial.get("copies", 1)))
        code_var = StringVar(value=initial.get("code", ""))
        code_touched = {"value": bool(initial.get("code"))}

        label(self.tr("tm_col_dept"), 1)
        dept_values = ["TUTTI"] + [d for d in department_options(self.templates) if d != "TUTTI"]
        dept_cb = ttk.Combobox(body, textvariable=dept_var, values=dept_values, width=32)
        dept_cb.grid(row=1, column=1, sticky="ew", pady=(0, 10))

        label(self.tr("tm_col_copies"), 2)
        copies_sp = ttk.Spinbox(body, from_=1, to=50, textvariable=copies_var, width=6)
        copies_sp.grid(row=2, column=1, sticky="w", pady=(0, 10))

        label(self.tr("tm_col_code"), 3)
        code_ent = ttk.Entry(body, textvariable=code_var, width=10)
        code_ent.grid(row=3, column=1, sticky="w", pady=(0, 10))
        code_ent.bind("<Key>", lambda _e: code_touched.__setitem__("value", True))

        tk.Label(body, text=self.tr("tm_form_hint"), bg=colors["card_body_bg"],
                 fg=colors["text_muted"], font=("Segoe UI", 8), anchor="w",
                 wraplength=420, justify="left"
                 ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 10))

        preview_var = StringVar()
        tk.Label(body, textvariable=preview_var, bg=colors["card_body_bg"],
                 fg=colors["text"], font=("Consolas", 10, "bold"), anchor="w"
                 ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 14))

        def refresh_preview(*_):
            if not code_touched["value"]:
                code_var.set(suggest_template_code(dept_var.get()))
            err = validate_template_fields(dept_var.get(), copies_var.get(), code_var.get())
            if err:
                preview_var.set("⚠ " + self.tr(err))
            else:
                preview_var.set(self.tr("tm_form_preview", name=build_template_filename(
                    dept_var.get(), int(copies_var.get()), code_var.get(), suffix)))

        dept_var.trace_add("write", refresh_preview)
        copies_var.trace_add("write", refresh_preview)
        code_var.trace_add("write", refresh_preview)
        refresh_preview()

        result: dict[str, Any] = {}

        def confirm(_evt=None):
            err = validate_template_fields(dept_var.get(), copies_var.get(), code_var.get())
            if err:
                messagebox.showwarning(self.tr("de_err_empty_title"), self.tr(err), parent=dlg)
                return
            result.update(department=dept_var.get().strip().upper(),
                          copies=int(copies_var.get()),
                          code=code_var.get().strip().upper())
            dlg.destroy()

        btns = tk.Frame(body, bg=colors["card_body_bg"])
        btns.grid(row=6, column=0, columnspan=2, sticky="e")
        self._button(btns, "de_cancel", style="Secondary.TButton",
                   command=dlg.destroy).pack(side=LEFT, padx=(0, 10))
        self._button(btns, "tm_form_ok", style="Primary.TButton",
                   command=confirm).pack(side=LEFT)
        dlg.bind("<Return>", confirm)
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        dept_cb.focus_set()
        dlg.grab_set()
        parent.wait_window(dlg)
        try:
            parent.grab_set()
        except tk.TclError:
            pass
        return result or None

    def open_template_manager(self, initial_files: list[Path] | None = None):
        colors = self._style_colors
        win = tk.Toplevel(self.root)
        win.title(self.tr("tm_title"))
        win.transient(self.root)
        win.grab_set()
        win.geometry("900x560")
        win.minsize(700, 420)
        win.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(win)
        except Exception:
            pass
        body = tk.Frame(win, bg=colors["card_body_bg"], padx=18, pady=18)
        body.pack(fill=BOTH, expand=True, padx=14, pady=14)
        ttk.Label(body, text=self.tr("tm_hint"), style="Muted.TLabel",
                  wraplength=820, justify="left").pack(anchor="w", pady=(0, 12))

        toolbar = tk.Frame(body, bg=colors["card_body_bg"])
        toolbar.pack(fill=X, side="bottom")

        list_wrap = tk.Frame(body, bg=colors["card_body_bg"])
        list_wrap.pack(fill=BOTH, expand=True, pady=(0, 12))
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)
        columns = ("file", "dept", "copies", "code", "status")
        tree = ttk.Treeview(list_wrap, columns=columns, show="headings", selectmode="extended")
        widths = {"file": 280, "dept": 170, "copies": 70, "code": 80, "status": 220}
        for col in columns:
            tree.heading(col, text=self.tr(f"tm_col_{col}"))
            tree.column(col, width=widths[col],
                        anchor="center" if col in ("copies", "code") else "w")
        tree.tag_configure("invalid", foreground="#c0392b" if self.theme.get() == "light" else "#ff8a7a")
        tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=sb.set)
        rows: dict[str, Path] = {}

        def folder() -> Path:
            return Path(self.template_dir.get()).expanduser()

        def reload():
            self.refresh_templates()
            tree.delete(*tree.get_children())
            rows.clear()
            ordered = sorted(self.templates, key=lambda t: (not t.is_for_every_department,
                                                            t.department.upper(),
                                                            t.path.name.lower()))
            for tpl in ordered:
                dept = (self.tr("tm_all_depts") if tpl.is_for_every_department
                        else tpl.department.upper())
                iid = tree.insert("", END, values=(tpl.display_name, dept, tpl.copies, tpl.code,
                                                   "✓ " + self.tr("tm_status_ok")))
                rows[iid] = tpl.path
            for path in self.ignored:
                iid = tree.insert("", END, values=(path.name, "—", "—", "—",
                                                   "⚠ " + self.tr("tm_status_invalid")),
                                  tags=("invalid",))
                rows[iid] = path

        def selected_paths() -> list[Path]:
            return [rows[i] for i in tree.selection() if i in rows]

        def require_one() -> Path | None:
            paths = selected_paths()
            if len(paths) != 1:
                messagebox.showwarning(self.tr("de_err_none_title"),
                                       self.tr("tm_err_select_one"), parent=win)
                return None
            return paths[0]

        def initial_for(path: Path) -> dict[str, Any]:
            tpl = parse_template(path)
            if tpl:
                return {"department": tpl.department.upper(), "copies": tpl.copies,
                        "code": tpl.code}
            return {"department": self.department.get() or "", "copies": 1, "code": ""}

        def add_files():
            sources = filedialog.askopenfilenames(
                parent=win, title=self.tr("tm_add"),
                filetypes=[(self.tr("tm_filetypes"), "*.docx *.doc *.xlsx *.xls *.pdf")])
            add_paths(sources)

        def add_paths(sources):
            if not sources:
                return
            dest_dir = folder()
            try:
                dest_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                messagebox.showerror(self.tr("tm_title"), str(exc), parent=win)
                return
            added = 0
            for src in map(Path, sources):
                if src.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue
                initial = initial_for(src)
                while True:
                    values = self._template_form(win, self.tr("tm_form_title_add"),
                                                 src.name, initial)
                    if values is None:
                        break
                    target = dest_dir / build_template_filename(
                        values["department"], values["copies"], values["code"], src.suffix)
                    if target.exists():
                        messagebox.showwarning(self.tr("de_err_exists_title"),
                                               self.tr("tm_err_exists", name=target.name),
                                               parent=win)
                        initial = values
                        continue
                    try:
                        shutil.copy2(src, target)
                        added += 1
                    except OSError as exc:
                        messagebox.showerror(self.tr("tm_title"), str(exc), parent=win)
                    break
            if added:
                reload()
                self.status.set(self.tr("tm_added", n=added))

        def edit_selected(_evt=None):
            path = require_one()
            if path is None:
                return
            initial = initial_for(path)
            while True:
                values = self._template_form(win, self.tr("tm_form_title_edit"),
                                             path.name, initial)
                if values is None:
                    return
                target = path.with_name(build_template_filename(
                    values["department"], values["copies"], values["code"], path.suffix))
                if target == path:
                    return
                if target.exists():
                    messagebox.showwarning(self.tr("de_err_exists_title"),
                                           self.tr("tm_err_exists", name=target.name),
                                           parent=win)
                    initial = values
                    continue
                try:
                    path.rename(target)
                except OSError as exc:
                    messagebox.showerror(self.tr("tm_title"), str(exc), parent=win)
                    return
                if str(path) in self.saved_hashes:
                    self.saved_hashes[str(target)] = self.saved_hashes.pop(str(path))
                    save_hashes(self.saved_hashes)
                if path in self.template_inclusion:
                    self.template_inclusion[target] = self.template_inclusion.pop(path)
                reload()
                self.status.set(self.tr("tm_renamed", name=target.name))
                return

        def delete_selected():
            paths = selected_paths()
            if not paths:
                messagebox.showwarning(self.tr("de_err_none_title"),
                                       self.tr("tm_err_select_any"), parent=win)
                return
            names = "\n".join("• " + p.name for p in paths[:10])
            if len(paths) > 10:
                names += "\n…"
            if not messagebox.askyesno(self.tr("tm_delete"),
                                       self.tr("tm_confirm_delete", n=len(paths), names=names,
                                               trash=TRASH_DIR_NAME), parent=win):
                return
            for path in paths:
                try:
                    move_template_to_trash(path, folder())
                except OSError as exc:
                    messagebox.showerror(self.tr("tm_title"), str(exc), parent=win)
            reload()
            self.status.set(self.tr("tm_deleted", n=len(paths)))

        def open_selected():
            path = require_one()
            if path is not None:
                open_folder(path)

        for text_key, cmd, style in (
            ("tm_add", add_files, "Accent.TButton"),
            ("tm_edit", edit_selected, "Secondary.TButton"),
            ("tm_open", open_selected, "Secondary.TButton"),
            ("tm_delete", delete_selected, "Secondary.TButton"),
        ):
            self._button(toolbar, text_key, style=style, command=cmd
                       ).pack(side=LEFT, padx=(0, 8))
        self._button(toolbar, "tm_close", style="Primary.TButton",
                   command=win.destroy).pack(side=RIGHT)
        self._button(toolbar, "tm_open_folder", style="Secondary.TButton",
                   command=lambda: open_folder(folder())).pack(side=RIGHT, padx=(0, 8))

        tree.bind("<Double-1>", edit_selected)
        tree.bind("<Return>", edit_selected)
        tree.bind("<Delete>", lambda _e: delete_selected())
        reload()
        if os.name == "nt":
            try:
                win._drop_target = FileDropTarget(win, lambda paths, _x, _y: add_paths(paths))
            except Exception:  # noqa: BLE001
                traceback.print_exc()
        if initial_files:
            win.after(150, lambda: add_paths(initial_files))

    def _parse_date_str(self, s: str):
        if not s:
            return None
        if isinstance(s, datetime):
            return s.date()
        if isinstance(s, date):
            return s
        s_norm = str(s).strip().replace("-", "/").replace(".", "/")
        for fmt in ("%d/%m/%Y", "%Y/%m/%d", "%d/%m/%y"):
            try:
                return datetime.strptime(s_norm, fmt).date()
            except ValueError:
                continue
        return None

    def _normalize_column_header(self, c: str) -> str:
        c_norm = str(c).strip().lower().replace("  ", " ")
        if any(k in c_norm for k in ("nome", "cognome", "name", "full name", "dipendente", "employee", "persona")):
            return "Nome"
        if any(k in c_norm for k in ("data", "date", "ingresso", "entrata", "entry")):
            return "Data"
        if any(k in c_norm for k in ("reparto", "department", "dipartimento", "settore", "area", "ufficio")):
            return "Reparto"
        if any(k in c_norm for k in ("ruolo", "role", "mansione", "qualifica", "position", "job")):
            return "Ruolo"
        if any(k in c_norm for k in ("note", "notes", "commento", "commenti", "comment", "osservazioni")):
            return "Note"
        return c

    def _coerce_date_value(self, val) -> str:
        if val is None:
            return ""
        if isinstance(val, datetime):
            return val.strftime("%d/%m/%Y")
        if isinstance(val, date):
            return val.strftime("%d/%m/%Y")
        s = str(val).strip()
        if not s:
            return ""
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S",
                    "%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        return s

    def _parse_batch_file(self, p: Path) -> list[dict[str, str]]:
        suffix = p.suffix.lower()
        if suffix == ".xls":
            raise ValueError(self.tr("bat_xls_unsupported"))
        if suffix == ".xlsx":
            wb = load_workbook(p, data_only=True, read_only=True)
            try:
                ws = wb.worksheets[0]
                rows_iter = ws.iter_rows(values_only=True)
                raw_rows = [list(r) for r in rows_iter]
            finally:
                wb.close()
            if not raw_rows:
                return []
            header = [str(c).strip().lower() if c is not None else "" for c in raw_rows[0]]
            norm = [self._normalize_column_header(c) for c in header]
            out: list[dict[str, str]] = []
            for r in raw_rows[1:]:
                if all(v is None or str(v).strip() == "" for v in r):
                    continue
                obj: dict[str, str] = {}
                for key, val in zip(norm, r):
                    if key == "Data":
                        obj[key] = self._coerce_date_value(val)
                    else:
                        obj[key] = "" if val is None else str(val).strip()
                out.append(obj)
            return out
        # CSV
        text = p.read_text(encoding="utf-8-sig", errors="replace")
        reader = csv.reader(text.splitlines())
        rows_list = list(reader)
        if not rows_list:
            return []
        header = [c.strip().lower() for c in rows_list[0]]
        norm = [self._normalize_column_header(c) for c in header]
        out = []
        for r in rows_list[1:]:
            if all(v.strip() == "" for v in r):
                continue
            while len(r) < len(norm):
                r.append("")
            obj = {}
            for k, v in zip(norm, r):
                if k == "Data":
                    obj[k] = self._coerce_date_value(v)
                else:
                    obj[k] = v.strip()
            out.append(obj)
        return out

    # ---------------------- History ---------------------------------------
    @staticmethod
    def _load_history() -> list[dict[str, Any]]:
        if not HISTORY_FILE.exists():
            return []
        try:
            history = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return []
        return [h for h in history if isinstance(h, dict)] if isinstance(history, list) else []

    def _write_history(self, history: list[dict[str, Any]]) -> None:
        try:
            HISTORY_FILE.write_text(
                json.dumps(history[:HISTORY_LIMIT], ensure_ascii=False, indent=2),
                encoding="utf-8")
        except OSError:
            # La cronologia e' accessoria: un errore qui non deve far fallire il dossier
            traceback.print_exc()

    def _save_history(self, output_path: Path, name: str, department: str, total: int,
                      entry_date: str = "") -> None:
        with self._history_lock:
            history = self._load_history()
            history.insert(0, {
                "path": str(output_path),
                "name": name,
                "entry_date": entry_date,
                "department": department,
                "documents": total,
                "ts": datetime.now().isoformat(timespec="seconds"),
            })
            self._write_history(history)

    def open_history(self):
        colors = self._style_colors
        win = tk.Toplevel(self.root)
        win.title(self.tr("hi_title"))
        win.transient(self.root)
        win.geometry("980x560")
        win.minsize(720, 400)
        win.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(win)
        except Exception:
            pass
        body = tk.Frame(win, bg=colors["card_body_bg"], padx=18, pady=18)
        body.pack(fill=BOTH, expand=True, padx=14, pady=14)

        top = tk.Frame(body, bg=colors["card_body_bg"])
        top.pack(fill=X, pady=(0, 12))
        tk.Label(top, text=self.tr("hi_search"), bg=colors["card_body_bg"], fg=colors["text"],
                 font=("Segoe UI Semibold", 9, "bold")).pack(side=LEFT, padx=(0, 10))
        query = StringVar()
        search = ttk.Entry(top, textvariable=query)
        search.pack(side=LEFT, fill=X, expand=True)
        count_var = StringVar()
        ttk.Label(top, textvariable=count_var, style="Count.TLabel").pack(side=RIGHT, padx=(12, 0))

        toolbar = tk.Frame(body, bg=colors["card_body_bg"])
        toolbar.pack(fill=X, side="bottom")

        list_wrap = tk.Frame(body, bg=colors["card_body_bg"])
        list_wrap.pack(fill=BOTH, expand=True, pady=(0, 12))
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)
        columns = ("ts", "name", "entry", "dept", "docs", "file")
        tree = ttk.Treeview(list_wrap, columns=columns, show="headings", selectmode="browse")
        widths = {"ts": 130, "name": 200, "entry": 100, "dept": 170, "docs": 60, "file": 280}
        for col in columns:
            tree.heading(col, text=self.tr(f"hi_col_{col}"))
            tree.column(col, width=widths[col], anchor="center" if col == "docs" else "w",
                        stretch=col in ("name", "file"))
        tree.tag_configure("missing", foreground=colors["text_muted"])
        tree.tag_configure("even", background=colors["row_even"])
        tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=sb.set)
        rows: dict[str, dict[str, Any]] = {}

        def fmt_ts(value: str) -> str:
            try:
                return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
            except (TypeError, ValueError):
                return str(value or "")

        def reload(*_):
            tree.delete(*tree.get_children())
            rows.clear()
            needle = query.get().strip().lower()
            history = self._load_history()
            shown = 0
            for entry in history:
                haystack = " ".join(str(entry.get(k, "")) for k in
                                    ("name", "department", "entry_date", "path")).lower()
                if needle and needle not in haystack:
                    continue
                path = Path(str(entry.get("path", "")))
                exists = path.exists()
                tags = [] if exists else ["missing"]
                if shown % 2:
                    tags.append("even")
                iid = tree.insert("", END, tags=tuple(tags), values=(
                    fmt_ts(entry.get("ts", "")), entry.get("name", ""),
                    entry.get("entry_date", "") or "—", entry.get("department", ""),
                    entry.get("documents", ""),
                    path.name if exists else f"⚠ {path.name} · {self.tr('hi_missing')}",
                ))
                rows[iid] = entry
                shown += 1
            count_var.set(self.tr("hi_count", shown=shown, total=len(history)))

        def current() -> dict[str, Any] | None:
            sel = tree.selection()
            if not sel or sel[0] not in rows:
                messagebox.showwarning(self.tr("de_err_none_title"),
                                       self.tr("hi_select"), parent=win)
                return None
            return rows[sel[0]]

        def open_pdf(_evt=None):
            entry = current()
            if entry is None:
                return
            path = Path(str(entry.get("path", "")))
            if not path.exists():
                messagebox.showwarning(self.tr("hi_title"), self.tr("hi_missing_body"), parent=win)
                return
            open_folder(path)

        def open_dir():
            entry = current()
            if entry is None:
                return
            folder = Path(str(entry.get("path", ""))).parent
            open_folder(folder if folder.exists() else Path(self.output_dir.get()))

        def reuse():
            entry = current()
            if entry is None:
                return
            self._apply_history_entry(entry)
            win.destroy()

        def remove():
            entry = current()
            if entry is None:
                return
            with self._history_lock:
                history = [h for h in self._load_history() if h != entry]
                self._write_history(history)
            reload()

        for text_key, cmd, style in (
            ("hi_open_pdf", open_pdf, "Accent.TButton"),
            ("hi_open_dir", open_dir, "Secondary.TButton"),
            ("hi_reuse", reuse, "Secondary.TButton"),
            ("hi_remove", remove, "Secondary.TButton"),
        ):
            self._button(toolbar, text_key, style=style, command=cmd
                       ).pack(side=LEFT, padx=(0, 8))
        self._button(toolbar, "tm_close", style="Primary.TButton",
                   command=win.destroy).pack(side=RIGHT)

        query.trace_add("write", reload)
        tree.bind("<Double-1>", open_pdf)
        tree.bind("<Return>", open_pdf)
        tree.bind("<Delete>", lambda _e: remove())
        win.bind("<Escape>", lambda _e: win.destroy())
        reload()
        search.focus_set()

    def _apply_history_entry(self, entry: dict[str, Any]) -> None:
        """Ricompila il modulo con i dati di un dossier dello storico."""
        self.employee_name.set(str(entry.get("name", "")))
        parsed = self._parse_date_str(str(entry.get("entry_date", "")))
        if parsed is not None:
            self.date_picker.set_date(parsed)
        options = department_options(self.templates)
        wanted = [d.strip().upper() for d in str(entry.get("department", "")).split("+")
                  if d.strip()]
        known = [d for d in wanted if d in options]
        if len(known) > 1:
            self.multi_dept_mode.set(True)
            self._toggle_multi_dept()
            for dept in known:
                if dept in self.multi_dept_values:
                    self.multi_dept_values[dept].set(True)
        elif known:
            if self.multi_dept_mode.get():
                self.multi_dept_mode.set(False)
                self._toggle_multi_dept()
            self.department.set(known[0])
        self.update_document_list()
        if len(known) != len(wanted):
            missing = ", ".join(d for d in wanted if d not in known)
            messagebox.showwarning(self.tr("hi_title"), self.tr("hi_dept_missing", d=missing))
        self.status.set(self.tr("hi_reused", name=entry.get("name", "")))
        self._toast(self.tr("hi_reused", name=entry.get("name", "")))

    # ---------------------- Modello Excel batch ---------------------------
    def download_batch_template(self):
        target = filedialog.asksaveasfilename(
            title=self.tr("bat_template_download"),
            initialdir=self.output_dir.get(),
            initialfile=self.tr("bat_template_filename"),
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not target:
            return
        departments = [d for d in department_options(self.templates)
                       if d.upper() not in ALL_DEPARTMENT_NAMES]
        try:
            write_batch_template(Path(target), departments,
                                 italian=self.language_code == "it")
        except OSError as exc:
            messagebox.showerror(self.tr("mb_error_title"), str(exc))
            return
        self.status.set(self.tr("bat_template_saved", name=Path(target).name))
        self._toast(self.tr("bat_template_saved", name=Path(target).name), "success",
                    (self.tr("ts_open"), lambda: open_folder(Path(target))))

    # ---------------------- Impostazioni ----------------------------------
    def open_settings(self):
        colors = self._style_colors
        win = tk.Toplevel(self.root)
        win.title(self.tr("st_title"))
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        win.configure(bg=colors["app_bg"])
        try:
            _set_app_icon(win)
        except Exception:
            pass
        body = tk.Frame(win, bg=colors["card_body_bg"], padx=22, pady=20)
        body.pack(fill=BOTH, expand=True, padx=14, pady=14)
        body.columnconfigure(0, weight=1)

        def section(text, row):
            tk.Label(body, text=text, bg=colors["card_body_bg"], fg=colors["accent_label_fg"],
                     font=("Segoe UI", 9, "bold"), anchor="w"
                     ).grid(row=row, column=0, sticky="w", pady=(0, 8))

        def hint(text, row):
            tk.Label(body, text=text, bg=colors["card_body_bg"], fg=colors["text_muted"],
                     font=("Segoe UI", 8), anchor="w", justify="left", wraplength=460
                     ).grid(row=row, column=0, sticky="w", pady=(0, 14))

        watermark_var = StringVar(value=str(self.settings.get("pdf_watermark") or ""))
        protect_var = BooleanVar(value=bool(self.settings.get("pdf_protect")))
        updates_var = BooleanVar(value=bool(self.settings.get("check_updates")))
        source_var = StringVar(value=str(self.settings.get("update_source") or ""))

        section(self.tr("st_pdf_section"), 0)
        tk.Label(body, text=self.tr("st_watermark"), bg=colors["card_body_bg"], fg=colors["text"],
                 font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                 ).grid(row=1, column=0, sticky="w", pady=(0, 4))
        wm_entry = ttk.Entry(body, textvariable=watermark_var, width=52)
        wm_entry.grid(row=2, column=0, sticky="ew", pady=(0, 4))
        hint(self.tr("st_watermark_hint"), 3)
        ttk.Checkbutton(body, text=self.tr("st_protect"), variable=protect_var
                        ).grid(row=4, column=0, sticky="w", pady=(0, 4))
        hint(self.tr("st_protect_hint"), 5)

        section(self.tr("st_update_section"), 6)
        ttk.Checkbutton(body, text=self.tr("st_check_updates"), variable=updates_var
                        ).grid(row=7, column=0, sticky="w", pady=(0, 8))
        tk.Label(body, text=self.tr("st_update_source"), bg=colors["card_body_bg"],
                 fg=colors["text"], font=("Segoe UI Semibold", 9, "bold"), anchor="w"
                 ).grid(row=8, column=0, sticky="w", pady=(0, 4))
        ttk.Entry(body, textvariable=source_var, width=52
                  ).grid(row=9, column=0, sticky="ew", pady=(0, 4))
        hint(self.tr("st_update_hint", version=APP_VERSION), 10)

        def store():
            self.settings["pdf_watermark"] = watermark_var.get().strip()
            self.settings["pdf_protect"] = bool(protect_var.get())
            self.settings["check_updates"] = bool(updates_var.get())
            self.settings["update_source"] = source_var.get().strip()
            self._persist_settings()

        def check_now():
            store()
            self.check_for_updates(manual=True)

        def save_and_close(_evt=None):
            store()
            win.destroy()

        btns = tk.Frame(body, bg=colors["card_body_bg"])
        btns.grid(row=11, column=0, sticky="ew", pady=(6, 0))
        self._button(btns, "st_check_now", style="Secondary.TButton",
                   command=check_now).pack(side=LEFT)
        self._button(btns, "de_save", style="Primary.TButton",
                   command=save_and_close).pack(side=RIGHT)
        self._button(btns, "de_cancel", style="Secondary.TButton",
                   command=win.destroy).pack(side=RIGHT, padx=(0, 10))
        win.bind("<Escape>", lambda _e: win.destroy())
        wm_entry.focus_set()

    # ---------------------- Aggiornamenti ---------------------------------
    def check_for_updates(self, manual: bool = False) -> None:
        # Campo vuoto = release pubblicate su GitHub
        source = str(self.settings.get("update_source") or "").strip() or DEFAULT_UPDATE_SOURCE
        if self._update_check_running:
            return
        self._update_check_running = True

        def work():
            try:
                manifest = fetch_update_manifest(source)
                self._post("update_result", {"manual": manual, "manifest": manifest})
            except Exception as error:  # noqa: BLE001 - rete/percorso non raggiungibile
                self._post("update_result", {"manual": manual, "error": str(error)})

        threading.Thread(target=work, daemon=True).start()

    def _handle_update_result(self, payload: dict[str, Any]) -> None:
        self._update_check_running = False
        manual = bool(payload.get("manual"))
        if payload.get("error"):
            # Controllo automatico: offline o sorgente irraggiungibile non deve disturbare
            if manual:
                messagebox.showerror(self.tr("up_title"),
                                     self.tr("up_error", e=payload["error"]))
            return
        self.settings["last_update_check"] = date.today().isoformat()
        self._persist_settings()
        manifest = payload.get("manifest") or {}
        remote = manifest.get("version", "")
        if not remote or parse_version(remote) <= parse_version(APP_VERSION):
            if manual:
                messagebox.showinfo(self.tr("up_title"),
                                    self.tr("up_latest", version=APP_VERSION))
            return
        notes = manifest.get("notes", "")
        body = self.tr("up_available", new=remote, current=APP_VERSION)
        if notes:
            body += "\n\n" + notes
        url = manifest.get("url", "")

        def open_update():
            if url.lower().startswith(("http://", "https://")):
                webbrowser.open(url)
            else:
                # Percorso di rete: si apre la cartella, l'utente avvia il setup
                setup = Path(url)
                open_folder(setup.parent if setup.suffix else setup)

        if manual:
            # Dalla finestra Impostazioni (modale) serve una risposta esplicita
            if not url:
                messagebox.showinfo(self.tr("up_title"), body)
            elif messagebox.askyesno(self.tr("up_title"), body + "\n\n" + self.tr("up_open")):
                open_update()
            return
        # Controllo automatico all'avvio: avviso discreto che non interrompe il lavoro
        self._toast(body, "info", (self.tr("up_get"), open_update) if url else None, 15000)


# --------------------------- MAIN -----------------------------------------

def _enable_dpi_awareness() -> None:
    if os.name != "nt":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main() -> None:
    if tk is None:
        raise SystemExit(
            "Tkinter non è disponibile in Python. Installa una versione di Python "
            "con il supporto Tk per avviare l'interfaccia desktop."
        )
    _enable_dpi_awareness()
    root = tk.Tk()
    FormazioniApp(root)
    root.mainloop()


if __name__ == "__main__":
    # Nell'exe i processi del batch parallelo rilanciano l'eseguibile stesso:
    # freeze_support li intercetta prima che aprano un'altra finestra.
    import multiprocessing

    multiprocessing.freeze_support()
    main()
