"""Integrazione con Windows: icone finestra, notifiche, trascinamento, cartelle."""

from __future__ import annotations

import ctypes
import os
import platform
import queue
import struct
import subprocess
import tempfile
import zlib
from pathlib import Path

from .config import ICON_ICO, ICON_PNG, PREVIEW_PREFIX
from .tkcompat import tk


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


def cleanup_old_previews() -> None:
    for old in Path(tempfile.gettempdir()).glob(f"{PREVIEW_PREFIX}*.pdf"):
        try:
            old.unlink()
        except OSError:
            pass  # ancora aperta nel visualizzatore: sara' rimossa al prossimo avvio


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
