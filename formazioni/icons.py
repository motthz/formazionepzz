"""Icone a linea dei pulsanti, disegnate al volo con Pillow."""

from __future__ import annotations

import re

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
    "st_check_now": "refresh", "st_open_log": "file", "secure_label": "lock", "bat_inline_title": "list",
    "ord_save": "check", "ord_delete": "trash", "ord_rename": "edit",
}


_LEADING_SYMBOLS_RE = re.compile(r"^[^\w(«\"'*]+", re.UNICODE)


def strip_leading_symbol(text: str) -> str:
    """Toglie l'emoji iniziale dai testi tradotti ("⬇  Genera" -> "Genera")."""
    return _LEADING_SYMBOLS_RE.sub("", text)
