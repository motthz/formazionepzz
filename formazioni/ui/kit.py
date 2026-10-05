"""Kit grafico di Formazioni PZZ.

Disegna con Pillow (anti-aliasing via supersampling) gli elementi ttk arrotondati:
pulsanti, campi, menu a tendina, caselle di spunta, barre di avanzamento e di
scorrimento, etichette "pillola" e i pannelli/card con angoli arrotondati e ombra.
Tutte le misure sono scalate sui DPI del monitor, cosi' restano nitide a 100-200%.

Uso (in FormazioniApp._configure_style, dopo gli style.configure di base):
    kit = UiKit(root)
    kit.install(style, palette)
"""
from __future__ import annotations

import functools
import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageDraw, ImageFilter, ImageTk

SS = 4  # fattore di supersampling per bordi morbidi

# Stili ttk le cui immagini hanno angoli trasparenti: il colore sotto gli angoli
# e' l'opzione "background" dello stile, che va adattata al contenitore.
_ADAPT_CLASSES = ("TButton", "TCombobox", "TEntry", "TCheckbutton", "TLabel", "TScrollbar")
_PILL_LABELS = ("Count.TLabel", "Gold.TLabel", "Secure.TLabel")


def _rgb(color: str) -> tuple[int, int, int]:
    color = color.lstrip("#")
    if len(color) == 3:
        color = "".join(c * 2 for c in color)
    return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)


def mix(c1: str, c2: str, t: float) -> str:
    """Colore intermedio fra c1 e c2 (t=0 -> c1, t=1 -> c2)."""
    a, b = _rgb(c1), _rgb(c2)
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def rounded_image(w: int, h: int, r: float, fill, outline=None, width: float = 1.0) -> Image.Image:
    """Rettangolo arrotondato RGBA con bordo opzionale, anti-aliased."""
    W, H = max(1, w) * SS, max(1, h) * SS
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    rad = max(0, r) * SS
    if outline:
        d.rounded_rectangle((0, 0, W - 1, H - 1), rad, fill=outline)
        inset = max(1, round(width * SS))
        d.rounded_rectangle((inset, inset, W - 1 - inset, H - 1 - inset),
                            max(0, rad - inset), fill=fill)
    else:
        d.rounded_rectangle((0, 0, W - 1, H - 1), rad, fill=fill)
    # Media dei blocchi SSxSS: e' l'anti-aliasing corretto ed e' molto piu' veloce di
    # LANCZOS, che con le immagini grandi degli elementi rallentava l'avvio
    return im.reduce(SS)


class UiKit:
    def __init__(self, root: tk.Misc) -> None:
        self.root = root
        try:
            self.scale = max(1.0, float(root.winfo_fpixels("1i")) / 96.0)
        except tk.TclError:
            self.scale = 1.0
        self._images: list[ImageTk.PhotoImage] = []
        self._old_images: list[ImageTk.PhotoImage] = []
        self._gen = 0
        self._variants: dict[str, str] = {}
        self._adapt_bound = False
        self.style: ttk.Style | None = None
        self.p: dict[str, str] = {}

    # ----------------------------------------------------------------- utils
    def px(self, value: float) -> int:
        return max(1, int(round(value * self.scale)))

    def photo(self, image: Image.Image) -> ImageTk.PhotoImage:
        img = ImageTk.PhotoImage(image, master=self.root)
        self._images.append(img)  # i PhotoImage vanno tenuti in vita
        return img

    def pill(self, w: int, h: int, fill: str, r: float | None = None) -> ImageTk.PhotoImage:
        return self.photo(rounded_image(w, h, h / 2 if r is None else r, fill))

    def _rr(self, w, h, r, fill, outline=None, width=1.0) -> ImageTk.PhotoImage:
        return self.photo(rounded_image(w, h, r, fill, outline, width))

    def _tile(self, sz: int) -> tuple[int, int]:
        """Dimensioni delle immagini degli elementi ttk estendibili. ttk riempie il
        widget ripetendo la parte centrale dell'immagine: con immagini di pochi pixel
        un campo largo richiedeva centinaia di disegni con trasparenza a ogni
        aggiornamento (scorrimento, ridimensionamento, schermo intero). Con un centro
        ampio ne bastano pochi. La dimensione minima del widget resta sz (opzioni
        width/height degli elementi)."""
        return sz + self.px(200), sz + self.px(48)

    # --------------------------------------------------------------- install
    def install(self, style: ttk.Style, p: dict[str, str]) -> None:
        """Crea elementi e layout ttk per la palette p (va richiamato a ogni cambio tema)."""
        self.style, self.p = style, p
        # Le immagini del tema precedente restano vive per un giro (i widget vecchi
        # vengono distrutti subito dopo); quelle di due temi fa si possono liberare.
        self._old_images, self._images = self._images, []
        self._gen += 1
        self._variants = {}
        g = f"pzz{self._gen}"

        self._buttons(style, g, p)
        self._fields(style, g, p)
        self._checkbutton(style, g, p)
        self._progressbar(style, g, p)
        self._scrollbars(style, g, p)
        self._pills(style, g, p)
        self._misc(style, p)

        if not self._adapt_bound:
            for cls in _ADAPT_CLASSES:
                self.root.bind_class(cls, "<Map>", self._on_map, add="+")
            self.root.bind_class("TProgressbar", "<Map>", self._watch_progress, add="+")
            self._adapt_bound = True

    def _buttons(self, style, g, p) -> None:
        r = self.px(8)
        sz = 2 * r + 4
        specs = (
            ("Primary.TButton", p["primary"], p["primary_hover"], p["primary_press"], None, p["on_primary"]),
            ("Accent.TButton", p["accent"], p["accent_hover"], p["accent_press"], None, p["on_accent"]),
            ("Secondary.TButton", p["secondary"], p["secondary_hover"], p["secondary_press"],
             p["secondary_border"], p["on_secondary"]),
            ("TButton", p["secondary"], p["secondary_hover"], p["secondary_press"],
             p["secondary_border"], p["on_secondary"]),
        )
        for name, base, hover, press, border, fg in specs:
            bw = self.px(1)
            tw, th = self._tile(sz)
            normal = self._rr(tw, th, r, base, border, bw)
            hov = self._rr(tw, th, r, hover, border and mix(border, p["focus"], 0.35), bw)
            prs = self._rr(tw, th, r, press, border, bw)
            dis = self._rr(tw, th, r, p["disabled_bg"])
            el = f"{g}.{name}.bd"
            style.element_create(el, "image", normal, ("disabled", dis), ("pressed", prs),
                                 ("active", hov), border=r, sticky="nsew", width=sz, height=sz)
            style.layout(name, [(el, {"sticky": "nsew", "children": [
                ("Button.padding", {"sticky": "nsew", "children": [
                    ("Button.label", {"sticky": "nsew"})]})]})])
            style.configure(name, background=p["surface"], foreground=fg, borderwidth=0,
                            focusthickness=0, anchor="center")
            style.map(name, background=[], foreground=[("disabled", p["disabled_fg"])])
        if not style.lookup("TButton", "font"):
            style.configure("TButton", font=("Segoe UI Semibold", 9), padding=(14, 8))

    def _field_images(self, fill, border, hover, focus, disabled):
        r = self.px(7)
        sz = 2 * r + 6
        tw, th = self._tile(sz)
        return r, (
            self._rr(tw, th, r, fill, border, self.px(1)),
            self._rr(tw, th, r, fill, hover, self.px(1)),
            self._rr(tw, th, r, fill, focus, max(1.5, self.scale * 1.6)),
            self._rr(tw, th, r, disabled, border, self.px(1)),
        )

    def _chevron(self, color: str) -> ImageTk.PhotoImage:
        w, h = self.px(22), self.px(16)
        im = Image.new("RGBA", (w * SS, h * SS), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        cx, cy, s = w * SS * 0.45, h * SS * 0.5, self.px(3.5) * SS
        lw = max(SS, round(self.scale * 1.7 * SS))
        pts = [(cx - s, cy - s / 2), (cx, cy + s / 2), (cx + s, cy - s / 2)]
        d.line(pts, fill=color, width=lw, joint="curve")
        for x, y in (pts[0], pts[2]):
            d.ellipse((x - lw / 2, y - lw / 2, x + lw / 2, y + lw / 2), fill=color)
        return self.photo(im.resize((w, h), Image.LANCZOS))

    def _fields(self, style, g, p) -> None:
        pad_y = self.px(6)
        # --- Entry
        r, (n, h, f, d) = self._field_images(p["field"], p["field_border"], p["field_hover"],
                                             p["focus"], p["disabled_bg"])
        el = f"{g}.Entry.field"
        style.element_create(el, "image", n, ("disabled", d), ("focus", f), ("hover", h),
                             border=r, sticky="nsew", width=2 * r + 6, height=2 * r + 6)
        style.layout("TEntry", [(el, {"sticky": "nswe", "children": [
            ("Entry.padding", {"sticky": "nswe", "children": [
                ("Entry.textarea", {"sticky": "nswe"})]})]})])
        style.configure("TEntry", background=p["surface"], foreground=p["text"],
                        padding=(self.px(10), pad_y), insertcolor=p["text"],
                        selectbackground=p["select_bg"], selectforeground=p["select_fg"])
        style.map("TEntry", bordercolor=[], lightcolor=[], darkcolor=[], fieldbackground=[],
                  foreground=[("disabled", p["disabled_fg"])])

        # --- Combobox (normale e variante scura per l'intestazione)
        variants = (
            ("TCombobox", p["surface"], p["field"], p["field_border"], p["field_hover"],
             p["focus"], p["text"], p["muted"]),
            ("Header.TCombobox", p["header_bg"], p["header_field"], p["header_border"],
             p["header_hover"], p["gold"], p["header_fg"], p["header_fg"]),
        )
        for name, bg, fill, border, hover, focus, fg, arrow in variants:
            r, (n, h, f, d) = self._field_images(fill, border, hover, focus, p["disabled_bg"])
            el = f"{g}.{name}.field"
            style.element_create(el, "image", n, ("disabled", d), ("focus", f), ("hover", h),
                                 ("pressed", f), border=r, sticky="nsew", width=2 * r + 6, height=2 * r + 6)
            arr = f"{g}.{name}.arrow"
            style.element_create(arr, "image", self._chevron(arrow),
                                 ("disabled", self._chevron(p["disabled_fg"])), sticky="")
            style.layout(name, [(el, {"sticky": "nswe", "children": [
                (arr, {"side": "right", "sticky": ""}),
                ("Combobox.padding", {"expand": "1", "sticky": "nswe", "children": [
                    ("Combobox.textarea", {"sticky": "nswe"})]})]})])
            style.configure(name, background=bg, foreground=fg, padding=(self.px(8), pad_y, 0, pad_y),
                            insertcolor=fg, selectbackground=fill, selectforeground=fg)
            style.map(name, bordercolor=[], lightcolor=[], darkcolor=[], fieldbackground=[],
                      background=[], arrowcolor=[],
                      foreground=[("disabled", p["disabled_fg"])],
                      selectbackground=[("readonly", fill)], selectforeground=[("readonly", fg)])

        # Lista a discesa dei combobox
        for opt, val in (("background", p["surface"]), ("foreground", p["text"]),
                         ("selectBackground", p["primary"]), ("selectForeground", p["on_primary"]),
                         ("font", ("Segoe UI", 10)), ("relief", "flat"), ("borderWidth", 0)):
            self.root.option_add(f"*TCombobox*Listbox.{opt}", val)

    def _checkbutton(self, style, g, p) -> None:
        box = self.px(18)
        gap = self.px(8)
        rad = self.px(5)
        bw = max(1.2, self.scale * 1.4)

        def make(fill, outline, check: str | None) -> ImageTk.PhotoImage:
            im = Image.new("RGBA", (box + gap, box), (0, 0, 0, 0))
            im.paste(rounded_image(box, box, rad, fill, outline, bw), (0, 0))
            if check:
                big = Image.new("RGBA", (box * SS, box * SS), (0, 0, 0, 0))
                d = ImageDraw.Draw(big)
                s = box * SS
                lw = max(SS, round(self.scale * 2.1 * SS))
                pts = [(s * 0.26, s * 0.52), (s * 0.43, s * 0.69), (s * 0.75, s * 0.33)]
                d.line(pts, fill=check, width=lw, joint="curve")
                for x, y in (pts[0], pts[2]):
                    d.ellipse((x - lw / 2, y - lw / 2, x + lw / 2, y + lw / 2), fill=check)
                im.alpha_composite(big.resize((box, box), Image.LANCZOS))
            return self.photo(im)

        off = make(p["field"], p["check_border"], None)
        off_h = make(p["field"], p["focus"], None)
        on = make(p["primary"], None, p["on_primary"])
        on_h = make(p["primary_hover"], None, p["on_primary"])
        dis = make(p["disabled_bg"], p["field_border"], None)
        dis_on = make(p["disabled_bg"], p["field_border"], p["disabled_fg"])
        el = f"{g}.Check.indicator"
        style.element_create(el, "image", off, ("disabled", "selected", dis_on), ("disabled", dis),
                             ("selected", "active", on_h), ("selected", on), ("active", off_h),
                             sticky="")
        style.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
            (el, {"side": "left", "sticky": ""}),
            ("Checkbutton.label", {"side": "left", "sticky": "nswe"})]})])
        style.configure("TCheckbutton", background=p["surface"], foreground=p["text"], padding=(0, 2))
        style.map("TCheckbutton", background=[], foreground=[("disabled", p["disabled_fg"])])

    def _progressbar(self, style, g, p) -> None:
        h = self.px(8)
        r = h / 2
        trough = self._rr(self.px(240), h, r, p["trough"])
        bar = self._rr(self.px(240), h, r, p["primary"])
        tr_el, bar_el = f"{g}.Pb.trough", f"{g}.Pb.bar"
        style.element_create(tr_el, "image", trough, border=(int(r) + 1, 0), sticky="nsew",
                             width=h * 3, height=h)
        # width=0: a valore 0 la barra non deve mostrare un segmento minimo
        style.element_create(bar_el, "image", bar, border=(int(r) + 1, 0), sticky="nsew", width=0)
        style.layout("Horizontal.TProgressbar", [(tr_el, {"sticky": "nswe", "children": [
            (bar_el, {"side": "left", "sticky": "ns"})]})])
        style.configure("Horizontal.TProgressbar", background=p["surface"], thickness=h)
        # A valore 0 ttk disegnerebbe comunque le estremita' arrotondate della barra:
        # le barre ferme usano uno stile con la sola guida (vedi _watch_progress).
        style.layout("Idle.Horizontal.TProgressbar", [(tr_el, {"sticky": "nswe"})])
        style.configure("Idle.Horizontal.TProgressbar", background=p["surface"], thickness=h)

    def _scrollbars(self, style, g, p) -> None:
        w = self.px(12)
        m = self.px(3)
        tw = w - 2 * m  # spessore del cursore
        length = self.px(160)  # immagine lunga: meno ripetizioni (vedi _tile)
        min_length = tw * 3

        def thumb(color, vertical: bool) -> ImageTk.PhotoImage:
            im = Image.new("RGBA", (w, length) if vertical else (length, w), (0, 0, 0, 0))
            pill = rounded_image(tw, length, tw / 2, color) if vertical else \
                rounded_image(length, tw, tw / 2, color)
            im.paste(pill, (m, 0) if vertical else (0, m))
            return self.photo(im)

        for orient, vertical in (("Vertical", True), ("Horizontal", False)):
            tr_el, th_el = f"{g}.{orient}.Sb.trough", f"{g}.{orient}.Sb.thumb"
            border = (0, tw // 2 + 1) if vertical else (tw // 2 + 1, 0)
            # Guida trasparente lunga: anche lei viene ripetuta a piastrelle
            empty = self.photo(Image.new("RGBA", (w, length) if vertical else (length, w),
                                         (0, 0, 0, 0)))
            style.element_create(tr_el, "image", empty, sticky="nsew", width=w, height=w)
            style.element_create(th_el, "image", thumb(p["thumb"], vertical),
                                 ("pressed", thumb(p["thumb_hover"], vertical)),
                                 ("active", thumb(p["thumb_hover"], vertical)),
                                 border=border, sticky="nsew",
                                 width=w if vertical else min_length, height=min_length if vertical else w)
            name = f"{orient}.TScrollbar"
            style.layout(name, [(tr_el, {"sticky": "ns" if vertical else "ew", "children": [
                (th_el, {"expand": "1", "sticky": "nswe"})]})])
            style.configure(name, background=p["surface"], width=w, arrowsize=w)
            style.map(name, background=[])

    def _pills(self, style, g, p) -> None:
        r = self.px(9)
        sz = 2 * r + 4
        for name, key in zip(_PILL_LABELS, ("count_bg", "gold_bg", "secure_bg")):
            el = f"{g}.{name}.pill"
            tw, th = self._tile(sz)
            style.element_create(el, "image", self._rr(tw, th, r, p[key]), border=r, sticky="nsew",
                                 width=sz, height=sz)
            style.layout(name, [(el, {"sticky": "nswe", "children": [
                ("Label.padding", {"sticky": "nswe", "children": [
                    ("Label.label", {"sticky": "nswe"})]})]})])
            style.configure(name, background=p["surface"])

    def _misc(self, style, p) -> None:
        style.configure("Treeview", borderwidth=0, relief="flat")
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        style.configure("Treeview.Heading", relief="flat", borderwidth=0)

    def _watch_progress(self, event) -> None:
        bar = event.widget
        if getattr(bar, "_pzz_watched", False):
            return
        bar._pzz_watched = True

        def tick() -> None:
            try:
                if str(bar.cget("orient")) != "horizontal":
                    return
                moving = float(bar.cget("value")) > 0 or str(bar.cget("mode")) == "indeterminate"
                want = "Horizontal.TProgressbar" if moving else "Idle.Horizontal.TProgressbar"
                if str(bar.cget("style")) != want:
                    bar.configure(style=want)
                bar.after(120, tick)
            except (tk.TclError, ValueError):
                pass

        tick()

    # ------------------------------------------------- adattamento angoli
    def _bg_of(self, widget) -> str | None:
        try:
            return str(widget.cget("background"))
        except tk.TclError:
            pass
        try:
            st = str(widget.cget("style")) or widget.winfo_class()
            return self.style.lookup(st, "background") or None
        except Exception:
            return None

    def _same(self, c1: str, c2: str) -> bool:
        try:
            return self.root.winfo_rgb(c1) == self.root.winfo_rgb(c2)
        except tk.TclError:
            return c1 == c2

    def _on_map(self, event) -> None:
        w = event.widget
        if self.style is None or not isinstance(w, ttk.Widget):
            return
        try:
            current = str(w.cget("style"))
            cls = w.winfo_class()
            if not current:
                if cls == "TScrollbar":
                    current = ("Vertical" if str(w.cget("orient")) == "vertical" else "Horizontal") + ".TScrollbar"
                else:
                    current = cls
            base = current.split(".", 1)[1] if current.startswith("Bg_") else current
            if cls == "TLabel" and base not in _PILL_LABELS:
                return
            parent = w.nametowidget(w.winfo_parent())
            bg = self._bg_of(parent)
            if not bg:
                return
            if self._same(bg, self.style.lookup(base, "background") or ""):
                target = base
            else:
                key = f"{bg}|{base}"
                target = self._variants.get(key)
                if target is None:
                    rgb = self.root.winfo_rgb(bg)
                    target = "Bg_%02x%02x%02x.%s" % (rgb[0] >> 8, rgb[1] >> 8, rgb[2] >> 8, base)
                    self.style.configure(target, background=bg)
                    self.style.map(target, background=[])
                    self._variants[key] = target
            if target != current:
                w.configure(style=target)
        except Exception:
            pass

    # ------------------------------------------------------ immagini grandi
    def header_image(self, w: int, h: int, base: str, glow: str, gold: str, primary: str,
                     flat_from: float = 0.55) -> ImageTk.PhotoImage:
        """Sfondo dell'intestazione: sfumatura, bagliori morbidi e filetto oro->primario.

        A destra di flat_from (frazione di w) il colore resta uguale a base, perche'
        li' ci sono controlli con sfondo pieno.
        """
        w, h = max(2, w), max(2, h)
        im = Image.new("RGB", (w, h), base)
        # Sfumatura orizzontale: piu' chiara a sinistra, base dalla zona controlli in poi
        lighter = mix(base, glow, 0.22)
        ramp = Image.linear_gradient("L").rotate(90).resize((w, h))  # 0 a sinistra -> 255 a destra
        ramp = ramp.point(lambda v: max(0, int((1 - (v / 255) / flat_from) * 255)))
        im = Image.composite(Image.new("RGB", (w, h), lighter), im, ramp)
        # Bagliori sfocati: disegnati e sfocati a 1/4 di risoluzione e poi ingranditi
        # (sono comunque sfumati), cosi' il ridisegno al ridimensionamento e' rapido.
        q = 4
        lw, lh = max(1, w // q), max(1, h // q)
        layer = Image.new("RGBA", (lw, lh), (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        R = h * 1.1 / q
        cx, cy = w * 0.30 / q, h * 1.35 / q
        d.ellipse((cx - R, cy - R, cx + R, cy + R), fill=_rgb(glow) + (70,))
        cx2, cy2, R2 = (w - h * 1.9) / q, h * 1.55 / q, h * 0.9 / q
        d.ellipse((cx2 - R2, cy2 - R2, cx2 + R2, cy2 + R2), fill=_rgb(glow) + (55,))
        layer = layer.filter(ImageFilter.GaussianBlur(h * 0.35 / q)).resize((w, h), Image.BILINEAR)
        im = Image.alpha_composite(im.convert("RGBA"), layer)
        # Puntinatura leggera in alto a sinistra
        dots = ImageDraw.Draw(im)
        step = self.px(14)
        dot = _rgb(mix(base, "#ffffff", 0.09)) + (255,)
        for yy in range(step, int(h * 0.75), step):
            for xx in range(int(w * 0.62), w - step, step):
                fade = (xx - w * 0.62) / (w * 0.38)
                if (xx // step + yy // step) % 2 == 0 and fade > 0.15 and yy > h * 0.55:
                    dots.point((xx, yy), fill=dot)
        # Filetto inferiore sfumato oro -> primario
        lh = self.px(3)
        line = Image.linear_gradient("L").rotate(90).resize((w, lh))
        band = Image.composite(Image.new("RGB", (w, lh), primary), Image.new("RGB", (w, lh), gold), line)
        im.paste(band, (0, h - lh))
        # Non va in self._images: lo tiene in vita chi lo usa, e a ogni ridimensionamento
        # le immagini vecchie si accumulerebbero in memoria.
        return ImageTk.PhotoImage(im.convert("RGB"), master=self.root)


class RoundedPanel(tk.Canvas):
    """Card con angoli arrotondati, bordo sottile e ombra morbida.

    Il contenuto va messo in panel.inner (un tk.Frame del colore della card).
    Si usa come un normale widget (grid/pack) al posto del vecchio frame-ombra.
    """

    def __init__(self, parent, kit: UiKit, *, outer_bg: str, fill: str, border: str,
                 shadow: str, shadow_alpha: int = 40, radius: float = 14,
                 padx: int = 20, pady: int = 16) -> None:
        super().__init__(parent, bg=outer_bg, highlightthickness=0, bd=0, width=10, height=10)
        self.kit = kit
        px = kit.px
        self._r = px(radius)
        self._blur = px(7)
        self._dy = px(3)
        self._m = self._blur + px(3)  # margine riservato all'ombra
        self._inset = self._m + max(px(1) + 1, int(self._r * 0.32))
        self._tpl = self._template(outer_bg, fill, border, shadow, shadow_alpha)
        self.inner = tk.Frame(self, bg=fill, padx=padx, pady=pady)
        self._win = self.create_window(self._inset, self._inset, window=self.inner, anchor="nw")
        # Ordine di disegno: riempimento, bordi, angoli (sopra ai bordi), contenuto
        self._fill_item = self.create_rectangle(0, 0, 0, 0, fill=fill, outline="")
        self._edge_items = [self.create_image(0, 0, anchor="nw") for _ in range(4)]
        self._corner_items = [self.create_image(0, 0, anchor="nw") for _ in range(4)]
        for item in (*self._corner_items, *self._edge_items, self._fill_item):
            self.tag_lower(item)
        self._size = (0, 0)
        self._req = 0
        try:
            screen = max(self.winfo_screenwidth(), self.winfo_screenheight())
        except tk.TclError:
            screen = 1920
        self._build_pieces(screen)
        self.bind("<Configure>", self._on_configure)
        self._poll()

    def required_width(self) -> int:
        """Larghezza minima perche' il contenuto non venga tagliato."""
        return self.inner.winfo_reqwidth() + 2 * self._inset

    def _template(self, outer_bg, fill, border, shadow, alpha) -> Image.Image:
        C = self._m + self._r
        K = 8
        S = 2 * C + K
        im = Image.new("RGBA", (S, S), _rgb(outer_bg) + (255,))
        sh = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        ImageDraw.Draw(sh).rounded_rectangle(
            (self._m, self._m + self._dy, S - self._m - 1, S - self._m - 1 + self._dy),
            self._r, fill=_rgb(shadow) + (alpha,))
        im = Image.alpha_composite(im, sh.filter(ImageFilter.GaussianBlur(self._blur / 2)))
        card = rounded_image(S - 2 * self._m, S - 2 * self._m, self._r, fill, border, max(1.0, self.kit.scale))
        im.alpha_composite(card, (self._m, self._m))
        return im.convert("RGB")

    def _build_pieces(self, length: int) -> None:
        """Angoli e bordi come immagini fisse: al ridimensionamento si spostano solo
        gli elementi del canvas, senza ridisegnare un'immagine grande quanto la card
        (che rendeva lenti scorrimento, schermo intero e trascinamento della finestra)."""
        T = self._tpl
        C = self._m + self._r
        tw, th = T.size
        mid = tw // 2
        photo = functools.partial(ImageTk.PhotoImage, master=self)
        self._corners = [photo(T.crop(box)) for box in (
            (0, 0, C, C), (tw - C, 0, tw, C), (0, th - C, C, th), (tw - C, th - C, tw, th))]
        # Bordi lunghi quanto lo schermo: la parte in eccesso finisce sotto gli
        # angoli o fuori dal canvas, che la ritaglia.
        self._edges = [
            photo(T.crop((mid, 0, mid + 1, C)).resize((length, C))),
            photo(T.crop((mid, th - C, mid + 1, th)).resize((length, C))),
            photo(T.crop((0, mid, C, mid + 1)).resize((C, length))),
            photo(T.crop((tw - C, mid, tw, mid + 1)).resize((C, length))),
        ]
        self._edge_len = length
        for item, image in zip(self._edge_items + self._corner_items, self._edges + self._corners):
            self.itemconfigure(item, image=image)

    def _on_configure(self, event) -> None:
        if (event.width, event.height) == self._size:
            return
        self._size = (event.width, event.height)
        self.itemconfigure(self._win, width=max(1, event.width - 2 * self._inset),
                           height=max(self.inner.winfo_reqheight(), event.height - 2 * self._inset))
        self._redraw()

    def _redraw(self) -> None:
        w, h = self._size
        if w < 2 or h < 2:
            return
        C = self._m + self._r
        if max(w, h) > self._edge_len:
            self._build_pieces(max(w, h) + 200)
        top, bottom, left, right = self._edge_items
        self.coords(self._fill_item, C, C, max(C, w - C), max(C, h - C))
        self.coords(top, C, 0)
        self.coords(bottom, C, h - C)
        self.coords(left, 0, C)
        self.coords(right, w - C, C)
        for item, (x, y) in zip(self._corner_items, ((0, 0), (w - C, 0), (0, h - C), (w - C, h - C))):
            self.coords(item, x, y)

    def _poll(self) -> None:
        # Tk non notifica i cambi di dimensione richiesta del contenuto:
        # li controlliamo periodicamente per adattare l'altezza della card.
        try:
            req = self.inner.winfo_reqheight()
        except tk.TclError:
            return
        if req != self._req:
            self._req = req
            self.configure(height=req + 2 * self._inset)
            w, h = self._size
            if w and h:
                self.itemconfigure(self._win, height=max(req, h - 2 * self._inset))
        self.after(150, self._poll)
