"""Widget riutilizzabili: tooltip e selettore data."""

from __future__ import annotations

import calendar
from datetime import date

from ..config import MONTH_KEYS
from ..tkcompat import BOTH, tk, ttk

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
