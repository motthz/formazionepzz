"""Tkinter opzionale: senza Tk si possono comunque usare le funzioni di generazione."""

try:
    import tkinter as tk
    from tkinter import BOTH, END, LEFT, RIGHT, BooleanVar, StringVar, X, filedialog, messagebox, ttk
except ImportError:
    tk = None  # type: ignore[assignment]
    BOTH = END = LEFT = RIGHT = X = None  # type: ignore[assignment]
    BooleanVar = StringVar = filedialog = messagebox = ttk = None  # type: ignore[assignment]

__all__ = ["BOTH", "END", "LEFT", "RIGHT", "X", "BooleanVar", "StringVar", "filedialog",
           "messagebox", "tk", "ttk"]
