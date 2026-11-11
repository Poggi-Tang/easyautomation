"""Repository branding shared by source GUIs and the lean configuration EXE."""
from __future__ import annotations

import os
from pathlib import Path
import sys
import tkinter as tk


def apply_window_icon(window):
    """Set this window's icon and the default inherited by its child windows."""
    resources = Path(__file__).resolve().parents[1] / "resources"
    if not resources.is_dir() and hasattr(sys, "_MEIPASS"):
        resources = Path(sys._MEIPASS) / "resources"
    if os.name == "nt":
        try:
            icon = str(resources / "easy-uiauto.ico")
            window.iconbitmap(default=icon)
            window.iconbitmap(icon)
            return True
        except tk.TclError:
            pass
    try:
        image = tk.PhotoImage(master=window, file=str(resources / "easy-uiauto.png"))
        window.iconphoto(True, image)
        window._easy_uiauto_icon = image
        return True
    except (tk.TclError, OSError):
        # Missing branding must not prevent opening the configuration editor.
        return False
