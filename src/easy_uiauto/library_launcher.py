"""Show startup failures for the console-free knowledge-library entry point."""

from __future__ import annotations

import tempfile
import traceback
from pathlib import Path
import sys


def main():
    # Frozen sys.executable points to this app, not to a Python interpreter.
    # Dispatch workers before importing Qt or constructing any main window.
    if len(sys.argv)>1 and sys.argv[1]=="--worker":
        from .library_tasks import main as run_worker
        return run_worker()
    try:
        from .qt_gui import main as run_gui

        run_gui()
    except Exception as error:
        path = Path(tempfile.gettempdir()) / "easy_uiauto" / "library-startup.log"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(traceback.format_exc(), encoding="utf-8")
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "知识库窗口无法启动",
            f"{error}\n\n请确认已安装 easy-uiauto[gui] 或完整功能依赖。\n详细日志：{path}",
            parent=root,
        )
        root.destroy()
        raise SystemExit(1)


if __name__ == "__main__":
    main()
