"""Shared Tk components; styles are scoped to avoid changing other windows."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .mcp.gui_icons import apply_window_icon


def scroll_text(parent, **options):
    frame = ttk.Frame(parent)
    text = tk.Text(frame, **options)
    vertical = ttk.Scrollbar(frame, command=text.yview)
    horizontal = ttk.Scrollbar(frame, orient="horizontal", command=text.xview)
    text.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    frame.columnconfigure(0, weight=1); frame.rowconfigure(0, weight=1)
    text.grid(row=0, column=0, sticky="nsew"); vertical.grid(row=0, column=1, sticky="ns")
    if options.get("wrap") == "none": horizontal.grid(row=1, column=0, sticky="ew")
    frame.pack(fill="both", expand=True, pady=5)
    return text


def scroll_tree(parent, **options):
    frame = ttk.Frame(parent)
    tree = ttk.Treeview(frame, **options)
    scrollbar = ttk.Scrollbar(frame, command=tree.yview)
    horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=horizontal.set)
    frame.columnconfigure(0, weight=1); frame.rowconfigure(0, weight=1)
    tree.grid(row=0, column=0, sticky="nsew"); scrollbar.grid(row=0, column=1, sticky="ns")
    horizontal.grid(row=1, column=0, sticky="ew")
    frame.pack(fill="both", expand=True)
    return tree


class ScrollForm(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
        self.body = ttk.Frame(self.canvas, padding=10)
        item = self.canvas.create_window(0, 0, window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(item, width=e.width))


class WindowPicker(tk.Toplevel):
    """A fixed application context, shared by learning and combination execution."""
    def __init__(self, master, app_label, callback):
        from .mcp import native_windows

        super().__init__(master)
        apply_window_icon(self)
        self.title("绑定运行窗口 · " + app_label); self.geometry("900x520"); self.minsize(640, 380)
        self.transient(master.winfo_toplevel())
        self.callback, self.rows, self.windows = callback, {}, []
        self.enumerate = native_windows.enumerate_windows
        self.search = tk.StringVar(self)
        ttk.Label(self, text="绑定到：" + app_label, padding=10).pack(anchor="w")
        row = ttk.Frame(self, padding=(10, 0)); row.pack(fill="x")
        ttk.Label(row, text="搜索窗口").pack(side="left")
        ttk.Entry(row, textvariable=self.search).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="刷新", command=self.refresh).pack(side="right")
        self.tree = scroll_tree(self, columns=("class",), selectmode="browse")
        self.tree.heading("#0", text="窗口标题"); self.tree.heading("class", text="窗口类名")
        self.tree.column("#0", width=540); self.tree.column("class", width=280)
        footer = ttk.Frame(self, padding=10); footer.pack(fill="x")
        ttk.Button(footer, text="取消", command=self.destroy).pack(side="right")
        ttk.Button(footer, text="绑定所选窗口", command=self.accept).pack(side="right", padx=8)
        self.search.trace_add("write", lambda *_a: self.filter())
        self.refresh()

    def refresh(self):
        self.windows = [w for w in self.enumerate(visible_only=True) if w.get("title", "").strip()
                        and not w["title"].startswith("easy-uiauto") and w["title"] != self.title()]
        self.filter()

    def filter(self):
        self.tree.delete(*self.tree.get_children()); self.rows.clear()
        query = self.search.get().casefold()
        for window in self.windows:
            if query not in (window["title"] + " " + window["class_name"]).casefold(): continue
            key = str(window["handle"]); self.rows[key] = window
            self.tree.insert("", "end", iid=key, text=window["title"], values=(window["class_name"],))

    def accept(self):
        selected = self.tree.selection()
        if not selected: return
        window = self.rows[selected[0]]
        self.destroy(); self.callback(window)


class TaskMonitor:
    """Keep cancellation visible while the editor is minimized; hide for captures."""
    def __init__(self, master, task, title):
        self.task, self.root = task, tk.Toplevel(master)
        apply_window_icon(self.root)
        self.root.withdraw(); self.root.title("easy-uiauto · 正在执行")
        self.root.geometry("440x115+20+40"); self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.message = tk.StringVar(self.root, value=title)
        ttk.Label(self.root, textvariable=self.message, padding=10, wraplength=420).pack(fill="x")
        self.button = ttk.Button(self.root, text="停止当前任务", command=self.stop)
        self.button.pack(pady=5)
        self.root.protocol("WM_DELETE_WINDOW", self.stop)
        self.root.update_idletasks()
        # A separate toplevel (not transient) remains visible when the editor minimizes.
        import ctypes
        self.hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id()) or self.root.winfo_id()
        user = ctypes.windll.user32
        user.SetWindowLongW(self.hwnd, -20, user.GetWindowLongW(self.hwnd, -20) | 0x08000000 | 0x80)
        # Map Tk's children once before the worker can activate its target. Later
        # native show/hide calls must not invoke Tk's deiconify focus behavior.
        self.root.deiconify(); self.root.update_idletasks()
        self.hidden = False; self.show()
        self.poll_id = self.root.after(30, self.poll)

    def show(self):
        import ctypes
        ctypes.windll.user32.ShowWindow(self.hwnd, 4)  # SW_SHOWNOACTIVATE

    def stop(self):
        if self.task is None: return
        self.task.cancel(); self.button.configure(state="disabled"); self.message.set("正在停止，已尝试的操作不会重试。")

    def poll(self):
        if self.task is None:
            self.poll_id = self.root.after(30, self.poll); return
        request = self.task.folder / "capture-request"
        if request.exists():
            import ctypes
            ctypes.windll.user32.ShowWindow(self.hwnd, 0)
            self.root.update_idletasks(); self.hidden = True
            try:
                # Echo the nonce: an old acknowledgement cannot permit a new capture.
                nonce = request.read_text(encoding="ascii")
                (self.task.folder / "capture-ack").write_text(nonce, encoding="ascii")
            except OSError: pass
        elif self.hidden:
            self.hidden = False; self.show()
        self.poll_id = self.root.after(30, self.poll)

    def close(self):
        self.root.after_cancel(self.poll_id); self.root.destroy()
