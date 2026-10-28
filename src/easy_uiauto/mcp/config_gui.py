"""Local Tk configuration editor; also runnable directly as a lean frozen app."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from queue import Empty, Queue
from tkinter import filedialog, messagebox, ttk

if __package__:
    from . import configuration
else:
    import configuration


class ConfigurationWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.closed = False
        self.busy = False
        self.results = Queue()
        self.poll_id = None
        self.root.title("easy-uiauto 配置")
        self.root.geometry("900x330")
        self.root.minsize(720, 330)
        self.root.protocol("WM_DELETE_WINDOW", self.close)

        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("TLabel", font=("Microsoft YaHei UI", 10))
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(10, 4))
        self.values = {name: tk.StringVar(root) for name in (
            "knowledge_dir", "api_url", "api_key", "api_model",
        )}
        self.status = tk.StringVar(root, value="API Model 可手动输入，也可点击“获取”后选择。")
        form = ttk.Frame(root, padding=(22, 20, 22, 16))
        form.pack(fill="both", expand=True)
        form.columnconfigure(1, weight=1)
        labels = (("knowledge_dir", "知识库路径"), ("api_url", "API URL"),
                  ("api_key", "API Key"), ("api_model", "API Model"))
        for row, (name, label) in enumerate(labels):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", padx=(0, 18), pady=9)
            if name == "api_model":
                field = ttk.Combobox(form, textvariable=self.values[name], state="normal")
                self.model_box = field
            else:
                # The local API Key field is deliberately readable, as requested.
                field = ttk.Entry(form, textvariable=self.values[name])
            field.configure(font=("Microsoft YaHei UI", 10))
            field.grid(row=row, column=1, sticky="ew", pady=9, ipady=3)
            if name == "knowledge_dir":
                self.directory_field = field
            if row in (1, 2):
                ttk.Label(form, text="", width=8).grid(row=row, column=2, padx=(14, 0))

        ttk.Button(form, text="浏览", width=8, command=self.browse).grid(
            row=0, column=2, padx=(14, 0), sticky="ew")
        self.fetch_button = ttk.Button(form, text="获取", width=8, command=self.fetch)
        self.fetch_button.grid(row=3, column=2, padx=(14, 0), sticky="ew")
        status_label = ttk.Label(form, textvariable=self.status, wraplength=810)
        status_label.grid(row=4, column=0, columnspan=3, sticky="w", pady=(9, 0))
        form.bind("<Configure>", lambda event: status_label.configure(
            wraplength=max(400, event.width - 44)))
        ttk.Separator(root).pack(fill="x")
        footer = ttk.Frame(root, padding=(22, 10, 22, 12))
        footer.pack(fill="x")
        ttk.Button(footer, text="关闭", width=9, command=self.close).pack(side="right")
        ttk.Button(footer, text="保存配置", width=12, command=self.save).pack(
            side="right", padx=(0, 12))
        try:
            for name, value in configuration.load_gui_configuration().items():
                self.values[name].set(value)
        except OSError:
            self.status.set("读取现有配置失败，可填写后重新保存。")
        self.initial = self.current_values()
        self.directory_field.focus_set()
        self.poll_id = self.root.after(100, self.poll)

    def current_values(self) -> dict[str, str]:
        return {name: value.get() for name, value in self.values.items()}

    def browse(self):
        directory = Path(self.values["knowledge_dir"].get()).expanduser()
        while not directory.is_dir() and directory.parent != directory:
            directory = directory.parent
        selected = filedialog.askdirectory(
            parent=self.root, title="选择知识库目录", mustexist=True,
            initialdir=str(directory) if directory.is_dir() else str(Path.home()),
        )
        if selected:
            self.values["knowledge_dir"].set(str(Path(selected)))

    def fetch(self):
        if self.busy:
            return
        api_url, api_key = self.values["api_url"].get(), self.values["api_key"].get()
        try:
            configuration.api_endpoints(api_url)
            if not api_key.strip() or "\r" in api_key or "\n" in api_key:
                raise ValueError("请先填写有效的 API Key")
        except ValueError as error:
            messagebox.showerror("无法获取模型", str(error), parent=self.root)
            return
        self.busy = True
        self.fetch_button.configure(state="disabled", text="获取中")
        self.status.set("正在获取模型列表…")
        threading.Thread(target=self.fetch_worker, args=(api_url, api_key), daemon=True).start()

    def fetch_worker(self, api_url: str, api_key: str):
        # Network work never touches Tk. Discard results if either field changed.
        try:
            result = configuration.fetch_models(api_url, api_key)
            self.results.put((api_url, api_key, result, None))
        except (ValueError, RuntimeError) as error:
            self.results.put((api_url, api_key, None, str(error)))
        except Exception:
            self.results.put((api_url, api_key, None, "获取失败，请检查网络或手动填写 API Model"))

    def poll(self):
        self.poll_id = None
        if self.closed:
            return
        try:
            api_url, api_key, models, error = self.results.get_nowait()
        except Empty:
            pass
        else:
            self.busy = False
            self.fetch_button.configure(state="normal", text="获取")
            if (api_url, api_key) != (self.values["api_url"].get(), self.values["api_key"].get()):
                self.status.set("API 配置已修改，请重新获取模型列表。")
            elif error:
                self.status.set(error)
                messagebox.showerror("获取模型失败", error, parent=self.root)
            else:
                self.model_box.configure(values=models)
                if not self.values["api_model"].get().strip():
                    self.values["api_model"].set(models[0])
                self.status.set(f"已获取 {len(models)} 个模型，可下拉选择或手动输入。")
        if not self.closed:
            self.poll_id = self.root.after(100, self.poll)

    def save(self):
        values = self.current_values()
        try:
            saved = configuration.save_gui_configuration(**values)
        except (ValueError, RuntimeError, OSError) as error:
            detail = str(error) if isinstance(error, (ValueError, RuntimeError)) else "目录或用户配置无法写入，请检查路径与权限。"
            messagebox.showerror("保存失败", detail, parent=self.root)
            return False
        self.values["knowledge_dir"].set(saved["knowledge_dir"])
        self.values["api_url"].set(saved["api_url"])
        self.initial = self.current_values()
        self.status.set("配置已保存。下次启动或调用新版 MCP 时生效。")
        return True

    def close(self):
        if self.current_values() != self.initial:
            answer = messagebox.askyesnocancel("尚未保存", "关闭前保存配置？", parent=self.root)
            if answer is None or (answer and not self.save()):
                return
        self.closed = True
        if self.poll_id is not None:
            self.root.after_cancel(self.poll_id)
        self.root.destroy()


def main():
    try:
        import ctypes
        # Set DPI awareness before creating Tk, without changing system-wide settings.
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass
    root = tk.Tk()
    ConfigurationWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
