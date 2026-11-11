"""Software instructions, separate from control records and executable actions."""
import tkinter as tk
from tkinter import messagebox, ttk

from .gui_components import scroll_text


class UsagePanel(ttk.Frame):
    def __init__(self, parent, library, on_saved, open_action):
        super().__init__(parent, padding=12)
        self.library, self.on_saved, self.open_action = library, on_saved, open_action
        self.app_id, self.record, self.dirty, self.loading = "", {}, False, False
        self.title = tk.StringVar(self); self.status = tk.StringVar(self, value="说明软件怎么用，可关联可执行的组合动作。")
        ttk.Label(self, text="使用知识", font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w")
        bar = ttk.Frame(self); bar.pack(fill="x", pady=10)
        ttk.Button(bar, text="新建使用知识", command=self.new).pack(side="left")
        self.save_button = ttk.Button(bar, text="保存使用知识", command=self.save); self.save_button.pack(side="left", padx=8)
        self.title_entry = ttk.Entry(self, textvariable=self.title); self.title_entry.pack(fill="x")
        self.body = scroll_text(self, wrap="word", undo=True, height=10)
        ttk.Label(self, text="关联组合动作（可多选；双击打开）").pack(anchor="w", pady=5)
        self.links = tk.Listbox(self, height=5, selectmode="extended", exportselection=False)
        self.links.pack(fill="x"); self.links.bind("<<ListboxSelect>>", self.changed)
        self.links.bind("<Double-1>", self.open_link)
        ttk.Label(self, textvariable=self.status, padding=(0, 10), wraplength=800).pack(fill="x")
        self.title.trace_add("write", self.changed); self.body.bind("<<Modified>>", self.body_changed)

    def changed(self, *_a):
        if not self.loading and self.record.get("editable"):
            selected = {self.actions[n]["id"] for n in self.links.curselection()}
            self.dirty = (self.title.get() != self.record.get("name", "")
                          or self.body.get("1.0", "end").rstrip() != self.record.get("body", "").rstrip()
                          or selected != set(self.record.get("workflow_ids", [])))

    def body_changed(self, _e=None):
        if self.body.edit_modified(): self.changed(); self.body.edit_modified(False)

    def maybe_save(self):
        if not self.dirty: return True
        answer = messagebox.askyesnocancel("使用知识未保存", "保存当前使用知识？", parent=self.winfo_toplevel())
        if answer is None: return False
        if answer: return self.save()
        self.dirty = False; return True

    def load(self, app_id, record):
        if not self.maybe_save(): return False
        self.app_id, self.record, self.loading = app_id, record, True
        try:
            editable = record.get("editable", False)
            self.title_entry.configure(state="normal"); self.title.set(record.get("name", ""))
            self.body.configure(state="normal"); self.body.delete("1.0", "end"); self.body.insert("1.0", record.get("body", "")); self.body.edit_modified(False)
            self.actions = self.library.workflows(app_id); self.links.configure(state="normal"); self.links.delete(0, "end")
            for n, flow in enumerate(self.actions):
                self.links.insert("end", flow["keyword"])
                if flow["id"] in record.get("workflow_ids", []): self.links.selection_set(n)
            self.title_entry.configure(state="normal" if editable else "readonly")
            self.body.configure(state="normal" if editable else "disabled")
            self.links.configure(state="normal" if editable else "disabled")
            self.save_button.configure(state="normal" if editable else "disabled")
            self.status.set("说明以 Markdown 保存；关联项引用当前软件的组合动作 ID。" if editable else "现有页面文档供查阅。可新建使用知识补充操作说明。")
        finally: self.loading = self.dirty = False
        return True

    def new(self):
        if self.app_id: self.load(self.app_id, dict(name="新使用知识", body="", editable=True))

    def save(self):
        if not self.app_id or not self.record.get("editable"): return True
        try:
            record = self.library.save_usage(self.app_id, self.title.get(), self.body.get("1.0", "end").rstrip(), document_id=self.record.get("id", ""), workflow_ids=[self.actions[n]["id"] for n in self.links.curselection()])
            self.record, self.dirty = record, False; self.status.set("使用知识已保存。")
            self.on_saved(self.app_id, record["id"]); return True
        except (ValueError, OSError) as error:
            messagebox.showerror("保存失败", str(error), parent=self.winfo_toplevel()); return False

    def open_link(self, _event=None):
        selected = self.links.curselection()
        if selected and self.maybe_save(): self.open_action(self.app_id, self.actions[selected[0]]["id"])
