"""Combination action editor; the public class name preserves older callers."""
from __future__ import annotations

from copy import deepcopy
import json
import re
import tkinter as tk
from tkinter import messagebox, ttk
from uuid import uuid4

from . import workflows
from .gui_components import ScrollForm, TaskMonitor, WindowPicker, scroll_text, scroll_tree
from .library_tasks import LibraryTask
from .mcp.gui_icons import apply_window_icon

ACTIONS = {"点击": "click", "右键点击": "right-click", "输入文字": "set-text", "检查控件": "inspect", "待完善": "manual"}
CHECKS = {"控件存在": "exists", "文字等于": "text", "窗口标题等于": "window-title", "只读 / 可编辑": "read-only", "菜单包含": "menu-items", "待完善": "manual"}


class WorkflowWindow:
    def __init__(self, root, library, windows=None, app_id="", busy=None, *, embedded=False, on_saved=None):
        self.root, self.library = root, library
        self.windows = windows if windows is not None else {}
        self.embedded, self.on_saved = embedded, on_saved
        if not embedded:
            apply_window_icon(root)
            root.title("easy-uiauto · 组合动作"); root.geometry("1200x820"); root.minsize(960, 640)
            root.protocol("WM_DELETE_WINDOW", self.close)
        self.app_id, self.flow, self.selected_id, self.selected_target = app_id, None, None, None
        self.task = self.monitor = self.poll_id = None
        self.dirty = self.editor_dirty = self.loading = False
        self.busy = busy or (lambda: False)
        self.keyword, self.app_choice, self.precondition = (tk.StringVar(root) for _ in range(3))
        self.parameters = tk.StringVar(root, value="{}")
        self.status = tk.StringVar(root, value="引用控件库，按顺序执行动作并检查预期。")
        self.bound_window = tk.StringVar(root, value="未绑定运行窗口")
        self.fields = {k: tk.StringVar(root) for k in ("title", "type", "label", "text", "expected", "timeout", "hit", "expected_description")}
        self.build()
        for var in (self.keyword, self.precondition, self.parameters, *self.fields.values()): var.trace_add("write", self.changed)
        for text in (self.source, self.editor): text.bind("<<Modified>>", lambda _e, w=text: self.text_changed(w))
        self.notebook.bind("<<NotebookTabChanged>>", self.tab_changed)
        self.reload()
        self.poll_id = root.after(150, self.poll)

    def build(self):
        footer = ttk.Frame(self.root, padding=8)
        if not self.embedded: footer.pack(side="bottom", fill="x")
        ttk.Label(footer, textvariable=self.status, wraplength=800).pack(fill="x")
        bar = ttk.Frame(self.root, padding=8); bar.pack(fill="x")
        self.apps = {f"{r.get('name', r['id'])} [{r['id']}]": r["id"] for r in self.library.apps()}
        self.app_combo = ttk.Combobox(bar, textvariable=self.app_choice, values=list(self.apps), state="readonly", width=25)
        self.app_choice.set(next((n for n, i in self.apps.items() if i == self.app_id), next(iter(self.apps), "")))
        self.app_id = self.apps.get(self.app_choice.get(), self.app_id)
        if not self.embedded: self.app_combo.pack(side="left", padx=4)
        self.app_combo.bind("<<ComboboxSelected>>", lambda _e: self.change_app())
        self.buttons = []
        for text, callback in (("新建", self.new), ("保存组合动作", self.save), ("绑定运行窗口", self.bind_window), ("检查配置", self.check), ("运行组合动作", self.run), ("停止", self.stop)):
            button = ttk.Button(bar, text=text, command=callback); button.pack(side="left", padx=(0, 5)); self.buttons.append(button)
        self.buttons[-1].configure(state="disabled")
        row = ttk.Frame(self.root, padding=(8, 0)); row.pack(fill="x")
        ttk.Label(row, text="动作标题 / 关键字").pack(side="left")
        self.keyword_entry = ttk.Entry(row, textvariable=self.keyword); self.keyword_entry.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Label(self.root, textvariable=self.bound_window, padding=(8, 5)).pack(anchor="w")
        panes = ttk.Panedwindow(self.root, orient="horizontal"); panes.pack(fill="both", expand=True, padx=8)
        left = ttk.Frame(panes)
        if not self.embedded: panes.add(left, weight=2)
        self.tree = scroll_tree(left, columns=("status",), selectmode="browse")
        self.tree.heading("#0", text="组合动作 / 步骤 / 预期"); self.tree.heading("status", text="配置")
        self.tree.column("#0", width=280); self.tree.column("status", width=75, stretch=False)
        self.tree.bind("<<TreeviewSelect>>", self.select)
        right = ttk.Frame(panes); panes.add(right, weight=3)
        self.notebook = ttk.Notebook(right); self.notebook.pack(fill="both", expand=True)
        main, source, advanced, results = (ttk.Frame(self.notebook) for _ in range(4))
        for frame, text in ((main, "动作与预期"), (source, "从步骤表导入"), (advanced, "高级 JSON"), (results, "运行结果")): self.notebook.add(frame, text=text)
        form = ScrollForm(main); form.pack(fill="both", expand=True); body = form.body
        for text, var in (("前置条件", self.precondition), ("参数默认值（JSON；输入文字可用 ${参数名}）", self.parameters)):
            ttk.Label(body, text=text).pack(anchor="w"); ttk.Entry(body, textvariable=var).pack(fill="x", pady=(3, 8))
        row = ttk.Frame(body); row.pack(fill="x", pady=8)
        for text, cmd in (("添加动作", self.add_step), ("添加预期", self.add_check), ("独立检查", self.add_assertion), ("上移", lambda: self.move(-1)), ("下移", lambda: self.move(1)), ("删除步骤", self.remove)):
            ttk.Button(row, text=text, command=cmd).pack(side="left", padx=(0, 4))
        self.binding_label = ttk.Label(body, text="选择步骤编辑动作及预期。", wraplength=650); self.binding_label.pack(anchor="w", pady=8)
        self.field_widgets = {}
        for key, text in (("title", "步骤标题"), ("type", "动作 / 预期类型"), ("label", "目标说明"), ("binding", "引用当前软件的控件（保存 ID，不复制 XPath）"), ("text", "输入文字"), ("expected", "预期值（菜单项用逗号分隔；只读填 true / false）"), ("expected_description", "预期说明"), ("timeout", "等待预期超时（秒，0 至 60）"), ("hit", "点击位置")):
            ttk.Label(body, text=text, wraplength=650).pack(anchor="w")
            if key == "binding":
                self.binding = ttk.Combobox(body, state="readonly"); self.binding.pack(fill="x", pady=5)
                self.binding.bind("<<ComboboxSelected>>", lambda _e: self.bind_target())
                ttk.Button(body, text="清除控件引用", command=self.clear_target).pack(anchor="w", pady=3)
                continue
            field = ttk.Combobox(body, textvariable=self.fields[key], state="readonly") if key in {"type", "hit"} else ttk.Entry(body, textvariable=self.fields[key])
            if key == "hit": field["values"] = ("控件中心", "容器空白处")
            field.pack(fill="x", pady=(3, 8)); self.field_widgets[key] = field
        ttk.Label(body, text="按列表顺序执行；不会重复执行控件库里的前置步骤。", wraplength=650).pack(anchor="w", pady=10)
        self.source = scroll_text(source, wrap="word", undo=True, height=10, width=50)
        ttk.Label(source, text="导入表头：动作标题（兼容用例名称）、前置条件、步骤、预期。", padding=8).pack(anchor="w")
        ttk.Button(source, text="转换为新组合动作", command=self.convert).pack(anchor="e", padx=8, pady=8)
        ttk.Label(advanced, text="保存时应用当前 JSON；软件与组合动作 ID 不可更换。", padding=8).pack(anchor="w")
        self.editor = scroll_text(advanced, wrap="none", undo=True, width=50, height=10)
        self.result_summary = tk.StringVar(self.root, value="尚未运行")
        ttk.Label(results, textvariable=self.result_summary, padding=8, wraplength=850).pack(anchor="w")
        self.result_tree = scroll_tree(results, columns=("result",), height=5)
        self.result_tree.heading("#0", text="动作 / 预期"); self.result_tree.heading("result", text="结果")
        self.result_tree.column("#0", width=350); self.result_tree.column("result", width=80)
        self.result_tree.bind("<<TreeviewSelect>>", self.result_selected)
        row = ttk.Frame(results); row.pack(fill="x")
        for when, text in (("before", "查看执行前截图"), ("after", "查看执行后截图")): ttk.Button(row, text=text, command=lambda w=when: self.show_result_image(w)).pack(side="left", padx=8)
        ttk.Button(row, text="原始 JSON", command=self.show_raw_result).pack(side="left")
        self.result = scroll_text(results, wrap="word", width=50, height=8); self.result.configure(state="disabled")

    def _text(self, widget, value):
        widget.configure(state="normal"); widget.delete("1.0", "end"); widget.insert("1.0", value); widget.edit_modified(False)
        if widget is self.result: widget.configure(state="disabled")

    def changed(self, *_args):
        if not self.loading and self.flow: self.dirty = True

    keyword_changed = changed

    def text_changed(self, widget):
        if widget.edit_modified():
            if not self.loading:
                self.dirty = True
                if widget is self.editor: self.editor_dirty = True
            widget.edit_modified(False)

    def tab_changed(self, _event=None):
        if self.loading or self.task or not self.flow or self.notebook.index("current") != 2: return
        if not self.editor_dirty and not self.editor.edit_modified() and self.commit_form():
            try:
                self.flow.update(keyword=self.keyword.get().strip(), precondition=self.precondition.get(), parameters=json.loads(self.parameters.get()))
                self.sync_editor()
            except ValueError as error: self.status.set("参数格式无效：" + str(error))

    def maybe_save(self):
        if not self.dirty: return True
        answer = messagebox.askyesnocancel("组合动作未保存", "保存当前组合动作的修改？", parent=self.root.winfo_toplevel())
        if answer is None: return False
        if answer: return self.save()
        self.dirty = self.editor_dirty = False; return True

    def change_app(self):
        if self.task or not self.maybe_save():
            self.app_choice.set(next((n for n, i in self.apps.items() if i == self.app_id), "")); return
        self.app_id = self.apps.get(self.app_choice.get(), ""); self.flow = None; self.reload()

    def show_app(self, app_id):
        if self.app_id == app_id: return
        self.app_id, self.flow, self.selected_id = app_id, None, None
        self.loading = True
        self.keyword.set(""); self.precondition.set(""); self.parameters.set("{}")
        self._text(self.source, ""); self.sync_editor(); self.load_form()
        self.loading = False; self.dirty = False; self.reload(); self.render_results({})

    def reload(self, select_id=""):
        self.tree.delete(*self.tree.get_children()); self.nodes, self.controls, self.choices = {}, {}, {}
        if not self.app_id: return
        controls = self.library.read(self.app_id)["controls"]; self.controls = {r["id"]: r for r in controls}
        self.choices = {f"{r.get('semantic_name') or r.get('name') or r['id']} [{r['id']}]": r["id"] for r in controls if r.get("status") != "quarantined"}
        self.binding["values"] = list(self.choices)
        for flow in self.library.workflows(self.app_id): self.insert_flow(flow)
        window = self.windows.get(self.app_id); self.bound_window.set("运行窗口：" + window["title"] if window else "未绑定运行窗口")
        if select_id and self.tree.exists(select_id): self.tree.selection_set(select_id); self.tree.see(select_id); self.select()

    def insert_flow(self, flow):
        key = flow["id"]; report = workflows.readiness(flow, list(self.controls.values())); pending = {i["id"] for i in report["issues"]}
        self.tree.insert("", "end", iid=key, text=flow["keyword"], values=("配置完整" if report["ready"] else "待完善",), open=True); self.nodes[key] = (flow, None)
        for n, step in enumerate(flow["steps"], 1):
            node = key + ":" + step["id"]; self.nodes[node] = (flow, step["id"])
            self.tree.insert(key, "end", iid=node, text=f"{n}. {step.get('description') or step['action']}", values=("待完善" if step["id"] in pending else "已配置",), open=True)
            for check in step.get("checks", []):
                child = key + ":" + check["id"]; self.nodes[child] = (flow, check["id"])
                self.tree.insert(node, "end", iid=child, text="预期：" + (check.get("description") or check.get("target", {}).get("label") or check["kind"]), values=("待完善" if check["id"] in pending else "已配置",))

    def select(self, _event=None):
        selected = self.tree.selection()
        if selected and selected[0] in self.nodes:
            flow, target_id = self.nodes[selected[0]]; self.load(self.app_id, flow, target_id)

    def load(self, app_id, flow, target_id=None):
        if self.task: return False
        same = self.flow and self.flow.get("id") == flow.get("id") and self.app_id == app_id
        if same:
            if not self.commit_form(): return False
            self.sync_editor()
        else:
            if not self.maybe_save(): return False
            self.app_id, self.flow, self.selected_id = app_id, deepcopy(flow), None; self.reload()
            self.loading = True
            try:
                self.keyword.set(flow["keyword"]); self.precondition.set(flow.get("precondition", "")); self.parameters.set(json.dumps(flow.get("parameters", {}), ensure_ascii=False))
                self._text(self.source, flow.get("source_case", "")); self.sync_editor(); self.render_results(flow.get("last_run", {}))
            finally: self.loading = False
            self.dirty = False
        self.selected_id = target_id; self.load_form()
        return True

    def item(self):
        return next((i for s in (self.flow or {}).get("steps", []) for i in [s, *s.get("checks", [])] if i["id"] == self.selected_id), None)

    def sync_editor(self):
        self._text(self.editor, json.dumps(self.flow or {}, ensure_ascii=False, indent=2)); self.editor_dirty = False

    def load_form(self):
        self.loading = True
        try:
            item = self.item() or {}; is_check = "kind" in item; target = item.get("target", {}); self.selected_target = item.get("target")
            mapping = CHECKS if is_check else ACTIONS; self.field_widgets["type"]["values"] = list(mapping)
            self.binding_label.configure(text=("预期检查：" if is_check else "动作步骤：") + item.get("id", "请选择步骤或添加动作。"))
            values = dict(title=item.get("description", ""), type=next((k for k, v in mapping.items() if v == item.get("kind" if is_check else "action")), ""), label=target.get("label", ""), text=item.get("text", ""),
                          expected=json.dumps(item["expected"], ensure_ascii=False) if "expected" in item else "", timeout=str(item.get("timeout_seconds", 10)), hit="容器空白处" if target.get("hit") == "blank" else "控件中心", expected_description=item.get("expected_description", ""))
            for key, var in self.fields.items(): var.set(values[key])
            self.binding.set(next((n for n, i in self.choices.items() if i == target.get("control_id")), ""))
            for key, widget in self.field_widgets.items():
                enabled = bool(item) and (key not in {"text", "timeout", "hit", "expected_description"} if is_check else key != "expected")
                widget.configure(state=("readonly" if key in {"type", "hit"} else "normal") if enabled else "disabled")
        finally: self.loading = False

    def commit_form(self):
        if not self.flow: return True
        try:
            if self.editor_dirty or self.editor.edit_modified():
                flow = json.loads(self.editor.get("1.0", "end")); workflows.validate(flow)
                if flow.get("id") != self.flow.get("id") or flow.get("app_id") != self.app_id: raise ValueError("不能更换软件或组合动作 ID")
                self.flow = flow; self.loading = True
                self.keyword.set(flow["keyword"]); self.precondition.set(flow.get("precondition", "")); self.parameters.set(json.dumps(flow.get("parameters", {}), ensure_ascii=False))
                self.loading = False; self.editor_dirty = False; self.editor.edit_modified(False); self.load_form()
            item = self.item()
            if item:
                is_check = "kind" in item; mapping = CHECKS if is_check else ACTIONS
                item["description"] = self.fields["title"].get().strip(); item["kind" if is_check else "action"] = mapping[self.fields["type"].get()]
                target = item.setdefault("target", {"control_id": ""}); target["label"] = self.fields["label"].get()
                if is_check:
                    value = self.fields["expected"].get().strip()
                    if item["kind"] in {"text", "window-title"}:
                        try: parsed = json.loads(value)
                        except ValueError: parsed = value
                        item["expected"] = parsed if isinstance(parsed, str) else value
                    elif item["kind"] == "read-only": item["expected"] = json.loads(value.lower())
                    elif item["kind"] == "menu-items":
                        try: item["expected"] = json.loads(value)
                        except ValueError: item["expected"] = [v.strip() for v in re.split(r"[\n,，、]", value) if v.strip()]
                else:
                    item.update(text=self.fields["text"].get(), timeout_seconds=float(self.fields["timeout"].get()), expected_description=self.fields["expected_description"].get())
                    target["hit"] = "blank" if self.fields["hit"].get() == "容器空白处" else "center"
            return True
        except (ValueError, TypeError, KeyError) as error:
            self.loading = False; messagebox.showerror("步骤无效", str(error), parent=self.root.winfo_toplevel()); return False

    def bind_target(self):
        item = self.item(); identifier = self.choices.get(self.binding.get())
        if self.task or not item or not identifier: return
        target = item.setdefault("target", {}); target.update(control_id=identifier, control_type=self.controls[identifier].get("control_type", ""))
        for step in self.flow["steps"]:
            if step.get("assertion_only") and item in step.get("checks", []): step["target"] = deepcopy(target)
        if not self.fields["label"].get(): self.fields["label"].set(self.controls[identifier].get("semantic_name") or self.controls[identifier].get("name", ""))
        self.selected_target = target; self.dirty = True

    def clear_target(self):
        if self.task or not self.item(): return
        self.item().setdefault("target", {})["control_id"] = ""; self.binding.set(""); self.dirty = True

    def refresh_draft(self, target_id=None):
        self.selected_id = target_id; self.sync_editor(); self.load_form(); self.dirty = True
        self.tree.delete(*self.tree.get_children()); self.nodes = {}; self.insert_flow(deepcopy(self.flow))
        if self.on_saved: self.on_saved(self.app_id, self.flow["id"], draft=self.flow, target_id=target_id)

    def new(self):
        if self.task or not self.app_id or not self.maybe_save(): return
        self.flow = dict(id="workflow-" + uuid4().hex[:12], kind="workflow", schema_version=1, app_id=self.app_id, keyword="新组合动作", precondition="软件已启动", parameters={}, steps=[])
        self.loading = True; self.keyword.set("新组合动作"); self.precondition.set("软件已启动"); self.parameters.set("{}"); self._text(self.source, ""); self.loading = False
        self.selected_id = None; self.sync_editor(); self.add_step()

    def add_step(self):
        if self.task or not self.flow or not self.commit_form(): return
        identifier = "step-" + uuid4().hex[:8]
        self.flow["steps"].append(dict(id=identifier, action="click", description="新动作", target=dict(label="", control_id=""), checks=[], timeout_seconds=10)); self.refresh_draft(identifier)

    def add_check(self):
        if self.task or not self.flow or not self.commit_form(): return
        step = next((s for s in self.flow["steps"] if s["id"] == self.selected_id or any(c["id"] == self.selected_id for c in s.get("checks", []))), None)
        if step is None: self.status.set("先选择需要添加预期的动作步骤。"); return
        identifier = "check-" + uuid4().hex[:8]
        step.setdefault("checks", []).append(dict(id=identifier, kind="exists", description="控件存在", target=dict(label="", control_id=""))); self.refresh_draft(identifier)

    def add_assertion(self):
        if self.task or not self.flow or not self.commit_form(): return
        self.add_step()
        if not self.task and self.item():
            self.item().update(action="inspect", description="独立检查", assertion_only=True); self.refresh_draft(self.selected_id); self.add_check()

    def move(self, offset):
        if self.task or not self.flow or not self.commit_form(): return
        item = self.item()
        if not item: return
        sequence = next((s["checks"] for s in self.flow["steps"] if item in s.get("checks", [])), self.flow["steps"])
        index = sequence.index(item); other = index + offset
        if 0 <= other < len(sequence): sequence[index], sequence[other] = sequence[other], sequence[index]; self.refresh_draft(item["id"])

    def remove(self):
        if self.task or not self.flow or not self.commit_form(): return
        item = self.item()
        if not item: return
        sequence = next((s["checks"] for s in self.flow["steps"] if item in s.get("checks", [])), self.flow["steps"])
        if sequence is self.flow["steps"] and len(sequence) == 1: self.status.set("至少保留一个动作步骤。"); return
        sequence.remove(item); self.refresh_draft()

    def convert(self):
        if self.task or not self.app_id: return
        try:
            flow = self.library.import_workflow_case(self.app_id, self.source.get("1.0", "end").strip(), self.keyword.get().strip())
            self.flow = None; self.dirty = False; self.reload(flow["id"])
            self.status.set(f"已生成组合动作，共 {len(flow['steps'])} 步；待完善的控件引用需绑定。")
            if self.on_saved: self.on_saved(self.app_id, flow["id"])
        except (ValueError, OSError) as error: messagebox.showerror("导入失败", str(error), parent=self.root.winfo_toplevel())

    def apply_schema(self): self.editor_dirty = True; return self.commit_form()
    def rename(self): return self.save()

    def save(self):
        if self.task or not self.flow or not self.commit_form(): return False
        try:
            self.flow.update(keyword=self.keyword.get().strip(), precondition=self.precondition.get(), parameters=json.loads(self.parameters.get()), source_case=self.source.get("1.0", "end").strip())
            flow = self.library.save_workflow(self.app_id, self.flow); target_id = self.selected_id
            self.flow = None; self.dirty = self.editor_dirty = False; self.reload(); self.load(self.app_id, flow, target_id)
            self.status.set("组合动作已保存。")
            if self.on_saved: self.on_saved(self.app_id, flow["id"])
            return True
        except (ValueError, OSError, TypeError, KeyError) as error:
            messagebox.showerror("保存失败", str(error), parent=self.root.winfo_toplevel()); return False

    def check(self):
        if not self.flow or self.task or self.dirty and not self.save(): return
        report = workflows.readiness(self.flow, self.library.read(self.app_id)["controls"])
        self._text(self.result, "配置完整，可绑定窗口运行。" if report["ready"] else "\n".join(i["id"] + "：" + i["message"] for i in report["issues"]))
        self.notebook.select(3); self.status.set("配置完整。" if report["ready"] else f"还有 {len(report['issues'])} 项待完善；未执行操作。")

    def bind_window(self):
        if self.task or not self.app_id: return
        app_id = self.app_id
        def accept(window):
            self.windows[app_id] = window
            if self.app_id == app_id: self.bound_window.set("运行窗口：" + window["title"])
        return WindowPicker(self.root, app_id, accept)

    def run(self):
        if self.task or not self.flow: return
        if self.busy(): self.status.set("学习任务尚未结束，请等待完成或停止。"); return
        if self.dirty and not self.save(): return
        if not workflows.readiness(self.flow, self.library.read(self.app_id)["controls"])["ready"]: self.check(); return
        if self.app_id not in self.windows: self.status.set("请先绑定运行窗口。"); self.bind_window(); return
        self.running_id = self.flow["id"]; self.owner = self.root.winfo_toplevel()
        try:
            self.owner.iconify(); self.owner.update_idletasks()
            self.monitor = TaskMonitor(self.owner, None, "正在执行组合动作")
            self.task = LibraryTask(dict(operation="workflow", vault=str(self.library.root), app_id=self.app_id, flow_id=self.running_id, window=self.windows[self.app_id], capture_handshake=True), timeout=300)
            self.monitor.task = self.task
            for button in self.buttons: button.configure(state="disabled")
            self.buttons[-1].configure(state="normal"); self.app_combo.configure(state="disabled"); self.keyword_entry.configure(state="disabled")
        except Exception as error:
            if self.task: self.task.close(); self.task = None
            if self.monitor: self.monitor.close(); self.monitor = None
            self.owner.deiconify(); messagebox.showerror("无法运行", str(error), parent=self.owner)

    def stop(self):
        if self.task: self.task.cancel()

    def render_results(self, execution):
        self.execution = execution; self.result_tree.delete(*self.result_tree.get_children()); self.result_nodes = {}
        self.result_summary.set("需要重新验证：步骤或控件定位已变化" if execution.get("outdated") else "上次通过" if execution.get("ok") else "上次未通过" if execution else "尚未运行")
        for n, step in enumerate(execution.get("steps", []), 1):
            key = str(n); self.result_nodes[key] = step
            self.result_tree.insert("", "end", iid=key, text=step.get("description") or step.get("id", key), values=("通过" if step.get("ok") else "未通过",), open=True)
            for i, check in enumerate(step.get("checks", [])):
                child = f"{key}:{i}"; self.result_nodes[child] = {**check, "before": step.get("before"), "after": step.get("after")}
                self.result_tree.insert(key, "end", iid=child, text="预期：" + (check.get("description") or check.get("id", str(i))), values=("通过" if check.get("ok") else "未通过",))
        passed = sum(bool(s.get("ok")) for s in execution.get("steps", []))
        self._text(self.result, f"完成通过的步骤：{passed} / {len(execution.get('steps', []))}\n"
                   + ("已经尝试输入，请检查目标软件当前状态。\n" if execution.get("input_attempted") and not execution.get("ok") else "")
                   + execution.get("error", "选择一个动作或预期，查看实际结果及运行截图。"))

    def result_selected(self, _event=None):
        selected = self.result_tree.selection()
        if not selected: return
        item = self.result_nodes[selected[0]]
        lines = [item.get("description") or item.get("id", ""), "结果：" + ("通过" if item.get("ok") else "未通过")]
        for key, label in (("expected", "预期"), ("observed", "实际"), ("exists", "实际存在"), ("read_only", "实际只读"), ("enabled", "实际可用"), ("missing", "缺失项目"), ("error", "错误")):
            if key in item: lines.append(label + "：" + json.dumps(item[key], ensure_ascii=False))
        if "action" in item:
            lines.append("输入已尝试：" + ("是" if item.get("input_attempted") else "否"))
            lines.extend("预期：" + (c.get("description") or c["id"]) + (" — 通过" if c.get("ok") else " — 未通过") for c in item.get("checks", []))
        self._text(self.result, "\n".join(lines))

    def show_raw_result(self):
        window = tk.Toplevel(self.root); window.title("组合动作 · 原始运行结果"); window.geometry("850x650")
        apply_window_icon(window)
        text = scroll_text(window, wrap="none", width=70, height=25)
        text.insert("1.0", json.dumps(getattr(self, "execution", {}), ensure_ascii=False, indent=2)); text.configure(state="disabled")

    def show_result_image(self, when):
        selected = self.result_tree.selection(); evidence = self.result_nodes.get(selected[0], {}) if selected else {}
        observation = evidence.get(when) or (self.execution.get("failure_observation", {}) if when == "after" else {})
        path = self.library.asset(self.app_id, observation.get("context_image") or observation.get("image", ""))
        if path is None: self.status.set("所选步骤没有保存此截图。"); return
        from .knowledge_gui import RegionCanvas
        window = tk.Toplevel(self.root); window.title("运行截图 · " + ("执行前" if when == "before" else "执行后")); window.geometry("1000x700")
        apply_window_icon(window)
        bar = ttk.Frame(window, padding=8); bar.pack(fill="x")
        ttk.Label(bar, text="缩放").pack(side="left")
        zoom = ttk.Combobox(bar, state="readonly", width=8, values=("适应", "50%", "100%", "150%", "200%")); zoom.set("适应"); zoom.pack(side="left", padx=8)
        viewer = RegionCanvas(window, lambda _r: None, lambda _r: None); viewer.pack(fill="both", expand=True); viewer.enabled = False; viewer.set_image(path)
        zoom.bind("<<ComboboxSelected>>", lambda _e: viewer.zoom(zoom.get()))
        ttk.Label(window, text=str(path), padding=8).pack(fill="x")

    def poll(self):
        self.poll_id = None
        if self.task:
            state = self.task.poll()
            if state["worker_running"]:
                message = state.get("message", "正在执行组合动作"); self.status.set(message)
                if self.monitor: self.monitor.message.set(message)
            else:
                task, self.task = self.task, None
                try:
                    if self.monitor: self.monitor.close(); self.monitor = None
                    self.flow = None; self.reload(self.running_id)
                    self.render_results(state.get("result") or state.get("execution") or dict(ok=False, error=state.get("error", "任务未完成"), steps=[]))
                    self.notebook.select(3)
                    self.status.set("组合动作及全部预期通过。" if state["state"] == "completed" and state.get("result", {}).get("ok") else state.get("error", "组合动作未通过，请查看运行结果。"))
                    if self.on_saved: self.on_saved(self.app_id, self.running_id)
                finally:
                    task.close(); self.owner.deiconify(); self.app_combo.configure(state="readonly"); self.keyword_entry.configure(state="normal")
                    for button in self.buttons: button.configure(state="normal")
                    self.buttons[-1].configure(state="disabled")
        self.poll_id = self.root.after(150, self.poll)

    def close(self):
        if self.task:
            messagebox.showinfo("任务运行中", "请先停止组合动作，工作进程退出后再关闭。", parent=self.root.winfo_toplevel()); return
        if not self.maybe_save(): return
        if self.poll_id: self.root.after_cancel(self.poll_id)
        self.root.destroy()
