"""Software → manual regions → automatically learned controls, in one Tk GUI."""

from __future__ import annotations

import argparse
import json
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from .knowledge_library import KnowledgeLibrary, PAGE_ID
from .library_tasks import LibraryTask
from .gui_components import ScrollForm, TaskMonitor, scroll_text, scroll_tree
from .mcp import configuration, knowledge, native_windows
from .mcp import control_hierarchy, dynamic_regions
from .mcp.gui_icons import apply_window_icon

COLORS = ("#007c91", "#6655bb", "#007e42", "#b94a00", "#bb2255", "#006bad")
REFERENCE_VIEW = "学习参考图"
BEFORE_VIEW = "本次执行前"
RESULT_VIEW = "本次执行结果"


class RegionCanvas(ttk.Frame):
    """Map drawing coordinates through zoom and scrolling to original pixels."""
    def __init__(self, master, on_draw, on_region):
        super().__init__(master)
        self.canvas = tk.Canvas(self, background="#eef1f5", highlightthickness=0)
        vertical = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        horizontal = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.image = self.photo = None
        self.scale = 1.
        self.fit_mode = True
        self.regions, self.controls = [], []
        self.active_region = self.active_control = ""
        self.drawing = False
        self.enabled = True
        self.start = self.rubber = None
        self.on_draw, self.on_region = on_draw, on_region
        self.canvas.bind("<ButtonPress-1>", self.press)
        self.canvas.bind("<B1-Motion>", self.drag)
        self.canvas.bind("<ButtonRelease-1>", self.release)
        self.canvas.bind("<Configure>", lambda _event: self.fit() if self.fit_mode else None)

    def set_image(self, path: Path | None):
        self.image = None
        if path:
            try:
                with Image.open(path) as image:
                    self.image = image.convert("RGB")
            except (OSError, ValueError):
                pass
        self.fit_mode = True
        self.fit()

    def fit(self):
        if self.image:
            width, height = max(200, self.canvas.winfo_width()-12), max(150, self.canvas.winfo_height()-12)
            self.scale = min(width/self.image.width, height/self.image.height, 1.)
        self.render()

    def zoom(self, value):
        if value == "适应":
            self.fit_mode = True
            self.fit()
        else:
            self.fit_mode = False
            self.scale = float(value.rstrip("%"))/100
            self.render()

    def point(self, x, y):
        if not self.image:
            return 0, 0
        return (max(0, min(self.image.width, round(self.canvas.canvasx(x)/self.scale))),
                max(0, min(self.image.height, round(self.canvas.canvasy(y)/self.scale))))

    def render(self):
        self.canvas.delete("all")
        if not self.image:
            self.canvas.create_text(30, 35, text="1. 选择软件窗口\n2. 添加区域：悬浮拾取，滚轮切换父子，Ctrl 确认\n3. 开始学习，控件自动归档", anchor="nw", fill="#596578", font=("Microsoft YaHei UI", 13))
            return
        width, height = max(1, round(self.image.width*self.scale)), max(1, round(self.image.height*self.scale))
        self.photo = ImageTk.PhotoImage(self.image.resize((width, height), Image.Resampling.LANCZOS))
        self.canvas.create_image(0, 0, image=self.photo, anchor="nw")
        self.canvas.configure(scrollregion=(0, 0, width+2, height+2))
        for index, region in enumerate(self.regions):
            n = region.get("normalized_rect")
            if not n:
                continue
            x1, y1, x2, y2 = (n["left"]*width, n["top"]*height, n["right"]*width, n["bottom"]*height)
            color = COLORS[index % len(COLORS)]
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=color,
                                         width=2 if region["id"] == self.active_region else 1)
            self.canvas.create_text(x1+2, max(0, y1-2), text=f"{index+1} {region['name']}",
                                    anchor="sw" if y1 > 18 else "nw", fill=color,
                                    font=("Microsoft YaHei UI", 9))
        for control in self.controls:
            n = control.get("normalized_rect")
            if not n:
                continue
            selected = control["id"] == self.active_control
            self.canvas.create_rectangle(n["left"]*width, n["top"]*height,
                                         n["right"]*width, n["bottom"]*height,
                                         outline="#d62728" if selected else "#2387cc", width=2 if selected else 1)

    def press(self, event):
        if not self.enabled or not self.image:
            return
        self.start = self.point(event.x, event.y)
        if self.drawing:
            x, y = self.start
            self.rubber = self.canvas.create_rectangle(x*self.scale, y*self.scale,
                                                       x*self.scale, y*self.scale, outline="#d62728", width=1)
        else:
            x, y = self.start[0]/self.image.width, self.start[1]/self.image.height
            for region in reversed(self.regions):
                r = region.get("normalized_rect", {})
                if r and r["left"] <= x <= r["right"] and r["top"] <= y <= r["bottom"]:
                    self.on_region(region["id"])
                    break

    def drag(self, event):
        if self.rubber is not None and self.start:
            x, y = self.point(event.x, event.y)
            self.canvas.coords(self.rubber, self.start[0]*self.scale, self.start[1]*self.scale,
                               x*self.scale, y*self.scale)

    def release(self, event):
        if self.rubber is None or not self.start:
            return
        start, end = self.start, self.point(event.x, event.y)
        self.canvas.delete(self.rubber)
        self.rubber = self.start = None
        left, right = sorted((start[0], end[0])); top, bottom = sorted((start[1], end[1]))
        if right-left >= 2 and bottom-top >= 2:
            self.on_draw(dict(left=left/self.image.width, top=top/self.image.height,
                              right=right/self.image.width, bottom=bottom/self.image.height))


class WindowChooser(tk.Toplevel):
    def __init__(self, master, library: KnowledgeLibrary, current_app: str, callback):
        super().__init__(master)
        apply_window_icon(self)
        self.title("选择软件窗口")
        self.geometry("1000x570")
        self.minsize(700, 420)
        self.transient(master)
        self.library, self.callback = library, callback
        self.rows = {}
        self.search = tk.StringVar(self)
        search_row = ttk.Frame(self, padding=(16, 10)); search_row.pack(fill="x")
        ttk.Label(search_row, text="搜索窗口").pack(side="left")
        ttk.Entry(search_row, textvariable=self.search).pack(side="left", fill="x", expand=True, padx=10)
        self.search.trace_add("write", lambda *_a: self.filter_windows())
        self.tree = scroll_tree(self, columns=("title", "class"), show="headings", selectmode="browse")
        self.tree.heading("title", text="窗口标题"); self.tree.heading("class", text="窗口类名")
        self.tree.column("title", width=660); self.tree.column("class", width=280)
        self.tree.bind("<<TreeviewSelect>>", self.changed)
        form = ttk.Frame(self, padding=(16, 0, 16, 12)); form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        self.name = tk.StringVar(self)
        self.app_id = tk.StringVar(self, value=current_app)
        self.auto_name, self.auto_id = "", current_app
        self.rebinding = bool(current_app)
        for row, (label, variable) in enumerate((("软件名称", self.name), ("知识库 ID", self.app_id))):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
            field = ttk.Combobox(form, textvariable=variable,
                                 values=[r["id"] for r in library.apps()]) if row else ttk.Entry(form, textvariable=variable)
            field.grid(row=row, column=1, sticky="ew", pady=5)
        footer = ttk.Frame(self, padding=(16, 0, 16, 16)); footer.pack(fill="x")
        ttk.Label(footer, text="可选择已有知识库 ID 继续学习，也可填写新 ID。").pack(side="left")
        ttk.Button(footer, text="刷新窗口", command=self.refresh).pack(side="right")
        ttk.Button(footer, text="取消", command=self.destroy).pack(side="right", padx=8)
        ttk.Button(footer, text="选择此窗口", command=self.accept).pack(side="right", padx=10)
        self.refresh()

    def refresh(self):
        self.tree.delete(*self.tree.get_children()); self.rows.clear()
        try:
            for window in native_windows.enumerate_windows(visible_only=True):
                if not window["title"].strip() or window["bounds"]["width"] < 2:
                    continue
                if window["title"] in {self.title(), self.master.title()}:
                    continue
                key = str(window["handle"])
                self.rows[key] = window
                self.tree.insert("", "end", iid=key, values=(window["title"], window["class_name"]))
        except Exception as error:
            messagebox.showerror("无法列出软件", str(error), parent=self)
        self.filter_windows()

    def filter_windows(self):
        self.tree.delete(*self.tree.get_children())
        query = self.search.get().strip().casefold()
        for key, window in self.rows.items():
            if query in (window["title"] + " " + window["class_name"]).casefold():
                self.tree.insert("", "end", iid=key, values=(window["title"], window["class_name"]))

    def changed(self, _event=None):
        selection = self.tree.selection()
        if not selection:
            return
        window = self.rows[selection[0]]
        if self.name.get() in {"", self.auto_name}:
            existing = next((r for r in self.library.apps() if r["id"] == self.app_id.get()), None)
            self.auto_name = existing.get("name", window["title"]) if existing else window["title"]
            self.name.set(self.auto_name)
        if not self.rebinding and self.app_id.get() in {"", self.auto_id}:
            self.auto_id = knowledge.stable_id("app", window["class_name"], window["title"])
            self.app_id.set(self.auto_id)

    def accept(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("请选择软件", "先在列表中选中目标软件窗口。", parent=self)
            return
        try:
            window = self.rows[selection[0]]
            self.library.bind(self.app_id.get().strip(), self.name.get(), window)
        except (ValueError, OSError) as error:
            messagebox.showerror("无法选择", str(error), parent=self)
            return
        app_id = self.app_id.get().strip()
        self.destroy()
        self.callback(app_id, window)


class KnowledgeWindow:
    def __init__(self, root: tk.Tk, vault: str | Path | None = None):
        self.root, self.library = root, KnowledgeLibrary(vault)
        apply_window_icon(root)
        self.root.title("easy-uiauto 知识库管理")
        self.root.geometry("1440x900"); self.root.minsize(1120, 720)
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.windows, self.indexes, self.nodes = {}, {}, {}
        self.hierarchies = {}
        self.region_trees = {}
        self.draw_parent_region = ""
        self.app_id, self.selected = "", None
        self.display_page_id = PAGE_ID
        self.task, self.task_operation, self.task_app = None, "", ""
        self.closed = self.loading = self.dirty = False
        self.poll_id = None
        self.preview_photo = None
        self.status = tk.StringVar(root, value="选择软件窗口 → 添加区域（悬浮拾取，Ctrl 确认）→ 开始学习。")
        self.path = tk.StringVar(root, value=str(self.library.root))
        self.name, self.description, self.aliases = (tk.StringVar(root) for _ in range(3))
        self.mode = tk.StringVar(root, value="identity")
        self.parent_selection = tk.StringVar(root)
        self.rect_values = {key: tk.StringVar(root) for key in ("left", "top", "right", "bottom")}
        self.draw_mode = tk.BooleanVar(root, value=False)
        self.enrich = tk.BooleanVar(root, value=False)
        self.show_controls = tk.BooleanVar(root, value=True)
        self.navigation_ids = []
        self.display_observation = {}
        self.display_after = False
        self.observation_mode = tk.StringVar(root, value=REFERENCE_VIEW)
        self.observation_control = None
        self._build()
        for variable in (self.name, self.description, self.aliases, self.mode, self.parent_selection, *self.rect_values.values()):
            variable.trace_add("write", self.changed)
        self.reload()
        self.poll_id = root.after(150, self.poll)

    def _build(self):
        style = ttk.Style(self.root)
        if "vista" in style.theme_names(): style.theme_use("vista")
        style.configure("Knowledge.TButton", padding=(6, 4))
        style.configure("Knowledge.Treeview", rowheight=27)
        # Reserve the status and stop area before the expanding workspace.
        footer = ttk.Frame(self.root, padding=(12, 8)); footer.pack(side="bottom", fill="x")
        self.stop_button = ttk.Button(footer, text="停止当前任务", command=self.stop, state="disabled")
        self.stop_button.pack(side="right")
        self.progress = ttk.Progressbar(footer, mode="determinate", length=100); self.progress.pack(side="right", padx=10)
        self.status_label = ttk.Label(footer, textvariable=self.status, wraplength=800)
        self.status_label.pack(side="left", fill="x", expand=True)
        footer.bind("<Configure>", lambda e: self.status_label.configure(wraplength=max(200, e.width-270)))
        top = ttk.Frame(self.root, padding=(12, 10, 12, 5)); top.pack(fill="x")
        ttk.Label(top, text="知识库路径").pack(side="left", padx=(0, 10))
        path_entry = ttk.Entry(top, textvariable=self.path); path_entry.pack(side="left", fill="x", expand=True)
        path_entry.bind("<Return>", lambda _e: self.change_vault(use_typed=True))
        self.path_button = ttk.Button(top, text="切换知识库…", command=self.change_vault); self.path_button.pack(side="left", padx=8)
        ttk.Button(top, text="配置", command=self.configure).pack(side="left")
        self.buttons = {}
        for entries in (
            (("1  选择软件", "选择软件窗口", self.choose_window), ("刷新截图", "刷新截图", self.capture),
             ("添加区域", "添加区域", self.add_region), ("修改区域", "修改区域", self.modify_region), ("3  开始学习", "开始学习", self.learn)),
            (("保存修改", "保存当前节点", self.save), ("删除所选", "删除所选", self.remove), ("验证定位", "验证定位", self.verify),
             ("点击", "执行点击", self.click), ("学习展开控件", "点击并学习展开内容", self.expand), ("导出", "导出当前软件（JSON）", self.export))):
            bar = ttk.Frame(self.root, padding=(12, 4)); bar.pack(fill="x")
            for key, label, callback in entries:
                button = ttk.Button(bar, text=label, command=callback, style="Knowledge.TButton")
                button.pack(side="left", padx=(0, 6)); self.buttons[key] = button
        options = ttk.Frame(self.root, padding=(12, 4, 12, 8)); options.pack(fill="x")
        self.draw_button = ttk.Checkbutton(options, text="手绘区域（备用）", variable=self.draw_mode, command=self.toggle_drawing)
        self.draw_button.pack(side="right", padx=6)
        ttk.Checkbutton(options, text="补充控件含义（API）", variable=self.enrich).pack(side="left")
        ttk.Checkbutton(options, text="显示控件框", variable=self.show_controls, command=self.overlays).pack(side="left", padx=12)
        ttk.Label(options, text="缩放").pack(side="left")
        self.zoom = ttk.Combobox(options, state="readonly", width=7, values=("适应", "50%", "100%", "150%", "200%")); self.zoom.set("适应"); self.zoom.pack(side="left", padx=6)
        self.zoom.bind("<<ComboboxSelected>>", lambda _e: self.viewer.zoom(self.zoom.get()))
        ttk.Label(options, text="画面").pack(side="left", padx=(10, 0))
        choice = ttk.Combobox(options, textvariable=self.observation_mode, state="readonly", width=17, values=(REFERENCE_VIEW, BEFORE_VIEW, RESULT_VIEW)); choice.pack(side="left", padx=6)
        choice.bind("<<ComboboxSelected>>", lambda _e: self.show_observation())
        splitter = ttk.Panedwindow(self.root, orient="horizontal"); splitter.pack(fill="both", expand=True, padx=12)
        tree_frame = ttk.Frame(splitter, width=300); splitter.add(tree_frame, weight=1)
        self.search = tk.StringVar(self.root)
        ttk.Label(tree_frame, text="知识库 / 软件 / 使用知识 · 控件库 · 组合动作").pack(anchor="w", pady=5)
        search = ttk.Entry(tree_frame, textvariable=self.search); search.pack(fill="x", pady=(0, 5))
        search.bind("<Return>", lambda _e: self.search_nodes())
        self.tree = scroll_tree(tree_frame, columns=("status",), selectmode="browse", style="Knowledge.Treeview")
        self.tree.heading("#0", text="知识节点"); self.tree.heading("status", text="状态")
        self.tree.column("#0", width=245); self.tree.column("status", width=80, stretch=False)
        self.tree.bind("<<TreeviewSelect>>", self.select)
        self.workspace = ttk.Frame(splitter); splitter.add(self.workspace, weight=4)
        self.control_workspace = ttk.Panedwindow(self.workspace, orient="horizontal")
        self.control_workspace.pack(fill="both", expand=True)
        self.viewer = RegionCanvas(self.control_workspace, self.drawn, self.select_region); self.control_workspace.add(self.viewer, weight=3)
        details = ttk.Frame(self.control_workspace, padding=(8, 0, 0, 0)); self.control_workspace.add(details, weight=2)
        self.detail_title = ttk.Label(details, text="未选择节点", font=("Microsoft YaHei UI", 12, "bold")); self.detail_title.pack(anchor="w", pady=8)
        notebook = ttk.Notebook(details); notebook.pack(fill="both", expand=True)
        scroll_form = ScrollForm(notebook); form = scroll_form.body
        locator, pictures, flow = (ttk.Frame(notebook, padding=8) for _ in range(3))
        for frame, label in ((scroll_form, "属性"), (locator, "LOCATION / 代码"), (pictures, "学习图片"), (flow, "前置步骤")): notebook.add(frame, text=label)
        self.property_fields = []
        for text, variable in (("名称 / 控件别名", self.name), ("说明", self.description), ("检索别名（逗号分隔）", self.aliases)):
            ttk.Label(form, text=text).pack(anchor="w", pady=(8, 3)); entry = ttk.Entry(form, textvariable=variable); entry.pack(fill="x"); self.property_fields.append(entry)
        ttk.Label(form, text="上级控件").pack(anchor="w", pady=(12, 3))
        self.parent_choice = ttk.Combobox(form, textvariable=self.parent_selection, state="readonly"); self.parent_choice.pack(fill="x")
        self.region_button = ttk.Button(form, text="将展开容器设为子区域", command=self.make_region); self.region_button.pack(anchor="w", pady=8)
        self.rect_title = ttk.Label(form, text="区域边界（原图像素）"); self.rect_title.pack(anchor="w", pady=(12, 4))
        self.rect_fields = {}
        row = ttk.Frame(form); row.pack(fill="x")
        for n, (label, key) in enumerate((("左", "left"), ("上", "top"), ("右", "right"), ("下", "bottom"))):
            ttk.Label(row, text=label).grid(row=n//2, column=(n%2)*2, padx=4, pady=4)
            entry = ttk.Entry(row, textvariable=self.rect_values[key], width=9); entry.grid(row=n//2, column=(n%2)*2+1); self.rect_fields[key] = entry
        ttk.Label(form, text="定位模式：identity 按身份取首项；position 按同类序号", wraplength=290).pack(anchor="w", pady=(12, 4))
        self.mode_choice = ttk.Combobox(form, textvariable=self.mode, state="readonly", values=("identity", "position")); self.mode_choice.pack(fill="x")
        self.attributes = tk.Text(form, height=10, width=30, wrap="word", relief="flat", background="#f5f7fa"); self.attributes.pack(fill="both", expand=True, pady=12)
        ttk.Label(locator, text="修改 LOCATION 后需重新验证。", wraplength=280).pack(anchor="w")
        self.location = scroll_text(locator, height=8, width=30, wrap="none", undo=True); self.location.bind("<<Modified>>", self.location_changed)
        ttk.Label(locator, text="Python 获取代码").pack(anchor="w")
        self.code = scroll_text(locator, height=6, width=30, wrap="none")
        ttk.Label(pictures, text="学习时保存的完整控件图", padding=(0, 8)).pack(anchor="w")
        self.picture_label = ttk.Label(pictures, text="尚无图片", anchor="center", relief="solid", borderwidth=1); self.picture_label.pack(fill="both", expand=True)
        self.picture_path = ttk.Label(pictures, text="", wraplength=280); self.picture_path.pack(fill="x", pady=8)
        ttk.Label(flow, text="单控件点击会依次执行入口；组合动作使用自己的显式步骤。", wraplength=290).pack(anchor="w")
        self.navigation_list = tk.Listbox(flow, height=7, exportselection=False); self.navigation_list.pack(fill="both", expand=True, pady=8)
        self.navigation_choice = ttk.Combobox(flow, state="readonly"); self.navigation_choice.pack(fill="x", pady=5)
        row = ttk.Frame(flow); row.pack(fill="x")
        for text, callback in (("添加", self.add_navigation), ("上移", lambda: self.move_navigation(-1)), ("下移", lambda: self.move_navigation(1)), ("移除", self.remove_navigation)):
            ttk.Button(row, text=text, command=callback, width=5).pack(side="left", padx=2)
        self.navigation_summary = ttk.Label(flow, text="", wraplength=290); self.navigation_summary.pack(fill="x", pady=10)
        from .usage_gui import UsagePanel
        from .workflow_gui import WorkflowWindow
        self.usage_panel = UsagePanel(self.workspace, self.library, self.usage_saved, self.open_action)
        self.combination_frame = ttk.Frame(self.workspace)
        self.workflow_panel = WorkflowWindow(self.combination_frame, self.library, self.windows, embedded=True, busy=lambda: bool(self.task), on_saved=self.combination_saved)
        self.monitor = None
        self.root.after_idle(lambda: splitter.sashpos(0, 325))

    def show_workspace(self, name):
        for frame in (self.control_workspace, self.usage_panel, self.combination_frame): frame.pack_forget()
        {"controls": self.control_workspace, "usage": self.usage_panel, "actions": self.combination_frame}[name].pack(fill="both", expand=True)

    def update_actions(self, kind):
        for key, button in self.buttons.items():
            enabled = bool(self.app_id)
            if key == "1  选择软件": enabled = True
            elif key in {"验证定位", "点击", "学习展开控件"}: enabled = kind == "control"
            elif key == "删除所选": enabled = kind in {"region", "control"}
            elif key == "修改区域": enabled = kind == "region" and self.selected is not None and not dynamic_regions.is_dynamic(self.selected[2]) and not self.selected[2].get("classification_only")
            elif key == "保存修改": enabled = kind in {"application", "region", "control", "controls-root", "usage", "combination"}
            button.configure(state="normal" if enabled else "disabled")

    def add_action_nodes(self, app_id, flow):
        from . import workflows
        key = f"workflow:{app_id}:{flow['id']}"
        if self.tree.exists(key):
            for node in [key, *self._descendants(key)]: self.nodes.pop(node, None)
            self.tree.delete(key)
        report = workflows.readiness(flow, self.indexes[app_id]["controls"], workflows.load_catalog(self.library.directory(app_id)))
        status = "待完善" if not report["ready"] else "需重验" if flow.get("last_run", {}).get("outdated") else "配置完整"
        self.tree.insert(f"actions-root:{app_id}", "end", iid=key, text=flow["keyword"], values=(status,), open=True)
        self.nodes[key] = (app_id, "combination", {"flow": flow, "target_id": None})
        for n, step in enumerate(flow["steps"], 1):
            node = key + ":" + step["id"]
            self.tree.insert(key, "end", iid=node, text=f"{n}. {step.get('description') or step['action']}", open=True)
            self.nodes[node] = (app_id, "combination", {"flow": flow, "target_id": step["id"]})
            for check in step.get("checks", []):
                child = key + ":" + check["id"]
                self.tree.insert(node, "end", iid=child, text="预期：" + (check.get("description") or check.get("target", {}).get("label") or check["kind"]))
                self.nodes[child] = (app_id, "combination", {"flow": flow, "target_id": check["id"]})

    def _descendants(self, key):
        result, pending = [], list(self.tree.get_children(key))
        while pending:
            item = pending.pop(); result.append(item); pending.extend(self.tree.get_children(item))
        return result

    def combination_saved(self, app_id, flow_id, *, draft=None, target_id=None):
        flow = draft or next((f for f in self.library.workflows(app_id) if f["id"] == flow_id), None)
        if flow is None: return
        self.loading = True
        try:
            self.add_action_nodes(app_id, flow)
            key = f"workflow:{app_id}:{flow_id}" + (":" + target_id if target_id else "")
            self.tree.selection_set(key); self.tree.see(key); self.selected = self.nodes[key]
        finally: self.loading = False

    def usage_saved(self, app_id, document_id):
        self.reload(f"usage:{app_id}:usage/{document_id}.md")

    def open_action(self, app_id, flow_id):
        key = f"workflow:{app_id}:{flow_id}"
        if self.tree.exists(key): self.tree.selection_set(key); self.tree.see(key); self.select()

    def search_nodes(self):
        query = self.search.get().strip().casefold()
        if not query: return
        keys = [k for k in self.nodes if query in self.tree.item(k, "text").casefold()]
        if not keys: self.status.set("没有匹配的知识节点。"); return
        current = self.tree.selection()
        index = (keys.index(current[0])+1) % len(keys) if current and current[0] in keys else 0
        self.tree.selection_set(keys[index]); self.tree.see(keys[index]); self.select()
        self.status.set(f"匹配 {len(keys)} 项；按回车查看下一个。")

    def changed(self, *_args):
        if not self.loading: self.dirty = True

    def location_changed(self, _event=None):
        if self.location.edit_modified():
            self.changed(); self.location.edit_modified(False)

    def _set_text(self, widget, value):
        widget.configure(state="normal"); widget.delete("1.0", "end"); widget.insert("1.0", value)

    def reload(self, select_key: str = ""):
        self.loading = True
        self.tree.delete(*self.tree.get_children()); self.nodes.clear(); self.indexes.clear(); self.hierarchies.clear(); self.region_trees.clear()
        try:
            self.tree.insert("", "end", iid="vault", text="知识库", open=True)
            self.nodes["vault"] = ("", "vault", {})
            for app in self.library.apps():
                app_id = app["id"]
                self.library.sync_dynamic_regions(app_id)
                index = self.library.read(app_id); self.indexes[app_id] = index
                app_key = f"app:{app_id}"
                self.tree.insert("vault", "end", iid=app_key, text=index["application"].get("name", app_id), open=app_id == self.app_id)
                self.nodes[app_key] = (app_id, "application", index["application"])
                for kind, label in (("usage-root", "使用知识"), ("controls-root", "控件库"), ("actions-root", "组合动作")):
                    category = f"{kind}:{app_id}"
                    self.tree.insert(app_key, "end", iid=category, text=label, open=kind == "controls-root")
                    self.nodes[category] = (app_id, kind, {})
                for document in self.library.usage_documents(app_id):
                    node = f"usage:{app_id}:{document['document_path']}"
                    self.tree.insert(f"usage-root:{app_id}", "end", iid=node, text=document["name"], values=("说明" if document["editable"] else "页面文档",))
                    self.nodes[node] = (app_id, "usage", document)
                for flow in index["workflows"]: self.add_action_nodes(app_id, flow)
                regions = {r["id"]: r for r in index["regions"]}
                hierarchy = control_hierarchy.build(index["controls"])
                self.hierarchies[app_id] = hierarchy
                graph = dynamic_regions.build_tree(index["controls"], index["regions"])
                self.region_trees[app_id] = graph
                records = {c["id"]: c for c in index["controls"]}
                region_pages = {control_hierarchy.region_key(c): c.get("page_id", "unknown") for c in index["controls"]}
                groups = {}
                for control in index["controls"]:
                    full = hierarchy["regions"][control["id"]]
                    page = region_pages[full]
                    if full not in regions:
                        regions[full] = {"id": full, "name": "待归档", "page_id": page, "classification_only": True}
                    groups.setdefault(full, []).append(control)
                display = {"region:"+i: ("region", r) for i, r in regions.items()
                           if not dynamic_regions.is_dynamic(r) or i in graph["dynamic"]}
                display.update({"control:"+i: ("control", r) for i, r in records.items() if i not in graph["containers"]})
                children = {}
                for node in display:
                    parent = graph["parents"].get(node)
                    children.setdefault(parent if parent in display else None, []).append(node)
                sort_key = lambda node: (display[node][1].get("classification_only", False),
                                        display[node][1].get("semantic_name") or display[node][1].get("name") or node, node)
                pending = [(node, f"controls-root:{app_id}") for node in reversed(sorted(children.get(None, []), key=sort_key))]
                while pending:
                    node, parent_key = pending.pop()
                    kind, record = display[node]
                    key = f"{kind}:{app_id}:{record['id']}"
                    name = record.get("semantic_name") or record.get("name") or record["id"]
                    if kind == "region":
                        status = "动态区域" if dynamic_regions.is_dynamic(record) else ""
                        name += "〔区域〕" if dynamic_regions.is_dynamic(record) else f"  ({len(groups.get(record['id'], []))})"
                    else:
                        status = {"verified": "已验证", "observed": "已采集", "quarantined": "隔离", "suspect": "待检查"}.get(record.get("status"), record.get("status", ""))
                    self.tree.insert(parent_key, "end", iid=key, text=name, values=(status,))
                    self.nodes[key] = (app_id, kind, record)
                    pending.extend((child, key) for child in reversed(sorted(children.get(node, []), key=sort_key)))
            key = select_key or (f"app:{self.app_id}" if self.app_id else "vault")
            if key and self.tree.exists(key):
                self.tree.selection_set(key); self.tree.see(key)
        finally:
            self.loading = False; self.dirty = False
        self.select()

    def _maybe_save(self):
        if not self.usage_panel.maybe_save() or not self.workflow_panel.maybe_save(): return False
        if not self.dirty: return True
        answer = messagebox.askyesnocancel("尚未保存", "保存当前节点的修改？", parent=self.root)
        if answer is None: return False
        if answer:
            return self.save()
        self.dirty = False
        return True

    def select(self, _event=None):
        if self.loading or self.task or self.workflow_panel.task: return
        selection = self.tree.selection()
        if not selection or selection[0] not in self.nodes: return
        next_node = self.nodes[selection[0]]
        same_flow = (next_node[1] == "combination" and self.workflow_panel.flow
                     and next_node[0] == self.workflow_panel.app_id
                     and next_node[2]["flow"]["id"] == self.workflow_panel.flow["id"])
        if same_flow:
            if not self.workflow_panel.commit_form(): return
        elif not self._maybe_save():
            if self.selected:
                previous = next((key for key, node in self.nodes.items() if node == self.selected), "")
                if previous: self.loading = True; self.tree.selection_set(previous); self.loading = False
            return
        # Saving another editor may rebuild the tree; keep the requested node.
        if self.tree.exists(selection[0]):
            if self.tree.selection() != selection:
                self.loading = True; self.tree.selection_set(selection[0]); self.loading = False
            next_node = self.nodes[selection[0]]
        if self.draw_parent_region and (next_node[1] != "region" or next_node[2]["id"] != self.draw_parent_region):
            self.draw_parent_region = ""; self.draw_mode.set(False); self.viewer.drawing = False
            self.viewer.canvas.configure(cursor="")
        self.selected = next_node
        self.app_id, kind, record = self.selected
        self.update_actions(kind)
        if kind in {"usage-root", "usage", "actions-root", "combination", "vault"}:
            self.show_workspace("actions" if kind in {"actions-root", "combination"} else "usage")
            if kind == "combination":
                self.workflow_panel.load(self.app_id, record["flow"], record.get("target_id"))
            elif kind == "actions-root":
                self.workflow_panel.show_app(self.app_id)
            elif kind != "vault":
                self.usage_panel.load(self.app_id, record if kind == "usage" else dict(name="使用知识", body="在这里维护软件操作说明，可关联组合动作。", editable=False))
            else:
                self.usage_panel.loading = True; self.usage_panel.record = {}; self.usage_panel.app_id = ""; self.usage_panel.title.set("知识库")
                self.usage_panel.body.configure(state="normal"); self.usage_panel.body.delete("1.0", "end")
                self.usage_panel.body.insert("1.0", "选择左侧软件，查看使用知识、管理控件库或编排组合动作。")
                self.usage_panel.body.edit_modified(False); self.usage_panel.body.configure(state="disabled")
                self.usage_panel.loading = self.usage_panel.dirty = False
                self.usage_panel.save_button.configure(state="disabled"); self.usage_panel.title_entry.configure(state="readonly")
                self.usage_panel.links.configure(state="normal"); self.usage_panel.links.delete(0, "end"); self.usage_panel.links.configure(state="disabled")
            return
        self.show_workspace("controls")
        if kind == "controls-root": kind, record = "application", self.indexes[self.app_id]["application"]
        self.loading = True
        try:
            self.name.set(record.get("semantic_name") or record.get("name", ""))
            self.description.set(record.get("description", ""))
            self.aliases.set(", ".join(record.get("aliases", [])))
            location = record.get("location", {})
            params = location.get("PARAMETERS") or {}
            if isinstance(params, str): params = json.loads(params)
            self.mode.set(params.get("match_mode", "identity"))
            self.loaded_mode = self.mode.get()
            self._set_text(self.location, json.dumps(location, ensure_ascii=False, indent=2) if location else "")
            self.location.edit_modified(False)
            self._set_text(self.code, self.library.python_code(record) if kind == "control" else "")
            self.code.configure(state="disabled")
            self.navigation_ids = list(record.get("navigation_path", [])) if kind == "control" or dynamic_regions.is_dynamic(record) else []
            self.render_navigation()
            self.parent_options = {"自动（入口 / UIA 父子）": "auto", "区域直属": ""}
            for item in self.indexes[self.app_id]["controls"]:
                if kind == "control" and item["id"] != record["id"]:
                    label = f"{item.get('semantic_name') or item.get('name') or item['id']}  [{item['id']}]"
                    self.parent_options[label] = item["id"]
            self.loaded_parent = record.get("hierarchy_parent", "auto") if kind == "control" else "auto"
            if self.loaded_parent not in self.parent_options.values():
                self.parent_options[f"上级已删除 [{self.loaded_parent}]"] = self.loaded_parent
            self.parent_choice["values"] = list(self.parent_options)
            self.parent_selection.set(next(label for label, identifier in self.parent_options.items() if identifier == self.loaded_parent))
            self.detail_title.configure(text="动态子区域" if kind == "region" and dynamic_regions.is_dynamic(record)
                                        else {"application": "软件", "region": "区域", "control": "控件"}[kind])
            editable_rect = kind == "region" and record.get("selection_source") != "uia-picker" and (dynamic_regions.is_dynamic(record) or
                record.get("source") == "manual" and record.get("page_id") == PAGE_ID and not record.get("classification_only"))
            for field in self.rect_fields.values(): field.configure(state="normal" if editable_rect else "readonly")
            self.rect_title.configure(text="区域边界（可编辑，原图像素）" if editable_rect else "区域边界（点击“修改区域”重新拾取）" if kind == "region" else "控件边界（只读，原图像素）")
            self.mode_choice.configure(state="readonly" if kind == "control" else "disabled")
            self.parent_choice.configure(state="readonly" if kind == "control" else "disabled")
            self.region_button.configure(state="normal" if kind == "control" else "disabled")
            self.property_fields[2].configure(state="normal" if kind == "control" else "disabled")
            attrs = {key: record.get(key, "") for key in ("id", "name", "automation_id", "class_name", "control_type", "status", "semantic_source", "region_assignment")}
            if kind == "control":
                hierarchy = self.hierarchies[self.app_id]
                attrs.update(hierarchy_parent=hierarchy["parents"][record["id"]],
                             hierarchy_source=hierarchy["sources"][record["id"]],
                             tree_region=hierarchy["regions"][record["id"]], geometry_region=record.get("region_id"))
            elif dynamic_regions.is_dynamic(record):
                attrs.update({k: record.get(k, "") for k in ("source", "source_control_id", "parent_control_id", "parent_region_id", "navigation_path")})
            elif kind == "region":
                attrs.update({k: record.get(k, "") for k in ("layout_origin", "region_evidence", "edited_by_user")})
            self._set_text(self.attributes, json.dumps(attrs, ensure_ascii=False, indent=2)); self.attributes.configure(state="disabled")
            page_id = record.get("page_id", PAGE_ID)
            self.display_page_id = page_id
            if page_id != PAGE_ID:
                self.draw_mode.set(False)
                self.viewer.drawing = False
                self.viewer.canvas.configure(cursor="")
            index = self.indexes[self.app_id]
            page = next((p for p in index["pages"] if p["id"] == page_id), {})
            control_key = (self.app_id, record["id"]) if kind == "control" or dynamic_regions.is_dynamic(record) else None
            if control_key != self.observation_control:
                self.observation_mode.set(REFERENCE_VIEW)
            self.observation_control = control_key
            self.show_observation()
            self.viewer.active_region = record["id"] if kind == "region" else f"{page_id}.{record.get('region_id', '')}"
            self.viewer.active_control = record["id"] if kind == "control" else ""
            n = self.display_observation.get("normalized_rect") or record.get("normalized_rect") or {}
            for key, value in self.rect_values.items():
                dimension = self.viewer.image.width if self.viewer.image and key in {"left", "right"} else self.viewer.image.height if self.viewer.image else 0
                value.set(str(round(n[key]*dimension)) if key in n else "")
            image_path = self.library.asset(self.app_id, record.get("image", ""))
            self.preview_photo = None
            if image_path:
                try:
                    with Image.open(image_path) as source:
                        preview = source.convert("RGB")
                    preview.thumbnail((320, 480))
                    self.preview_photo = ImageTk.PhotoImage(preview)
                    self.picture_label.configure(image=self.preview_photo, text="")
                except (OSError, ValueError):
                    self.picture_label.configure(image="", text="图片无法读取，控件记录仍可查看")
            else:
                self.picture_label.configure(image="", text="尚无图片")
            self.picture_path.configure(text=record.get("image", "") + (f"\n原图：{source.width} × {source.height} 像素" if image_path and self.preview_photo else ""))
            self.overlays()
        finally:
            self.loading = False; self.dirty = False

    def show_observation(self):
        if not self.selected: return
        app_id, kind, record = self.selected
        index = self.indexes[app_id]
        page_id = record.get("page_id", PAGE_ID)
        page = next((p for p in index["pages"] if p["id"] == page_id), {})
        observation = record.get("visual_observation", {}) if kind == "control" or dynamic_regions.is_dynamic(record) else {}
        self.display_after = False
        mode = {"找到时 / 点击前": REFERENCE_VIEW, "点击后结果（即时）": RESULT_VIEW}.get(
            self.observation_mode.get(), self.observation_mode.get())
        if kind != "control" and mode != REFERENCE_VIEW:
            mode = REFERENCE_VIEW
            self.status.set("区域显示学习参考图；查看运行前后画面请选中操作控件。")
        if mode != self.observation_mode.get(): self.observation_mode.set(mode)
        runtime = kind == "control" and mode in {BEFORE_VIEW, RESULT_VIEW}
        if runtime:
            execution = record.get("last_click_execution", {})
            step = next((s for s in reversed(execution.get("steps", []))
                         if s.get("control_id") == record["id"]), {})
            observation = step.get("after" if mode == RESULT_VIEW else "before", {})
            if mode == RESULT_VIEW and execution.get("failure_observation"):
                observation = execution["failure_observation"]
            self.display_after = mode == RESULT_VIEW
            if not observation:
                detail = execution.get("failure_capture_error") or step.get("after_capture_error", "尚无对应的运行截图")
                self.status.set(f"{mode}：{detail}；可切换到学习参考图。")
        source = self.library.asset(app_id, observation.get("context_image", "")) if observation else None
        if source is None and not runtime:
            source = self.library.asset(app_id, page.get("layout_image") or page.get("image", ""))
            observation = {}
            self.display_after = False
        if source is None and runtime:
            observation = {}
        self.display_observation = observation
        self.viewer.set_image(source)
        self.zoom.set("适应")
        regions = [r for r in index["regions"] if r.get("page_id") == page_id
                   and not r.get("classification_only") and not dynamic_regions.is_dynamic(r)]
        if observation:
            from .mcp.control_observation import reframe
            main_rect = observation["main_rect"]
            regions = [{**r, "normalized_rect": reframe(r["normalized_rect"], main_rect)}
                       for r in regions if r.get("normalized_rect")]
            if not self.display_after:
                regions += [r for r in index["regions"] if dynamic_regions.is_dynamic(r)
                            and r.get("visual_observation", {}).get("context_image") == observation.get("context_image")]
            self.draw_mode.set(False); self.viewer.drawing = False
            self.viewer.canvas.configure(cursor="")
        self.viewer.regions = regions
        self.overlays()

    def overlays(self):
        if not self.selected: return
        _app, kind, record = self.selected
        controls = self.indexes.get(self.app_id, {}).get("controls", [])
        if kind == "region":
            if dynamic_regions.is_dynamic(record):
                graph = self.region_trees[self.app_id]
                descendants, pending = set(), ["region:"+record["id"]]
                while pending:
                    node = pending.pop()
                    if node in descendants: continue
                    descendants.add(node); pending.extend(graph["children"].get(node, []))
                controls = [{**c, "normalized_rect": c["visual_observation"]["normalized_rect"]} for c in controls
                            if "control:"+c["id"] in descendants and c.get("visual_observation", {}).get("context_image") == self.display_observation.get("context_image")]
            else:
                hierarchy = self.hierarchies.get(self.app_id, {})
                controls = [c for c in controls if hierarchy.get("regions", {}).get(c["id"]) == record["id"]
                            and c.get("page_id") == record.get("page_id") and not c.get("navigation_path")]
        elif kind == "control":
            missing_runtime = self.observation_mode.get() in {BEFORE_VIEW, RESULT_VIEW} and not self.display_observation
            controls = [] if self.display_after or missing_runtime else [{**record, "normalized_rect":
                self.display_observation.get("normalized_rect") or record.get("normalized_rect", {})}]
        else: controls = []
        self.viewer.controls = controls if self.show_controls.get() else []
        self.viewer.render()

    def select_region(self, region_id):
        key = f"region:{self.app_id}:{region_id}"
        if self.tree.exists(key): self.tree.selection_set(key); self.tree.see(key)

    def toggle_drawing(self):
        self.draw_parent_region = ""
        if self.draw_mode.get() and self.selected and self.selected[1] == "region" and dynamic_regions.is_dynamic(self.selected[2]):
            if not self._maybe_save():
                self.draw_mode.set(False); return
            self.observation_mode.set(REFERENCE_VIEW); self.show_observation()
            self.draw_parent_region = self.selected[2]["id"]
            self.draw_mode.set(True)
        elif self.draw_mode.get() and (self.display_page_id != PAGE_ID or self.display_observation) and self.app_id:
            if not self._maybe_save():
                self.draw_mode.set(False)
                return
            self.tree.selection_set(f"app:{self.app_id}")
            self.select()
        if self.draw_mode.get() and not self.draw_parent_region and (not self.app_id or not self.library.page(self.app_id).get("layout_image")):
            self.draw_mode.set(False)
            messagebox.showinfo("先获取截图", "请选择软件并获取截图后再划分区域。", parent=self.root)
        self.viewer.drawing = self.draw_mode.get()
        self.viewer.canvas.configure(cursor="crosshair" if self.draw_mode.get() else "")
        if self.viewer.drawing:
            self.status.set("在所选动态区域内拖动画出子区域；同级区域不能重叠。" if self.draw_parent_region
                            else "拖动画出一个区域，松开保存；可在右侧改名、精调像素边界。区域不能重叠。")

    def drawn(self, rectangle):
        try:
            if self.draw_parent_region:
                identifier = self.library.save_dynamic_region(self.app_id, "子区域", rectangle, parent_region_id=self.draw_parent_region)
                self.draw_mode.set(False); self.viewer.drawing = False; self.draw_parent_region = ""
                self.reload(f"region:{self.app_id}:{identifier}")
                self.status.set("子区域已保存，已学习控件按同一展开状态归档；可继续选中此区域细分。")
                return
            regions = self.library.regions(self.app_id)
            identifier = self.library.save_region(self.app_id, f"区域 {len(regions)+1}", rectangle)
            self.reload(f"region:{self.app_id}:{identifier}")
            self.status.set("手绘区域已保存。可继续添加区域，完成后点击“开始学习”。")
        except (ValueError, OSError) as error:
            messagebox.showerror("无法保存区域", str(error), parent=self.root)

    def make_region(self):
        if self.task or not self.selected or self.selected[1] != "control" or not self._maybe_save(): return
        try:
            identifier = self.library.make_dynamic_region(self.app_id, self.selected[2]["id"])
            self.reload(f"region:{self.app_id}:{identifier}")
            self.status.set("已设为动态区域；选中后勾选“手绘区域（备用）”可继续细分。")
        except (ValueError, OSError) as error:
            messagebox.showerror("无法设为区域", str(error), parent=self.root)

    def confirm(self):
        if not self.app_id: return
        if not self._maybe_save(): return
        try:
            confirmed = self.library.confirm_layout(self.app_id)
            self.status.set(f"已确认 {len(confirmed['regions'])} 个区域，可以开始学习。")
        except (ValueError, OSError) as error:
            messagebox.showerror("无法确认", str(error), parent=self.root)

    def choose_window(self):
        if self.task or not self._maybe_save(): return
        WindowChooser(self.root, self.library, self.app_id, self.bound)

    def bound(self, app_id, window):
        self.app_id = app_id; self.windows[app_id] = window
        self.reload(f"app:{app_id}")
        self.capture()

    def capture(self):
        if not self.app_id or self.app_id not in self.windows:
            self.choose_window(); return
        if not self._maybe_save(): return
        self.start_task("capture", {"window": self.windows[self.app_id]}, hide=True)

    def add_region(self):
        self.pick_region()

    def modify_region(self):
        if not self.selected or self.selected[1] != "region": return
        record = self.selected[2]
        if dynamic_regions.is_dynamic(record) or record.get("classification_only"): return
        self.pick_region(record)

    def pick_region(self, record=None):
        if self.task or self.workflow_panel.task or not self.app_id or not self._maybe_save(): return
        app_id = self.app_id
        payload = {"region_id": record["id"] if record else "",
                   "region_name": record.get("name", "") if record else "",
                   "region_description": record.get("description", "") if record else ""}
        def begin(bound_app, window):
            if bound_app != app_id: return
            self.app_id = app_id; self.windows[app_id] = window
            self.start_task("pick-region", {**payload, "window": window}, hide=True)
        if app_id in self.windows:
            begin(app_id, self.windows[app_id])
        else:
            application = self.indexes[app_id]["application"]
            title, class_name = application.get("window_name"), application.get("window_class")
            matches = [w for w in native_windows.matching_windows(title) if w["visible"]
                       and w["class_name"] == class_name] if title and class_name else []
            if len(matches) == 1:
                begin(app_id, matches[0])
            else:
                WindowChooser(self.root, self.library, app_id, begin)

    def automatic_partition(self):
        if self.task or not self.app_id or not self._maybe_save(): return
        if self.app_id not in self.windows:
            messagebox.showinfo("请选择软件", "请先绑定要自动划分的软件窗口。", parent=self.root); return
        page = self.library.page(self.app_id)
        image = self.library.asset(self.app_id, page.get("layout_image", ""))
        if image is None:
            self.capture(); return
        self.start_task("partition", {"window": self.windows[self.app_id], "image": str(image)})

    def render_navigation(self):
        records = self.indexes.get(self.app_id, {}).get("controls", [])
        self.navigation_records = {r["id"]: r for r in records}
        self.navigation_list.delete(0, "end")
        names = []
        for index, identifier in enumerate(self.navigation_ids, 1):
            record = self.navigation_records.get(identifier, {})
            name = record.get("semantic_name") or record.get("name") or f"已删除：{identifier}"
            names.append(name)
            self.navigation_list.insert("end", f"{index}. {name}")
        selected_id = self.selected[2]["id"] if self.selected and self.selected[1] == "control" else ""
        self.navigation_candidates = [r for r in records if r["id"] != selected_id and r.get("location")]
        self.navigation_choice["values"] = [f"{r.get('semantic_name') or r.get('name') or r['id']}  [{r['id']}]"
                                            for r in self.navigation_candidates]
        self.navigation_choice.set("")
        dynamic = self.selected and self.selected[1] == "region" and dynamic_regions.is_dynamic(self.selected[2])
        leaf = self.selected[2].get("semantic_name") or self.selected[2].get("name") or selected_id if selected_id else "请选择控件"
        self.navigation_summary.configure(text=("区域展开路径：" + " → ".join(names) + "（区域本身不点击）") if dynamic
                                          else "执行路径：" + " → ".join([*names, leaf]))

    def add_navigation(self):
        if self.task or not self.selected or self.selected[1] != "control": return
        index = self.navigation_choice.current()
        if index < 0: return
        record = self.navigation_candidates[index]
        proposed = list(self.navigation_ids)
        for identifier in [*record.get("navigation_path", []), record["id"]]:
            if identifier not in proposed: proposed.append(identifier)
        try:
            from .library_actions import validate_plan
            validate_plan(self.indexes[self.app_id]["controls"], self.selected[2]["id"], proposed)
            self.navigation_ids = proposed
            self.render_navigation(); self.changed()
        except (ValueError, RuntimeError) as error:
            messagebox.showerror("无法添加步骤", str(error), parent=self.root)

    def move_navigation(self, direction):
        if self.task: return
        selected = self.navigation_list.curselection()
        if not selected: return
        index, destination = selected[0], selected[0]+direction
        if 0 <= destination < len(self.navigation_ids):
            self.navigation_ids[index], self.navigation_ids[destination] = self.navigation_ids[destination], self.navigation_ids[index]
            self.render_navigation(); self.navigation_list.selection_set(destination); self.changed()

    def remove_navigation(self):
        if self.task: return
        selected = self.navigation_list.curselection()
        if selected:
            del self.navigation_ids[selected[0]]
            self.render_navigation(); self.changed()

    def expand(self):
        if self.selected and self.selected[1] == "control": self.learn(expand=True)

    def click(self):
        if self.task or not self.selected or self.selected[1] != "control": return
        if not self._maybe_save(): return
        if self.app_id not in self.windows:
            messagebox.showinfo("请选择软件", "先选择此知识库对应的实际软件窗口，再执行点击。", parent=self.root); return
        try:
            self.library.click_plan(self.app_id, self.selected[2]["id"])
        except (ValueError, RuntimeError) as error:
            messagebox.showerror("无法执行", str(error), parent=self.root); return
        self.start_task("click", {"vault": str(self.library.root), "app_id": self.app_id,
                                  "control_id": self.selected[2]["id"], "window": self.windows[self.app_id]}, hide=True)

    def learn(self, expand=False):
        if not self.app_id: return
        if self.app_id not in self.windows:
            messagebox.showinfo("请选择软件", "先重新选择要学习的软件窗口。", parent=self.root); return
        if not self._maybe_save(): return
        try:
            if not self.library.page(self.app_id).get("layout_confirmed"):
                if self.library.page(self.app_id).get("requires_layout_review"):
                    raise ValueError("窗口尺寸或类型已改变，请用“修改区域”重新拾取已有区域，再开始学习")
                self.library.confirm_layout(self.app_id)
            layout = self.library.learning_layout(self.app_id, self.windows[self.app_id])
            if self.enrich.get() and not configuration.vision_configuration_status()["ready"]:
                raise ValueError("请先在配置中填写含义补充服务，或关闭该选项")
        except (ValueError, OSError) as error:
            messagebox.showerror("无法开始学习", str(error), parent=self.root); return
        payload = {"layout": layout, "window_name": self.windows[self.app_id]["title"],
                                  "vault": str(self.library.root),
                                  "semantic_mode": "api" if self.enrich.get() else "native"}
        if expand:
            if not self.selected or self.selected[1] != "control": return
            payload.update(app_id=self.app_id, control_id=self.selected[2]["id"], window=self.windows[self.app_id])
        self.start_task("expand" if expand else "learn", payload, hide=True)

    def verify(self):
        if not self.selected or self.selected[1] != "control": return
        if not self._maybe_save(): return
        record = self.selected[2]
        if not record.get("location"):
            messagebox.showinfo("无法定位", "此记录只有图像观察，没有 UIA LOCATION。", parent=self.root); return
        self.start_task("verify", {"location": record["location"], "window_scope": record.get("window_scope", {"kind": "main"}),
                                   "window": self.windows.get(self.app_id)})

    def start_task(self, operation, payload, *, hide=False):
        if self.task: return
        if getattr(getattr(self, "workflow_panel", None), "task", None):
            self.status.set("流程正在运行，请等待完成或在流程窗口中停止。"); return
        self.task_app, self.task_operation = self.app_id, operation
        self.task_control = payload.get("control_id", "")
        self.task_region = {"region_id": payload.get("region_id", ""), "name": payload.get("region_name", ""),
                            "description": payload.get("region_description", "")}
        if hide:
            self.root.iconify(); self.root.update_idletasks()
        try:
            if hide and operation != "pick-region": self.monitor = TaskMonitor(self.root, None, "正在执行知识库任务")
            self.task = LibraryTask({"operation": operation, **payload, "capture_handshake": hide and operation != "pick-region"}, timeout=600 if operation in {"learn", "expand"} else 300 if operation == "pick-region" else 90 if operation == "click" else 12 if operation == "partition" else 30)
            if self.monitor: self.monitor.task = self.task
        except Exception as error:
            if self.task: self.task.close(); self.task = None
            if self.monitor: self.monitor.close(); self.monitor = None
            self.root.deiconify(); messagebox.showerror("无法启动任务", str(error), parent=self.root); return
        for button in self.buttons.values(): button.configure(state="disabled")
        self.path_button.configure(state="disabled"); self.draw_button.configure(state="disabled")
        self.viewer.enabled = False
        self.stop_button.configure(state="normal")
        self.progress.configure(mode="indeterminate"); self.progress.start(12)
        self.status.set("任务已启动，可用“停止”结束本次工作进程。")

    def stop(self):
        if self.task:
            self.task.cancel(); self.status.set("正在停止本次任务，已经保存的控件仍保留。")
        elif self.workflow_panel.task: self.workflow_panel.stop()

    def poll(self):
        self.poll_id = None
        if self.closed: return
        if not self.task and self.selected and self.selected[1] in {"combination", "actions-root"}:
            self.status.set(self.workflow_panel.status.get())
            self.stop_button.configure(state="normal" if self.workflow_panel.task else "disabled")
        if self.task:
            state = self.task.poll()
            if state["state"] == "running" or state["worker_running"]:
                message = state.get("message", "正在处理")
                self.status.set(f"{message}  ·  {state['elapsed_seconds']:.1f} 秒")
                if self.monitor: self.monitor.message.set(self.status.get())
            else:
                task, operation, app_id = self.task, self.task_operation, self.task_app
                try:
                    if state["state"] == "completed":
                        result = state["result"]
                        region_key = ""
                        if operation == "capture": self.library.save_capture(app_id, result)
                        if operation == "partition": result = self.library.apply_automatic_layout(app_id, result)
                        if operation == "pick-region":
                            region_key = self.library.save_picked_region(app_id, result, **self.task_region)
                        self.task = None
                        self.app_id = app_id
                        self.reload(f"region:{app_id}:{region_key}" if region_key else f"control:{app_id}:{self.task_control}" if self.task_control else f"app:{app_id}")
                        if operation in {"learn", "expand"}:
                            self.status.set(f"学习完成：保存 {result['controls_saved_this_scan']} 个控件，待归档 {result.get('unassigned_controls', 0)} 个，用时 {result['elapsed_seconds']:.1f} 秒。" + (" 达到扫描上限，请检查完整性。" if result.get("truncated") else ""))
                            if operation == "expand":
                                self.status.set(self.status.get() + f" 其中 {result.get('navigation_controls', 0)} 个控件已记录前置步骤。")
                            for key, node in self.nodes.items():
                                if node[0] == app_id and node[1] == "region": self.tree.item(key, open=True)
                        elif operation == "capture":
                            self.status.set("软件窗口截图已获取。点击“添加区域”进入悬浮拾取；已有区域可选中后点击“修改区域”。")
                        elif operation == "pick-region":
                            self.status.set("区域已保存。可继续添加区域，或点击“开始学习”。" if self.library.page(app_id).get("layout_confirmed")
                                            else "区域已保存；窗口尺寸已变化，请继续修改其余区域后再学习。")
                        elif operation == "partition":
                            self.status.set(f"自动划分完成：生成 {result['suggested']} 个区域，保留 {result['preserved']} 个人工区域。可改名、调整边界、增删区域后确认并学习。" + result.get("warning", ""))
                        elif operation == "click":
                            self.observation_mode.set(RESULT_VIEW)
                            self.show_observation()
                            if self.viewer.image:
                                self.status.set(result.get("message", "点击已执行") + "；当前显示运行结果，学习参考图保持不变。")
                        else: self.status.set(result.get("message", "定位成功"))
                    else:
                        self.task = None
                        self.reload(f"control:{app_id}:{self.task_control}" if self.task_control else f"app:{app_id}")
                        if operation == "click":
                            self.observation_mode.set(RESULT_VIEW)
                            self.show_observation()
                        self.status.set(state.get("error", "任务已停止"))
                        if state["state"] == "failed": messagebox.showerror("任务未完成", state.get("error", "任务失败"), parent=self.root)
                except Exception as error:
                    self.task = None
                    self.status.set(f"无法保存任务结果：{error}")
                    messagebox.showerror("无法保存结果", str(error), parent=self.root)
                finally:
                    if self.monitor: self.monitor.close(); self.monitor = None
                    task.close()
                    for button in self.buttons.values(): button.configure(state="normal")
                    self.path_button.configure(state="normal"); self.draw_button.configure(state="normal")
                    self.viewer.enabled = True
                    self.stop_button.configure(state="disabled"); self.progress.stop(); self.progress.configure(mode="determinate", value=0)
                    if self.selected: self.update_actions(self.selected[1])
                    self.root.deiconify(); self.root.lift()
        self.poll_id = self.root.after(150, self.poll)

    def save(self):
        if not self.selected: return True
        app_id, kind, record = self.selected
        if kind == "usage": return self.usage_panel.save()
        if kind == "combination": return self.workflow_panel.save()
        if kind not in {"application", "controls-root", "region", "control"}: return True
        try:
            if kind == "region":
                if dynamic_regions.is_dynamic(record):
                    if not self.viewer.image: raise ValueError("区域没有可用学习参考图")
                    rect = {k: int(v.get())/(self.viewer.image.width if k in {"left", "right"} else self.viewer.image.height)
                            for k, v in self.rect_values.items()}
                    self.library.save_dynamic_region(app_id, self.name.get(), rect, region_id=record["id"], description=self.description.get())
                    key = self.tree.selection()[0]
                    self.dirty = False; self.reload(key); self.status.set("动态区域已保存，XPath 和前置动作保持原值。")
                    return True
                if (record.get("source") != "manual" or record.get("page_id") != PAGE_ID
                        or record.get("classification_only")):
                    raise ValueError("请在本软件的人工布局中划分区域；历史自动区域仅供查看")
                if not self.viewer.image: raise ValueError("区域没有可用截图")
                rect = {k: int(v.get())/(self.viewer.image.width if k in {"left", "right"} else self.viewer.image.height)
                        for k, v in self.rect_values.items()}
                self.library.save_region(app_id, self.name.get(), rect, record["id"], description=self.description.get())
            elif kind == "control":
                text = self.location.get("1.0", "end").strip()
                edited_location = json.loads(text) if text else None
                mode_changed = self.mode.get() != getattr(self, "loaded_mode", "identity")
                locator_changed = edited_location != (record.get("location") or None)
                selected_parent = self.parent_options.get(self.parent_selection.get(), "auto")
                self.library.update_control(app_id, record["id"], name=self.name.get(), description=self.description.get(),
                                            aliases=[s.strip() for s in self.aliases.get().replace("，", ",").split(",") if s.strip()],
                                            location=edited_location if locator_changed or mode_changed else None,
                                            match_mode=self.mode.get() if mode_changed else None,
                                            navigation_path=self.navigation_ids if self.navigation_ids != record.get("navigation_path", []) else None,
                                            hierarchy_parent=selected_parent if selected_parent != self.loaded_parent else None)
            else:
                directory = self.library.directory(app_id)
                metadata, body = knowledge.read_markdown(directory / "app.md")
                if not self.name.get().strip(): raise ValueError("软件名称不能为空")
                metadata.update(name=self.name.get().strip(), description=self.description.get())
                knowledge.write_markdown(directory / "app.md", metadata, body)
                knowledge.rebuild_index(directory)
            key = self.tree.selection()[0] if self.tree.selection() else ""
            self.dirty = False; self.reload(key); self.status.set("修改已保存到现有知识库。")
            return True
        except (ValueError, OSError, TypeError, KeyError, RuntimeError) as error:
            messagebox.showerror("保存失败", str(error), parent=self.root); return False

    def remove(self):
        if not self.selected or self.selected[1] not in {"region", "control"}: return
        app_id, kind, record = self.selected
        if not messagebox.askyesno("删除所选", "记录将移入软件目录的归档中，图片保留。删除区域后，其中控件会重新归档。是否继续？", parent=self.root): return
        try:
            self.library.archive(app_id, kind, record["id"])
            self.dirty = False; self.reload(f"app:{app_id}"); self.status.set("已移入 .easy_uiauto/archive，未删除图片。")
        except (ValueError, OSError) as error: messagebox.showerror("删除失败", str(error), parent=self.root)

    def export(self):
        if not self.app_id: return
        path = filedialog.asksaveasfilename(parent=self.root, title="导出软件知识", defaultextension=".json", initialfile=f"{self.app_id}.json", filetypes=[("JSON", "*.json")])
        if path:
            try: self.library.export(self.app_id, path); self.status.set(f"已导出：{path}")
            except (ValueError, OSError) as error: messagebox.showerror("导出失败", str(error), parent=self.root)

    def change_vault(self, *, use_typed=False):
        if use_typed:
            path = self.path.get().strip()
        else:
            initial = Path(self.path.get()).expanduser()
            path = filedialog.askdirectory(parent=self.root, title="打开知识库根目录",
                                           initialdir=str(initial if initial.is_dir() else self.library.root))
        if not path: return
        try:
            library = self.prepare_vault_switch(path)
            if library.root == self.library.root:
                self.path.set(str(self.library.root)); return
            if not self._maybe_save(): return
            self.apply_vault_switch(library)
        except (ValueError, OSError, RuntimeError) as error:
            self.path.set(str(self.library.root))
            messagebox.showerror("无法切换知识库", str(error), parent=self.root)

    def prepare_vault_switch(self, path):
        """Validate the destination and guard active tasks before changing settings."""
        if not str(path).strip() or not Path(path).expanduser().is_absolute():
            raise ValueError("请输入知识库的绝对路径。")
        library = KnowledgeLibrary(path)
        if library.root == self.library.root: return library
        if self.task or self.workflow_panel.task:
            raise ValueError("当前有学习或组合动作任务正在运行，请结束任务后再切换知识库。")
        if library.root.exists() and not library.root.is_dir():
            raise ValueError("知识库路径必须是目录，不能是文件。")
        for app in library.apps(): library.read(app["id"])
        return library

    def _replace_library(self, library):
        self.library = library; self.path.set(str(library.root))
        self.usage_panel.library = self.workflow_panel.library = library
        panel = self.workflow_panel
        panel.dirty = panel.editor_dirty = False
        panel.app_id = None; panel.show_app("")
        panel.apps = {f"{r.get('name', r['id'])} [{r['id']}]": r["id"] for r in library.apps()}
        panel.app_choice.set(""); panel.app_combo["values"] = list(panel.apps)
        panel.binding["values"] = (); panel.bound_window.set("未绑定运行窗口")
        self.usage_panel.record = {}; self.usage_panel.dirty = False; self.usage_panel.actions = []
        self.app_id, self.selected = "", None; self.windows.clear(); self.dirty = False
        self.loading = True
        try:
            self.name.set(""); self.description.set(""); self.aliases.set("")
            for value in self.rect_values.values(): value.set("")
            self._set_text(self.location, ""); self.location.edit_modified(False)
            self._set_text(self.code, ""); self.code.configure(state="disabled")
            self.picture_label.configure(image="", text="尚无图片"); self.picture_path.configure(text="")
            self.detail_title.configure(text="未选择节点")
            self.draw_parent_region = ""; self.draw_mode.set(False); self.viewer.drawing = False
            self.viewer.canvas.configure(cursor="")
            self.navigation_ids = []
            self.display_page_id = PAGE_ID
            self.display_observation = {}; self.display_after = False; self.observation_control = None
            self.observation_mode.set(REFERENCE_VIEW)
            self.viewer.regions = []; self.viewer.controls = []; self.viewer.set_image(None)
        finally: self.loading = False
        self.reload()

    def apply_vault_switch(self, library):
        previous, windows = self.library, dict(self.windows)
        selected = self.tree.selection()
        try:
            self._replace_library(library)
        except Exception as error:
            try:
                self._replace_library(previous); self.windows.update(windows)
                if selected and self.tree.exists(selected[0]):
                    self.tree.selection_set(selected[0]); self.select()
            except Exception as recovery_error: error.add_note(f"Restoring the previous vault also failed: {recovery_error}")
            raise
        self.status.set(f"已切换知识库：{self.library.root}")

    def before_configuration_save(self, values):
        self.pending_config_library = self.prepare_vault_switch(values["knowledge_dir"])
        return self.pending_config_library.root == self.library.root or self._maybe_save()

    def configuration_saved(self, saved):
        library = self.pending_config_library
        if library.root != self.library.root:
            self.apply_vault_switch(library)
        return f"配置已保存。当前 GUI 使用：{self.library.root}；API 设置用于后续读取配置的任务。"

    def configure(self):
        from .mcp.config_gui import ConfigurationWindow

        panel = getattr(self, "config_panel", None)
        if panel is not None and not panel.closed and panel.root.winfo_exists():
            panel.root.deiconify(); panel.root.lift(); return panel
        window = tk.Toplevel(self.root); window.transient(self.root)
        self.config_panel = ConfigurationWindow(window, knowledge_dir=self.library.root,
            before_save=self.before_configuration_save, on_saved=self.configuration_saved)
        window.update_idletasks()
        if window.winfo_viewable(): window.grab_set()
        return self.config_panel

    def open_workflows(self):
        if self.task or not self.app_id: return
        key = f"actions-root:{self.app_id}"
        self.tree.selection_set(key); self.tree.see(key); self.select()

    def close(self):
        workflow = getattr(self, "workflow_panel", None)
        if workflow and workflow.task:
            messagebox.showinfo("组合动作运行中", "先停止当前组合动作，工作进程退出后再关闭。", parent=self.root); return
        if self.task and not messagebox.askyesno("任务进行中", "停止本次任务并关闭知识库窗口？", parent=self.root): return
        config = getattr(self, "config_panel", None)
        if config is not None and not config.closed and not config.close(): return
        if not self.task and not self._maybe_save(): return
        self.closed = True
        if self.poll_id: self.root.after_cancel(self.poll_id)
        if workflow and workflow.poll_id: workflow.root.after_cancel(workflow.poll_id)
        if self.monitor: self.monitor.close()
        if self.task: self.task.close()
        self.root.destroy()


def main():
    # Keep old class imports available for existing integrations; new launches
    # use Qt without nesting a Tk event loop in the GUI process.
    from .qt_gui import main as run_gui
    return run_gui()


if __name__ == "__main__":
    main()
