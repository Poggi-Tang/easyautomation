"""Task-focused Qt knowledge manager over the existing vault and worker services."""
from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path
from uuid import uuid4

from PySide6.QtCore import QItemSelectionModel, QSettings, QSignalBlocker, QThreadPool, QTimer, Qt
from PySide6.QtGui import QAction, QIcon, QKeySequence, QShortcut
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QComboBox, QDialog, QFileDialog, QFormLayout,
    QDialogButtonBox, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QMainWindow, QMenu, QMessageBox,
    QPlainTextEdit, QProgressBar, QPushButton, QScrollArea, QSplitter, QStackedWidget,
    QSizePolicy, QTabWidget, QTextBrowser, QToolButton, QTreeView, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget)

from ..knowledge_library import KnowledgeLibrary, PAGE_ID
from ..library_tasks import LibraryTask
from ..mcp import configuration, knowledge, dynamic_regions, native_windows
from ..locator_images import reference_images, is_image_location, record_location
from ..mcp.protocol import CONTROL_IDENTITY_KEYS, normalize_location, with_control_identity
from .. import workflows
from .models import KnowledgeModel, KnowledgeFilter, ROLE
from .widgets import ImageView, IoJob, StepEditor, ControlChooser, button, ACTION_NAMES
from .dialogs import ParameterDialog, SettingsDialog, WindowDialog
from .icons import icon, IconLabel, action_icon_name, set_button_icon, decorate_button_box, decorate_search


class KnowledgeWindow(QMainWindow):
    def __init__(self, vault=None, *, remember=True):
        super().__init__();self.library=KnowledgeLibrary(vault);self.remember=remember
        self.snapshots={};self.windows={};self.current={};self.flow=None;self.dirty=False;self.loading=False
        self.task=None;self.task_done=None;self.jobs=[];self.load_serial=0;self.active_dialog=None
        self.entry_condition=None;self.result={};self.last_key="vault";self.context_buttons=[]
        self.syncing_identity=False;self.identity_error="";self.locator_snapshot={}
        self.deleting=False;self.selection_caption="知识库"
        self.model_report=None;self.model_check_job=None;self.model_check_again=False;self.model_dialog=None
        self.setWindowTitle("easy-uiauto 知识库管理")
        self.setWindowIcon(QIcon(str(Path(__file__).resolve().parents[1]/"resources/easy-uiauto.ico")))
        screen=QApplication.primaryScreen().availableGeometry()
        self.setMinimumSize(min(1000,screen.width()-30),min(650,screen.height()-30))
        self.resize(min(1440,screen.width()-50),min(900,screen.height()-50))
        self.settings=QSettings("easy-uiauto","KnowledgeGUI") if remember else None
        self._build();self.reload();self.refresh_image_models()
        self.timer=QTimer(self);self.timer.setInterval(120);self.timer.timeout.connect(self.poll);self.timer.start()
        QShortcut(QKeySequence.StandardKey.Save,self,self.save)

    def _build(self):
        main=QWidget();outer=QVBoxLayout(main);outer.setContentsMargins(12,10,12,8);self.setCentralWidget(main)
        top=QHBoxLayout();self.vault_label=QLabel();self.vault_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.vault_label.setSizePolicy(QSizePolicy.Policy.Ignored,QSizePolicy.Policy.Preferred)
        self.vault_label.setToolTip(str(self.library.root));top.addWidget(self.vault_label,1)
        top.addWidget(button("切换知识库",self.change_vault));self.connection=IconLabel("未连接软件窗口",icon_name="status-pending");self.connection.setMaximumWidth(250);top.addWidget(self.connection)
        self.connect_button=button("选择窗口",self.choose_window,name="choose-window");top.addWidget(self.connect_button)
        self.model_button=button("标注模型：检查中…",self.open_image_models,name="image-models");top.addWidget(self.model_button)
        top.addWidget(button("设置",self.configure));outer.addLayout(top)
        self.main_split=QSplitter();outer.addWidget(self.main_split,1)
        nav=QWidget();nav_layout=QVBoxLayout(nav);nav_layout.setContentsMargins(0,0,6,0)
        nav_layout.addWidget(QLabel("知识库"));self.search=QLineEdit();self.search.setAccessibleName("搜索知识节点");self.search.setPlaceholderText("搜索软件、区域、控件或动作")
        decorate_search(self.search)
        nav_layout.addWidget(self.search);self.tree=QTreeView();self.tree.setAccessibleName("知识库导航")
        self.model=KnowledgeModel(self);self.proxy=KnowledgeFilter(self);self.proxy.setSourceModel(self.model)
        self.tree.setModel(self.proxy);self.tree.setUniformRowHeights(True);self.tree.setAlternatingRowColors(False)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tree.setToolTip("Ctrl 点选多项，Shift 连选；Delete 删除选中的区域或控件")
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.tree_menu)
        self.delete_shortcut=QShortcut(QKeySequence(Qt.Key.Key_Delete),self.tree)
        self.delete_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self.delete_shortcut.setAutoRepeat(False);self.delete_shortcut.activated.connect(self.remove)
        self.tree.header().setStretchLastSection(False)
        self.tree.header().setSectionResizeMode(0,QHeaderView.ResizeMode.Stretch)
        self.tree.header().setSectionResizeMode(1,QHeaderView.ResizeMode.ResizeToContents)
        nav_layout.addWidget(self.tree,1);self.main_split.addWidget(nav)
        self.search.textChanged.connect(self.proxy.setFilterFixedString)
        self.search.textChanged.connect(lambda text:self.tree.expandAll() if text else None)
        self.tree.selectionModel().currentChanged.connect(self.select)
        self.tree.selectionModel().selectionChanged.connect(lambda *_:self.build_toolbar() if not self.loading else None)
        work=QWidget();work.setAccessibleName("知识工作区");layout=QVBoxLayout(work);layout.setContentsMargins(8,0,0,0)
        self.caption=QLabel("知识库");self.caption.setWordWrap(True);layout.addWidget(self.caption)
        toolbar_widget=QWidget();toolbar_widget.setAccessibleName("操作栏");self.toolbar=QHBoxLayout(toolbar_widget);self.toolbar.setContentsMargins(0,0,0,0)
        toolbar_scroll=QScrollArea();toolbar_scroll.setWidgetResizable(True);toolbar_scroll.setWidget(toolbar_widget)
        toolbar_scroll.setAccessibleName("操作栏滚动区")
        toolbar_scroll.setFrameShape(QScrollArea.Shape.NoFrame);toolbar_scroll.setFixedHeight(47)
        toolbar_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff);layout.addWidget(toolbar_scroll)
        self.work_split=QSplitter();self.work_split.setAccessibleName("内容与属性");layout.addWidget(self.work_split,1);self.main_split.addWidget(work)
        self.center=QStackedWidget();self.center.setAccessibleName("知识内容");self.work_split.addWidget(self.center)
        self.overview=QTextBrowser();self.center.addWidget(self.overview)
        self.image=ImageView();self.center.addWidget(self.image)
        self.usage=QPlainTextEdit();self.usage.setAccessibleName("使用知识正文");self.usage.textChanged.connect(self.changed);self.center.addWidget(self.usage)
        self.flow_page=QWidget();flow_layout=QVBoxLayout(self.flow_page);flow_layout.setContentsMargins(0,0,0,0)
        self.flow_title=QLineEdit();self.flow_title.setAccessibleName("组合动作标题");self.flow_title.textEdited.connect(self.changed);flow_layout.addWidget(self.flow_title)
        self.flow_steps=StepEditor(self.choose_control,inline=False);self.flow_steps.changed.connect(self.changed);flow_layout.addWidget(self.flow_steps)
        self.center.addWidget(self.flow_page)
        self.results_page=QWidget();results_layout=QVBoxLayout(self.results_page);results_layout.setContentsMargins(0,0,0,0)
        self.result_summary=IconLabel();results_layout.addWidget(self.result_summary)
        self.result_tree=QTreeWidget();self.result_tree.setHeaderLabels(["动作 / 预期","结果"]);self.result_tree.setColumnWidth(0,370)
        results_layout.addWidget(self.result_tree,1);self.result_text=QTextBrowser();self.result_text.setMaximumHeight(150);results_layout.addWidget(self.result_text)
        self.result_image=ImageView();results_layout.addWidget(self.result_image,2)
        row=QHBoxLayout();row.addWidget(button("执行前",lambda:self.result_picture("before")));row.addWidget(button("执行后",lambda:self.result_picture("after")));row.addStretch();row.addWidget(button("返回步骤",self.show_steps));results_layout.addLayout(row)
        self.result_tree.currentItemChanged.connect(self.select_result);self.center.addWidget(self.results_page)
        self.details=QTabWidget();self.details.setMinimumWidth(270);self.work_split.addWidget(self.details)
        scroll=QScrollArea();scroll.setWidgetResizable(True);properties=QWidget()
        properties_layout=QVBoxLayout(properties);self.form=QFormLayout()
        self.form.setFormAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignTop)
        self.form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignTop)
        self.form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.DontWrapRows)
        self.form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        properties_layout.addLayout(self.form);properties_layout.addStretch(1);scroll.setWidget(properties)
        self.name=QLineEdit();self.name.setAccessibleName("名称");self.description=QPlainTextEdit();self.description.setMaximumHeight(85)
        self.form.addRow("名称",self.name);self.form.addRow("说明",self.description)
        self.reference_info=QLabel();self.reference_info.setWordWrap(True);self.reference_info.setTextFormat(Qt.TextFormat.PlainText)
        self.reference_info.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignTop)
        self.reference_info.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse);self.form.addRow(self.reference_info)
        self.state_name=QLineEdit();self.form.addRow("状态名称",self.state_name)
        self.hierarchy=QComboBox();self.form.addRow("上级控件",self.hierarchy)
        self.native_fields={}
        for key,label in (("automation_id","AutomationID"),("name","Name"),("class_name","ClassName"),("control_type","ControlType")):
            field=QLineEdit();field.setAccessibleName(label);field.setPlaceholderText("（空）")
            field.textChanged.connect(self.edit_native_identity)
            self.native_fields[key]=field;self.form.addRow(label,field)
        self.entry_summary=QLabel("无");self.entry_summary.setWordWrap(True);self.form.addRow("前置步骤",self.entry_summary)
        self.condition_label=QLabel("无");self.condition_label.setWordWrap(True);self.form.addRow("状态条件",self.condition_label)
        for field in (self.entry_summary,self.condition_label):
            field.setAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignTop)
            field.setSizePolicy(QSizePolicy.Policy.Ignored,QSizePolicy.Policy.Preferred)
        self.condition_kind=QComboBox();self.condition_kind.addItem("控件存在","exists");self.condition_kind.addItem("Tab 已选中","selected")
        self.form.addRow("条件类型",self.condition_kind)
        self.pick_condition_button=button("拾取状态条件",self.pick_condition);self.form.addRow(self.pick_condition_button)
        self.details.addTab(scroll,icon("sliders"),"属性")
        entries=QWidget();entries_layout=QVBoxLayout(entries)
        self.entry_editor=StepEditor(self.choose_control,entries);self.entry_editor.hide();self.entry_editor.changed.connect(self.entry_changed)
        self.entry_preview=QTextBrowser();entries_layout.addWidget(self.entry_preview,1)
        self.edit_entries_button=button("编辑前置步骤",self.edit_entries);entries_layout.addWidget(self.edit_entries_button)
        row=QHBoxLayout();self.pick_entry_button=button("拾取入口",self.pick_entry);self.record_entry_button=button("录制步骤",self.record_entries)
        row.addWidget(self.pick_entry_button);row.addWidget(self.record_entry_button);entries_layout.addLayout(row)
        self.details.addTab(entries,icon("workflow"),"前置步骤")
        advanced=QWidget();advanced_layout=QVBoxLayout(advanced)
        self.mode=QComboBox();self.mode.addItem("身份（名称 / 标识）","identity");self.mode.addItem("位置（同类序号）","position")
        advanced_layout.addWidget(self.mode)
        self.template_selector=QComboBox();self.template_selector.setAccessibleName("定位图片")
        advanced_layout.addWidget(QLabel("定位图片（用于属性定位未找到时匹配）"));advanced_layout.addWidget(self.template_selector)
        self.template_path=QLineEdit();self.template_path.setReadOnly(True);self.template_path.setAccessibleName("定位图片路径")
        advanced_layout.addWidget(self.template_path)
        self.template_preview=ImageView();self.template_preview.setAccessibleName("定位图片预览")
        self.template_preview.setMinimumHeight(100);self.template_preview.setMaximumHeight(180);advanced_layout.addWidget(self.template_preview)
        self.template_selector.currentIndexChanged.connect(self.show_locator_image)
        self.locator=QPlainTextEdit();self.locator.setAccessibleName("LOCATION 定位信息");advanced_layout.addWidget(self.locator)
        self.details.addTab(advanced,icon("sliders"),"高级")
        self.details.addTab(self.flow_steps.details,icon("edit"),"步骤编辑")
        self.name.textEdited.connect(self.changed);self.description.textChanged.connect(self.changed)
        self.state_name.textEdited.connect(self.changed);self.condition_kind.currentIndexChanged.connect(self.changed)
        self.hierarchy.currentIndexChanged.connect(self.changed)
        self.locator.textChanged.connect(self.edit_locator);self.mode.currentIndexChanged.connect(self.changed)
        self.main_split.setSizes([280,1120]);self.work_split.setSizes([780,340])
        self.status_label=IconLabel("正在加载知识库…",icon_name="status-busy");self.status_label.setWordWrap(True);self.statusBar().addWidget(self.status_label,1)
        self.progress=QProgressBar();self.progress.setFixedWidth(110);self.progress.hide();self.statusBar().addPermanentWidget(self.progress)
        self.stop_button=button("停止",self.stop);self.stop_button.setEnabled(False);self.statusBar().addPermanentWidget(self.stop_button)
        if self.settings:
            for widget,key in ((self.main_split,"navigation"),(self.work_split,"workspace")):
                value=self.settings.value(key)
                if value:widget.restoreState(value)

    @property
    def app_id(self):return self.current.get("app_id","")

    @property
    def record(self):return self.current.get("record",{})

    @property
    def scoped(self):return self.library.scoped(self.record.get("page_id",PAGE_ID))

    def error(self,title,error):
        self.status_label.set_icon("status-error");self.status_label.setText(str(error));QMessageBox.warning(self,title,str(error))

    def async_io(self,function,done):
        job=IoJob(function);self.jobs.append(job)
        def finished(value,error):
            if job in self.jobs:self.jobs.remove(job)
            if error:self.error("读取失败",error)
            else:done(value)
        job.signals.done.connect(finished);QThreadPool.globalInstance().start(job)

    def open_image_models(self):
        if self.model_dialog is None:
            from .image_models_dialog import ImageModelsDialog
            self.model_dialog=ImageModelsDialog(self)
        if self.model_report:self.model_dialog.set_report(self.model_report)
        self.update_model_buttons()
        self.model_dialog.show();self.model_dialog.raise_();self.model_dialog.activateWindow()
        self.refresh_image_models()

    def update_model_buttons(self):
        if self.model_dialog:
            downloading=bool(self.task and self.task.operation=="download-image-models")
            self.model_dialog.set_busy(bool(self.task),checking=bool(self.model_check_job),downloading=downloading)

    def apply_model_report(self,report):
        self.model_report=report
        count=sum(model["state"]=="ready" for model in report["models"])
        label="已就绪" if report["ok"] else "环境缺失" if report["runtime_issues"] else f"待补齐 {count}/3"
        self.model_button.setText("标注模型："+label)
        self.model_button.setToolTip(report["message"]+"；点击查看和下载")
        if self.model_dialog:self.model_dialog.set_report(report)
        self.update_model_buttons()

    def refresh_image_models(self):
        if self.model_check_job or self.task:return
        from ..image_models import inspect_models
        job=IoJob(inspect_models);self.model_check_job=job;self.jobs.append(job)
        self.model_button.setText("标注模型：检查中…")
        if self.model_dialog and not self.model_dialog.paths_dirty:self.model_dialog.message.setText("正在检查标注模型…")
        self.update_model_buttons()
        def finished(report,error):
            self.jobs.remove(job);self.model_check_job=None
            if self.model_check_again:
                self.model_check_again=False;self.refresh_image_models();return
            if error:
                self.model_button.setText("标注模型：检查失败");self.model_button.setToolTip(str(error))
                if self.model_dialog:self.model_dialog.message.setText("检查失败："+str(error))
                self.update_model_buttons()
            else:self.apply_model_report(report)
        job.signals.done.connect(finished);QThreadPool.globalInstance().start(job)

    def download_image_models(self):
        if self.task or self.model_check_job:return
        if self.model_dialog and self.model_dialog.paths_dirty:
            self.model_dialog.save_paths();return
        self.start_task("download-image-models",{},self.apply_model_report)
        if self.task:
            self.model_button.setText("标注模型：下载中…")
            if self.model_dialog:self.model_dialog.message.setText("正在准备下载缺失或损坏的模型…")
        self.update_model_buttons()

    def save_image_model_paths(self,paths):
        if self.task:return False
        from ..image_models import save_model_paths
        try:save_model_paths(paths)
        except (OSError,ValueError) as error:
            if self.model_dialog:self.model_dialog.message.setText("路径未保存："+str(error))
            return False
        if self.model_dialog:self.model_dialog.paths_dirty=False
        if self.model_check_job:self.model_check_again=True
        self.refresh_image_models()
        return True

    def reload(self,key=None,*,message=None):
        self.load_serial+=1;serial=self.load_serial;self.vault_label.setText("知识库："+self.library.root.name)
        self.vault_label.setToolTip(str(self.library.root));library=self.library
        expanded=[k for k,item in self.model.nodes.items() if self.tree.isExpanded(self.proxy.mapFromSource(item.index()))]
        target=key or self.current.get("key","vault")
        self.reload_key=target
        def loaded(snapshots):
            if serial!=self.load_serial:return
            self.loading=True;self.snapshots=snapshots;self.model.load(library.root.name,snapshots,self.windows)
            self.tree.header().setSectionResizeMode(0,QHeaderView.ResizeMode.Stretch)
            self.tree.header().setSectionResizeMode(1,QHeaderView.ResizeMode.ResizeToContents)
            for identifier in ["vault",*expanded]:
                if identifier in self.model.nodes:self.tree.setExpanded(self.proxy.mapFromSource(self.model.nodes[identifier].index()),True)
            desired=self.reload_key
            self.loading=False;self.goto(desired if desired in self.model.nodes else "vault")
            if not self.task:self.status_label.set_icon("");self.status_label.setText(message or "已加载知识库；选择软件、区域或组合动作。")
        self.async_io(lambda:{app["id"]:library.read(app["id"]) for app in library.apps()},loaded)

    def goto(self,key):
        item=self.model.nodes.get(key)
        if item:
            index=self.proxy.mapFromSource(item.index())
            blocker=QSignalBlocker(self.tree.selectionModel())
            self.tree.selectionModel().setCurrentIndex(index,QItemSelectionModel.SelectionFlag.ClearAndSelect|QItemSelectionModel.SelectionFlag.Rows)
            del blocker
            self.tree.scrollTo(index);self.select(index)

    def maybe_save(self):
        if self.deleting:return False
        if not self.dirty:return True
        answer=QMessageBox.question(self,"未保存修改","保存当前修改？",QMessageBox.StandardButton.Save|QMessageBox.StandardButton.Discard|QMessageBox.StandardButton.Cancel)
        if answer==QMessageBox.StandardButton.Cancel:return False
        if answer==QMessageBox.StandardButton.Save:return self.save()
        self.dirty=False;return True

    def select(self,index,*_):
        if self.loading or not index.isValid():return
        index=index.siblingAtColumn(0)
        node=index.data(ROLE)
        if not node:return
        serial_before=self.load_serial
        if self.current.get("key")!=node["key"] and not self.maybe_save():
            self.loading=True
            old=self.model.nodes.get(self.current.get("key"))
            if old:self.tree.setCurrentIndex(self.proxy.mapFromSource(old.index()))
            self.loading=False;return
        if self.load_serial!=serial_before:self.reload_key=node["key"]
        self.current=deepcopy(node);self.loading=True;self.dirty=False;self.flow=None
        kind,r=node["kind"],node["record"]
        label={"vault":"知识库","controls-root":"控件库","usage-root":"使用知识","actions-root":"组合动作"}.get(kind,kind)
        self.caption.setText(" / ".join(s for s in (self.snapshots.get(self.app_id,{}).get("application",{}).get("name",""),r.get("semantic_name") or r.get("name") or r.get("keyword") or label) if s))
        self.selection_caption=self.caption.text()
        window=self.windows.get(self.app_id)
        self.connection.setText(window["title"] if window else "未连接软件窗口")
        self.connection.set_icon("status-success" if window else "status-pending")
        self.name.setText(r.get("semantic_name") or r.get("name", ""));self.description.setPlainText(r.get("description",""))
        self.description.setPlaceholderText(r.get("description_reason") or "尚未补充说明")
        self.reference_info.setText((f"来源：{r.get('source','')}\n标准 Markdown：{r.get('document_path','')}" if kind=="reference" else
            ("说明来源：根据软件理解推断，尚未交互验证" if r.get("description_method")=="software-understanding" else
             "说明来源：参考资料推断，尚未交互验证") if r.get("description_source")=="reference-inference" else
            "说明来源：实际交互观察" if r.get("description_source")=="interaction-observation" else ""))
        if kind=="control" and r.get("description_status") in {"unmatched","error"}:
            self.reference_info.setText(("参考说明未生成：" if r["description_status"]=="error" else "说明待确认：")+
                                        r.get("description_reason","未找到明确依据"))
        if kind=="control" and r.get("reference_evidence"):
            evidence=r["reference_evidence"][0]
            source=next((d for d in self.snapshots[self.app_id].get("references",[]) if d["id"]==evidence["reference_id"]),{})
            basis=("推断依据："+evidence['basis'][:1000] if evidence.get('basis') else "原文："+evidence.get('quote','')[:600])
            self.reference_info.setText(self.reference_info.text()+f"\n资料：{source.get('name','已移除的参考资料')}\n{basis}")
        try:native_location=record_location(r) if kind=="control" else {}
        except (ValueError,TypeError):native_location={}
        for key,field in self.native_fields.items():
            value=r.get(key,"") if is_image_location(native_location) else native_location.get(CONTROL_IDENTITY_KEYS[key],r.get(key,""))
            field.setText(str(value or ""));field.setReadOnly(kind!="control" or is_image_location(native_location))
        self.details.setVisible(kind in {"region","control","area","workflow","reference"})
        for i in range(4):self.details.setTabVisible(i,(i==3) if kind=="workflow" else i==0 if kind=="reference" else i<3)
        for field in [self.description,self.state_name,self.hierarchy,*self.native_fields.values(),
                      self.entry_summary,self.condition_label,self.condition_kind,self.pick_condition_button]:
            self.form.setRowVisible(field,kind!="reference")
        self.details.setCurrentIndex(3 if kind=="workflow" else 0)
        self.name.setEnabled(kind in {"region","control","application","usage"})
        self.locator.setPlainText(json.dumps(native_location if kind=="control" else r.get("location",{}),ensure_ascii=False,indent=2));self.locator.setReadOnly(kind!="control")
        self.locator_snapshot=deepcopy(native_location if kind=="control" else r.get("location",{}));self.identity_error=""
        self.refresh_locator_images()
        mode=(native_location.get("PARAMETERS") or {}).get("match_mode","identity")
        self.mode.setCurrentIndex(max(0,self.mode.findData(mode)));self.mode.setEnabled(kind=="control" and not is_image_location(native_location))
        self.entry_condition=None;state={}
        if kind in {"region","control","area"}:state=self.library.record_state(self.app_id,r,snapshot=self.snapshots[self.app_id])
        self.entry_editor.load(self.library.entry_steps(self.app_id,r,snapshot=self.snapshots[self.app_id]) if kind in {"control","region"} else state.get("steps",[]))
        self.entry_dirty=False;self.entry_condition=deepcopy(state.get("condition"))
        self.state_name.setText(state.get("name", ""));self.state_name.setVisible(kind=="region")
        self.hierarchy.clear();self.hierarchy.addItem("自动","auto");self.hierarchy.addItem("区域直属","")
        if kind=="control":
            for other in self.snapshots[self.app_id]["controls"]:
                if other["id"]!=r["id"]:self.hierarchy.addItem(other.get("semantic_name") or other.get("name") or other["id"],other["id"])
            self.hierarchy.setCurrentIndex(max(0,self.hierarchy.findData(r.get("hierarchy_parent","auto"))))
        self.hierarchy.setEnabled(kind=="control")
        self.condition_kind.setCurrentIndex(max(0,self.condition_kind.findData((self.entry_condition or {}).get("kind","exists"))))
        self.pick_condition_button.setEnabled(kind=="region");self.update_entry_summary()
        editable_entry=kind=="control" or kind=="region" and r.get("source")=="manual"
        for widget in (self.edit_entries_button,self.pick_entry_button,self.record_entry_button):widget.setEnabled(editable_entry)
        if kind=="workflow":
            self.flow=deepcopy(r);self.flow_title.setText(r["keyword"]);self.flow_steps.set_parameters(r.get("parameters",{}));self.flow_steps.load(r["steps"]);self.center.setCurrentWidget(self.flow_page)
        elif kind=="usage":self.usage.setPlainText(r.get("body",""));self.usage.setReadOnly(not r.get("editable"));self.center.setCurrentWidget(self.usage)
        elif kind=="reference":self.usage.setPlainText(r.get("body",""));self.usage.setReadOnly(False);self.center.setCurrentWidget(self.usage)
        elif kind in {"region","control","area"}:self.show_reference()
        else:
            index=self.snapshots.get(self.app_id,{})
            self.overview.setPlainText(("选择软件窗口 → 添加区域 → 学习区域\n\n" if kind in {"vault","application","controls-root"} else "")+
                (f"控件：{len(index.get('controls',[]))}    区域：{len(index.get('regions',[]))}    组合动作：{len(index.get('workflows',[]))}\n\n" if index else "")+
                {"vault":"从左侧选择已有软件，或点击“添加软件”。","actions-root":"选择已有组合动作查看步骤，或新建组合动作。","usage-root":"选择文档查看软件使用说明，也可以新建说明。",
                 "references-root":"添加用户手册、测试用例、脑图或网页。首次使用设置中的 API 整理为标准 Markdown；可直接编辑，学习时读取修订后的内容。参考资料只用于推断，不会自动执行其中的步骤。"}.get(kind,"选中区域可查看参考图和前置步骤。"))
            self.center.setCurrentWidget(self.overview)
        self.loading=False;self.dirty=False;self.build_toolbar();self.adapt_details()

    def changed(self,*_):
        if not self.loading:
            self.dirty=True;self.status_label.setText("有未保存修改；Ctrl+S 保存。")

    def edit_native_identity(self,*_):
        if self.loading or self.syncing_identity or self.current.get("kind")!="control":return
        self.changed()
        try:
            location=with_control_identity(json.loads(self.locator.toPlainText()),
                {key:field.text() for key,field in self.native_fields.items()})
        except (ValueError,TypeError) as error:
            self.identity_error=str(error);self.status_label.setText("定位属性尚未同步："+str(error));return
        self.syncing_identity=True
        try:self.locator.setPlainText(json.dumps(location,ensure_ascii=False,indent=2))
        finally:self.syncing_identity=False
        self.locator_snapshot=location;self.identity_error=""
        self.refresh_locator_images()

    def edit_locator(self):
        if self.loading or self.syncing_identity or self.current.get("kind")!="control":return
        self.changed()
        try:
            location=normalize_location(json.loads(self.locator.toPlainText()))
            if is_image_location(location):
                self.locator_snapshot=deepcopy(location);self.identity_error="";self.refresh_locator_images();return
            if not location["Xpath"]:raise ValueError("控件 LOCATION 必须保留 XPath")
            raw_location=deepcopy(location)
            old=normalize_location(self.locator_snapshot)
            identity={}
            for field,key in CONTROL_IDENTITY_KEYS.items():
                top=location.get(key,"");leaf=location["Xpath"][-1].get(key,"")
                top_changed=top!=old.get(key,"")
                leaf_changed=leaf!=(old["Xpath"][-1].get(key,"") if old["Xpath"] else "")
                if top_changed and leaf_changed and top!=leaf:
                    raise ValueError(f"{key} 与 XPath 末级属性不一致，请修正后保存")
                identity[field]=leaf if leaf_changed and not top_changed else top
            location=with_control_identity(location,identity)
        except (ValueError,TypeError) as error:
            self.identity_error=str(error);return
        self.syncing_identity=True
        try:
            for key,field in self.native_fields.items():field.setText(identity[key])
        finally:self.syncing_identity=False
        self.locator_snapshot=raw_location;self.identity_error=""
        self.refresh_locator_images()

    def refresh_locator_images(self):
        self.template_selector.clear()
        if self.current.get("kind")!="control":self.show_locator_image();return
        try:record={**self.record,"location":json.loads(self.locator.toPlainText())}
        except (ValueError,TypeError):self.show_locator_image();return
        for path in reference_images(self.library.directory(self.app_id),record):
            self.template_selector.addItem(path.name,str(path))
        self.show_locator_image()

    def show_locator_image(self,*_):
        path=self.template_selector.currentData()
        self.template_path.setText(path or "没有可用的定位图片")
        self.template_preview.show_image(path)
        if not path:
            self.template_preview.canvas.clear();self.template_preview.canvas.addText("未配置图片，或图片文件不存在。")

    def entry_changed(self):self.entry_dirty=True;self.update_entry_summary();self.changed()

    def edit_entries(self):
        dialog=QDialog(self);dialog.setWindowTitle("前置步骤");dialog.resize(900,660)
        layout=QVBoxLayout(dialog);editor=StepEditor(self.choose_control);editor.load(self.entry_editor.steps);layout.addWidget(editor)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("应用");buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(buttons)
        buttons.accepted.connect(dialog.accept);buttons.rejected.connect(dialog.reject);layout.addWidget(buttons)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            self.entry_editor.load(editor.steps);self.entry_changed()

    def update_entry_summary(self):
        self.entry_summary.setText(" → ".join(s.get("description") or s.get("target",{}).get("label","") for s in self.entry_editor.steps) or "无")
        self.condition_label.setText((self.entry_condition or {}).get("label", "无"))
        self.entry_preview.setPlainText("\n".join(f"{i}. {ACTION_NAMES.get(s['action'],s['action'])} · {s.get('description','')}" for i,s in enumerate(self.entry_editor.steps,1)) or "无前置步骤")

    def build_toolbar(self):
        while self.toolbar.count():
            item=self.toolbar.takeAt(0)
            if item.widget():
                item.widget().hide();item.widget().deleteLater()
        self.context_buttons=[];kind=self.current.get("kind","vault")
        selected=self.selected_nodes();multiple=len(selected)>1;deletable=self.deletable_nodes(selected)
        self.caption.setText(f"已选择 {len(selected)} 项 · 当前预览：{self.selection_caption}" if multiple else self.selection_caption)
        self.details.setEnabled(not multiple);self.center.setEnabled(not multiple)
        actions=[]
        if kind=="vault":actions=[("添加软件",lambda:self.choose_window(new=True))]
        elif kind in {"application","controls-root","area"}:actions=[("添加区域",self.add_region),("查看区域",self.review_layout),("刷新截图",self.capture)]
        elif kind=="region":actions=[("打开区域",self.open_region),("查看区域",self.review_layout),("修改边界",lambda:self.pick_region(modify=True)),("学习区域",self.learn),("添加状态",self.add_variant),("保存",self.save)]
        elif kind=="control":actions=[("验证定位",self.verify),("执行",self.execute_control),("学习展开内容",self.expand),("插入前置步骤",self.insert_entries),("保存",self.save)]
        elif kind in {"actions-root","workflow"}:actions=[("新建",self.new_flow),("参数",self.parameters),("运行",self.run_flow),("运行结果",self.show_results),("保存",self.save)]
        if kind=="workflow":actions.insert(2,("插入前置步骤",self.insert_entries))
        if kind=="region" and dynamic_regions.is_dynamic(self.record):
            actions=[("打开区域",self.open_region),("细分区域",self.subdivide),("重新学习展开内容",self.relearn_dynamic),("保存",self.save)]
        elif kind in {"usage-root","usage"}:actions=[("新建说明",self.new_usage),("保存",self.save)]
        elif kind=="references-root":actions=[("添加参考文件",self.import_reference_files),("添加网页",self.import_reference_url)]
        elif kind=="reference":actions=[("保存",self.save),("重新整理",self.reimport_reference),
                                        ("停用参考" if self.record.get("enabled",True) else "启用参考",self.toggle_reference)]
        if multiple:
            actions=[(f"删除所选（{len(deletable)}）",self.remove),("取消多选",self.clear_multiselection)]
        elif deletable:
            actions.append(("删除所选",self.remove))
        for text,slot in actions:
            value=button(text,slot);value.setEnabled(not self.task and not self.deleting and (not text.startswith("删除所选") or bool(deletable)))
            self.toolbar.addWidget(value);self.context_buttons.append(value)
        # Keep the button a child before changing visibility; an unparented
        # visible widget briefly becomes a native window and can steal focus.
        more=QToolButton(self.toolbar.parentWidget());more.setText("更多");more.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup);menu=QMenu(more)
        set_button_icon(more);more.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        extras=[("导出软件知识",self.export)] if self.app_id else []
        if self.app_id:extras.append(('处理弹窗',self.inspect_popups))
        if kind in {"application","usage-root","usage","control","region"}:
            extras.extend([("添加参考文件",self.import_reference_files),("添加参考网页",self.import_reference_url)])
        if kind in {"application","controls-root","control","region","reference","references-root"}:
            extras.append(("补充控件说明",self.supplement_descriptions))
        if kind=="control" or kind=="region" and self.record.get("source")=="manual":extras.extend([("拾取入口控件",self.pick_entry),("录制步骤",self.record_entries)])
        if kind=="control":
            if self.record.get("reference_evidence"):
                extras.append(("查看参考依据",lambda:self.goto(f"reference:{self.app_id}:{self.record['reference_evidence'][0]['reference_id']}")))
            extras.append(("复制 Python 定位代码",lambda:QApplication.clipboard().setText(self.library.python_code(self.record))))
            if self.record.get("navigation_path") and self.record.get("control_type") in {"MenuControl","PaneControl","GroupControl","ToolBarControl","WindowControl","CustomControl","ListControl","TreeControl","TableControl","DataGridControl"}:
                extras.append(("设为展开子区域",self.promote_region))
        if kind in {"workflow","actions-root"}:extras.append(("导入步骤表",self.import_steps))
        if multiple:extras=[]
        for text,slot in extras:
            action=menu.addAction(text);action.triggered.connect(slot);action.setEnabled(bool(self.app_id) and not self.task)
            action.setIcon(icon(action_icon_name(text)))
        more.setMenu(menu);self.toolbar.addWidget(more);more.setVisible(bool(extras));self.toolbar.addStretch()
        self.toggle_details=button("属性",lambda:self.details.setVisible(not self.details.isVisible()));self.toolbar.addWidget(self.toggle_details)

    def selected_nodes(self):
        return [index.data(ROLE) for index in self.tree.selectionModel().selectedRows(0) if index.data(ROLE)]

    @staticmethod
    def deletable_nodes(nodes):
        return [node for node in nodes if node["kind"] in {"control","reference"} or node["kind"]=="region" and (
            dynamic_regions.is_dynamic(node["record"]) or node["record"].get("source")=="manual"
            and not node["record"].get("classification_only"))]

    def clear_multiselection(self):
        index=self.tree.currentIndex()
        if index.isValid():
            self.tree.selectionModel().select(index,QItemSelectionModel.SelectionFlag.ClearAndSelect|QItemSelectionModel.SelectionFlag.Rows)

    def tree_menu(self,position):
        index=self.tree.indexAt(position)
        if not index.isValid():return
        index=index.siblingAtColumn(0)
        if not self.tree.selectionModel().isSelected(index):self.goto(index.data(ROLE)["key"])
        count=len(self.deletable_nodes(self.selected_nodes()))
        menu=QMenu(self.tree);action=menu.addAction(f"删除所选（{count}）")
        action.setIcon(icon("trash"))
        action.setEnabled(bool(count) and not self.task and not self.deleting);action.triggered.connect(self.remove)
        menu.exec(self.tree.viewport().mapToGlobal(position))

    def adapt_details(self):
        relevant=self.current.get("kind") in {"region","control","area","workflow"}
        self.details.setVisible(relevant and self.width()>=1280)

    def resizeEvent(self,event):
        super().resizeEvent(event)
        if hasattr(self,"details"):self.adapt_details()

    def choose_control(self):
        index=self.snapshots.get(self.app_id,{})
        pages={p["id"]:p for p in index.get("pages",[])}
        regions=[]
        for region in index.get("regions",[]):
            state=region.get("state_name") or pages.get(region.get("page_id"),{}).get("region_state",{}).get("name","")
            regions.append({**region,"name":region.get("name","")+(" / "+state if state else "")})
        dialog=ControlChooser(index.get("controls",[]),regions,self)
        return dialog.selected_record() if dialog.exec()==QDialog.DialogCode.Accepted else None

    def inspect_popups(self):
        if self.task:return
        expected=self.bound_window()
        if expected:self.start_task('inspect-popups',dict(window=expected),lambda state:self.show_popups(state,expected))

    def show_popups(self,state,expected):
        from .popup_dialog import PopupDialog
        existing=self.active_dialog
        if isinstance(existing,PopupDialog) and all(getattr(existing,'expected_window',{}).get(k)==expected.get(k)
                                                    for k in ('handle','process_id','class_name')):
            existing.set_state(state);existing.show();return
        def run(operation,popup):
            if self.task:return
            def completed(result):
                dialog.set_state(result)
                if operation!='focus-popup':dialog.show()
            self.start_task(operation,dict(window=expected,popup=popup),completed,hide=operation!='inspect-popups')
        dialog=PopupDialog(state,run,self);dialog.expected_window=expected;self.active_dialog=dialog
        dialog.setWindowTitle('处理软件弹窗 — '+(expected.get('title') or self.app_id))
        def finished(*_):
            if self.active_dialog is dialog:self.active_dialog=None
        dialog.finished.connect(finished);dialog.show()

    def choose_window(self,new=False):
        if self.task or not self.maybe_save():return
        application=None if new else self.snapshots.get(self.app_id,{}).get("application")
        dialog=WindowDialog(self,application=application)
        if dialog.exec()!=QDialog.DialogCode.Accepted:return
        app_id=dialog.app_id;self.library.bind(app_id,dialog.name.text().strip(),dialog.window);self.windows[app_id]=dialog.window
        self.reload("app:"+app_id)
        self.start_task("capture",{"window":dialog.window},lambda result:self.captured(app_id,result),hide=True)

    def bound_window(self):
        if self.app_id in self.windows:return self.windows[self.app_id]
        if self.app_id:
            from ..mcp import native_windows
            app=self.snapshots[self.app_id]["application"]
            matches=[w for w in native_windows.matching_windows(app.get("window_name","")) if w["visible"] and w["class_name"]==app.get("window_class")]
            if len(matches)==1:self.windows[self.app_id]=matches[0];return matches[0]
        self.choose_window();return None

    def capture(self):
        window=self.bound_window()
        if window:
            app_id=self.app_id;self.start_task("capture",{"window":window},lambda r:self.captured(app_id,r),hide=True)

    def captured(self,app_id,result):
        self.library.save_capture(app_id,result);self.reload("app:"+app_id)

    def add_region(self):self.pick_region()

    def pick_region(self,modify=False,variant=None):
        if not self.app_id or self.task or not self.maybe_save():return
        window=self.bound_window()
        if not window:return
        app_id=self.app_id;record=deepcopy(self.record);scope=self.scoped if modify else self.library
        def done(result):
            target_scope=scope
            title=record.get("name","") if modify else ""
            if not title and not variant:
                title,ok=QInputDialog.getText(self,"区域名称","名称",text=result["selection"]["native_name"] or "区域")
                if not ok:return
            if variant:
                identifier=self.library.save_variant(app_id,record,result,**variant)
            else:
                if result["capture"].get("window_scope",{}).get("kind")=="owned" and not modify:
                    target_scope=self.library.scoped("window-"+uuid4().hex[:10])
                try:
                    identifier=target_scope.save_picked_region(app_id,result,region_id=record["id"] if modify else "",name=title,
                                                        description=record.get("description","") if modify else "")
                except ValueError as error:
                    if modify or "重叠" not in str(error):raise
                    conflicts=[r for r in target_scope.regions(app_id) if self.overlapping(r["normalized_rect"],result["selection"]["normalized_rect"])]
                    if len(conflicts)!=1:raise
                    existing=conflicts[0];box=QMessageBox(self);box.setWindowTitle("区域重叠")
                    box.setText("与“"+existing["name"]+"”重叠。若是另一 Tab 的内容，请添加状态。")
                    update=box.addButton("更新已有区域",QMessageBox.ButtonRole.AcceptRole)
                    state=box.addButton("添加状态",QMessageBox.ButtonRole.ActionRole)
                    box.addButton("取消",QMessageBox.ButtonRole.RejectRole);box.exec()
                    if box.clickedButton()==update:
                        identifier=target_scope.save_picked_region(app_id,result,region_id=existing["id"],name=existing["name"],description=existing.get("description",""))
                    elif box.clickedButton()==state:
                        from .state_dialog import StateDialog
                        dialog=StateDialog(self);self.active_dialog=dialog
                        try:
                            if dialog.exec()!=QDialog.DialogCode.Accepted:return
                            identifier=self.library.save_variant(app_id,existing,result,**dialog.values())
                        finally:self.active_dialog=None
                    else:return
            self.reload(f"region:{app_id}:{identifier}")
        self.start_task("pick-region",{"window":window},done,hide=True)

    @staticmethod
    def overlapping(a,b):
        return min(a["right"],b["right"])>max(a["left"],b["left"]) and min(a["bottom"],b["bottom"])>max(a["top"],b["top"])

    def add_variant(self):
        if self.current.get("kind") not in {"region","area"}:return
        if not self.maybe_save():return
        from .state_dialog import StateDialog
        dialog=StateDialog(self);self.active_dialog=dialog
        if dialog.exec()==QDialog.DialogCode.Accepted:
            values=dialog.values();self.active_dialog=None;self.pick_region(variant=values)
        else:self.active_dialog=None

    def subdivide(self):
        if not self.maybe_save():return
        record=self.record;observation=record.get("visual_observation",{})
        path=self.library.asset(self.app_id,observation.get("context_image",""))
        if path is None:self.error("无法细分","学习参考图缺失，请重新学习展开内容");return
        from .boundary_dialog import BoundaryDialog
        dialog=BoundaryDialog(path,record,self)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            try:
                identifier=self.library.save_dynamic_region(self.app_id,dialog.name.text(),dialog.view.rectangle,parent_region_id=record["id"])
                self.reload(f"region:{self.app_id}:{identifier}")
            except Exception as error:self.error("无法保存子区域",error)

    def promote_region(self):
        if not self.maybe_save():return
        try:
            identifier=self.library.make_dynamic_region(self.app_id,self.record["id"])
            self.reload(f"region:{self.app_id}:{identifier}")
        except Exception as error:self.error("无法建立子区域",error)

    def relearn_dynamic(self):
        route=self.record.get("navigation_path",[])
        record=next((r for r in self.snapshots[self.app_id]["controls"] if route and r["id"]==route[-1]),None)
        if record:self.expand_record(record)
        else:self.error("无法重新学习","缺少展开入口，请选择原入口控件重新学习")

    def pick_control(self,callback):
        window=self.bound_window()
        if not window:return
        app_id=self.app_id;scope=self.scoped
        def done(result):
            record=scope.save_captured_control(app_id,result)
            # Keep unsaved entry edits while refreshing the captured record catalogue.
            self.snapshots[app_id]=self.library.read(app_id);callback(record)
        self.start_task("pick-region",{"window":window,"match_mode":"identity"},done,hide=True)

    def pick_entry(self):
        if not self.app_id:return
        def added(record):
            self.entry_editor.steps.append(workflows.click_step(record));self.entry_editor.render();self.entry_changed()
            self.details.setVisible(True);self.details.setCurrentIndex(1)
        self.pick_control(added)

    def pick_condition(self):
        def selected(record):
            self.entry_condition=dict(location=record["location"],window_scope=record.get("window_scope",{"kind":"main"}),
                                      label=record.get("semantic_name") or record["name"],kind="selected" if record["control_type"]=="TabItemControl" else "exists")
            self.condition_kind.setCurrentIndex(max(0,self.condition_kind.findData(self.entry_condition["kind"])));self.update_entry_summary();self.changed()
        self.pick_control(selected)

    def record_entries(self):
        window=self.bound_window()
        if not window:return
        app_id=self.app_id;scope=self.scoped
        def done(result):
            from ..entry_recording import archive_actions
            steps=archive_actions(scope,app_id,result["actions"])
            self.snapshots[app_id]=self.library.read(app_id);self.entry_editor.steps.extend(steps);self.entry_editor.render();self.entry_changed()
            self.details.setVisible(True);self.details.setCurrentIndex(1)
        self.start_task("record-steps",{"window":window},done,hide=True)

    def open_region(self):
        if not self.maybe_save():return
        window=self.bound_window()
        if window:self.start_task("open-region",dict(vault=str(self.library.root),app_id=self.app_id,page_id=self.record.get("page_id",PAGE_ID),region_id=self.record["id"],window=window),lambda r:self.status_label.setText(r["message"]),hide=True)

    def review_layout(self,*,for_learning=False):
        if not self.app_id or self.task or not self.maybe_save():return False
        from .layout_dialog import LayoutDialog
        app_id=self.app_id;key=self.current["key"]
        dialog=LayoutDialog(self.library,app_id,self.snapshots[app_id],self,
                            page_id=self.record.get("page_id",PAGE_ID),
                            region_id=self.record.get("id","") if self.current.get("kind") in {"region","area"} else "",
                            for_learning=for_learning)
        accepted=dialog.exec()==QDialog.DialogCode.Accepted
        if dialog.edit_region_id:
            self.goto(f"region:{app_id}:{dialog.edit_region_id}");self.pick_region(modify=True)
        elif accepted:
            self.snapshots[app_id]=self.library.read(app_id)
            if not for_learning:self.reload(key)
        return accepted

    def learn(self):
        if self.current.get("kind")!="region" or not self.maybe_save():return
        window=self.bound_window()
        if not window:return
        scope=self.scoped;app_id=self.app_id;record=deepcopy(self.record)
        if not scope.page(app_id).get("layout_confirmed") and not self.review_layout(for_learning=True):return
        try:layout=scope.learning_layout(app_id,window,region_id=record["id"])
        except Exception as error:self.error("无法学习",error);return
        self.start_task("learn",dict(layout=layout,window=window,window_name=window["title"],vault=str(self.library.root),
                                    region_id=record["id"],semantic_mode="native"),lambda _r:self.reload(f"region:{app_id}:{record['id']}"),hide=True)

    def verify(self):
        if not self.maybe_save():return
        window=self.bound_window()
        if window:self.start_task("verify",dict(vault=str(self.library.root),app_id=self.app_id,
            **{key:deepcopy(self.record.get(key)) for key in ("name","location","image","image_variants")},
            window_scope=self.record.get("window_scope",{"kind":"main"}),window=window),lambda r:self.status_label.setText(r["message"]),hide=True)

    def execute_control(self):
        if not self.maybe_save():return
        window=self.bound_window()
        if not window:return
        app_id,record=self.app_id,deepcopy(self.record)
        route=self.library.entry_steps(app_id,record,snapshot=self.snapshots[app_id])
        summary=" → ".join(s.get("description","") for s in route)
        if summary and QMessageBox.question(self,"执行控件","前置步骤："+summary+"\n然后点击："+(record.get("semantic_name") or record.get("name","")))!=QMessageBox.StandardButton.Yes:return
        self.start_task("click",dict(vault=str(self.library.root),app_id=app_id,page_id=record.get("page_id",PAGE_ID),control_id=record["id"],window=window,include_region_entry=True),lambda r:self.display_result(r),hide=True)

    def expand(self):
        self.expand_record(deepcopy(self.record))

    def expand_record(self,record):
        if not self.maybe_save():return
        window=self.bound_window()
        if not window:return
        app_id=self.app_id
        layout=self.scoped.learning_layout(app_id,window)
        self.start_task("expand",dict(vault=str(self.library.root),app_id=app_id,control_id=record["id"],window=window,
            layout=layout,window_name=window["title"],semantic_mode="native"),lambda _r:self.reload(f"control:{app_id}:{record['id']}"),hide=True)

    def new_flow(self):
        if not self.app_id or not self.maybe_save():return
        name,ok=QInputDialog.getText(self,"新建组合动作","动作标题",text="新建工程")
        if not ok or not name.strip():return
        self.flow=dict(id="workflow-"+uuid4().hex[:12],kind="workflow",app_id=self.app_id,keyword=name.strip(),parameters={},precondition="软件已启动",steps=[])
        self.current={**self.current,"kind":"workflow","record":deepcopy(self.flow)}
        self.loading=True;self.flow_title.setText(name);self.flow_steps.set_parameters({});self.flow_steps.load([]);self.loading=False
        self.center.setCurrentWidget(self.flow_page)
        for i in range(4):self.details.setTabVisible(i,i==3)
        self.details.setCurrentIndex(3);self.adapt_details();self.dirty=True;self.build_toolbar()

    def parameters(self):
        if not self.flow:return
        dialog=ParameterDialog(self.flow.get("parameters",{}),self,editable=True)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            self.flow["parameters"]=dialog.values();self.flow_steps.set_parameters(self.flow["parameters"]);self.changed()

    def insert_entries(self):
        if self.current.get("kind")=="workflow":
            selected,is_check=self.flow_steps.current()
            if not selected:return
            step=next((s for s in self.flow_steps.steps if s is selected or selected in s.get("checks",[])),None)
            records={r["id"]:r for r in self.snapshots[self.app_id]["controls"]}
            record=records.get(selected.get("target",{}).get("control_id"))
            if not record:return
            route=self.library.entry_steps(self.app_id,record,snapshot=self.snapshots[self.app_id])
            if not route:self.status_label.setText("此目标没有前置步骤。");return
            summary="\n".join(f"{i}. {s.get('description','')}" for i,s in enumerate(route,1))
            if QMessageBox.question(self,"插入前置步骤","在当前动作之前插入：\n"+summary)!=QMessageBox.StandardButton.Yes:return
            index=self.flow_steps.steps.index(step);self.flow_steps.steps[index:index]=route
            self.flow_steps.render(route[0]["id"]);self.changed();return
        record=self.record;steps=self.library.entry_steps(self.app_id,record,snapshot=self.snapshots[self.app_id])
        if not steps:self.status_label.setText("此控件没有前置步骤。请选择组合动作并添加目标控件。");return
        self.new_flow()
        if self.flow:
            self.flow_steps.load([*steps,workflows.click_step(record)]);self.changed()

    def run_flow(self):
        if not self.flow or not self.save():return
        window=self.bound_window()
        if not window:return
        parameters=self.flow.get("parameters",{})
        if parameters:
            dialog=ParameterDialog(parameters,self)
            if dialog.exec()!=QDialog.DialogCode.Accepted:return
            parameters=dialog.values()
        self.start_task("workflow",dict(vault=str(self.library.root),app_id=self.app_id,flow_id=self.flow["id"],window=window,parameters=parameters),self.display_result,hide=True)

    def display_result(self,result):
        self.result=result;self.result_tree.clear()
        self.result_summary.setText("通过" if result.get("ok") else "失败："+result.get("error",""))
        self.result_summary.set_icon("status-success" if result.get("ok") else "status-error")
        for index,step in enumerate(result.get("steps",[]),1):
            row=QTreeWidgetItem([f"{index}. {step.get('description') or step.get('action','')}","通过" if step.get("ok") else "未通过"])
            row.setData(0,Qt.ItemDataRole.UserRole,step);self.result_tree.addTopLevelItem(row)
            row.setIcon(0,icon("control"));row.setIcon(1,icon("status-success" if step.get("ok") else "status-error"))
            for check in step.get("checks",[]):
                child=QTreeWidgetItem(["预期："+(check.get("description") or check.get("kind","")),"通过" if check.get("ok") else "未通过"]);child.setData(0,Qt.ItemDataRole.UserRole,check);row.addChild(child)
                child.setIcon(0,icon("check-list"));child.setIcon(1,icon("status-success" if check.get("ok") else "status-error"))
            row.setExpanded(True)
        self.center.setCurrentWidget(self.results_page);self.details.hide()
        if self.result_tree.topLevelItemCount():self.result_tree.setCurrentItem(self.result_tree.topLevelItem(0))

    def show_results(self):self.display_result((self.flow or {}).get("last_run") or self.result)

    def show_steps(self):
        if self.flow:
            self.center.setCurrentWidget(self.flow_page);self.adapt_details()

    def select_result(self,*_):
        item=self.result_tree.currentItem()
        if item:
            data=item.data(0,Qt.ItemDataRole.UserRole)
            text=[data.get("description") or data.get("action") or data.get("kind","")]
            if "expected" in data:text.append("预期："+str(data["expected"]))
            if "observed" in data:text.append("实际："+str(data["observed"]))
            if "exists" in data:text.append("存在："+("是" if data["exists"] else "否"))
            if "visible" in data:text.append("可见："+("是" if data["visible"] else "否"))
            if data.get("error"):text.append("失败原因："+data["error"])
            if data.get("resolved_by"):
                text.append("定位方式："+("图片匹配" if data["resolved_by"]=="image" else "UIA"))
                if data["resolved_by"]=="image":
                    text.append(f"匹配分数：{data.get('score',0):.3f}；唯一性：{data.get('uniqueness',0):.3f}")
            if "input_attempted" in data:text.append("输入已尝试" if data["input_attempted"] else "仅读取状态")
            if data.get("capture_error"):text.append("截图失败："+data["capture_error"])
            self.result_text.setPlainText("\n".join(text));self.result_picture("after")

    def result_picture(self,when):
        item=self.result_tree.currentItem()
        if not item:return
        record=item.data(0,Qt.ItemDataRole.UserRole);observation=record.get(when) or record.get("observation") or self.result.get("failure_observation",{})
        path=self.library.asset(getattr(self,"result_app_id",self.app_id),observation.get("context_image",""))
        rectangle=observation.get("normalized_rect")
        self.result_image.show_image(path,[(rectangle,record.get("description",""),True)] if rectangle else ())

    def show_reference(self):
        r=self.record;page=next((p for p in self.snapshots[self.app_id]["pages"] if p["id"]==r.get("page_id")),{})
        observation=r.get("visual_observation",{})
        ref=observation.get("context_image") or page.get("layout_image") or page.get("image") or r.get("image")
        n=observation.get("normalized_rect") or r.get("normalized_rect")
        path=self.library.asset(self.app_id,ref or "")
        markers=[(n,r.get("semantic_name") or r.get("name",""),True)] if n and ref!=r.get("image") else []
        self.image.show_image(path,markers);self.center.setCurrentWidget(self.image)

    def save(self):
        if self.loading:return True
        if self.deleting:return False
        if self.task:self.status_label.setText("任务运行中，完成后再保存修改。");return False
        kind=self.current.get("kind");app_id=self.app_id;r=self.record
        try:
            if kind=="workflow" and self.flow:
                self.flow.update(keyword=self.flow_title.text().strip(),steps=deepcopy(self.flow_steps.steps))
                saved=self.library.save_workflow(app_id,self.flow);key=f"workflow:{app_id}:{saved['id']}"
            elif kind=="control":
                if self.identity_error:raise ValueError(self.identity_error)
                location=normalize_location(json.loads(self.locator.toPlainText()))
                if not is_image_location(location):
                    location=with_control_identity(location,{key:field.text() for key,field in self.native_fields.items()})
                self.library.update_control(app_id,r["id"],name=self.name.text(),description=self.description.toPlainText(),aliases=r.get("aliases",[]),
                    location=location,match_mode=self.mode.currentData(),
                    entry_steps=self.entry_editor.steps if self.entry_dirty else None,hierarchy_parent=self.hierarchy.currentData())
                key=self.current["key"]
            elif kind=="region":
                if dynamic_regions.is_dynamic(r):
                    self.library.save_dynamic_region(app_id,self.name.text(),r["normalized_rect"],region_id=r["id"],description=self.description.toPlainText())
                    key=self.current["key"];self.dirty=False;self.reload(key);return True
                if r.get("source")!="manual":raise ValueError("历史自动区域请重新添加后维护")
                condition=deepcopy(self.entry_condition)
                if condition:condition["kind"]=self.condition_kind.currentData()
                self.scoped.configure_region(app_id,r["id"],steps=self.entry_editor.steps,condition=condition,state_name=self.state_name.text())
                self.scoped.save_region(app_id,self.name.text(),r["normalized_rect"],r["id"],description=self.description.toPlainText())
                key=self.current["key"]
            elif kind=="usage":
                saved=self.library.save_usage(app_id,self.name.text(),self.usage.toPlainText(),document_id=r.get("id","") if r.get("editable") else "")
                key=f"usage:{app_id}:{saved['document_path']}"
            elif kind=="reference":
                self.library.save_reference(app_id,r["id"],self.usage.toPlainText(),enabled=r.get("enabled",True))
                key=self.current["key"]
            else:return True
        except Exception as error:self.error("无法保存",error);return False
        self.dirty=False;self.status_label.setText("已保存");self.reload(key);return True

    def new_usage(self):
        if not self.app_id or not self.maybe_save():return
        self.current={**self.current,"kind":"usage","record":{"name":"新说明","editable":True,"body":""}}
        self.loading=True;self.name.setText("新说明");self.usage.clear();self.usage.setReadOnly(False);self.loading=False
        self.center.setCurrentWidget(self.usage);self.details.setVisible(True);self.dirty=True;self.build_toolbar()

    def import_steps(self):
        if not self.app_id or not self.maybe_save():return
        text,ok=QInputDialog.getMultiLineText(self,"导入步骤表","粘贴动作标题、步骤和预期的 Markdown 表格")
        if not ok:return
        try:flow=self.library.import_workflow_case(self.app_id,text);self.reload(f"workflow:{self.app_id}:{flow['id']}")
        except Exception as error:self.error("无法导入",error)

    def import_reference_files(self):
        if not self.app_id or not self.maybe_save():return
        paths,_=QFileDialog.getOpenFileNames(self,"添加参考文件","",
            "参考资料 (*.pdf *.docx *.md *.markdown *.txt *.html *.htm *.xmind *.mm *.opml)")
        if paths:self.import_references(paths)

    def supplement_descriptions(self):
        if not self.app_id or not self.maybe_save():return
        kind=self.current.get("kind");app_id=self.app_id;key=self.current["key"]
        identifiers=None
        if kind=="control":identifiers=[self.record["id"]]
        elif kind=="region":
            local=self.record["id"].removeprefix(self.record.get("page_id",PAGE_ID)+".")
            identifiers=[r["id"] for r in self.snapshots[app_id]["controls"]
                         if r.get("page_id")==self.record.get("page_id") and r.get("region_id") in {local,self.record["id"]}]
        def completed(result):self.reload(key,message=result.get("message"))
        self.start_task("supplement-descriptions",dict(vault=str(self.library.root),app_id=app_id,control_ids=identifiers),completed)

    def import_reference_url(self):
        if not self.app_id or not self.maybe_save():return
        url,ok=QInputDialog.getText(self,"添加参考网页","网页地址（http / https）")
        if ok and url.strip():
            if not url.strip().lower().startswith(("http://","https://")):
                self.error("地址无效","请输入 http:// 或 https:// 网页地址");return
            self.import_references([url.strip()])

    def import_references(self,sources,*,force=False):
        app_id=self.app_id
        def imported(result):
            identifiers=result.get("reference_ids",[])
            self.reload(f"reference:{app_id}:{identifiers[-1]}" if identifiers else "references-root:"+app_id)
        self.start_task("import-reference",dict(vault=str(self.library.root),app_id=app_id,sources=sources,force=force),imported)

    def reimport_reference(self):
        if not self.maybe_save():return
        record=self.record
        if QMessageBox.question(self,"重新整理参考资料","使用保存的原始资料重新理解，覆盖当前标准 Markdown？")!=QMessageBox.StandardButton.Yes:return
        asset=self.library.asset(self.app_id,record["source_asset"])
        if not asset:self.error("无法重新整理","原始资料已不存在");return
        self.import_references([str(asset)],force=True)

    def toggle_reference(self):
        if not self.maybe_save():return
        self.library.save_reference(self.app_id,self.record["id"],self.usage.toPlainText(),enabled=not self.record.get("enabled",True))
        self.reload(self.current["key"])

    def export(self):
        if not self.app_id:return
        path,_=QFileDialog.getSaveFileName(self,"导出软件知识",self.app_id+".json","JSON (*.json)")
        if path:
            try:self.library.export(self.app_id,path);self.status_label.setText("已导出软件知识")
            except Exception as error:self.error("无法导出",error)

    def remove(self):
        if self.task or self.deleting:return
        selected=deepcopy(self.selected_nodes());nodes=self.deletable_nodes(selected)
        if not nodes or not self.maybe_save():return
        lines=[f"{self.snapshots.get(n['app_id'],{}).get('application',{}).get('name',n['app_id'])} / "
               +(n['record'].get('semantic_name') or n['record'].get('name') or n['record']['id']) for n in nodes]
        prompt=f"删除选中的 {len(nodes)} 项记录？\n\n"+"\n".join(lines[:12])
        if len(lines)>12:prompt+=f"\n……另有 {len(lines)-12} 项"
        prompt+="\n\n记录移入归档。未选中的子项保留，相关前置步骤和组合动作可能需要重新配置。"
        if len(selected)>len(nodes):prompt+=f"\n另有 {len(selected)-len(nodes)} 项软件、分类或只读区域节点保留。"
        if QMessageBox.question(self,"删除所选",prompt,QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,
                                QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        library=self.library;app_id=self.app_id or nodes[0]["app_id"]
        items=[dict(app_id=n["app_id"],kind=n["kind"],id=n["record"]["id"],page_id=n["record"].get("page_id",PAGE_ID)) for n in nodes]
        self.deleting=True;self.centralWidget().setEnabled(False);self.progress.setRange(0,0);self.progress.show()
        self.status_label.setText(f"正在删除 {len(items)} 项记录…")
        def archive():
            try:return library.archive_many(items),None
            except Exception as error:return None,error
        def finished(value):
            self.deleting=False;self.centralWidget().setEnabled(True);self.progress.hide()
            paths,error=value
            if error:self.error("无法删除",error);self.build_toolbar();return
            self.dirty=False;self.search.clear()
            self.reload("controls-root:"+app_id,message=f"已删除 {len(paths)} 项记录；原记录已归档。")
        self.async_io(archive,finished)

    def configure(self):
        def before(values):
            if Path(values["knowledge_dir"]).resolve()!=self.library.root:
                if self.task:raise ValueError("任务运行中，不能切换知识库")
                if not self.maybe_save():raise ValueError("已取消切换知识库")
        def saved(values):
            path=Path(values["knowledge_dir"]).resolve()
            if path!=self.library.root:
                self.library=KnowledgeLibrary(path);self.windows.clear();self.current={};self.dirty=False;self.reload("vault")
        SettingsDialog(self.library.root,self,before_save=before,on_saved=saved).exec()

    def change_vault(self):
        if self.task or not self.maybe_save():return
        path=QFileDialog.getExistingDirectory(self,"选择知识库",str(self.library.root))
        if path:self.library=KnowledgeLibrary(path);self.windows.clear();self.current={};self.reload("vault")

    def start_task(self,operation,payload,done,*,hide=False):
        if self.task or self.deleting:return
        try:
            self.task=LibraryTask({"operation":operation,**payload,"capture_handshake":hide and operation not in {"pick-region","record-steps"}},
                                 timeout=3600 if operation in {"import-reference","supplement-descriptions","download-image-models"} else 600 if operation in {"learn","expand","record-steps"} else 300 if operation=="pick-region" else 90)
        except Exception as error:self.error("任务无法启动",error);return
        self.task_done=done;self.task_hidden=hide;self.stop_button.setEnabled(True);self.progress.setRange(0,0);self.progress.show()
        process=getattr(self.task,"process",None)
        if hide and getattr(process,"pid",None):
            native_windows.allow_foreground(process.pid)
        self.task_app_id=payload.get("app_id") or payload.get("layout",{}).get("app_id") or self.app_id
        self.task_window=payload.get('window')
        self.was_maximized=self.isMaximized();self.tree.setEnabled(False);self.search.setEnabled(False)
        for value in self.context_buttons:value.setEnabled(False)
        self.connect_button.setEnabled(False)
        if hide:
            if self.active_dialog:self.active_dialog.hide()
            self.showMinimized()
            task=self.task
            def ready():
                if self.task is task and task.folder.exists():
                    (task.folder/"gui-ready").write_text("ready",encoding="ascii")
            QTimer.singleShot(50,ready)
        self.status_label.set_icon("status-busy");self.status_label.setText("任务已启动")
        self.update_model_buttons()

    def poll(self):
        if not self.task:return
        state=self.task.poll();self.status_label.setText(f"{state.get('message','正在处理')} · {state['elapsed_seconds']:.1f} 秒")
        if self.task.operation=="download-image-models" and self.model_dialog:
            self.model_dialog.message.setText(self.status_label.text())
        request=self.task.folder/"capture-request"
        if request.exists():
            try:
                nonce=request.read_text(encoding="ascii");(self.task.folder/"capture-ack").write_text(nonce,encoding="ascii")
            except OSError:pass
        if state["worker_running"]:return
        task,callback=self.task,self.task_done;self.task=None;self.task_done=None
        self.stop_button.setEnabled(False);self.progress.hide();self.connect_button.setEnabled(True)
        self.tree.setEnabled(True);self.search.setEnabled(True)
        keep_target_foreground=task.operation=='focus-popup' and state['state']=='completed'
        if not keep_target_foreground:
            self.showMaximized() if self.was_maximized else self.showNormal()
        self.result_app_id=self.task_app_id
        if self.active_dialog and not keep_target_foreground:self.active_dialog.show()
        try:
            if state["state"]=="completed":
                callback(state["result"]);self.status_label.setText("任务完成" if not state["result"].get("message") else state["result"]["message"])
                self.status_label.set_icon("status-error" if state["result"].get("ok") is False else "status-success")
            elif state["state"]=="cancelled":self.status_label.set_icon("status-pending");self.status_label.setText("已取消本次任务")
            else:
                self.status_label.setText(state.get("error","任务失败"))
                self.status_label.set_icon("status-error")
                if state.get("execution"):self.display_result(state["execution"])
                if state.get('popup_state') and self.task_window:
                    self.show_popups(state['popup_state'],self.task_window)
                else:self.error("任务未完成",state.get("error","任务失败"))
        except Exception as error:self.error("结果未保存",error)
        finally:
            task.close();self.build_toolbar()
            if task.operation=="download-image-models":self.refresh_image_models()
            self.update_model_buttons()

    def stop(self):
        if self.task:self.task.cancel();self.status_label.setText("正在停止当前任务…")

    def closeEvent(self,event):
        if self.deleting:self.status_label.setText("正在完成删除，请稍候再关闭。");event.ignore();return
        if self.task:
            if QMessageBox.question(self,"任务运行中","停止当前任务并关闭？")!=QMessageBox.StandardButton.Yes:event.ignore();return
            self.task.close();self.task=None
        if not self.maybe_save():event.ignore();return
        if self.settings:
            self.settings.setValue("navigation",self.main_split.saveState());self.settings.setValue("workspace",self.work_split.saveState())
        self.timer.stop();event.accept()
