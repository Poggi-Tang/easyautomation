"""Short native dialogs; persistent configuration stays in the shared service."""
from __future__ import annotations

from pathlib import Path
import os

from PySide6.QtCore import QThreadPool, Qt
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QFormLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox, QPushButton,
    QCheckBox, QSpinBox, QPlainTextEdit, QTableWidget, QTableWidgetItem, QTreeWidget, QTreeWidgetItem, QVBoxLayout)

from ..mcp import configuration, native_windows, knowledge
from .widgets import IoJob, button
from .icons import icon, decorate_button_box, decorate_search


class WindowDialog(QDialog):
    def __init__(self, parent=None, *, application=None):
        super().__init__(parent);self.setWindowTitle("选择软件窗口");self.resize(760,500)
        self.application=application or {};self.window=None;self.windows=[]
        layout=QVBoxLayout(self)
        form=QFormLayout();self.name=QLineEdit(self.application.get("name",""));form.addRow("软件名称",self.name);layout.addLayout(form)
        row=QHBoxLayout();self.search=QLineEdit();self.search.setPlaceholderText("搜索窗口标题或类名");self.search.setAccessibleName("搜索窗口")
        decorate_search(self.search)
        row.addWidget(self.search);row.addWidget(button("刷新",self.refresh));layout.addLayout(row)
        self.table=QTableWidget(0,3);self.table.setHorizontalHeaderLabels(["窗口标题","类名","进程"])
        self.table.setColumnWidth(0,440);self.table.setColumnWidth(1,180)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers);layout.addWidget(self.table)
        self.table.itemSelectionChanged.connect(self.selected);self.table.itemDoubleClicked.connect(lambda *_:self.accept())
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("选择窗口");buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(buttons)
        buttons.accepted.connect(self.accept);buttons.rejected.connect(self.reject);layout.addWidget(buttons)
        self.search.textChanged.connect(self.filter);self.refresh()

    def refresh(self):
        self.windows=[w for w in native_windows.enumerate_windows(visible_only=True)
                      if w["title"] and w["process_id"]!=os.getpid()]
        self.table.setRowCount(len(self.windows))
        for i,w in enumerate(self.windows):
            for c,value in enumerate((w["title"],w["class_name"],str(w["process_id"]))):self.table.setItem(i,c,QTableWidgetItem(value))
            self.table.item(i,0).setIcon(icon("window"))
        self.filter(self.search.text())

    def filter(self,text):
        for i,w in enumerate(self.windows):self.table.setRowHidden(i,text.casefold() not in (w["title"]+" "+w["class_name"]).casefold())

    def selected(self):
        i=self.table.currentRow()
        if i>=0:
            self.window=self.windows[i]
            if not self.name.text():self.name.setText(self.window["title"])

    def accept(self):
        self.selected()
        if not self.window or not self.name.text().strip():return
        super().accept()

    @property
    def app_id(self):return self.application.get("id") or knowledge.slugify(self.name.text().strip())


class CapturedControlDialog(QDialog):
    def __init__(self, selection, snapshot, parent=None, *, region_id=""):
        super().__init__(parent);self.setWindowTitle("添加控件");self.resize(520,350)
        self.selection=selection;self.snapshot=snapshot
        layout=QVBoxLayout(self);form=QFormLayout();layout.addLayout(form)
        self.region=QComboBox();self.region.addItem("未分区",None)
        pages={p["id"]:p for p in snapshot["pages"]}
        scope=selection.get("window_scope",{"kind":"main"})
        for region in snapshot["regions"]:
            page=pages.get(region.get("page_id"),{})
            if region.get("window_scope",page.get("window_scope",{"kind":"main"}))!=scope:continue
            state=region.get("state_name") or page.get("region_state",{}).get("name","")
            label=region.get("name",region["id"])+(f" / {state}" if state else "")
            self.region.addItem(label,region)
            if region["id"]==region_id:self.region.setCurrentIndex(self.region.count()-1)
        self.name=QLineEdit();self.description=QPlainTextEdit()
        self.name.setAccessibleName("控件名称");self.description.setAccessibleName("控件说明")
        self.region.setAccessibleName("所属区域")
        self.description.setPlaceholderText("说明控件用途、输入要求或使用注意事项")
        form.addRow("所属区域",self.region);form.addRow("名称",self.name);form.addRow("说明",self.description)
        native=QLabel(f"{selection['control_type']} · {selection.get('native_name') or '无原生名称'}")
        native.setWordWrap(True);form.addRow("拾取结果",native)
        self.message=QLabel();self.message.setWordWrap(True);layout.addWidget(self.message)
        self.buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        self.buttons.button(QDialogButtonBox.StandardButton.Save).setText("保存")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(self.buttons);layout.addWidget(self.buttons)
        self.buttons.accepted.connect(self.accept);self.buttons.rejected.connect(self.reject)
        self.region.currentIndexChanged.connect(self.destination_changed);self.destination_changed()
        if region_id and not self.region.currentData():
            self.message.setText("所选控件不在原区域的窗口内，请选择对应窗口的区域，或保存为未分区。")

    def destination_changed(self):
        from ..knowledge_library import KnowledgeLibrary, PAGE_ID
        region=self.region.currentData() or {};self.match_error=""
        try:
            old=KnowledgeLibrary.matching_captured_control(self.snapshot["controls"],self.selection,
                region.get("page_id",PAGE_ID),region.get("navigation_path",[]))
        except ValueError as error:old={};self.match_error=str(error)
        if not self.name.isModified():
            self.name.setText(old.get("semantic_name") or self.selection.get("native_name") or self.selection["control_type"])
        if not self.description.document().isModified():self.description.setPlainText(old.get("description",""))
        self.message.setText(self.match_error or ("已存在该控件，保存将更新原记录。" if old else
            "保存后可在属性中调整定位和前置步骤；所属区域的前置步骤会自动继承。"))
        self.buttons.button(QDialogButtonBox.StandardButton.Save).setEnabled(not self.match_error)

    def values(self):
        from ..knowledge_library import PAGE_ID
        region=self.region.currentData() or {}
        return dict(page_id=region.get("page_id",PAGE_ID),region_id=region.get("id",""),
                    name=self.name.text().strip(),description=self.description.toPlainText())

    def accept(self):
        if self.match_error:return
        if not self.name.text().strip():self.message.setText("请输入控件名称");self.name.setFocus();return
        super().accept()


class DeepLearningDialog(QDialog):
    def __init__(self, region, flows, parent=None, *, sessions=(), controls=(), mode=None):
        from ..relearning import MODES
        super().__init__(parent);self.setWindowTitle("深度学习 · "+region.get("name","区域"));self.resize(760,790)
        self.sessions=list(sessions)
        layout=QVBoxLayout(self)
        hint=QLabel("从当前区域探索交互分支，自动保存新区域、控件和组合动作。请使用测试工程；未完成的分支会保留在学习进度中。")
        hint.setWordWrap(True);layout.addWidget(hint);form=QFormLayout();layout.addLayout(form)
        self.mode=QComboBox()
        for value in ("resume","improve","update","new"):self.mode.addItem(MODES[value],value)
        self.mode.setCurrentIndex(self.mode.findData(mode or ("resume" if sessions else "new")))
        self.session=QComboBox()
        for graph in sessions:
            status={"partial":"部分完成","complete":"已完成","running":"待继续"}.get(graph.get("status"),"待检查")
            self.session.addItem(f"{graph.get('updated_at','')} · {status} · {len(graph.get('states',{}))} 个状态",graph["id"])
        if not sessions:self.session.addItem("暂无历史记录，将建立首次学习记录","")
        self.mode_hint=QLabel();self.mode_hint.setWordWrap(True)
        form.addRow("本次目的",self.mode);form.addRow("已有学习记录",self.session);form.addRow(self.mode_hint)
        self.instructions=QPlainTextEdit();self.instructions.setPlaceholderText("学习范围、需要覆盖的业务流程，以及参考资料中的要求")
        self.instructions.setMaximumHeight(110)
        self.inputs=QPlainTextEdit("{}");self.inputs.setPlaceholderText('例如：{"工程名称":"学习测试", "保存路径":"D:/测试工程"}')
        self.inputs.setMaximumHeight(110)
        self.actions=QSpinBox();self.actions.setRange(1,2000);self.actions.setValue(100)
        self.depth=QSpinBox();self.depth.setRange(1,100);self.depth.setValue(24)
        self.reset=QComboBox();self.reset.addItem("未配置；无法返回时保留进度并停止","")
        for flow in flows:self.reset.addItem(flow["keyword"],flow["id"])
        self.commit=QCheckBox("允许新建、保存等测试数据提交（不包括删除或外部发送）")
        form.addRow("学习要求",self.instructions);form.addRow("测试输入（JSON）",self.inputs)
        form.addRow("本次交互步数",self.actions);form.addRow("探索深度",self.depth);form.addRow("返回起点动作",self.reset)
        layout.addWidget(self.commit)
        self.scope=QTreeWidget();self.scope.setHeaderLabels(["指定内容（不勾选表示整个区域）","类型"]);self.scope.setMaximumHeight(150)
        for kind,records in (("control",controls),("workflow",flows)):
            for record in records:
                item=QTreeWidgetItem([record.get("semantic_name") or record.get("keyword") or record.get("name") or record["id"],
                                      "控件" if kind=="control" else "组合动作"])
                item.setData(0,Qt.ItemDataRole.UserRole,(kind,record["id"]));item.setCheckState(0,Qt.CheckState.Unchecked)
                item.setToolTip(0,item.text(0))
                self.scope.addTopLevelItem(item)
        self.scope.header().setStretchLastSection(False)
        self.scope.header().setSectionResizeMode(0,QHeaderView.ResizeMode.Stretch)
        self.scope.header().setSectionResizeMode(1,QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.scope)
        self.pending=QTreeWidget();self.pending.setHeaderLabels(["未完成分支","原因"]);self.pending.setMaximumHeight(150)
        self.pending.setColumnWidth(0,280)
        layout.addWidget(self.pending)
        retry_hint=QLabel("仅在确认上次操作未生效，或测试数据已恢复后，勾选结果不明的分支允许重试。其余未完成分支会自动重新分析。")
        retry_hint.setWordWrap(True);layout.addWidget(retry_hint)
        self.mode.currentIndexChanged.connect(self.update_mode)
        self.session.currentIndexChanged.connect(self.select_session)
        self.select_session();self.update_mode()
        self.message=QLabel();self.message.setWordWrap(True);layout.addWidget(self.message)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("开始深度学习")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(buttons);buttons.accepted.connect(self.accept);buttons.rejected.connect(self.reject);layout.addWidget(buttons)

    def values(self):
        import json
        inputs=json.loads(self.inputs.toPlainText())
        if not isinstance(inputs,dict):raise ValueError("测试输入必须为 JSON 对象")
        selected={"control":[],"workflow":[]}
        if self.mode.currentData()=="improve":
            for n in range(self.scope.topLevelItemCount()):
                item=self.scope.topLevelItem(n)
                if item.checkState(0)==Qt.CheckState.Checked:
                    kind,identifier=item.data(0,Qt.ItemDataRole.UserRole);selected[kind].append(identifier)
        retry=[self.pending.topLevelItem(n).data(0,Qt.ItemDataRole.UserRole) for n in range(self.pending.topLevelItemCount())
               if self.pending.topLevelItem(n).checkState(0)==Qt.CheckState.Checked]
        return dict(inputs=inputs,instructions=self.instructions.toPlainText().strip(),max_actions=self.actions.value(),
                    max_depth=self.depth.value(),allow_commit=self.commit.isChecked(),mode=self.mode.currentData(),
                    session_id=self.session.currentData() if self.mode.currentData()!="new" else "",
                    control_ids=selected["control"],workflow_ids=selected["workflow"],
                    retry_branches=retry if self.mode.currentData()!="new" else [],reset_workflow_id=self.reset.currentData())

    def update_mode(self):
        mode=self.mode.currentData();self.session.setEnabled(mode!="new");self.scope.setEnabled(mode=="improve")
        self.pending.setEnabled(mode!="new")
        self.mode_hint.setText({
            "resume":"保留已完成结果，重新分析受阻分支；可以补充输入和参考资料后继续。请打开原起点或已记录的结果界面。",
            "improve":"重新分析指定内容，更新原控件和组合动作；保留人工修改，并报告冲突。",
            "update":"请先打开并确认当前区域边界。比较新旧控件、更新定位，并验证受影响流程；未找到的控件保留为待定位。",
            "new":"建立新的学习记录；已有控件继续按身份合并。已有学习建议选择继续、改进或适配，避免重复生成流程。"}[mode])
        self.populate_pending()

    def select_session(self):
        import json
        graph=next((g for g in self.sessions if g["id"]==self.session.currentData()),{})
        options=graph.get("options",{})
        self.inputs.setPlainText(json.dumps(options.get("inputs",{}),ensure_ascii=False,indent=2))
        self.instructions.setPlainText(options.get("instructions",""));self.commit.setChecked(options.get("allow_commit",False))
        self.reset.setCurrentIndex(max(0,self.reset.findData(options.get("reset_workflow_id",""))))
        self.populate_pending()

    def populate_pending(self):
        self.pending.clear()
        graph=next((g for g in self.sessions if g["id"]==self.session.currentData()),{})
        for node in graph.get("states",{}).values():
            for branch in node["branches"]:
                if branch["status"] in {"done","cycle"} and self.mode.currentData()=="resume":continue
                item=QTreeWidgetItem([node["name"]+" / "+branch["name"],branch.get("reason") or branch["status"]])
                item.setToolTip(0,item.text(0));item.setToolTip(1,item.text(1))
                item.setData(0,Qt.ItemDataRole.UserRole,branch["id"])
                if branch["status"] in {"uncertain","running","observed","done"}:
                    item.setCheckState(0,Qt.CheckState.Unchecked)
                self.pending.addTopLevelItem(item)

    def accept(self):
        try:self.values()
        except ValueError as error:self.message.setText(str(error));return
        super().accept()


class ParameterDialog(QDialog):
    def __init__(self, parameters, parent=None, *, editable=False):
        super().__init__(parent);self.setWindowTitle("参数" if editable else "运行参数");self.resize(460,260)
        self.editable=editable;layout=QVBoxLayout(self);self.table=QTableWidget(0,2)
        self.table.setHorizontalHeaderLabels(["参数名称","默认值" if editable else "本次输入"])
        self.table.setColumnWidth(0,150);self.table.setColumnWidth(1,240);layout.addWidget(self.table)
        for key,value in parameters.items():
            row=self.table.rowCount();self.table.insertRow(row)
            item=QTableWidgetItem(key)
            if not editable:item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row,0,item);self.table.setItem(row,1,QTableWidgetItem(str(value)))
        if editable:
            row=QHBoxLayout();row.addWidget(button("添加参数",self.add));row.addWidget(button("移除参数",self.remove));row.addStretch();layout.addLayout(row)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("保存" if editable else "运行")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(buttons)
        buttons.accepted.connect(self.accept);buttons.rejected.connect(self.reject);layout.addWidget(buttons)

    def add(self):
        i=self.table.rowCount();self.table.insertRow(i);self.table.setItem(i,0,QTableWidgetItem("参数"+str(i+1)));self.table.setItem(i,1,QTableWidgetItem(""))

    def remove(self):
        if self.table.currentRow()>=0:self.table.removeRow(self.table.currentRow())

    def values(self):
        result={}
        for i in range(self.table.rowCount()):
            key=self.table.item(i,0).text().strip() if self.table.item(i,0) else ""
            if not key or key in result:raise ValueError("参数名称不能为空或重复")
            result[key]=self.table.item(i,1).text() if self.table.item(i,1) else ""
        return result

    def accept(self):
        try:self.values()
        except ValueError as error:QMessageBox.warning(self,"参数无效",str(error));return
        super().accept()


class SettingsDialog(QDialog):
    def __init__(self, current_vault, parent=None, *, before_save=None, on_saved=None):
        super().__init__(parent);self.setWindowTitle("设置");self.resize(660,310)
        self.before_save=before_save or (lambda _values:None);self.on_saved=on_saved or (lambda _values:None)
        self.job=None;self.initial=configuration.load_gui_configuration();self.initial["knowledge_dir"]=str(current_vault)
        layout=QVBoxLayout(self);form=QFormLayout();self.fields={}
        for key,label in (("knowledge_dir","知识库路径"),("api_url","API URL"),("api_key","API Key"),("api_model","API Model")):
            field=QComboBox() if key=="api_model" else QLineEdit()
            if key=="api_model":field.setEditable(True);field.setCurrentText(self.initial.get(key,""))
            else:field.setText(self.initial.get(key,""))
            field.setAccessibleName(label);self.fields[key]=field
            row=QHBoxLayout();row.addWidget(field)
            if key=="knowledge_dir":row.addWidget(button("浏览",self.browse))
            elif key=="api_model":self.fetch_button=button("获取",self.fetch);row.addWidget(self.fetch_button)
            else:spacer=QLabel();spacer.setMinimumWidth(70);row.addWidget(spacer)
            form.addRow(label,row)
        layout.addLayout(form)
        self.message=QLabel("API Key 明文显示；API Model 支持输入和选择。");self.message.setWordWrap(True);layout.addWidget(self.message)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("保存");buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(buttons)
        buttons.accepted.connect(self.save);buttons.rejected.connect(self.reject);layout.addWidget(buttons)

    def values(self):return {k:(v.currentText() if k=="api_model" else v.text()) for k,v in self.fields.items()}

    def browse(self):
        path=QFileDialog.getExistingDirectory(self,"选择知识库目录",self.fields["knowledge_dir"].text())
        if path:self.fields["knowledge_dir"].setText(path)

    def fetch(self):
        if self.job:return
        values=self.values();self.fetch_button.setEnabled(False);self.message.setText("正在获取模型列表…")
        self.job=IoJob(lambda:configuration.fetch_models(values["api_url"],values["api_key"]))
        self.job.signals.done.connect(self.fetched);QThreadPool.globalInstance().start(self.job)

    def fetched(self,result,error):
        self.job=None;self.fetch_button.setEnabled(True)
        if error:self.message.setText(str(error));return
        model=self.fields["api_model"];current=model.currentText();model.clear();model.addItems(result);model.setCurrentText(current or result[0]);self.message.setText(f"已获取 {len(result)} 个模型")

    def save(self):
        try:
            values=self.values()
            directory=Path(values["knowledge_dir"]).expanduser()
            if not directory.is_absolute():raise ValueError("知识库路径须为绝对路径")
            if any(values[k].strip() for k in ("api_url","api_key","api_model")):
                values=configuration.validate_gui_configuration(**values)
                self.before_save(values);configuration.save_gui_configuration(**values)
            else:
                values["knowledge_dir"]=str(directory.resolve());self.before_save(values)
                configuration.save_knowledge_directory(values["knowledge_dir"])
            self.on_saved(values)
        except Exception as error:self.message.setText(str(error));return
        self.accept()
