"""Short native dialogs; persistent configuration stays in the shared service."""
from __future__ import annotations

from pathlib import Path
import os

from PySide6.QtCore import QThreadPool, Qt
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton,
    QTableWidget, QTableWidgetItem, QVBoxLayout)

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
