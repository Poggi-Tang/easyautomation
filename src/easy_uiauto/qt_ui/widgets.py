"""Reusable image, IO, and action editing widgets without live UIA objects."""
from __future__ import annotations

from copy import deepcopy
import json
from collections import OrderedDict
from pathlib import Path
from uuid import uuid4

from PySide6.QtCore import QObject, QRunnable, Signal, Qt, QRectF, QThreadPool
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QFormLayout, QGraphicsScene,
    QGraphicsView, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget, QDialog, QDialogButtonBox)

from .. import workflows
from .icons import icon, set_button_icon, decorate_button_box, decorate_search

ACTION_NAMES = {"click":"点击", "double-click":"双击", "right-click":"右击", "hover":"悬停", "set-text":"输入",
                "press-key":"按键", "hotkey":"组合键", "inspect":"检查", "call":"调用组合动作"}
CHECK_NAMES = {"exists":"存在", "not-visible":"不可见", "text":"文本等于", "window-title":"窗口标题",
               "read-only":"只读", "menu-items":"菜单包含", "node-exists":"包含树节点",
               "node-absent":"不含树节点", "selected":"已选中"}


class Signals(QObject):
    done = Signal(object, object)


class IoJob(QRunnable):
    def __init__(self, function):
        super().__init__(); self.function = function; self.signals = Signals()

    def run(self):
        try: result, error = self.function(), None
        except Exception as failure: result, error = None, failure
        try: self.signals.done.emit(result, error)
        except RuntimeError:
            # Closing the last window can destroy Qt signal objects while a
            # read-only background check is finishing. Its result is no longer needed.
            from shiboken6 import isValid
            if isValid(self.signals): raise


def button(text, slot, *, name="", icon_name=None):
    value = QPushButton(text); value.setAccessibleName(text)
    set_button_icon(value, icon_name)
    if name: value.setObjectName(name)
    value.clicked.connect(slot)
    return value


class ImageView(QGraphicsView):
    def __init__(self, parent=None, *, async_load=False):
        super().__init__(parent); self.canvas = QGraphicsScene(self); self.setScene(self.canvas)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setAccessibleName("区域与控件图片")
        self.image_size = (0, 0); self.fit_mode = True
        self.async_load=async_load;self.image_jobs=[];self.image_cache=OrderedDict()
        self.image_key=None;self.pending_key=None;self.request_serial=0;self.markers=();self.overlays=[]

    def show_image(self, path=None, markers=()):
        self.markers=deepcopy(markers)
        try:
            stat=Path(path).stat() if path else None
            key=(str(path),stat.st_mtime_ns,stat.st_size) if stat else None
        except OSError:key=None
        if key and key==self.image_key:
            self.draw_markers();return
        if key and key==self.pending_key:return
        self.request_serial+=1;serial=self.request_serial;self.pending_key=None
        if key in self.image_cache:
            self.image_cache.move_to_end(key);self.display_image(key,self.image_cache[key]);return
        if key and self.async_load:
            self.pending_key=key;self.canvas.clear();self.overlays=[];self.image_key=None;self.image_size=(0,0)
            self.canvas.addText("正在加载参考图…")
            job=IoJob(lambda:QImage(str(path)));self.image_jobs.append(job)
            def loaded(value,error):
                from shiboken6 import isValid
                self.image_jobs.remove(job)
                if not isValid(self) or serial!=self.request_serial:return
                self.pending_key=None
                self.display_image(key,QPixmap.fromImage(value) if error is None else QPixmap())
            job.signals.done.connect(loaded);QThreadPool.globalInstance().start(job)
        else:self.display_image(key,QPixmap.fromImage(QImage(str(path))) if key else QPixmap())

    def display_image(self,key,pixmap):
        self.canvas.clear();self.overlays=[];self.image_size=(0,0);self.image_key=None
        if pixmap.isNull():
            self.canvas.setSceneRect(QRectF());self.canvas.addText("暂无图片；选择窗口并学习区域后可查看。");return
        self.image_key=key;self.image_cache[key]=pixmap;self.image_cache.move_to_end(key)
        # Bound memory as well as entry count; large desktop captures can exceed 4K.
        while len(self.image_cache)>1 and (len(self.image_cache)>3 or
                sum(p.width()*p.height()*4 for p in self.image_cache.values())>64*1024*1024):
            self.image_cache.popitem(last=False)
        pixmap.setDevicePixelRatio(1.0)
        self.image_size = (pixmap.width(),pixmap.height()); self.canvas.addPixmap(pixmap)
        self.canvas.setSceneRect(QRectF(0,0,pixmap.width(),pixmap.height()))
        self.draw_markers();self.fit()

    def draw_markers(self):
        for item in self.overlays:self.canvas.removeItem(item)
        self.overlays=[]
        for r, label, active in self.markers:
            pen = QPen(QColor("#ff1744" if active else "#007c91"), 2 if active else 1); pen.setCosmetic(True)
            width,height=self.image_size
            x,y=r["left"]*width,r["top"]*height
            self.overlays.append(self.canvas.addRect(x,y,(r["right"]-r["left"])*width,(r["bottom"]-r["top"])*height,pen))
            if label:
                text=self.canvas.addText(label); text.setDefaultTextColor(pen.color()); text.setPos(x,max(0,y-23))
                self.overlays.append(text)

    def fit(self):
        self.fit_mode=True
        self.fitInView(self.canvas.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def resizeEvent(self,event):
        super().resizeEvent(event)
        if self.fit_mode: self.fit()

    def wheelEvent(self,event):
        self.fit_mode=False; self.scale(1.15 if event.angleDelta().y()>0 else 1/1.15,
                                       1.15 if event.angleDelta().y()>0 else 1/1.15)


def control_label(record, regions=()):
    region_id = record.get("region_id", "")
    region = next((r.get("name", "") for r in regions if r["id"] in {region_id, str(record.get("page_id"))+"."+region_id}), "")
    return " / ".join(s for s in (region,record.get("semantic_name") or record.get("name") or record["id"]) if s)


class ControlChooser(QDialog):
    def __init__(self, controls, regions=(), parent=None):
        super().__init__(parent); self.setWindowTitle("选择控件"); self.resize(640,480)
        layout=QVBoxLayout(self); self.search=QLineEdit(); self.search.setPlaceholderText("搜索名称或区域")
        self.search.setAccessibleName("搜索控件"); decorate_search(self.search); layout.addWidget(self.search)
        self.tree=QTreeWidget(); self.tree.setHeaderLabels(["区域 / 控件","类型"]); self.tree.setColumnWidth(0,420)
        layout.addWidget(self.tree)
        self.records={r["id"]:r for r in controls if r.get("location")}
        for r in self.records.values():
            item=QTreeWidgetItem([control_label(r,regions),r.get("control_type","")]);item.setData(0,Qt.ItemDataRole.UserRole,r["id"]);self.tree.addTopLevelItem(item)
            item.setIcon(0,icon("control"))
        self.search.textChanged.connect(lambda text: [self.tree.topLevelItem(i).setHidden(text.casefold() not in self.tree.topLevelItem(i).text(0).casefold()) for i in range(self.tree.topLevelItemCount())])
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("选择");buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(buttons)
        buttons.accepted.connect(lambda: self.accept() if self.tree.currentItem() else None); buttons.rejected.connect(self.reject); layout.addWidget(buttons)
        self.tree.itemDoubleClicked.connect(lambda *_:self.accept())

    def selected_record(self):
        item=self.tree.currentItem()
        return self.records[item.data(0,Qt.ItemDataRole.UserRole)] if item else None


class StepEditor(QWidget):
    changed=Signal()
    def __init__(self, choose_control, parent=None, *, inline=True, choose_workflow=None):
        super().__init__(parent); self.choose_control=choose_control; self.steps=[]; self.loading=False;self.parameters={}
        self.choose_workflow=choose_workflow;self.invalid_values={}
        layout=QVBoxLayout(self);layout.setContentsMargins(0,0,0,0)
        bar=QHBoxLayout()
        for text,slot in (("添加动作",self.add_action),("添加预期",self.add_check),("上移",lambda:self.move(-1)),("下移",lambda:self.move(1)),("移除",self.remove)):
            bar.addWidget(button(text,slot))
        if choose_workflow:bar.insertWidget(1,button("调用组合动作",self.add_call))
        bar.addStretch();layout.addLayout(bar)
        self.tree=QTreeWidget();self.tree.setHeaderLabels(["步骤 / 预期","区域 / 控件","值"])
        self.tree.setColumnWidth(0,190);self.tree.setColumnWidth(1,240)
        self.tree.setAccessibleName("动作步骤与预期");self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        layout.addWidget(self.tree)
        self.details=QWidget();details_layout=QVBoxLayout(self.details)
        form=QFormLayout(); self.kind=QComboBox();self.value=QLineEdit();self.description=QLineEdit();self.target=QLabel("未选择")
        self.value_source=QComboBox();self.value_source.addItem("固定值","")
        form.addRow("动作 / 预期",self.kind);form.addRow("目标",self.target);form.addRow("标题",self.description)
        form.addRow("值来源",self.value_source);form.addRow("输入 / 预期",self.value)
        self.validation_error=QLabel();self.validation_error.setWordWrap(True)
        form.addRow(self.validation_error)
        self.value_source.currentIndexChanged.connect(self.source_changed)
        details_layout.addLayout(form)
        row=QHBoxLayout();row.addWidget(button("选择目标",self.change_target));row.addStretch();details_layout.addLayout(row);details_layout.addStretch()
        if inline:layout.addWidget(self.details)
        self.tree.currentItemChanged.connect(self.select)
        self.kind.currentIndexChanged.connect(self.edit); self.value.textEdited.connect(self.edit);self.description.textEdited.connect(self.edit)

    def load(self, steps):
        self.invalid_values.clear();self.steps=deepcopy(steps);self.render()

    def validated_steps(self):
        if self.invalid_values:raise ValueError("调用组合动作的参数必须是 JSON 对象，请修正输入")
        return deepcopy(self.steps)

    def current(self):
        item=self.tree.currentItem()
        if not item:return None,None
        identifier,is_check=item.data(0,Qt.ItemDataRole.UserRole)
        for step in self.steps:
            if step["id"]==identifier:return step,False
            for check in step.get("checks",[]):
                if check["id"]==identifier:return check,True
        return None,is_check

    def render(self, identifier=None):
        self.loading=True;self.tree.clear();selected=None
        for index,step in enumerate(self.steps,1):
            item=QTreeWidgetItem([f"{index}. {ACTION_NAMES.get(step['action'],step['action'])} · {step.get('description','')}",step.get("target",{}).get("label",""),str(step.get("text",""))])
            if step["action"]=="call":
                item.setText(1,step.get("workflow_id",""));item.setText(2,json.dumps(step.get("parameters",{}),ensure_ascii=False))
            item.setData(0,Qt.ItemDataRole.UserRole,(step["id"],False));self.tree.addTopLevelItem(item)
            item.setIcon(0,icon("control"))
            if step["id"]==identifier:selected=item
            for check in step.get("checks",[]):
                child=QTreeWidgetItem(["预期："+CHECK_NAMES.get(check["kind"],check["kind"]),check.get("target",{}).get("label",""),str(check.get("expected",""))])
                child.setData(0,Qt.ItemDataRole.UserRole,(check["id"],True));item.addChild(child)
                child.setIcon(0,icon("check-list"))
                if check["id"]==identifier:selected=child
            item.setExpanded(True)
        self.loading=False
        if selected:self.tree.setCurrentItem(selected)
        elif self.tree.topLevelItemCount():self.tree.setCurrentItem(self.tree.topLevelItem(0))
        else:self.select()

    def select(self,*_):
        if self.loading:return
        self.loading=True;item,is_check=self.current(); self.kind.clear()
        for key,label in (CHECK_NAMES if is_check else ACTION_NAMES).items():
            if key!="call" or self.choose_workflow:self.kind.addItem(label,key)
        self.kind.setEnabled(bool(item));self.value.setEnabled(bool(item));self.description.setEnabled(bool(item))
        if item:
            self.kind.setCurrentIndex(max(0,self.kind.findData(item.get("kind" if is_check else "action"))))
            self.target.setText(item.get("target",{}).get("label","未选择")); self.description.setText(item.get("description",""))
            value=item.get("expected" if is_check else "text","")
            self.value.setText(", ".join(value) if isinstance(value,list) else "true" if value is True else "false" if value is False else str(value))
            if item.get("action")=="call":
                self.target.setText(item.get("workflow_id",""));self.value.setText(json.dumps(item.get("parameters",{}),ensure_ascii=False))
                if item["id"] in self.invalid_values:self.value.setText(self.invalid_values[item["id"]])
        else:self.value.clear();self.description.clear();self.target.setText("请选择步骤")
        self.value_source.clear();self.value_source.addItem("固定值","")
        for name in self.parameters:self.value_source.addItem("参数："+name,name)
        value=self.value.text();parameter=value[2:-1] if value.startswith("${") and value.endswith("}") else ""
        self.value_source.setCurrentIndex(max(0,self.value_source.findData(parameter)))
        is_call=bool(item and item.get("action")=="call")
        self.value.setReadOnly(bool(self.value_source.currentData()) and not is_call)
        self.value_source.setEnabled(bool(item) and not is_call)
        self.validation_error.setText("请输入有效的 JSON 对象，例如 {}" if item and item["id"] in self.invalid_values else "")
        self.loading=False

    def set_parameters(self,parameters):
        self.parameters=dict(parameters);self.select()

    def source_changed(self,*_):
        if self.loading:return
        parameter=self.value_source.currentData()
        self.value.setReadOnly(bool(parameter))
        if parameter:self.value.setText("${"+parameter+"}")
        self.edit()

    def edit(self,*_):
        if self.loading:return
        item,is_check=self.current()
        if not item:return
        previous=item.get("kind" if is_check else "action")
        action=self.kind.currentData()
        if not is_check and action!=previous and (action=="call" or previous=="call"):
            if action=="call":
                record=self.choose_workflow() if self.choose_workflow else None
                if not record:self.select();return
                item.update(workflow_id=record["id"],parameters={},checks=[],description=record["keyword"])
                item.pop("target",None);item.pop("text",None)
            else:
                record=self.choose()
                if not record:self.select();return
                item.update(target=dict(control_id=record["id"],label=record.get("semantic_name") or record.get("name","")),text="")
                item.pop("workflow_id",None);item.pop("parameters",None)
            item["action"]=action;self.invalid_values.pop(item["id"],None)
            self.render(item["id"]);self.changed.emit();return
        item["kind" if is_check else "action"]=action
        item["description"]=self.description.text()
        if item.get("action")=="call":
            try:
                parameters=json.loads(self.value.text() or "{}")
                if not isinstance(parameters,dict):raise ValueError("parameters must be an object")
            except ValueError:
                self.invalid_values[item["id"]]=self.value.text()
                self.validation_error.setText("请输入有效的 JSON 对象，例如 {}")
                self.changed.emit();return
            self.invalid_values.pop(item["id"],None);self.validation_error.clear()
            item["parameters"]=parameters;item.pop("target",None);item.pop("text",None);item["checks"]=[]
            value=self.value.text()
        else:
            value=self.value.text()
            if is_check and item["kind"]=="read-only":value=value.lower() in {"true","是","只读"}
            if is_check and item["kind"]=="menu-items":value=[v.strip() for v in value.split(",") if v.strip()]
            item["expected" if is_check else "text"]=value
        row=self.tree.currentItem();row.setText(2,str(value))
        row.setText(0,("预期："+CHECK_NAMES.get(item["kind"],item["kind"])) if is_check else
                    f"{self.steps.index(item)+1}. {ACTION_NAMES.get(item['action'],item['action'])} · {item.get('description','')}")
        self.changed.emit()

    def choose(self):return self.choose_control()

    def add_action(self):
        record=self.choose()
        if not record:return
        step=workflows.click_step(record);self.steps.append(step);self.render(step["id"]);self.changed.emit()

    def add_call(self):
        record=self.choose_workflow() if self.choose_workflow else None
        if not record:return
        step=dict(id="call-"+uuid4().hex[:10],action="call",workflow_id=record["id"],description=record["keyword"],parameters={})
        self.steps.append(step);self.render(step["id"]);self.changed.emit()

    def add_check(self):
        item,is_check=self.current()
        if not item:return
        parent=next((s for s in self.steps if s is item or item in s.get("checks",[])),None)
        if parent.get("action")=="call":return
        record=self.choose()
        if not record:return
        check=dict(id="check-"+uuid4().hex[:10],kind="exists",target=dict(control_id=record["id"],label=record.get("semantic_name") or record.get("name","")))
        parent.setdefault("checks",[]).append(check);self.render(check["id"]);self.changed.emit()

    def change_target(self):
        item,_=self.current()
        if not item:return
        if item.get("action")=="call":
            record=self.choose_workflow() if self.choose_workflow else None
            if record:item["workflow_id"]=record["id"];item["description"]=record["keyword"];self.render(item["id"]);self.changed.emit()
            return
        record=self.choose()
        if record:item["target"]=dict(control_id=record["id"],label=record.get("semantic_name") or record.get("name",""));self.render(item["id"]);self.changed.emit()

    def sequence(self,item):
        if item in self.steps:return self.steps
        return next(s["checks"] for s in self.steps if item in s.get("checks",[]))

    def move(self,delta):
        item,_=self.current()
        if not item:return
        items=self.sequence(item); index=items.index(item); target=index+delta
        if 0<=target<len(items):items[index],items[target]=items[target],items[index];self.render(item["id"]);self.changed.emit()

    def remove(self):
        item,_=self.current()
        if item:self.invalid_values.pop(item["id"],None);self.sequence(item).remove(item);self.render();self.changed.emit()
