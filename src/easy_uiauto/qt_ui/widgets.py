"""Reusable image, IO, and action editing widgets without live UIA objects."""
from __future__ import annotations

from copy import deepcopy
from uuid import uuid4

from PySide6.QtCore import QObject, QRunnable, Signal, Qt, QRectF
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QFormLayout, QGraphicsScene,
    QGraphicsView, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget, QDialog, QDialogButtonBox)

from .. import workflows
from .icons import icon, set_button_icon, decorate_button_box, decorate_search

ACTION_NAMES = {"click":"点击", "double-click":"双击", "right-click":"右击", "set-text":"输入",
                "press-key":"按键", "hotkey":"组合键", "inspect":"检查"}
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
    def __init__(self, parent=None):
        super().__init__(parent); self.canvas = QGraphicsScene(self); self.setScene(self.canvas)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setAccessibleName("区域与控件图片")
        self.image_size = (0, 0); self.fit_mode = True

    def show_image(self, path=None, markers=()):
        self.canvas.clear(); self.image_size = (0, 0)
        pixmap = QPixmap(str(path)) if path else QPixmap()
        if pixmap.isNull(): self.canvas.addText("暂无图片；选择窗口并学习区域后可查看。"); return
        pixmap.setDevicePixelRatio(1.0)
        self.image_size = (pixmap.width(),pixmap.height()); self.canvas.addPixmap(pixmap)
        self.canvas.setSceneRect(QRectF(0,0,pixmap.width(),pixmap.height()))
        for r, label, active in markers:
            pen = QPen(QColor("#ff1744" if active else "#007c91"), 2 if active else 1); pen.setCosmetic(True)
            width,height=self.image_size
            x,y=r["left"]*width,r["top"]*height
            self.canvas.addRect(x,y,(r["right"]-r["left"])*width,(r["bottom"]-r["top"])*height,pen)
            if label:
                text=self.canvas.addText(label); text.setDefaultTextColor(pen.color()); text.setPos(x,max(0,y-23))
        self.fit()

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
    def __init__(self, choose_control, parent=None, *, inline=True):
        super().__init__(parent); self.choose_control=choose_control; self.steps=[]; self.loading=False;self.parameters={}
        layout=QVBoxLayout(self);layout.setContentsMargins(0,0,0,0)
        bar=QHBoxLayout()
        for text,slot in (("添加动作",self.add_action),("添加预期",self.add_check),("上移",lambda:self.move(-1)),("下移",lambda:self.move(1)),("移除",self.remove)):
            bar.addWidget(button(text,slot))
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
        self.value_source.currentIndexChanged.connect(self.source_changed)
        details_layout.addLayout(form)
        row=QHBoxLayout();row.addWidget(button("选择目标",self.change_target));row.addStretch();details_layout.addLayout(row);details_layout.addStretch()
        if inline:layout.addWidget(self.details)
        self.tree.currentItemChanged.connect(self.select)
        self.kind.currentIndexChanged.connect(self.edit); self.value.textEdited.connect(self.edit);self.description.textEdited.connect(self.edit)

    def load(self, steps):
        self.steps=deepcopy(steps);self.render()

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
        for key,label in (CHECK_NAMES if is_check else ACTION_NAMES).items():self.kind.addItem(label,key)
        self.kind.setEnabled(bool(item));self.value.setEnabled(bool(item));self.description.setEnabled(bool(item))
        if item:
            self.kind.setCurrentIndex(max(0,self.kind.findData(item.get("kind" if is_check else "action"))))
            self.target.setText(item.get("target",{}).get("label","未选择")); self.description.setText(item.get("description",""))
            value=item.get("expected" if is_check else "text","")
            self.value.setText(", ".join(value) if isinstance(value,list) else "true" if value is True else "false" if value is False else str(value))
        else:self.value.clear();self.description.clear();self.target.setText("请选择步骤")
        self.value_source.clear();self.value_source.addItem("固定值","")
        for name in self.parameters:self.value_source.addItem("参数："+name,name)
        value=self.value.text();parameter=value[2:-1] if value.startswith("${") and value.endswith("}") else ""
        self.value_source.setCurrentIndex(max(0,self.value_source.findData(parameter)))
        self.value.setReadOnly(bool(self.value_source.currentData()));self.value_source.setEnabled(bool(item))
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
        item["kind" if is_check else "action"]=self.kind.currentData()
        item["description"]=self.description.text()
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

    def add_check(self):
        item,is_check=self.current()
        if not item:return
        parent=next((s for s in self.steps if s is item or item in s.get("checks",[])),None)
        record=self.choose()
        if not record:return
        check=dict(id="check-"+uuid4().hex[:10],kind="exists",target=dict(control_id=record["id"],label=record.get("semantic_name") or record.get("name","")))
        parent.setdefault("checks",[]).append(check);self.render(check["id"]);self.changed.emit()

    def change_target(self):
        item,_=self.current()
        if not item:return
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
        if item:self.sequence(item).remove(item);self.render();self.changed.emit()
