"""State creation requires an explicit entrance and observable condition."""
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QVBoxLayout

from .widgets import StepEditor, button
from .icons import icon, decorate_button_box, decorate_search
from .. import workflows


class StateDialog(QDialog):
    def __init__(self, manager):
        super().__init__(manager);self.manager=manager;self.condition=None
        self.setWindowTitle("添加区域状态");self.resize(780,650)
        layout=QVBoxLayout(self);form=QFormLayout();self.name=QLineEdit();self.name.setPlaceholderText("例如：绘图")
        form.addRow("状态名称",self.name);self.condition_text=QLabel("请选择选中 Tab 或该状态独有控件")
        self.condition_text.setWordWrap(True);form.addRow("状态条件",self.condition_text);layout.addLayout(form)
        row=QHBoxLayout();row.addWidget(button("选择条件控件",self.choose_condition));row.addWidget(button("拾取条件控件",self.pick_condition));row.addStretch();layout.addLayout(row)
        layout.addWidget(QLabel("前置步骤"));self.editor=StepEditor(manager.choose_control);layout.addWidget(self.editor)
        row=QHBoxLayout();row.addWidget(button("拾取入口",self.pick_entry));row.addStretch();layout.addLayout(row)
        self.error=QLabel();self.error.setWordWrap(True);layout.addWidget(self.error)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("选择区域");buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(buttons)
        buttons.accepted.connect(self.accept);buttons.rejected.connect(self.reject);layout.addWidget(buttons)

    def set_condition(self,record):
        self.condition=dict(location=record["location"],window_scope=record.get("window_scope",{"kind":"main"}),
                            label=record.get("semantic_name") or record.get("name", ""),
                            kind="selected" if record.get("control_type")=="TabItemControl" else "exists")
        self.condition_text.setText(self.condition["label"]+"（"+("已选中" if self.condition["kind"]=="selected" else "存在")+"）")
        if not self.name.text():self.name.setText(record.get("semantic_name") or record.get("name", ""))

    def choose_condition(self):
        record=self.manager.choose_control()
        if record:self.set_condition(record)

    def pick_condition(self):self.manager.pick_control(self.set_condition)

    def pick_entry(self):
        def append(record):
            self.editor.steps.append(workflows.click_step(record));self.editor.render()
            if record.get("control_type")=="TabItemControl":self.set_condition(record)
        self.manager.pick_control(append)

    def values(self):return dict(name=self.name.text().strip(),steps=self.editor.steps,condition=self.condition)

    def accept(self):
        if not self.name.text().strip() or not self.editor.steps or not self.condition:
            self.error.setText("请填写状态名称、添加前置步骤，并选择状态条件。");return
        super().accept()
