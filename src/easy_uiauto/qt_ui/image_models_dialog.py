"""Local annotation models with inline status and persistent file choices."""
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QDialog, QFileDialog, QGridLayout, QHBoxLayout,
                              QLabel, QLineEdit, QVBoxLayout)

from .widgets import button
from .icons import IconLabel, set_button_icon


class ImageModelsDialog(QDialog):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner
        self.paths_dirty = False
        self.loading_paths = False
        self.busy = self.checking = self.downloading = False
        self.report = None
        self.setWindowTitle("本地标注模型")
        self.resize(900, 270)
        layout = QVBoxLayout(self)
        hint = QLabel("选择文件后自动保存；输入路径后按回车或移开焦点保存，清空恢复默认位置。")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        grid = QGridLayout()
        grid.setColumnStretch(2, 1)
        for column, label in enumerate(("模型", "状态", "文件路径")):
            grid.addWidget(QLabel(label), 0, column)
        self.path_fields = {}
        self.status_labels = {}
        self.browse_buttons = []
        for row, (key, label) in enumerate((("yolo", "图形检测 · YOLO"),
                                            ("det", "文字定位 · OCR"),
                                            ("rec", "文字识别 · OCR")), 1):
            grid.addWidget(QLabel(label), row, 0)
            status = IconLabel("检查中…", icon_name="status-busy")
            status.setMinimumWidth(64)
            status.setAccessibleName(label + "状态")
            self.status_labels[key] = status
            grid.addWidget(status, row, 1)
            field = QLineEdit()
            field.setAccessibleName(label + "路径")
            field.setPlaceholderText("留空使用默认位置")
            field.setClearButtonEnabled(True)
            field.textChanged.connect(self.paths_changed)
            field.editingFinished.connect(self.save_paths)
            self.path_fields[key] = field
            grid.addWidget(field, row, 2)
            browse = button("选择文件", lambda _checked=False, model=key: self.choose_file(model))
            browse.setAccessibleName("选择" + label + "文件")
            browse.setAutoDefault(False)
            self.browse_buttons.append(browse)
            grid.addWidget(browse, row, 3)
        layout.addLayout(grid)
        self.runtime = QLabel()
        self.runtime.setWordWrap(True)
        self.runtime.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.runtime.hide()
        layout.addWidget(self.runtime)
        self.message = QLabel("正在检查标注模型…")
        self.message.setWordWrap(True)
        layout.addWidget(self.message)
        row = QHBoxLayout()
        self.refresh_button = button("重新检查", self.refresh)
        self.download_button = button("下载 / 修复", self.download_or_stop)
        self.close_button = button("关闭", self.close)
        for action in (self.refresh_button, self.download_button, self.close_button):
            action.setAutoDefault(False)
        row.addWidget(self.refresh_button)
        row.addWidget(self.download_button)
        row.addStretch()
        row.addWidget(self.close_button)
        layout.addLayout(row)
        self.set_busy(False)

    def set_report(self, report):
        self.report = report
        if not self.paths_dirty:
            self.loading_paths = True
            try:
                for model in report["models"]:
                    field = self.path_fields[model["id"]]
                    configured = model.get("configured_path", "")
                    field.setText(configured)
                    field.setPlaceholderText("留空使用默认位置" if configured else "默认：" + model["path"])
                    field.setToolTip("实际使用路径：" + model["path"])
            finally:
                self.loading_paths = False
            self.message.setText(report["message"])
        self.update_statuses()
        issues = report.get("runtime_issues", [])
        self.runtime.setText("；".join(issues) + "。请安装完整标注运行环境，模型下载仅补齐模型文件。" if issues else "")
        self.runtime.setVisible(bool(issues))

    def update_statuses(self):
        if not self.report:
            return
        names = {"ready": "已就绪", "missing": "未下载", "invalid": "需修复", "unreadable": "无法读取"}
        symbols = {"ready": "status-success", "missing": "status-warning", "invalid": "status-warning", "unreadable": "status-error"}
        for model in self.report["models"]:
            status = self.status_labels[model["id"]]
            changed = self.path_fields[model["id"]].text().strip() != model.get("configured_path", "")
            status.setText("待检查" if changed else names.get(model["state"], model["detail"]))
            status.set_icon("status-pending" if changed else symbols.get(model["state"], "status-pending"))
            status.setToolTip("路径已修改，等待检查" if changed else model["detail"])

    def paths_changed(self):
        if self.loading_paths:
            return
        self.paths_dirty = True
        self.message.setText("按回车或移开焦点自动保存路径。")
        self.update_statuses()
        self.set_busy(self.busy, checking=self.checking, downloading=self.downloading)

    def choose_file(self, key):
        current = self.path_fields[key].text().strip()
        if not current and self.report:
            current = next(model["path"] for model in self.report["models"] if model["id"] == key)
        extension = "YOLO 模型 (*.pt)" if key == "yolo" else "OCR 模型 (*.onnx)"
        filename, _ = QFileDialog.getOpenFileName(self, "选择模型文件", current, extension)
        if filename:
            self.path_fields[key].setText(str(Path(filename)))
            self.save_paths()

    def save_paths(self):
        if self.loading_paths or not self.paths_dirty:
            return True
        return self.owner.save_image_model_paths({key: field.text() for key, field in self.path_fields.items()})

    def refresh(self):
        if self.save_paths():
            self.owner.refresh_image_models()

    def download_or_stop(self):
        if self.downloading:
            self.owner.stop()
        else:
            self.owner.download_image_models()

    def done(self, result):
        self.save_paths()
        super().done(result)

    def set_busy(self, busy, *, checking=False, downloading=False):
        self.busy, self.checking, self.downloading = busy, checking, downloading
        self.refresh_button.setEnabled(not busy and not checking)
        missing = self.report and any(model["state"] != "ready" for model in self.report["models"])
        self.download_button.setText("停止下载" if downloading else "下载 / 修复")
        set_button_icon(self.download_button)
        self.download_button.setAccessibleName(self.download_button.text())
        self.download_button.setEnabled(downloading or (not busy and not checking and not self.paths_dirty and bool(missing)))
        for widget in [*self.path_fields.values(), *self.browse_buttons]:
            widget.setEnabled(not busy)
