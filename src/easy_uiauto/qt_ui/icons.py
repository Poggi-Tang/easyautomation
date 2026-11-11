"""Shared SVG icons for buttons, navigation, and explicit status indicators."""
from functools import lru_cache
from pathlib import Path
import re

from PySide6.QtCore import QByteArray, QRect, QRectF, QSize, Qt
from PySide6.QtGui import QIcon, QIconEngine, QPainter, QPalette, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QApplication, QLabel, QLineEdit

ICON_DIRECTORY = Path(__file__).resolve().parents[1] / "resources/icons"


class SvgIconEngine(QIconEngine):
    """Render the original vectors at the requested DPI, including state colors."""
    def __init__(self, source):
        super().__init__()
        self.source = source
        self.renderers = {}

    def clone(self):
        return SvgIconEngine(self.source)

    def renderer(self, mode):
        color = None
        if mode == QIcon.Mode.Disabled:
            color = "#9CA3AF"
        elif mode == QIcon.Mode.Selected:
            # Some Windows palettes use a pale selection background.
            if QApplication.palette().color(QPalette.ColorRole.Highlight).lightness() < 150:
                color = "#FFFFFF"
        if color not in self.renderers:
            svg = re.sub(r' color="#[0-9A-Fa-f]{6}"', f' color="{color}"', self.source, count=1) if color else self.source
            self.renderers[color] = QSvgRenderer(QByteArray(svg.encode("utf-8")))
        return self.renderers[color]

    def paint(self, painter, rect, mode, state):
        side = min(rect.width(), rect.height())
        square = QRectF(rect.x()+(rect.width()-side)/2, rect.y()+(rect.height()-side)/2, side, side)
        self.renderer(mode).render(painter, square)

    def pixmap(self, size, mode, state):
        pixmap = QPixmap(size)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        self.paint(painter, QRect(0, 0, size.width(), size.height()), mode, state)
        painter.end()
        return pixmap

    def scaledPixmap(self, size, mode, state, scale):
        pixels = QSize(round(size.width()*scale), round(size.height()*scale))
        pixmap = self.pixmap(pixels, mode, state)
        pixmap.setDevicePixelRatio(scale)
        return pixmap


@lru_cache(maxsize=64)
def icon(name):
    if not name:
        return QIcon()
    return QIcon(SvgIconEngine((ICON_DIRECTORY / (name + ".svg")).read_text(encoding="utf-8")))


ACTION_ICONS = {
    "切换知识库": "folder-open", "选择文件": "folder-open", "浏览": "folder-open",
    "选择窗口": "window", "添加软件": "window", "处理弹窗": "window", "查看弹窗": "window",
    "设置": "settings", "新增": "add", "新建说明": "add", "添加动作": "add", "添加参数": "add",
    "新建": "workflow", "保存": "save", "应用": "save", "保存子区域": "save",
    "编辑前置步骤": "edit", "修改边界": "edit", "修改所选区域": "edit", "补充控件说明": "edit",
    "刷新": "refresh", "重新检查": "refresh", "重新整理": "refresh", "获取": "refresh",
    "重新学习展开内容": "refresh", "下载 / 修复": "download",
    "添加参考文件": "import", "导入步骤表": "import", "导出软件知识": "export",
    "删除所选": "trash", "移除": "minus", "移除参数": "minus",
    "执行": "play", "运行": "play", "停止": "stop", "停止下载": "stop",
    "更多": "more", "属性": "sliders", "参数": "sliders", "返回步骤": "back",
    "关闭所选弹窗": "close", "取消多选": "close", "上移": "arrow-up", "下移": "arrow-down",
    "复制 Python 定位代码": "copy", "停用参考": "toggle-on", "启用参考": "toggle-off",
    "刷新截图": "capture", "执行前": "capture", "执行后": "capture",
    "添加区域": "region", "打开区域": "region", "查看区域": "region", "选择区域": "region",
    "设为展开子区域": "region", "添加状态": "layers", "学习区域": "scan",
    "确认区域并学习": "scan", "确认区域": "region", "选择": "target",
    "拾取入口": "target", "拾取入口控件": "target", "拾取状态条件": "target",
    "拾取条件控件": "target", "选择条件控件": "target", "选择目标": "target",
    "验证定位": "target-check", "细分区域": "split", "学习展开内容": "expand",
    "插入前置步骤": "route", "录制步骤": "record", "添加预期": "check-list",
    "运行结果": "results", "添加网页": "globe", "添加参考网页": "globe",
    "查看参考依据": "documents",
}


def action_icon_name(text):
    if text.startswith("标注模型："):
        return "model-chip"
    return ACTION_ICONS.get(text.split("（", 1)[0], "")


def set_button_icon(button, name=None, *, size=18):
    name = action_icon_name(button.text()) if name is None else name
    button.setIcon(icon(name))
    button.setIconSize(QSize(size, size))
    button.setProperty("iconName", name)


def decorate_button_box(box):
    for button in box.buttons():
        set_button_icon(button)


def decorate_search(field):
    action = field.addAction(icon("search"), QLineEdit.ActionPosition.LeadingPosition)
    action.setEnabled(False)


STATUS_ICONS = {
    "已连接": "status-success", "已就绪": "status-success", "定位已验证": "status-success",
    "图片定位已验证": "status-success", "已整理": "status-success", "通过": "status-success",
    "未连接": "status-pending", "已采集": "status-pending", "待检查": "status-pending",
    "待复核": "status-warning", "待确认": "status-warning", "功能待确认": "status-warning",
    "图片定位待验证": "status-warning", "定位待复核": "status-warning", "待完善": "status-warning",
    "已隔离": "status-blocked", "已停用": "toggle-off", "可运行": "play",
    "有前置": "route", "展开区域": "expand",
}


def status_icon_name(text):
    return STATUS_ICONS.get(text, "")


class IconLabel(QLabel):
    """Keep ordinary QLabel text/accessibility while painting a separate icon."""
    def __init__(self, text="", *, icon_name="", parent=None):
        super().__init__(text, parent)
        self.icon_name = ""
        self.set_icon(icon_name)

    def set_icon(self, name):
        self.icon_name = name
        self.setContentsMargins(22 if name else 0, 0, 0, 0)
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.icon_name:
            painter = QPainter(self)
            mode = QIcon.Mode.Normal if self.isEnabled() else QIcon.Mode.Disabled
            icon(self.icon_name).paint(painter, QRect(0, max(0, (self.height()-16)//2), 16, 16), mode=mode)
            painter.end()
