"""Draw a child region in the immutable learning frame, in image pixels."""
from PySide6.QtCore import Qt,Signal
from PySide6.QtGui import QPen,QColor
from PySide6.QtWidgets import QDialog,QDialogButtonBox,QLabel,QLineEdit,QVBoxLayout
from .widgets import ImageView
from .icons import icon, decorate_button_box, decorate_search


class RegionImage(ImageView):
    selected=Signal(object)
    def __init__(self,parent=None):
        super().__init__(parent);self.start=None;self.outline=None;self.rectangle=None
        self.setDragMode(self.DragMode.NoDrag)

    def point(self,position):
        point=self.mapToScene(position)
        width,height=self.image_size
        return max(0,min(width,point.x())),max(0,min(height,point.y()))

    def mousePressEvent(self,event):
        if event.button()==Qt.MouseButton.LeftButton and all(self.image_size):
            self.start=self.point(event.position().toPoint());event.accept()
        else:super().mousePressEvent(event)

    def mouseMoveEvent(self,event):
        if self.start:
            end=self.point(event.position().toPoint());x,y=self.start
            if self.outline:self.canvas.removeItem(self.outline)
            pen=QPen(QColor("#ff1744"),2);pen.setCosmetic(True)
            self.outline=self.canvas.addRect(min(x,end[0]),min(y,end[1]),abs(end[0]-x),abs(end[1]-y),pen)
            event.accept()
        else:super().mouseMoveEvent(event)

    def mouseReleaseEvent(self,event):
        if self.start and event.button()==Qt.MouseButton.LeftButton:
            end=self.point(event.position().toPoint());x,y=self.start;self.start=None
            if abs(end[0]-x)>=2 and abs(end[1]-y)>=2:
                width,height=self.image_size
                self.rectangle=dict(left=min(x,end[0])/width,top=min(y,end[1])/height,
                                    right=max(x,end[0])/width,bottom=max(y,end[1])/height)
                self.selected.emit(self.rectangle)
            event.accept()
        else:super().mouseReleaseEvent(event)


class BoundaryDialog(QDialog):
    def __init__(self,path,region,parent=None):
        super().__init__(parent);self.setWindowTitle("细分区域");self.resize(1000,700)
        layout=QVBoxLayout(self);self.name=QLineEdit();self.name.setPlaceholderText("子区域名称");layout.addWidget(self.name)
        self.message=QLabel("在学习参考图的上级区域内拖框；取消不会写入知识库。");layout.addWidget(self.message)
        self.view=RegionImage();self.view.show_image(path,[(region["normalized_rect"],region.get("name",""),False)]);layout.addWidget(self.view,1)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("保存子区域");buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        decorate_button_box(buttons)
        buttons.accepted.connect(self.accept);buttons.rejected.connect(self.reject);layout.addWidget(buttons)

    def accept(self):
        if not self.name.text().strip() or not self.view.rectangle:
            self.message.setText("请填写名称并拖框选择子区域。");return
        super().accept()
