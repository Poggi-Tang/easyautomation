"""Review saved region boundaries without recapturing or operating the app."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QHBoxLayout,
    QLabel, QSplitter, QTreeWidget, QTreeWidgetItem, QVBoxLayout)

from ..mcp import manual_layout
from .widgets import ImageView, button
from .icons import icon, decorate_button_box, decorate_search


class LayoutDialog(QDialog):
    def __init__(self, library, app_id, snapshot, parent=None, *, page_id="",
                 region_id="", for_learning=False):
        super().__init__(parent)
        self.library=library;self.app_id=app_id;self.snapshot=snapshot
        self.for_learning=for_learning;self.edit_region_id="";self.confirmed_page_id=""
        self.initial_region=region_id;self.regions=[];self.page={}
        self.setWindowTitle("确认区域" if for_learning else "查看区域");self.resize(1100,740)
        layout=QVBoxLayout(self)
        hint=QLabel("查看已保存的布局图和区域边界；此处不会刷新截图或操作目标软件。")
        hint.setWordWrap(True);layout.addWidget(hint)
        row=QHBoxLayout();row.addWidget(QLabel("区域状态"));self.pages=QComboBox()
        self.pages.setAccessibleName("选择区域状态");row.addWidget(self.pages,1);layout.addLayout(row)
        self.message=QLabel();self.message.setWordWrap(True);layout.addWidget(self.message)
        split=QSplitter();split.setAccessibleName("区域边界总览");layout.addWidget(split,1)
        self.tree=QTreeWidget();self.tree.setAccessibleName("已保存区域")
        self.tree.setHeaderLabels(["区域","边界"]);self.tree.setColumnWidth(0,160);split.addWidget(self.tree)
        self.image=ImageView();split.addWidget(self.image);split.setSizes([300,800])
        footer=QHBoxLayout();self.modify=button("修改所选区域",self.edit_region);footer.addWidget(self.modify)
        footer.addStretch();self.buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        self.confirm=self.buttons.button(QDialogButtonBox.StandardButton.Ok)
        self.confirm.setText("确认区域并学习" if for_learning else "确认区域")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("返回")
        decorate_button_box(self.buttons)
        self.buttons.accepted.connect(self.accept);self.buttons.rejected.connect(self.reject)
        footer.addWidget(self.buttons);layout.addLayout(footer)
        for page in snapshot.get("pages",[]):
            regions=self.page_regions(page["id"])
            if not regions:continue
            label=page.get("region_state",{}).get("name") or "、".join(r["name"] for r in regions)
            self.pages.addItem(icon("layers"),label,page["id"])
        self.pages.setCurrentIndex(max(0,self.pages.findData(page_id)))
        self.pages.setEnabled(not for_learning)
        self.pages.currentIndexChanged.connect(self.show_page)
        self.tree.currentItemChanged.connect(self.show_selection)
        self.show_page()

    def page_regions(self,page_id):
        return [r for r in self.snapshot.get("regions",[]) if r.get("page_id")==page_id
                and r.get("source")=="manual" and not r.get("classification_only")]

    def show_page(self,*_):
        page_id=self.pages.currentData()
        self.page=next((p for p in self.snapshot.get("pages",[]) if p["id"]==page_id),{})
        self.regions=self.page_regions(page_id);self.tree.clear();selected=None
        size=self.page.get("observed_window_size",{})
        for region in self.regions:
            r=region["normalized_rect"];width,height=size.get("width",0),size.get("height",0)
            pixels=[round(r[k]*(width if k in {"left","right"} else height)) for k in ("left","top","right","bottom")]
            item=QTreeWidgetItem([region["name"],", ".join(map(str,pixels))])
            item.setData(0,Qt.ItemDataRole.UserRole,region["id"]);self.tree.addTopLevelItem(item)
            item.setIcon(0,icon("region"))
            if region["id"]==self.initial_region:selected=item
        self.tree.setCurrentItem(selected or self.tree.topLevelItem(0));self.show_selection()
        status="已确认" if self.page.get("layout_confirmed") else "待确认"
        reason="窗口尺寸或类型发生变化，请逐个检查边界；有偏差时先修改。" if self.page.get("requires_layout_review") else "请检查图中边界是否对应要学习的区域。"
        self.message.setText(f"{status} · 布局图 {size.get('width',0)} × {size.get('height',0)} · {len(self.regions)} 个区域\n{reason}")
        try:
            manual_layout.validate_regions(self.regions)
            if not all(self.image.image_size):raise ValueError("布局图缺失，请返回软件节点刷新截图。")
        except ValueError as error:
            self.confirm.setEnabled(False);self.message.setText(str(error))
        else:self.confirm.setEnabled(True)

    def show_selection(self,*_):
        item=self.tree.currentItem();identifier=item.data(0,Qt.ItemDataRole.UserRole) if item else ""
        path=self.library.asset(self.app_id,self.page.get("layout_image",""))
        self.image.show_image(path,[(r["normalized_rect"],r["name"],r["id"]==identifier) for r in self.regions])
        self.modify.setEnabled(bool(identifier))

    def edit_region(self):
        item=self.tree.currentItem()
        if item:
            self.edit_region_id=item.data(0,Qt.ItemDataRole.UserRole);self.reject()

    def accept(self):
        if not self.confirm.isEnabled():return
        try:self.library.scoped(self.page["id"]).confirm_layout(self.app_id)
        except Exception as error:self.message.setText(str(error));return
        self.confirmed_page_id=self.page["id"];super().accept()
