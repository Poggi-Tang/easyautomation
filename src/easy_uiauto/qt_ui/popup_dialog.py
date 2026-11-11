"""Owned-window inspection with explicit focus and normal close actions."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog,QVBoxLayout,QHBoxLayout,QLabel,QTreeWidget,QTreeWidgetItem,QDialogButtonBox
from .widgets import button
from .icons import icon, decorate_button_box, decorate_search


class PopupDialog(QDialog):
    def __init__(self,state,on_action,parent=None):
        super().__init__(parent);self.setWindowTitle('处理软件弹窗');self.resize(760,400)
        self.on_action=on_action
        layout=QVBoxLayout(self);self.summary=QLabel();self.summary.setWordWrap(True);layout.addWidget(self.summary)
        self.tree=QTreeWidget();self.tree.setHeaderLabels(['弹窗标题','状态','窗口类名'])
        self.tree.setColumnWidth(0,370);self.tree.setColumnWidth(1,120);layout.addWidget(self.tree)
        row=QHBoxLayout();self.refresh_button=button('刷新',lambda:self.on_action('inspect-popups',None))
        self.focus_button=button('查看弹窗',lambda:self.act('focus-popup'))
        self.close_button=button('关闭所选弹窗',lambda:self.act('close-popup'));self.close_button.setAutoDefault(False)
        for value in (self.refresh_button,self.focus_button,self.close_button):row.addWidget(value)
        row.addStretch();layout.addLayout(row)
        hint=QLabel('关闭使用弹窗自身的关闭请求。若出现保存确认，请在软件中处理；原动作不会自动重试。')
        hint.setWordWrap(True);layout.addWidget(hint)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.button(QDialogButtonBox.StandardButton.Close).setText('返回');decorate_button_box(buttons);buttons.rejected.connect(self.reject);layout.addWidget(buttons)
        self.tree.currentItemChanged.connect(self.update_actions);self.set_state(state)

    def set_state(self,state):
        self.summary.setText(state.get('message',''));self.tree.clear()
        for popup in state.get('popups',[]):
            status='阻挡当前操作' if popup.get('blocking') else '被其他弹窗阻挡' if not popup.get('enabled',True) else '未阻挡当前窗口'
            item=QTreeWidgetItem([popup.get('title') or '（无标题）',status,popup.get('class_name','')])
            item.setData(0,Qt.ItemDataRole.UserRole,popup);self.tree.addTopLevelItem(item)
            item.setIcon(0,icon("window"))
            item.setIcon(1,icon("status-warning" if popup.get('blocking') else "status-blocked" if not popup.get('enabled',True) else "status-success"))
        if self.tree.topLevelItemCount():self.tree.setCurrentItem(self.tree.topLevelItem(0))
        self.update_actions()

    def update_actions(self,*_):
        item=self.tree.currentItem();popup=item.data(0,Qt.ItemDataRole.UserRole) if item else None
        enabled=bool(popup and popup.get('enabled',True))
        self.focus_button.setEnabled(enabled);self.close_button.setEnabled(enabled)

    def act(self,operation):
        item=self.tree.currentItem()
        if item:self.on_action(operation,item.data(0,Qt.ItemDataRole.UserRole))
