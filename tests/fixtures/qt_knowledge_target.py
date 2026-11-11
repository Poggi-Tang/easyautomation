"""Disposable Qt target: same-position Tabs and a real project-creation dialog."""
import argparse
import json
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QApplication,QDialog,QDialogButtonBox,QFormLayout,QGroupBox,
    QLabel,QLineEdit,QMainWindow,QPushButton,QSplitter,QTabWidget,QTreeWidget,QTreeWidgetItem,QVBoxLayout,QWidget)


def main():
    parser=argparse.ArgumentParser();parser.add_argument("--events",type=Path,required=True);args=parser.parse_args()
    app=QApplication([]);events=[];args.events.parent.mkdir(parents=True,exist_ok=True)
    def event(event_name,**values):
        events.append(dict(event=event_name,**values));args.events.write_text(json.dumps(events,ensure_ascii=False),encoding="utf-8")
    args.events.write_text("[]",encoding="utf-8")
    window=QMainWindow();window.setWindowTitle("easy-uiauto Qt Acceptance Target");window.resize(850,600);window.move(60,80)
    splitter=QSplitter();window.setCentralWidget(splitter)
    projects=QTreeWidget();projects.setHeaderLabels(["工程名称"]);projects.setAccessibleName("工程树");projects.setObjectName("projectTree")
    splitter.addWidget(projects);tabs=QTabWidget();tabs.setAccessibleName("工作区");splitter.addWidget(tabs);splitter.setSizes([250,600])
    for title,label in (("主页","主页区域"),("绘图","绘图区域")):
        page=QWidget();layout=QVBoxLayout(page);group=QGroupBox(label);group.setAccessibleName(label)
        form=QVBoxLayout(group);form.addWidget(QLabel(label+"内容"));action=QPushButton(title+"动作");action.setAccessibleName(title+"动作")
        action.clicked.connect(lambda checked=False,t=title:event("action",tab=t));form.addWidget(action);layout.addWidget(group);layout.addStretch()
        tabs.addTab(page,title)
    tabs.currentChanged.connect(lambda index:event("tab",name=tabs.tabText(index)))
    def new_project():
        event("new-dialog");dialog=QDialog(window);dialog.setWindowTitle("新建工程");dialog.setModal(True);dialog.resize(400,180)
        form=QFormLayout(dialog);name=QLineEdit();name.setAccessibleName("工程名称");name.setObjectName("projectName");form.addRow("工程名称",name)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("确认");buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        def accept():
            if not name.text().strip():return
            QTreeWidgetItem(projects,[name.text().strip()]);event("confirm",name=name.text().strip());dialog.accept()
        buttons.accepted.connect(accept);buttons.rejected.connect(dialog.reject);form.addRow(buttons);dialog.exec();event("dialog-closed")
    file=window.menuBar().addMenu("文件");file.addAction("新建工程",new_project)
    window.show();print("fixture_ready hwnd="+str(int(window.winId())),flush=True);app.exec()


if __name__=="__main__":main()
