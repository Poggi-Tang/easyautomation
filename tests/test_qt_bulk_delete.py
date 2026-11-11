import pytest

pytest.importorskip("PySide6")
from PySide6.QtCore import QItemSelectionModel,QTimer,Qt
from PySide6.QtTest import QTest
from PySide6.QtGui import QAction
from PySide6.QtWidgets import QAbstractItemView,QApplication,QMenu,QMessageBox,QPushButton

from easy_uiauto.qt_ui.window import KnowledgeWindow
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library,add_control
from test_qt_knowledge_gui import app,settle


@pytest.fixture
def window(app,library):
    first=add_control(library)
    for identifier,name in (("control-second","打开二"),("control-third","保留")):
        knowledge.save_control(library.directory("demo"),{**first,"id":identifier,"semantic_name":name})
    view=KnowledgeWindow(library.root,remember=False)
    settle(app,lambda:"demo" in view.snapshots and not view.jobs)
    view.tree.expandAll();view.show();app.processEvents()
    try:yield view
    finally:
        settle(app,lambda:not view.jobs and not view.deleting)
        view.task=None;view.dirty=False;view.close();app.processEvents()


def select(window,keys,current=None):
    flags=QItemSelectionModel.SelectionFlag
    model=window.tree.selectionModel();model.clearSelection()
    for key in keys:
        index=window.proxy.mapFromSource(window.model.nodes[key].index())
        assert index.isValid()
        model.select(index,flags.Select|flags.Rows)
    key=current or keys[-1]
    model.setCurrentIndex(window.proxy.mapFromSource(window.model.nodes[key].index()),flags.NoUpdate)


def ids(window):return {n["record"].get("id") for n in window.selected_nodes()}


def test_native_ctrl_and_shift_select_rows_and_status_column_preview(app,window):
    assert window.tree.selectionMode()==QAbstractItemView.SelectionMode.ExtendedSelection
    window.goto("control:demo:control-one");app.processEvents()
    second=window.proxy.mapFromSource(window.model.nodes["control:demo:control-second"].index())
    QTest.mouseClick(window.tree.viewport(),Qt.MouseButton.LeftButton,Qt.KeyboardModifier.ControlModifier,
                     window.tree.visualRect(second).center());app.processEvents()
    assert ids(window)=={"control-one","control-second"}
    assert not window.details.isEnabled() and "已选择 2 项" in window.caption.text()
    first=window.proxy.mapFromSource(window.model.nodes["control:demo:control-one"].index())
    QTest.mouseClick(window.tree.viewport(),Qt.MouseButton.LeftButton,pos=window.tree.visualRect(first).center());app.processEvents()
    third=window.proxy.mapFromSource(window.model.nodes["control:demo:control-third"].index())
    QTest.mouseClick(window.tree.viewport(),Qt.MouseButton.LeftButton,Qt.KeyboardModifier.ShiftModifier,
                     window.tree.visualRect(third).center());app.processEvents()
    assert ids(window)=={"control-one","control-second","control-third"}
    window.tree.setCurrentIndex(second.siblingAtColumn(1));app.processEvents()
    assert window.record["id"]=="control-second"


def test_batch_delete_confirms_once_and_preserves_unselected_control(app,window,monkeypatch):
    select(window,["control:demo:control-one","control:demo:control-second"])
    questions=[]
    monkeypatch.setattr(QMessageBox,"question",lambda _parent,_title,text,*a:(questions.append(text),QMessageBox.StandardButton.Yes)[1])
    button=next(b for b in window.findChildren(QPushButton) if b.text()=="删除所选（2）")
    button.click();assert window.deleting and not window.centralWidget().isEnabled()
    settle(app,lambda:not window.deleting and not window.jobs)
    assert len(questions)==1 and "2 项" in questions[0]
    assert [r["id"] for r in window.library.read("demo")["controls"]]==["control-third"]
    assert window.centralWidget().isEnabled() and "已删除 2 项" in window.status_label.text()


def test_cancellation_does_not_change_records_or_selection(window,monkeypatch):
    select(window,["control:demo:control-one","control:demo:control-second"])
    monkeypatch.setattr(QMessageBox,"question",lambda *a:QMessageBox.StandardButton.No)
    before=window.library.read("demo")
    window.remove()
    assert not window.deleting and window.library.read("demo")==before
    assert ids(window)=={"control-one","control-second"}


def test_delete_shortcut_only_acts_in_navigation(app,window,monkeypatch):
    questions=[]
    monkeypatch.setattr(QMessageBox,"question",lambda *a:(questions.append(a),QMessageBox.StandardButton.No)[1])
    window.goto("control:demo:control-one");window.activateWindow();window.name.setFocus();app.processEvents()
    QTest.keyClick(window.name,Qt.Key.Key_Delete);app.processEvents()
    assert not questions
    window.dirty=False;window.tree.setFocus();app.processEvents()
    QTest.keyClick(window.tree,Qt.Key.Key_Delete);app.processEvents()
    assert len(questions)==1 and len(window.library.read("demo")["controls"])==3


def test_filtered_selection_uses_visible_proxy_rows_and_protects_categories(app,window,monkeypatch):
    window.search.setText("打开");window.tree.expandAll();app.processEvents()
    window.tree.setFocus();window.tree.selectAll()
    assert ids(window)>={"control-one","control-second"} and "control-third" not in ids(window)
    questions=[]
    monkeypatch.setattr(QMessageBox,"question",lambda _p,_t,text,*a:(questions.append(text),QMessageBox.StandardButton.Yes)[1])
    window.remove();settle(app,lambda:not window.deleting and not window.jobs)
    assert [r["id"] for r in window.library.read("demo")["controls"]]==["control-third"]
    assert "2 项" in questions[0] and "节点保留" in questions[0]
    assert not window.search.text() and window.current["kind"]=="controls-root"


def test_clearing_multiselection_preserves_unsaved_current_edits(window):
    window.goto("control:demo:control-one")
    window.name.setText("未保存的名称");window.changed()
    index=window.proxy.mapFromSource(window.model.nodes["control:demo:control-second"].index())
    window.tree.selectionModel().select(index,QItemSelectionModel.SelectionFlag.Select|QItemSelectionModel.SelectionFlag.Rows)
    window.clear_multiselection()
    assert window.dirty and window.name.text()=="未保存的名称" and window.details.isEnabled()
    assert ids(window)=={"control-one"}


def test_right_click_preserves_the_existing_multiselection(app,window,monkeypatch):
    select(window,["control:demo:control-one","control:demo:control-second"])
    menus=[]
    class Menu:
        def __init__(self,parent):self.actions=[]
        def addAction(self,text):
            action=QAction(text);self.actions.append(action);return action
        def exec(self,*a):menus.append([action.text() for action in self.actions])
    monkeypatch.setattr("easy_uiauto.qt_ui.window.QMenu",Menu)
    index=window.proxy.mapFromSource(window.model.nodes["control:demo:control-one"].index())
    window.tree.expandAll();window.tree.scrollTo(index);app.processEvents()
    assert window.tree.indexAt(window.tree.visualRect(index).center()).isValid()
    window.tree_menu(window.tree.visualRect(index).center())
    assert ids(window)=={"control-one","control-second"}
    assert menus==[["删除所选（2）"]]


def test_background_failure_reenables_gui_and_keeps_records(app,window,monkeypatch):
    select(window,["control:demo:control-one","control:demo:control-second"])
    monkeypatch.setattr(QMessageBox,"question",lambda *a:QMessageBox.StandardButton.Yes)
    def fail(_items):raise PermissionError("locked record")
    monkeypatch.setattr(window.library,"archive_many",fail)
    errors=[];window.error=lambda title,error:errors.append(str(error))
    window.remove();settle(app,lambda:not window.deleting and not window.jobs)
    assert errors==["locked record"] and window.centralWidget().isEnabled()
    assert len(window.library.read("demo")["controls"])==3


def test_real_context_menu_triggers_one_batch(app,window,monkeypatch):
    select(window,["control:demo:control-one","control:demo:control-second"])
    window.tree.expandAll();app.processEvents()
    monkeypatch.setattr(QMessageBox,"question",lambda *a:QMessageBox.StandardButton.Yes)
    triggered=[]
    def choose():
        menu=QApplication.activePopupWidget()
        if isinstance(menu,QMenu):
            triggered.append(menu.actions()[0].text());menu.actions()[0].trigger();menu.close()
    QTimer.singleShot(30,choose)
    def dismiss():
        menu=QApplication.activePopupWidget()
        if isinstance(menu,QMenu):menu.close()
    QTimer.singleShot(300,dismiss)
    index=window.proxy.mapFromSource(window.model.nodes["control:demo:control-one"].index())
    window.tree_menu(window.tree.visualRect(index).center())
    settle(app,lambda:not window.deleting and not window.jobs)
    assert triggered==["删除所选（2）"]
    assert [r["id"] for r in window.library.read("demo")["controls"]]==["control-third"]
