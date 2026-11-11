"""Optional Qt UI keeps edits, state grouping and navigation free of UIA input."""
import time
from copy import deepcopy

import pytest

pytest.importorskip("PySide6")
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

from easy_uiauto.qt_ui.window import KnowledgeWindow
from easy_uiauto.qt_ui.widgets import StepEditor
from easy_uiauto import workflows
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library,add_control
from test_region_picker import picked


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


@pytest.mark.parametrize("record,expected",[
    (dict(status="verified",control_type="ButtonControl"),"定位已验证"),
    (dict(status="observed",control_type="VisualElement"),"图片定位待验证"),
    (dict(status="suspect",control_type="ButtonControl"),"待复核"),
    (dict(status="quarantined",semantic_status="uncertain",
          verification=dict(location="passed",image="passed")),"功能待确认"),
    (dict(status="quarantined",semantic_status="uncertain",
          verification=dict(location="failed",image="passed")),"图片定位待验证"),
    (dict(status="quarantined",semantic_status="uncertain",quarantine=dict(active=True,reason="failed"),
          verification=dict(location="passed",image="passed")),"已隔离"),
])
def test_tree_status_distinguishes_native_visual_and_unconfirmed_records(record,expected):
    from easy_uiauto.qt_ui.models import control_status
    assert control_status(record)==expected


def settle(app,predicate,seconds=5):
    deadline=time.monotonic()+seconds
    while time.monotonic()<deadline:
        app.processEvents()
        if predicate():return
        time.sleep(.01)
    raise AssertionError("Qt state did not settle")


def test_step_edits_update_the_saved_definition_instead_of_a_variant_copy(app):
    record=dict(id="button",name="Button")
    editor=StepEditor(lambda:record);editor.load([workflows.click_step(record)])
    editor.value.setText("输入值");editor.kind.setCurrentIndex(editor.kind.findData("set-text"));editor.edit()
    assert editor.steps[0]["text"]=="输入值" and editor.steps[0]["action"]=="set-text"
    editor.add_check();assert len(editor.steps[0]["checks"])==1
    editor.kind.setCurrentIndex(editor.kind.findData("node-exists"));editor.value.setText("${工程名称}");editor.edit()
    assert editor.steps[0]["checks"][0]["expected"]=="${工程名称}"
    editor.deleteLater();app.processEvents()


def test_parameter_selector_sets_the_same_input_and_assertion_value(app):
    editor=StepEditor(lambda:dict(id="edit",name="名称"))
    editor.set_parameters({"工程名称":"demo"});editor.add_action()
    editor.kind.setCurrentIndex(editor.kind.findData("set-text"))
    editor.value_source.setCurrentIndex(editor.value_source.findData("工程名称"))
    assert editor.steps[0]["text"]=="${工程名称}"
    editor.add_check();editor.kind.setCurrentIndex(editor.kind.findData("text"))
    editor.value_source.setCurrentIndex(editor.value_source.findData("工程名称"))
    assert editor.steps[0]["checks"][0]["expected"]=="${工程名称}"
    editor.deleteLater();app.processEvents()


def test_image_drag_maps_back_to_original_pixels_after_scaling(app,tmp_path):
    from PIL import Image
    from PySide6.QtCore import QPointF
    from PySide6.QtTest import QTest
    from easy_uiauto.qt_ui.boundary_dialog import RegionImage
    path=tmp_path/"frame.png";Image.new("RGB",(100,100),"white").save(path)
    view=RegionImage();view.resize(400,400);view.show_image(path);view.show();app.processEvents()
    start=view.mapFromScene(QPointF(20,30));end=view.mapFromScene(QPointF(80,90))
    QTest.mousePress(view.viewport(),Qt.MouseButton.LeftButton,pos=start)
    QTest.mouseMove(view.viewport(),end)
    QTest.mouseRelease(view.viewport(),Qt.MouseButton.LeftButton,pos=end)
    assert view.rectangle==pytest.approx(dict(left=.2,top=.3,right=.8,bottom=.9),abs=.005)
    view.close();view.deleteLater();app.processEvents()


def test_new_window_loads_existing_vault_and_selection_does_not_run_actions(app,library,tmp_path,monkeypatch):
    identifier=library.save_picked_region("demo",picked(tmp_path));add_control(library)
    calls=[]
    monkeypatch.setattr("easy_uiauto.qt_ui.window.LibraryTask",lambda *a,**kw:calls.append(a))
    window=KnowledgeWindow(library.root,remember=False)
    try:
        settle(app,lambda:"demo" in window.snapshots)
        window.goto("region:demo:"+identifier);app.processEvents()
        assert window.entry_summary.text()=="无" and not calls
        assert window.name.text()=="工具栏"
        window.resize(1100,700);window.show();app.processEvents()
        assert window.statusBar().isVisible() and not window.details.isVisible()
    finally:window.dirty=False;window.close();app.processEvents()


def test_qt_tree_groups_same_location_states_and_keeps_control_id(app,library,tmp_path):
    identifier=library.save_picked_region("demo",picked(tmp_path));entry=add_control(library)
    variant=library.save_variant("demo",library.regions("demo")[0],picked(tmp_path,color="green"),name="绘图",steps=[workflows.click_step(entry)],
                                 condition=dict(location=entry["location"],kind="exists",label="绘图"))
    window=KnowledgeWindow(library.root,remember=False)
    try:
        settle(app,lambda:"area:demo:"+identifier in window.model.nodes)
        parent=window.model.nodes["area:demo:"+identifier]
        assert parent.rowCount()==2
        assert "region:demo:"+variant in window.model.nodes
        assert "control:demo:"+entry["id"] in window.model.nodes
    finally:window.dirty=False;window.close();app.processEvents()


def test_save_during_navigation_does_not_jump_back_after_async_reload(app,library,tmp_path,monkeypatch):
    from PySide6.QtWidgets import QMessageBox
    identifier=library.save_picked_region("demo",picked(tmp_path));record=add_control(library)
    window=KnowledgeWindow(library.root,remember=False)
    try:
        settle(app,lambda:"demo" in window.snapshots)
        window.goto("region:demo:"+identifier);window.name.setText("Edited");window.changed()
        monkeypatch.setattr(QMessageBox,"question",lambda *_a,**_k:QMessageBox.StandardButton.Save)
        key="control:demo:"+record["id"];window.goto(key)
        settle(app,lambda:not window.jobs)
        assert window.current["key"]==key and library.regions("demo")[0]["name"]=="Edited"
    finally:window.dirty=False;window.close();app.processEvents()


def test_embedded_settings_cancellation_blocks_configuration_write(app,tmp_path,monkeypatch):
    from easy_uiauto.qt_ui.dialogs import SettingsDialog
    from easy_uiauto.mcp import configuration
    saved=[]
    monkeypatch.setattr(configuration,"load_gui_configuration",lambda:dict(knowledge_dir=str(tmp_path),api_url="",api_key="",api_model=""))
    monkeypatch.setattr(configuration,"save_knowledge_directory",lambda value:saved.append(value))
    def cancel(_values):raise ValueError("已取消切换知识库")
    dialog=SettingsDialog(tmp_path,before_save=cancel)
    dialog.save()
    assert not saved and "取消" in dialog.message.text()
    dialog.reject();dialog.deleteLater();app.processEvents()


def test_initial_vault_selection_renders_its_actions_even_after_model_auto_selects_root(app,tmp_path):
    window=KnowledgeWindow(tmp_path,remember=False)
    try:
        settle(app,lambda:not window.jobs)
        assert window.current["key"]=="vault"
        assert [button.text() for button in window.context_buttons]==["添加软件"]
        assert window.caption.text()=="知识库"
    finally:window.dirty=False;window.close();app.processEvents()


def test_layout_review_shows_saved_boundaries_and_cancel_does_not_change_knowledge(app,library,tmp_path):
    from easy_uiauto.qt_ui.layout_dialog import LayoutDialog
    first=library.save_picked_region("demo",picked(tmp_path))
    library.save_picked_region("demo",picked(tmp_path,rect=dict(left=0,top=.5,right=1,bottom=1)))
    library.save_capture("demo",picked(tmp_path,size=(800,600),color="green")["capture"])
    before=deepcopy(library.read("demo"));path=library.asset("demo",library.page("demo")["layout_image"])
    image=path.read_bytes()
    dialog=LayoutDialog(library,"demo",before,page_id=library.page_id,region_id=first,for_learning=True)
    try:
        assert dialog.tree.topLevelItemCount()==2 and dialog.image.image_size==(800,600)
        assert "待确认" in dialog.message.text() and "800 × 600" in dialog.message.text()
        assert not dialog.pages.isEnabled() and dialog.confirm.text()=="确认区域并学习"
        assert dialog.tree.currentItem().data(0,Qt.ItemDataRole.UserRole)==first
        dialog.reject()
        assert library.read("demo")==before and path.read_bytes()==image
    finally:dialog.deleteLater();app.processEvents()


@pytest.mark.parametrize("confirm",[False,True])
def test_learning_routes_to_review_and_requires_explicit_confirmation(app,library,tmp_path,monkeypatch,confirm):
    from easy_uiauto.qt_ui.layout_dialog import LayoutDialog
    identifier=library.save_picked_region("demo",picked(tmp_path))
    record=add_control(library)
    library.save_capture("demo",picked(tmp_path,size=(800,600),color="green")["capture"])
    before=deepcopy(library.read("demo"));calls=[];dialogs=[]
    def review(dialog):
        dialogs.append(dialog)
        if confirm:dialog.accept()
        else:dialog.reject()
        return dialog.result()
    monkeypatch.setattr(LayoutDialog,"exec",review)
    window=KnowledgeWindow(library.root,remember=False)
    try:
        settle(app,lambda:"demo" in window.snapshots)
        window.goto("region:demo:"+identifier)
        assert "查看区域" in [b.text() for b in window.context_buttons]
        item=window.model.nodes["region:demo:"+identifier]
        assert item.parent().child(item.row(),1).text()=="待确认"
        window.windows["demo"]=dict(title="Test",handle=123,process_id=456,class_name="Window")
        monkeypatch.setattr(window,"start_task",lambda *a,**kw:calls.append((a,kw)))
        window.learn()
        assert len(dialogs)==1 and library.page("demo")["layout_confirmed"]==confirm
        assert len(calls)==int(confirm)
        if confirm:
            assert calls[0][0][0]=="learn"
            assert calls[0][0][1]["layout"]["image_size"]==dict(width=800,height=600)
            assert calls[0][0][1]["region_id"]==identifier
        else:assert library.read("demo")==before
        assert next(r for r in library.read("demo")["controls"] if r["id"]==record["id"])==before["controls"][0]
    finally:window.dirty=False;window.close();app.processEvents()


def test_review_modification_returns_to_the_selected_region_without_learning(app,library,tmp_path,monkeypatch):
    from easy_uiauto.qt_ui.layout_dialog import LayoutDialog
    first=library.save_picked_region("demo",picked(tmp_path))
    second=library.save_picked_region("demo",picked(tmp_path,rect=dict(left=0,top=.5,right=1,bottom=1)))
    library.save_capture("demo",picked(tmp_path,size=(800,600),color="green")["capture"])
    def review(dialog):
        selected=next(dialog.tree.topLevelItem(i) for i in range(dialog.tree.topLevelItemCount())
                      if dialog.tree.topLevelItem(i).data(0,Qt.ItemDataRole.UserRole)==second)
        dialog.tree.setCurrentItem(selected);dialog.modify.click();return dialog.result()
    monkeypatch.setattr(LayoutDialog,"exec",review)
    window=KnowledgeWindow(library.root,remember=False);edits=[];tasks=[]
    try:
        settle(app,lambda:"demo" in window.snapshots);window.goto("region:demo:"+first)
        window.windows["demo"]=dict(title="Test",handle=123,process_id=456,class_name="Window")
        monkeypatch.setattr(window,"pick_region",lambda **kw:edits.append((window.record["id"],kw)))
        monkeypatch.setattr(window,"start_task",lambda *a,**kw:tasks.append(a))
        window.learn()
        assert edits==[(second,dict(modify=True))] and not tasks
        assert not library.page("demo")["layout_confirmed"]
    finally:window.dirty=False;window.close();app.processEvents()


def test_review_confirms_only_the_chosen_state(app,library,tmp_path):
    from easy_uiauto.qt_ui.layout_dialog import LayoutDialog
    base=library.save_picked_region("demo",picked(tmp_path));entry=add_control(library)
    variant=library.save_variant("demo",library.regions("demo")[0],picked(tmp_path,color="green"),name="绘图",
        steps=[workflows.click_step(entry)],condition=dict(location=entry["location"],kind="exists",label="绘图"))
    state_page=next(r["page_id"] for r in library.read("demo")["regions"] if r["id"]==variant)
    for page_id in (library.page_id,state_page):
        scope=library.scoped(page_id);page=scope.page("demo");page["layout_confirmed"]=False
        knowledge.save_page(library.directory("demo"),page)
    before=deepcopy(library.read("demo"))
    dialog=LayoutDialog(library,"demo",before,page_id=library.page_id,region_id=base)
    try:
        dialog.pages.setCurrentIndex(dialog.pages.findData(state_page));dialog.accept()
        assert dialog.confirmed_page_id==state_page
        assert library.scoped(state_page).page("demo")["layout_confirmed"]
        assert not library.page("demo")["layout_confirmed"]
        after=library.read("demo")
        assert after["regions"]==before["regions"] and after["controls"]==before["controls"]
    finally:dialog.deleteLater();app.processEvents()


def test_saved_layout_is_accessible_from_software_without_binding_or_capture(app,library,tmp_path,monkeypatch):
    from easy_uiauto.qt_ui.layout_dialog import LayoutDialog
    library.save_picked_region("demo",picked(tmp_path));before=deepcopy(library.read("demo"))
    dialogs=[]
    def review(dialog):dialogs.append(dialog);dialog.reject();return dialog.result()
    monkeypatch.setattr(LayoutDialog,"exec",review)
    window=KnowledgeWindow(library.root,remember=False);tasks=[]
    try:
        settle(app,lambda:"demo" in window.snapshots);window.goto("app:demo")
        monkeypatch.setattr(window,"start_task",lambda *a,**kw:tasks.append(a))
        next(b for b in window.context_buttons if b.text()=="查看区域").click()
        assert len(dialogs)==1 and not tasks and not window.windows
        assert library.read("demo")==before
    finally:window.dirty=False;window.close();app.processEvents()


def test_missing_layout_picture_cannot_be_confirmed(app,library,tmp_path):
    from easy_uiauto.qt_ui.layout_dialog import LayoutDialog
    library.save_picked_region("demo",picked(tmp_path))
    path=library.asset("demo",library.page("demo")["layout_image"]);path.unlink()
    dialog=LayoutDialog(library,"demo",library.read("demo"))
    try:
        assert not dialog.confirm.isEnabled() and "布局图缺失" in dialog.message.text()
        dialog.accept();assert not dialog.confirmed_page_id
    finally:dialog.reject();dialog.deleteLater();app.processEvents()
