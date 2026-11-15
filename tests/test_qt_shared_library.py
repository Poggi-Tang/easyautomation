"""External library commits update navigation without losing in-progress edits."""
import json
from threading import Event, get_ident

import pytest
pytest.importorskip("PySide6")
from PySide6.QtCore import QPersistentModelIndex, QTimer

from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.mcp import server
from easy_uiauto.qt_ui.window import KnowledgeWindow
from test_knowledge_library import library, add_control
from test_qt_knowledge_gui import app, settle
from test_shared_control_library import legacy


@pytest.fixture
def window(app, library, monkeypatch):
    monkeypatch.setattr(KnowledgeWindow, "refresh_image_models", lambda _:None)
    view=KnowledgeWindow(library.root, remember=False)
    view.external_timer.stop()
    settle(app, lambda:not view.jobs)
    view.refresh_external_changes();settle(app, lambda:not view.jobs)
    yield view
    settle(app, lambda:not view.jobs and not view.image.image_jobs)
    view.dirty=False;view.close();view.deleteLater();app.processEvents()


def refresh(app, window):
    window.refresh_external_changes();settle(app, lambda:not window.jobs)


def test_external_import_updates_tree_without_reset_or_reading_unchanged_app(app, library, window, monkeypatch):
    library.bind("other", "Other", dict(title="Other", class_name="Window"))
    refresh(app, window)
    item=window.model.nodes["app:other"];persistent=QPersistentModelIndex(item.index())
    resets=[];reads=[];window.model.modelReset.connect(lambda:resets.append(True))
    original=KnowledgeLibrary.read
    def read(store, app_id):
        reads.append((app_id,get_ident()));return original(store,app_id)
    result=json.loads(server.upsert_control_record(json.dumps(legacy()), app_id="demo", vault=str(library.root)))
    monkeypatch.setattr(KnowledgeLibrary,"read",read)
    refresh(app,window)
    assert "control:demo:"+result["id"] in window.model.nodes
    assert not resets and persistent.isValid() and window.model.nodes["app:other"] is item
    assert reads and all(name=="demo" and thread!=get_ident() for name,thread in reads)
    reads.clear();refresh(app,window)
    assert reads==[]
    window.goto("control:demo:"+result["id"])
    assert window.record["id"]==result["id"] and window.name.text()=="Open"


def test_external_edit_keeps_local_input_and_conflict_prevents_overwrite(app, library, window, monkeypatch):
    control=add_control(library);refresh(app,window)
    window.goto("control:demo:"+control["id"])
    window.name.setText("Local edit");window.changed()
    library.update_control("demo",control["id"],name="External edit",description="External description",aliases=[])
    refresh(app,window)
    assert window.dirty and window.name.text()=="Local edit"
    assert window.snapshots["demo"]["controls"][0]["semantic_name"]=="External edit"
    errors=[];monkeypatch.setattr(window,"error",lambda title,error:errors.append(str(error)))
    window.save();settle(app,lambda:not window.jobs)
    assert errors and "其他入口" in errors[0]
    assert window.dirty and window.name.text()=="Local edit"
    assert library.read("demo")["controls"][0]["semantic_name"]=="External edit"


def test_external_delete_refreshes_rows_and_protects_dirty_deleted_record(app, library, window, monkeypatch):
    control=add_control(library);refresh(app,window)
    key="control:demo:"+control["id"];window.goto(key)
    window.name.setText("Unsaved");window.changed()
    library.archive("demo","control",control["id"]);refresh(app,window)
    assert key not in window.model.nodes and window.dirty and window.name.text()=="Unsaved"
    errors=[];monkeypatch.setattr(window,"error",lambda title,error:errors.append(str(error)))
    window.save();settle(app,lambda:not window.jobs)
    assert errors and library.read("demo")["controls"]==[] and window.dirty


def test_slow_external_read_does_not_block_event_loop_or_replace_switched_vault(app, library, window, monkeypatch, tmp_path):
    add_control(library)
    entered=Event();release=Event();original=KnowledgeLibrary.read;ticks=[]
    def read(store,app_id):
        if store.root==library.root:entered.set();assert release.wait(5)
        return original(store,app_id)
    monkeypatch.setattr(KnowledgeLibrary,"read",read)
    timer=QTimer();timer.setInterval(1);timer.timeout.connect(lambda:ticks.append(True));timer.start()
    try:
        window.refresh_external_changes()
        settle(app,lambda:entered.is_set() and len(ticks)>3)
        window.library=KnowledgeLibrary(tmp_path/"new-vault");window.reload("vault")
    finally:release.set();timer.stop()
    settle(app,lambda:not window.jobs)
    assert not window.snapshots and "app:demo" not in window.model.nodes
