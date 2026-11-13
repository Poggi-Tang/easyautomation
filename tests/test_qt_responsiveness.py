"""Keep incremental navigation and asynchronous persistence safe under slow IO."""
from copy import deepcopy
from pathlib import Path
from threading import Event, get_ident

import pytest

pytest.importorskip("PySide6")
from PySide6.QtCore import QPersistentModelIndex, QTimer
from PySide6.QtWidgets import QGraphicsPixmapItem, QInputDialog

from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.mcp import knowledge
from easy_uiauto.qt_ui.models import KnowledgeModel, ROLE
from easy_uiauto.qt_ui.widgets import ImageView
from easy_uiauto.qt_ui.window import KnowledgeWindow
from test_knowledge_library import library, add_control
from test_qt_knowledge_gui import app, settle
from test_region_picker import picked


@pytest.fixture
def window(app, library, monkeypatch):
    monkeypatch.setattr(KnowledgeWindow, "refresh_image_models", lambda _: None)
    view = KnowledgeWindow(library.root, remember=False)
    settle(app, lambda: not view.jobs)
    yield view
    settle(app, lambda: not view.jobs and not view.image.image_jobs)
    view.dirty = False
    view.close()
    view.deleteLater()
    app.processEvents()


def test_tree_status_stays_at_right_after_initial_load_and_vault_switch(app,window,tmp_path):
    from PySide6.QtWidgets import QHeaderView
    window.resize(1400,850);window.show();window.main_split.setSizes([500,900])
    for switched in (False,True):
        if switched:
            window.library=KnowledgeLibrary(tmp_path/"other-vault")
            window.reload("vault")
            settle(app,lambda:not window.jobs)
        for width in (500,650):
            window.main_split.setSizes([width,1400-width]);app.processEvents()
            header=window.tree.header()
            assert window.model.columnCount()==header.count()==2
            assert header.sectionResizeMode(0)==QHeaderView.ResizeMode.Stretch
            assert header.sectionResizeMode(1)==QHeaderView.ResizeMode.ResizeToContents
            assert header.sectionViewportPosition(1)+header.sectionSize(1)==window.tree.viewport().width()
            assert window.tree.columnWidth(0)>window.tree.columnWidth(1)


def test_refresh_changes_only_one_app_and_keeps_unaffected_indexes(app, library, window, monkeypatch):
    library.bind("other", "Other", dict(title="Other", class_name="Window", handle=2, process_id=3))
    window.reload("vault")
    settle(app, lambda: not window.jobs)
    other = window.model.nodes["app:other"]
    persistent = QPersistentModelIndex(other.index())
    resets, reads, builds = [], [], []
    window.model.modelReset.connect(lambda: resets.append(True))
    read = KnowledgeLibrary.read
    describe = KnowledgeModel.describe
    def tracked_read(store, app_id):
        reads.append((app_id, get_ident()))
        return read(store, app_id)
    def tracked_describe(*args):
        builds.append(get_ident())
        return describe(*args)
    monkeypatch.setattr(KnowledgeLibrary, "read", tracked_read)
    monkeypatch.setattr(KnowledgeModel, "describe", staticmethod(tracked_describe))
    add_control(library)
    window.reload("control:demo:control-one")
    settle(app, lambda: not window.jobs)
    assert [identifier for identifier, _ in reads] == ["demo"]
    assert all(thread != get_ident() for _, thread in reads)
    assert builds and all(thread != get_ident() for thread in builds)
    assert not resets and persistent.isValid()
    assert window.model.nodes["app:other"] is other
    assert window.current["record"]["id"] == "control-one"


def test_incremental_move_preserves_children_when_state_group_disappears(app):
    snapshot = dict(application=dict(name="Demo"), usage=[], references=[], workflows=[], controls=[],
                    pages=[dict(id="first"), dict(id="second", region_state=dict(area_id="first.a", name="Tab"))],
                    regions=[dict(id="first.a", page_id="first", name="Area", source="manual"),
                             dict(id="second.a", page_id="second", name="Area", source="manual")])
    model = KnowledgeModel()
    model.load("Vault", {"demo": snapshot}, {})
    region = model.nodes["region:demo:first.a"]
    stable = QPersistentModelIndex(model.nodes["actions-root:demo"].index())
    assert region.parent() is model.nodes["area:demo:first.a"]
    snapshot = deepcopy(snapshot)
    snapshot["regions"].pop()
    model.load("Vault", {"demo": snapshot}, {})
    assert model.nodes["region:demo:first.a"] is region
    assert region.parent() is model.nodes["controls-root:demo"]
    assert "area:demo:first.a" not in model.nodes
    assert stable.isValid()
    assert region.data(ROLE)["record"]["id"] == "first.a"


def test_region_save_keeps_event_loop_alive_and_retains_picker_file(app, library, window, tmp_path, monkeypatch):
    result = picked(tmp_path)
    entered, release = Event(), Event()
    calls, ticks, cleaned = [], [], []
    original = KnowledgeLibrary.save_picked_region
    def slow_save(store, *args, **kwargs):
        entered.set()
        assert release.wait(5)
        assert Path(result["capture"]["image"]).exists()
        return original(store, *args, **kwargs)
    class PickerTask:
        def close(self):
            cleaned.append(True)
            Path(result["capture"]["image"]).unlink()
    monkeypatch.setattr(KnowledgeLibrary, "save_picked_region", slow_save)
    monkeypatch.setattr(QInputDialog, "getText", lambda *a, **k: ("区域", True))
    monkeypatch.setattr(window, "bound_window", lambda: dict(handle=1))
    monkeypatch.setattr(window, "start_task", lambda op, payload, callback, **kw: calls.append(callback))
    window.goto("app:demo")
    existing = window.model.nodes["app:demo"]
    window.pick_region()
    window.result_task = PickerTask()
    timer = QTimer();timer.setInterval(1);timer.timeout.connect(lambda: ticks.append(True));timer.start()
    try:
        calls[0](result)
        # Simulate poll's cleanup immediately after the picker callback returns.
        window.cleanup_result()
        settle(app, lambda: entered.is_set() and len(ticks) >= 3)
        assert window.saving and not cleaned
        assert not window.centralWidget().isEnabled()
    finally:
        release.set();timer.stop()
    settle(app, lambda: not window.jobs)
    assert cleaned == [True] and not window.saving
    assert window.model.nodes["app:demo"] is existing
    assert window.current["kind"] == "region"
    assert library.asset("demo", window.record["image"]).is_file()


def test_failed_save_keeps_edits_and_does_not_resume_navigation(app, library, window, tmp_path, monkeypatch):
    identifier = library.save_picked_region("demo", picked(tmp_path))
    window.reload("region:demo:" + identifier)
    settle(app, lambda: not window.jobs)
    errors, resumed = [], []
    def fail(*args, **kwargs):raise OSError("disk unavailable")
    monkeypatch.setattr(KnowledgeLibrary, "configure_region", fail)
    monkeypatch.setattr(window, "error", lambda title, error: errors.append(str(error)))
    window.name.setText("Unwritten edit");window.changed()
    assert window.save(done=lambda: resumed.append(True))
    settle(app, lambda: not window.jobs)
    assert errors == ["disk unavailable"] and not resumed
    assert window.dirty and window.name.text() == "Unwritten edit"
    assert window.centralWidget().isEnabled()
    assert library.regions("demo")[0]["name"] != "Unwritten edit"


def test_selection_reuses_image_and_toolbar_but_updates_markers(app, library, window, tmp_path):
    identifier = library.save_picked_region("demo", picked(tmp_path))
    second = library.save_picked_region("demo", picked(tmp_path, rect=dict(left=0, top=.5, right=1, bottom=1)))
    window.reload("region:demo:" + identifier)
    settle(app, lambda: not window.jobs and not window.image.image_jobs)
    buttons = list(window.context_buttons)
    pixmap = next(i for i in window.image.canvas.items() if isinstance(i, QGraphicsPixmapItem))
    window.image.fit_mode = False;window.image.scale(2, 2)
    scale = window.image.transform().m11()
    window.goto("region:demo:" + second)
    assert window.context_buttons == buttons
    assert next(i for i in window.image.canvas.items() if isinstance(i, QGraphicsPixmapItem)) is pixmap
    assert window.image.transform().m11() == scale
    assert window.image.markers[0][0]["top"] == .5


def test_cached_image_invalidates_when_file_changes(app, tmp_path):
    from PIL import Image
    path = tmp_path / "image.png"
    Image.new("RGB", (20, 30), "red").save(path)
    view = ImageView()
    view.show_image(path)
    assert view.image_size == (20, 30)
    Image.new("RGB", (40, 50), "blue").save(path)
    view.show_image(path)
    assert view.image_size == (40, 50)
    view.show_image(None)
    assert view.image_size == (0, 0)
    view.deleteLater()


def test_slow_image_cannot_replace_a_more_recent_selection(app, tmp_path, monkeypatch):
    from PIL import Image
    from easy_uiauto.qt_ui import widgets
    first, second = tmp_path / "first.png", tmp_path / "second.png"
    Image.new("RGB", (20, 30), "red").save(first)
    Image.new("RGB", (40, 50), "blue").save(second)
    entered, release = Event(), Event()
    original = widgets.QImage
    def slow_image(path):
        if Path(path) == first:
            entered.set();assert release.wait(5)
        return original(path)
    monkeypatch.setattr(widgets, "QImage", slow_image)
    view = ImageView(async_load=True)
    try:
        view.show_image(first)
        settle(app, entered.is_set)
        view.show_image(second)
        settle(app, lambda: view.image_size == (40, 50))
    finally:release.set()
    settle(app, lambda: not view.image_jobs)
    assert view.image_size == (40, 50)
    view.deleteLater()


@pytest.mark.parametrize("update", [False, True])
def test_overlap_retry_or_cancel_releases_picker_only_after_completion(app, library, window, tmp_path, monkeypatch, update):
    from easy_uiauto.qt_ui import window as window_module
    identifier = library.save_picked_region("demo", picked(tmp_path))
    result = picked(tmp_path, color="green")
    callbacks, cleaned, dialogs = [], [], []
    real_box = window_module.QMessageBox
    class Box:
        ButtonRole = real_box.ButtonRole
        def __init__(self, *args):self.buttons=[]
        def setWindowTitle(self, _):pass
        def setText(self, _):pass
        def addButton(self, label, role):self.buttons.append(label);return label
        def exec(self):
            assert Path(result["capture"]["image"]).exists() and not cleaned
            dialogs.append(True)
        def clickedButton(self):return self.buttons[0 if update else -1]
    class Picker:
        def close(self):cleaned.append(True);Path(result["capture"]["image"]).unlink()
    monkeypatch.setattr(window_module, "QMessageBox", Box)
    monkeypatch.setattr(QInputDialog, "getText", lambda *a, **kw: ("New", True))
    monkeypatch.setattr(window, "bound_window", lambda: dict(handle=1))
    monkeypatch.setattr(window, "start_task", lambda op, payload, done, **kw: callbacks.append(done))
    window.goto("app:demo");window.pick_region()
    window.result_task = Picker();callbacks[0](result);window.cleanup_result()
    settle(app, lambda: not window.jobs)
    assert dialogs == [True] and cleaned == [True]
    assert [r["id"] for r in library.regions("demo")] == [identifier]
    from PIL import Image
    with Image.open(library.asset("demo", library.page("demo")["layout_image"])) as image:
        assert image.getpixel((0, 0)) == ((0, 128, 0) if update else (0, 0, 255))


def test_workflow_waits_for_save_before_starting_worker(app, library, window, monkeypatch):
    from easy_uiauto import workflows
    control = add_control(library)
    window.reload("actions-root:demo");settle(app, lambda: not window.jobs)
    monkeypatch.setattr(QInputDialog, "getText", lambda *a, **kw: ("Example", True))
    window.new_flow();window.flow_steps.load([workflows.click_step(control)])
    entered, release = Event(), Event()
    original = KnowledgeLibrary.save_workflow
    def slow_save(store, *args, **kwargs):
        entered.set();assert release.wait(5)
        return original(store, *args, **kwargs)
    calls=[]
    monkeypatch.setattr(KnowledgeLibrary, "save_workflow", slow_save)
    monkeypatch.setattr(window, "bound_window", lambda: dict(handle=1))
    monkeypatch.setattr(window, "start_task", lambda *a, **kw: calls.append(a))
    try:
        window.run_flow();settle(app, entered.is_set)
        assert not calls
    finally:release.set()
    settle(app, lambda: not window.jobs)
    assert len(calls) == 1 and calls[0][0] == "workflow"
    assert calls[0][1]["flow_id"] == library.read("demo")["workflows"][0]["id"]


def test_picker_rebuilds_once_and_does_not_rewrite_unaffected_controls(library, tmp_path, monkeypatch):
    add_control(library)
    counts, writes = [], []
    rebuild = knowledge.rebuild_index
    save_control = knowledge.save_control
    def counted(directory):counts.append(True);return rebuild(directory)
    def tracked(directory, record):writes.append(record["id"]);return save_control(directory, record)
    monkeypatch.setattr(knowledge, "rebuild_index", counted)
    monkeypatch.setattr(knowledge, "save_control", tracked)
    identifier = library.save_picked_region("demo", picked(tmp_path))
    assert len(counts) == 1 and writes == ["control-one"]
    counts.clear();writes.clear()
    library.save_picked_region("demo", picked(tmp_path, rect=dict(left=0, top=.5, right=1, bottom=1)))
    assert len(counts) == 1 and not writes
    region = next(r for r in library.regions("demo") if r["id"] == identifier)
    path = library.asset("demo", region["image"])
    before = path.stat().st_mtime_ns
    library.configure_region("demo", identifier, steps=[], name="Renamed", description="Description")
    assert path.stat().st_mtime_ns == before and not writes
    saved = next(r for r in library.regions("demo") if r["id"] == identifier)
    assert saved["name"] == "Renamed" and saved["description"] == "Description"
