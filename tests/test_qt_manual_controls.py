"""Exercise manual addition without starting a desktop picker or native input."""
from pathlib import Path
from threading import Event

import pytest

pytest.importorskip("PySide6")
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QDialog

from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.qt_ui.dialogs import CapturedControlDialog
from test_knowledge_library import library
from test_region_picker import picked
from test_qt_knowledge_gui import app, settle
from test_qt_responsiveness import window


def test_dialog_edits_existing_details_and_requires_name(app, library, tmp_path):
    region = library.save_picked_region("demo", picked(tmp_path))
    result = picked(tmp_path)
    library.save_captured_control("demo", result, name="已有名称", description="已有说明", region_id=region)
    dialog = CapturedControlDialog(result["selection"], library.read("demo"), region_id=region)
    assert dialog.name.text() == "已有名称" and dialog.description.toPlainText() == "已有说明"
    assert dialog.values()["region_id"] == region and "更新原记录" in dialog.message.text()
    dialog.name.clear();dialog.accept()
    assert dialog.result() != QDialog.DialogCode.Accepted
    dialog.name.setText("新名称");dialog.description.setPlainText("新说明");dialog.accept()
    assert dialog.values()["name"] == "新名称" and dialog.values()["description"] == "新说明"
    assert dialog.result() == QDialog.DialogCode.Accepted
    dialog.deleteLater()


@pytest.mark.parametrize("cancel", [True, False])
def test_manual_add_cancels_or_saves_asynchronously_without_tree_reset(app, library, window, tmp_path, monkeypatch, cancel):
    region = library.save_picked_region("demo", picked(tmp_path))
    window.reload("region:demo:"+region);settle(app, lambda: not window.jobs)
    add = next(b for b in window.context_buttons if b.text() == "添加控件")
    assert not add.icon().isNull()
    result = picked(tmp_path);calls=[];cleaned=[];resets=[];ticks=[]
    entered, release = Event(), Event()
    original = KnowledgeLibrary.save_captured_control
    def save(store, *args, **kwargs):
        entered.set();assert release.wait(5)
        assert Path(result["capture"]["image"]).is_file()
        return original(store, *args, **kwargs)
    def fill(dialog):
        assert dialog.values()["region_id"] == region
        dialog.name.setText("手动按钮");dialog.description.setPlainText("手动说明")
        return QDialog.DialogCode.Rejected if cancel else QDialog.DialogCode.Accepted
    class PickerTask:
        def close(self):
            cleaned.append(True);Path(result["capture"]["image"]).unlink()
    monkeypatch.setattr(KnowledgeLibrary, "save_captured_control", save)
    monkeypatch.setattr(CapturedControlDialog, "exec", fill)
    monkeypatch.setattr(window, "bound_window", lambda: dict(handle=1))
    monkeypatch.setattr(window, "start_task", lambda op, payload, callback, **kw: calls.append((op, payload, callback, kw)))
    window.model.modelReset.connect(lambda: resets.append(True))
    existing = window.model.nodes["region:demo:"+region]
    add.click()
    op, payload, callback, options = calls[0]
    assert op == "pick-region" and payload["match_mode"] == "identity" and options["hide"]
    window.result_task = PickerTask()
    timer=QTimer();timer.setInterval(1);timer.timeout.connect(lambda: ticks.append(True));timer.start()
    try:
        callback(result);window.cleanup_result()
        if not cancel:
            settle(app, lambda: entered.is_set() and len(ticks) >= 3)
            assert window.saving and not cleaned
    finally:release.set();timer.stop()
    settle(app, lambda: not window.jobs)
    assert cleaned == [True] and not resets
    assert window.model.nodes["region:demo:"+region] is existing
    if cancel:assert not library.read("demo")["controls"]
    else:
        assert window.current["kind"] == "control"
        assert window.record["semantic_name"] == "手动按钮" and window.record["description"] == "手动说明"
        assert window.model.nodes[window.current["key"]].parent() is existing
