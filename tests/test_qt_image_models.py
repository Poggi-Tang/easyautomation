"""The GUI exposes missing models and keeps downloads outside the GUI thread."""
from copy import deepcopy
import time

import pytest

pytest.importorskip("PySide6")
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest

from easy_uiauto import image_models, image_model_paths
from easy_uiauto.qt_ui.window import KnowledgeWindow


def settle(app, predicate):
    deadline = time.monotonic() + 5
    while time.monotonic() < deadline:
        app.processEvents()
        if predicate():
            return
        time.sleep(.01)
    pytest.fail("model status did not settle")


@pytest.fixture
def window(tmp_path, monkeypatch):
    monkeypatch.setenv(image_model_paths.CONFIG_ENV, str(tmp_path / "model-paths.json"))
    app = QApplication.instance() or QApplication([])
    report = dict(ok=False, runtime_issues=[], message="标注模型待补齐（2/3 已就绪）", models=[
        dict(id=key, name=key, path=str(tmp_path / key),
             state="missing" if key == "yolo" else "ready",
             detail="未下载" if key == "yolo" else "已就绪（校验通过）")
        for key in ("yolo", "det", "rec")])
    def inspect():
        configured = image_model_paths.load_model_paths()
        result = deepcopy(report)
        for model in result["models"]:
            model["configured_path"] = configured.get(model["id"], "")
            model["path"] = configured.get(model["id"], model["path"])
        return result
    monkeypatch.setattr(image_models, "inspect_models", inspect)
    gui = KnowledgeWindow(tmp_path / "vault", remember=False)
    settle(app, lambda: not gui.jobs)
    yield app, gui, report
    settle(app, lambda: not gui.jobs)
    gui.dirty = False
    gui.close()
    app.processEvents()


def test_startup_check_and_dialog_recheck_enable_only_missing_downloads(window):
    app, gui, report = window
    assert "待补齐 2/3" in gui.model_button.text()
    gui.model_button.click()
    settle(app, lambda: not gui.model_check_job)
    dialog = gui.model_dialog
    assert len(dialog.path_fields) == 3
    assert dialog.status_labels["yolo"].text() == "未下载"
    assert dialog.status_labels["yolo"].icon_name == "status-warning"
    assert report["models"][0]["path"] in dialog.path_fields["yolo"].placeholderText()
    assert dialog.download_button.isEnabled()
    report["models"][0].update(state="ready", detail="已就绪（校验通过）")
    report.update(ok=True, message="标注模型已就绪（3/3）")
    dialog.refresh_button.click()
    settle(app, lambda: not gui.model_check_job)
    assert "已就绪" in gui.model_button.text()
    assert not dialog.download_button.isEnabled()
    assert dialog.status_labels["yolo"].icon_name == "status-success"


def test_one_click_starts_background_download_with_long_timeout_and_rechecks_on_cancel(window, monkeypatch, tmp_path):
    app, gui, report = window
    calls = []

    class Task:
        operation = "download-image-models"
        folder = tmp_path
        cancelled = False

        def __init__(self, request, *, timeout):
            calls.append((request, timeout))

        def cancel(self):
            self.cancelled = True

        def poll(self):
            return dict(state="cancelled" if self.cancelled else "running",
                        worker_running=not self.cancelled, elapsed_seconds=1,
                        message="正在下载 yolo：1 MB")

        def close(self):
            pass

    monkeypatch.setattr("easy_uiauto.qt_ui.window.LibraryTask", Task)
    gui.open_image_models()
    settle(app, lambda: not gui.model_check_job)
    gui.model_dialog.download_button.click()
    assert calls[0][0]["operation"] == "download-image-models"
    assert calls[0][1] == 3600
    assert gui.model_dialog.download_button.isEnabled()
    assert gui.model_dialog.download_button.text() == "停止下载"
    assert gui.model_dialog.download_button.property("iconName") == "stop"
    assert not gui.model_dialog.path_fields["yolo"].isEnabled()
    gui.poll()
    assert "1 MB" in gui.model_dialog.message.text()
    gui.model_dialog.download_button.click()
    gui.poll()
    settle(app, lambda: not gui.model_check_job)
    assert gui.task is None and gui.model_dialog.download_button.isEnabled()
    assert gui.model_dialog.download_button.text() == "下载 / 修复"
    assert gui.model_dialog.download_button.property("iconName") == "download"
    assert "待补齐" in gui.model_button.text()


def test_select_autosave_reopen_and_clear_model_paths(window, monkeypatch, tmp_path):
    from PySide6.QtWidgets import QFileDialog
    app, gui, report = window
    gui.open_image_models()
    settle(app, lambda: not gui.model_check_job)
    dialog = gui.model_dialog
    custom = tmp_path / "我的模型" / "det.onnx"
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (str(custom), ""))
    dialog.browse_buttons[1].click()
    settle(app, lambda: not gui.model_check_job)
    assert image_model_paths.load_model_paths() == {"det": str(custom)}
    assert not dialog.paths_dirty
    assert str(custom) in dialog.path_fields["det"].toolTip()
    dialog.close()
    gui.open_image_models()
    settle(app, lambda: not gui.model_check_job)
    assert dialog.path_fields["det"].text() == str(custom)
    field = dialog.path_fields["det"]
    field.setFocus()
    field.clear()
    QTest.keyClick(field, Qt.Key.Key_Return)
    settle(app, lambda: not gui.model_check_job)
    assert image_model_paths.load_model_paths() == {}
    assert report["models"][1]["path"] in field.placeholderText()
    assert dialog.isVisible()  # Enter saves; it must not activate a dialog button.


def test_invalid_path_save_preserves_draft_and_does_not_change_configuration(window):
    app, gui, report = window
    gui.open_image_models()
    settle(app, lambda: not gui.model_check_job)
    dialog = gui.model_dialog
    field = dialog.path_fields["yolo"]
    field.setFocus()
    field.setText("relative.pt")
    QTest.keyClick(field, Qt.Key.Key_Return)
    assert dialog.paths_dirty and "路径未保存" in dialog.message.text()
    assert image_model_paths.load_model_paths() == {}
    assert not dialog.download_button.isEnabled()
    dialog.close()
    gui.open_image_models()
    settle(app, lambda: not gui.model_check_job)
    assert field.text() == "relative.pt" and "路径未保存" in dialog.message.text()


@pytest.mark.parametrize("finish", ["tab", "close"])
def test_typed_paths_autosave_on_focus_change_or_close(window, tmp_path, finish):
    app, gui, _ = window
    gui.open_image_models()
    settle(app, lambda: not gui.model_check_job)
    dialog = gui.model_dialog
    field = dialog.path_fields["yolo"]
    field.setFocus()
    custom = str(tmp_path / "custom.pt")
    field.setText(custom)
    if finish == "tab":
        QTest.keyClick(field, Qt.Key.Key_Tab)
    else:
        dialog.close()
    settle(app, lambda: not gui.model_check_job)
    assert image_model_paths.load_model_paths() == {"yolo": custom}
    assert not dialog.paths_dirty


def test_path_change_during_check_ignores_old_report_and_checks_latest_path(window, monkeypatch, tmp_path):
    from threading import Event
    app, gui, report = window
    gui.open_image_models()
    settle(app, lambda: not gui.model_check_job)
    dialog = gui.model_dialog
    started, release = Event(), Event()
    inspect = image_models.inspect_models
    calls = []

    def slow_first_check():
        result = inspect()
        calls.append(result)
        if len(calls) == 1:
            started.set()
            assert release.wait(5)
        return result

    monkeypatch.setattr(image_models, "inspect_models", slow_first_check)
    dialog.refresh_button.click()
    settle(app, started.is_set)
    try:
        field = dialog.path_fields["yolo"]
        assert field.isEnabled() and not dialog.download_button.isEnabled()
        field.setText(str(tmp_path / "latest.pt"))
        QTest.keyClick(field, Qt.Key.Key_Return)
        assert image_model_paths.load_model_paths() == {"yolo": field.text()}
    finally:
        release.set()
    settle(app, lambda: not gui.model_check_job)
    assert len(calls) == 2
    assert gui.model_report["models"][0]["path"] == str(tmp_path / "latest.pt")
    assert dialog.download_button.isEnabled()
