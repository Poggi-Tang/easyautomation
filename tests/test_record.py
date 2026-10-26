"""Recorder startup, queued control tracking, cleanup, and highlight worker."""

from __future__ import annotations

import queue
import subprocess
import threading
import time
from unittest.mock import Mock

import pytest

from easy_uiauto import highlight, record
from easy_uiauto.draw import ScreenLineBox


class Listener:
    def __init__(self, **_callbacks):
        self.ident = None
        self.start = Mock(side_effect=self._start)
        self.stop = Mock()
        self.join = Mock()
        self.is_alive = Mock(return_value=False)

    def _start(self):
        self.ident = -1


class FakeTk:
    """Small event-loop fake so the queue handoff can be tested without a GUI."""

    def __init__(self):
        self.callbacks = queue.Queue()
        self.stopped = threading.Event()
        self.quitting = threading.Event()
        self.owner_thread = None

    def withdraw(self):
        pass

    def after(self, _delay_ms, callback):
        self.callbacks.put(callback)

    def mainloop(self):
        self.owner_thread = threading.get_ident()
        while not self.stopped.is_set() and not self.quitting.is_set():
            try:
                callback = self.callbacks.get(timeout=0.05)
            except queue.Empty:
                continue
            callback()

    def destroy(self):
        self.stopped.set()

    def quit(self):
        self.quitting.set()


def recorder(monkeypatch):
    monkeypatch.setattr(record.mouse, "Listener", Listener)
    monkeypatch.setattr(record.keyboard, "Listener", Listener)
    return record.RecordThread(close_callback=Mock())


def test_ui_start_failure_never_starts_or_joins_listeners(monkeypatch):
    value = recorder(monkeypatch)
    value.start_ui_thread = Mock(side_effect=RuntimeError("Tk failed"))
    value.run()
    assert str(value.run_error) == "Tk failed"
    assert value.started.is_set() and value._stop_complete.is_set()
    assert not value.running
    for listener in (value.mouse_listener, value.keyboard_listener):
        listener.start.assert_not_called()
        listener.stop.assert_not_called()
        listener.join.assert_not_called()
    assert value.cleanup_errors == []
    value.start_ui_thread.assert_called_once()


def test_partial_listener_start_only_joins_the_started_listener(monkeypatch):
    value = recorder(monkeypatch)
    value.start_ui_thread = Mock()
    value.keyboard_listener.start.side_effect = OSError("keyboard hook failed")
    value.run()
    assert str(value.run_error) == "keyboard hook failed"
    value.mouse_listener.stop.assert_called_once()
    value.mouse_listener.join.assert_called_once()
    value.keyboard_listener.join.assert_not_called()
    assert value.cleanup_errors == []


def test_stop_is_idempotent_even_before_start(monkeypatch):
    value = recorder(monkeypatch)
    value.stop()
    value.stop()
    value.close_callback.assert_called_once()
    assert value._stop_complete.is_set()
    assert value.ui_cmd_queue.get_nowait() == {"cmd": "close"}


def test_callback_failure_still_closes_ui_thread(monkeypatch):
    value = recorder(monkeypatch)
    value.close_callback.side_effect = ValueError("callback failed")
    with pytest.raises(ValueError, match="callback failed"):
        value.stop()
    assert value.ui_cmd_queue.get_nowait() == {"cmd": "close"}
    assert value._stop_complete.is_set()


def test_run_record_propagates_startup_failure_and_write_false_creates_no_file(
    monkeypatch, tmp_path,
):
    monkeypatch.chdir(tmp_path)
    value = recorder(monkeypatch)
    value.start_ui_thread = Mock(side_effect=RuntimeError("Tk failed"))
    monkeypatch.setattr(record, "RecordThread", lambda **_kwargs: value)
    with pytest.raises(RuntimeError, match="Tk failed"):
        record.run_record(write_file=False)
    assert list(tmp_path.iterdir()) == []


def test_missing_worker_is_reported_before_start(monkeypatch, tmp_path):
    monkeypatch.setattr(highlight, "__file__", str(tmp_path / "highlight.py"))
    popen = Mock()
    monkeypatch.setattr(highlight.subprocess, "Popen", popen)
    with pytest.raises(RuntimeError, match="missing from"):
        highlight.HighlightProcess().start()
    popen.assert_not_called()


def substitute_child(monkeypatch, code):
    popen = subprocess.Popen

    def launch(args, **kwargs):
        return popen([args[0], "-E", "-P", "-c", code], **kwargs)

    monkeypatch.setattr(highlight.subprocess, "Popen", launch)


def test_child_exit_is_a_fast_error_instead_of_a_startup_timeout(monkeypatch):
    substitute_child(monkeypatch,
                     "import sys; print('No module named easy_uiauto._highlight_worker', "
                     "file=sys.stderr); sys.exit(2)")
    worker = highlight.HighlightProcess()
    started = time.monotonic()
    with pytest.raises(RuntimeError, match="exited before ready.*No module named"):
        worker.start(timeout=8)
    assert time.monotonic() - started < 4
    assert worker.status()["exitcode"] == 2


def test_unresponsive_child_is_terminated_at_startup_deadline(monkeypatch):
    substitute_child(monkeypatch, "import time; time.sleep(60)")
    worker = highlight.HighlightProcess()
    with pytest.raises(TimeoutError, match="start timed out"):
        worker.start(timeout=0.1)
    assert worker.status()["stopped"]
    assert worker.status()["forced_termination"]


def test_worker_keeps_interpreter_site_packages_available(monkeypatch):
    substitute_child(
        monkeypatch,
        "import json, sys; sys.stdin.readline(); "
        "print(json.dumps({'event': 'error' if sys.flags.no_user_site or "
        "sys.flags.isolated else 'ready', 'message': 'site packages disabled'}), "
        "flush=True); sys.stdin.readline()",
    )
    worker = highlight.HighlightProcess()
    try:
        worker.start(timeout=3)
        assert worker.status()["running"]
    finally:
        status = worker.stop()
    assert status["exitcode"] == 0
    assert status["error"] is None


def test_real_worker_starts_from_an_unrelated_cwd_and_stops_cleanly(monkeypatch, tmp_path):
    # The child ignores PYTHONPATH and the cwd, but retains installed dependencies.
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("PYTHONPATH", str(tmp_path / "unrelated-package"))
    worker = highlight.HighlightProcess()
    try:
        worker.start(timeout=8)
        assert worker.status()["running"]
        worker.send({"cmd": "rect", "rect": {"left": 30, "top": 30,
                    "right": 100, "bottom": 80, "width": 70, "height": 50}})
    finally:
        status = worker.stop()
    assert status["exitcode"] == 0
    assert status["error"] is None
    assert not status["forced_termination"]


def test_track_control_queues_the_control_for_the_tk_thread(monkeypatch):
    value = recorder(monkeypatch)

    class Control:
        Name = "button"
        ControlTypeName = "ButtonControl"
        ClassName = "Button"

    control = Control()
    value.track_control(control)
    assert value.ui_cmd_queue.get_nowait() == {"cmd": "set_control", "control": control}


def test_click_updates_the_highlight_control_without_callback_signature_error(monkeypatch):
    value = recorder(monkeypatch)

    class Control:
        Name = "button"
        ControlTypeName = "ButtonControl"
        ClassName = "Button"

    control = Control()
    xpath = [{"Name": "window"}, {"Name": "button"}]
    monkeypatch.setattr(record, "get_control_info", Mock(return_value=(xpath, control)))
    monkeypatch.setattr(record, "CoInitialize", Mock())
    monkeypatch.setattr(record, "CoUninitialize", Mock())
    value.on_click(10, 20, record.Button.left, True)
    assert value.ui_cmd_queue.get_nowait() == {"cmd": "set_control", "control": control}
    assert value.actions_data[-1]["ACTION"] == "点击"


def test_drag_updates_source_then_destination_highlight(monkeypatch):
    value = recorder(monkeypatch)

    class Control:
        def __init__(self, name):
            self.Name = name
            self.ControlTypeName = "ButtonControl"
            self.ClassName = "Button"

    source = Control("source")
    destination = Control("destination")
    source_xpath = [{"Name": "window"}, {"Name": "source"}]
    destination_xpath = [{"Name": "window"}, {"Name": "destination"}]
    lookup = Mock(side_effect=[(source_xpath, source), (destination_xpath, destination)])
    monkeypatch.setattr(record, "get_control_info", lookup)
    monkeypatch.setattr(record, "CoInitialize", Mock())
    monkeypatch.setattr(record, "CoUninitialize", Mock())
    value.last_click_time = time.time() - 1
    value.last_click_position = (10, 20)

    value.on_click(30, 40, record.Button.left, False)

    assert value.ui_cmd_queue.get_nowait() == {"cmd": "set_control", "control": source}
    assert value.ui_cmd_queue.get_nowait() == {"cmd": "set_control", "control": destination}
    assert value.actions_data[-1]["ACTION"] == "拖拽"


def test_ui_queue_is_applied_on_tk_thread_and_stop_closes_it(monkeypatch):
    value = recorder(monkeypatch)
    root = FakeTk()
    boxes = []

    class FakeBox:
        def __init__(self, **kwargs):
            self.root = kwargs["root"]
            self.controls = []
            self.destroyed = False
            boxes.append(self)

        def set_control(self, control):
            self.controls.append((control, threading.get_ident()))

        def set_style(self, **_kwargs):
            pass

        def destroy(self):
            self.destroyed = True
            self.root.quit()

    monkeypatch.setattr(record.tk, "Tk", lambda: root)
    monkeypatch.setattr(record, "ScreenLineBox", FakeBox)
    value.start_ui_thread()
    control = object()
    value.track_control(control)
    deadline = time.monotonic() + 2
    while (not boxes or not boxes[0].controls) and time.monotonic() < deadline:
        time.sleep(0.01)
    assert boxes[0].controls == [(control, root.owner_thread)]

    value.stop()
    assert value.ui_thread is not None and not value.ui_thread.is_alive()
    assert boxes[0].destroyed


def test_escape_stops_recording_and_closes_highlight(monkeypatch):
    value = recorder(monkeypatch)
    root = FakeTk()
    boxes = []

    class FakeBox:
        def __init__(self, **_kwargs):
            self.destroyed = False
            boxes.append(self)

        def set_control(self, _control):
            pass

        def set_style(self, **_kwargs):
            pass

        def destroy(self):
            self.destroyed = True
            root.quit()

    monkeypatch.setattr(record.tk, "Tk", lambda: root)
    monkeypatch.setattr(record, "ScreenLineBox", FakeBox)
    value.start_ui_thread()
    value.on_press(record.keyboard.Key.esc)
    assert not value.running
    assert value._stop_complete.is_set()
    assert value.ui_thread is not None and not value.ui_thread.is_alive()
    assert boxes[0].destroyed


def test_screen_line_box_cancels_pending_tracking_tick_on_destroy():
    box = ScreenLineBox.__new__(ScreenLineBox)
    box._destroyed = False
    box._tracking_active = True
    box._tick_after_id = "pending-tick"
    box._owns_root = False
    box.root = Mock()
    box.win_top = Mock()
    box.win_right = Mock()
    box.win_bottom = Mock()
    box.win_left = Mock()

    box.destroy()
    box.destroy()

    box.root.after_cancel.assert_called_once_with("pending-tick")
    box.root.quit.assert_called_once()
    for window in box._wins():
        window.destroy.assert_called_once()
    assert box._tick_after_id is None
