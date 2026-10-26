"""Recorder startup/cleanup and the real isolated highlight process."""

from __future__ import annotations

import subprocess
import time
from unittest.mock import Mock

import pytest

from easy_uiauto import highlight, record


class Listener:
    def __init__(self, **_callbacks):
        self.ident = None
        self.start = Mock(side_effect=self._start)
        self.stop = Mock()
        self.join = Mock()
        self.is_alive = Mock(return_value=False)

    def _start(self):
        self.ident = -1


def recorder(monkeypatch):
    monkeypatch.setattr(record.mouse, "Listener", Listener)
    monkeypatch.setattr(record.keyboard, "Listener", Listener)
    value = record.RecordThread(close_callback=Mock())
    value._ui_highlight = Mock()
    value._ui_highlight.stop.return_value = {"error": None, "forced_termination": False}
    return value


def test_highlight_failure_never_starts_or_joins_listeners(monkeypatch):
    value = recorder(monkeypatch)
    value._ui_highlight.start.side_effect = RuntimeError("missing worker")
    value.run()
    assert str(value.run_error) == "missing worker"
    assert value.started.is_set() and value._stop_complete.is_set()
    assert not value.running
    for listener in (value.mouse_listener, value.keyboard_listener):
        listener.start.assert_not_called()
        listener.stop.assert_not_called()
        listener.join.assert_not_called()
    assert value.cleanup_errors == []
    value._ui_highlight.stop.assert_called_once()


def test_partial_listener_start_only_joins_the_started_listener(monkeypatch):
    value = recorder(monkeypatch)
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
    value._ui_highlight.stop.assert_called_once()
    assert value._stop_complete.is_set()


def test_callback_failure_still_closes_highlight(monkeypatch):
    value = recorder(monkeypatch)
    value.close_callback.side_effect = ValueError("callback failed")
    with pytest.raises(ValueError, match="callback failed"):
        value.stop()
    value._ui_highlight.stop.assert_called_once()
    assert value._stop_complete.is_set()


def test_run_record_propagates_startup_failure_and_write_false_creates_no_file(
    monkeypatch, tmp_path,
):
    monkeypatch.chdir(tmp_path)
    value = recorder(monkeypatch)
    value._ui_highlight.start.side_effect = RuntimeError("missing worker")
    monkeypatch.setattr(record, "RecordThread", lambda **_kwargs: value)
    with pytest.raises(RuntimeError, match="missing worker"):
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


def test_track_control_sends_rect_before_cross_process_xpath(monkeypatch):
    value = recorder(monkeypatch)
    value._ui_highlight.send = Mock()

    class Rect:
        left, top, right, bottom = 10, 20, 110, 80

    class Control:
        Name = "button"
        ControlTypeName = "ButtonControl"
        ClassName = "Button"
        BoundingRectangle = Rect()

    value.track_control(Control(), [{"Name": "button"}])
    assert value._ui_highlight.send.call_args_list[0].args[0] == {
        "cmd": "rect",
        "rect": {
            "left": 10, "top": 20, "right": 110, "bottom": 80,
            "width": 100, "height": 60,
        },
    }
    assert value._ui_highlight.send.call_args_list[1].args[0]["cmd"] == "set_xpath"
