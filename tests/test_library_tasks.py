import json
from pathlib import Path
import time
from types import SimpleNamespace

import pytest

from easy_uiauto import library_tasks


@pytest.mark.parametrize("winerror", [5, 32, 33])
def test_status_publication_waits_for_a_brief_windows_reader(monkeypatch, tmp_path, winerror):
    path = tmp_path / "state.json"
    path.write_text('{"state":"running"}', encoding="utf-8")
    replace = Path.replace
    attempts, waits = [], []
    def busy_reader(source, destination):
        attempts.append(source.read_text(encoding="utf-8"))
        if len(attempts) < 3:
            assert json.loads(path.read_text(encoding="utf-8")) == {"state":"running"}
            error = PermissionError("Destination is held by a reader"); error.winerror = winerror
            raise error
        return replace(source, destination)
    monkeypatch.setattr(Path, "replace", busy_reader)
    monkeypatch.setattr(library_tasks.time, "sleep", waits.append)
    library_tasks._write(path, {"state":"completed"})
    assert json.loads(path.read_text(encoding="utf-8")) == {"state":"completed"}
    assert len(attempts) == 3 and waits == [.01,.02]


@pytest.mark.parametrize("winerror,attempts", [(5,6), (None,1)])
def test_status_publication_does_not_hide_persistent_permission_errors(monkeypatch, tmp_path, winerror, attempts):
    path = tmp_path / "state.json"
    path.write_text('{"state":"running"}', encoding="utf-8")
    error = PermissionError("Persistent access denial"); error.winerror = winerror
    calls, waits = [], []
    def denied(*args): calls.append(args); raise error
    monkeypatch.setattr(Path, "replace", denied)
    monkeypatch.setattr(library_tasks.time, "sleep", waits.append)
    with pytest.raises(PermissionError) as failure:
        library_tasks._write(path, {"state":"completed"})
    assert failure.value is error and len(calls) == attempts
    assert sum(waits) <= .151
    assert json.loads(path.read_text(encoding="utf-8")) == {"state":"running"}


class FakeProcess:
    def __init__(self):
        self.returncode = None
        self.terminations = 0

    def poll(self): return self.returncode
    def terminate(self): self.terminations += 1; self.returncode = -15
    def kill(self): self.returncode = -9
    def wait(self, timeout=None): return self.returncode


def fake_tasks(monkeypatch, tmp_path):
    process = FakeProcess()
    clock = {"now": 100.}
    monkeypatch.setattr(library_tasks, "tempfile", SimpleNamespace(gettempdir=lambda: str(tmp_path)))
    monkeypatch.setattr(library_tasks, "time", SimpleNamespace(monotonic=lambda: clock["now"]))
    calls = []
    def spawn(*args, **kwargs): calls.append((args, kwargs)); return process
    monkeypatch.setattr(library_tasks.subprocess, "Popen", spawn)
    return process, clock, calls


def test_timeout_stops_only_its_worker_and_cleans_disposable_job(monkeypatch, tmp_path):
    process, clock, calls = fake_tasks(monkeypatch, tmp_path)
    task = library_tasks.LibraryTask({"operation": "capture"}, timeout=5)
    assert task.poll()["state"] == "running"
    clock["now"] += 6
    state = task.poll()
    assert state["state"] == "failed" and "预算" in state["error"]
    assert process.terminations == 1
    assert "--worker" in calls[0][0][0]
    task.close()
    assert not task.folder.exists()


def test_cancellation_has_grace_period_and_is_not_misreported_as_success(monkeypatch, tmp_path):
    process, clock, _calls = fake_tasks(monkeypatch, tmp_path)
    task = library_tasks.LibraryTask({"operation": "learn"})
    task.cancel()
    assert (task.folder / "cancel").is_file()
    assert task.poll()["worker_running"]
    clock["now"] += 3
    assert task.poll()["state"] == "cancelled"
    assert process.terminations == 1
    task.close()


def test_completed_state_is_preserved_after_worker_exit(monkeypatch, tmp_path):
    process, _clock, _calls = fake_tasks(monkeypatch, tmp_path)
    task = library_tasks.LibraryTask({"operation": "verify"})
    library_tasks._write(task.folder / "state.json", {"state": "completed", "result": {"ok": True}})
    process.returncode = 0
    assert task.poll()["result"]["ok"]
    task.close()


def test_cancelled_picker_never_commits_a_late_completed_result(monkeypatch, tmp_path):
    process, _clock, _calls = fake_tasks(monkeypatch, tmp_path)
    task = library_tasks.LibraryTask({"operation": "pick-region"})
    library_tasks._write(task.folder / "state.json", {"state": "completed", "result": {"selection": {}}})
    task.cancel(); process.returncode = 0
    state = task.poll()
    assert state["state"] == "cancelled" and "result" not in state
    task.close()


def test_partition_timeout_returns_saved_pixel_draft_without_hanging_gui(monkeypatch, tmp_path):
    process, clock, _calls = fake_tasks(monkeypatch, tmp_path)
    task = library_tasks.LibraryTask({"operation": "partition"}, timeout=5)
    library_tasks._write(task.folder / "state.json", {"state": "running", "partial_result": {
        "regions": [{"id": "content"}], "method": "pixel-dividers"}})
    clock["now"] += 6
    state = task.poll()
    assert process.terminations == 1 and state["state"] == "completed"
    assert state["result"]["regions"] == [{"id": "content"}]
    assert "超时" in state["result"]["warning"]
    task.close()


def test_cancelled_partition_does_not_apply_the_draft(monkeypatch, tmp_path):
    _process, clock, _calls = fake_tasks(monkeypatch, tmp_path)
    task = library_tasks.LibraryTask({"operation": "partition"}, timeout=5)
    library_tasks._write(task.folder / "state.json", {"state": "running", "partial_result": {"regions": [{}]}})
    task.cancel(); clock["now"] += 6
    assert task.poll()["state"] == "cancelled"
    task.close()


def test_actual_worker_returns_failure_and_next_job_still_runs(tmp_path, monkeypatch):
    monkeypatch.setattr(library_tasks, "tempfile", SimpleNamespace(gettempdir=lambda: str(tmp_path)))
    for _ in range(2):
        task = library_tasks.LibraryTask({"operation": "unknown"}, timeout=15)
        try:
            deadline = time.monotonic()+15
            state = task.poll()
            while state["worker_running"] and time.monotonic() < deadline:
                time.sleep(.05)
                state = task.poll()
            assert state["state"] == "failed"
            assert "未知" in state["error"]
        finally:
            task.close()


def test_invalid_selected_window_fails_before_activation(tmp_path, monkeypatch):
    from easy_uiauto.ctrl import Controller
    from easy_uiauto.mcp import native_windows

    folder = tmp_path / "job"; folder.mkdir()
    library_tasks._write(folder / "request.json", {"operation": "capture", "window": {"handle": 999}})
    monkeypatch.setattr(native_windows, "matching_windows", lambda *_a, **_kw: [])
    def forbidden(*_args, **_kwargs): raise AssertionError("Activation was attempted")
    monkeypatch.setattr(Controller, "activate_window_handle", forbidden)
    library_tasks._worker(folder)
    state = json.loads((folder / "state.json").read_text(encoding="utf-8"))
    assert state["state"] == "failed" and "失效" in state["error"]
