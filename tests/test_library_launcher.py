import sys
from types import SimpleNamespace

import pytest

from easy_uiauto import library_launcher


def test_startup_failure_is_visible_logged_and_returns_failure(tmp_path, monkeypatch):
    import tkinter
    from tkinter import messagebox

    def fail(): raise ImportError("missing optional GUI dependency")
    events = []
    monkeypatch.setitem(sys.modules, "easy_uiauto.qt_gui", SimpleNamespace(main=fail))
    monkeypatch.setattr(library_launcher, "tempfile", SimpleNamespace(gettempdir=lambda: str(tmp_path)))
    monkeypatch.setattr(tkinter, "Tk", lambda: SimpleNamespace(
        withdraw=lambda: None, destroy=lambda: events.append("closed")))
    monkeypatch.setattr(messagebox, "showerror", lambda title, detail, **kwargs: events.append(detail))
    with pytest.raises(SystemExit) as failure:
        library_launcher.main()
    assert failure.value.code == 1
    assert "missing optional GUI dependency" in events[0] and events[-1] == "closed"
    assert (tmp_path / "easy_uiauto" / "library-startup.log").is_file()


def test_worker_dispatch_does_not_initialize_the_gui(monkeypatch):
    calls=[]
    monkeypatch.setattr(sys,"argv",["app.exe","--worker","job"])
    monkeypatch.setitem(sys.modules,"easy_uiauto.library_tasks",SimpleNamespace(main=lambda:calls.append("worker")))
    monkeypatch.setitem(sys.modules,"easy_uiauto.qt_gui",SimpleNamespace(main=lambda:pytest.fail("GUI opened for worker")))
    library_launcher.main()
    assert calls==["worker"]
