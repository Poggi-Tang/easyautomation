"""Configuration persistence and the active vault must stay in sync."""
import os
from pathlib import Path
import subprocess
import sys

import pytest

from easy_uiauto.mcp import configuration
from test_knowledge_library import library, add_control
from test_knowledge_organization import simple_flow


def test_form_validation_normalizes_without_creating_directories(tmp_path, monkeypatch):
    destination = tmp_path / "new-vault"
    monkeypatch.setenv("CONFIG_TEST_ROOT", str(tmp_path))
    environment = dict(os.environ)
    values = configuration.validate_gui_configuration(
        str(Path("$CONFIG_TEST_ROOT") / "new-vault"),
        " https://example.test/v1/models ", " test-key ", " custom-model ",
    )
    assert values == dict(knowledge_dir=str(destination),
                          api_url="https://example.test/v1/chat/completions",
                          api_key="test-key", api_model="custom-model")
    assert not destination.exists() and dict(os.environ) == environment


@pytest.mark.parametrize("field,value", [
    ("knowledge_dir", ""), ("knowledge_dir", "relative/vault"),
    ("api_url", "invalid"), ("api_key", ""), ("api_key", "key\nheader"),
    ("api_model", ""),
])
def test_invalid_form_is_rejected_before_any_writes(tmp_path, field, value):
    values = dict(knowledge_dir=str(tmp_path / "new"), api_url="https://example.test/v1",
                  api_key="test-key", api_model="test-model")
    values[field] = value
    with pytest.raises(ValueError): configuration.validate_gui_configuration(**values)
    assert not (tmp_path / "new").exists()


GUI_SETUP = r'''
import sys, tkinter as tk
from pathlib import Path
from unittest.mock import patch
from easy_uiauto.knowledge_gui import KnowledgeWindow, RESULT_VIEW
from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.mcp import configuration
from easy_uiauto.mcp.config_gui import ConfigurationWindow

old_root = Path(sys.argv[1]); destination = old_root.parent / "destination"
other = KnowledgeLibrary(destination)
other.bind("second", "第二个软件", dict(title="Second", class_name="Window"))
settings = dict(knowledge_dir=str(destination), api_url="https://example.test/v1",
                api_key="test-key", api_model="custom-model")
writes, errors = [], []
def persist(**values):
    writes.append(values.copy())
    Path(values["knowledge_dir"]).mkdir(parents=True, exist_ok=True)
    return {k: v for k, v in values.items() if k != "api_key"}
def choose(key):
    app.tree.selection_set(key); app.select(); root.update()
def record():
    return next(r for r in app.library.read("demo")["controls"] if r["id"] == "control-one")
def cleanup():
    app.task = app.workflow_panel.task = None
    app.dirty = app.workflow_panel.dirty = app.usage_panel.dirty = False
    if getattr(app, "config_panel", None) and not app.config_panel.closed:
        app.config_panel.initial = app.config_panel.current_values()
    app.close()
root = tk.Tk(); root.withdraw()
app = KnowledgeWindow(root, old_root)
patches = [patch.object(configuration, "load_gui_configuration", lambda: settings.copy()),
           patch.object(configuration, "save_gui_configuration", persist),
           patch("tkinter.messagebox.showerror", lambda *a, **kw: errors.append(a)),
           patch("tkinter.messagebox.askyesnocancel", lambda *a, **kw: None)]
for guard in patches: guard.start()
'''


def run_gui(library, code):
    result = subprocess.run([sys.executable, "-X", "utf8", "-c", GUI_SETUP + code,
                             str(library.root)], capture_output=True, text=True,
                            encoding="utf-8", timeout=40)
    assert result.returncode == 0, result.stdout + result.stderr


def test_embedded_configuration_switches_panels_and_preserves_api_only_edits(library):
    add_control(library); library.save_workflow("demo", simple_flow())
    run_gui(library, r'''
try:
    choose("workflow:demo:workflow-demo:step-1")
    assert app.workflow_panel.flow and app.workflow_panel.controls
    choose("control:demo:control-one")
    assert app.viewer.image
    windows = app.windows
    windows["demo"] = dict(title="Test", handle=123)
    app.name.set("尚未保存的控件名")
    selection = app.tree.selection()
    root.deiconify(); root.update()
    config = app.configure()
    root.update()
    assert config.root.winfo_viewable() and root.grab_current() == config.root
    assert app.configure() is config
    # An explicit --vault is authoritative when opening settings.
    assert config.values["knowledge_dir"].get() == str(old_root)
    config.values["api_model"].set("another-model")
    assert config.save() and len(writes) == 1
    assert app.tree.selection() == selection and app.name.get() == "尚未保存的控件名"
    assert app.dirty and app.windows is windows and windows["demo"]["handle"] == 123
    assert app.workflow_panel.flow and record()["semantic_name"] == "打开"
    # The same path is allowed while jobs are active; it does not reset their state.
    app.task = app.workflow_panel.task = object()
    assert config.save() and len(writes) == 2
    app.task = app.workflow_panel.task = None
    app.dirty = False
    app.display_observation = {"old": "evidence"}
    app.observation_mode.set(RESULT_VIEW)
    config.values["knowledge_dir"].set(str(destination))
    assert config.save() and len(writes) == 3
    root.update()
    assert app.library.root == destination and app.path.get() == str(destination)
    assert app.usage_panel.library is app.library is app.workflow_panel.library
    assert app.tree.get_children("vault") == ("app:second",)
    assert not app.tree.exists("app:demo") and app.selected[1] == "vault"
    assert app.windows is windows is app.workflow_panel.windows and not windows
    assert app.workflow_panel.flow is None and not app.workflow_panel.controls
    assert app.workflow_panel.app_id == app.usage_panel.app_id == app.app_id == ""
    assert app.workflow_panel.bound_window.get() == "未绑定运行窗口"
    assert not app.workflow_panel.binding["values"]
    assert not app.viewer.image and not app.display_observation
    assert not app.dirty and not errors
    assert str(destination) in config.status.get()
    assert config.current_values() == config.initial
    # The standalone configuration editor still saves without parent callbacks.
    standalone = ConfigurationWindow(tk.Toplevel(root))
    assert standalone.before_save is None and standalone.on_saved is None
    assert standalone.save() and len(writes) == 4
    assert standalone.close()
finally:
    cleanup()
''')


def test_configuration_cancellation_task_guards_and_persistence_failure(library):
    add_control(library)
    run_gui(library, r'''
try:
    choose("control:demo:control-one")
    app.name.set("应保存到旧知识库")
    config = app.configure()
    config.values["knowledge_dir"].set(str(destination))
    initial = config.initial.copy()
    assert not config.save() and not writes  # Unsaved edits: Cancel.
    assert app.library.root == old_root and app.dirty
    assert app.name.get() == "应保存到旧知识库" and config.initial == initial
    for owner in (app, app.workflow_panel):
        owner.task = object()
        assert not config.save() and not writes
        owner.task = None
    assert len(errors) == 2 and all("运行" in error[1] for error in errors)
    app.dirty = False
    blocked = old_root.parent / "is-a-file"; blocked.write_text("file")
    config.values["knowledge_dir"].set(str(blocked))
    assert not config.save() and not writes
    assert app.library.root == old_root
    config.values["knowledge_dir"].set(str(destination))
    with patch.object(configuration, "save_gui_configuration", side_effect=OSError("denied")):
        assert not config.save() and not writes
    assert app.library.root == old_root and config.initial == initial
    # Saving pending edits commits them using the previous library before switching.
    app.name.set("应保存到旧知识库")
    with patch("tkinter.messagebox.askyesnocancel", return_value=True):
        assert config.save()
    assert len(writes) == 1 and app.library.root == destination
    old = KnowledgeLibrary(old_root)
    assert old.read("demo")["controls"][0]["semantic_name"] == "应保存到旧知识库"
    assert not (destination / "applications/demo").exists()
finally:
    cleanup()
''')


def test_configuration_reports_saved_settings_if_gui_refresh_fails(library):
    add_control(library)
    run_gui(library, r'''
try:
    choose("control:demo:control-one")
    app.windows["demo"] = dict(title="Test", handle=123)
    selected = app.tree.selection()
    config = app.configure(); config.values["knowledge_dir"].set(str(destination))
    original = app._replace_library
    def fail_once(library):
        original(library)
        if library.root == destination: raise OSError("refresh failed")
    with patch.object(app, "_replace_library", fail_once):
        assert not config.save()
    assert len(writes) == 1 and writes[0]["knowledge_dir"] == str(destination)
    assert app.library.root == old_root and app.path.get() == str(old_root)
    assert app.usage_panel.library is app.library is app.workflow_panel.library
    assert app.tree.selection() == selected and app.windows["demo"]["handle"] == 123
    assert "配置已保存" in config.status.get() and "切换知识库失败" in config.status.get()
    assert errors and config.current_values() == config.initial
    assert config.save() and app.library.root == destination
    # Closing the parent also respects cancellation in its configuration window.
    config.values["api_model"].set("unsaved-model")
    app.close()
    assert not app.closed and not config.closed and root.winfo_exists()
finally:
    cleanup()
''')
