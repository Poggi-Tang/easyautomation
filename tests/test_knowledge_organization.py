"""Vault organization and edit safety, with GUI checks isolated in one Tk process."""
import json
from pathlib import Path
import subprocess
import sys

import pytest

from easy_uiauto import workflows
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library, add_control


def simple_flow(control_id="control-one"):
    return dict(id="workflow-demo", kind="workflow", app_id="demo", keyword="打开", precondition="软件已启动", parameters={},
                steps=[dict(id="step-1", action="click", description="打开按钮", timeout_seconds=10,
                            target=dict(control_id=control_id, label="打开"), checks=[])])


def test_usage_links_are_app_scoped_and_existing_documents_are_not_rewritten(library):
    flow = library.save_workflow("demo", simple_flow())
    before = {p: p.read_bytes() for p in library.directory("demo").rglob("*.md")}
    note = library.save_usage("demo", "如何打开", "先选择文件，再打开。", workflow_ids=[flow["id"]])
    assert all(p.read_bytes() == data for p, data in before.items())
    assert next(r for r in library.read("demo")["usage"] if r["editable"])["body"].strip() == "先选择文件，再打开。"
    assert not next(r for r in library.usage_documents("demo") if r["document_path"] == "app.md")["editable"]
    library.bind("other", "Other", dict(title="Other", class_name="Window"))
    with pytest.raises(ValueError, match="组合动作"):
        library.save_usage("other", "跨软件引用", "", workflow_ids=[flow["id"]])
    with pytest.raises(ValueError, match="ID"):
        library.save_usage("demo", "坏路径", "", document_id="../outside")
    updated = library.save_usage("demo", "修改标题", "新说明", document_id=note["id"], workflow_ids=[flow["id"]])
    assert updated["id"] == note["id"] and len([r for r in library.usage_documents("demo") if r["editable"]]) == 1


def test_locator_changes_invalidate_results_without_copying_control_location(library):
    control = add_control(library)
    flow = library.save_workflow("demo", simple_flow())
    result = dict(ok=True, definition_hash=workflows.definition_hash(flow), control_hash=workflows.control_hash(flow, [control]))
    library.save_workflow_execution("demo", flow["id"], result)
    assert not library.workflows("demo")[0]["last_run"]["outdated"]
    control["location"]["WindowName"] = "Different"
    knowledge.save_control(library.directory("demo"), control)
    saved = library.workflows("demo")[0]
    assert saved["last_run"]["outdated"]
    assert "location" not in saved["steps"][0]["target"]


def test_new_import_label_is_compatible_with_old_case_label():
    source = (Path(__file__).parents[1]/"docs/examples/workflows/new-project-context-menu.md").read_text(encoding="utf-8")
    assert workflows.import_case(source.replace("用例名称", "动作标题"), "demo")["steps"] == workflows.import_case(source, "demo")["steps"]


def test_combination_editor_saves_visible_edits_and_preserves_window_binding_context(library):
    add_control(library); library.save_workflow("demo", simple_flow())
    code = r'''
import json, sys, tkinter as tk
from copy import deepcopy
from easy_uiauto.knowledge_gui import KnowledgeWindow
from easy_uiauto.mcp.config_gui import ConfigurationWindow
import easy_uiauto.workflow_gui as module
root=tk.Tk(); app=KnowledgeWindow(root, sys.argv[1])
try:
    root.geometry("1120x720"); root.update()
    assert app.stop_button.winfo_ismapped() and app.progress.winfo_ismapped()
    assert app.stop_button.winfo_rootx()+app.stop_button.winfo_width() <= root.winfo_rootx()+root.winfo_width()
    assert app.tree.get_children("") == ("vault",)
    assert app.tree.get_children("app:demo") == ("usage-root:demo", "controls-root:demo", "actions-root:demo")
    app.tree.selection_set("control:demo:control-one"); app.select(); root.update()
    assert all(str(w.cget("state")) == "readonly" for w in app.rect_fields.values())
    size=app.buttons["点击"].winfo_reqwidth()
    config=ConfigurationWindow(tk.Toplevel(root)); root.update()
    assert app.buttons["点击"].winfo_reqwidth() == size
    config.close()
    key="workflow:demo:workflow-demo:step-1"
    app.tree.selection_set(key); app.select(); root.update()
    panel=app.workflow_panel
    panel.fields["title"].set("修改动作标题")
    assert panel.save()
    assert panel.flow["steps"][0]["description"] == "修改动作标题"
    data=deepcopy(panel.flow); data["steps"][0]["description"]="JSON 直接保存"
    panel.editor.delete("1.0", "end"); panel.editor.insert("1.0", json.dumps(data, ensure_ascii=False))
    assert panel.save()
    assert panel.flow["steps"][0]["description"] == "JSON 直接保存"
    panel.source.insert("end", "新增来源说明"); assert panel.save()
    assert "新增来源说明" in panel.flow["source_case"]
    failures=[]
    module.messagebox.showerror=lambda *_a, **_k: failures.append(True)
    data=deepcopy(panel.flow); data["app_id"]="wrong-app"
    panel.editor.delete("1.0", "end"); panel.editor.insert("1.0", json.dumps(data))
    assert not panel.save() and failures
    panel.sync_editor(); panel.dirty=False
    captured={}
    module.WindowPicker=lambda _master, app_id, callback: captured.update(app_id=app_id, callback=callback)
    panel.bind_window(); panel.app_id="other"
    captured["callback"](dict(title="Selected", handle=42))
    assert panel.windows["demo"]["handle"] == 42 and "other" not in panel.windows
    panel.app_id="demo"
    app.tree.selection_set("usage-root:demo"); app.select(); app.usage_panel.new()
    app.usage_panel.title.set("操作说明"); app.usage_panel.body.insert("end", "如何使用")
    assert app.usage_panel.save()
    assert any(r["name"] == "操作说明" for r in app.library.usage_documents("demo"))
finally:
    app.dirty=app.workflow_panel.dirty=app.usage_panel.dirty=False; app.close()
'''
    result = subprocess.run([sys.executable, "-X", "utf8", "-c", code, str(library.root)], capture_output=True, text=True, encoding="utf-8", timeout=30)
    assert result.returncode == 0, result.stdout + result.stderr


def test_capture_guard_refuses_contaminated_capture_and_restores_scanner(tmp_path, monkeypatch):
    from easy_uiauto import library_tasks
    from easy_uiauto.mcp import scanner
    captured = []
    original = lambda _rect: captured.append(True)
    monkeypatch.setattr(scanner, "_screenshot_window", original)
    counter = iter(range(100))
    monkeypatch.setattr(library_tasks.time, "monotonic", lambda: next(counter))
    with pytest.raises(TimeoutError):
        with library_tasks._capture_guard(tmp_path, True): scanner._screenshot_window({})
    assert captured == [] and scanner._screenshot_window is original
    assert not (tmp_path/"capture-request").exists()
