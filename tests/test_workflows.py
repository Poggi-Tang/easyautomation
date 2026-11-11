from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace

import pytest

from easy_uiauto import workflows
from easy_uiauto.mcp.protocol import location_from_xpath
from easy_uiauto.mcp.results import ActionError
from test_knowledge_library import library

CASE = (Path(__file__).parents[1]/"docs/examples/workflows/new-project-context-menu.md").read_text(encoding="utf-8")


def record(identifier, kind="ButtonControl", name=None):
    return dict(id=identifier, name=name or identifier, control_type=kind, status="observed",
        location=location_from_xpath([dict(ControlType="WindowControl", Name="Test"),
                                     dict(ControlType=kind, Name=name or identifier)]))


def bound_case():
    flow = workflows.import_case(CASE, "demo")
    controls = []
    for n, (_id, target) in enumerate(workflows.targets(flow)):
        target["control_id"] = f"control-{n}"
        controls.append(record(target["control_id"], target.get("control_type") or "TextControl", target["label"]))
    return flow, controls


def test_table_compiles_to_one_keyword_with_exact_scope_and_all_expectations():
    flow = workflows.import_case(CASE, "demo")
    assert flow["keyword"] == "项目管理.新建项目-右键菜单打开对话框"
    assert [s["action"] for s in flow["steps"]] == ["right-click", "click", "inspect"]
    assert flow["steps"][0]["target"]["hit"] == "blank"
    menu = flow["steps"][0]["checks"][0]
    assert menu["expected"] == ["新建项目", "打开项目", "导入", "关闭所有项目", "粘贴", "展开所有", "折叠所有", "重新加载"]
    assert flow["steps"][1]["checks"][0]["expected"] == "新建项目"
    assert flow["steps"][1]["checks"][0]["target"]["control_type"] == "WindowControl"
    checks = flow["steps"][2]["checks"]
    assert [c["target"]["label"] for c in checks] == ["选择存储路径", "选择", "设置项目名称", "取消", "确定"]
    assert checks[0]["expected"] is True and checks[2]["expected"] is False
    assert flow["steps"][-1]["action"] == "inspect"  # Never adds a confirm click.
    assert not workflows.readiness(flow, [])["ready"]
    assert len(list(workflows.targets(flow))) == 10


def test_unsupported_prose_is_preserved_as_unexecutable_instead_of_guessed():
    flow = workflows.import_case(CASE.replace("右键单击【项目管理】空白处", "完成一切初始化"), "demo")
    assert flow["steps"][0]["action"] == "manual"
    assert any("尚未转换" in i["message"] for i in workflows.readiness(flow, [])["issues"])
    with pytest.raises(ValueError): workflows.import_case("random text", "demo")


def test_binding_requires_one_exact_name_and_rejects_type_mismatch_and_quarantine():
    flow = workflows.import_case(CASE, "demo")
    controls = [record("one", "MenuItemControl", "新建项目"), record("two", "MenuItemControl", "新建项目")]
    proposed = workflows.suggest_bindings(flow, controls)
    assert not proposed["steps"][1]["target"]["control_id"]
    proposed = workflows.suggest_bindings(flow, controls[:1])
    assert proposed["steps"][1]["target"]["control_id"] == "one"
    flow, controls = bound_case()
    controls[0]["status"] = "quarantined"
    controls[0]["quarantine"] = dict(active=True,reason='Runtime positioning failed')
    assert not workflows.readiness(flow, controls)["ready"]
    controls[0]["status"] = "observed"
    controls[0].pop('quarantine')
    controls[2]["control_type"] = "ButtonControl"
    assert not workflows.readiness(flow, controls)["ready"]


def test_saving_flow_keeps_existing_control_knowledge_unchanged(library):
    from test_knowledge_library import add_control
    add_control(library)
    original = deepcopy(library.read("demo")["controls"])
    flow = library.import_workflow_case("demo", CASE)
    assert flow["status"] == "draft"
    path = library.directory("demo")/"workflows"/(flow["id"]+".md")
    assert path.is_file() and library.read("demo")["controls"] == original
    assert library.workflows("demo")[0]["keyword"] == flow["keyword"]
    library.save_workflow_execution("demo", flow["id"], dict(ok=False, failed_step="step-1", input_attempted=False))
    assert library.workflows("demo")[0]["last_run"]["failed_step"] == "step-1"
    assert library.read("demo")["controls"] == original


@pytest.fixture
def runtime(monkeypatch):
    import uiautomation
    from easy_uiauto import library_actions
    from easy_uiauto.ctrl import Controller

    flow, controls = bound_case()
    events = []
    monkeypatch.setattr(library_actions, "bind_window", lambda _w: events.append("bind") or object())
    monkeypatch.setattr(library_actions, "owned_windows", lambda _w: [])
    monkeypatch.setattr(Controller, "activate_window_handle", lambda *_a, **_k: events.append("activate"))
    monkeypatch.setattr(uiautomation, "GetForegroundWindow", lambda: 1)
    monkeypatch.setattr(uiautomation, "IsIconic", lambda _h: False)
    menu_expected = flow["steps"][0]["checks"][0]["expected"]
    objects = {}
    for (_id, target) in workflows.targets(flow):
        read_only = target["label"] == "选择存储路径"
        c = SimpleNamespace(Name=target["label"], IsEnabled=True, IsOffscreen=False,
            GetValuePattern=lambda read_only=read_only: SimpleNamespace(IsReadOnly=read_only),
            GetChildren=lambda: [])
        if target.get("control_type") == "MenuControl":
            c.GetChildren = lambda: [SimpleNamespace(Name=name, ControlTypeName="MenuItemControl", IsOffscreen=False, GetChildren=lambda: []) for name in menu_expected]
        objects[target["control_id"]] = c
    def resolve(r, *_args):
        if r["id"] == flow["steps"][1]["target"]["control_id"] and "right-click" not in events:
            raise ActionError("control_not_found", "menu closed")
        events.append("find:"+r["id"])
        return objects[r["id"]]
    monkeypatch.setattr(library_actions, "resolve_record", resolve)
    def execute(action, **_kwargs): events.append(action); return dict(executed=True)
    monkeypatch.setattr(Controller, "execute_resolved", execute)
    monkeypatch.setattr(workflows, "_blank_offset", lambda _c: (10, 20))
    return flow, controls, events, objects


def test_executes_ordered_actions_once_then_checks_properties_without_confirm_click(runtime):
    flow, controls, events, _objects = runtime
    original = deepcopy(controls)
    result = workflows.execute(flow, controls, dict(handle=1))
    assert result["ok"] and len(result["steps"]) == 3
    assert all(s["ok"] for s in result["steps"])
    assert events.index("right-click") < events.index("click")
    assert events.count("right-click") == events.count("click") == 1
    assert events.index("find:control-0") < events.index("activate")
    assert sum(len(s["checks"]) for s in result["steps"]) == 7
    assert controls == original


def test_missing_bindings_fail_before_activation_or_any_input(runtime):
    flow, controls, events, _objects = runtime
    flow["steps"][-1]["checks"][-1]["target"]["control_id"] = ""
    with pytest.raises(ValueError, match="尚不可运行"):
        workflows.execute(flow, controls, dict(handle=1))
    assert events == []


def test_failed_menu_assertion_stops_before_new_and_records_failed_check(runtime):
    flow, controls, events, objects = runtime
    flow["steps"][0]["timeout_seconds"] = 0
    menu_target = flow["steps"][0]["checks"][0]["target"]["control_id"]
    objects[menu_target].GetChildren = lambda: []
    with pytest.raises(ValueError) as failed: workflows.execute(flow, controls, dict(handle=1))
    result = failed.value.execution_result
    assert result["failed_step"] == "step-1" and result["failed_check"] == "step-1.check-1"
    assert events.count("right-click") == 1 and "click" not in events


def test_editability_failure_is_not_reported_as_pass(runtime):
    flow, controls, events, objects = runtime
    flow["steps"][-1]["timeout_seconds"] = 0
    edit = flow["steps"][-1]["checks"][2]["target"]["control_id"]
    objects[edit].GetValuePattern = lambda: SimpleNamespace(IsReadOnly=True)
    with pytest.raises(ValueError) as failed: workflows.execute(flow, controls, dict(handle=1))
    assert failed.value.execution_result["failed_check"] == "step-3.check-3"
    assert events.count("click") == 1


def test_unknown_readonly_pattern_and_missing_parameter_stop_without_guessing(runtime):
    flow, controls, events, _objects = runtime
    flow["steps"][-1]["text"] = "${project_name}"
    with pytest.raises(ValueError, match="缺少流程参数"):
        workflows.execute(flow, controls, dict(handle=1))
    assert events == []
    with pytest.raises(ValueError, match="ValuePattern"):
        workflows._check(dict(kind="read-only", expected=False), SimpleNamespace(GetValuePattern=lambda: None))


def test_interrupted_input_is_not_retried(runtime, monkeypatch):
    from easy_uiauto.ctrl import Controller
    flow, controls, events, _objects = runtime
    def fail(*_a, **_k): events.append("attempt"); raise RuntimeError("partial input")
    monkeypatch.setattr(Controller, "execute_resolved", fail)
    with pytest.raises(RuntimeError) as failed: workflows.execute(flow, controls, dict(handle=1))
    assert events.count("attempt") == 1 and failed.value.execution_result["input_attempted"]


def test_gui_uses_keyword_as_root_and_imports_without_performing_actions(library):
    # Start one Tk interpreter, as the application does. Repeated Tk destruction
    # in this Windows Python runtime can intermittently fail loading init.tcl.
    import subprocess
    import sys
    code = '''
import sys, tkinter as tk
from pathlib import Path
from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.workflow_gui import WorkflowWindow
import easy_uiauto.workflow_gui as module
root = tk.Tk(); root.withdraw(); panel = None
try:
    panel = WorkflowWindow(tk.Toplevel(root), KnowledgeLibrary(sys.argv[1]), {}, "demo")
    panel.source.insert("1.0", Path(sys.argv[2]).read_text(encoding="utf-8")); panel.convert()
    assert panel.tree.get_children("") == (panel.flow["id"],)
    assert panel.tree.item(panel.flow["id"], "text") == panel.flow["keyword"]
    assert len(panel.tree.get_children(panel.flow["id"])) == 3
    calls = []
    module.LibraryTask = lambda *_a, **_k: calls.append("run")
    panel.run()
    assert calls == [] and "待完善" in panel.status.get()
    panel.keyword.set("changed-keyword")
    assert panel.dirty and panel.save()
    assert panel.flow["keyword"] == "changed-keyword"
finally:
    if panel: panel.dirty = False; panel.close()
    root.destroy()
'''
    completed = subprocess.run([sys.executable, "-X", "utf8", "-c", code, str(library.root),
        str(Path(__file__).parents[1]/"docs/examples/workflows/new-project-context-menu.md")],
        capture_output=True, text=True, encoding="utf-8", timeout=30)
    assert completed.returncode == 0, completed.stdout+completed.stderr


@pytest.mark.parametrize("change", [
    lambda f: f.update(precondition=[]),
    lambda f: f.update(steps=[None]),
    lambda f: f["steps"][0].update(checks=[None]),
    lambda f: f["steps"][0]["target"].update(label=[]),
    lambda f: f["steps"][0].update(action=[]),
    lambda f: f["steps"][0]["checks"][0].update(kind=[]),
    lambda f: f["steps"][0]["target"].update(hit=[]),
])
def test_malformed_schema_reports_validation_error(change):
    flow = workflows.import_case(CASE, "demo"); change(flow)
    with pytest.raises(ValueError): workflows.validate(flow)


def test_import_cannot_overwrite_bindings_and_changed_definition_marks_old_result(library):
    flow = library.import_workflow_case("demo", CASE)
    library.save_workflow_execution("demo", flow["id"], dict(ok=True, definition_hash=workflows.definition_hash(flow),
        control_hash=workflows.control_hash(flow, library.read("demo")["controls"])))
    with pytest.raises(ValueError, match="已存在"):
        library.import_workflow_case("demo", CASE)
    saved = library.workflows("demo")[0]
    assert saved["last_run"]["ok"] and not saved["last_run"]["outdated"]
    saved["steps"][0]["description"] = "修改后的步骤"
    changed = library.save_workflow("demo", saved)
    assert changed["last_run"]["outdated"]


def test_foreground_change_during_capture_stops_before_input_and_saves_failure(runtime, monkeypatch):
    import uiautomation
    flow, controls, events, _objects = runtime
    def before(*_args):
        monkeypatch.setattr(uiautomation, "GetForegroundWindow", lambda: 2)
        return dict(image="execution-before.png")
    with pytest.raises(ActionError) as failed:
        workflows.execute(flow, controls, dict(handle=1), before_action=before,
                          after_action=lambda *_a: dict(image="failure.png"))
    result = failed.value.execution_result
    assert not result["input_attempted"] and "right-click" not in events
    assert result["failure_observation"]["image"] == "failure.png"


def test_blank_hit_requires_exact_container_and_checks_points_within_bounds(monkeypatch):
    import uiautomation
    from easy_uiauto.mcp import scanner
    container, child = object(), object()
    hits = []
    monkeypatch.setattr(scanner, "_rect", lambda _c: dict(left=100, top=200, width=2, height=2))
    monkeypatch.setattr(uiautomation, "ControlsAreSame", lambda a, b: a is b)
    monkeypatch.setattr(uiautomation, "ControlFromPoint", lambda x, y: hits.append((x, y)) or child)
    with pytest.raises(ValueError, match="空白处"): workflows._blank_offset(container)
    assert all(100 <= x < 102 and 200 <= y < 202 for x, y in hits)
    monkeypatch.setattr(uiautomation, "ControlFromPoint", lambda _x, _y: container)
    assert workflows._blank_offset(container) == (1, 1)


def test_result_evidence_is_saved_after_each_action_and_inspection(runtime):
    flow, controls, _events, _objects = runtime
    result = workflows.execute(flow, controls, dict(handle=1),
        before_action=lambda *_a: dict(image="before.png"), after_action=lambda *_a: dict(image="after.png"))
    assert [step["after"]["image"] for step in result["steps"]] == ["after.png"]*3
    assert result["steps"][0]["before"]["image"] == "before.png"
    assert sum(bool(step.get("completed")) for step in result["steps"]) == 2


def test_workflow_worker_saves_failure_result_and_does_not_modify_learning(library, monkeypatch, tmp_path):
    from easy_uiauto import library_tasks
    library.import_workflow_case("demo", CASE)
    flow = library.workflows("demo")[0]
    def fail(*_a, **_k):
        error = ValueError("assertion failed")
        error.execution_result = dict(ok=False, failed_step="step-2", input_attempted=True, steps=[])
        raise error
    monkeypatch.setattr(workflows, "execute", fail)
    folder = tmp_path/"job"; folder.mkdir()
    library_tasks._write(folder/"request.json", dict(operation="workflow", vault=str(library.root),
                        app_id="demo", flow_id=flow["id"], window={}))
    library_tasks._worker(folder)
    import json
    state = json.loads((folder/"state.json").read_text(encoding="utf-8"))
    assert state["state"] == "failed"
    assert library.workflows("demo")[0]["last_run"]["failed_step"] == "step-2"
