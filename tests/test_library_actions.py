from copy import deepcopy
from types import SimpleNamespace

import pytest

from easy_uiauto import library_actions as actions
from easy_uiauto.mcp import knowledge
from easy_uiauto.mcp.protocol import location_from_xpath
from easy_uiauto.mcp.results import ActionError


def record(identifier, **extra):
    return {"id": identifier, "name": identifier, "app_id": "demo", "status": "observed",
            "location": location_from_xpath([{"ControlType": "WindowControl", "Name": "App"},
                                             {"ControlType": "MenuItemControl", "Name": identifier}]), **extra}


def test_plan_preserves_explicit_order_and_ignores_xpath_ancestry():
    controls = [record("File"), record("Recent", navigation_path=["File"]),
                record("New", navigation_path=["File", "Recent"])]
    assert [r["id"] for r in actions.validate_plan(controls, "New")] == ["File", "Recent", "New"]
    controls[-1]["navigation_path"] = []
    assert len(actions.validate_plan(controls, "New")) == 1


@pytest.mark.parametrize("path", [["missing"], ["New"], ["File", "File"], "File", ["File"]*21])
def test_invalid_navigation_is_rejected_before_any_input(path):
    with pytest.raises(ValueError):
        actions.validate_plan([record("File"), record("New")], "New", path)


def test_quarantined_prerequisite_and_invalid_index_block_entire_plan():
    file = record("File", quarantine={"active": True})
    new = record("New", navigation_path=["File"])
    with pytest.raises(ActionError, match="隔离"):
        actions.validate_plan([file, new], "New")
    file.pop("quarantine")
    new["location"]["Xpath"][-1]["foundIndex"] = "bad"
    with pytest.raises(ActionError):
        actions.validate_plan([file, new], "New")


@pytest.fixture
def replay(monkeypatch):
    events = []
    window = object()
    monkeypatch.setattr(actions, "bind_window", lambda _e: window)
    monkeypatch.setattr(actions, "owned_windows", lambda _e: [])
    monkeypatch.setattr(actions.Controller, "activate_window_handle", lambda *_a, **_kw: events.append("activate"))
    import uiautomation
    monkeypatch.setattr(uiautomation, "GetForegroundWindow", lambda: 1)
    monkeypatch.setattr(uiautomation, "IsIconic", lambda _h: False)

    def resolve(record, expected, main):
        assert main is window
        if record["id"] == "New" and "click:File" not in events:
            raise ActionError("control_not_found", "menu is closed")
        events.append("find:"+record["id"])
        return SimpleNamespace(Name=record["id"])

    monkeypatch.setattr(actions, "resolve_record", resolve)
    def execute(_action, *, control):
        events.append("click:"+control.Name)
        return {"executed": True}
    monkeypatch.setattr(actions.Controller, "execute_resolved", execute)
    return events, [record("File"), record("New")]


def test_replay_activates_after_first_lookup_and_locates_leaf_after_opening_menu(replay):
    events, plan = replay
    snapshots = []
    result = actions.execute_plan(plan, {"handle": 1}, checkpoint=lambda r: snapshots.append(deepcopy(r)))
    assert events == ["find:File", "activate", "click:File", "find:New", "click:New"]
    assert result["ok"] and all(s["completed"] for s in result["steps"])
    assert snapshots[0]["steps"][0]["input_attempted"] and not snapshots[0]["steps"][0]["completed"]


def test_missing_entry_does_not_activate_or_click(replay, monkeypatch):
    events, plan = replay
    def missing(*_a): raise ActionError("control_not_found", "missing")
    monkeypatch.setattr(actions, "resolve_record", missing)
    with pytest.raises(ActionError) as failure:
        actions.execute_plan(plan, {"handle": 1})
    assert events == [] and not failure.value.execution_result["input_attempted"]


def test_each_target_is_captured_after_lookup_before_click_and_after_is_separate(replay):
    events, plan = replay
    def before(record, control, _window):
        assert control.Name == record["id"]
        events.append("before:"+record["id"])
        return {"phase": "before-action", "context_image": record["id"]+"-before.png"}
    def after(record, _window):
        events.append("after:"+record["id"])
        return {"phase": "after-action", "context_image": record["id"]+"-after.png"}
    result = actions.execute_plan(plan, {"handle": 1}, before_click=before, after_click=after)
    assert events == ["find:File", "activate", "before:File", "click:File", "after:File",
                      "find:New", "before:New", "click:New", "after:New"]
    assert result["steps"][-1]["before"]["context_image"] == "New-before.png"
    assert result["steps"][-1]["after"]["context_image"] == "New-after.png"


def test_pre_capture_failure_sends_no_input(replay):
    events, plan = replay
    def failed(*_a): raise RuntimeError("target disappeared before capture")
    with pytest.raises(RuntimeError) as failure:
        actions.execute_plan(plan, {"handle": 1}, before_click=failed)
    assert events == ["find:File", "activate"]
    assert not failure.value.execution_result["input_attempted"]


def test_post_capture_failure_does_not_repeat_or_misreport_successful_input(replay):
    events, plan = replay
    def failed(*_a): raise RuntimeError("result image unavailable")
    result = actions.execute_plan(plan, {"handle": 1}, after_click=failed)
    assert result["ok"] and len([e for e in events if e.startswith("click:")]) == 2
    assert all(s["after_capture_error"] == "result image unavailable" for s in result["steps"])


def test_failed_input_is_not_retried_and_leaf_is_not_executed(replay, monkeypatch):
    events, plan = replay
    def fails(_action, *, control):
        events.append("attempt:"+control.Name)
        raise RuntimeError("UIA click failed after input")
    monkeypatch.setattr(actions.Controller, "execute_resolved", fails)
    with pytest.raises(RuntimeError) as failure:
        actions.execute_plan(plan, {"handle": 1})
    assert events == ["find:File", "activate", "attempt:File"]
    assert failure.value.execution_result["input_attempted"]


def test_failed_input_captures_failure_scene_without_retrying(replay, monkeypatch):
    events, plan = replay
    def fails(_action, *, control):
        events.append("attempt:"+control.Name)
        raise RuntimeError("partially sent input")
    monkeypatch.setattr(actions.Controller, "execute_resolved", fails)
    def failure_image(_record, _window):
        events.append("failure-image")
        return {"context_image": "failure.png", "purpose": "execution-evidence"}
    with pytest.raises(RuntimeError) as failure:
        actions.execute_plan(plan, {"handle": 1}, after_click=failure_image)
    assert events == ["find:File", "activate", "attempt:File", "failure-image"]
    assert failure.value.execution_result["failure_observation"]["context_image"] == "failure.png"


def test_foreground_interference_stops_before_input(replay, monkeypatch):
    events, plan = replay
    import uiautomation
    monkeypatch.setattr(uiautomation, "GetForegroundWindow", lambda: 999)
    with pytest.raises(ActionError, match="前台"):
        actions.execute_plan(plan, {"handle": 1})
    assert events == ["find:File", "activate"]


def test_open_menu_is_not_toggled_closed(replay, monkeypatch):
    events, plan = replay
    monkeypatch.setattr(actions, "resolve_record", lambda r, *_a: SimpleNamespace(Name=r["id"], IsEnabled=True))
    result = actions.execute_plan(plan, {"handle": 1})
    assert events == ["activate", "click:New"]
    assert result["steps"][0]["skipped"] and not result["steps"][0]["input_attempted"]


def test_visible_disabled_next_control_does_not_skip_enabling_step(replay, monkeypatch):
    events, plan = replay
    monkeypatch.setattr(actions, "resolve_record", lambda r, *_a: SimpleNamespace(
        Name=r["id"], IsEnabled=r["id"] != "New" or "click:File" in events))
    result = actions.execute_plan(plan, {"handle": 1})
    assert events == ["activate", "click:File", "click:New"]
    assert result["ok"]


def test_minimized_window_is_located_before_restore_and_rechecked_before_click(replay, monkeypatch):
    events, plan = replay
    import uiautomation
    monkeypatch.setattr(uiautomation, "IsIconic", lambda _h: True)
    def resolve(r, *_a, require_visible=True):
        events.append("visible" if require_visible else "preflight")
        return SimpleNamespace(Name=r["id"])
    monkeypatch.setattr(actions, "resolve_record", resolve)
    result = actions.execute_plan(plan[:1], {"handle": 1})
    assert result["ok"] and events == ["preflight", "activate", "visible", "click:File"]


def test_malformed_leaf_selector_blocks_plan_before_opening_menu():
    new = record("New", navigation_path=["File"])
    new["location"]["Xpath"][-1]["ControlType"] = "NotAControl"
    with pytest.raises(ActionError, match="invalid_argument"):
        actions.validate_plan([record("File"), new], "New")


def test_rescan_preserves_navigation_and_click_outcomes_without_unquarantining():
    old = record("New", navigation_path=["File"], window_scope={"kind": "owned", "class_name": "#32768"},
                 quarantine={"active": True}, last_click_execution={"ok": True})
    result = knowledge.merge_control_state(old, record("New"))
    assert result["navigation_path"] == ["File"]
    assert result["window_scope"]["class_name"] == "#32768"
    assert result["last_click_execution"]["ok"]
    assert knowledge.runtime_quarantined(result)


def test_owned_scope_does_not_include_same_process_unrelated_windows(monkeypatch):
    windows = [dict(handle=1, process_id=10, owner_handle=0),
               dict(handle=2, process_id=10, owner_handle=1),
               dict(handle=3, process_id=10, owner_handle=2),
               dict(handle=4, process_id=10, owner_handle=0),
               dict(handle=5, process_id=99, owner_handle=1)]
    monkeypatch.setattr(actions.native_windows, "enumerate_windows", lambda **_kw: windows)
    assert [w["handle"] for w in actions.owned_windows(dict(handle=1, process_id=10))] == [2, 3]


def test_native_menu_without_gw_owner_requires_exact_menu_loop_owner(monkeypatch):
    windows = [dict(handle=1, process_id=10, owner_handle=0),
               dict(handle=2, process_id=10, owner_handle=0, class_name="#32768"),
               dict(handle=3, process_id=10, owner_handle=0, class_name="#32768"),
               dict(handle=4, process_id=99, owner_handle=0, class_name="#32768")]
    monkeypatch.setattr(actions.native_windows, "enumerate_windows", lambda **_kw: windows)
    calls = []
    monkeypatch.setattr(actions.native_windows, "menu_owner", lambda h: calls.append(h) or {2:1, 3:999}[h])
    assert [w["handle"] for w in actions.owned_windows(dict(handle=1, process_id=10))] == [2]
    assert calls == [2, 3]  # Never bind a different process merely by its menu class.


def test_owned_dialog_full_xpath_is_rebased_at_exact_bound_root_without_rewriting_record(monkeypatch):
    import uiautomation
    dialog = SimpleNamespace(ControlTypeName="WindowControl", Name="Dialog", ClassName="DialogClass", AutomationId="")
    result = SimpleNamespace(Name="Input")
    target = record("Input", window_scope=dict(kind="owned", class_name="DialogClass", title="Dialog"))
    target["location"] = location_from_xpath([
        dict(ControlType="WindowControl", Name="App", searchDepth=1),
        dict(ControlType="WindowControl", Name="Dialog", ClassName="DialogClass", searchDepth=2),
        dict(ControlType="EditControl", Name="Input", searchDepth=3)])
    before = deepcopy(target)
    monkeypatch.setattr(actions, "owned_windows", lambda _e: [dict(handle=2, class_name="DialogClass", title="Dialog")])
    monkeypatch.setattr(uiautomation, "ControlFromHandle", lambda _h: dialog)
    def observe(location, *, window):
        assert window is dialog
        assert [s["Name"] for s in location["Xpath"]] == ["Dialog", "Input"]
        assert [s["searchDepth"] for s in location["Xpath"]] == [1, 2]
        return result
    monkeypatch.setattr(actions, "observe_control", observe)
    assert actions.resolve_record(target, dict(handle=1), object(), require_visible=False) is result
    assert target == before
    target["location"]["Xpath"][1]["Name"] = "Other dialog"
    with pytest.raises(ActionError, match="上下文不一致"):
        actions.resolve_record(target, dict(handle=1), object(), require_visible=False)
