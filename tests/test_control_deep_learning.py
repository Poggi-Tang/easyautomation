"""Single-entry exploration uses a deterministic app, without desktop/API calls."""
from copy import deepcopy

import pytest

from easy_uiauto import relearning, workflows
from easy_uiauto.deep_learning import Explorer
from test_deep_learning import Backend
from test_knowledge_library import library
from test_relearning import region


@pytest.mark.parametrize("mode", ["new", "resume", "improve", "update"])
def test_single_control_explores_descendants_and_archives_complete_flow(library, region, mode):
    backend = Backend(library)
    explorer = Explorer(library, "demo", region, backend, entry_control_id="toolbar", mode=mode)
    result = explorer.run()
    assert result["complete"] and backend.executed == ["toolbar", "fill"]
    assert backend.planned == ["root", "dialog", "done"]
    catalog = workflows.load_catalog(library.directory("demo"))
    roots = explorer.graph["entry_workflow_ids"]
    assert len(roots) == 1
    expanded = workflows.expand_calls(catalog[roots[0]], catalog)
    assert [s["target"]["control_id"] for s in expanded["steps"]] == ["toolbar", "fill", "result"]


def test_selected_improvement_without_prior_deep_scan_explores_new_controls(library, region):
    backend = Backend(library)
    result = Explorer(library, "demo", region, backend, mode="improve", control_ids=["toolbar"]).run()
    assert result["complete"] and backend.executed == ["toolbar", "fill"]


def test_child_missing_input_is_reported_and_can_resume_from_child(library, region):
    first = Explorer(library, "demo", region, Backend(library, blocked=True), entry_control_id="toolbar")
    result = first.run()
    assert not result["complete"] and result["unfinished"][0]["reason"] == "缺少工程路径"
    assert not result["workflow_ids"]
    backend = Backend(library); backend.current = "dialog"
    second = Explorer(library, "demo", region, backend, entry_control_id="toolbar", inputs={"路径":"D:/test"})
    result = second.run()
    assert result["complete"] and backend.executed == ["fill"]
    assert second.graph["id"] == first.graph["id"]


def test_resume_keeps_entry_scope_and_stops_at_budget(library, region):
    first = Explorer(library, "demo", region, Backend(library), entry_control_id="toolbar", max_actions=1)
    result = first.run()
    assert not result["complete"] and result["unfinished"]
    backend = Backend(library)
    second = Explorer(library, "demo", region, backend, entry_control_id="toolbar")
    assert second.run()["complete"] and backend.executed == ["toolbar", "fill"]
    assert second.graph["id"] == first.graph["id"]


def test_region_and_control_histories_cannot_select_each_other(library, region):
    region_run = Explorer(library, "demo", region, Backend(library)); region_run.run()
    control_run = Explorer(library, "demo", region, Backend(library), entry_control_id="toolbar")
    control_run.run()
    other = Explorer(library, "demo", region, Backend(library), entry_control_id="menu"); other.run()
    directory = library.directory("demo")
    assert [g["id"] for g in relearning.sessions(directory, region["id"])] == [region_run.graph["id"]]
    assert relearning.select_session(directory, region["id"], entry_control_id="toolbar")["id"] == control_run.graph["id"]
    with pytest.raises(ValueError, match="不属于"):
        Explorer(library, "demo", region, Backend(library), entry_control_id="toolbar", session_id=other.graph["id"])
    with pytest.raises(ValueError, match="不属于"):
        Explorer(library, "demo", region, Backend(library), session_id=control_run.graph["id"])


def test_click_and_right_click_of_selected_control_both_continue(library, region):
    backend = Backend(library); original = backend.plan
    def plan(observation, route):
        result = original(observation, route)
        if observation["state_id"] == "root":
            branch = deepcopy(result["branches"][1]); branch["name"] = "右键工具栏按钮"
            branch["steps"][0]["action"] = "right-click"
            result["branches"].append(branch)
        return result
    backend.plan = plan
    explorer = Explorer(library, "demo", region, backend, entry_control_id="toolbar")
    assert explorer.run()["complete"]
    assert backend.executed == ["toolbar", "fill", "toolbar"]
    assert len(explorer.graph["entry_workflow_ids"]) == 2


def test_return_to_root_does_not_enable_unrelated_entries(library, region):
    backend = Backend(library); original_plan = backend.plan; original_execute = backend.execute
    def plan(observation, route):
        result = original_plan(observation, route)
        if observation["state_id"] == "dialog":
            branch = deepcopy(result["branches"][0]); branch["name"] = "返回区域"
            branch["steps"][0]["action"] = "right-click"
            result["branches"].append(branch)
        return result
    def execute(steps):
        if steps[0]["action"] == "right-click":
            backend.executed.append("back"); backend.current = "root"; return dict(ok=True)
        return original_execute(steps)
    backend.plan = plan; backend.execute = execute
    explorer = Explorer(library, "demo", region, backend, entry_control_id="toolbar")
    assert explorer.run()["complete"]
    assert backend.executed == ["toolbar", "fill", "back"]
    assert "root" not in explorer.followup_states()


def test_unrelated_first_action_cannot_smuggle_entry_in_later_steps(library, region):
    backend = Backend(library); original = backend.plan
    def plan(observation, route):
        result = original(observation, route)
        if observation["state_id"] == "root":
            result["branches"] = [dict(name="错误入口", steps=[
                workflows.click_step(backend.controls["menu"]), workflows.click_step(backend.controls["toolbar"])])]
        return result
    backend.plan = plan
    result = Explorer(library, "demo", region, backend, entry_control_id="toolbar").run()
    assert not result["complete"] and not backend.executed
    assert any("未规划交互" in issue["reason"] for issue in result["issues"])


def test_unreached_control_is_partial_without_executing_siblings(library, region):
    backend = Backend(library)
    result = Explorer(library, "demo", region, backend, entry_control_id="absent").run()
    assert not result["complete"] and not backend.executed
    assert any("尚未到达" in issue["reason"] for issue in result["issues"])


def test_visible_sibling_in_changed_main_window_is_still_excluded(library, region):
    backend = Backend(library); observe = backend.observe; plan = backend.plan
    def observed(**kwargs):
        result = observe(**kwargs)
        if backend.current == "dialog":result["controls"].append(backend.controls["menu"])
        return result
    def planned(observation, route):
        result = plan(observation, route)
        if observation["state_id"] == "dialog":
            result["branches"].append(dict(name="无关主菜单", steps=[workflows.click_step(backend.controls["menu"])]))
        return result
    backend.observe = observed; backend.plan = planned
    result = Explorer(library, "demo", region, backend, entry_control_id="toolbar").run()
    assert result["complete"] and backend.executed == ["toolbar", "fill"]


@pytest.mark.parametrize("smuggled",[False,True])
def test_entry_only_origin_still_excludes_known_siblings_after_navigation(library,region,smuggled):
    backend=Backend(library);observe=backend.observe;plan=backend.plan
    def observed(**kwargs):
        result=observe(**kwargs)
        result["excluded_entry_control_ids"]=["menu","context"]
        if backend.current=="root":result["controls"]=[backend.controls["toolbar"]]
        elif backend.current=="dialog":result["controls"].append(backend.controls["menu"])
        return result
    def planned(observation,route):
        result=plan(observation,route)
        if backend.current=="root":result["branches"]=result["branches"][1:2]
        elif backend.current=="dialog":
            steps=([workflows.click_step(backend.controls["fill"])] if smuggled else [])
            steps.append(workflows.click_step(backend.controls["menu"]))
            result["branches"].append(dict(name="无关入口",steps=steps))
        return result
    backend.observe=observed;backend.plan=planned
    explorer=Explorer(library,"demo",region,backend,entry_control_id="toolbar")
    result=explorer.run()
    assert result["complete"] and backend.executed==["toolbar","fill"]
    assert explorer.graph["excluded_entry_control_ids"]==["context","menu"]


def test_start_screen_cannot_count_as_control_workflow_completion(library, region):
    backend = Backend(library)
    backend.plan = lambda observation, route:dict(terminal=True,branches=[],completion_checks=[
        dict(kind="exists",target=dict(control_id="toolbar"))])
    result = Explorer(library, "demo", region, backend, entry_control_id="toolbar").run()
    assert not result["complete"] and not backend.executed and not result["workflow_ids"]
    assert any("起点" in issue["reason"] for issue in result["issues"])


def test_improve_preserves_unselected_historical_branches(library, region):
    first = Explorer(library, "demo", region, Backend(library)); first.run()
    siblings = deepcopy([b for b in first.graph["states"]["root"]["branches"] if b["name"] != "toolbar"])
    backend = Backend(library); original = backend.plan
    def plan(observation, route):
        result = original(observation, route)
        if observation["state_id"] == "root":result["branches"] = result["branches"][1:2]
        return result
    backend.plan = plan
    second = Explorer(library, "demo", region, backend, mode="improve", control_ids=["toolbar"])
    assert second.run()["complete"] and backend.executed == ["toolbar", "fill"]
    assert [b for b in second.graph["states"]["root"]["branches"] if b["name"] != "toolbar"] == siblings
