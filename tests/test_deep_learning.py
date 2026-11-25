"""Branch traversal tests use a deterministic app model, never desktop input."""
from copy import deepcopy

import pytest

from easy_uiauto.deep_learning import Explorer, validate_plan
from easy_uiauto import workflows
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library, add_control
from test_region_picker import picked


class Backend:
    def __init__(self, library, *, blocked=False, fail=False):
        self.library=library;self.current="root";self.executed=[];self.planned=[];self.fail=fail;self.blocked=blocked
        base=add_control(library);self.controls={}
        for identifier in ("menu","toolbar","context","fill","result"):
            record={**base,"id":identifier,"name":identifier,"semantic_name":identifier,"actions":[] if identifier=="result" else ["click"]}
            knowledge.save_control(library.directory("demo"),record);self.controls[identifier]=record

    def observe(self, **_):
        ids={"root":["menu","toolbar","context"],"dialog":["fill"],"done":["result"]}[self.current]
        return dict(state_id=self.current,name=self.current,controls=[self.controls[i] for i in ids])

    def plan(self, observation, route):
        state=observation["state_id"];self.planned.append(state)
        if state=="done":return dict(name="创建成功",terminal=True,completion_checks=[
            dict(kind="text",target={"control_id":"result"},expected="创建成功")],branches=[])
        branches=[dict(name=i,steps=[workflows.click_step(self.controls[i])],effect="navigation",replayable=True)
                  for i in (["menu","toolbar","context"] if state=="root" else ["fill"])]
        if self.blocked and state=="dialog":branches=[dict(name="填写信息",control_id="fill",blocked="缺少工程路径")]
        return dict(name=state,branches=branches)

    def execute(self, steps):
        identifier=steps[0]["target"]["control_id"];self.executed.append(identifier)
        if self.fail:raise InterruptedError("cancelled during input")
        self.current="done" if identifier=="fill" else "dialog"
        return dict(ok=True,steps=[dict(input_attempted=True,ok=True)])

    def assert_state(self, observation):assert self.current==observation["state_id"]
    def verify(self, observation, checks):assert self.current=="done" and checks[0]["expected"]=="创建成功"
    def restore(self, parent, child, steps, route):self.current=parent["state_id"]


def test_three_entries_share_one_verified_suffix(library,tmp_path):
    identifier=library.save_picked_region("demo",picked(tmp_path));region=library.regions("demo")[0]
    backend=Backend(library);explorer=Explorer(library,"demo",region,backend)
    result=explorer.run()
    assert result["complete"] and backend.executed==["menu","fill","toolbar","context"]
    assert backend.planned==["root","dialog","done"]
    roots=explorer.graph["entry_workflow_ids"]
    catalog=workflows.load_catalog(library.directory("demo"))
    assert len(roots)==3
    calls={catalog[i]["steps"][-1]["workflow_id"] for i in roots}
    assert len(calls)==1
    for flow in roots:
        expanded=workflows.expand_calls(catalog[flow],catalog)
        assert [s["action"] for s in expanded["steps"]]==["click","click","inspect"]
        assert workflows.readiness(catalog[flow],library.read("demo")["controls"],catalog)["ready"]
    assert explorer.path.is_file()
    assert any("深度学习进度" in r["name"] for r in library.usage_documents("demo"))


@pytest.mark.parametrize("provider_title", ["【工程管理】新建工程-结果校验", "创建成功", None])
def test_generated_titles_follow_outcome_and_distinguish_entry_paths(library, tmp_path, monkeypatch, provider_title):
    library.save_picked_region("demo", picked(tmp_path))
    region = {**library.regions("demo")[0], "name": "工程管理"}
    backend = Backend(library)
    original_plan = backend.plan
    labels = {"menu": "文件菜单", "toolbar": "工具栏", "context": "右键菜单", "fill": "填写并确认"}

    def plan(observation, route):
        response = original_plan(observation, route)
        for branch in response["branches"]:
            branch["name"] = labels[branch["name"]]
        if response.get("terminal"):
            response["workflow_title"] = provider_title
        return response

    monkeypatch.setattr(backend, "plan", plan)
    feature = "新建工程" if provider_title == "【工程管理】新建工程-结果校验" else "创建成功"
    heading = f"【工程管理】{feature}"
    existing_title = f"{heading}-文件菜单／填写并确认"
    library.save_workflow("demo", dict(keyword=existing_title, name=existing_title,
        steps=[dict(id="manual", action="manual")]))

    explorer = Explorer(library, "demo", region, backend)
    assert explorer.run()["complete"]
    catalog = workflows.load_catalog(library.directory("demo"))
    generated = [catalog[identifier] for identifier in explorer.graph["workflow_ids"]]
    assert {flow["keyword"] for flow in generated} == {
        f"{heading}-结果校验", f"{heading}-填写并确认", f"{existing_title}（2）",
        f"{heading}-工具栏／填写并确认", f"{heading}-右键菜单／填写并确认",
    }
    assert all(flow["name"] == flow["keyword"] for flow in generated)
    assert {catalog[identifier]["keyword"] for identifier in explorer.graph["entry_workflow_ids"]} == {
        f"{existing_title}（2）", f"{heading}-工具栏／填写并确认", f"{heading}-右键菜单／填写并确认",
    }


@pytest.mark.parametrize("correct_result", [True, False])
def test_business_learning_archives_only_after_the_target_result_is_verified(library, monkeypatch, correct_result):
    from easy_uiauto.business_processes import from_topics

    process, = from_topics([dict(title="新建工程", process_title="【工程管理】新建工程-文件菜单",
                                 description="创建工程", steps=["打开菜单", "填写名称并确定"],
                                 expected=["工程出现在树中"], quote="新建工程")], dict(id="guide", name="用户手册"))
    backend = Backend(library)
    backend.business_process = process
    original_plan = backend.plan

    def plan(observation, route):
        result = original_plan(observation, route)
        if observation["state_id"] == "root":
            result["branches"] = result["branches"][:1]
        if result.get("terminal"):
            result["workflow_title"] = "【工程管理】新建工程-结果校验" if correct_result else "【文件】打开-结果校验"
        return result

    monkeypatch.setattr(backend, "plan", plan)
    explorer = Explorer(library, "demo", dict(id=process["id"], name=process["name"]), backend)
    run = explorer.run()
    catalog = workflows.load_catalog(library.directory("demo"))
    assert backend.executed == ["menu", "fill"]
    if correct_result:
        assert run["complete"]
        entry, = [catalog[i] for i in explorer.graph["entry_workflow_ids"]]
        assert entry["keyword"] == process["name"] and entry["business_process_id"] == process["id"]
        assert sum(flow.get("business_process_id") == process["id"] for flow in catalog.values()) == 1
    else:
        assert not run["complete"] and not catalog


def test_missing_inputs_are_partial_not_successful_workflows(library,tmp_path):
    library.save_picked_region("demo",picked(tmp_path));backend=Backend(library,blocked=True)
    result=Explorer(library,"demo",library.regions("demo")[0],backend).run()
    assert not result["complete"] and not result["workflow_ids"]
    assert result["unfinished"][0]["reason"]=="缺少工程路径"
    assert backend.executed==["menu","toolbar","context"]


def test_resume_replays_navigation_but_not_completed_business_suffix(library,tmp_path):
    library.save_picked_region("demo",picked(tmp_path));region=library.regions("demo")[0]
    first=Backend(library)
    report=Explorer(library,"demo",region,first,max_actions=1).run()
    assert not report["complete"] and first.executed==["menu"]
    second=Backend(library)
    report=Explorer(library,"demo",region,second,max_actions=20).run()
    assert report["complete"] and second.executed==["menu","fill","toolbar","context"]


def test_interrupted_input_is_never_retried_on_resume(library,tmp_path):
    library.save_picked_region("demo",picked(tmp_path));region=library.regions("demo")[0]
    with pytest.raises(InterruptedError):Explorer(library,"demo",region,Backend(library,fail=True)).run()
    backend=Backend(library);report=Explorer(library,"demo",region,backend).run()
    assert "menu" not in backend.executed and not report["complete"]
    assert any(b["status"]=="uncertain" for b in report["unfinished"])


def test_planner_cannot_fabricate_targets_or_claim_completion_without_checks(library):
    record=add_control(library)
    with pytest.raises(ValueError,match="未观察"):
        validate_plan(dict(branches=[dict(name="bad",steps=[workflows.click_step(dict(id="invented"))])]),[record])
    with pytest.raises(ValueError,match="完成状态"):
        validate_plan(dict(branches=[],terminal=True),[record])


def test_invalid_plan_preserves_partial_progress(library,tmp_path):
    import json
    library.save_picked_region("demo",picked(tmp_path));backend=Backend(library)
    backend.plan=lambda *args:dict(branches=[],terminal=True,completion_checks=[dict(kind="exists")])
    explorer=Explorer(library,"demo",library.regions("demo")[0],backend)
    with pytest.raises(ValueError,match="target"):
        explorer.run()
    saved=json.loads(explorer.path.read_text(encoding="utf-8"))
    assert saved["status"]=="partial" and "target" in saved["stop_reason"]
    assert not backend.executed
    assert any("深度学习进度" in r["name"] for r in library.usage_documents("demo"))


def test_invalid_plan_rolls_back_provisional_native_observation(library,tmp_path):
    from types import SimpleNamespace
    from easy_uiauto.region_learning import NativeBackend
    from easy_uiauto.mcp import knowledge

    library.save_picked_region("demo",picked(tmp_path));backend=Backend(library)
    directory = library.directory('demo')
    control = add_control(library)
    path = knowledge.find_control_record(directory,control['id'])[0]
    original = path.read_bytes()
    new_page = directory/'pages/provisional.md'
    state = SimpleNamespace(directory=directory,library=library,revision=None,cache={})
    backend.observation_transaction = lambda: NativeBackend.observation_transaction(state)
    observe = backend.observe
    def provisional(**kwargs):
        path.write_text('changed',encoding='utf-8')
        new_page.write_text('provisional',encoding='utf-8')
        return observe(**kwargs)
    backend.observe = provisional
    backend.plan = lambda *args: dict(branches=[],terminal='unknown')
    with pytest.raises(ValueError,match='terminal'):
        Explorer(library,'demo',library.regions('demo')[0],backend).run()
    assert path.read_bytes()==original and not new_page.exists()
    assert not backend.executed


@pytest.mark.parametrize('value,expected',[(' true ',True),('False',False),(False,False)])
def test_plan_normalizes_only_unambiguous_boolean_strings(value,expected):
    from easy_uiauto.deep_learning import normalize_plan
    assert normalize_plan(dict(branches=[],terminal=value))['terminal'] is expected


def test_failed_restore_preserves_verified_flow_and_reports_stop(library,tmp_path):
    library.save_picked_region("demo",picked(tmp_path));backend=Backend(library)
    def fail(*args):raise ValueError("cannot return")
    backend.restore=fail
    report=Explorer(library,"demo",library.regions("demo")[0],backend).run()
    assert not report["complete"] and "cannot return" in report["message"]
    assert report["workflow_ids"] and backend.executed==["menu","fill"]


def test_archive_failure_keeps_progress_and_cannot_claim_complete(library,tmp_path,monkeypatch):
    import json
    library.save_picked_region("demo",picked(tmp_path))
    explorer=Explorer(library,"demo",library.regions("demo")[0],Backend(library))
    def fail(*args):raise OSError("disk unavailable")
    monkeypatch.setattr(library,"save_workflow",fail)
    result=explorer.run()
    assert not result["complete"] and "归档未完成" in result["message"]
    assert json.loads(explorer.path.read_text(encoding="utf-8"))["status"]=="partial"


def test_generated_shared_workflow_reuses_user_parameter_defaults(library,tmp_path):
    library.save_picked_region("demo",picked(tmp_path));backend=Backend(library)
    explorer=Explorer(library,"demo",library.regions("demo")[0],backend,inputs={"结果":"创建成功"})
    explorer.run()
    catalog=workflows.load_catalog(library.directory("demo"))
    result_flow=next(f for f in catalog.values() if f["steps"][0]["action"]=="inspect")
    assert result_flow["steps"][0]["checks"][0]["expected"]=="${结果}"
    root=catalog[explorer.graph["entry_workflow_ids"][0]]
    expanded=workflows.expand_calls(root,catalog,parameters={"结果":"另一个结果"})
    assert expanded["steps"][-1]["checks"][0]["expected"]=="另一个结果"


@pytest.mark.parametrize("terminal", [False, True])
def test_inspection_and_click_do_not_cover_search_input_or_context_menu(library, tmp_path, terminal):
    library.save_picked_region("demo", picked(tmp_path))
    backend = Backend(library)
    backend.controls["menu"]["actions"] = ["click", "set-text", "right-click"]
    backend.observe = lambda **_:dict(state_id="root", controls=[backend.controls["menu"]])
    calls = []
    def plan(observation, route):
        calls.append(deepcopy(observation.get("planning_feedback", {})))
        step = workflows.click_step(backend.controls["menu"], action="inspect")
        step["checks"] = [dict(kind="exists", target=dict(control_id="menu"))]
        return dict(terminal=terminal, branches=[dict(name="仅观察", steps=[step])],
                    completion_checks=[dict(kind="exists", target=dict(control_id="menu"))] if terminal else [])
    backend.plan = plan
    backend.execute = lambda steps:dict(ok=True)
    backend.verify = lambda *args:dict(ok=True)
    explorer = Explorer(library, "demo", library.regions("demo")[0], backend)
    result = explorer.run()
    assert not result["complete"]
    assert len(calls) == 3
    assert {gap["action"] for gap in calls[1]["missing_interactions"]} == {"click", "set-text", "right-click"}
    assert {gap["action"] for gap in explorer.graph["states"]["root"]["interaction_gaps"]} == {"click", "set-text", "right-click"}


def test_missing_search_action_is_replanned_then_executed(library, tmp_path):
    library.save_picked_region("demo", picked(tmp_path))
    backend = Backend(library)
    backend.controls["menu"].update(name="搜索", semantic_name="搜索", actions=["click", "set-text", "right-click"])
    seen = []
    def observe(**kwargs):
        ids = ["menu"] if backend.current == "root" else ["result"]
        return dict(state_id=backend.current, controls=[backend.controls[i] for i in ids])
    def plan(observation, route):
        if backend.current == "done":
            return dict(terminal=True, branches=[], completion_checks=[dict(kind="text", expected="搜索结果", target=dict(control_id="result"))])
        seen.append(deepcopy(observation.get("planning_feedback", {})))
        actions = ["click"] if len(seen) == 1 else ["click", "set-text", "right-click"]
        return dict(branches=[dict(name=action, effect="input" if action == "set-text" else "navigation", replayable=action != "set-text",
            steps=[dict(action=action, target=dict(control_id="menu"), **({"text":"电阻"} if action == "set-text" else {}))]) for action in actions])
    restored = []
    def execute(steps):
        action = steps[0]["action"]; backend.executed.append(action)
        backend.current = "done" if action == "set-text" else "root"
        return dict(ok=True, **({"input_values":{"menu":"原查询"}} if action == "set-text" else {}))
    def restore(parent, child, steps, route, **kwargs):
        restored.append(kwargs); backend.current = parent["state_id"]
    backend.observe, backend.plan, backend.execute, backend.restore = observe, plan, execute, restore
    backend.verify = lambda *args:dict(ok=True)
    explorer = Explorer(library, "demo", library.regions("demo")[0], backend)
    result = explorer.run()
    assert result["complete"] and result["workflow_ids"]
    assert result["coverage"] == dict(executed_branches=3, verified_outcomes=1, blocked_branches=0, missing_actions=0)
    assert backend.executed == ["click", "set-text", "right-click"]
    assert {gap["action"] for gap in seen[1]["missing_interactions"]} == {"set-text", "right-click"}
    assert {"input_values":{"menu":"原查询"}} in restored
    assert explorer.graph["states"]["root"]["branches"][1]["recovery_verified"]


def test_legacy_complete_session_rechecks_action_coverage(library, tmp_path):
    import json
    library.save_picked_region("demo", picked(tmp_path)); region = library.regions("demo")[0]
    first = Explorer(library, "demo", region, Backend(library))
    assert first.run()["complete"]
    graph = json.loads(first.path.read_text(encoding="utf-8")); graph.pop("coverage_version")
    first.path.write_text(json.dumps(graph), encoding="utf-8")
    backend = Backend(library); backend.controls["menu"]["actions"].append("set-text")
    second = Explorer(library, "demo", region, backend)
    assert not second.run()["complete"]
    assert any(gap["action"] == "set-text" for gap in second.graph["states"]["root"]["interaction_gaps"])
