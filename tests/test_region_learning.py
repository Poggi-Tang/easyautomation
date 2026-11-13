"""Native adapter tests mock all desktop and API boundaries."""
from copy import deepcopy
from types import SimpleNamespace

import pytest
from PIL import Image

from easy_uiauto import region_actions, region_learning, workflows
from easy_uiauto.mcp import knowledge, library_tools, scanner
from test_knowledge_library import library, add_control
from test_region_picker import picked


def test_basic_learning_never_replays_entry_steps_or_activates(library,tmp_path,monkeypatch):
    from easy_uiauto import library_actions
    from easy_uiauto.ctrl import Controller
    identifier=library.save_picked_region("demo",picked(tmp_path));entry=add_control(library)
    library.configure_region("demo",identifier,steps=[workflows.click_step(entry)],condition=dict(kind="exists",location=entry["location"]))
    region=library.regions("demo")[0];expected=dict(handle=123)
    def forbidden(*a,**kw):raise AssertionError("basic learning sent UI input")
    monkeypatch.setattr(Controller,"activate_window_handle",forbidden)
    monkeypatch.setattr(workflows,"execute",forbidden)
    monkeypatch.setattr(library_actions,"bind_window",lambda _:object())
    monkeypatch.setattr(library_actions,"resolve_record",lambda *a,**k:object())
    monkeypatch.setattr(region_actions,"target_window",lambda e,s:e)
    monkeypatch.setattr(region_actions,"condition_matches",lambda *a:True)
    assert region_actions.visible_region(library,"demo",region,expected)==expected
    monkeypatch.setattr(region_actions,"condition_matches",lambda *a:False)
    with pytest.raises(ValueError,match="手动打开"):region_actions.visible_region(library,"demo",region,expected)


@pytest.fixture
def backend(library,tmp_path,monkeypatch):
    library.save_picked_region("demo",picked(tmp_path));region=library.regions("demo")[0]
    monkeypatch.setattr(region_learning.configuration,"load_gui_configuration",lambda:dict(
        api_url="https://example.invalid/v1",api_key="test",api_model="test"))
    expected=dict(handle=123,process_id=456,class_name="Window",title="Test")
    return region_learning.NativeBackend(library,"demo",region,expected,tmp_path)


def test_new_state_archives_regions_and_controls_without_replacing_manual_map(backend,library,monkeypatch):
    image=Image.new("RGB",(400,300),"white")
    frame=dict(state_id="dialog",image=image,target=backend.expected,scope={"kind":"main"},issues=[])
    monkeypatch.setattr(backend,"stable_frame",lambda:frame)
    before=deepcopy(library.regions("demo"));entry=add_control(library);route=[workflows.click_step(entry)]
    calls=[]
    def scan(*args,**kw):
        calls.append(kw);layout=kw["manual_layout"]
        record={**entry,"id":"dialog-input","page_id":layout["page_id"],"region_id":layout["regions"][0]["id"]}
        knowledge.save_control(library.directory("demo"),record)
        return dict(observed_control_ids=[record["id"]],truncated=False)
    monkeypatch.setattr(scanner,"scan_window",scan)
    result=backend.observe(route=route)
    assert result["controls"][0]["id"]=="dialog-input" and result["region_ids"]
    assert library.regions("demo")==before
    region=next(r for r in library.read("demo")["regions"] if r["id"]==result["region_ids"][0])
    assert region["learning_source"]=="deep-learning" and region["entry_steps"][0]["target"]["control_id"]==entry["id"]
    assert calls[0]["observation_owner"]==backend.expected and calls[0]["strategy"]=="local-visual"
    assert backend.observe(route=[]) is result and len(calls)==1


def test_planner_missing_input_and_commit_policy_leave_explicit_blocks(backend,library,monkeypatch):
    entry=add_control(library)
    observation=dict(controls=[entry],image=Image.new("RGB",(10,10)))
    plan=dict(branches=[dict(name="输入",effect="input",steps=[{**workflows.click_step(entry,action="set-text"),"text":"${missing}"}]),
                        dict(name="提交",effect="commit",steps=[workflows.click_step(entry)])])
    monkeypatch.setattr(region_learning,"_request_plan",lambda *args:deepcopy(plan))
    result=backend.plan(observation,[])
    assert "缺少" in result["branches"][0]["blocked"] and "提交" in result["branches"][1]["blocked"]


@pytest.mark.parametrize("reason", ["planner", "missing-input", "commit"])
def test_blocked_api_steps_without_ids_save_progress_and_resume(backend,library,monkeypatch,reason):
    from easy_uiauto.deep_learning import Explorer
    entry=add_control(library)
    observation=dict(state_id="root",controls=[entry],image=Image.new("RGB",(10,10)))
    branch=dict(name="输入",effect="input",steps=[dict(action="set-text",target=dict(control_id=entry["id"]),
                text="${missing}" if reason=="missing-input" else "test",
                checks=[dict(kind="exists",target=dict(control_id=entry["id"]))])])
    if reason=="planner":branch["blocked"]="缺少测试资料"
    if reason=="commit":branch["effect"]="commit"
    raw=dict(branches=[branch])
    monkeypatch.setattr(region_learning,"_request_plan",lambda *args:deepcopy(raw))
    monkeypatch.setattr(backend,"observe",lambda **kwargs:observation)
    def forbidden(*args):raise AssertionError("blocked branch sent desktop input")
    monkeypatch.setattr(backend,"execute",forbidden)
    monkeypatch.setattr(backend,"assert_state",forbidden)
    for _ in range(2):
        explorer=Explorer(library,"demo",backend.region,backend)
        result=explorer.run()
        assert result["states"]==1 and result["actions"]==0 and not result["complete"]
        assert result["unfinished"][0]["status"]=="blocked"
        assert not explorer.graph["stop_reason"] and explorer.path.is_file()
        saved=explorer.graph["states"]["root"]["branches"][0]["steps"][0]
        assert saved["id"] and saved["checks"][0]["id"]
    assert "id" not in raw["branches"][0]["steps"][0]


@pytest.mark.parametrize("plan,message", [
    ([], "branches"),
    (dict(branches=[None]), "名称"),
    (dict(branches=[dict(name="输入",steps=None)]), "步骤必须为列表"),
    (dict(branches=[dict(name="输入",steps=[None])]), "步骤必须为对象"),
    (dict(branches=[dict(name="输入",steps=[{}])]), "action"),
    (dict(branches=[dict(name="输入",effect=[],steps=[])]), "effect"),
    (dict(branches=[dict(name="输入",steps=[dict(action="click",target=None)])]), "控件引用"),
    (dict(branches=[dict(name="输入",steps=[dict(action="click",checks=None)])]), "校验必须为列表"),
    (dict(branches=[],completion_checks=[dict(kind="exists")]), "target"),
    (dict(branches=[],completion_checks=[dict(target=dict(control_id="c"))]), "kind"),
    (dict(branches=[],return_steps=[{}]), "action"),
    (dict(branches=[],unresolved="未完成"), "unresolved"),
    (dict(branches=[],semantics=[None]), "语义说明"),
])
def test_malformed_provider_plan_reports_field_error(backend,library,monkeypatch,plan,message):
    entry=add_control(library)
    monkeypatch.setattr(region_learning,"_request_plan",lambda *args:deepcopy(plan))
    with pytest.raises(ValueError,match=message):
        backend.plan(dict(controls=[entry],image=Image.new("RGB",(10,10))),[])


def test_restore_does_not_replay_a_commit_route(backend,monkeypatch):
    backend.reset_workflow_id="reset"
    monkeypatch.setattr(backend,"stable_frame",lambda:dict(state_id="after"))
    with pytest.raises(ValueError,match="不能重复"):
        backend.restore(dict(state_id="before"),dict(controls=[]),[],[dict(learning_replayable=False)])


def test_restore_cannot_use_delete_as_return_navigation(backend,library,monkeypatch):
    record={**add_control(library),"name":"删除","semantic_name":"删除工程"}
    monkeypatch.setattr(backend,"stable_frame",lambda:dict(state_id="after"))
    def forbidden(*args):raise AssertionError("unexpected desktop input")
    monkeypatch.setattr(backend,"execute",forbidden)
    with pytest.raises(ValueError,match="删除"):
        backend.restore(dict(state_id="before"),dict(controls=[record]),[workflows.click_step(record)],[])


def test_state_fingerprint_accounts_for_input_values_and_rejects_failed_traversal(backend,monkeypatch):
    import uiautomation
    from easy_uiauto import library_actions
    from test_mcp_scanner import FakeControl
    field=FakeControl("Name","EditControl",(10,10,30,30))
    root=FakeControl("Test","WindowControl",(0,0,100,100),[field])
    value={"text":"first"};field.GetValuePattern=lambda:SimpleNamespace(Value=value["text"])
    monkeypatch.setattr(library_actions,"bind_window",lambda _:root)
    monkeypatch.setattr(library_actions,"owned_windows",lambda _:[])
    monkeypatch.setattr(uiautomation,"GetForegroundWindow",lambda:123)
    monkeypatch.setattr(uiautomation,"ControlFromHandle",lambda _:root)
    monkeypatch.setattr(scanner,"_window_rect",lambda _:dict(left=0,top=0,width=100,height=100))
    monkeypatch.setattr(scanner,"_screenshot_window",lambda _:Image.new("RGB",(100,100)))
    first=backend.frame()["state_id"];value["text"]="second"
    assert backend.frame()["state_id"]!=first
    def fail():raise RuntimeError("UIA provider failed")
    field.GetChildren=fail
    with pytest.raises(RuntimeError,match="UIA provider"):backend.frame()


def test_invalid_options_and_missing_api_fail_before_desktop(library,tmp_path,monkeypatch):
    for options in ({"max_depth":True},{"resume":"yes"},{"inputs":[]},{"unknown":1}):
        with pytest.raises(ValueError):region_learning.validate_options(options)
    monkeypatch.setattr(region_learning.configuration,"load_gui_configuration",lambda:{})
    with pytest.raises(ValueError,match="API"):
        region_learning.NativeBackend(library,"demo",{}, {},tmp_path)


def test_mcp_prepares_deep_job_with_same_app_and_explicit_options(library,tmp_path,monkeypatch):
    region=library.save_picked_region("demo",picked(tmp_path))
    monkeypatch.setattr(library_tools,"_window",lambda *a:dict(handle=123,title="Test"))
    options=dict(inputs={"工程名称":"test"},allow_commit=True,max_actions=25)
    result=library_tools._request("deep-learn-region","demo",vault=str(library.root),region_id=region,learning_options=options)
    assert result["operation"]=="deep-learn" and result["region_id"]==region
    assert result["learning_options"]["inputs"]==options["inputs"]
    assert result["app_id"]=="demo"


def test_update_reuses_region_and_node_identity_when_interface_changes(backend,library,monkeypatch):
    from easy_uiauto import library_actions, image_models
    record=add_control(library);record.update(region_id=backend.region["id"],semantic_source="manual",semantic_name="人工按钮")
    knowledge.save_control(library.directory("demo"),record)
    node=dict(id="old-state",fingerprint="old-state",name="主界面",route=[],region_ids=[backend.region["id"]],
              controls=[record["id"]],branches=[],issues=[],terminal=False,verified_terminal=False)
    backend.configure_learning(dict(root="old-state",states={"old-state":node}),"update",[],[])
    image=Image.new("RGB",(400,300),"white")
    frame=dict(state_id="updated-state",image=image,target=backend.expected,scope={"kind":"main"},issues=[])
    monkeypatch.setattr(backend,"stable_frame",lambda:deepcopy(frame))
    monkeypatch.setattr(image_models,"inspect_models",lambda **_:dict(ok=True))
    monkeypatch.setattr(library_actions,"activate_application",lambda _:None)
    monkeypatch.setattr(library_actions,"bind_window",lambda _:object())
    monkeypatch.setattr(library_actions,"resolve_record",lambda *a,**k:object())
    def forbidden(*a,**k):raise AssertionError("update replayed old region entry")
    monkeypatch.setattr(region_actions,"open_region",forbidden)
    layouts=[]
    def scan(*args,**kwargs):
        layouts.append(kwargs["manual_layout"])
        changed=deepcopy(record);changed["location"]["Name"]="new-location";changed["semantic_name"]="auto"
        knowledge.save_control(library.directory("demo"),changed)
        return dict(observed_control_ids=[record["id"]],truncated=False)
    monkeypatch.setattr(scanner,"scan_window",scan)
    observed=backend.observe(root=True)
    assert observed["state_id"]=="updated-state" and observed["node_id"]=="old-state"
    assert layouts[0]["page_id"]==backend.region["page_id"] and observed["region_ids"]==[backend.region["id"]]
    assert observed["controls"][0]["semantic_name"]=="人工按钮"
    assert observed["controls"][0]["location"]["Name"]=="new-location"


def test_update_verifies_affected_flow_with_real_assertion_evidence(backend,library,monkeypatch):
    record=add_control(library)
    step=workflows.click_step(record,action="inspect")
    step["checks"]=[dict(id="check",kind="exists",target=dict(control_id=record["id"]))]
    flow=library.save_workflow("demo",dict(id="verify",keyword="验证入口",entry_state="root",steps=[step]))
    backend.configure_learning(dict(states={},root=""),"update",[],[flow["id"]])
    backend.revision.summarize()
    monkeypatch.setattr(backend,"stable_frame",lambda:dict(state_id="new-root"))
    calls=[]
    def execute(flow,controls,*a,**kw):
        calls.append(flow["id"])
        return dict(ok=True,definition_hash=workflows.definition_hash(flow),control_hash=workflows.control_hash(flow,controls),
                    steps=[dict(ok=True,checks=[dict(ok=True)])])
    monkeypatch.setattr(workflows,"execute",execute)
    graph=dict(states={"root":dict(fingerprint="new-root",branches=[])})
    backend.validate_affected_workflows(graph,10)
    summary=backend.revision.summarize()
    assert calls==["verify"] and summary["workflows"]["verify"]["status"]=="passed"
    saved=workflows.load_catalog(library.directory("demo"))["verify"]
    assert saved["last_run"]["ok"] and saved["update_validation"]["status"]=="passed"


def test_update_validation_does_not_execute_unsupported_or_unapproved_commit(backend,library,monkeypatch):
    record=add_control(library)
    step={**workflows.click_step(record),"learning_effect":"commit"}
    step["checks"]=[dict(id="check",kind="exists",target=dict(control_id=record["id"]))]
    flow=library.save_workflow("demo",dict(id="save",keyword="保存",entry_state="root",steps=[step]))
    backend.configure_learning(dict(states={},root=""),"update",[],[flow["id"]]);backend.revision.summarize()
    def forbidden(*a,**kw):raise AssertionError("sent unapproved input")
    monkeypatch.setattr(workflows,"execute",forbidden)
    backend.validate_affected_workflows(dict(states={}),10)
    result=backend.revision.summarize()["workflows"]["save"]
    assert result["status"]=="pending" and "启用提交" in result["reason"]


def test_planner_improves_selected_semantics_and_keeps_manual_and_observed_effects(backend,library,monkeypatch):
    base=add_control(library);records=[]
    for identifier,extra in (("selected",{}),("other",{}),("manual",dict(semantic_source="manual")),
                             ("observed",dict(description_source="interaction-observation"))):
        record={**deepcopy(base),"id":identifier,"semantic_name":"旧名称","description":"旧说明",**extra}
        knowledge.save_control(library.directory("demo"),record);records.append(record)
    backend.configure_learning(dict(states={},root=""),"improve",["selected","manual","observed"],[])
    def plan(settings,payload,image):
        assert payload["learning_mode"]=="improve" and "selected" in payload["selected_control_ids"]
        return dict(branches=[],semantics=[dict(control_id=r["id"],name="改进名称",description="改进说明") for r in records])
    monkeypatch.setattr(region_learning,"_request_plan",plan)
    backend.plan(dict(controls=records,image=Image.new("RGB",(10,10))),[])
    current={r["id"]:r for r in library.read("demo")["controls"]}
    assert current["selected"]["description"]=="改进说明"
    assert all(current[i]["description"]=="旧说明" for i in ("other","manual","observed"))


def test_preflight_failure_is_recorded_as_no_input(backend,library,monkeypatch):
    def fail(*a,**kw):raise ValueError("invalid target before execution")
    monkeypatch.setattr(workflows,"execute",fail)
    with pytest.raises(ValueError) as caught:backend.execute([workflows.click_step(add_control(library))])
    assert caught.value.execution_result["input_attempted"] is False


def test_full_update_reuses_flows_and_verifies_changed_navigation(backend,library,monkeypatch):
    from easy_uiauto import image_models, library_actions
    from easy_uiauto.deep_learning import Explorer
    base=add_control(library);machine=dict(state="root",version=1);observed_pages=[];executed=[]
    monkeypatch.setattr(image_models,"inspect_models",lambda **_:dict(ok=True))
    monkeypatch.setattr(library_actions,"activate_application",lambda _:None)
    monkeypatch.setattr(library_actions,"bind_window",lambda _:object())
    monkeypatch.setattr(library_actions,"resolve_record",lambda *a,**k:object())
    monkeypatch.setattr(region_actions,"open_region",lambda *a,**k:None)
    def frame():
        return dict(state_id=f"v{machine['version']}-{machine['state']}",image=Image.new("RGB",(400,300)),
                    target=backend.expected,scope={"kind":"main"},issues=[])
    def scan(*args,**kwargs):
        layout=kwargs["manual_layout"];observed_pages.append((machine["version"],machine["state"],layout["page_id"]))
        ids={"root":["menu","toolbar"],"dialog":["fill"],"done":["result"]}[machine["state"]]
        for identifier in ids:
            record={**deepcopy(base),"id":identifier,"name":identifier,"semantic_name":identifier,
                    "page_id":layout["page_id"],"region_id":layout["regions"][0]["id"],"actions":["click"]}
            # Layout IDs may be local; persist the full region identity like the scanner.
            if not record["region_id"].startswith(layout["page_id"]+"."):
                record["region_id"]=layout["page_id"]+"."+record["region_id"]
            record["location"]["Name"]=f"v{machine['version']}-{identifier}"
            knowledge.save_control(library.directory("demo"),record)
        return dict(observed_control_ids=ids,truncated=False)
    monkeypatch.setattr(scanner,"scan_window",scan)
    def plan(settings,payload,image):
        ids=[r["id"] for r in payload["controls"] if r["id"] in {"menu","toolbar","fill","result"}]
        if machine["state"]=="done":
            return dict(name="完成",branches=[],terminal=True,completion_checks=[dict(kind="text",expected="完成",target=dict(control_id="result"))])
        return dict(name=machine["state"],branches=[dict(name=i,effect="navigation",replayable=True,
            steps=[dict(action="click",target=dict(control_id=i))]) for i in ids],unresolved=[])
    monkeypatch.setattr(region_learning,"_request_plan",plan)
    def configure(instance):
        monkeypatch.setattr(instance,"stable_frame",frame)
        def execute(steps):
            for step in steps:
                if step["action"]=="inspect":continue
                identifier=step["target"]["control_id"];executed.append((machine["version"],identifier))
                machine["state"]="done" if identifier=="fill" else "dialog"
            return dict(ok=True,input_attempted=True,steps=[dict(ok=True,checks=[dict(ok=True)])])
        monkeypatch.setattr(instance,"execute",execute)
        def restore(parent,*args):machine["state"]=parent["state_id"].split("-",1)[1]
        monkeypatch.setattr(instance,"restore",restore)
    configure(backend)
    first=Explorer(library,"demo",backend.region,backend)
    first.run();original=workflows.load_catalog(library.directory("demo"))
    machine.update(version=2,state="root")
    second_backend=region_learning.NativeBackend(library,"demo",backend.region,backend.expected,backend.folder)
    configure(second_backend)
    second=Explorer(library,"demo",backend.region,second_backend,mode="update")
    result=second.run();updated=workflows.load_catalog(library.directory("demo"))
    assert set(updated)==set(original) and second.graph["id"]==first.graph["id"]
    assert "menu" in result["changes"]["updated"] and result["changes"]["workflows"]
    assert all(r["status"]=="passed" for r in result["changes"]["workflows"].values())
    for state in ("dialog","done"):
        assert {p for _,s,p in observed_pages if s==state}.__len__()==1
    assert [i for v,i in executed if v==2]==["menu","fill","toolbar"]
