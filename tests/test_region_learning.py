"""Native adapter tests mock all desktop and API boundaries."""
from copy import deepcopy
from types import SimpleNamespace

import pytest
from PIL import Image, ImageDraw

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
    frame=dict(state_id="dialog",image=image,target={**backend.expected,"handle":789,"title":"Dialog"},
               scope=dict(kind="owned",class_name="Window",title="Dialog"),root=object(),issues=[])
    monkeypatch.setattr(backend,"stable_frame",lambda:frame)
    before=deepcopy(library.regions("demo"));entry=add_control(library);route=[workflows.click_step(entry)]
    frame["visible_windows"]={backend.surface_key(frame["target"])}
    backend.transition=dict(before_windows=set(),controls={entry["id"]},steps=route)
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
    assert "normalized_rect" not in region and library.asset("demo",region["image"]).is_file()
    assert region["parent_control_id"] == entry["id"]
    assert region["window_scope"] == frame["scope"]
    assert calls[0]["observation_owner"]==backend.expected and calls[0]["strategy"]=="local-visual"
    assert backend.observe(route=[]) is result and len(calls)==1


@pytest.fixture
def content_frame(backend, monkeypatch):
    from easy_uiauto import library_actions
    from test_mcp_scanner import FakeControl
    panel = FakeControl("工具栏", "PaneControl", (50, 60, 450, 180), automation_id="101")
    panel.ClassName = "NativePanel"
    root = FakeControl("Test", "WindowControl", (50, 60, 450, 360), [panel])
    monkeypatch.setattr(library_actions, "resolve_record", lambda *a, **kw: panel)
    return dict(state_id="new-content", image=Image.new("RGB", (400, 300)), target=backend.expected,
                scope={"kind":"main"}, bounds=scanner._rect(root), root=root, issues=[])


@pytest.mark.parametrize("drawn", [False, True])
def test_changed_content_retains_manual_region_and_control_entry_routes(backend, library, monkeypatch, content_frame, drawn):
    from easy_uiauto.mcp import automatic_layout
    if drawn:
        backend.region.pop("location")
        reference = backend.library.asset(backend.app_id, backend.region["image"])
        draw = ImageDraw.Draw(content_frame["image"])
        draw.rectangle((35, 22, 126, 45), fill="red")
        draw.polygon([(40, 28), (70, 76), (110, 30)], fill="white")
        content_frame["image"].crop((10, 10, 200, 100)).save(reference)
        knowledge.save_region(backend.directory, backend.region)
    before = deepcopy(library.read("demo"))
    entry = add_control(library)
    entry.update(region_id=backend.region["id"], entry_steps=[], semantic_source="manual")
    knowledge.save_control(backend.directory, entry)
    route = [workflows.click_step(entry)]
    monkeypatch.setattr(backend, "stable_frame", lambda: content_frame)
    monkeypatch.setattr(automatic_layout, "propose", lambda *a: pytest.fail("repartitioned manual content area"))
    scans = []
    def scan(*args, **kwargs):
        layout = kwargs["manual_layout"]; scans.append(layout)
        assert layout["restrict_regions"] and layout["page_id"] == backend.region["page_id"]
        assert [r["id"] for r in layout["regions"]] == [backend.region["id"].split(".", 1)[1]]
        assert kwargs["navigation_context"] is None
        new = {**entry, "id":"diagram", "page_id":layout["page_id"], "region_id":layout["regions"][0]["id"]}
        new.pop("entry_steps")
        # Simulate the scanner's preservation of saved prerequisites on rescans.
        old = knowledge.find_control_record(backend.directory, new["id"])
        knowledge.save_control(backend.directory, knowledge.merge_control_state(old[1] if old else {}, new))
        return dict(observed_control_ids=[entry["id"], new["id"]], truncated=False)
    monkeypatch.setattr(scanner, "scan_window", scan)
    first = backend.observe(route=route)
    content_frame["state_id"] = "another-content-state"
    second = backend.observe(route=[workflows.click_step(entry, action="double-click")])
    assert first["node_id"] != second["node_id"]
    assert first["region_ids"] == second["region_ids"] == [backend.region["id"]]
    snapshot = library.read("demo")
    assert snapshot["regions"] == before["regions"] and snapshot["pages"] == before["pages"]
    assert len(snapshot["controls"]) == 2 and len(scans) == 2
    records = {r["id"]:r for r in snapshot["controls"]}
    assert records[entry["id"]]["entry_steps"] == []
    steps = library.entry_steps("demo", records["diagram"])
    assert len(steps) == 1 and steps[0]["action"] == "click"
    assert steps[0]["target"]["control_id"] == entry["id"]


def test_resume_old_state_uses_valid_manual_region_without_deleting_old_page(backend, library, monkeypatch, content_frame):
    old = {**backend.region, "id":"deep-old.region", "page_id":"deep-old", "learning_source":"deep-learning"}
    monkeypatch.setattr(knowledge, "utc_now", lambda: "2026-01-01T00:00:00+00:00")
    path = knowledge.save_region(backend.directory, old)
    old = knowledge.read_markdown(path)[0]
    saved = path.read_bytes()
    monkeypatch.setattr(knowledge, "utc_now", lambda: "2026-01-01T00:00:01+00:00")
    backend.previous_graph = dict(states={"old-node":dict(id="old-node", fingerprint=content_frame["state_id"],
                                                        region_ids=[old["id"]], route=[])})
    monkeypatch.setattr(backend, "stable_frame", lambda:content_frame)
    monkeypatch.setattr(scanner, "scan_window", lambda *a, **kw:dict(observed_control_ids=[], truncated=False))
    result = backend.observe()
    assert result["node_id"] == "old-node" and result["region_ids"] == [backend.region["id"]]
    assert next(r for r in library.read("demo")["regions"] if r["id"] == old["id"]) == old
    assert path.read_bytes() == saved


@pytest.mark.parametrize("change", ["condition", "missing", "outside", "replacement", "owned", "other-window"])
def test_different_layout_does_not_reuse_manual_region(backend, library, monkeypatch, content_frame, change):
    from easy_uiauto import library_actions
    from easy_uiauto.mcp.results import ActionError
    if change == "condition":
        backend.region["state_condition"] = {"kind":"selected"}
        monkeypatch.setattr(region_actions, "condition_matches", lambda *a:False)
    elif change == "missing":
        def missing(*a, **kw):raise ActionError("control_not_found", "panel absent")
        monkeypatch.setattr(library_actions, "resolve_record", missing)
    elif change == "outside":
        content_frame["root"]._children[0].BoundingRectangle.bottom = 420
    elif change == "replacement":
        content_frame["root"]._children[0].AutomationId = "different-panel"
    elif change == "owned":
        content_frame["scope"] = {"kind":"owned"}
    else:
        content_frame["target"] = {**backend.expected, "handle":999}
    assert backend.existing_region_layout(content_frame) is None


@pytest.mark.parametrize("code,reason", [("ambiguous_match", ""), ("control_not_found", "uia_access_denied")])
def test_region_probe_errors_stop_before_scanning(backend, monkeypatch, content_frame, code, reason):
    from easy_uiauto import library_actions
    from easy_uiauto.mcp.results import ActionError
    def fail(*a, **kw):raise ActionError(code, "cannot read panel", reason=reason)
    monkeypatch.setattr(library_actions, "resolve_record", fail)
    monkeypatch.setattr(backend, "stable_frame", lambda:content_frame)
    monkeypatch.setattr(scanner, "scan_window", lambda *a, **kw:pytest.fail("scanned after failed probe"))
    with pytest.raises(ActionError, match=code):backend.observe()


def test_resized_main_rebinds_region_without_writing_a_new_page(backend, library, monkeypatch, content_frame):
    before = library.read("demo")
    content_frame["bounds"].update(width=500,right=550)
    content_frame["image"]=Image.new("RGB",(500,300))
    content_frame["root"]._children[0].BoundingRectangle.right=550
    monkeypatch.setattr(backend, "stable_frame", lambda:content_frame)
    layouts=[]
    def scan(*a,**kw):
        layouts.append(kw["manual_layout"])
        return dict(observed_control_ids=[],truncated=False)
    monkeypatch.setattr(scanner,"scan_window",scan)
    observed=backend.observe()
    assert observed["region_ids"]==[backend.region["id"]]
    assert len(layouts)==1 and layouts[0]["image_size"]==dict(width=500,height=300)
    assert layouts[0]["restrict_regions"] and len(layouts[0]["regions"])==1
    assert library.read("demo") == before


def test_same_window_size_uses_moved_native_region_without_saving_coordinates(
    backend, library, monkeypatch, content_frame,
):
    panel = content_frame["root"]._children[0]
    backend.region.pop("normalized_rect")
    knowledge.save_region(backend.directory, backend.region)
    original = deepcopy(library.regions("demo"))
    layouts = []
    monkeypatch.setattr(backend, "stable_frame", lambda: content_frame)

    def scan(*_args, **kwargs):
        layouts.append(kwargs["manual_layout"])
        return dict(observed_control_ids=[], truncated=False)

    monkeypatch.setattr(scanner, "scan_window", scan)
    first = backend.observe()
    panel.BoundingRectangle.bottom = 240
    second = backend.observe()
    assert first["state_id"] == second["state_id"]
    assert second["region_ids"] == [backend.region["id"]]
    assert [layout["regions"][0]["normalized_rect"]["bottom"] for layout in layouts] == [.4, .6]
    assert library.regions("demo") == original


@pytest.mark.parametrize("duplicate", [False, True])
def test_live_region_image_fallback_requires_a_unique_match(
    backend, content_frame, monkeypatch, duplicate,
):
    from easy_uiauto import library_actions
    from easy_uiauto.mcp.results import ActionError

    def missing(*_args, **_kwargs):
        raise ActionError("control_not_found", "UIA region absent")

    monkeypatch.setattr(library_actions, "resolve_record", missing)
    template = Image.new("RGB", (40, 24), "white")
    draw = ImageDraw.Draw(template)
    draw.polygon([(2, 3), (33, 20), (15, 22)], fill="black")
    draw.rectangle((24, 3, 29, 12), fill="red")
    ref = backend.library.asset(backend.app_id, backend.region["image"])
    template.save(ref)
    content_frame["image"].paste(template, (80, 50))
    if duplicate:
        content_frame["image"].paste(template, (210, 130))
    layout = backend.existing_region_layout(content_frame)
    if duplicate:
        assert layout is None
    else:
        assert layout["geometry_adapted"] == "image-location"
        assert layout["regions"][0]["normalized_rect"]["left"] == .2
        assert layout["regions"][0]["normalized_rect"]["right"] == .3


def test_duplicate_native_region_identity_is_not_selected_by_old_ordinal(backend, content_frame):
    from easy_uiauto.mcp.results import ActionError
    from test_mcp_scanner import FakeControl

    duplicate = FakeControl("另一个标题", "PaneControl", (50, 200, 450, 300), automation_id="101")
    duplicate.ClassName = "NativePanel"
    content_frame["root"]._children.append(duplicate)
    with pytest.raises(ActionError, match="多个"):
        backend.existing_region_layout(content_frame)


def test_unidentified_area_stops_without_repartitioning_or_writing(backend, library, monkeypatch, content_frame):
    before = deepcopy(library.regions("demo")); page = deepcopy(library.page("demo"))
    content_frame["root"]._children[0].AutomationId = "unknown"
    backend.previous_graph = dict(states={"old-node":dict(id="old-node", fingerprint=content_frame["state_id"],
                                                        region_ids=[backend.region["id"]], route=[])})
    monkeypatch.setattr(backend, "stable_frame", lambda:content_frame)
    monkeypatch.setattr(scanner, "scan_window", lambda *a, **kw:pytest.fail("scanned outside selected region"))
    snapshot=deepcopy(library.read("demo"))
    with pytest.raises(ValueError,match="未执行整窗"):
        backend.observe()
    assert library.read("demo")==snapshot
    assert library.regions("demo") == before and library.page("demo") == page


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


def test_mcp_single_control_infers_region_and_rejects_wrong_scope(library,tmp_path,monkeypatch):
    region=library.save_picked_region("demo",picked(tmp_path))
    control=library.save_captured_control("demo",picked(tmp_path),name="新建",region_id=region)
    monkeypatch.setattr(library_tools,"_window",lambda *args:dict(handle=123,title="Test"))
    result=library_tools._request("deep-learn-control","demo",vault=str(library.root),control_id=control["id"],
                                  learning_options=dict(mode="new",max_actions=20))
    assert result["operation"]=="deep-learn" and result["region_id"]==region
    assert result["learning_options"]["entry_control_id"]==control["id"]
    assert result["learning_options"]["mode"]=="new"
    for options in (dict(control_id="missing"),dict(control_id=control["id"],region_id="wrong"),
                    dict(control_id=control["id"],learning_options={"entry_control_id":"other"})):
        with pytest.raises(ValueError):library_tools._request("deep-learn-control","demo",vault=str(library.root),**options)


def test_worker_rejects_control_from_other_region_before_constructing_native_backend(library,tmp_path,monkeypatch):
    region=library.save_picked_region("demo",picked(tmp_path))
    record=add_control(library)
    monkeypatch.setattr(region_learning,"NativeBackend",lambda *a,**k:pytest.fail("constructed backend for wrong region"))
    with pytest.raises(ValueError,match="所属区域"):
        region_learning.learn_region(library,"demo",region,{},tmp_path,entry_control_id=record["id"])
    with pytest.raises(ValueError,match="文本"):
        region_learning.validate_options(dict(entry_control_id=1))


def test_planner_learns_child_semantics_without_changing_unrelated_root_entry(backend,library,monkeypatch):
    base=add_control(library);records=[]
    for identifier in ("entry","child","unrelated"):
        record={**deepcopy(base),"id":identifier,"semantic_name":"旧名称","description":"旧说明"}
        knowledge.save_control(library.directory("demo"),record);records.append(record)
    backend.configure_learning(dict(states={},root=""),"improve",["entry"],[])
    def plan(settings,payload,image):
        assert payload["selected_control_ids"]==[] and payload["exploration_scope"]["followup"]
        return dict(branches=[],semantics=[dict(control_id=r["id"],description="后续说明") for r in records])
    monkeypatch.setattr(region_learning,"_request_plan",plan)
    backend.plan(dict(controls=records,image=Image.new("RGB",(10,10)),exploration_scope=dict(
        followup=True,selected_control_ids=["entry"],excluded_entry_control_ids=["unrelated"])),[])
    current={r["id"]:r for r in library.read("demo")["controls"]}
    assert current["child"]["description"]=="后续说明" and current["unrelated"]["description"]=="旧说明"


@pytest.mark.parametrize("resolved", [True, False])
def test_manual_entry_omitted_by_scan_requires_fresh_visible_locator(backend,library,monkeypatch,resolved):
    from easy_uiauto import library_actions,image_models
    record=add_control(library);record.update(region_id=backend.region["id"],semantic_source="manual")
    knowledge.save_control(library.directory("demo"),record)
    sibling={**deepcopy(record),"id":"unrelated-sibling"}
    descendant={**deepcopy(record),"id":"known-followup","entry_steps":[workflows.click_step(record)]}
    knowledge.save_control(backend.directory,sibling);knowledge.save_control(backend.directory,descendant)
    backend.configure_learning(dict(states={},root="",entry_control_id=record["id"]),"new",[record["id"]],[])
    opened=[]
    def open_entry(lib,app,entry,*a,**k):
        assert not entry.get("location")  # Entry verification below also supports image locators.
        opened.append(entry["id"])
    monkeypatch.setattr(region_actions,"open_region",open_entry)
    monkeypatch.setattr(image_models,"inspect_models",lambda **k:dict(ok=True))
    monkeypatch.setattr(library_actions,"bind_window",lambda *a:object())
    def resolve(*a,**k):
        if not resolved:raise ValueError("stale locator")
        return SimpleNamespace(IsEnabled=True)
    monkeypatch.setattr(library_actions,"resolve_record",resolve)
    monkeypatch.setattr(scanner,"_window_rect",lambda _:dict(left=10,top=10,width=20,height=20))
    monkeypatch.setattr(scanner,"scan_window",lambda *a,**k:pytest.fail("single-entry origin must not rescan the region"))
    frame=dict(state_id="root",image=Image.new("RGB",(400,300)),target=backend.expected,scope={"kind":"main"},
               bounds=dict(left=0,top=0,width=400,height=300),root=object(),issues=[])
    monkeypatch.setattr(backend,"stable_frame",lambda:deepcopy(frame))
    observed=backend.observe(root=True)
    assert observed["excluded_entry_control_ids"]==[sibling["id"]]
    assert opened==[record["id"]]
    assert (record["id"] in [r["id"] for r in observed["controls"]])==resolved
    if not resolved:assert any("指定入口尚未定位" in reason for reason in observed["issues"])


def test_scoped_post_learning_validation_does_not_replay_unrelated_workflow(backend,library,monkeypatch):
    record=add_control(library)
    step=workflows.click_step(record,action="inspect")
    step["checks"]=[dict(id="check",kind="exists",target=dict(control_id=record["id"]))]
    flow=library.save_workflow("demo",dict(id="unrelated",keyword="其他入口流程",entry_state="root",steps=[step]))
    backend.configure_learning(dict(states={},root=""),"update",[],[flow["id"]]);backend.revision.summarize()
    monkeypatch.setattr(workflows,"execute",lambda *a,**k:pytest.fail("executed unrelated workflow"))
    backend.validate_affected_workflows(dict(entry_control_id=record["id"],run_workflow_ids=[]),10)
    report=backend.revision.summarize()["workflows"][flow["id"]]
    assert report["status"]=="pending" and "本次控件学习" in report["reason"]


def test_update_reuses_region_and_node_identity_when_interface_changes(backend,library,monkeypatch,content_frame):
    from easy_uiauto import library_actions, image_models
    record=add_control(library);record.update(region_id=backend.region["id"],semantic_source="manual",semantic_name="人工按钮")
    knowledge.save_control(library.directory("demo"),record)
    node=dict(id="old-state",fingerprint="old-state",name="主界面",route=[],region_ids=[backend.region["id"]],
              controls=[record["id"]],branches=[],issues=[],terminal=False,verified_terminal=False)
    backend.configure_learning(dict(root="old-state",states={"old-state":node}),"update",[],[])
    frame={**content_frame,"state_id":"updated-state"}
    monkeypatch.setattr(backend,"stable_frame",lambda:deepcopy(frame))
    monkeypatch.setattr(image_models,"inspect_models",lambda **_:dict(ok=True))
    monkeypatch.setattr(library_actions,"activate_application",lambda _:None)
    monkeypatch.setattr(library_actions,"bind_window",lambda _:object())
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
    for identifier,extra in (("selected",{}),("english",{}),("other",{}),
                             ("manual",dict(semantic_source="manual")),
                             ("observed",dict(description_source="interaction-observation"))):
        record={**deepcopy(base),"id":identifier,"semantic_name":"旧名称","description":"旧说明",**extra}
        knowledge.save_control(library.directory("demo"),record);records.append(record)
    backend.configure_learning(dict(states={},root=""),"improve",["selected","english","manual","observed"],[])
    def plan(settings,payload,image):
        assert payload["learning_mode"]=="improve" and "selected" in payload["selected_control_ids"]
        return dict(branches=[],semantics=[dict(control_id=r["id"],name="改进名称",
            description="English only." if r["id"]=="english" else "改进说明") for r in records])
    monkeypatch.setattr(region_learning,"_request_plan",plan)
    backend.plan(dict(controls=records,image=Image.new("RGB",(10,10))),[])
    current={r["id"]:r for r in library.read("demo")["controls"]}
    assert current["selected"]["description"]=="改进说明"
    assert all(current[i]["description"]=="旧说明" for i in ("english","other","manual","observed"))


def test_preflight_failure_is_recorded_as_no_input(backend,library,monkeypatch):
    def fail(*a,**kw):raise ValueError("invalid target before execution")
    monkeypatch.setattr(workflows,"execute",fail)
    with pytest.raises(ValueError) as caught:backend.execute([workflows.click_step(add_control(library))])
    assert caught.value.execution_result["input_attempted"] is False


def test_full_update_reuses_flows_and_verifies_changed_navigation(backend,library,monkeypatch,content_frame):
    from easy_uiauto import image_models, library_actions
    from easy_uiauto.deep_learning import Explorer
    base=add_control(library);machine=dict(state="root",version=1);observed_pages=[];executed=[]
    monkeypatch.setattr(image_models,"inspect_models",lambda **_:dict(ok=True))
    monkeypatch.setattr(library_actions,"activate_application",lambda _:None)
    monkeypatch.setattr(library_actions,"bind_window",lambda _:object())
    monkeypatch.setattr(region_actions,"open_region",lambda *a,**k:None)
    def frame():
        return dict(state_id=f"v{machine['version']}-{machine['state']}",image=Image.new("RGB",(400,300)),
                    target=backend.expected if machine["state"]=="root" else {**backend.expected,"handle":789,"title":"Dialog"},
                    scope={"kind":"main"} if machine["state"]=="root" else dict(kind="owned",title="Dialog",class_name="Window"),
                    root=content_frame["root"] if machine["state"]=="root" else object(),
                    bounds=content_frame["bounds"],visible_windows={(789,456,"Window")} if machine["state"]!="root" else set(),issues=[])
    def scan(*args,**kwargs):
        layout=kwargs["manual_layout"];observed_pages.append((machine["version"],machine["state"],layout["page_id"]))
        ids={"root":["menu","toolbar"],"dialog":["fill"],"done":["result"]}[machine["state"]]
        for identifier in ids:
            record={**deepcopy(base),"id":identifier,"name":identifier,"semantic_name":identifier,
                    "page_id":layout["page_id"],"region_id":layout["regions"][0]["id"],"actions":[] if identifier=="result" else ["click"]}
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
            instance.transition=dict(before_windows=frame()["visible_windows"],controls=set(instance.observed_controls),steps=deepcopy(steps))
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



def test_planner_receives_observed_search_terms_and_action_feedback(backend, library, monkeypatch):
    record = add_control(library)
    record.update(observed_value="电阻", class_name="SearchLineEdit", current_state="enabled", help_text="搜索元件")
    feedback = dict(missing_interactions=[dict(control_id=record["id"], action="set-text")])
    returned = dict(input_values={record["id"]:"旧查询"})
    def request(settings, payload, image):
        assert payload["test_inputs"] == {}
        assert payload["controls"][0]["observed_value"] == "电阻"
        assert payload["controls"][0]["class_name"] == "SearchLineEdit"
        assert payload["planning_feedback"] == feedback
        assert payload["return_context"] == returned
        return dict(branches=[])
    monkeypatch.setattr(region_learning, "_request_plan", request)
    backend.plan(dict(controls=[record], image=Image.new("RGB", (10, 10)), planning_feedback=feedback, return_context=returned), [])


def test_deep_execution_evidence_preserves_properties_and_images_without_geometry():
    evidence = {"steps": [{"input_attempted": True, "resolved_by": "image", "rect": {"left": 1},
                           "before": {"context_image": "scene.png", "control_image": "control.png",
                                      "context_size": {"width": 100}, "normalized_rect": {"left": .1}},
                           "checks": [{"text": "验证通过", "rect": {"left": 2}}]}],
                "failure_observation": {"main_rect": {"top": 0}, "context_image": "result.png"}}
    proof = region_learning._execution_evidence(evidence)
    assert proof["steps"] == [{"input_attempted": True, "resolved_by": "image",
                              "before": {"context_image": "scene.png", "control_image": "control.png"},
                              "checks": [{"text": "验证通过"}]}]
    assert proof["failure_observation"] == {"context_image": "result.png"}
    assert evidence["steps"][0]["rect"] == {"left": 1}


def test_input_capture_uses_live_value_not_placeholder_and_preserves_first_value(backend, library, monkeypatch):
    from easy_uiauto import library_actions
    record = add_control(library); record["observed_value"] = "请输入关键词"
    value = SimpleNamespace(Value="原始查询", IsReadOnly=False)
    control = SimpleNamespace(IsPassword=False, GetValuePattern=lambda:value)
    monkeypatch.setattr(library_actions, "owned_windows", lambda *_:[])
    monkeypatch.setattr(library, "capture_before_click", lambda *args:dict(ok=True))
    def execute(flow, *args, **kwargs):
        kwargs["before_action"](record, control, object())
        value.Value = "第一次搜索"
        kwargs["before_action"](record, control, object())
        return dict(ok=True, input_attempted=True)
    monkeypatch.setattr(workflows, "execute", execute)
    steps = [dict(action="set-text", target=dict(control_id=record["id"]), text="电阻", learning_effect="input")]
    assert backend.execute(steps)["input_values"] == {record["id"]:"原始查询"}


@pytest.mark.parametrize("password,pattern", [(True, None), (False, None), (False, SimpleNamespace(IsReadOnly=True))])
def test_unsafe_or_unrecoverable_input_stops_before_writing(backend, library, monkeypatch, password, pattern):
    record = add_control(library)
    control = SimpleNamespace(IsPassword=password, GetValuePattern=lambda:pattern)
    def execute(flow, *args, **kwargs):
        kwargs["before_action"](record, control, object())
        pytest.fail("input was allowed")
    monkeypatch.setattr(workflows, "execute", execute)
    with pytest.raises(ValueError) as error:
        backend.execute([dict(action="set-text", target=dict(control_id=record["id"]), text="x", learning_effect="input")])
    assert error.value.execution_result["input_attempted"] is False


@pytest.mark.parametrize("original", ["", "原查询"])
def test_restoration_writes_exact_original_and_verifies_parent(backend, library, monkeypatch, original):
    record = add_control(library)
    parent = dict(state_id="parent"); child = dict(state_id="child", controls=[record])
    machine = dict(state="child"); executed = []
    monkeypatch.setattr(backend, "stable_frame", lambda:dict(state_id=machine["state"]))
    def execute(steps, **kwargs):
        assert kwargs["literal_text"]
        executed.extend(steps); machine["state"] = "parent"
        return dict(ok=True)
    monkeypatch.setattr(backend, "execute", execute)
    backend.restore(parent, child, [], [], input_values={record["id"]:original})
    assert len(executed) == 1 and executed[0]["action"] == "set-text" and executed[0]["text"] == original


def test_restoration_rejects_invented_original_value(backend, library, monkeypatch):
    record = add_control(library)
    monkeypatch.setattr(backend, "stable_frame", lambda:dict(state_id="child"))
    monkeypatch.setattr(backend, "execute", lambda *_:pytest.fail("invented restoration executed"))
    with pytest.raises(ValueError, match="实际读取"):
        backend.restore(dict(state_id="parent"), dict(state_id="child", controls=[record]),
            [dict(action="set-text", target=dict(control_id=record["id"]), text="")], [], input_values={record["id"]:"原查询"})


def test_restoration_failure_is_not_accepted_after_successful_write(backend, library, monkeypatch):
    record = add_control(library)
    monkeypatch.setattr(backend, "stable_frame", lambda:dict(state_id="child"))
    monkeypatch.setattr(backend, "execute", lambda *_, **kwargs:dict(ok=True))
    with pytest.raises(ValueError, match="无法确认已返回"):
        backend.restore(dict(state_id="parent"), dict(state_id="child", controls=[record]), [], [], input_values={record["id"]:""})



def test_verified_parent_restores_transition_control_scope(backend, monkeypatch):
    backend.observed_controls = {"filtered-item"}
    backend.excluded_entry_ids = {"unrelated"}
    parent = dict(state_id="parent", controls=[dict(id="search"), dict(id="original-item"), dict(id="unrelated")])
    monkeypatch.setattr(backend, "stable_frame", lambda:parent)
    backend.assert_state(parent)
    assert backend.observed_controls == {"search", "original-item"}



def test_original_search_text_is_not_interpreted_as_a_parameter(backend, library, monkeypatch):
    record = add_control(library)
    original = "${business_value}"
    backend.inputs = dict(business_value="must-not-replace-original")
    def execute(flow, *args, **kwargs):
        expanded = workflows.substitute(flow["steps"], flow["parameters"])
        assert expanded[0]["text"] == original
        return dict(ok=True, input_attempted=True)
    monkeypatch.setattr(workflows, "execute", execute)
    result = backend.execute([dict(action="set-text", target=dict(control_id=record["id"]), text=original)], literal_text=True)
    assert result["ok"]
