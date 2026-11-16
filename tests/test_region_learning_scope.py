"""Learning scope must be proved before any new region is archived."""
from copy import deepcopy
from types import SimpleNamespace

import pytest
from PIL import Image

from easy_uiauto import library_actions, region_learning, workflows
from easy_uiauto.mcp import knowledge, scanner
from test_knowledge_library import library, add_control
from test_region_learning import backend, content_frame


def dialog_frame(backend):
    target={**backend.expected,"handle":789,"title":"Project"}
    return dict(state_id="dialog",image=Image.new("RGB",(400,300),"white"),target=target,
                scope=dict(kind="owned",class_name="Window",title="Project"),
                root=SimpleNamespace(Name="Project",AutomationId="project-dialog",ClassName="Dialog",ControlTypeName="WindowControl"),
                visible_windows={backend.surface_key(target)},issues=[])


@pytest.mark.parametrize("entry_only",[False,True])
@pytest.mark.parametrize("change",["missing","moved","condition","replacement"])
@pytest.mark.parametrize("owned",[False,True])
def test_scope_failure_never_creates_unrelated_regions(backend,library,monkeypatch,content_frame,entry_only,change,owned):
    from easy_uiauto import region_actions
    from easy_uiauto.mcp.results import ActionError
    entry=add_control(library)
    if owned:
        scope=dict(kind="owned",class_name="Window",title="Project")
        backend.region["window_scope"]=scope
        content_frame.update(scope=scope,target={**backend.expected,"handle":789,"title":"Project"})
    if entry_only:backend.previous_graph["entry_control_id"]=entry["id"]
    if change=="missing":
        def missing(*a,**kw):raise ActionError("control_not_found","panel gone")
        monkeypatch.setattr(library_actions,"resolve_record",missing)
    elif change=="moved":content_frame["root"]._children[0].BoundingRectangle.bottom+=20
    elif change=="replacement":content_frame["root"]._children[0].AutomationId="other-panel"
    else:
        backend.region["state_condition"]={"kind":"selected"}
        monkeypatch.setattr(region_actions,"condition_matches",lambda *a:False)
    before=deepcopy(library.read("demo"))
    monkeypatch.setattr(backend,"stable_frame",lambda:content_frame)
    monkeypatch.setattr(scanner,"scan_window",lambda *a,**kw:pytest.fail("out-of-scope scan"))
    with pytest.raises(ValueError,match="未执行整窗"):
        backend.observe(route=[workflows.click_step(entry)])
    assert library.read("demo")==before


@pytest.mark.parametrize("evidence",["none","already-open","outside-control","wrong-route","not-owned"])
def test_unproven_popup_does_not_write_any_records(backend,library,monkeypatch,evidence):
    entry=add_control(library);route=[workflows.click_step(entry)];frame=dialog_frame(backend)
    key=backend.surface_key(frame["target"])
    if evidence!="none":
        backend.transition=dict(before_windows={key} if evidence=="already-open" else set(),
                                controls=set() if evidence=="outside-control" else {entry["id"]},steps=route)
    if evidence=="wrong-route":route=[workflows.click_step(entry,action="right-click")]
    if evidence=="not-owned":frame["visible_windows"]=set()
    before=deepcopy(library.read("demo"))
    monkeypatch.setattr(backend,"stable_frame",lambda:frame)
    monkeypatch.setattr(scanner,"scan_window",lambda *a,**kw:pytest.fail("unrelated popup scan"))
    with pytest.raises(ValueError,match="未确认由本次"):
        backend.observe(route=route)
    assert library.read("demo")==before


def test_action_evidence_admits_new_dialog_and_reuses_its_regions(backend,library,monkeypatch):
    entry=add_control(library);route=[workflows.click_step(entry)];frame=dialog_frame(backend)
    backend.observed_controls={entry["id"]}
    monkeypatch.setattr(library_actions,"owned_windows",lambda _:[])
    monkeypatch.setattr(library,"capture_before_click",lambda *a:dict(ok=True))
    def execute(flow,*a,**kw):
        kw["before_action"](entry,object(),object())
        return dict(ok=True,input_attempted=True)
    monkeypatch.setattr(workflows,"execute",execute)
    backend.execute(route)
    monkeypatch.setattr(backend,"stable_frame",lambda:frame)
    layouts=[]
    def scan(*a,**kw):
        layout=kw["manual_layout"];layouts.append(deepcopy(layout))
        assert layout["window_handle"]==789 and layout["restrict_regions"]
        return dict(observed_control_ids=[],truncated=False)
    monkeypatch.setattr(scanner,"scan_window",scan)
    first=backend.observe(route=route);before=library.read("demo")
    frame["state_id"]="dialog-filled";frame["image"].putpixel((20,20),(0,0,0))
    second=backend.observe(route=route)
    after=library.read("demo")
    assert first["region_ids"]==second["region_ids"]
    assert len(after["regions"])==len(before["regions"]) and len(after["pages"])==len(before["pages"])
    assert layouts[0]["page_id"]==layouts[1]["page_id"]
    regions=[r for r in after["regions"] if r["id"] in first["region_ids"]]
    assert all(r["parent_control_id"]==entry["id"] and r["learning_origin_region"]==backend.region["id"] for r in regions)
    # A later run can resume this proved dialog without having just clicked it.
    fresh=region_learning.NativeBackend(library,"demo",backend.region,backend.expected,backend.folder)
    fresh.check_surface(frame,dict(region_ids=second["region_ids"],route=route),True,())


def test_existing_manual_region_inside_dialog_keeps_its_scope(backend,library,monkeypatch,content_frame):
    scope=dict(kind="owned",class_name="Window",title="Project")
    backend.region["window_scope"]=scope
    content_frame.update(scope=scope,target={**backend.expected,"handle":789,"title":"Project"})
    knowledge.save_region(backend.directory,backend.region)
    before=deepcopy(library.read("demo"))
    monkeypatch.setattr(backend,"stable_frame",lambda:content_frame)
    def scan(*a,**kw):
        layout=kw["manual_layout"]
        assert layout["restrict_regions"] and len(layout["regions"])==1
        return dict(observed_control_ids=[],truncated=False)
    monkeypatch.setattr(scanner,"scan_window",scan)
    observed=backend.observe()
    assert observed["region_ids"]==[backend.region["id"]]
    assert library.read("demo")==before


def test_frame_ignores_unrelated_tool_window_when_main_is_active(backend,monkeypatch):
    import uiautomation
    from test_mcp_scanner import FakeControl
    main=FakeControl("Test","WindowControl",(0,0,400,300))
    tool={**backend.expected,"handle":789,"title":"Unrelated tool"}
    monkeypatch.setattr(library_actions,"bind_window",lambda _:main)
    monkeypatch.setattr(library_actions,"owned_windows",lambda _:[tool])
    monkeypatch.setattr(uiautomation,"GetForegroundWindow",lambda:123)
    def bind(handle):
        assert handle==123,"selected an unrelated owned window"
        return main
    monkeypatch.setattr(uiautomation,"ControlFromHandle",bind)
    monkeypatch.setattr(scanner,"_screenshot_window",lambda _:Image.new("RGB",(400,300)))
    assert backend.frame()["scope"]=={"kind":"main"}
