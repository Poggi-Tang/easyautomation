from pathlib import Path

import pytest
from PIL import Image

from easy_uiauto import local_image_worker, utils
from easy_uiauto.mcp import knowledge, scanner, control_observation
from test_manual_region_scan import setup
from test_mcp_scanner import FakeControl


def element(identifier, kind, rect, text=""):
    left,top,right,bottom=rect
    return dict(id=identifier,entity_type=kind,rect=dict(left=left,top=top,right=right,bottom=bottom,
                width=right-left,height=bottom-top),observed_value=text,confidence=.95,
                semantic_name=text or "Graphic",source="rapidocr" if text else "yolo")


@pytest.fixture
def local(setup,monkeypatch):
    layout,native,root=setup
    monkeypatch.setattr(control_observation,"_check_foreground",lambda _w:None)
    calls=[]
    detections=[element(1,"graphic-control",(40,30,160,90)),element(2,"text",(65,45,135,70),"Save")]
    def analyze(source,output,*,progress):
        with Image.open(source) as image:calls.append(image.size)
        output=Path(output)/str(len(calls));output.mkdir(parents=True)
        Image.new("RGB",(400,150),"white").save(output/"combined.png")
        (output/"elements.json").write_text('{}')
        return dict(elements=detections,element_count=len(detections),
                    coordinates_json=str(output/"elements.json"),annotated_image=str(output/"combined.png"))
    monkeypatch.setattr(local_image_worker,"analyze",analyze)
    window=scanner.uiautomation.ControlFromHandle(1)
    button=window._children[0]
    points=[]
    def info(x,y,**kw):
        assert kw=={"raise_errors":True}
        points.append((x,y))
        return [dict(ControlType="WindowControl",Name="Example"),dict(ControlType="GroupControl",Name="layout"),
                dict(ControlType="ButtonControl",Name="Save")],button
    monkeypatch.setattr(utils,"get_control_info",info)
    def forbidden(*a,**kw):raise AssertionError("Full UIA traversal or fresh XPath generation was used")
    monkeypatch.setattr(scanner,"_walk_controls",forbidden)
    monkeypatch.setattr(scanner,"_native_action_controls",forbidden)
    monkeypatch.setattr(scanner,"get_control_xpath",forbidden)
    return layout,root,window,button,detections,calls,points


def scan(layout,root):
    return scanner.scan_window("Example","","","","test",manual_layout=layout,semantic_mode="native",root=root)


def records(root):return knowledge.load_index(knowledge.app_dir("chosen-app",root))["controls"]


@pytest.mark.parametrize("failure_first", [False, True])
def test_one_stale_point_keeps_other_controls_and_failed_visual_observation(local, monkeypatch, failure_first):
    layout, root, _window, _button, detections, *_ = local
    detections[1] = element(2, "text-label", (20, 100, 100, 125), "Status")
    if failure_first:
        detections.reverse()
    original = utils.get_control_info
    def lookup(x, y, **kwargs):
        if y > 90:
            raise ValueError("Recorded control is no longer a child of its parent")
        return original(x, y, **kwargs)
    monkeypatch.setattr(utils, "get_control_info", lookup)
    result = scan(layout, root)
    saved = records(root)
    assert any(r["control_type"] == "ButtonControl" for r in saved)
    failed = next(r for r in saved if r["control_type"] == "VisualElement")
    assert not failed["location"]["Xpath"]
    assert failed["verification"]["reason"] == "point_binding_failed"
    assert result["visual_only_observations"] == 1
    assert result["failures"][0]["diagnostics"]["error_type"] == "ValueError"


@pytest.mark.parametrize("error", [PermissionError("access denied"), RuntimeError("provider disconnected")])
def test_global_lookup_failures_are_not_silently_downgraded(local, monkeypatch, error):
    layout, root, *_ = local
    def lookup(*a, **kw):
        raise error
    monkeypatch.setattr(utils, "get_control_info", lookup)
    with pytest.raises(type(error), match=str(error)):
        scan(layout, root)


def test_basic_scan_analyzes_semantics_without_activation_or_input(local,monkeypatch):
    from easy_uiauto import library_actions
    from easy_uiauto.ctrl import Controller
    layout,root,window,button,*_=local
    window.NativeWindowHandle=1
    monkeypatch.setattr(library_actions,"bind_window",lambda _:window)
    monkeypatch.setattr(library_actions,"owned_windows",lambda _:[])
    def forbidden(*a,**kw):raise AssertionError("basic learning attempted desktop input")
    monkeypatch.setattr(scanner,"_activate_window",forbidden)
    monkeypatch.setattr(Controller,"activate_window_handle",forbidden)
    monkeypatch.setattr(Controller,"execute_resolved",forbidden)
    candidates=[]
    def semantics(image,items,*a):
        candidates.extend(items)
        return {1:dict(semantic_name="保存文档",description="保存当前文档",confidence=.95)}
    monkeypatch.setattr(scanner,"analyze_control_semantics",semantics)
    scanner.scan_window("Example","","","","test",manual_layout=layout,semantic_mode="api",root=root,
                        observation_owner=dict(handle=1))
    saved=records(root)[0]
    assert candidates[0]["uia_name"]=="Save" and candidates[0]["relative_rect"]["left"]==40
    assert saved["semantic_name"]=="保存文档" and saved["region_id"]=="toolbar"
    assert saved["location"]["Xpath"][-1]["Name"]=="Save"


def test_gui_region_default_annotates_then_uses_recorder_point_method(local):
    layout,root,_window,_button,_detections,calls,points=local
    result=scan(layout,root)
    assert result["strategy"]=="local-visual" and result["coverage"]["uia_enumerated"]==0
    assert calls==[(400,150)] and points==[(100,60),(100,57)]
    saved=records(root)
    assert len(saved)==1 and result["visual_only_observations"]==0
    native=next(r for r in saved if r["control_type"]=="ButtonControl")
    assert len(native["location"]["Xpath"])==3
    assert native["location"]["Xpath"][1]["ControlType"]=="GroupControl"
    assert native["detection"]["method"]=="local-yolo-rapidocr+get_control_info"
    assert native["observed_value"]=="Save"
    assert len(native["visual_components"])==2
    assert all(c["native_evidence"] for c in native["visual_components"])
    assert native["visual_components"][1]["observed_value"]=="Save"
    assert result["annotated_controls"]==1
    directory=knowledge.app_dir("chosen-app",root)
    with Image.open(directory/native["image"]) as crop:assert crop.size==(120,60)


@pytest.mark.parametrize("reverse",[False,True])
def test_nested_icon_and_label_bind_once_regardless_of_detection_order(local,reverse):
    layout,root,_window,button,detections,*_=local
    # The outer/inner icon IoU is only .2, just like the real toolbar duplicates.
    detections[:]=[element(1,"graphic-control",(40,30,160,90)),
                   element(2,"graphic-control",(70,40,110,70)),
                   element(3,"text-label",(80,45,110,65),"N")]
    if reverse:detections.reverse()
    button.IsEnabled=False
    result=scan(layout,root)
    saved=records(root)
    assert len(saved)==1 and not result["failures"]
    assert result["visual_targets"]==3 and result["visual_only_observations"]==0
    assert saved[0]["control_type"]=="ButtonControl" and saved[0]["entity_type"]=="graphic-control"
    assert saved[0]["observed_value"]=="N" and len(saved[0]["visual_components"])==3
    assert saved[0]["is_enabled_at_scan"] is False
    again=scan(layout,root)
    assert again["observed_control_ids"]==result["observed_control_ids"] and len(records(root))==1


def test_distinct_graphics_and_text_on_shared_custom_canvas_are_retained(local):
    layout,root,_window,button,detections,*_=local
    button.ControlTypeName="CustomControl"
    detections[:]=[element(1,"graphic-control",(40,30,95,90)),
                   element(2,"graphic-control",(105,30,160,90)),
                   element(3,"text-label",(45,35,100,70),"Cell")]
    result=scan(layout,root)
    assert len(records(root))==3 and result["visual_only_observations"]==2
    assert {f["reason"] for f in result["failures"]}=={"shared_native_control"}
    assert any(r["observed_value"]=="Cell" and r["control_type"]=="VisualElement" for r in records(root))


def test_nested_objects_on_custom_canvas_are_not_merged_by_containment_alone(local):
    layout,root,_window,button,*_=local
    button.ControlTypeName="CustomControl"
    detections=local[4]
    detections[:]=[element(1,"graphic-control",(40,30,160,90)),
                   element(2,"graphic-control",(60,40,120,70))]
    result=scan(layout,root)
    assert len(records(root))==2 and result["visual_only_observations"]==1
    assert result["failures"][0]["reason"]=="shared_native_control"


def test_missing_uia_text_remains_an_independent_visual_record(local,monkeypatch):
    layout,root,_window,_button,detections,*_=local
    detections[:]=[element(1,"text-label",(20,100,100,125),"Status")]
    monkeypatch.setattr(utils,"get_control_info",lambda *a,**kw:([],None))
    result=scan(layout,root)
    record=records(root)[0]
    assert result["visual_only_observations"]==1 and record["observed_value"]=="Status"
    assert record["control_type"]=="VisualElement" and not record["location"]["Xpath"]


@pytest.mark.parametrize("kind",["GroupControl","ToolBarControl","WindowControl","MenuControl"])
def test_layout_hits_are_only_visual_observations(local,kind):
    layout,root,_window,button,*_=local
    button.ControlTypeName=kind
    result=scan(layout,root)
    assert result["uia_controls_seen"]==0
    assert all(r["control_type"]=="VisualElement" and not r["location"]["Xpath"] for r in records(root))
    assert all(r["verification"]["reason"]=="layout_container" for r in records(root))


def test_independent_ocr_text_control_is_saved(local,monkeypatch):
    layout,root,window,_button,detections,*_=local
    label=FakeControl("Label","TextControl",(20,100,100,125));label._parent=window
    detections[:]=[element(1,"text",(20,100,100,125),"Label")]
    monkeypatch.setattr(utils,"get_control_info",lambda x,y,**kw:([dict(ControlType="WindowControl",Name="Example"),
                          dict(ControlType="TextControl",Name="Label")],label))
    scan(layout,root)
    assert records(root)[0]["control_type"]=="TextControl"
    assert records(root)[0]["observed_value"]=="Label"


def test_foreign_window_and_missing_points_never_gain_native_locations(local,monkeypatch):
    layout,root,_window,button,*_=local
    other=FakeControl("Other","WindowControl",(0,0,400,300),[button])
    scan(layout,root)
    assert all(r["verification"]["reason"]=="outside_selected_window" for r in records(root))
    monkeypatch.setattr(utils,"get_control_info",lambda x,y,**kw:([],None))
    scan(layout,root)
    assert len(records(root))==2
    assert all(not r["location"]["Xpath"] for r in records(root))


def test_empty_detection_does_not_fall_back_or_quarantine_existing(local):
    layout,root,_window,_button,detections,*_=local
    first=scan(layout,root)
    detections.clear()
    second=scan(layout,root)
    assert second["controls_saved_this_scan"]==0 and second["controls_retained_unobserved"]==1
    assert {r["id"] for r in records(root)}==set(first["observed_control_ids"])


def test_inference_error_does_not_overwrite_control_references(local,monkeypatch):
    layout,root,*_=local
    scan(layout,root)
    before=records(root)
    monkeypatch.setattr(local_image_worker,"analyze",lambda *a,**kw:(_ for _ in ()).throw(RuntimeError("no local runtime")))
    with pytest.raises(RuntimeError,match="no local runtime"):scan(layout,root)
    assert records(root)==before


def test_relearning_preserves_human_semantics_for_native_and_visual_controls(local,monkeypatch):
    layout,root,_window,_button,detections,*_=local
    # Independent text without its own UIA node must remain human-editable.
    detections[1]=element(2,"text-label",(20,100,100,125),"Status")
    info=utils.get_control_info
    monkeypatch.setattr(utils,"get_control_info",lambda x,y,**kw:([],None) if y>90 else info(x,y,**kw))
    scan(layout,root)
    directory=knowledge.app_dir("chosen-app",root)
    saved=records(root)
    for record in saved:
        record.update(semantic_source="manual",semantic_name="人工名称",intent="manual-label",description="我写的说明")
        knowledge.save_control(directory,record)
    result=scan(layout,root)
    assert result["controls_saved_this_scan"]==2
    assert not any("KeyError" in f.get("error","") for f in result["failures"])
    assert all(r["semantic_name"]=="人工名称" and r["description"]=="我写的说明" for r in records(root))


def test_screen_origin_is_added_once_including_negative_monitors(local,monkeypatch):
    layout,root,window,button,detections,_calls,points=local
    for control in (window,button):
        r=control.BoundingRectangle
        r.left-=800;r.right-=800;r.top+=50;r.bottom+=50
    monkeypatch.setattr(scanner,"_window_rect",lambda _w:dict(left=-800,top=50,right=-400,bottom=350,width=400,height=300))
    scan(layout,root)
    assert points==[(-700,110),(-700,107)]
    assert any(r["control_type"]=="ButtonControl" for r in records(root))


@pytest.mark.parametrize("referenced",[False,True])
def test_only_unused_legacy_layout_records_are_archived_with_images(local,referenced):
    layout,root,*_=local
    directory=knowledge.initialize_app("chosen-app","Example",root)
    image=directory/"images/controls/legacy.png";image.parent.mkdir(parents=True,exist_ok=True)
    Image.new("RGB",(20,20),"black").save(image)
    knowledge.save_control(directory,dict(id="legacy",app_id="chosen-app",page_id="manual-main",region_id="toolbar",
        control_type="GroupControl",semantic_name="Layout",name="Layout",actions=[],scan_strategy="manual-regions",
        semantic_source="uia-metadata",status="observed",image="images/controls/legacy.png",location={},
        normalized_rect=dict(left=0,top=0,right=.5,bottom=.1)))
    if referenced:
        knowledge.write_markdown(directory/"workflows/use.md",dict(kind="workflow",id="use",control_id="legacy"),"")
    result=scan(layout,root)
    assert result["archived_layout_records"]==int(not referenced)
    assert any(r["id"]=="legacy" for r in records(root))==referenced
    if not referenced:
        archive=next((directory/".easy_uiauto/archive").rglob("legacy.md"))
        saved,_body=knowledge.read_markdown(archive)
        assert (directory/saved["image"]).is_file()


def test_permission_errors_from_point_method_are_not_mislabeled_as_missing(local,monkeypatch):
    layout,root,*_=local
    def denied(*a,**kw):raise PermissionError("UIA E_ACCESSDENIED")
    monkeypatch.setattr(utils,"get_control_info",denied)
    with pytest.raises(PermissionError,match="UIA E_ACCESSDENIED"):scan(layout,root)
    assert records(root)==[]


@pytest.mark.parametrize("raise_errors",[False,True])
def test_point_method_keeps_legacy_result_and_balances_com_on_permission_failure(monkeypatch,raise_errors):
    events=[]
    monkeypatch.setattr(utils,"CoInitialize",lambda:events.append("init"))
    monkeypatch.setattr(utils,"CoUninitialize",lambda:events.append("uninit"))
    monkeypatch.setattr(utils,"push_message",lambda _m:None)
    monkeypatch.setattr(utils.uiautomation,"SetGlobalSearchTimeout",lambda _t:None)
    def denied(*a):raise PermissionError("E_ACCESSDENIED")
    monkeypatch.setattr(utils.uiautomation,"ControlFromPoint",denied)
    if raise_errors:
        with pytest.raises(PermissionError,match="E_ACCESSDENIED"):utils.get_control_info(10,20,raise_errors=True)
    else:assert utils.get_control_info(10,20)==([],None)
    assert events==["init","uninit"]
