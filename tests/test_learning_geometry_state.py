"""Resize/repaint tolerance must retain both scope and meaningful state changes."""
from copy import deepcopy

import pytest
from PIL import Image, ImageDraw

from easy_uiauto import state_matching, region_learning, library_actions, image_models
from easy_uiauto.mcp import image_storage, knowledge, manual_layout, scanner
from test_knowledge_library import library, add_control
from test_region_learning import backend, content_frame


def screen(size=(400,300)):
    image=Image.new("RGB",(400,300),"white");draw=ImageDraw.Draw(image)
    draw.rectangle((20,40,180,120),fill="navy");draw.rectangle((200,40,380,120),outline="black",width=2)
    draw.text((30,150),"Project / input: 123",fill="black")
    return image.resize(size,Image.Resampling.NEAREST)


def evidence(image, native="same-native", scope=None):
    return dict(version=1,native=native,scope=scope or dict(kind="main"),image=state_matching.visual_evidence(image))


def test_scaled_and_slightly_repainted_image_matches_but_local_change_does_not():
    first=screen();scaled=screen((800,600))
    assert state_matching.matches(evidence(first),evidence(scaled))
    repaint=first.copy();repaint.putpixel((399,299),(252,252,252))
    assert state_matching.matches(evidence(first),evidence(repaint))
    changed=first.copy();ImageDraw.Draw(changed).rectangle((210,55,230,75),fill="black")
    assert not state_matching.matches(evidence(first),evidence(changed))


@pytest.mark.parametrize("change",["value","scope","missing","invalid-image"])
def test_matching_image_never_overrides_different_native_state(change):
    first=evidence(screen());second=deepcopy(first)
    if change=="value":second["native"]="different-input-value"
    elif change=="scope":second["scope"]=dict(kind="owned",title="Dialog")
    elif change=="missing":second={}
    else:second["image"]="bad-base64"
    assert not state_matching.matches(first,second)


def test_resume_matches_same_state_after_resize_and_reuses_existing_node(backend,monkeypatch,content_frame):
    old=evidence(screen())
    node=dict(id="original",fingerprint="old-pixels",state_evidence=old,route=[],region_ids=[backend.region["id"]])
    backend.previous_graph=dict(root="original",states={"original":node})
    content_frame.update(state_id="scaled-pixels",state_evidence=evidence(screen((800,600))))
    monkeypatch.setattr(backend,"stable_frame",lambda:content_frame)
    monkeypatch.setattr(image_models,"inspect_models",lambda **kw:dict(ok=True))
    monkeypatch.setattr(library_actions,"activate_application",lambda _:None)
    monkeypatch.setattr(scanner,"scan_window",lambda *a,**kw:dict(observed_control_ids=[],truncated=False))
    observed=backend.observe(root=True)
    assert observed["node_id"]=="original" and observed["region_ids"]==[backend.region["id"]]
    assert observed["state_evidence"]==content_frame["state_evidence"]
    # Repainting during the same session must not introduce duplicate states.
    content_frame["state_id"]="repainted-again"
    assert backend.observe() is observed


def test_ambiguous_history_is_not_merged(backend):
    value=evidence(screen())
    backend.previous_graph=dict(states={i:dict(id=i,fingerprint=i,state_evidence=value) for i in ("one","two")})
    assert backend.previous_node(dict(state_id="fresh",state_evidence=value),[],True) is None


def test_resize_during_learning_refreshes_observation_without_adding_a_state(backend, monkeypatch, content_frame):
    from easy_uiauto import region_actions
    content_frame.update(image=screen(),state_evidence=evidence(screen()))
    monkeypatch.setattr(backend,"stable_frame",lambda:content_frame)
    monkeypatch.setattr(image_models,"inspect_models",lambda **kw:dict(ok=True))
    monkeypatch.setattr(library_actions,"activate_application",lambda _:None)
    monkeypatch.setattr(region_actions,"open_region",lambda *a,**kw:None)
    scans=[]
    def scan(*a,**kw):
        scans.append(kw["manual_layout"])
        return dict(observed_control_ids=[],truncated=False)
    monkeypatch.setattr(scanner,"scan_window",scan)
    first=backend.observe(root=True)
    content_frame.update(state_id="scaled",image=screen((800,600)),state_evidence=evidence(screen((800,600))))
    content_frame["bounds"].update(width=800,height=600,right=850,bottom=660)
    content_frame["root"]._children[0].BoundingRectangle.right=850
    second=backend.observe()
    assert second is not first and second["image"].size==(800,600)
    assert second["node_id"]==first["node_id"] and second["region_ids"]==first["region_ids"]
    assert len(scans)==2 and scans[-1]["image_size"]==dict(width=800,height=600)


def test_legacy_session_uses_actual_before_frame_and_current_locators(backend,library,monkeypatch,content_frame):
    record=add_control(library)
    reference=image_storage.save_interaction_image(backend.directory,screen())
    node=dict(id="legacy",fingerprint="old",controls=[record["id"]],branches=[dict(evidence=dict(steps=[dict(
        before=dict(context_image=reference,main_rect=dict(left=0.,top=0.,right=1.,bottom=1.)))]))])
    backend.previous_graph=dict(states={node["id"]:node})
    content_frame["image"]=screen((800,600))
    assert backend.previous_node(content_frame,[],True)==node
    def missing(*a,**kw):raise ValueError("missing control")
    monkeypatch.setattr(library_actions,"resolve_record",missing)
    assert backend.previous_node(content_frame,[],True) is None


def test_drawn_region_scales_only_when_reference_image_matches(tmp_path,monkeypatch):
    old=screen();path=tmp_path/"layout.png";old.save(path)
    region=dict(id="canvas",name="画布",normalized_rect=dict(left=0,top=0,right=1,bottom=.5))
    layout=dict(image_size=dict(width=400,height=300),regions=[region],reference_image=str(path),restrict_regions=True)
    bounds=dict(left=50,top=70,right=850,bottom=670,width=800,height=600)
    monkeypatch.setattr(scanner,"_screenshot_window",lambda _:screen((800,600)))
    adapted=manual_layout.adapt_to_window(layout,bounds,window=object())
    assert adapted["image_size"]==dict(width=800,height=600) and adapted["geometry_adapted"]=="scaled-reference"
    assert adapted["regions"]==layout["regions"] and layout["image_size"]==dict(width=400,height=300)
    assert manual_layout.segmentation(adapted,bounds)["regions"][0]["rect"]["bottom"]==300
    changed=screen((800,600));ImageDraw.Draw(changed).rectangle((0,0,400,100),fill="black")
    monkeypatch.setattr(scanner,"_screenshot_window",lambda _:changed)
    with pytest.raises(ValueError,match="边界"):
        manual_layout.segmentation(manual_layout.adapt_to_window(layout,bounds,window=object()),bounds)


def test_native_region_resize_keeps_semantics_and_reference_bounds(backend,library,content_frame):
    original=library.learning_layout("demo",backend.expected,region_id=backend.region["id"])
    before=deepcopy(original)
    bounds={**content_frame["bounds"],"width":800,"right":850}
    content_frame["root"]._children[0].BoundingRectangle.right=850
    adapted=manual_layout.adapt_to_window(original,bounds,window=content_frame["root"],owner=backend.expected)
    assert original==before
    assert adapted["image_size"]==dict(width=800,height=300)
    assert adapted["regions"][0]["id"]==original["regions"][0]["id"]
    assert adapted["regions"][0]["location"]==original["regions"][0]["location"]
    assert adapted["restrict_regions"] and len(adapted["regions"])==1


def test_native_resize_rejects_replacement_panel(backend,library,content_frame):
    layout=library.learning_layout("demo",backend.expected,region_id=backend.region["id"])
    bounds={**content_frame["bounds"],"width":800,"right":850}
    content_frame["root"]._children[0].AutomationId="different-panel"
    with pytest.raises(ValueError,match="不同控件"):
        manual_layout.adapt_to_window(layout,bounds,window=content_frame["root"],owner=backend.expected)


def test_planner_uses_branch_history_without_embedding_state_image_in_text(backend, monkeypatch):
    previous = dict(id="existing", branches=[], state_evidence={"image":"a" * 170000})
    def request(_settings, payload, _image):
        assert payload["previous_learning"] == dict(id="existing", branches=[])
        return dict(branches=[])
    monkeypatch.setattr(region_learning,"_request_plan",request)
    backend.plan(dict(controls=[],image=screen(),previous_learning=previous), [])
    assert len(previous["state_evidence"]["image"]) == 170000
