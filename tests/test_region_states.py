"""Independent state namespaces and region navigation must not corrupt legacy data."""
from copy import deepcopy
from types import SimpleNamespace

import pytest

from easy_uiauto import workflows
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library, add_control
from test_region_picker import picked


def test_identical_geometry_in_another_state_preserves_old_image_controls_and_locator(library,tmp_path):
    base_id=library.save_picked_region("demo",picked(tmp_path))
    base=next(r for r in library.regions("demo") if r["id"]==base_id)
    entry=add_control(library);control={**entry,"region_id":base_id.split(".",1)[1]};knowledge.save_control(library.directory("demo"),control)
    before=deepcopy(library.read("demo"));image_before=library.asset("demo",base["image"]).read_bytes()
    condition=dict(location=entry["location"],kind="selected",label="绘图")
    variant=library.save_variant("demo",base,picked(tmp_path,color="green"),name="绘图",steps=[workflows.click_step(entry)],condition=condition)
    state_id=variant.split(".",1)[0];scoped=library.scoped(state_id)
    assert scoped.page("demo")["region_state"]["area_id"]==base_id
    assert scoped.regions("demo")[0]["normalized_rect"]==base["normalized_rect"]
    assert library.asset("demo",base["image"]).read_bytes()==image_before
    assert library.read("demo")["controls"]==before["controls"]
    assert len(library.regions("demo"))==1
    with pytest.raises(ValueError,match="重叠"):scoped.save_picked_region("demo",picked(tmp_path,color="green"))


def test_each_default_region_has_independent_entry_steps(library,tmp_path):
    first=library.save_picked_region("demo",picked(tmp_path))
    second=library.save_picked_region("demo",picked(tmp_path,rect=dict(left=0,top=.6,right=1,bottom=1)))
    entry=add_control(library)
    library.configure_region("demo",first,steps=[workflows.click_step(entry)])
    records=library.regions("demo")
    assert len(library.record_state("demo",next(r for r in records if r["id"]==first))["steps"])==1
    assert not library.record_state("demo",next(r for r in records if r["id"]==second))["steps"]


def test_variant_requires_entry_and_guard_before_writing_any_page(library,tmp_path):
    identifier=library.save_picked_region("demo",picked(tmp_path));base=library.regions("demo")[0]
    before=deepcopy(library.read("demo"))
    with pytest.raises(ValueError,match="前置步骤"):
        library.save_variant("demo",base,picked(tmp_path),name="另一状态")
    assert library.read("demo")==before
    entry=add_control(library);before=deepcopy(library.read("demo"))
    with pytest.raises(ValueError,match="状态条件"):
        library.save_variant("demo",base,picked(tmp_path),name="另一状态",steps=[workflows.click_step(entry)],
                             condition=dict(location=entry["location"],kind="invalid"))
    assert library.read("demo")==before


def test_local_learning_layout_filters_one_region_without_rewriting_whole_page(library,tmp_path):
    identifier=library.save_picked_region("demo",picked(tmp_path))
    layout=library.learning_layout("demo",dict(handle=1,process_id=2,class_name="Window"),region_id=identifier)
    assert layout["restrict_regions"] and len(layout["regions"])==1


def test_state_ids_and_known_locations_remain_stable_on_rename(library,tmp_path):
    identifier=library.save_picked_region("demo",picked(tmp_path));entry=add_control(library)
    before=entry["location"]
    library.configure_region("demo",identifier,steps=[workflows.click_step(entry)],state_name="绘图")
    library.configure_region("demo",identifier,steps=[workflows.click_step(entry)],state_name="另一个名称")
    assert library.regions("demo")[0]["id"]==identifier
    assert library.read("demo")["controls"][0]["location"]==before


def test_control_route_override_and_repeated_checks_get_new_ids(library):
    record=add_control(library)
    step=workflows.click_step(record)
    step["checks"]=[dict(id="check-one",kind="exists",target=step["target"])]
    second=deepcopy(step);second["id"]="step-two";second["checks"][0]["id"]="check-two"
    library.update_control("demo",record["id"],name="名称",description="",aliases=[],entry_steps=[step,second])
    record=library.read("demo")["controls"][0]
    route=library.entry_steps("demo",record)
    assert len(route)==2 and route[0]["target"]==route[1]["target"]
    assert route[0]["checks"][0]["id"]!=route[1]["checks"][0]["id"]
    assert record["entry_steps"][0]["id"]==step["id"]


def test_manual_control_entry_survives_rescan_merge():
    route=[dict(id="step",action="click",target=dict(control_id="button"))]
    record=knowledge.merge_control_state(dict(entry_steps=route),dict(id="control",status="observed"))
    assert record["entry_steps"]==route


def test_picked_control_keeps_its_own_context_instead_of_using_a_different_page(library,tmp_path):
    from PIL import Image
    library.save_picked_region("demo",picked(tmp_path))
    selection=picked(tmp_path,color="green",rect=dict(left=.2,top=.1,right=.8,bottom=.4))
    record=library.save_captured_control("demo",selection)
    observation=record["visual_observation"]
    assert observation["normalized_rect"]==selection["selection"]["normalized_rect"]
    assert observation["purpose"]=="learning-reference"
    with Image.open(library.asset("demo",observation["context_image"])) as image:
        assert image.getpixel((0,0))==(0,128,0)
