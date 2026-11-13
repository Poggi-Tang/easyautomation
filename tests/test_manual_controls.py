"""Manual control capture persists state, routes and existing user annotations."""
import pytest
from PIL import Image

from easy_uiauto import workflows
from easy_uiauto.mcp import dynamic_regions, knowledge
from test_knowledge_library import library, add_control
from test_region_picker import picked


def test_manual_control_saves_region_images_and_inherits_entry_steps(library, tmp_path):
    region = library.save_picked_region("demo", picked(tmp_path))
    entry = add_control(library)
    library.configure_region("demo", region, steps=[workflows.click_step(entry)])
    result = picked(tmp_path, color="green", rect=dict(left=.2, top=.1, right=.4, bottom=.3))
    record = library.save_captured_control("demo", result, name="开始", description="启动运行", region_id=region)
    assert record["semantic_name"] == "开始" and record["description"] == "启动运行"
    assert record["region_id"] == region and record["semantic_source"] == "manual"
    assert record["location"] == result["selection"]["location"] and record["status"] == "observed"
    assert library.entry_steps("demo", record)[0]["target"]["control_id"] == entry["id"]
    with Image.open(library.asset("demo", record["image"])) as image:
        assert image.size == (80, 60) and image.getpixel((0, 0)) == (0, 128, 0)
    # Editing layout geometry must not undo the explicit destination.
    library._reassign(library.directory("demo"), [])
    saved = knowledge.find_control_record(library.directory("demo"), record["id"])[1]
    assert saved["region_id"] == region
    # A later scan may suggest another geometric assignment; keep the user choice.
    knowledge.save_control(library.directory("demo"), {**saved, "region_id":"automatic", "region_source":"manual"})
    assert knowledge.find_control_record(library.directory("demo"), record["id"])[1]["region_id"] == region


def test_repick_reuses_learned_id_and_keeps_annotations_and_quarantine(library, tmp_path):
    result = picked(tmp_path)
    old = {**add_control(library), "location": result["selection"]["location"],
           "semantic_name":"已有名称", "description":"已有说明", "aliases":["启动"],
           "entry_steps":[], "status":"quarantined", "quarantine":{"active":True,"reason":"failed"}}
    knowledge.save_control(library.directory("demo"), old)
    record = library.save_captured_control("demo", result, region_id="", name="新名称", description="新说明")
    assert record["id"] == old["id"] and len(library.read("demo")["controls"]) == 1
    assert record["semantic_name"] == "新名称" and record["description"] == "新说明"
    assert record["aliases"] == ["启动"] and record["entry_steps"] == []
    assert record["status"] == "quarantined" and record["quarantine"]["active"]
    # Picking the same record as an entry must not reset manual edits or grouping.
    again = library.save_captured_control("demo", result)
    assert again["semantic_name"] == "新名称" and again["region_source"] == "manual-selection"


def test_same_control_in_two_states_keeps_separate_records(library, tmp_path):
    region = library.save_picked_region("demo", picked(tmp_path))
    entry = add_control(library)
    variant = library.save_variant("demo", library.regions("demo")[0], picked(tmp_path), name="绘图",
        steps=[workflows.click_step(entry)], condition=dict(location=entry["location"], kind="exists"))
    state = next(r for r in library.read("demo")["regions"] if r["id"] == variant)
    first = library.save_captured_control("demo", picked(tmp_path), name="主页", region_id=region)
    second = library.scoped(state["page_id"]).save_captured_control("demo", picked(tmp_path), name="绘图", region_id=variant)
    assert first["id"] != second["id"] and first["page_id"] != second["page_id"]
    assert library.entry_steps("demo", second)[0]["target"]["control_id"] == entry["id"]


def test_manual_dynamic_destination_keeps_route_despite_different_frame(library, tmp_path):
    entry = add_control(library)
    region = dict(id="dynamic-child", app_id="demo", page_id=library.page_id, kind="region",
        name="子菜单", source="manual-dynamic", navigation_path=[entry["id"]], parent_control_id=entry["id"],
        window_scope={"kind":"main"}, visual_observation={"context_image":"old-frame.png"})
    knowledge.save_region(library.directory("demo"), region)
    record = library.save_captured_control("demo", picked(tmp_path), name="菜单项", region_id=region["id"])
    snapshot = library.read("demo")
    tree = dynamic_regions.build_tree(snapshot["controls"], snapshot["regions"])
    assert tree["parents"]["control:"+record["id"]] == "region:"+region["id"]
    assert record["navigation_path"] == [entry["id"]]
    assert library.entry_steps("demo", record)[0]["target"]["control_id"] == entry["id"]


@pytest.mark.parametrize("invalid", ["missing-region", "wrong-window", "bad-rectangle"])
def test_invalid_destination_does_not_write_a_control(library, tmp_path, invalid):
    region = library.save_picked_region("demo", picked(tmp_path))
    result = picked(tmp_path)
    if invalid == "missing-region":region = "gone"
    elif invalid == "wrong-window":result["selection"]["window_scope"] = dict(kind="owned", class_name="Popup")
    else:result["selection"]["normalized_rect"]["right"] = -1
    with pytest.raises(ValueError):library.save_captured_control("demo", result, name="按钮", region_id=region)
    assert not library.read("demo")["controls"]


def test_ambiguous_duplicate_is_not_silently_merged(library, tmp_path):
    result = picked(tmp_path)
    first = library.save_captured_control("demo", result, region_id="")
    knowledge.save_control(library.directory("demo"), {**first, "id":"duplicate"})
    with pytest.raises(ValueError, match="多个"):
        library.save_captured_control("demo", result, region_id="")
    assert len(library.read("demo")["controls"]) == 2


def test_repick_does_not_overwrite_a_record_moved_to_another_route(library, tmp_path):
    entry = add_control(library)
    result = picked(tmp_path)
    first = library.save_captured_control("demo", result, name="菜单按钮", region_id="")
    knowledge.save_control(library.directory("demo"), {**first, "navigation_path":[entry["id"]]})
    second = library.save_captured_control("demo", result, name="主页按钮", region_id="")
    assert first["id"] != second["id"] and second["navigation_path"] == []
    saved = knowledge.find_control_record(library.directory("demo"), first["id"])[1]
    assert saved["navigation_path"] == [entry["id"]] and saved["semantic_name"] == "菜单按钮"
