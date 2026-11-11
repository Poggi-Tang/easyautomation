import json

import pytest
from PIL import Image

from easy_uiauto.knowledge_library import KnowledgeLibrary, PAGE_ID
from easy_uiauto.mcp import knowledge
from easy_uiauto.mcp.protocol import location_from_xpath


@pytest.fixture
def library(tmp_path):
    store = KnowledgeLibrary(tmp_path / "vault")
    store.bind("demo", "测试软件", {"title": "Test", "class_name": "Window", "handle": 123, "process_id": 456})
    image = tmp_path / "capture.png"
    Image.new("RGB", (400, 300), "white").save(image)
    store.save_capture("demo", {"image": str(image), "width": 400, "height": 300,
                                "window_name": "Test", "window_class": "Window"})
    return store


def add_control(library, *, status="observed", quarantine=None):
    location = location_from_xpath([{"ControlType": "WindowControl", "Name": "Test"},
                                    {"ControlType": "ButtonControl", "Name": "Open", "foundIndex": 1}])
    record = {"id": "control-one", "app_id": "demo", "page_id": PAGE_ID, "region_id": "unassigned",
              "name": "Open", "semantic_name": "打开", "description": "", "status": status,
              "location": location, "actions": [], "command": f"{PAGE_ID}.unassigned.open",
              "normalized_rect": dict(left=.1, top=.1, right=.2, bottom=.2)}
    if quarantine: record["quarantine"] = quarantine
    knowledge.save_control(library.directory("demo"), record)
    return record


def test_binding_and_capture_use_existing_markdown_schema_without_runtime_handles(library):
    index = library.read("demo")
    assert index["application"]["name"] == "测试软件"
    assert index["pages"][0]["region_source"] == "manual"
    metadata = json.dumps(index)
    assert '"handle"' not in metadata and '"process_id"' not in metadata
    assert library.asset("demo", index["pages"][0]["layout_image"]).is_file()


def test_navigation_edits_persist_and_removed_prerequisite_blocks_replay(library):
    parent = add_control(library)
    target = {**parent, "id": "new", "name": "New", "navigation_path": []}
    knowledge.save_control(library.directory("demo"), target)
    updated = library.update_control("demo", "new", name="新建", description="", aliases=[],
                                     navigation_path=[parent["id"]])
    assert updated["navigation_path"] == [parent["id"]]
    assert [r["id"] for r in library.click_plan("demo", "new")] == [parent["id"], "new"]
    library.save_execution("demo", "new", {"ok": True, "input_attempted": True})
    saved = next(r for r in library.read("demo")["controls"] if r["id"] == "new")
    assert saved["status"] == target["status"] and saved["last_click_execution"]["ok"]
    library.archive("demo", "control", parent["id"])
    with pytest.raises(ValueError, match="删除"):
        library.click_plan("demo", "new")


@pytest.mark.parametrize("learning", [False, True])
def test_execution_evidence_does_not_update_learning_reference_unless_learning(library, monkeypatch, learning):
    from copy import deepcopy
    from easy_uiauto.mcp import control_observation, image_storage, scanner

    record = add_control(library)
    directory = library.directory("demo")
    path, variants, digest = scanner._save_control_images(directory, record["id"], Image.new("RGB", (30, 20), "white"), {})
    context = image_storage.save_interaction_image(directory, Image.new("RGB", (400, 300), "white"))
    record.update(image=path.relative_to(directory).as_posix(), image_sha256=digest, image_variants=variants,
                  visual_observation={"phase": "located", "purpose": "learning-reference", "context_image": context})
    knowledge.save_control(directory, record)
    original = deepcopy(library.read("demo")["controls"][0])
    runtime_context = image_storage.save_interaction_image(directory, Image.new("RGB", (400, 300), "blue"))
    monkeypatch.setattr(control_observation, "capture", lambda *_a, **_kw: (
        {"phase": "before-action", "context_image": runtime_context}, Image.new("RGB", (30, 20), "blue")))
    evidence = library.capture_before_click("demo", record["id"], object(), object(), update_reference=learning)
    observed = library.read("demo")["controls"][0]
    keys = ("image", "image_sha256", "image_variants", "visual_observation", "location", "status")
    if learning:
        assert observed["image_sha256"] != original["image_sha256"]
        assert observed["visual_observation"]["purpose"] == "learning-reference"
    else:
        assert observed == original
    assert evidence["purpose"] == "execution-evidence"
    library.save_execution("demo", record["id"], {"ok": True, "steps": [
        {"control_id": record["id"], "before": evidence, "after": {"context_image": runtime_context}}]})
    saved = library.read("demo")["controls"][0]
    assert {k: saved[k] for k in keys} == {k: observed[k] for k in keys}
    assert library.asset("demo", evidence["context_image"]).exists()
    assert library.asset("demo", evidence["control_image"]).exists()


def test_region_creation_saves_image_and_reassigns_controls(library):
    add_control(library)
    identifier = library.save_region("demo", "工具栏", dict(left=0, top=0, right=1, bottom=.5))
    index = library.read("demo")
    assert index["controls"][0]["region_id"] == identifier.removeprefix(PAGE_ID + ".")
    assert index["controls"][0]["region_assignment"] == "contained"
    path = library.asset("demo", index["regions"][0]["image"])
    with Image.open(path) as image: assert image.size == (400, 150)
    library.save_region("demo", "下半区", dict(left=0, top=.5, right=1, bottom=1), identifier)
    index = library.read("demo")
    assert len(index["regions"]) == 1
    assert index["controls"][0]["region_id"] == "unassigned"
    assert ".unassigned." in index["controls"][0]["command"]


def test_overlapping_region_does_not_commit_partial_data(library):
    library.save_region("demo", "上", dict(left=0, top=0, right=1, bottom=.5))
    with pytest.raises(ValueError, match="重叠"):
        library.save_region("demo", "重叠", dict(left=0, top=.4, right=1, bottom=1))
    assert len(library.regions("demo")) == 1


def test_resize_invalidates_layout_confirmation(library, tmp_path):
    library.save_region("demo", "全部", dict(left=0, top=0, right=1, bottom=1))
    library.confirm_layout("demo")
    runtime = dict(handle=123, process_id=456, class_name="Window")
    assert library.learning_layout("demo", runtime)["image_size"] == dict(width=400, height=300)
    image = tmp_path / "large.png"; Image.new("RGB", (600, 500)).save(image)
    library.save_capture("demo", dict(image=str(image), width=600, height=500, window_name="Test", window_class="Window"))
    with pytest.raises(ValueError, match="确认"):
        library.learning_layout("demo", runtime)


def test_alias_does_not_replace_uia_name_and_locator_edits_invalidate_verification(library):
    record = add_control(library, status="verified")
    edited = library.update_control("demo", record["id"], name="业务入口", description="人工说明", aliases=["入口"])
    assert edited["name"] == "Open" and edited["semantic_name"] == "业务入口"
    edited = library.update_control("demo", record["id"], name="业务入口", description="", aliases=[],
                                    location=record["location"], match_mode="position")
    assert edited["status"] == "observed"
    assert edited["verification"]["location"] == "pending"
    assert edited["location"]["PARAMETERS"]["match_mode"] == "position"


def test_archiving_region_preserves_children_and_images(library):
    record = add_control(library)
    identifier = library.save_region("demo", "区域", dict(left=0, top=0, right=1, bottom=.5))
    region = library.regions("demo")[0]
    image = library.asset("demo", region["image"])
    archived = library.archive("demo", "region", identifier)
    assert archived.is_file() and image.is_file()
    assert library.read("demo")["controls"][0]["region_id"] == "unassigned"
    assert knowledge.find_control_record(library.directory("demo"), record["id"])


def test_binding_again_preserves_application_metadata(library):
    path = library.directory("demo") / "app.md"
    record, body = knowledge.read_markdown(path)
    record["version"] = "2.1"
    knowledge.write_markdown(path, record, body)
    library.bind("demo", "改名", dict(title="Other", class_name="NewWindow"))
    assert knowledge.read_markdown(path)[0]["version"] == "2.1"


@pytest.mark.parametrize("app_id", ["../outside", "C:/outside", "demo/sub", "", "DEMO"])
def test_application_paths_are_checked(library, app_id):
    with pytest.raises(ValueError): library.directory(app_id)


def test_export_and_python_example_are_valid_and_portable(library, tmp_path):
    record = add_control(library)
    compile(library.python_code(record), "example.py", "exec")
    path = library.export("demo", tmp_path / "export.json")
    assert json.loads(path.read_text(encoding="utf-8"))["controls"][0]["location"]["Name"] == "Open"
    assert library.asset("demo", "../../../../other.png") is None


def test_region_identifier_cannot_escape_its_directory(library):
    with pytest.raises(ValueError, match="ID"):
        library.save_region("demo", "区域", dict(left=0, top=0, right=1, bottom=.5), "../../outside")
    assert not library.regions("demo")
