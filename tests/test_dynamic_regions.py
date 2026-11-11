from copy import deepcopy

import pytest
from PIL import Image

from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.mcp import dynamic_regions, image_storage, knowledge
from easy_uiauto.mcp.protocol import location_from_xpath


@pytest.fixture
def popup(tmp_path):
    store = KnowledgeLibrary(tmp_path / "vault")
    store.bind("demo", "Demo", dict(title="Test", class_name="Window"))
    image = tmp_path / "layout.png"
    Image.new("RGB", (400, 300), "white").save(image)
    store.save_capture("demo", dict(image=str(image), width=400, height=300, window_name="Test", window_class="Window"))
    static = store.save_region("demo", "工具栏", dict(left=0, top=0, right=1, bottom=1))
    directory = store.directory("demo")
    frame = Image.new("RGB", (400, 300), "blue")
    context = image_storage.save_interaction_image(directory, frame)
    main = dict(ControlType="WindowControl", Name="Test")
    menu = dict(ControlType="MenuControl", Name="File", ClassName="#32768", foundIndex=1)
    records = []
    for identifier, kind, path, n, navigation in (
        ("file", "MenuItemControl", [main, dict(ControlType="MenuItemControl", Name="File")],
         dict(left=.1, top=.1, right=.2, bottom=.15), []),
        ("popup", "MenuControl", [main, menu], dict(left=.1, top=.2, right=.6, bottom=.6), ["file"]),
        ("new", "MenuItemControl", [main, menu, dict(ControlType="MenuItemControl", Name="New", foundIndex=1)],
         dict(left=.15, top=.25, right=.55, bottom=.32), ["file"]),
        ("open", "MenuItemControl", [main, menu, dict(ControlType="MenuItemControl", Name="Open", foundIndex=2)],
         dict(left=.15, top=.4, right=.55, bottom=.47), ["file"]),
    ):
        record = dict(id=identifier, app_id="demo", page_id="manual-main", region_id=static.removeprefix("manual-main."),
                      name="File" if identifier == "popup" else identifier.title(), control_type=kind,
                      location=location_from_xpath(path), navigation_path=navigation, status="verified", normalized_rect=n,
                      image=context, visual_observation=dict(context_image=context, normalized_rect=n,
                          context_size=dict(width=400, height=300), main_rect=dict(left=0, top=0, right=1, bottom=1),
                          purpose="learning-reference"))
        knowledge.save_control(directory, record)
        records.append(record)
    store.sync_dynamic_regions("demo")
    return store, records, static


def get_graph(store):
    index = store.read("demo")
    return dynamic_regions.build_tree(index["controls"], index["regions"])


def test_popup_becomes_region_and_native_children_follow_it_without_changing_replay(popup):
    store, records, static = popup
    original = deepcopy(store.read("demo")["controls"])
    graph = get_graph(store)
    region_id = graph["containers"]["popup"]
    assert graph["parents"]["region:"+region_id] == "control:file"
    assert graph["parents"]["control:new"] == graph["parents"]["control:open"] == "region:"+region_id
    assert "control:popup" not in graph["parents"]
    assert graph["parents"]["control:file"] == "region:"+static
    assert [r["id"] for r in store.click_plan("demo", "new")] == ["file", "new"]
    store.sync_dynamic_regions("demo")
    assert store.read("demo")["controls"] == original
    assert len(get_graph(store)["dynamic"]) == 1
    region = next(r for r in store.read("demo")["regions"] if r["id"] == region_id)
    assert region["image"] == records[1]["image"] and region["name"] == "File 子菜单"
    assert len(store.layout_regions("demo")) == 1  # Dynamic overlap is a different state.


def test_automatic_per_command_wrappers_are_flattened_into_the_real_popup(popup):
    store, records, _static = popup
    directory = store.directory("demo")
    for command, label in ((records[2], "创建"), (records[3], "打开")):
        path = command["location"]["Xpath"]
        wrapper_step = dict(ControlType="MenuControl", Name=label)
        wrapper = {**records[1], "id": label, "name": label,
            "visual_observation": {**command["visual_observation"]},
            "location": location_from_xpath([*path[:-1], wrapper_step])}
        knowledge.save_control(directory, wrapper)
        knowledge.save_control(directory, {**command, "location": location_from_xpath([*path[:-1], wrapper_step, path[-1]])})
    original = deepcopy(store.read("demo")["controls"])
    store.sync_dynamic_regions("demo")
    graph = get_graph(store)
    root = graph["containers"]["popup"]
    assert len(graph["dynamic"]) == 1
    assert graph["containers"]["创建"] == graph["containers"]["打开"] == root
    assert graph["parents"]["control:new"] == graph["parents"]["control:open"] == "region:"+root
    assert store.read("demo")["controls"] == original
    assert [r["id"] for r in store.click_plan("demo", "new")] == ["file", "new"]
    store.sync_dynamic_regions("demo")
    assert len(get_graph(store)["dynamic"]) == 1


def test_duplicate_wrapper_region_is_automatically_retired_but_manual_group_is_kept(popup):
    store, records, _static = popup
    directory = store.directory("demo")
    root = get_graph(store)["containers"]["popup"]
    path = records[1]["location"]["Xpath"]
    step = dict(ControlType="MenuControl", Name="Duplicate")
    wrapper = {**records[1], "id": "duplicate", "name": "Duplicate", "location": location_from_xpath([*path, step])}
    knowledge.save_control(directory, wrapper)
    for command in records[2:]:
        knowledge.save_control(directory, {**command, "location": location_from_xpath([*path, step, command["location"]["Xpath"][-1]])})
    knowledge.save_region(directory, {**get_graph(store)["dynamic"][root], "id": "old-duplicate", "source_control_id": "duplicate"})
    manual = store.save_dynamic_region("demo", "人工分组", dict(left=.1, top=.2, right=.6, bottom=.35), parent_region_id="old-duplicate")
    store.sync_dynamic_regions("demo")
    assert "old-duplicate" not in {r["id"] for r in store.read("demo")["regions"]}
    assert list((directory / ".easy_uiauto" / "archive").rglob("old-duplicate.md"))
    assert manual in get_graph(store)["dynamic"]
    assert len(get_graph(store)["dynamic"]) == 2
    assert get_graph(store)["parents"]["region:"+manual] == "region:"+root


def test_a_real_single_item_popup_keeps_its_region_in_a_different_expanded_state(popup):
    store, records, _static = popup
    directory = store.directory("demo")
    path = [*records[1]["location"]["Xpath"], dict(ControlType="MenuControl", Name="Submenu")]
    source = {**records[1], "id": "submenu", "navigation_path": ["file", "new"], "location": location_from_xpath(path)}
    leaf = {**records[2], "id": "only-item", "navigation_path": ["file", "new"],
            "location": location_from_xpath([*path, dict(ControlType="MenuItemControl", Name="Only")])}
    for r in (source, leaf): knowledge.save_control(directory, r)
    store.sync_dynamic_regions("demo")
    graph = get_graph(store)
    assert len(graph["dynamic"]) == 2
    assert graph["parents"]["region:"+graph["containers"]["submenu"]] == "control:new"
    assert graph["parents"]["control:only-item"] == "region:"+graph["containers"]["submenu"]


def test_subregions_can_nest_and_only_partition_the_same_learned_frame(popup):
    store, _records, _static = popup
    root = get_graph(store)["containers"]["popup"]
    upper = store.save_dynamic_region("demo", "创建", dict(left=.1, top=.2, right=.6, bottom=.35), parent_region_id=root)
    lower = store.save_dynamic_region("demo", "打开", dict(left=.1, top=.35, right=.6, bottom=.6), parent_region_id=root)
    nested = store.save_dynamic_region("demo", "常用", dict(left=.12, top=.23, right=.58, bottom=.34), parent_region_id=upper)
    graph = get_graph(store)
    assert graph["parents"]["control:new"] == "region:"+nested
    assert graph["parents"]["control:open"] == "region:"+lower
    assert graph["parents"]["region:"+nested] == "region:"+upper
    assert graph["parents"]["region:"+upper] == "region:"+root
    region = graph["dynamic"][nested]
    with Image.open(store.asset("demo", region["image"])) as crop:
        assert crop.size == (184, 33) and crop.getpixel((0, 0)) == (0, 0, 255)
    knowledge.save_control(store.directory("demo"), {**store.read("demo")["controls"][0],
        "id": "outside", "name": "Other", "navigation_path": [],
        "location": location_from_xpath([dict(ControlType="WindowControl", Name="Test"), dict(ControlType="ButtonControl", Name="Other")])})
    assert get_graph(store)["parents"]["control:outside"].startswith("region:manual-main.")


def test_wrong_scope_or_navigation_never_attaches_to_same_named_popup(popup):
    store, records, _static = popup
    for field, value in (("navigation_path", []), ("window_scope", dict(kind="owned", title="Other", class_name="#32768"))):
        other = {**records[2], "id": "other", field: value}
        knowledge.save_control(store.directory("demo"), other)
        assert get_graph(store)["parents"]["control:other"] != "region:"+get_graph(store)["containers"]["popup"]


def test_multilevel_popup_regions_attach_to_their_own_entry_and_survive_missing_source(popup):
    store, records, _static = popup
    directory = store.directory("demo")
    recent = {**records[2], "id": "recent", "name": "Recent"}
    next_menu = dict(ControlType="MenuControl", Name="Recent", foundIndex=2)
    path = [dict(ControlType="WindowControl", Name="Test"), next_menu]
    second = {**records[1], "id": "second", "name": "Recent", "navigation_path": ["file", "recent"],
              "location": location_from_xpath(path)}
    leaf = {**records[2], "id": "project", "navigation_path": ["file", "recent"],
            "location": location_from_xpath([*path, dict(ControlType="MenuItemControl", Name="Project")])}
    for r in (recent, second, leaf): knowledge.save_control(directory, r)
    store.sync_dynamic_regions("demo")
    graph = get_graph(store)
    second_region = graph["containers"]["second"]
    assert graph["parents"]["region:"+second_region] == "control:recent"
    assert graph["parents"]["control:project"] == "region:"+second_region
    store.archive("demo", "control", "second")
    graph = get_graph(store)
    assert "second" not in graph["containers"]
    assert graph["parents"]["control:project"] == "control:recent"


def test_explicit_container_conversion_and_manual_parent_override(popup):
    store, records, static = popup
    panel = {**records[1], "id": "panel", "control_type": "PaneControl", "name": "Options",
             "location": location_from_xpath([dict(ControlType="WindowControl", Name="Test"),
                                                dict(ControlType="PaneControl", Name="Options")])}
    knowledge.save_control(store.directory("demo"), panel)
    store.sync_dynamic_regions("demo")
    assert "panel" not in get_graph(store)["containers"]
    region = store.make_dynamic_region("demo", "panel")
    assert get_graph(store)["containers"]["panel"] == region
    with pytest.raises(ValueError, match="容器"):
        store.make_dynamic_region("demo", "new")
    store.update_control("demo", "new", name="New", description="", aliases=[], hierarchy_parent="")
    assert get_graph(store)["parents"]["control:new"] == "region:"+static
    assert knowledge.merge_control_state({**panel, "region_behavior": "region"}, panel)["region_behavior"] == "region"


def test_different_reference_frame_does_not_reuse_old_subdivision(popup):
    store, records, _static = popup
    root = get_graph(store)["containers"]["popup"]
    child = store.save_dynamic_region("demo", "创建", dict(left=.1, top=.2, right=.6, bottom=.35), parent_region_id=root)
    directory = store.directory("demo")
    next_frame = image_storage.save_interaction_image(directory, Image.new("RGB", (400, 300), "green"))
    for r in (records[1], records[2]):
        knowledge.save_control(directory, {**r, "visual_observation": {**r["visual_observation"], "context_image": next_frame}})
    store.sync_dynamic_regions("demo")
    assert get_graph(store)["parents"]["control:new"] == "region:"+root
    assert child in get_graph(store)["dynamic"]


def test_overlapping_outside_and_parent_shrinking_are_rejected_without_partial_writes(popup):
    store, _records, _static = popup
    root = get_graph(store)["containers"]["popup"]
    child = store.save_dynamic_region("demo", "创建", dict(left=.1, top=.2, right=.6, bottom=.35), parent_region_id=root)
    before = deepcopy(store.read("demo")["regions"])
    for rect in (dict(left=.1, top=.3, right=.6, bottom=.5), dict(left=0, top=0, right=.6, bottom=.6)):
        with pytest.raises(ValueError): store.save_dynamic_region("demo", "错误", rect, parent_region_id=root)
    assert store.read("demo")["regions"] == before
    nested = store.save_dynamic_region("demo", "子", dict(left=.12, top=.21, right=.58, bottom=.34), parent_region_id=child)
    with pytest.raises(ValueError, match="排除"):
        store.save_dynamic_region("demo", "创建", dict(left=.2, top=.2, right=.6, bottom=.35), region_id=child)
    assert get_graph(store)["parents"]["region:"+nested] == "region:"+child


def test_rename_and_archive_do_not_change_learning_pixels_or_resurrect_deleted_region(popup):
    store, records, _static = popup
    root = get_graph(store)["containers"]["popup"]
    original = deepcopy(store.read("demo")["controls"])
    store.save_dynamic_region("demo", "文件操作", records[1]["normalized_rect"], region_id=root)
    store.sync_dynamic_regions("demo")
    assert get_graph(store)["dynamic"][root]["name"] == "文件操作"
    child = store.save_dynamic_region("demo", "创建", dict(left=.1, top=.2, right=.6, bottom=.35), parent_region_id=root)
    store.archive("demo", "region", root)
    store.sync_dynamic_regions("demo")
    assert "popup" not in get_graph(store)["containers"]
    assert get_graph(store)["parents"]["region:"+child] == "control:file"
    updated = store.read("demo")["controls"]
    for old in original:
        current = next(r for r in updated if r["id"] == old["id"])
        assert all(current[k] == old[k] for k in ("location", "navigation_path", "image", "status"))


def test_gui_region_node_and_drawing_preserve_fixed_reference_and_click_path(popup, monkeypatch):
    import tkinter as tk
    from easy_uiauto.knowledge_gui import KnowledgeWindow

    store, records, _static = popup
    root = tk.Tk(); root.withdraw()
    gui = None
    try:
        gui = KnowledgeWindow(root, store.root)
        region_id = get_graph(store)["containers"]["popup"]
        key = "region:demo:"+region_id
        assert gui.tree.parent(key) == "control:demo:file"
        assert gui.tree.parent("control:demo:new") == key
        assert not gui.tree.exists("control:demo:popup")
        gui.tree.selection_set(key); gui.select()
        assert gui.detail_title.cget("text") == "动态子区域"
        assert gui.viewer.image.getpixel((0, 0)) == (0, 0, 255)
        assert "File" in gui.navigation_summary.cget("text")
        assert len(gui.viewer.controls) == 2
        gui.observation_mode.set("本次执行结果"); gui.show_observation()
        assert gui.observation_mode.get() == "学习参考图"
        calls = []
        monkeypatch.setattr(gui, "start_task", lambda *a, **kw: calls.append((a, kw)))
        gui.click()
        assert not calls  # A region is not an additional input action.
        gui.draw_mode.set(True); gui.toggle_drawing()
        assert gui.draw_parent_region == region_id and gui.viewer.drawing
        gui.drawn(dict(left=.1, top=.2, right=.6, bottom=.35))
        child = gui.selected[2]["id"]
        assert gui.tree.parent("region:demo:"+child) == key
        assert gui.tree.parent("control:demo:new") == "region:demo:"+child
        assert gui.viewer.image.getpixel((0, 0)) == (0, 0, 255)
        gui.tree.selection_set("control:demo:new"); gui.select()
        assert gui.selected[2]["location"] == records[2]["location"]
        assert gui.navigation_ids == ["file"]
    finally:
        if gui: gui.dirty = False; gui.close()
        else: root.destroy()
