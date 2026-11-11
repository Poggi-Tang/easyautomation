import pytest

from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.mcp import knowledge, scanner
from test_mcp_scanner import FakeControl, _test_screen


@pytest.fixture
def setup(monkeypatch, tmp_path):
    button = FakeControl("Save", "ButtonControl", (40, 30, 160, 90), automation_id="save")
    label = FakeControl("Heading", "TextControl", (20, 220, 180, 250))
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [button, label])
    native = dict(handle=1, visible=True, title="Example", class_name="WindowControl", process_id=123)
    monkeypatch.setattr(scanner.native_windows, "matching_windows", lambda *_a, **_kw: [native])
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", lambda _h: window)
    monkeypatch.setattr(scanner, "_activate_window", lambda _w: None)
    monkeypatch.setattr(scanner, "_window_rect", lambda _w: dict(left=0, top=0, right=400, bottom=300, width=400, height=300))
    monkeypatch.setattr(scanner, "_process_name", lambda _p: "test-process")
    monkeypatch.setattr(scanner, "_screenshot_window", lambda _r: _test_screen())
    monkeypatch.setattr(scanner, "get_control_xpath", lambda c: [
        dict(ControlType="WindowControl", Name="Example"), dict(ControlType=c.ControlTypeName, Name=c.Name)])
    monkeypatch.setattr(scanner, "_verify_location", lambda _l, r, *_a: (True, r, "same control"))
    def forbidden(*_a, **_kw): raise AssertionError("Remote segmentation was called")
    monkeypatch.setattr(scanner, "segment_interface", forbidden)
    monkeypatch.setattr(scanner, "_trim_window_capture", forbidden)
    monkeypatch.setattr(scanner, "analyze_control_semantics", forbidden)
    layout = {"app_id": "chosen-app", "page_id": "manual-main", "window_handle": 1,
              "window_class": "WindowControl", "window_process_id": 123,
              "image_size": dict(width=400, height=300),
              "regions": [dict(id="toolbar", name="工具栏", normalized_rect=dict(left=0, top=0, right=1, bottom=.5))]}
    return layout, native, tmp_path


def run(layout, root, **kwargs):
    return scanner.scan_window("Example", "", "", "", "test", root=root,
                               manual_layout=layout, semantic_mode="native", strategy="full-uia", **kwargs)


def test_manual_scan_uses_exact_window_and_regions_without_api(setup):
    layout, _native, root = setup
    result = run(layout, root)
    assert result["strategy"] == "manual-regions"
    assert result["app_id"] == "chosen-app"
    assert result["controls_saved_this_scan"] == 2 and result["unassigned_controls"] == 1
    index = knowledge.load_index(knowledge.app_dir("chosen-app", root))
    controls = {r["name"]: r for r in index["controls"]}
    assert controls["Save"]["region_id"] == "toolbar"
    assert controls["Heading"]["region_id"] == "unassigned"
    assert len(index["regions"]) == 2
    assert controls["Save"]["location"]["Xpath"][-1]["Name"] == "Save"
    assert all(r["control_type"] != "WindowControl" for r in index["controls"])


def test_repeat_learning_updates_instead_of_duplicate_records(setup):
    layout, _native, root = setup
    first = run(layout, root)
    second = run(layout, root)
    assert first["observed_control_ids"] == second["observed_control_ids"]
    assert second["knowledge_controls_total"] == 2


def test_expanded_menu_container_saves_dynamic_region_during_learning(setup, monkeypatch):
    layout, _native, root = setup
    run(layout, root)
    store = KnowledgeLibrary(root)
    opener = next(r for r in store.read("chosen-app")["controls"] if r["name"] == "Save")
    main = scanner.uiautomation.ControlFromHandle(1)
    before_keys = {scanner._control_key(c) for c, _d in scanner._walk_controls(main, 12, 3000)}
    leaf = FakeControl("New", "MenuItemControl", (40, 100, 140, 125))
    menu = FakeControl("File", "MenuControl", (30, 95, 150, 160), [leaf])
    main._children.append(menu)
    menu._parent = main
    monkeypatch.setattr(scanner, "get_control_xpath", lambda c: (
        [dict(ControlType="WindowControl", Name="Example"), dict(ControlType="MenuControl", Name="File")]
        + ([dict(ControlType="MenuItemControl", Name="New")] if c is leaf else [])
        if c is leaf or c is menu else [dict(ControlType="WindowControl", Name="Example"), dict(ControlType=c.ControlTypeName, Name=c.Name)]))
    result = run(layout, root, navigation_context=dict(path=[opener["id"]], before_keys=before_keys, popup_handles=[]))
    assert not result["failures"]
    index = store.read("chosen-app")
    from easy_uiauto.mcp.dynamic_regions import build_tree
    graph = build_tree(index["controls"], index["regions"])
    container = next(r for r in index["controls"] if r["control_type"] == "MenuControl")
    new = next(r for r in index["controls"] if r["name"] == "New")
    region = graph["containers"][container["id"]]
    assert graph["parents"]["control:"+new["id"]] == "region:"+region
    assert graph["parents"]["region:"+region] == "control:"+opener["id"]


def test_wrong_size_fails_before_creating_application(setup):
    layout, _native, root = setup
    layout["image_size"]["width"] = 800
    with pytest.raises(ValueError, match="尺寸"): run(layout, root)
    assert not (root / "applications").exists()


def test_changed_native_window_identity_is_not_rebound(setup):
    layout, native, root = setup
    native["process_id"] = 999
    with pytest.raises(ValueError, match="失效"): run(layout, root)
    assert not (root / "applications").exists()


def test_gui_layout_metadata_survives_scan(setup):
    layout, _native, root = setup
    store = KnowledgeLibrary(root)
    store.bind("chosen-app", "自定义软件名", dict(title="Example", class_name="WindowControl"))
    screenshot = root / "source.png"; _test_screen().save(screenshot)
    store.save_capture("chosen-app", dict(image=str(screenshot), width=400, height=300, window_name="Example", window_class="WindowControl"))
    store.save_region("chosen-app", "工具栏", dict(left=0, top=0, right=1, bottom=.5), "toolbar")
    store.confirm_layout("chosen-app")
    run(layout, root)
    assert store.page("chosen-app")["layout_image"].endswith("-layout.png")
    assert store.page("chosen-app")["layout_confirmed"] is True
    assert store.read("chosen-app")["application"]["name"] == "自定义软件名"


def test_learning_retains_auto_origin_and_manual_edits_for_later_repartition(setup):
    from easy_uiauto.mcp.automatic_layout import propose

    _layout, native, root = setup
    store = KnowledgeLibrary(root)
    store.bind("chosen-app", "Demo", dict(title="Example", class_name="WindowControl"))
    screenshot = root / "source.png"; _test_screen().save(screenshot)
    store.save_capture("chosen-app", dict(image=str(screenshot), width=400, height=300, window_name="Example", window_class="WindowControl"))
    proposal = propose(_test_screen())
    store.apply_automatic_layout("chosen-app", proposal)
    automatic = store.layout_regions("chosen-app")
    edited = max(automatic, key=lambda r: (r["normalized_rect"]["right"]-r["normalized_rect"]["left"])
                 *(r["normalized_rect"]["bottom"]-r["normalized_rect"]["top"]))
    store.save_region("chosen-app", "我的工作区", edited["normalized_rect"], edited["id"])
    store.confirm_layout("chosen-app")
    layout = store.learning_layout("chosen-app", native)
    run(layout, root)
    after = store.layout_regions("chosen-app")
    assert all(r.get("layout_origin", "").startswith("auto-") for r in after)
    record = next(r for r in after if r["id"] == edited["id"])
    assert record["edited_by_user"] and record["name"] == "我的工作区" and record["image"]
    result = store.apply_automatic_layout("chosen-app", proposal)
    assert result["preserved"] == 1
    assert next(r for r in store.layout_regions("chosen-app") if r["id"] == edited["id"])["name"] == "我的工作区"


def test_traversal_limit_still_reports_truncation_after_excluding_window_root(setup):
    layout, _native, root = setup
    assert run(layout, root, max_controls=2)["truncated"] is True


def test_optional_semantics_receives_manual_regions_without_resegmenting(setup, monkeypatch):
    from easy_uiauto.mcp.manual_layout import native_semantics

    layout, _native, root = setup
    calls = []
    def semantics(_image, candidates, _page, regions, *_args):
        calls.append([r["name"] for r in regions])
        return {c["id"]: native_semantics(c) for c in candidates}
    monkeypatch.setattr(scanner, "analyze_control_semantics", semantics)
    result = scanner.scan_window("Example", "https://example.invalid", "key", "model", "test",
                                 manual_layout=layout, semantic_mode="api", root=root, strategy="full-uia")
    assert result["controls_saved_this_scan"] == 2
    assert calls == [["工具栏"]]


def test_rescan_preserves_user_selected_position_mode(setup):
    layout, _native, root = setup
    run(layout, root)
    store = KnowledgeLibrary(root)
    record = next(r for r in store.read("chosen-app")["controls"] if r["name"] == "Save")
    store.update_control("chosen-app", record["id"], name="保存", description="", aliases=[],
                         location=record["location"], match_mode="position")
    run(layout, root)
    updated = next(r for r in store.read("chosen-app")["controls"] if r["id"] == record["id"])
    assert updated["location"]["PARAMETERS"]["match_mode"] == "position"


def test_owned_popup_outside_main_window_keeps_full_crop_and_navigation(setup, monkeypatch):
    from PIL import Image
    from easy_uiauto import library_actions

    layout, _native, root = setup
    run(layout, root)
    store = KnowledgeLibrary(root)
    parent = next(c for c in store.read("chosen-app")["controls"] if c["name"] == "Save")
    main = scanner.uiautomation.ControlFromHandle(1)
    new = FakeControl("New", "ButtonControl", (420, 30, 500, 80), automation_id="new")
    popup = FakeControl("Menu", "WindowControl", (410, 20, 520, 100), [new])
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", lambda h: popup if h == 2 else main)
    monkeypatch.setattr(library_actions, "owned_windows", lambda _e: [dict(handle=2, title="Menu", class_name="#32768")])
    monkeypatch.setattr(scanner, "_screenshot_window", lambda r: Image.new("RGB", (r["width"], r["height"]), "white"))
    context = dict(path=[parent["id"]], before_keys={scanner._control_key(c) for c, _d in scanner._walk_controls(main, 12, 3000)}, popup_handles=[2])
    result = run(layout, root, navigation_context=context)
    assert not result["failures"], result["failures"]
    record = next(c for c in store.read("chosen-app")["controls"] if c["name"] == "New")
    assert record["navigation_path"] == [parent["id"]]
    assert record["window_scope"] == dict(kind="owned", class_name="#32768", title="Menu")
    assert record["region_id"] == "unassigned"
    with Image.open(store.asset("chosen-app", record["image"])) as crop:
        assert crop.size == (80, 50)
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", lambda _h: main)
    run(layout, root)
    retained = next(c for c in store.read("chosen-app")["controls"] if c["id"] == record["id"])
    assert retained["navigation_path"] == [parent["id"]]
    assert retained["status"] != "quarantined"


def test_expanded_scan_does_not_replace_clicked_entry_with_after_pixels(setup, monkeypatch):
    from PIL import Image
    from easy_uiauto import library_actions

    layout, _native, root = setup
    run(layout, root)
    store = KnowledgeLibrary(root)
    previous = next(r for r in store.read("chosen-app")["controls"] if r["name"] == "Save")
    main = scanner.uiautomation.ControlFromHandle(1)
    monkeypatch.setattr(library_actions, "owned_windows", lambda _e: [])
    monkeypatch.setattr(scanner, "_screenshot_window", lambda r: Image.new("RGB", (r["width"], r["height"]), "white"))
    result = run(layout, root, navigation_context=dict(path=[previous["id"]], before_keys=set(), popup_handles=[]))
    current = next(r for r in store.read("chosen-app")["controls"] if r["id"] == previous["id"])
    assert not result["failures"]
    assert current["image_sha256"] == previous["image_sha256"]
    assert current["visual_observation"] == previous["visual_observation"]
