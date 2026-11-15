"""Qt menu items must be current, scoped knowledge even without point hits."""
from copy import deepcopy

import pytest
from PIL import Image

from easy_uiauto import library_actions, region_learning, utils, workflows
from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.mcp import control_observation, dynamic_regions, knowledge, local_learning, scanner
from test_manual_region_scan import setup
from test_mcp_scanner import FakeControl


NAMES = ["新建项目", "打开项目", "最近打开项目", "导入", "系统设置"]


@pytest.fixture(params=["MenuControl", "WindowControl"])
def popup(setup, monkeypatch, request):
    layout, native, root = setup
    main = scanner.uiautomation.ControlFromHandle(1)
    items = [FakeControl(name, "MenuItemControl", (355, 55+i*28, 565, 79+i*28))
             for i, name in enumerate(NAMES)]
    menu = FakeControl("menu_文件", request.param, (350, 50, 570, 210), items, parent=main)
    menu.ClassName = "Qt5152QWindowPopupSaveBits"
    menu.NativeWindowHandle = 2
    for item in items:
        item.NativeWindowHandle = 0
    popup_native = dict(handle=2, title="", class_name=menu.ClassName, process_id=123,
                        owner_handle=1, visible=True, enabled=True)
    monkeypatch.setattr(scanner.native_windows, "matching_windows",
                        lambda _name, **kw: [popup_native if kw.get("handle") == 2 else native])
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", lambda h: menu if h == 2 else main)
    monkeypatch.setattr(library_actions, "owned_windows", lambda _e: [popup_native])
    monkeypatch.setattr(library_actions, "bind_window", lambda _e: main)
    monkeypatch.setattr(control_observation, "_check_foreground", lambda _w: None)
    monkeypatch.setattr(scanner, "_window_rect", scanner._rect)
    monkeypatch.setattr(scanner, "_screenshot_window",
                        lambda r: Image.new("RGB", (r["width"], r["height"]), "white"))
    monkeypatch.setattr(local_learning, "detect", lambda *a, **kw: ([], []))
    monkeypatch.setattr(utils, "get_control_info", lambda *a, **kw: pytest.fail("menu needs no point hit"))
    monkeypatch.setattr(scanner, "_walk_controls", lambda *a, **kw: pytest.fail("no full-tree scan"))
    scope = dict(kind="owned", title="", class_name=menu.ClassName)
    return layout, native, root, main, menu, items, popup_native, scope


def scan_menu(popup, *, layout_overrides=None, **kwargs):
    layout, native, root, _main, _menu, _items, popup_native, scope = popup
    layout = {**deepcopy(layout), "window_handle": 2, "window_class": popup_native["class_name"],
              "window_scope": scope, "image_size": dict(width=220, height=160),
              "regions": [dict(id="menu", name="文件菜单", normalized_rect=dict(left=0, top=0, right=1, bottom=1))]}
    layout.update(layout_overrides or {})
    return scanner.scan_window("", "", "", "", "test", root=root, manual_layout=layout,
                               semantic_mode="native", observation_owner=native, **kwargs)


def current_records(root, result):
    snapshot = KnowledgeLibrary(root).read("chosen-app")
    return [r for r in snapshot["controls"] if r["id"] in result["observed_control_ids"]]


def test_owned_qt_menu_has_five_current_items_and_stable_ids_without_detections(popup):
    result = scan_menu(popup)
    records = current_records(popup[2], result)
    assert {r["name"] for r in records} == set(NAMES)
    assert result["coverage"]["uia_enumerated"] == 5 and not result["truncated"]
    for record in records:
        assert record["window_scope"] == popup[-1]
        assert record["location"]["Xpath"][0]["Name"] == "menu_文件"
        assert record["location"]["Xpath"][-1]["Name"] == record["name"]
        assert "click" in record["actions"]
        assert record["detection"]["method"] == "popup-menu-uia"
    assert scan_menu(popup)["observed_control_ids"] == result["observed_control_ids"]


def test_unobserved_or_hidden_menu_items_are_not_imported_from_history(popup):
    first = scan_menu(popup)
    menu, items = popup[4:6]
    menu._children = items[:3]
    items[1].IsOffscreen = True
    second = scan_menu(popup)
    assert {r["name"] for r in current_records(popup[2], second)} == {NAMES[0], NAMES[2]}
    assert len(KnowledgeLibrary(popup[2]).read("chosen-app")["controls"]) == len(first["observed_control_ids"])


def test_menu_limit_and_provider_failure_do_not_claim_full_coverage(popup):
    result = scan_menu(popup, max_controls=2)
    assert result["truncated"] and len(result["observed_control_ids"]) == 2
    def disconnected():
        raise RuntimeError("menu provider disconnected")
    popup[4].GetChildren = disconnected
    with pytest.raises(RuntimeError, match="provider disconnected"):
        scan_menu(popup)


def test_nested_menu_containers_keep_duplicate_name_identity_and_exclude_closed_submenus(popup):
    menu, items = popup[4:6]
    items[1].Name = items[0].Name
    hidden_child = FakeControl("未展开的子项", "MenuItemControl", (360, 60, 540, 75))
    items[2]._children = [hidden_child]
    hidden_child._parent = items[2]
    wrapper = FakeControl("", "PaneControl", (350, 50, 570, 210), items, parent=menu)
    menu._children = [wrapper]
    result = scan_menu(popup)
    records = current_records(popup[2], result)
    assert len(records) == 5 and not any(r["name"] == hidden_child.Name for r in records)
    same_names = [r for r in records if r["name"] == items[0].Name]
    assert len(same_names) == 2
    assert {r["location"]["Xpath"][-1]["foundIndex"] for r in same_names} == {1, 2}


def test_saved_popup_locator_resolves_in_owned_window_without_cross_window_fallback(popup, monkeypatch):
    from easy_uiauto import popup_guard
    records = current_records(popup[2], scan_menu(popup))
    monkeypatch.setattr(popup_guard, "assert_unblocked", lambda *a: None)
    for record in records:
        control = library_actions.resolve_record(record, popup[1], popup[3])
        assert control is next(i for i in popup[5] if i.Name == record["name"])


def test_region_restriction_and_foreign_popup_are_still_enforced(popup, monkeypatch):
    result = scan_menu(popup, layout_overrides=dict(restrict_regions=True, regions=[dict(
        id="menu", name="菜单上部", normalized_rect=dict(left=0, top=0, right=1, bottom=.35))]))
    assert {r["name"] for r in current_records(popup[2], result)} == set(NAMES[:2])
    monkeypatch.setattr(library_actions, "owned_windows", lambda _: [])
    with pytest.raises(ValueError, match="不属于"):
        scan_menu(popup)


@pytest.mark.parametrize("previously_seen", [False, True])
def test_expanded_popup_items_attach_under_opener_and_keep_popup_relative_coordinates(popup, previously_seen):
    layout, _native, root, _main, _menu, _items, _popup_native, scope = popup
    before_ids = scan_menu(popup)["observed_control_ids"] if previously_seen else []
    directory = knowledge.initialize_app("chosen-app", "Example", root)
    opener = dict(id="file-entry", name="文件", app_id="chosen-app", page_id=layout["page_id"],
                  region_id="toolbar", control_type="MenuItemControl", status="observed",
                  location={"Xpath": [{"ControlType": "WindowControl", "Name": "Example"},
                                      {"ControlType": "MenuItemControl", "Name": "文件"}]})
    knowledge.save_control(directory, opener)
    result = scanner.scan_window("Example", "", "", "", "test", root=root, manual_layout=layout,
        semantic_mode="native", navigation_context=dict(path=[opener["id"]], popup_handles=[2], before_ids=[opener["id"], *before_ids]))
    snapshot = KnowledgeLibrary(root).read("chosen-app")
    tree = dynamic_regions.build_tree(snapshot["controls"], snapshot["regions"])
    region = next(r for r in snapshot["regions"] if r.get("source") == "local-visual-container")
    assert region["window_scope"] == scope
    assert tree["parents"]["region:"+region["id"]] == "control:"+opener["id"]
    records = current_records(root, result)
    assert {r["name"] for r in records} == set(NAMES)
    for record in records:
        assert record["window_scope"] == scope and record["navigation_path"] == [opener["id"]]
        assert tree["parents"]["control:"+record["id"]] == "region:"+region["id"]
        assert record["detection"]["normalized_rect"]["left"] == pytest.approx(355/400)


def test_deep_observation_planner_and_tree_receive_current_popup_items(popup, monkeypatch, tmp_path):
    layout, native, root, _main, menu, _items, popup_native, scope = popup
    # Seed the entry in its main-page region; the menu is a separate state page.
    scan_menu(popup)
    library = KnowledgeLibrary(root, page_id=layout["page_id"])
    origin = library.regions("chosen-app")[0]
    entry = {**library.read("chosen-app")["controls"][0], "id": "file-entry", "name": "文件",
             "window_scope": {"kind": "main"}, "region_id": origin["id"]}
    knowledge.save_control(library.directory("chosen-app"), entry)
    monkeypatch.setattr(region_learning.configuration, "load_gui_configuration",
                        lambda: dict(api_url="https://example.invalid/v1", api_key="test", api_model="test"))
    backend = region_learning.NativeBackend(library, "chosen-app", origin, native, tmp_path)
    frame = dict(state_id="file-menu", image=Image.new("RGB", (220,160), "white"), target=popup_native,
                 scope=scope, bounds=scanner._rect(menu), root=menu, issues=[])
    monkeypatch.setattr(backend, "stable_frame", lambda: frame)
    observation = backend.observe(route=[workflows.click_step(entry)])
    assert {r["name"] for r in observation["controls"]} == set(NAMES)
    assert all(r["window_scope"] == scope for r in observation["controls"])
    def plan(_settings, payload, _image):
        assert {r["name"] for r in payload["controls"]} == set(NAMES)
        return dict(branches=[dict(name=r["name"], effect="navigation", replayable=True,
                                  steps=[workflows.click_step(r)]) for r in payload["controls"]])
    monkeypatch.setattr(region_learning, "_request_plan", plan)
    planned = backend.plan(observation, [])
    assert len(planned["branches"]) == 5
    assert not any(b.get("blocked") for b in planned["branches"])
    snapshot = library.read("chosen-app")
    tree = dynamic_regions.build_tree(snapshot["controls"], snapshot["regions"])
    for record in observation["controls"]:
        parent = tree["parents"]["control:"+record["id"]]
        assert tree["parents"][parent] == "control:file-entry"
        assert library.entry_steps("chosen-app", record)[0]["target"]["control_id"] == entry["id"]


def test_qt_tree_places_deep_popup_region_under_entry_without_duplicate_root():
    pytest.importorskip("PySide6")
    from easy_uiauto.qt_ui.models import KnowledgeModel
    region = dict(id="main.toolbar", page_id="main", name="工具栏")
    popup_region = dict(id="popup.menu", page_id="popup", name="文件菜单", learning_source="deep-learning",
                        window_scope=dict(kind="owned", class_name="Qt5152QWindowPopupSaveBits", title=""),
                        parent_control_id="file")
    entry = dict(id="file", page_id="main", region_id=region["id"], name="文件", control_type="MenuItemControl")
    child = dict(id="new", page_id="popup", region_id=popup_region["id"], name="新建", control_type="MenuItemControl")
    snapshot = dict(application=dict(name="Demo"), usage=[], workflows=[],
                    pages=[dict(id="main"), dict(id="popup")], regions=[region, popup_region], controls=[entry, child])
    rows = KnowledgeModel.describe("Vault", {"demo": snapshot}, {})
    keys = [row[1] for row in rows]
    by_key = {row[1]: row for row in rows}
    assert keys.count("region:demo:popup.menu") == 1
    assert by_key["region:demo:popup.menu"][0] == "control:demo:file"
    assert by_key["control:demo:new"][0] == "region:demo:popup.menu"
    assert keys.index("control:demo:file") < keys.index("region:demo:popup.menu") < keys.index("control:demo:new")
