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


@pytest.fixture(params=["MenuControl", "WindowControl", "PaneControl"])
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
        assert record["region_id"] == region["id"]
        assert record["command"].startswith(layout["page_id"]+"."+region["id"]+".")
        assert tree["parents"]["control:"+record["id"]] == "region:"+region["id"]
        assert record["detection"]["normalized_rect"]["left"] == pytest.approx(355/400)


def test_deep_observation_planner_and_tree_receive_current_popup_items(popup, monkeypatch, tmp_path):
    layout, native, root, _main, menu, _items, popup_native, scope = popup
    # A popup observation must use the original page without replacing its layout.
    scanner.scan_window("Example", "", "", "", "test", root=root, manual_layout=layout, semantic_mode="native")
    library = KnowledgeLibrary(root, page_id=layout["page_id"])
    original_page = library.page("chosen-app")
    assert not original_page.get("layout_confirmed")
    origin = library.regions("chosen-app")[0]
    entry = {"id": "file-entry", "name": "文件", "app_id":"chosen-app", "page_id":layout["page_id"],
             "control_type":"MenuItemControl", "status":"observed",
             "window_scope": {"kind": "main"}, "region_id": origin["id"]}
    knowledge.save_control(library.directory("chosen-app"), entry)
    monkeypatch.setattr(region_learning.configuration, "load_gui_configuration",
                        lambda: dict(api_url="https://example.invalid/v1", api_key="test", api_model="test"))
    backend = region_learning.NativeBackend(library, "chosen-app", origin, native, tmp_path)
    backend.configure_learning(dict(states={},root="",entry_control_id=entry["id"]),"new",[entry["id"]],[])
    frame = dict(state_id="file-menu", image=Image.new("RGB", (220,160), "white"), target=popup_native,
                 scope=scope, bounds=scanner._rect(menu), root=menu, issues=[])
    frame["visible_windows"]={backend.surface_key(popup_native)}
    backend.transition=dict(before_windows=set(),controls={entry["id"]},steps=[workflows.click_step(entry)])
    monkeypatch.setattr(backend, "stable_frame", lambda: frame)
    observation = backend.observe(route=[workflows.click_step(entry)])
    assert observation["state_id"] == frame["state_id"]
    assert {r["name"] for r in observation["controls"]} == set(NAMES)
    assert all(r["window_scope"] == scope for r in observation["controls"])
    def plan(_settings, payload, _image):
        assert {r["name"] for r in payload["controls"]} == set(NAMES)
        assert all(r["normalized_rect"]["left"]==pytest.approx(5/220,abs=1e-6) for r in payload["controls"])
        return dict(branches=[dict(name=r["name"], effect="navigation", replayable=True,
                                  steps=[workflows.click_step(r)]) for r in payload["controls"]])
    monkeypatch.setattr(region_learning, "_request_plan", plan)
    planned = backend.plan(observation, [])
    assert len(planned["branches"]) == 5
    assert not any(b.get("blocked") for b in planned["branches"])
    snapshot = library.read("chosen-app")
    assert len(snapshot["pages"]) == 1
    assert len(snapshot["controls"]) == 6
    tree = dynamic_regions.build_tree(snapshot["controls"], snapshot["regions"])
    for record in observation["controls"]:
        parent = tree["parents"]["control:"+record["id"]]
        assert tree["parents"][parent] == "control:file-entry"
        assert library.entry_steps("chosen-app", record)[0]["target"]["control_id"] == entry["id"]
        assert record["page_id"] == layout["page_id"]
        region = next(r for r in snapshot["regions"] if r["id"] == record["region_id"])
        assert region["parent_control_id"] == entry["id"]
        saved=next(r for r in snapshot["controls"] if r["id"]==record["id"])
        assert "normalized_rect" not in saved and "normalized_rect" not in record
        assert "normalized_rect" not in saved["detection"]
        assert "normalized_rect" not in region and "visual_observation" not in region
        assert region["location"]["Xpath"] and library.asset("chosen-app", region["image"]).is_file()
    assert library.page("chosen-app") == original_page
    repeat = scanner.scan_window("Example", "", "", "", "test", root=root, manual_layout=layout,
        semantic_mode="native", navigation_context=dict(path=[entry["id"]], popup_handles=[2]))
    assert set(repeat["observed_control_ids"]) == {r["id"] for r in observation["controls"]}
    assert len(library.read("chosen-app")["controls"]) == 6


def seed_main(popup):
    layout, _native, root, _main, _menu, _items, _popup_native, _scope = popup
    directory=knowledge.initialize_app("chosen-app", "Example", root)
    knowledge.save_page(directory, dict(id=layout["page_id"], app_id="chosen-app", name="人工区域布局",
        layout_confirmed=True, observed_window_size=layout["image_size"]))
    for region in layout["regions"]:
        knowledge.save_region(directory, {**region,"id":layout["page_id"]+"."+region["id"],
            "page_id":layout["page_id"],"app_id":"chosen-app","source":"manual"})
    entry=dict(id="file-entry",name="文件",app_id="chosen-app",page_id=layout["page_id"],
        region_id="toolbar",status="observed",control_type="MenuItemControl")
    knowledge.save_control(directory,entry)
    return KnowledgeLibrary(root,page_id=layout["page_id"]),entry


def expand(popup, entry):
    return scanner.scan_window("Example", "", "", "", "test", root=popup[2], manual_layout=popup[0],
        semantic_mode="native", navigation_context=dict(path=[entry["id"]], popup_handles=[2]))


def test_popup_scan_never_detects_main_or_merges_toolbar_qaction(popup, monkeypatch):
    library,entry=seed_main(popup)
    item=popup[5][0];item.AutomationId="shared-action"
    twin={**entry,"id":"toolbar-twin","name":item.Name,"automation_id":item.AutomationId,
          "location":{"Xpath":[{"ControlType":"WindowControl","Name":"Example"},
                               {"ControlType":"MenuItemControl","Name":item.Name}]}}
    knowledge.save_control(library.directory("chosen-app"),twin)
    original=knowledge.find_control_record(library.directory("chosen-app"),twin["id"])[1]
    monkeypatch.setattr(local_learning,"detect",lambda *a,**kw:pytest.fail("native popup must not rescan main or run detector"))
    first=expand(popup,entry);second=expand(popup,entry)
    snapshot=library.read("chosen-app")
    assert first["observed_control_ids"]==second["observed_control_ids"]
    assert len(snapshot["controls"])==7
    assert next(r for r in snapshot["controls"] if r["id"]==twin["id"])==original
    assert not any(r.get("region_id")=="unassigned" for r in snapshot["controls"])
    assert len(snapshot["pages"])==1


@pytest.mark.parametrize("legacy_scope", [False,True])
def test_standalone_popup_records_move_to_entry_without_new_ids(popup,legacy_scope):
    old=scan_menu(popup,layout_overrides={"page_id":"old-menu-page"})
    library,entry=seed_main(popup)
    if legacy_scope:
        for record in current_records(popup[2],old):
            record["window_scope"]={"kind":"main"}
            record["location"]["Xpath"][0]["foundIndex"]=1
            record["location"]["Xpath"].insert(0,{"ControlType":"WindowControl","Name":"Example"})
            knowledge.save_control(library.directory("chosen-app"),record)
    result=expand(popup,entry)
    assert result["observed_control_ids"]==old["observed_control_ids"]
    assert len(library.read("chosen-app")["controls"])==6
    assert all(r["page_id"]==popup[0]["page_id"] for r in current_records(popup[2],result))


def test_popup_commands_execute_through_saved_location(popup,monkeypatch):
    from easy_uiauto.mcp import ui_cli
    library,entry=seed_main(popup)
    result=expand(popup,entry)
    monkeypatch.setattr(scanner.native_windows,"enumerate_windows",lambda **kw:[popup[1],popup[6]])
    monkeypatch.setattr(ui_cli,"_window_rect",lambda w,*a:scanner._rect(w))
    monkeypatch.setattr(ui_cli,"_screenshot_window",scanner._screenshot_window)
    monkeypatch.setattr(ui_cli,"_trim_window_capture",lambda image,bounds,w:(image,bounds))
    executed=[]
    monkeypatch.setattr(ui_cli,"_perform_action",lambda action,text,control,rectangle,**kw:executed.append(control.Name))
    for record in current_records(popup[2],result):
        response=ui_cli.execute(library.directory("chosen-app"),record["command"]+".click",confirm=True)
        assert response["resolved_by"]=="location" and response["action_executed"]
    assert set(executed)==set(NAMES)


def test_qt_combo_popup_list_items_are_current_and_attached(popup):
    library,entry=seed_main(popup)
    popup[4].ClassName="QComboBoxPrivateContainer"
    for item in popup[5]:item.ControlTypeName="ListItemControl"
    result=expand(popup,entry)
    records=current_records(popup[2],result)
    assert len(records)==5 and all(r["control_type"]=="ListItemControl" for r in records)
    assert all(r["region_id"].startswith("menu-region-") for r in records)


def test_five_menus_thirty_items_keep_distinct_regions_and_repeat_without_duplicates(popup):
    library,entry=seed_main(popup)
    all_ids=set()
    menu=popup[4]
    for number,(label,count) in enumerate((("文件",5),("视图",8),("仿真",13),("辅助应用",1),("帮助",3))):
        opener={**entry,"id":entry["id"] if number==0 else f"entry-{number}","name":label}
        knowledge.save_control(library.directory("chosen-app"),opener)
        menu.Name="menu_"+label
        menu._children=[FakeControl(f"{label}项{i}","MenuItemControl",(355,55+i*10,565,64+i*10),parent=menu)
                        for i in range(count)]
        for item in menu._children:item.NativeWindowHandle=0
        first=expand(popup,opener);second=expand(popup,opener)
        assert first["observed_control_ids"]==second["observed_control_ids"]
        records=current_records(popup[2],first)
        assert len(records)==count
        all_ids.update(first["observed_control_ids"])
        regions={r["id"]:r for r in library.read("chosen-app")["regions"]}
        for record in records:
            assert regions[record["region_id"]]["parent_control_id"]==opener["id"]
            assert "click" in record["actions"]
            assert record["command"].startswith("manual-main."+record["region_id"]+".")
    assert len(all_ids)==30
    snapshot=library.read("chosen-app")
    assert len(snapshot["controls"])==35 and len(snapshot["pages"])==1
    assert len([r for r in snapshot["regions"] if r.get("parent_control_id")])==5


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
