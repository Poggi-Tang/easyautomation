"""Parent/child selection, input filtering, and non-destructive region persistence."""
from copy import deepcopy
from types import SimpleNamespace

import pytest
from PIL import Image

from easy_uiauto.region_picker import ControlPath, PickerKeys, normalized_selection, selection_metadata
from easy_uiauto.mcp.protocol import location_from_xpath
from test_knowledge_library import library


class Node:
    def __init__(self, parent=None, **values):
        self.parent = parent
        self.AutomationId = values.get("id", "101")
        self.Name = values.get("name", "区域")
        self.ClassName = "NativePanel"
        self.ControlTypeName = "PaneControl"
        self.IsOffscreen = False
        self.BoundingRectangle = SimpleNamespace(left=-90, top=20, right=80, bottom=90)

    def GetParentControl(self): return self.parent


def test_wheel_up_parent_down_original_child_and_never_an_unrelated_sibling():
    desktop = Node(); window = Node(desktop); parent = Node(window); child = Node(parent)
    model = ControlPath(window, lambda a, b: a is b)
    model.reset(child)
    assert model.current is child
    assert model.scroll(1) is parent
    assert model.scroll(1) is window
    assert model.scroll(1) is window  # Never ascend to the desktop.
    assert model.scroll(-1) is parent
    assert model.scroll(-1) is child
    assert model.scroll(-1) is child
    other = Node(window); model.reset(other)
    assert model.level == 0 and model.current is other


def test_other_windows_and_cyclic_provider_paths_cannot_be_regions():
    model = ControlPath(Node(), lambda a, b: a is b)
    with pytest.raises(ValueError, match="所选软件"): model.reset(Node(Node()))
    cyclic = Node(); cyclic.parent = cyclic
    with pytest.raises(ValueError): model.reset(cyclic)
    assert model.current is None


def test_empty_automation_id_is_displayed_literally_and_bad_geometry_is_rejected():
    control = Node(id="", name="")
    metadata = selection_metadata(control)
    assert metadata["automation_id"] == metadata["native_name"] == ""
    assert metadata["class_name"] == "NativePanel" and metadata["control_type"] == "PaneControl"
    control.IsOffscreen = True
    with pytest.raises(ValueError, match="可见区域"): selection_metadata(control)


def test_normalization_supports_negative_monitors_and_clips_resize_borders():
    bounds = dict(left=-100, top=10, right=100, bottom=110, width=200, height=100)
    assert normalized_selection(dict(left=-108, top=10, right=100, bottom=110), bounds) == dict(left=0,top=0,right=1,bottom=1)
    assert normalized_selection(dict(left=-90, top=20, right=80, bottom=90), bounds) == dict(left=.05,top=.1,right=.9,bottom=.8)
    with pytest.raises(ValueError): normalized_selection(dict(left=120,top=20,right=180,bottom=70), bounds)


def test_ctrl_and_escape_are_consumed_but_other_keys_are_not():
    keys = PickerKeys()
    assert keys.event(65, True) == (False, None)
    assert keys.event(0xA2, True) == (True, "confirm")
    assert keys.event(0xA2, True) == (True, None)  # Key-repeat.
    assert keys.event(0xA2, False) == (True, None)
    assert keys.event(0xA3, True) == (True, "confirm")
    assert keys.event(0x1B, True) == (True, "cancel")


def test_ctrl_already_held_when_picker_starts_requires_release_and_fresh_press():
    keys = PickerKeys([0x11, 0xA2])
    assert keys.event(0xA2, True) == (True, None)
    assert keys.event(0xA2, False) == (False, None)
    assert keys.event(0xA2, True) == (True, "confirm")
    assert PickerKeys([0xA2]).event(0x1B, True) == (True, "cancel")


def test_both_ctrl_keys_held_before_start_are_released_without_sticking():
    keys = PickerKeys([0x11, 0xA2, 0xA3])
    assert keys.event(0xA2, False) == (False, None)
    assert keys.event(0xA3, True) == (True, None)
    assert keys.event(0xA3, False) == (False, None)
    assert keys.event(0xA2, True) == (True, "confirm")


def picked(tmp_path, *, rect=None, name="工具栏", size=(400, 300), color="blue"):
    image = tmp_path / f"picked-{color}.png"; Image.new("RGB", size, color).save(image)
    location = location_from_xpath([
        dict(ControlType="WindowControl", Name="Test", match_mode="position"),
        dict(ControlType="PaneControl", Name="工具栏", foundIndex=2, match_mode="position"),
    ], parameters={"match_mode":"position"})
    return dict(selection=dict(automation_id="101", native_name=name, class_name="NativePanel",
        control_type="PaneControl", location=location,
        normalized_rect=rect or dict(left=0,top=0,right=1,bottom=.4)),
        capture=dict(image=str(image), width=size[0], height=size[1], window_name="Test", window_class="Window"))


def test_ctrl_confirmation_saves_native_region_and_exact_location(library, tmp_path):
    result = picked(tmp_path)
    identifier = library.save_picked_region("demo", result)
    region = next(r for r in library.regions("demo") if r["id"] == identifier)
    assert region["selection_source"] == "uia-picker" and region["name"] == "工具栏"
    assert region["geometry_role"] == "uia-control-region"
    assert region["location"]["Xpath"] == result["selection"]["location"]["Xpath"]
    assert region["location"]["PARAMETERS"]["match_mode"] == "position"
    assert region["automation_id"] == "101" and region["class_name"] == "NativePanel"
    assert library.page("demo")["layout_confirmed"]
    with Image.open(library.asset("demo", region["image"])) as image: assert image.size == (400,120)


def test_modify_region_keeps_id_and_name_and_does_not_append_another_region(library, tmp_path):
    identifier = library.save_picked_region("demo", picked(tmp_path))
    replacement = picked(tmp_path, rect=dict(left=0,top=0,right=1,bottom=.5), name="新 UIA 名称", color="green")
    saved = library.save_picked_region("demo", replacement, region_id=identifier, name="自定义区域", description="说明")
    assert saved == identifier and len(library.regions("demo")) == 1
    region = library.regions("demo")[0]
    assert region["name"] == "自定义区域" and region["native_name"] == "新 UIA 名称"
    assert region["description"] == "说明" and region["normalized_rect"]["bottom"] == .5


def test_overlapping_or_removed_region_does_not_replace_reference_or_records(library, tmp_path):
    identifier = library.save_picked_region("demo", picked(tmp_path))
    before = deepcopy(library.read("demo"))
    source = library.asset("demo", library.page("demo")["layout_image"])
    image_before = source.read_bytes()
    replacement = picked(tmp_path, color="green")
    with pytest.raises(ValueError, match="重叠"): library.save_picked_region("demo", replacement)
    assert library.read("demo") == before and source.read_bytes() == image_before
    with pytest.raises(ValueError, match="不存在"):
        library.save_picked_region("demo", replacement, region_id="manual-main.deleted")
    assert library.read("demo") == before and source.read_bytes() == image_before


def test_resize_requires_all_picked_regions_to_be_reviewed(library, tmp_path):
    first = library.save_picked_region("demo", picked(tmp_path))
    bottom = dict(left=0,top=.5,right=1,bottom=1)
    second = library.save_picked_region("demo", picked(tmp_path, rect=bottom))
    new = picked(tmp_path, size=(800,600), color="green")
    library.save_picked_region("demo", new, region_id=first)
    assert not library.page("demo")["layout_confirmed"]
    library.save_picked_region("demo", picked(tmp_path, size=(800,600), rect=bottom, color="green"), region_id=second)
    assert library.page("demo")["layout_confirmed"]


def test_missing_xpath_cannot_commit_a_region(library, tmp_path):
    result = picked(tmp_path); result["selection"]["location"]["Xpath"] = []
    with pytest.raises(ValueError, match="XPath"): library.save_picked_region("demo", result)
    assert not library.regions("demo")


@pytest.mark.parametrize("count", [0, 1, 2])
def test_existing_software_rebinds_only_an_exact_unique_visible_window(monkeypatch, count):
    import easy_uiauto.knowledge_gui as module
    candidates = [dict(handle=i, visible=True, title="Native caption", class_name="Native") for i in range(count)]
    monkeypatch.setattr(module.native_windows, "matching_windows", lambda title: candidates)
    dialogs, tasks = [], []
    monkeypatch.setattr(module, "WindowChooser", lambda *args: dialogs.append(args))
    gui = SimpleNamespace(task=None, workflow_panel=SimpleNamespace(task=None), app_id="demo", windows={},
        indexes={"demo":{"application":{"window_name":"Native caption", "window_class":"Native"}}},
        root=object(), library=object(), _maybe_save=lambda: True,
        start_task=lambda *args, **kwargs: tasks.append((args,kwargs)))
    module.KnowledgeWindow.pick_region(gui)
    if count == 1:
        assert not dialogs and len(tasks) == 1 and gui.windows["demo"] == candidates[0]
        assert tasks[0][0][0] == "pick-region" and tasks[0][1]["hide"]
    else:
        assert len(dialogs) == 1 and not tasks and not gui.windows
