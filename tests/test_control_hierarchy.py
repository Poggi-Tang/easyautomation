from copy import deepcopy

import pytest

from easy_uiauto.mcp import control_hierarchy as hierarchy, knowledge
from easy_uiauto.mcp.protocol import location_from_xpath


WINDOW = dict(ControlType="WindowControl", Name="Demo")
PANEL = dict(ControlType="PaneControl", Name="Options", foundIndex=1)


def record(identifier, path=None, **extra):
    return {"id": identifier, "page_id": "main", "region_id": "toolbar",
            "name": identifier, "location": location_from_xpath(path or [WINDOW, dict(ControlType="ButtonControl", Name=identifier)]), **extra}


def test_native_parent_uses_full_xpath_prefix_without_mutating_locators():
    parent = record("panel", [WINDOW, PANEL])
    child = record("button", [WINDOW, PANEL, dict(ControlType="ButtonControl", Name="Save")])
    original = deepcopy([parent, child])
    result = hierarchy.build([parent, child])
    assert result["parents"] == {"panel": None, "button": "panel"}
    assert result["sources"]["button"] == "uia"
    assert [parent, child] == original


def test_same_named_sibling_panels_use_recorded_ordinals():
    p2 = {**PANEL, "foundIndex": 2}
    controls = [record("one", [WINDOW, PANEL]), record("two", [WINDOW, p2]),
                record("child", [WINDOW, p2, dict(ControlType="ButtonControl", Name="OK")])]
    assert hierarchy.build(controls)["parents"]["child"] == "two"


def test_menu_layers_follow_navigation_and_inherit_entry_region():
    controls = [record("file"), record("recent", navigation_path=["file"], region_id="content"),
                record("project", navigation_path=["file", "recent"], region_id="unassigned")]
    original = deepcopy(controls)
    result = hierarchy.build(controls)
    assert result["parents"] == {"file": None, "recent": "file", "project": "recent"}
    assert set(result["regions"].values()) == {"main.toolbar"}
    assert controls == original  # Physical region and executable path stay intact.


def test_native_container_does_not_swallow_controls_in_other_regions():
    controls = [record("panel", [WINDOW, PANEL], region_id="unassigned"),
                record("child", [WINDOW, PANEL, dict(ControlType="ButtonControl", Name="OK")])]
    assert hierarchy.build(controls)["parents"]["child"] is None


def test_manual_parent_and_explicit_region_root_override_automatic_relations():
    controls = [record("file"), record("edit"),
                record("new", navigation_path=["file"], hierarchy_parent="edit")]
    assert hierarchy.build(controls)["parents"]["new"] == "edit"
    controls[-1]["hierarchy_parent"] = ""
    assert hierarchy.build(controls)["parents"]["new"] is None
    controls[-1]["hierarchy_parent"] = "auto"
    assert hierarchy.build(controls)["parents"]["new"] == "file"


@pytest.mark.parametrize("parent", ["child", "panel", "missing", [], None])
def test_invalid_manual_parent_and_cycles_are_rejected(parent):
    controls = [record("panel", [WINDOW, PANEL]),
                record("child", [WINDOW, PANEL, dict(ControlType="ButtonControl", Name="OK")])]
    with pytest.raises(ValueError):
        hierarchy.validate_parent(controls, "panel", parent)


def test_corrupt_cycle_or_deleted_parent_never_drops_controls_or_hangs_tree():
    controls = [record("a", hierarchy_parent="b"), record("b", hierarchy_parent="a"),
                record("c", hierarchy_parent="deleted"), record("d", hierarchy_parent=[])]
    result = hierarchy.build(controls)
    assert result["parents"] == {"a": None, "b": None, "c": None, "d": None}
    assert len(result["warnings"]) == 2
    assert result["sources"]["c"] == "missing-parent"


def test_display_tree_supports_deep_chains_without_python_recursion():
    controls = [record(str(i), hierarchy_parent=str(i-1) if i else "") for i in range(1200)]
    result = hierarchy.build(controls)
    assert result["parents"]["1199"] == "1198"
    assert result["regions"]["1199"] == "main.toolbar"


def test_rescan_retains_manual_parent_override():
    old = record("button", app_id="demo", status="observed", hierarchy_parent="parent")
    assert knowledge.merge_control_state(old, record("button"))["hierarchy_parent"] == "parent"
