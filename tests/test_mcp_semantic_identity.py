"""Regression cases from adjacent menu items and misleading native buttons."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from PIL import Image

from easy_uiauto.mcp import knowledge, scanner, semantic_identity, ui_cli


@pytest.mark.parametrize(
    "name,intent",
    [
        ("全选(A)\tCtrl+A", "goto"),
        ("时间/日期(D)\tF5", "select-all"),
        ("替换(R)...\tCtrl+H", "find-previous"),
        ("转到(G)...\tCtrl+G", "replace"),
        ("查找上一个(V)\tShift+F3", "find-next"),
        ("关闭", "maximize-window"),
        ("格式(O)", "open-view-menu"),
        ("查看(V)", "open-help-menu"),
    ],
)
def test_known_native_action_cannot_take_adjacent_items_meaning(name, intent):
    evidence = semantic_identity.check(
        dict(intent=intent), dict(name=name, control_type="MenuItemControl")
    )
    assert evidence["status"] == "conflict"


@pytest.mark.parametrize(
    "name,intent,label",
    [
        ("全选（A）\tCtrl+A", "select-all", "Select all"),
        ("保存(&S)...\tCtrl+S", "save-document", "Save"),
        ("&Find Previous...", "find-previous", "查找上一个"),
        ("字体(F)...", "open-font-dialog", "Font"),
        ("关于测试软件(A)", "open-about-dialog", "About"),
    ],
)
def test_cross_language_labels_and_accelerators_are_not_conflicts(name, intent, label):
    result = semantic_identity.check(
        dict(intent=intent, semantic_name=label), dict(name=name, control_type="MenuItemControl")
    )
    assert result["status"] == "consistent"


def test_editable_values_are_not_native_action_identity():
    assert (
        semantic_identity.check(
            dict(intent="select-all"), dict(name="保存", control_type="EditControl")
        )["status"]
        == "not-comparable"
    )


def test_checkbox_with_matching_geometry_cannot_take_another_option_meaning():
    evidence = semantic_identity.check(
        dict(intent="toggle-match-case", semantic_name="Match case"),
        dict(name="循环(R)", control_type="CheckBoxControl"),
    )
    assert evidence["status"] == "conflict"
    unknown = semantic_identity.check(
        dict(intent="set-direction", semantic_name="Direction"),
        dict(name="循环(R)", control_type="CheckBoxControl"),
    )
    assert unknown["status"] == "unresolved"


def test_labeled_input_identity_uses_label_not_its_editable_value():
    native = dict(name="查找内容(N):", control_type="EditControl")
    assert (
        semantic_identity.check(dict(intent="enter-search-text"), native)["status"] == "consistent"
    )
    literal = semantic_identity.literal_semantics(native)
    assert literal["intent"] == "enter-search-text" and "set-text" in literal["actions"]
    assert (
        semantic_identity.literal_semantics(dict(name="Save", control_type="EditControl")) is None
    )


def test_replace_all_cannot_be_published_as_single_replacement():
    native = dict(name="全部替换(A)", control_type="ButtonControl")
    evidence = semantic_identity.check(
        dict(intent="replace-current-match", semantic_name="Replace"), native
    )
    assert evidence["status"] == "conflict"
    assert semantic_identity.literal_semantics(native)["intent"] == "replace-all"


def test_legacy_semantic_conflict_removed_from_catalog_without_deleting_evidence(tmp_path):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = dict(
        id="bad",
        app_id="example",
        page_id="main",
        region_id="menu",
        command="main.menu.goto",
        actions=["click"],
        name="全选(A)\tCtrl+A",
        control_type="MenuItemControl",
        semantic_name="Go to",
        intent="goto",
        status="verified",
        semantic_status="verified",
    )
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    assert knowledge.available_commands(directory) == []
    assert knowledge.find_control_record(directory, "bad") is not None


def test_runtime_identity_conflict_stops_before_image_or_ocr_fallback(monkeypatch, tmp_path):
    image = tmp_path / "save.png"
    Image.new("RGB", (20, 20), "white").save(image)
    record = dict(
        name="保存",
        semantic_name="Save",
        intent="save-file",
        control_type="MenuItemControl",
        image="save.png",
        location={},
    )
    monkeypatch.setattr(
        ui_cli,
        "resolve_location",
        lambda *_a, **_kw: SimpleNamespace(Name="全选", ControlTypeName="MenuItemControl"),
    )
    forbidden = Mock(side_effect=AssertionError("Conflicting identity must not fall back"))
    monkeypatch.setattr(ui_cli, "locate_template", forbidden)
    monkeypatch.setattr(ui_cli, "_locate_record_text", forbidden)
    with pytest.raises(RuntimeError, match="semantic identity conflict"):
        ui_cli._resolve_verified_control_in_context(tmp_path, record, object(), {}, object())
    forbidden.assert_not_called()


def test_visual_target_rejects_adjacent_conflicting_native_item():
    class Control:
        Name = "全选(A)\tCtrl+A"
        AutomationId = "25"
        ControlTypeName = "MenuItemControl"
        BoundingRectangle = SimpleNamespace(left=0, top=0, right=80, bottom=25)

        def GetParentControl(self):
            return None

    control = Control()
    window = SimpleNamespace(
        ControlTypeName="WindowControl",
        AutomationId="window",
        Name="Example",
        BoundingRectangle=SimpleNamespace(left=0, top=0, right=400, bottom=300),
    )
    selected = scanner.control_from_visual_target(
        window,
        scanner._rect(window),
        dict(
            intent="goto",
            semantic_name="Go to",
            actions=["click"],
            relative_rect=dict(left=0, top=0, right=80, bottom=25, width=80, height=25),
        ),
        point_lookup=lambda *_a: control,
    )
    assert selected is None
