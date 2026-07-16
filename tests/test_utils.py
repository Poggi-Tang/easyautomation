from __future__ import annotations

import time

import pytest

from conftest import FakeControl, FakeRect
from easy_uiauto import utils


def test_disassemble_location_preserves_control_type():
    location = utils.package_location("w", "n", "c", "ButtonControl", 2, "a", [], "", {})
    assert utils.disassemble_location(location)[3] == "ButtonControl"


def test_empty_xpath_uses_dictionary_fallback(monkeypatch):
    window = FakeControl(name="window", control_type="WindowControl")
    target = FakeControl(parent=window)

    def strategy(dictionary, parent=None, timeout=10):
        return window if dictionary["ControlType"] == "WindowControl" else target

    monkeypatch.setattr(utils, "strategy_dictionary", strategy)
    monkeypatch.setattr(utils, "set_top_window", lambda _title: True)
    location = utils.package_location(
        "window", "control", "FakeClass", "ButtonControl", 1, "", [], "", {}
    )

    assert utils.find_control(location) is target


def test_xpath_string_is_parsed_without_eval(monkeypatch):
    expected = FakeControl()
    captured = []
    monkeypatch.setattr(
        utils,
        "strategy_xpath",
        lambda name, class_name, control_type, xpath, debug=False, **_kwargs: (
            captured.append(xpath) or expected
        ),
    )
    location = utils.package_location(
        "",
        "control",
        "FakeClass",
        "ButtonControl",
        1,
        "",
        "[{'ControlType': 'ButtonControl', 'Name': 'control'}]",
        "",
        {},
    )

    assert utils.find_control(location) is expected
    assert captured == [[{"ControlType": "ButtonControl", "Name": "control"}]]


def test_malicious_xpath_string_is_rejected(monkeypatch):
    called = []
    monkeypatch.setattr(utils, "strategy_xpath", lambda *_args, **_kwargs: called.append(True))
    location = utils.package_location(
        "", "", "", "", 0, "", "__import__('os').getcwd()", "", {}
    )

    assert utils.find_control(location) is None
    assert called == []


def test_malformed_xpath_nodes_are_rejected():
    assert utils.find_control_by_xpath([1]) is None
    assert utils.find_control(
        utils.package_location("", "", "", "", 0, "", "[1]", "", {})
    ) is None


def test_auto_scroll_uses_correct_direction(monkeypatch):
    calls = []
    monkeypatch.setattr(utils.pyautogui, "scroll", lambda value, **_kwargs: calls.append(value))
    horizontal = []
    monkeypatch.setattr(
        utils.pyautogui, "hscroll", lambda value, **_kwargs: horizontal.append(value)
    )
    utils.auto_scroll(3, direction="up")
    utils.auto_scroll(4, direction="down")
    utils.auto_scroll(2, direction="left")
    utils.auto_scroll(5, direction="right")
    assert calls == [3, -4]
    assert horizontal == [-2, 5]


def test_cache_key_changes_with_found_index():
    first = [{"ControlType": "ListItemControl", "Name": "row", "foundIndex": 1}]
    second = [{"ControlType": "ListItemControl", "Name": "row", "foundIndex": 2}]
    assert utils.generate_cache_key(first) != utils.generate_cache_key(second)


def test_expired_cache_is_not_returned(monkeypatch):
    control = FakeControl()
    xpath = [{"ControlType": "ButtonControl", "Name": "control", "foundIndex": 1}]
    key, _ = utils.generate_cache_keys(xpath)
    utils.CONTROL_CACHE[key] = control
    utils.CACHE_METADATA[key] = {"last_verified": time.monotonic() - 10}
    monkeypatch.setattr(utils, "CONTROL_CACHE_TIMEOUT", 0.1)
    monkeypatch.setattr(utils, "_find_control_from_root", lambda _xpath, debug=False: None)

    assert utils.find_control_by_xpath(xpath) is None


def test_use_cache_false_does_not_store_control(monkeypatch):
    utils.CONTROL_CACHE.clear()
    utils.CACHE_METADATA.clear()
    control = FakeControl()
    xpath = [{"ControlType": "ButtonControl", "Name": "control"}]
    monkeypatch.setattr(utils, "_find_control_from_root", lambda _xpath, debug=False: control)
    assert utils.find_control_by_xpath(xpath, use_cache=False) is control
    assert utils.CONTROL_CACHE == {}


def test_cache_stats_and_clear_release_entries():
    utils.CONTROL_CACHE.clear()
    utils.CACHE_METADATA.clear()
    utils._store_cached_control("one", FakeControl())

    stats = utils.get_control_cache_stats()
    removed = utils.clear_control_cache()

    assert stats["entries"] == 1
    assert removed == 1
    assert utils.CONTROL_CACHE == {}
    assert utils.CACHE_METADATA == {}


def test_cache_rejects_reused_control_with_changed_identity():
    utils.CONTROL_CACHE.clear()
    utils.CACHE_METADATA.clear()
    control = FakeControl(name="first")
    utils._store_cached_control("key", control)
    control.Name = "second"
    assert utils._get_cached_control("key") is None


def test_get_control_xpath_records_skipped_depths():
    desktop = FakeControl(name="desktop", control_type="DesktopControl", rect=FakeRect(0, 0, 500, 500))
    window = FakeControl(name="window", control_type="WindowControl", parent=desktop)
    group = FakeControl(name="group", control_type="GroupControl", parent=window)
    leaf = FakeControl(name="leaf", control_type="ButtonControl", parent=group)

    xpath = utils.get_control_xpath(leaf)

    assert [item["Name"] for item in xpath] == ["window", "leaf"]
    assert [item["searchDepth"] for item in xpath] == [1, 2]


def test_get_control_xpath_treats_win32_desktop_pane_as_root():
    desktop = FakeControl(
        name="desktop",
        class_name="#32769",
        control_type="PaneControl",
        rect=FakeRect(0, 0, 500, 500),
    )
    taskbar = FakeControl(
        name="taskbar",
        class_name="Shell_TrayWnd",
        control_type="PaneControl",
        parent=desktop,
    )

    assert utils.get_control_xpath(taskbar) == [
        {
            "ControlType": "PaneControl",
            "searchDepth": 1,
            "Name": "taskbar",
            "ClassName": "Shell_TrayWnd",
        }
    ]


def test_found_index_uses_same_identity_fields_as_replay():
    parent = FakeControl(control_type="ListControl")
    FakeControl(
        name="row",
        control_type="ListItemControl",
        automation_id="first",
        parent=parent,
    )
    target = FakeControl(
        name="row",
        control_type="ListItemControl",
        automation_id="second",
        parent=parent,
    )
    xpath = utils.get_control_xpath(target)
    assert xpath[-1]["foundIndex"] == 1


def test_found_index_treats_empty_attributes_as_replay_wildcards():
    parent = FakeControl(control_type="ListControl")
    FakeControl(
        name="row",
        control_type="ListItemControl",
        automation_id="first",
        parent=parent,
    )
    target = FakeControl(
        name="row",
        control_type="ListItemControl",
        automation_id="",
        parent=parent,
    )

    xpath = utils.get_control_xpath(target)

    assert xpath[-1]["foundIndex"] == 2


def test_xpath_replay_uses_exact_depth(monkeypatch):
    captured = []

    class Found:
        def Refind(self, **_kwargs):
            return True

    class Root:
        def Control(self, **conditions):
            captured.append(conditions)
            return Found()

    monkeypatch.setattr(utils.uiautomation, "GetRootControl", Root)

    result = utils._find_control_from_root(
        [{"ControlType": "ButtonControl", "Name": "same", "searchDepth": 2}]
    )

    assert result is not None
    assert captured[0]["Depth"] == 2
    assert captured[0]["searchDepth"] == 2


def test_strategy_dictionary_honors_explicit_timeout(monkeypatch):
    calls = []

    class Found:
        def Refind(self, **kwargs):
            calls.append(kwargs)
            return True

    class Root:
        def Control(self, **_conditions):
            return Found()

    assert utils.strategy_dictionary(
        {"ControlType": "ButtonControl"}, parent=Root(), timeout=1.25
    ) is not None
    assert calls == [{"maxSearchSeconds": 1.25, "raiseException": False}]


def test_uiautomation_initializer_is_reused_per_thread(monkeypatch):
    calls = []
    monkeypatch.delattr(utils._UIA_THREAD_STATE, "initializer", raising=False)
    monkeypatch.setattr(
        utils.uiautomation,
        "UIAutomationInitializerInThread",
        lambda: calls.append("init") or object(),
    )
    utils.ensure_uiautomation_thread()
    utils.ensure_uiautomation_thread()
    assert calls == ["init"]


def test_image_control_adapter(monkeypatch, tmp_path):
    image = tmp_path / "button.png"
    image.write_bytes(b"not-an-image-but-location-is-mocked")
    monkeypatch.setattr(utils.pyautogui, "locateOnScreen", lambda *_args, **_kwargs: (10, 20, 30, 40))
    control = utils.find_control(
        utils.package_location("", "", "", "", 0, "", [], str(image), {})
    )
    assert control.ControlTypeName == "ImageControl"
    assert utils.get_control_coordinates(control) == (10, 20, 40, 60)


def test_image_locator_with_window_name_does_not_return_parent(monkeypatch, tmp_path):
    image = tmp_path / "button.png"
    image.write_bytes(b"image")
    parent = FakeControl(name="window", control_type="WindowControl")
    monkeypatch.setattr(utils, "strategy_dictionary", lambda *_args, **_kwargs: parent)
    monkeypatch.setattr(utils.pyautogui, "locateOnScreen", lambda *_args, **_kwargs: (1, 2, 3, 4))

    control = utils.find_control(
        utils.package_location("window", "", "", "", 0, "", [], str(image), {})
    )

    assert control.ControlTypeName == "ImageControl"


def test_desktop_lookup_has_no_win_d_side_effect(monkeypatch):
    expected = FakeControl(class_name="WorkerW", control_type="PaneControl")
    calls = []
    monkeypatch.setattr(utils.uiautomation, "SendKeys", lambda *args: calls.append(args))
    monkeypatch.setattr(utils, "strategy_xpath", lambda *_args, **_kwargs: expected)

    control = utils.find_control(
        utils.package_location(
            "",
            "",
            "WorkerW",
            "PaneControl",
            1,
            "",
            [{"ControlType": "PaneControl", "ClassName": "WorkerW"}],
            "",
            {},
        )
    )

    assert control is expected
    assert calls == []


def test_negative_virtual_screen_coordinates_remain_valid():
    assert utils.RectStruct(-400, 20, -200, 120).check_pos == (-300, 70)


def test_invalid_image_confidence_returns_none(monkeypatch, tmp_path):
    image = tmp_path / "button.png"
    image.write_bytes(b"image")
    monkeypatch.setattr(
        utils.pyautogui,
        "locateOnScreen",
        lambda *_args, **_kwargs: pytest.fail("invalid confidence must fail first"),
    )

    assert utils._find_image_control(
        {"Img": str(image), "PARAMETERS": {"confidence": "high"}}
    ) is None


def test_virtual_list_lookup_does_not_use_global_mouse_scroll(monkeypatch):
    class Container:
        def GetScrollPattern(self):
            raise RuntimeError("pattern unavailable")

    results = iter([None, Container()])
    monkeypatch.setattr(
        utils,
        "find_control_by_xpath",
        lambda *_args, **_kwargs: next(results),
    )
    calls = []
    monkeypatch.setattr(utils, "auto_scroll", lambda *_args, **_kwargs: calls.append(True))

    control = utils.strategy_xpath(
        "row",
        "",
        "ListItemControl",
        [
            {"ControlType": "ListControl"},
            {"ControlType": "ListItemControl", "Name": "row"},
        ],
    )

    assert control is None
    assert calls == []
