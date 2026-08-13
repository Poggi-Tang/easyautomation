"""Tests for the canonical easy_uiauto LOCATION flow."""

from __future__ import annotations

import json
from types import SimpleNamespace

from easy_uiauto import utils
from easy_uiauto.mcp import server
from easy_uiauto.mcp.protocol import (
    build_location,
    location_from_xpath,
    normalize_location,
    stable_location,
)

XPATH = [
    {
        "ControlType": "WindowControl",
        "Name": "Example",
        "ClassName": "WindowClass",
        "searchDepth": 1,
    },
    {
        "ControlType": "ButtonControl",
        "Name": "Save",
        "ClassName": "ButtonClass",
        "AutomationId": "saveButton",
        "foundIndex": 2,
        "searchDepth": 3,
    },
]


def _fake_control():
    return SimpleNamespace(
        Name="Save",
        ClassName="ButtonClass",
        ControlTypeName="ButtonControl",
        AutomationId="saveButton",
        BoundingRectangle=SimpleNamespace(left=10, top=20, right=110, bottom=60),
        IsEnabled=True,
        IsVisible=True,
        FrameworkId="Win32",
    )


def test_location_from_xpath_matches_recorded_location_shape() -> None:
    location = location_from_xpath(XPATH, parameters={"x": -1, "y": -1})

    assert location == {
        "WindowName": "Example",
        "Name": "Save",
        "ClassName": "ButtonClass",
        "ControlType": "ButtonControl",
        "foundIndex": 2,
        "AutomationId": "saveButton",
        "Xpath": XPATH,
        "Img": "",
        "PARAMETERS": {"x": -1, "y": -1},
    }


def test_normalize_location_accepts_recorded_action_wrapper() -> None:
    location = location_from_xpath(XPATH)
    action = {"TEST_ID": "1", "ACTION": "点击", "LOCATION": location}

    assert normalize_location(action) == location


def test_stable_location_uses_automation_id_instead_of_dynamic_name() -> None:
    location = stable_location(location_from_xpath(XPATH))

    assert location["Name"] == ""
    assert "Name" not in location["Xpath"][-1]
    assert location["AutomationId"] == "saveButton"
    assert location["Xpath"][-1]["AutomationId"] == "saveButton"
    assert location["Xpath"][0]["Name"] == "Example"


def test_normalize_location_accepts_vector_record_without_losing_indexes() -> None:
    record = {
        "window_name": "Example",
        "name": "Save",
        "class_name": "ButtonClass",
        "control_type": "ButtonControl",
        "automation_id": "saveButton",
        "xpath": [
            {"control_type": "WindowControl", "name": "Example", "search_depth": 1},
            {
                "control_type": "ButtonControl",
                "name": "Save",
                "found_index": 2,
                "search_depth": 3,
            },
        ],
    }

    location = normalize_location(record)

    assert location["Xpath"][-1]["foundIndex"] == 2
    assert location["Xpath"][-1]["searchDepth"] == 3


def test_find_control_passes_complete_location_to_core(monkeypatch) -> None:
    expected = location_from_xpath(XPATH)
    received: dict = {}

    def fake_find(location, debug=False):
        received.update(location)
        return _fake_control()

    monkeypatch.setattr(server, "_find_control", fake_find)

    result = json.loads(server.find_control(location={"LOCATION": expected}))

    assert received == expected
    assert result["LOCATION"] == expected
    assert result["bounds"] == {"left": 10, "top": 20, "right": 110, "bottom": 60}


def test_get_control_at_position_returns_reusable_location(monkeypatch) -> None:
    monkeypatch.setattr(server, "get_control_info", lambda _x, _y: (XPATH, _fake_control()))

    result = json.loads(server.get_control_at_position(50, 40))

    assert result["LOCATION"] == location_from_xpath(XPATH)
    assert result["LOCATION"]["Xpath"][-1]["foundIndex"] == 2


def test_capture_record_contains_canonical_location(monkeypatch) -> None:
    monkeypatch.setattr(server, "get_control_info", lambda _x, _y: (XPATH, _fake_control()))
    monkeypatch.setattr(server, "_learning_log", lambda _message: None)

    record = json.loads(server.capture_control_record_at_position(50, 40))

    assert record["LOCATION"] == location_from_xpath(XPATH)
    assert record["xpath"][-1]["found_index"] == 2
    assert record["xpath"][-1]["search_depth"] == 3


def test_core_find_control_keeps_legacy_empty_xpath_fallback(monkeypatch) -> None:
    control = _fake_control()
    window = SimpleNamespace()

    def fake_strategy_dictionary(selector, parent=None):
        if selector["ControlType"] == "WindowControl":
            return window
        assert parent is window
        return control

    monkeypatch.setattr(utils, "CURRENT_APP_NAME", "Example")
    monkeypatch.setattr(utils, "_find_window_scope", lambda *_args: window)
    monkeypatch.setattr(utils, "strategy_xpath", lambda *_args, **_kwargs: False)
    monkeypatch.setattr(utils, "strategy_dictionary", fake_strategy_dictionary)
    location = build_location(
        window_name="Example",
        name="Save",
        class_name="ButtonClass",
        control_type="ButtonControl",
        automation_id="saveButton",
    )

    assert utils.find_control(location) is control


def test_core_find_control_prefers_automation_id_before_xpath(monkeypatch) -> None:
    control = _fake_control()
    window = SimpleNamespace()
    selectors = []

    def fake_strategy_dictionary(selector, parent=None):
        selectors.append(selector)
        if selector.get("ControlType") == "WindowControl":
            return window
        assert parent is window
        if selector.get("AutomationId") == "saveButton":
            return control
        return False

    monkeypatch.setattr(utils, "CURRENT_APP_NAME", "Example")
    monkeypatch.setattr(utils, "_find_window_scope", lambda *_args: window)
    monkeypatch.setattr(utils, "strategy_dictionary", fake_strategy_dictionary)
    monkeypatch.setattr(
        utils,
        "strategy_xpath",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("XPath must not run after an AutomationId match")
        ),
    )

    assert utils.find_control(location_from_xpath(XPATH)) is control
    assert selectors[0] == {
        "ControlType": "ButtonControl",
        "ClassName": "ButtonClass",
        "AutomationId": "saveButton",
    }


def test_core_find_control_does_not_call_xpath_when_path_is_empty(monkeypatch) -> None:
    control = _fake_control()
    window = SimpleNamespace()

    def fake_strategy_dictionary(selector, parent=None):
        if selector.get("ControlType") == "WindowControl":
            return window
        if selector.get("AutomationId"):
            return False
        return control

    monkeypatch.setattr(utils, "CURRENT_APP_NAME", "Example")
    monkeypatch.setattr(utils, "_find_window_scope", lambda *_args: window)
    monkeypatch.setattr(utils, "strategy_dictionary", fake_strategy_dictionary)
    monkeypatch.setattr(
        utils,
        "strategy_xpath",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("empty XPath must never be resolved")
        ),
    )
    location = build_location(
        window_name="Example",
        name="Save",
        class_name="ButtonClass",
        control_type="ButtonControl",
        automation_id="stale-id",
    )

    assert utils.find_control(location) is control


def test_core_find_control_uses_xpath_before_property_combination(monkeypatch) -> None:
    control = _fake_control()
    window = SimpleNamespace()
    combination_used = False

    def fake_strategy_dictionary(selector, parent=None):
        nonlocal combination_used
        if selector.get("ControlType") == "WindowControl":
            return window
        if selector.get("AutomationId"):
            return False
        combination_used = True
        return control

    monkeypatch.setattr(utils, "CURRENT_APP_NAME", "Example")
    monkeypatch.setattr(utils, "_find_window_scope", lambda *_args: window)
    monkeypatch.setattr(utils, "strategy_dictionary", fake_strategy_dictionary)
    monkeypatch.setattr(utils, "strategy_xpath", lambda *_args, **_kwargs: control)

    assert utils.find_control(location_from_xpath(XPATH)) is control
    assert combination_used is False
