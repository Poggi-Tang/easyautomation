"""Identity/position contracts across recording, normalization and execution."""

from copy import deepcopy
from unittest.mock import Mock

import pytest

from easy_uiauto import ctrl, utils
from easy_uiauto.mcp import scanner, server, service, ui_cli
from easy_uiauto.mcp.protocol import location_from_xpath, normalize_location, with_match_mode
from easy_uiauto.mcp.results import ActionError
from test_mcp_scanner import FakeControl


@pytest.fixture
def grid(monkeypatch):
    cells = [FakeControl("1.0" if i < 5 else str(i), "DataItemControl",
                         (i * 20, 10, i * 20 + 18, 30), automation_id=f"cell-{i}")
             for i in range(9)]
    table = FakeControl("Values", "TableControl", (0, 0, 190, 190), list(cells))
    window = FakeControl("Example", "WindowControl", (0, 0, 200, 200), [table])
    monkeypatch.setattr(utils, "_native_window_scope", lambda *_: window)
    monkeypatch.setattr(utils, "set_top_window", Mock(side_effect=AssertionError("Lookup activated a window")))
    return window, table, cells, location_from_xpath(utils.get_control_xpath(cells[4]))


@pytest.mark.parametrize("ordinal", [1, 5, 99])
def test_identity_ignores_recorded_ordinal_and_id_for_repeated_names(grid, ordinal):
    window, _, cells, location = grid
    location["foundIndex"] = location["Xpath"][-1]["foundIndex"] = ordinal
    assert utils.observe_control(location, window=window) is cells[0]
    assert utils.find_control(location) is cells[0]
    assert utils.find_control_by_xpath(location["Xpath"]) is cells[0]
    assert scanner.resolve_location(location, window=window) is cells[0]


def test_position_roundtrip_keeps_clicked_sibling_without_xy(grid):
    window, _, cells, _ = grid
    path = utils.get_control_xpath(cells[4], match_mode="position")
    assert path[-1]["foundIndex"] == 5
    assert path[-1]["AutomationId"] == "cell-4"
    assert utils.observe_control(location_from_xpath(path), window=window) is cells[4]


def test_position_counts_only_type_siblings_ignoring_name_class_and_id(grid):
    window, table, cells, location = grid
    table._children.insert(0, FakeControl("Header", "TextControl", (0, 0, 100, 10)))
    location = with_match_mode(location, "position")
    cells[4].Name, cells[4].ClassName, cells[4].AutomationId = "changed", "ChangedClass", "new-id"
    assert utils.observe_control(location, window=window) is cells[4]
    table._children[1], table._children[5] = table._children[5], table._children[1]
    assert utils.observe_control(location, window=window) is cells[0]


def test_same_record_can_switch_modes_without_mutation(grid):
    window, _, cells, location = grid
    original = deepcopy(location)
    assert utils.find_control(location, match_mode="position") is cells[4]
    assert utils.find_control_by_xpath(location["Xpath"], match_mode="position") is cells[4]
    assert utils.find_control(location) is cells[0]
    assert location == original


def test_per_step_modes_allow_identity_parent_and_positional_cell(grid):
    window, _, cells, location = grid
    location["Xpath"][-1]["match_mode"] = "position"
    cache = {}
    assert utils.observe_control(location, window=window, prefix_cache=cache) is cells[4]
    location["Xpath"][-1]["match_mode"] = "identity"
    assert utils.observe_control(location, window=window, prefix_cache=cache) is cells[0]


def test_identity_does_not_fall_back_to_id_when_name_is_missing(grid):
    window, _, _, location = grid
    location["Name"] = location["Xpath"][-1]["Name"] = "does-not-exist"
    with pytest.raises(ActionError, match="control_not_found"):
        utils.observe_control(location, window=window)


@pytest.mark.parametrize("index", [0, 10])
def test_invalid_position_never_switches_to_identity(grid, index):
    window, _, _, location = grid
    location = with_match_mode(location, "position")
    location["foundIndex"] = location["Xpath"][-1]["foundIndex"] = index
    with pytest.raises(ActionError) as raised:
        utils.observe_control(location, window=window)
    assert raised.value.code == ("invalid_argument" if index == 0 else "control_not_found")
    if index:
        assert raised.value.reason == "index_out_of_range"


def test_normalization_and_vector_xpath_preserve_mode(grid):
    _, _, _, location = grid
    normalized = normalize_location({"LOCATION": {**location, "match_mode": "position"}})
    assert normalized["PARAMETERS"]["match_mode"] == "position"
    assert all(step["match_mode"] == "position" for step in normalized["Xpath"])
    vector_step = server._normalize_xpath_step(normalized["Xpath"][-1])
    assert normalize_location({"xpath": [vector_step]})["Xpath"][-1]["match_mode"] == "position"


def test_bad_mode_is_rejected_before_uia(monkeypatch):
    forbidden = Mock(side_effect=AssertionError("Invalid mode touched UIA"))
    monkeypatch.setattr(utils, "_native_window_scope", forbidden)
    with pytest.raises(ActionError) as raised:
        utils.find_control({"WindowName": "Example", "Name": "1.0", "match_mode": "unknown"})
    assert raised.value.code == "invalid_argument"
    forbidden.assert_not_called()


def test_replay_override_reaches_legacy_wrapper_without_mutating_record(monkeypatch, grid):
    _, _, _, location = grid
    action = {"ACTION": "点击", "TEST_ID": 1, "LOCATION": location}
    original = deepcopy(action)
    operation = Mock(return_value="success")
    monkeypatch.setitem(ctrl.Execute_Function, "点击", operation)
    assert ctrl.run_action(action, match_mode="position") == "success"
    forwarded = operation.call_args.kwargs
    assert forwarded["PARAMETERS"]["match_mode"] == "position"
    assert all(step["match_mode"] == "position" for step in forwarded["Xpath"])
    assert action == original


def test_native_root_attributes_are_respected_by_scanner(monkeypatch):
    first = FakeControl("Example", "WindowControl", (0, 0, 100, 100))
    second = FakeControl("Example", "WindowControl", (0, 0, 100, 100))
    first.ClassName, second.ClassName = "First", "Second"
    entries = [{"title": "Example", "handle": 1, "visible": True},
               {"title": "Example", "handle": 2, "visible": True}]
    monkeypatch.setattr(scanner.native_windows, "enumerate_windows", lambda **_: entries)
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", lambda hwnd: {1: first, 2: second}[hwnd])
    location = {"Xpath": [{"ControlType": "WindowControl", "Name": "Example", "ClassName": "Second"}]}
    assert utils.find_control(location) is second
    assert scanner.resolve_location(location) is second


def test_tcp_builder_preserves_mode(monkeypatch):
    operation = Mock(return_value="success")
    monkeypatch.setattr(service.Controller, "left_click", operation)
    adapter = object.__new__(service.UIAutomationService)
    adapter._handle_click({"window_name": "Example", "control_type": "DataItemControl",
                           "found_index": 5, "match_mode": "position"}, "test")
    assert operation.call_args.kwargs["PARAMETERS"]["match_mode"] == "position"


def test_tcp_smart_find_returns_a_replayable_positional_xpath(grid):
    import json
    window, _, cells, location = grid
    adapter = object.__new__(service.UIAutomationService)
    result = adapter._handle_smart_find({"window_name": "Example", "match_mode": "position",
                                       "xpath": json.dumps(location["Xpath"])}, "test")
    assert all(step["match_mode"] == "position" for step in result["xpath"])
    assert utils.observe_control(location_from_xpath(result["xpath"]), window=window) is cells[4]


def test_tcp_smart_find_preserves_mixed_path_intent(grid):
    import json
    window, _, cells, location = grid
    location["Xpath"][-1]["match_mode"] = "position"
    adapter = object.__new__(service.UIAutomationService)
    result = adapter._handle_smart_find({"window_name": "Example",
                                       "xpath": json.dumps(location["Xpath"])}, "test")
    assert result["xpath"] == location["Xpath"]
    assert utils.observe_control(location_from_xpath(result["xpath"]), window=window) is cells[4]


@pytest.mark.parametrize("method", ["click", "type_text", "set_text"])
def test_mcp_action_builders_preserve_mode_and_text(monkeypatch, method):
    monkeypatch.setattr(server, "_mode_blocks_operation", lambda: None)
    monkeypatch.setattr(server, "_auto_learn_location", Mock())
    operation = Mock(return_value="Success")
    target = {"click": "left_click", "type_text": "input_text", "set_text": "set_text"}[method]
    monkeypatch.setattr(server.Controller, target, operation)
    arguments = dict(window_name="Example", control_type="DataItemControl", found_index=5,
                     match_mode="position")
    if method != "click":
        arguments["text"] = "new value"
    getattr(server, method)(**arguments)
    operation.assert_called_once()
    parameters = operation.call_args.kwargs["PARAMETERS"]
    assert parameters["match_mode"] == "position"
    if method != "click":
        assert "new value" in parameters.values()


def test_positional_parent_failure_never_uses_visual_fallback(monkeypatch, tmp_path):
    from PIL import Image
    from test_mcp_ui_cli import _record
    record = _record()
    record["location"]["match_mode"] = "position"
    path = tmp_path / record["image"]
    path.parent.mkdir(parents=True)
    Image.new("RGB", (20, 20), "white").save(path)
    monkeypatch.setattr(ui_cli, "resolve_location", Mock(side_effect=ActionError("control_not_found", "parent missing")))
    fallback = Mock(side_effect=AssertionError("Position failure used identity evidence"))
    monkeypatch.setattr(ui_cli, "locate_template", fallback)
    monkeypatch.setattr(ui_cli, "_locate_record_text", fallback)
    with pytest.raises(ActionError, match="control_not_found"):
        ui_cli._resolve_verified_control_in_context(tmp_path, record, object(), {}, None)
    fallback.assert_not_called()


@pytest.mark.parametrize("mode", ["identity", "position"])
def test_reference_image_cannot_replace_the_selected_control(monkeypatch, tmp_path, mode):
    from PIL import Image
    from test_mcp_ui_cli import _record
    record = _record()
    record["location"]["match_mode"] = mode
    path = tmp_path / record["image"]
    path.parent.mkdir(parents=True)
    image = Image.new("RGB", (20, 20), "white")
    image.save(path)
    control = FakeControl("Save", "ButtonControl", (0, 0, 20, 20), automation_id="save")
    monkeypatch.setattr(ui_cli, "resolve_location", Mock(return_value=control))
    monkeypatch.setattr(ui_cli, "_crop_control", Mock(return_value=image))
    monkeypatch.setattr(ui_cli, "_image_similarity", Mock(return_value=0.1))
    fallback = Mock(side_effect=AssertionError("Selected control was replaced using another image"))
    monkeypatch.setattr(ui_cli, "locate_template", fallback)
    monkeypatch.setattr(ui_cli, "_locate_record_text", fallback)
    with pytest.raises(ActionError) as raised:
        ui_cli._resolve_verified_control_in_context(tmp_path, record, object(), {}, image)
    assert raised.value.reason == f"{mode}_verification_failed"
    fallback.assert_not_called()


@pytest.mark.parametrize("mode", ["identity", "position"])
def test_recorder_passes_selected_mode_to_capture(monkeypatch, mode):
    from easy_uiauto import record
    from test_record import recorder
    value = recorder(monkeypatch)
    value.match_mode = mode
    control = FakeControl("1.0", "DataItemControl", (0, 0, 20, 20))
    capture = Mock(return_value=([{"Name": "Example"}, {"Name": "1.0"}], control))
    monkeypatch.setattr(record, "get_control_info", capture)
    monkeypatch.setattr(record, "CoInitialize", Mock())
    monkeypatch.setattr(record, "CoUninitialize", Mock())
    value.on_click(0, 10, record.Button.left, True)
    capture.assert_called_once_with(0, 10, match_mode=mode)
