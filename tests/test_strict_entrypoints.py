"""All entrypoints share strict selection and terminal lookup failures."""

from copy import deepcopy
from unittest.mock import Mock

import pytest
from PIL import Image

from easy_uiauto import utils
from easy_uiauto.mcp import scanner, ui_cli
from easy_uiauto.mcp.protocol import location_from_xpath, normalize_location
from easy_uiauto.mcp.results import ActionError
from test_mcp_scanner import FakeControl
from test_mcp_ui_cli import _record


@pytest.fixture
def tree(monkeypatch):
    first = FakeControl("Save", "ButtonControl", (10, 10, 30, 30))
    second = FakeControl("Save", "ButtonControl", (40, 10, 60, 30))
    window = FakeControl("Example", "WindowControl", (0, 0, 200, 200), [first, second])
    monkeypatch.setattr(utils, "_native_window_scope", lambda *_args: window)
    monkeypatch.setattr(scanner, "_find_window", lambda *_args: window)
    monkeypatch.setattr(utils, "set_top_window", Mock(side_effect=AssertionError("Read activated a window")))
    return window, first, second


def location(index=None, mode="identity"):
    step = {"ControlType": "ButtonControl", "Name": "Save", "match_mode": mode}
    if index is not None:
        step["foundIndex"] = index
    return location_from_xpath([{"ControlType": "WindowControl", "Name": "Example"}, step])


@pytest.mark.parametrize("entry", ["observe", "find", "legacy_flag", "xpath", "scan"])
@pytest.mark.parametrize("index,code", [(-1, "invalid_argument"),
                                       (True, "invalid_argument"), (False, "invalid_argument"),
                                       (1.5, "invalid_argument"),
                                       ("bad", "invalid_argument")])
def test_every_entry_rejects_invalid_selection(tree, entry, index, code):
    window, *_ = tree
    value = location(index)
    calls = {
        "observe": lambda: utils.observe_control(value),
        "find": lambda: utils.find_control(value),
        "legacy_flag": lambda: utils.find_control(value, strict=False),
        "xpath": lambda: utils.find_control_by_xpath(value["Xpath"]),
        "scan": lambda: scanner.resolve_location(value, window=window),
    }
    with pytest.raises(ActionError) as raised:
        calls[entry]()
    assert raised.value.code == code
    if index == 3:
        assert raised.value.reason == "index_out_of_range"


@pytest.mark.parametrize("index", [1, 2])
def test_exact_indices_agree_across_public_and_scanner_entrypoints(tree, index):
    window, first, second = tree
    value = location(index, "position")
    expected = (first, second)[index - 1]
    assert utils.find_control(value) is expected
    assert utils.find_control_by_xpath(value["Xpath"]) is expected
    assert scanner.resolve_location(value, window=window) is expected


def test_unique_id_cannot_ignore_explicit_invalid_index(tree):
    window, first, _ = tree
    first.AutomationId = "unique"
    value = location(3, "position")
    value["AutomationId"] = value["Xpath"][-1]["AutomationId"] = "unique"
    with pytest.raises(ActionError) as raised:
        scanner.resolve_location(value, window=window)
    assert raised.value.reason == "index_out_of_range"


def test_top_level_index_cannot_be_bypassed_by_xpath_leaf(tree):
    window, *_ = tree
    value = location(mode="position")
    value["foundIndex"] = 3
    with pytest.raises(ActionError) as raised:
        scanner.resolve_location(value, window=window)
    assert raised.value.reason == "index_out_of_range"


def test_scanner_does_not_choose_between_same_named_visible_windows(monkeypatch):
    monkeypatch.setattr(scanner.native_windows, "matching_windows", lambda *_a, **_kw: [
        {"visible": True, "handle": 1}, {"visible": True, "handle": 2}])
    bind = Mock(side_effect=AssertionError("Ambiguous window must not bind UIA"))
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", bind)
    with pytest.raises(ActionError, match="ambiguous_match"):
        scanner._find_window("Example")
    bind.assert_not_called()


def test_scanner_preserves_shared_invalid_argument_code(tree):
    window, *_ = tree
    with pytest.raises(ActionError) as raised:
        scanner.resolve_location({"Xpath": "not an array"}, window=window)
    assert raised.value.code == "invalid_argument"


def test_scanner_resolves_workerw_context_with_shared_core(monkeypatch):
    icon = FakeControl("App", "ListItemControl", (10, 10, 40, 40))
    shell = FakeControl("", "PaneControl", (0, 0, 200, 200), [icon])
    shell.ClassName = "WorkerW"
    desktop = FakeControl("Program Manager", "WindowControl", (0, 0, 200, 200))
    desktop.ClassName = "Progman"
    monkeypatch.setattr(scanner, "_desktop_shell_roots", lambda *_a: [desktop, shell])
    value = {"WindowName": "Program Manager", "Xpath": [
        {"ControlType": "PaneControl", "ClassName": "WorkerW"},
        {"ControlType": "ListItemControl", "Name": "App"}]}
    assert scanner.resolve_location(value, window=desktop) is icon


def test_cached_prefixes_are_scoped_to_window_and_leaf_is_always_rechecked(tree):
    window, first, second = tree
    group = FakeControl("Tools", "GroupControl", (0, 0, 100, 100), [first])
    window._children = [group]
    value = location()
    value["Xpath"].insert(1, {"ControlType": "GroupControl", "Name": "Tools"})
    cache = {}
    assert scanner.resolve_location(value, window=window, prefix_cache=cache) is first
    group._children[:] = [second]
    assert scanner.resolve_location(value, window=window, prefix_cache=cache) is second
    other = FakeControl("Example", "WindowControl", (0, 0, 100, 100), [])
    with pytest.raises(ActionError, match="control_not_found"):
        scanner.resolve_location(value, window=other, prefix_cache=cache)


def test_legacy_xpath_literals_are_data_and_cannot_execute(tmp_path):
    value = location(1)
    assert normalize_location({"Xpath": repr(value["Xpath"])})["Xpath"] == value["Xpath"]
    marker = tmp_path / "should-not-exist"
    with pytest.raises(ValueError):
        normalize_location({"Xpath": f"__import__('pathlib').Path({str(marker)!r}).touch()"})
    assert not marker.exists()


@pytest.mark.parametrize("code,reason", [("ambiguous_match", ""), ("invalid_argument", ""),
                                        ("control_not_found", "index_out_of_range"),
                                        ("action_failed", "")])
def test_terminal_lookup_failure_never_uses_template_or_ocr(monkeypatch, tmp_path, code, reason):
    record = deepcopy(_record())
    path = tmp_path / record["image"]
    path.parent.mkdir(parents=True)
    image = Image.new("RGB", (10, 10), "white")
    image.save(path)
    monkeypatch.setattr(ui_cli, "resolve_location", Mock(side_effect=ActionError(code, "lookup failed", reason=reason)))
    template = Mock(side_effect=AssertionError("Terminal failure used image fallback"))
    ocr = Mock(side_effect=AssertionError("Terminal failure used OCR fallback"))
    monkeypatch.setattr(ui_cli, "locate_template", template)
    monkeypatch.setattr(ui_cli, "_locate_record_text", ocr)
    with pytest.raises(ActionError) as raised:
        ui_cli._resolve_verified_control_in_context(tmp_path, record, object(), {}, image)
    assert raised.value.code == code
    template.assert_not_called()
    ocr.assert_not_called()
