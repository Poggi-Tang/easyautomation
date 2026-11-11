"""Recorded UIA names may differ from native captions, including after restart."""
from copy import deepcopy
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from easy_uiauto import ctrl, utils
from easy_uiauto.mcp import native_windows, scanner
from easy_uiauto.mcp.protocol import location_from_xpath
from easy_uiauto.mcp.results import ActionError
from test_mcp_scanner import FakeControl


@pytest.fixture
def app(monkeypatch):
    button = FakeControl("新建", "ButtonControl", (10, 10, 80, 40))
    window = FakeControl("SimuNPS", "WindowControl", (0, 0, 400, 300), [button])
    window.ClassName = "QtWindow"
    window.NativeWindowHandle = 0x123456789
    button.GetTopLevelControl = lambda: window
    entries = [dict(handle=window.NativeWindowHandle, title="SimuNPS-EMT", class_name="QtWindow", visible=True)]
    monkeypatch.setattr(native_windows, "enumerate_windows", lambda **_kw: entries)
    bind = Mock(side_effect=lambda handle: window if handle == window.NativeWindowHandle else None)
    monkeypatch.setattr(utils.uiautomation, "ControlFromHandle", bind)
    forbidden = Mock(side_effect=AssertionError("Must not enumerate desktop UIA or activate by title"))
    monkeypatch.setattr(utils.uiautomation, "GetRootControl", forbidden)
    monkeypatch.setattr(utils, "set_top_window", forbidden)
    location = location_from_xpath(utils.get_control_xpath(button))
    return window, button, entries, bind, location


def test_recording_and_all_lookup_entries_keep_uia_name_despite_native_caption(app):
    window, button, _entries, _bind, location = app
    original = deepcopy(location)
    assert location["WindowName"] == location["Xpath"][0]["Name"] == "SimuNPS"
    assert utils.find_control(location) is button
    assert utils.find_control_by_xpath(location["Xpath"]) is button
    assert scanner.resolve_location(location) is button
    action = utils.format_action_data_by_xpath("点击", 1, location["Xpath"], {"x": -1, "y": -1}, "")
    assert action["LOCATION"]["WindowName"] == "SimuNPS"
    assert utils.find_control(action) is button
    assert location == original


def test_replay_uses_current_resolved_handle_and_does_not_look_up_caption_again(app, monkeypatch):
    window, button, entries, _bind, location = app
    activate = Mock()
    monkeypatch.setattr(ctrl.Controller, "activate_window_handle", activate)
    with utils.strict_control_resolution(activate=True): assert utils.find_control(location) is button
    activate.assert_called_once_with(0x123456789, window)
    window.NativeWindowHandle = entries[0]["handle"] = 0x223456789
    entries[0]["title"] = "SimuNPS-EMT — changed caption"
    with utils.strict_control_resolution(activate=True): assert utils.find_control(location) is button
    assert activate.call_args.args == (0x223456789, window)


def test_legacy_click_locates_then_activates_resolved_window_once(app, monkeypatch):
    window, button, _entries, _bind, location = app
    actions, activate = Mock(), Mock()
    monkeypatch.setattr(ctrl, "show_ctrl_area", lambda _c: None)
    monkeypatch.setattr(ctrl, "correct_ctrl_position", lambda _c: (3, 4))
    monkeypatch.setattr(ctrl.Controller, "activate_window_handle", activate)
    monkeypatch.setattr(ctrl.Controller, "execute_resolved", actions)
    result = ctrl.Controller.left_click("新建", **location)
    assert "成功" in result
    activate.assert_called_once_with(window.NativeWindowHandle, window)
    actions.assert_called_once_with("click", control=button, x=3, y=4)


def test_native_title_hint_cannot_override_uia_identity(app, monkeypatch):
    window, button, entries, bind, location = app
    wrong = FakeControl("Other", "WindowControl", (0, 0, 400, 300))
    wrong.ClassName = window.ClassName
    entries.insert(0, dict(handle=77, title="SimuNPS", class_name="QtWindow", visible=True))
    bind.side_effect = lambda handle: wrong if handle == 77 else window
    assert utils.find_control(location) is button


@pytest.mark.parametrize("ordinal", [None, 3])
def test_alias_windows_preserve_ambiguity_and_invalid_index_errors(app, monkeypatch, ordinal):
    window, _button, entries, bind, location = app
    other = FakeControl("SimuNPS", "WindowControl", (0, 0, 400, 300)); other.ClassName = window.ClassName
    entries.append(dict(handle=77, title="Another caption", class_name="QtWindow", visible=True))
    bind.side_effect = lambda handle: other if handle == 77 else window
    activate = Mock(); monkeypatch.setattr(ctrl.Controller, "_activate_target", activate)
    if ordinal is not None: location["Xpath"][0]["foundIndex"] = ordinal
    with pytest.raises(ActionError) as failed:
        with utils.strict_control_resolution(activate=True): utils.find_control(location)
    assert failed.value.code == ("ambiguous_match" if ordinal is None else "window_invalid")
    if ordinal: assert failed.value.reason == "index_out_of_range"
    activate.assert_not_called()


def test_native_title_activation_keeps_its_explicit_caption_semantics(app, monkeypatch):
    window, _button, _entries, _bind, _location = app
    activate = Mock(); monkeypatch.setattr(ctrl.Controller, "activate_window_handle", activate)
    assert ctrl.Controller.activate_window("SimuNPS-EMT") is True
    activate.assert_called_once_with(window.NativeWindowHandle, control=window)
    assert ctrl.Controller.activate_window("") is False


def test_missing_handle_is_reported_instead_of_silently_skipping_activation():
    with pytest.raises(ActionError, match="window_invalid"):
        ctrl.Controller._activate_target(SimpleNamespace(NativeWindowHandle=0))


def test_class_shortlist_does_not_bind_unrelated_provider(app):
    _window, button, entries, bind, location = app
    entries.insert(0, dict(handle=77, title="Unresponsive unrelated app", class_name="OtherClass", visible=True))
    assert utils.find_control(location) is button
    bind.assert_called_once_with(entries[1]["handle"])


def test_uia_metadata_failure_is_not_misreported_as_missing_window(app):
    _window, _button, _entries, bind, location = app
    bind.side_effect = RuntimeError("Provider unavailable")
    with pytest.raises(ActionError) as failed: utils.find_control(location)
    assert failed.value.code == "action_failed"


@pytest.mark.parametrize("with_xpath", [False, True])
def test_scanner_cannot_bypass_uia_identity_rejection_with_native_title_fallback(app, monkeypatch, with_xpath):
    window, _button, entries, _bind, location = app
    entries[0]["title"] = "SimuNPS"
    window.Name = "Unrelated UIA window"
    if with_xpath: location["Xpath"][0].pop("ClassName")
    else: location["Xpath"] = []
    fallback = Mock(side_effect=AssertionError("Rejected identity must not use native title fallback"))
    monkeypatch.setattr(scanner, "_find_window", fallback)
    with pytest.raises(ActionError) as failed: scanner.resolve_location(location)
    assert failed.value.code == "window_invalid"
    assert failed.value.reason == "window_identity_mismatch"
    fallback.assert_not_called()
