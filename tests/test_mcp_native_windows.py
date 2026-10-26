"""Native inventory and learning lookup must never enumerate UIA desktop roots."""

import ctypes
from ctypes import wintypes
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from easy_uiauto.mcp import diagnostics, interaction_learning, native_windows, scanner


@pytest.mark.parametrize("minimized", [False, True])
def test_activation_preserves_maximized_layout_and_only_restores_minimized(
        monkeypatch, minimized):
    api = SimpleNamespace(IsWindow=lambda _h: True, IsIconic=lambda _h: minimized,
                          ShowWindow=Mock(), SetForegroundWindow=lambda _h: True,
                          GetForegroundWindow=lambda: 42)
    monkeypatch.setattr(native_windows, "_api", lambda: (api, None))
    monkeypatch.setattr(scanner.time, "sleep", lambda _s: None)
    scanner._activate_window(SimpleNamespace(NativeWindowHandle=42))
    if minimized:
        api.ShowWindow.assert_called_once_with(42, 9)
    else:
        api.ShowWindow.assert_not_called()


def test_activation_denied_is_reported_even_if_focus_call_does_not_raise(monkeypatch):
    api = SimpleNamespace(IsWindow=lambda _h: True, IsIconic=lambda _h: False,
                          ShowWindow=Mock(), SetForegroundWindow=lambda _h: False,
                          GetForegroundWindow=lambda: 7)
    monkeypatch.setattr(native_windows, "_api", lambda: (api, None))
    monkeypatch.setattr(scanner.time, "sleep", lambda _s: None)
    with pytest.raises(RuntimeError, match="foreground access was denied"):
        scanner._activate_window(SimpleNamespace(NativeWindowHandle=42, SetFocus=Mock()))


def _entry(handle, title="Editor", class_name="EditorClass", visible=True):
    return dict(handle=handle, title=title, class_name=class_name, visible=visible,
                enabled=True, process_id=8, owner_handle=0, z_order=0, foreground=False,
                bounds=dict(left=-800, top=-100, right=0, bottom=500, width=800, height=600))


def _forbid_root():
    raise AssertionError("UIA desktop root must not be accessed")


def test_native_enumeration_retains_full_width_handles_and_native_metadata(monkeypatch):
    handle = 0x123456789

    def window_rect(_handle, output):
        output._obj.left, output._obj.top = -800, -100
        output._obj.right, output._obj.bottom = 0, 500
        return True

    def process_id(_handle, output):
        output._obj.value = 8
        return 123

    def text(_handle, output, _length):
        output.value = "Editor"
        return 6

    def class_name(_handle, output, _length):
        output.value = "EditorClass"
        return 11

    api = SimpleNamespace(
        GetForegroundWindow=lambda: handle,
        EnumWindows=lambda callback, _data: callback(handle, 0),
        IsWindow=lambda _handle: True,
        IsWindowVisible=lambda _handle: True,
        IsWindowEnabled=lambda _handle: True,
        GetWindowTextLengthW=lambda _handle: 6,
        GetWindowTextW=text, GetClassNameW=class_name, GetWindowRect=window_rect,
        GetWindowThreadProcessId=process_id, GetWindow=lambda _handle, _flag: 0)
    monkeypatch.setattr(native_windows, "_api", lambda: (api, lambda callback: callback))
    result = native_windows.enumerate_windows()
    assert result == [{**_entry(handle), "foreground": True}]


def test_enum_callback_errors_are_not_silently_swallowed(monkeypatch):
    api = SimpleNamespace(GetForegroundWindow=lambda: 0,
                          EnumWindows=lambda callback, _data: callback(99, 0),
                          IsWindow=Mock(side_effect=OSError("native failure")))
    monkeypatch.setattr(native_windows, "_api", lambda: (api, lambda callback: callback))
    with pytest.raises(RuntimeError, match="enumerate native"):
        native_windows.enumerate_windows()


def test_private_native_signatures_do_not_modify_shared_pygetwindow_callback():
    shared = ctypes.windll.user32.EnumWindows
    signature = shared.argtypes
    api, _callback = native_windows._api()
    assert api.EnumWindows is not shared
    assert shared.argtypes is signature
    assert api.GetWindow.restype is wintypes.HWND


def test_exact_title_precedes_partial_and_handle_tracks_a_renamed_window(monkeypatch):
    entries = [_entry(1, "Editor - other"), _entry(2, "Editor")]
    monkeypatch.setattr(native_windows, "enumerate_windows", lambda: entries)
    assert native_windows.matching_windows("Editor", partial=True) == [entries[1]]
    assert native_windows.matching_windows("EDITOR", partial=True) == entries
    assert native_windows.matching_windows("old title", handle=2) == [entries[1]]
    assert native_windows.matching_windows("missing") == []


def test_scanner_binds_only_selected_handle_without_touching_other_uia_providers(monkeypatch):
    target = object()
    monkeypatch.setattr(scanner.uiautomation, "GetRootControl", _forbid_root)
    monkeypatch.setattr(native_windows, "enumerate_windows", lambda: [
        _entry(1, "Unresponsive app"), _entry(2, "Editor - other"), _entry(3)])
    bind = Mock(side_effect=lambda handle: target if handle == 3 else _forbid_root())
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", bind)
    assert scanner._find_window("Editor") is target
    bind.assert_called_once_with(3)


def test_scanner_missing_target_does_not_fall_back_to_desktop_uia(monkeypatch):
    monkeypatch.setattr(scanner.uiautomation, "GetRootControl", _forbid_root)
    monkeypatch.setattr(native_windows, "enumerate_windows", lambda: [])
    bind = Mock(side_effect=_forbid_root)
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", bind)
    with pytest.raises(RuntimeError, match="Window not found"):
        scanner._find_window("Missing")
    with pytest.raises(ValueError, match="empty"):
        scanner._find_window(" ")
    bind.assert_not_called()


def test_interaction_snapshot_reads_only_native_metadata(monkeypatch):
    monkeypatch.setattr(interaction_learning.uiautomation, "GetRootControl", _forbid_root)
    monkeypatch.setattr(native_windows, "enumerate_windows", lambda **_kw: [_entry(7)])
    result = interaction_learning.snapshot_top_level_windows()
    assert len(result) == 1 and result[0]["handle"] == 7
    assert result[0]["metadata_source"] == "win32"
    assert result[0]["rect"]["left"] == -800


def test_popup_learning_binds_only_added_window_handles(monkeypatch):
    root = object()
    monkeypatch.setattr(interaction_learning.uiautomation, "GetRootControl", _forbid_root)
    bind = Mock(return_value=root)
    monkeypatch.setattr(interaction_learning.uiautomation, "ControlFromHandle", bind)
    monkeypatch.setattr(scanner, "_walk_controls", lambda *_args: [])
    interaction_learning._popup_controls({"added": [{"handle": 41}]})
    bind.assert_called_once_with(41)


def test_readiness_probe_binds_one_visible_window_without_enumerating_desktop(monkeypatch):
    from contextlib import nullcontext

    import uiautomation

    monkeypatch.setattr(native_windows, "enumerate_windows", lambda **_kw: [_entry(7)])
    monkeypatch.setattr(uiautomation, "UIAutomationInitializerInThread", lambda: nullcontext())
    monkeypatch.setattr(uiautomation, "GetRootControl", _forbid_root)
    bind = Mock(return_value=SimpleNamespace(ControlTypeName="WindowControl"))
    monkeypatch.setattr(uiautomation, "ControlFromHandle", bind)
    assert "no child enumeration" in diagnostics._test_uia()
    bind.assert_called_once_with(7)
