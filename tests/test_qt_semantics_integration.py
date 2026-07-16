from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path

import pytest
import uiautomation

from easy_uiauto.ctrl import Controller

RUN_WINDOWS_INTEGRATION = os.environ.get("EASY_UIAUTO_RUN_WINDOWS_TESTS") == "1"

pytestmark = [
    pytest.mark.windows_integration,
    pytest.mark.skipif(
        not RUN_WINDOWS_INTEGRATION,
        reason="set EASY_UIAUTO_RUN_WINDOWS_TESTS=1 to run desktop integration tests",
    ),
]

FIXTURE = Path(__file__).parent / "fixtures" / "qt_fixture.py"


def _launch_fixture(
    title: str,
    *,
    tree_expanded: bool = False,
    editable_combo: bool = False,
):
    environment = {
        **os.environ,
        "PYTHONUNBUFFERED": "1",
        "QT_QPA_PLATFORM": "windows",
    }
    if tree_expanded:
        environment["EASY_UIAUTO_QT_TREE_EXPANDED"] = "1"
    if editable_combo:
        environment["EASY_UIAUTO_QT_EDITABLE_COMBO"] = "1"
    process = subprocess.Popen(
        [sys.executable, str(FIXTURE), title],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=environment,
    )
    assert process.stdout is not None
    output: queue.Queue[str] = queue.Queue()
    threading.Thread(target=lambda: output.put(process.stdout.readline()), daemon=True).start()
    try:
        ready = output.get(timeout=10).strip()
    except queue.Empty as error:
        process.terminate()
        stderr = process.stderr.read() if process.stderr else ""
        raise AssertionError(f"fixture did not become ready: {stderr}") from error
    assert ready == "READY"
    return process


def _wait_for_window(title: str, timeout: float = 10):
    window = uiautomation.WindowControl(Name=title, searchDepth=1)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if window.Exists(0):
            return window
        time.sleep(0.1)
    raise AssertionError(f"window did not appear: {title}")


def _wait_for_status(window, name: str, timeout: float = 3):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        status = window.TextControl(Name=name, searchDepth=15)
        if status.Exists(0):
            return status
        time.sleep(0.05)
    raise AssertionError(f"Qt application state did not reach {name!r}")


def _stop_fixture(process, window) -> None:
    try:
        import win32con
        import win32gui

        handle = int(window.NativeWindowHandle)
        if handle and win32gui.IsWindow(handle):
            win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
        process.wait(timeout=5)
    except (OSError, subprocess.TimeoutExpired):
        process.terminate()
        try:
            process.wait(timeout=3)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=3)


def _select_action_args(title: str, control):
    return (
        "Qt semantic selection",
        title,
        control.Name,
        control.ClassName,
        control.ControlTypeName,
        1,
        control.AutomationId,
        [],
        "",
        {},
    )


def _named_select_action_args(title: str, control, item_name: str):
    arguments = list(_select_action_args(title, control))
    arguments[-1] = {"选择项": item_name}
    return tuple(arguments)


def _expand_action_args(title: str, control, should_expand: bool):
    arguments = list(_select_action_args(title, control))
    arguments[0] = "Qt semantic expand/collapse"
    arguments[-1] = {"展开": should_expand}
    return tuple(arguments)


def _input_action_args(title: str, control, value: str):
    arguments = list(_select_action_args(title, control))
    arguments[0] = "Qt editable combo input"
    arguments[-1] = {"输入文本": value}
    return tuple(arguments)


def test_qt_tree_table_and_tab_selection_update_real_widget_state():
    title = f"Easy UIAuto Qt semantics {uuid.uuid4().hex[:8]}"
    process = _launch_fixture(title)
    window = _wait_for_window(title)
    try:
        targets = (
            (window.TreeItemControl(Name="Root", searchDepth=12), "tree-selection:Root"),
            (window.DataItemControl(Name="row-a", searchDepth=12), "table:row-a"),
            (window.TabItemControl(Name="Second", searchDepth=12), "tab:Second"),
        )
        for target, expected_status in targets:
            assert target.Exists(2)
            assert target.ProcessId == window.ProcessId
            message = Controller.select_control(*_select_action_args(title, target))
            assert "成功" in message
            _wait_for_status(window, expected_status)

            refreshed = uiautomation.Control.CreateControlFromControl(target)
            selection = refreshed.GetPattern(uiautomation.PatternId.SelectionItemPattern)
            assert selection is not None
            assert selection.IsSelected

        assert not window.TextControl(Name="First page", searchDepth=12).Exists(0)
        assert window.TextControl(Name="Second page", searchDepth=12).Exists(0)
    finally:
        _stop_fixture(process, window)


def test_qt_editable_combobox_inner_edit_supports_verified_value_input():
    title = f"Easy UIAuto Qt editable combo {uuid.uuid4().hex[:8]}"
    process = _launch_fixture(title, editable_combo=True)
    window = _wait_for_window(title)
    try:
        combo = window.ComboBoxControl(searchDepth=12)
        edit = combo.EditControl(searchDepth=3)
        assert edit.Exists(2)
        assert edit.ProcessId == window.ProcessId

        message = Controller.input_text(*_input_action_args(title, edit, "Beta"))
        assert "成功" in message
        _wait_for_status(window, "combo:Beta")

        refreshed_combo = window.ComboBoxControl(searchDepth=12)
        refreshed_edit = refreshed_combo.EditControl(searchDepth=3)
        assert (
            refreshed_edit.GetPattern(uiautomation.PatternId.ValuePattern).Value == "Beta"
        )
        assert (
            refreshed_combo.GetPattern(uiautomation.PatternId.ValuePattern).Value == "Beta"
        )
    finally:
        _stop_fixture(process, window)


def test_qt_combobox_provider_success_is_not_a_committed_selection():
    title = f"Easy UIAuto Qt combo probe {uuid.uuid4().hex[:8]}"
    process = _launch_fixture(title)
    window = _wait_for_window(title)
    try:
        combo = window.ComboBoxControl(searchDepth=12)
        assert combo.Exists(2)
        assert combo.ProcessId == window.ProcessId
        assert combo.GetPattern(uiautomation.PatternId.ValuePattern).Value == "Alpha"

        value_pattern = combo.GetPattern(uiautomation.PatternId.ValuePattern)
        assert value_pattern.SetValue("Beta", waitTime=0)

        expand = combo.GetPattern(uiautomation.PatternId.ExpandCollapsePattern)
        assert expand.Expand(waitTime=0)
        beta = window.ListItemControl(Name="Beta", searchDepth=12)
        assert beta.Exists(2)
        assert beta.ProcessId == window.ProcessId
        selection = beta.GetPattern(uiautomation.PatternId.SelectionItemPattern)
        assert selection.Select(waitTime=0)
        assert beta.GetPattern(uiautomation.PatternId.InvokePattern).Invoke(waitTime=0)
        assert beta.GetPattern(
            uiautomation.PatternId.LegacyIAccessiblePattern
        ).DoDefaultAction(waitTime=0)
        time.sleep(0.2)

        refreshed_combo = window.ComboBoxControl(searchDepth=12)
        assert refreshed_combo.ProcessId == window.ProcessId
        actual_value = refreshed_combo.GetPattern(uiautomation.PatternId.ValuePattern).Value
        if actual_value == "Beta":
            pytest.skip("this Qt provider now commits semantic combo-box selection")
        assert actual_value == "Alpha"
        _wait_for_status(window, "combo:Alpha")
        assert not window.TextControl(Name="combo:Beta", searchDepth=12).Exists(0)

        message = Controller.select_control(
            *_named_select_action_args(title, refreshed_combo, "Beta")
        )
        assert "异常" in message
        assert (
            window.ComboBoxControl(searchDepth=12)
            .GetPattern(uiautomation.PatternId.ValuePattern)
            .Value
            == "Alpha"
        )
    finally:
        _stop_fixture(process, window)


@pytest.mark.parametrize(
    ("initially_expanded", "operation", "initial_state", "unchanged_status"),
    [
        (False, "Expand", 0, "tree:collapsed"),
        (True, "Collapse", 1, "tree:expanded"),
    ],
)
def test_qt_tree_expand_provider_result_requires_postcondition_verification(
    initially_expanded: bool,
    operation: str,
    initial_state: int,
    unchanged_status: str,
):
    title = f"Easy UIAuto Qt tree probe {uuid.uuid4().hex[:8]}"
    process = _launch_fixture(title, tree_expanded=initially_expanded)
    window = _wait_for_window(title)
    try:
        root = window.TreeItemControl(Name="Root", searchDepth=12)
        pattern = root.GetPattern(uiautomation.PatternId.ExpandCollapsePattern)
        assert int(pattern.ExpandCollapseState) == initial_state
        assert getattr(pattern, operation)(waitTime=0)
        time.sleep(0.2)

        refreshed_root = window.TreeItemControl(Name="Root", searchDepth=12)
        refreshed_state = int(
            refreshed_root.GetPattern(
                uiautomation.PatternId.ExpandCollapsePattern
            ).ExpandCollapseState
        )
        target_state = 1 - initial_state
        if refreshed_state == target_state:
            pytest.skip(f"this Qt provider now performs {operation.lower()}")
        assert refreshed_state == initial_state
        _wait_for_status(window, unchanged_status)

        message = Controller.expand_collapse_control(
            *_expand_action_args(title, refreshed_root, not initially_expanded)
        )
        assert "异常" in message
        final_root = window.TreeItemControl(Name="Root", searchDepth=12)
        assert (
            int(
                final_root.GetPattern(
                    uiautomation.PatternId.ExpandCollapsePattern
                ).ExpandCollapseState
            )
            == initial_state
        )
    finally:
        _stop_fixture(process, window)
