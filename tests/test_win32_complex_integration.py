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
from easy_uiauto.utils import (
    find_control,
    find_control_by_xpath,
    get_control_xpath,
    package_location,
)

RUN_WINDOWS_INTEGRATION = os.environ.get("EASY_UIAUTO_RUN_WINDOWS_TESTS") == "1"
if RUN_WINDOWS_INTEGRATION:
    import win32gui
else:
    win32gui = None


pytestmark = [
    pytest.mark.windows_integration,
    pytest.mark.skipif(
        not RUN_WINDOWS_INTEGRATION,
        reason="set EASY_UIAUTO_RUN_WINDOWS_TESTS=1 to run desktop integration tests",
    ),
]


FIXTURE = Path(__file__).parent / "fixtures" / "win32_fixture.py"
LVM_FIRST = 0x1000
LVM_GETNEXTITEM = LVM_FIRST + 12
LVM_GETTOPINDEX = LVM_FIRST + 39
LVNI_SELECTED = 0x0002
TVM_FIRST = 0x1100
TVM_GETNEXTITEM = TVM_FIRST + 10
TVGN_ROOT = 0x0000
TVGN_CARET = 0x0009
TCM_FIRST = 0x1300
TCM_GETCURSEL = TCM_FIRST + 11


def _launch_fixture(title):
    process = subprocess.Popen(
        [sys.executable, str(FIXTURE), title],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env={**os.environ, "PYTHONUNBUFFERED": "1"},
    )
    assert process.stdout is not None
    output = queue.Queue()
    threading.Thread(target=lambda: output.put(process.stdout.readline()), daemon=True).start()
    try:
        ready = output.get(timeout=10).strip()
    except queue.Empty as error:
        process.terminate()
        stderr = process.stderr.read() if process.stderr else ""
        raise AssertionError(f"fixture did not become ready: {stderr}") from error
    if ready != "READY":
        stderr = process.stderr.read() if process.stderr else ""
        raise AssertionError(f"fixture failed before ready: {stderr}")
    return process


def _wait_until(predicate, description, timeout=5):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        value = predicate()
        if value:
            return value
        time.sleep(0.05)
    raise AssertionError(f"timed out waiting for {description}")


def _location(title, name, control_type, automation_id=""):
    return package_location(title, name, "", control_type, 1, automation_id, [], "", {})


def _action_args(title, name, control_type, parameters, automation_id=""):
    return (
        "complex Win32 integration action",
        title,
        name,
        "",
        control_type,
        1,
        automation_id,
        [],
        "",
        parameters,
    )


def _assert_xpath_roundtrip(control):
    xpath = get_control_xpath(control)
    assert xpath
    replayed = find_control_by_xpath(xpath, use_cache=False)
    assert replayed is not None
    assert uiautomation.ControlsAreSame(control, replayed)


def test_native_common_controls_have_verified_semantic_state_transitions():
    title = f"Easy UIAuto Win32 Complex {uuid.uuid4().hex[:8]}"
    process = _launch_fixture(title)
    try:
        window = _wait_until(
            lambda: (
                candidate
                if (candidate := uiautomation.WindowControl(Name=title, searchDepth=1)).Exists(0)
                else None
            ),
            "fixture window",
            timeout=10,
        )
        window.SetActive()

        list_view = find_control(_location(title, "", "ListControl", "106"))
        assert list_view is not None
        assert list_view.ClassName == "SysListView32"
        _assert_xpath_roundtrip(list_view)

        list_item = find_control(_location(title, "List item 23", "ListItemControl"))
        assert list_item is not None
        _assert_xpath_roundtrip(list_item)
        message = Controller.select_control(
            *_action_args(title, "List item 23", "ListItemControl", {})
        )
        assert "成功" in message
        selected_pattern = list_item.GetPattern(uiautomation.PatternId.SelectionItemPattern)
        assert selected_pattern is not None
        _wait_until(lambda: selected_pattern.IsSelected, "list item selection")
        assert win32gui.SendMessage(
            int(list_view.NativeWindowHandle), LVM_GETNEXTITEM, -1, LVNI_SELECTED
        ) == 23

        first_item = find_control(_location(title, "List item 01", "ListItemControl"))
        assert first_item is not None
        message = Controller.select_control(
            *_action_args(title, "List item 01", "ListItemControl", {})
        )
        assert "成功" in message
        first_selection = first_item.GetPattern(uiautomation.PatternId.SelectionItemPattern)
        assert first_selection is not None
        _wait_until(lambda: first_selection.IsSelected, "replacement list item selection")
        assert not selected_pattern.IsSelected
        assert win32gui.SendMessage(
            int(list_view.NativeWindowHandle), LVM_GETNEXTITEM, -1, LVNI_SELECTED
        ) == 1

        tree = find_control(_location(title, "", "TreeControl", "107"))
        root = find_control(_location(title, "Root node", "TreeItemControl"))
        assert tree is not None
        assert root is not None
        assert tree.ClassName == "SysTreeView32"
        _assert_xpath_roundtrip(root)
        root_expand = root.GetPattern(uiautomation.PatternId.ExpandCollapsePattern)
        assert root_expand is not None
        assert int(root_expand.ExpandCollapseState) == 0

        message = Controller.expand_collapse_control(
            *_action_args(title, "Root node", "TreeItemControl", {"展开": True})
        )
        assert "成功" in message
        _wait_until(
            lambda: int(root_expand.ExpandCollapseState) == 1,
            "tree root expansion",
        )
        child = _wait_until(
            lambda: find_control(_location(title, "Child node B", "TreeItemControl")),
            "expanded child tree item",
        )
        message = Controller.select_control(
            *_action_args(title, "Child node B", "TreeItemControl", {})
        )
        assert "成功" in message
        child_selection = child.GetPattern(uiautomation.PatternId.SelectionItemPattern)
        assert child_selection is not None
        _wait_until(lambda: child_selection.IsSelected, "tree child selection")
        tree_handle = int(tree.NativeWindowHandle)
        root_handle = win32gui.SendMessage(tree_handle, TVM_GETNEXTITEM, TVGN_ROOT, 0)
        selected_tree_handle = win32gui.SendMessage(tree_handle, TVM_GETNEXTITEM, TVGN_CARET, 0)
        assert root_handle
        assert selected_tree_handle
        assert selected_tree_handle != root_handle

        message = Controller.expand_collapse_control(
            *_action_args(title, "Root node", "TreeItemControl", {"展开": False})
        )
        assert "成功" in message
        _wait_until(
            lambda: int(root_expand.ExpandCollapseState) == 0,
            "tree root collapse",
        )

        tabs = find_control(_location(title, "", "TabControl", "108"))
        advanced = find_control(_location(title, "Advanced", "TabItemControl"))
        general = find_control(_location(title, "General", "TabItemControl"))
        assert tabs is not None
        assert advanced is not None
        assert general is not None
        assert tabs.ClassName == "SysTabControl32"
        _assert_xpath_roundtrip(advanced)
        message = Controller.select_control(
            *_action_args(title, "Advanced", "TabItemControl", {})
        )
        assert "成功" in message
        advanced_selection = advanced.GetPattern(uiautomation.PatternId.SelectionItemPattern)
        general_selection = general.GetPattern(uiautomation.PatternId.SelectionItemPattern)
        assert advanced_selection is not None
        assert general_selection is not None
        _wait_until(lambda: advanced_selection.IsSelected, "tab selection")
        assert not general_selection.IsSelected
        assert win32gui.SendMessage(int(tabs.NativeWindowHandle), TCM_GETCURSEL, 0, 0) == 1

        scroll = list_view.GetPattern(uiautomation.PatternId.ScrollPattern)
        assert scroll is not None
        assert scroll.VerticallyScrollable
        assert scroll.SetScrollPercent(-1, 0, waitTime=0)
        list_handle = int(list_view.NativeWindowHandle)
        _wait_until(
            lambda: win32gui.SendMessage(list_handle, LVM_GETTOPINDEX, 0, 0) == 0,
            "list scroll reset",
        )
        assert scroll.SetScrollPercent(-1, 100, waitTime=0)
        _wait_until(
            lambda: win32gui.SendMessage(list_handle, LVM_GETTOPINDEX, 0, 0) > 0,
            "native list scroll position",
        )
        assert scroll.VerticalScrollPercent > 0
        assert scroll.SetScrollPercent(-1, 0, waitTime=0)
        _wait_until(
            lambda: win32gui.SendMessage(list_handle, LVM_GETTOPINDEX, 0, 0) == 0,
            "final list scroll restoration",
        )
    finally:
        process.terminate()
        process.wait(timeout=5)
