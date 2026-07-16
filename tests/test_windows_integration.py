from __future__ import annotations

import os
import queue
import subprocess
import sys
import tempfile
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
    import win32con
    import win32gui
else:
    win32con = None
    win32gui = None

pytestmark = [
    pytest.mark.windows_integration,
    pytest.mark.skipif(
        not RUN_WINDOWS_INTEGRATION,
        reason="set EASY_UIAUTO_RUN_WINDOWS_TESTS=1 to run desktop integration tests",
    ),
]


FIXTURES = Path(__file__).parent / "fixtures"


def _launch_fixture(script_name, title):
    process = subprocess.Popen(
        [sys.executable, str(FIXTURES / script_name), title],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env={**os.environ, "PYTHONUNBUFFERED": "1", "QT_QPA_PLATFORM": "windows"},
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


def _wait_for_window(title, timeout=10):
    window = uiautomation.WindowControl(Name=title, searchDepth=1)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if window.Exists(0):
            return window
        time.sleep(0.1)
    raise AssertionError(f"window did not appear: {title}")


def _wait_for_control(control, description, timeout=10):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if control.Exists(0):
            return control
        time.sleep(0.1)
    raise AssertionError(f"control did not appear: {description}")


def _location(title, name, control_type, class_name="", automation_id=""):
    return package_location(
        title,
        name,
        class_name,
        control_type,
        1,
        automation_id,
        [],
        "",
        {},
    )


def _action_args(title, name, control_type, parameters, class_name="", automation_id=""):
    return (
        "integration action",
        title,
        name,
        class_name,
        control_type,
        1,
        automation_id,
        [],
        "",
        parameters,
    )


@pytest.mark.parametrize(
    ("script_name", "framework"),
    [("win32_fixture.py", "Win32"), ("qt_fixture.py", "Qt")],
)
def test_real_desktop_input_click_and_xpath_roundtrip(script_name, framework):
    title = f"Easy UIAuto {framework} {uuid.uuid4().hex[:8]}"
    process = _launch_fixture(script_name, title)
    try:
        window = _wait_for_window(title)
        window.SetActive()
        edit = find_control(_location(title, "Input" if framework == "Qt" else "", "EditControl"))
        assert edit is not None
        xpath = get_control_xpath(edit)
        assert xpath
        assert find_control(
            package_location(title, edit.Name, edit.ClassName, edit.ControlTypeName, 1,
                             edit.AutomationId, xpath, "", {})
        ) is not None

        text = f"{framework.lower()}-input"
        message = Controller.input_text(
            *_action_args(title, edit.Name, "EditControl", {"输入文本": text}, edit.ClassName,
                          edit.AutomationId)
        )
        assert "成功" in message
        assert edit.GetPattern(uiautomation.PatternId.ValuePattern).Value == text

        button = find_control(_location(title, "Apply", "ButtonControl"))
        assert button is not None
        message = Controller.left_click(
            *_action_args(title, "Apply", "ButtonControl", {"x": -1, "y": -1})
        )
        assert "成功" in message

        deadline = time.monotonic() + 3
        expected = f"applied:{text}"
        while time.monotonic() < deadline:
            status = find_control(_location(title, expected, "TextControl"))
            if status is not None:
                break
            time.sleep(0.1)
        else:
            raise AssertionError(f"status did not update to {expected}")

        checkbox = find_control(_location(title, "Enabled", "CheckBoxControl"))
        assert checkbox is not None
        message = Controller.toggle_control(
            *_action_args(title, "Enabled", "CheckBoxControl", {"选中": True})
        )
        assert "成功" in message
        assert checkbox.GetPattern(uiautomation.PatternId.TogglePattern).ToggleState == 1

        combo = find_control(_location(title, "", "ComboBoxControl"))
        assert combo is not None
        if framework == "Win32":
            message = Controller.select_control(
                *_action_args(
                    title,
                    combo.Name,
                    "ComboBoxControl",
                    {"选择项": "Beta"},
                    combo.ClassName,
                    combo.AutomationId,
                )
            )
            assert "成功" in message
            assert combo.GetPattern(uiautomation.PatternId.ValuePattern).Value == "Beta"
        else:
            assert find_control(_location(title, "", "TreeControl")) is not None
            assert find_control(_location(title, "", "TableControl")) is not None
            assert find_control(_location(title, "", "TabControl")) is not None
    finally:
        process.terminate()
        process.wait(timeout=5)


def test_real_explorer_selection_and_xpath_roundtrip():
    with tempfile.TemporaryDirectory(prefix="easy-uiauto-explorer-") as temp_path:
        sandbox = Path(temp_path)
        (sandbox / "alpha.txt").write_text("alpha", encoding="utf-8")
        (sandbox / "beta.txt").write_text("beta", encoding="utf-8")
        subprocess.Popen(["explorer.exe", str(sandbox)])
        window = _wait_for_window(sandbox.name)
        handle = int(window.NativeWindowHandle)
        try:
            assert window.ClassName == "CabinetWClass"
            for index, file_name in enumerate(("alpha.txt", "beta.txt")):
                item = _wait_for_control(
                    window.ListItemControl(Name=file_name, searchDepth=15),
                    file_name,
                )
                assert item.AutomationId == str(index)
                xpath = get_control_xpath(item)
                assert xpath
                replayed = find_control(
                    package_location(
                        sandbox.name,
                        item.Name,
                        item.ClassName,
                        item.ControlTypeName,
                        1,
                        item.AutomationId,
                        xpath,
                        "",
                        {},
                    )
                )
                assert replayed is not None
                assert uiautomation.ControlsAreSame(item, replayed)

            alpha = window.ListItemControl(Name="alpha.txt", searchDepth=15)
            selection = alpha.GetPattern(uiautomation.PatternId.SelectionItemPattern)
            assert selection is not None
            assert selection.Select(waitTime=0)
            assert selection.IsSelected
        finally:
            if win32gui.IsWindow(handle):
                win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
            deadline = time.monotonic() + 5
            while win32gui.IsWindow(handle) and time.monotonic() < deadline:
                time.sleep(0.1)


@pytest.mark.parametrize("class_name", ["WorkerW", "Shell_TrayWnd"])
def test_desktop_shell_xpath_roundtrip_is_read_only(class_name):
    shell_control = _wait_for_control(
        uiautomation.PaneControl(ClassName=class_name, searchDepth=1),
        class_name,
    )
    xpath = get_control_xpath(shell_control)
    assert xpath
    replayed = find_control_by_xpath(xpath, use_cache=False)
    assert replayed is not None
    assert uiautomation.ControlsAreSame(shell_control, replayed)
