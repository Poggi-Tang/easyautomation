from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import uuid
from pathlib import Path

import pytest

pytest.importorskip("mcp")

import anyio
import uiautomation
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

RUN_WINDOWS_INTEGRATION = os.environ.get("EASY_UIAUTO_RUN_WINDOWS_TESTS") == "1"
RUN_SIMUNPS_INTEGRATION = os.environ.get("EASY_UIAUTO_RUN_SIMUNPS_TESTS") == "1"
FIXTURES = Path(__file__).parent / "fixtures"


def _launch_fixture(script_name: str, title: str, **environment: str) -> subprocess.Popen[str]:
    process = subprocess.Popen(
        [sys.executable, str(FIXTURES / script_name), title],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env={
            **os.environ,
            "PYTHONUNBUFFERED": "1",
            "QT_QPA_PLATFORM": "windows",
            **environment,
        },
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
    if ready != "READY":
        stderr = process.stderr.read() if process.stderr else ""
        raise AssertionError(f"fixture failed before ready: {stderr}")
    return process


def _stop_fixture(process: subprocess.Popen[str]) -> None:
    process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=3)


def _location(
    window_name: str,
    name: str = "",
    control_type: str = "",
    *,
    class_name: str = "",
    automation_id: str = "",
) -> dict[str, object]:
    return {
        "WindowName": window_name,
        "Name": name,
        "ClassName": class_name,
        "ControlType": control_type,
        "foundIndex": 1,
        "AutomationId": automation_id,
        "Xpath": [],
        "Img": "",
        "PARAMETERS": {},
    }


async def _tool(session: ClientSession, name: str, arguments: dict[str, object]):
    result = await session.call_tool(name, arguments)
    assert result.structuredContent is not None
    return result.structuredContent


async def _run_with_mcp(scenario, environment: dict[str, str] | None = None):
    server = StdioServerParameters(
        command=sys.executable,
        args=["-m", "easy_uiauto.mcp"],
        env=environment,
    )
    async with stdio_client(server) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            return await scenario(session)


@pytest.mark.windows_integration
@pytest.mark.skipif(
    not RUN_WINDOWS_INTEGRATION,
    reason="set EASY_UIAUTO_RUN_WINDOWS_TESTS=1 to run desktop integration tests",
)
def test_mcp_stdio_win32_actions_tree_cache_and_highlight_roundtrip():
    title = f"Easy UIAuto MCP Win32 {uuid.uuid4().hex[:8]}"
    process = _launch_fixture("win32_fixture.py", title)

    async def scenario(session: ClientSession):
        window = await _tool(session, "find_control", {"location": {"WindowName": title}})
        assert window["ok"] is True
        window_ref = window["data"]["ref"]

        children = await _tool(
            session,
            "list_children",
            {"reference": window_ref, "offset": 0, "limit": 20},
        )
        tree = await _tool(
            session,
            "get_control_tree",
            {"reference": window_ref, "max_depth": 1, "max_nodes": 12},
        )
        assert children["ok"] is True
        assert children["data"]["total"] >= 5
        assert tree["ok"] is True
        assert tree["data"]["node_count"] > 1

        edit_location = _location(title, control_type="EditControl")
        edit = await _tool(session, "find_control", {"location": edit_location})
        assert edit["ok"] is True
        value = f"mcp-win32-{uuid.uuid4().hex[:6]}"
        edited = await _tool(
            session,
            "perform_action",
            {
                "action": "输入文本",
                "reference": edit["data"]["ref"],
                "parameters": {"输入文本": value},
            },
        )
        assert edited["ok"] is True
        assert edited["data"]["verified"] is True

        button_location = _location(title, "Apply", "ButtonControl")
        button = await _tool(session, "find_control", {"location": button_location})
        assert button["ok"] is True
        expected_status = f"applied:{value}"
        clicked = await _tool(
            session,
            "perform_action",
            {
                "action": "点击",
                "reference": button["data"]["ref"],
                "observe_location": _location(title, expected_status, "TextControl"),
                "expected": {"path": "name", "equals": expected_status},
            },
        )
        assert clicked["ok"] is True
        assert clicked["data"]["verified"] is True

        highlighted = await _tool(
            session,
            "start_highlight",
            {"reference": button["data"]["ref"], "color": "#00FF00", "alpha": 0.7},
        )
        assert highlighted["ok"] is True, highlighted
        highlight_id = highlighted["data"]["session_id"]
        updated = await _tool(
            session,
            "update_highlight",
            {"session_id": highlight_id, "line_width": 3},
        )
        stopped = await _tool(
            session,
            "stop_highlight",
            {"session_id": highlight_id},
        )
        assert updated["ok"] is True
        assert stopped["ok"] is True
        assert stopped["data"]["running"] is False
        assert stopped["data"]["forced_termination"] is False

        stats = await _tool(session, "cache_stats", {})
        cleared = await _tool(session, "clear_caches", {})
        assert stats["ok"] is True
        assert stats["data"]["mcp_references"] >= 3
        assert cleared["ok"] is True
        assert cleared["data"]["mcp_references_removed"] >= 3

    try:
        anyio.run(_run_with_mcp, scenario)
    finally:
        _stop_fixture(process)


@pytest.mark.windows_integration
@pytest.mark.skipif(
    not RUN_WINDOWS_INTEGRATION,
    reason="set EASY_UIAUTO_RUN_WINDOWS_TESTS=1 to run desktop integration tests",
)
def test_mcp_stdio_qt_verified_actions_and_provider_false_success():
    title = f"Easy UIAuto MCP Qt {uuid.uuid4().hex[:8]}"
    process = _launch_fixture("qt_fixture.py", title)

    async def scenario(session: ClientSession):
        edit_location = _location(title, "Input", "EditControl")
        edit = await _tool(session, "find_control", {"location": edit_location})
        assert edit["ok"] is True
        value = f"mcp-qt-{uuid.uuid4().hex[:6]}"
        edited = await _tool(
            session,
            "perform_action",
            {
                "action": "输入文本",
                "reference": edit["data"]["ref"],
                "parameters": {"输入文本": value},
            },
        )
        assert edited["ok"] is True
        assert edited["data"]["verified"] is True

        tab_location = _location(title, "Second", "TabItemControl")
        tab = await _tool(session, "find_control", {"location": tab_location})
        assert tab["ok"] is True
        selected = await _tool(
            session,
            "perform_action",
            {
                "action": "选择",
                "reference": tab["data"]["ref"],
                "observe_location": _location(title, "tab:Second", "TextControl"),
                "expected": {"path": "name", "equals": "tab:Second"},
            },
        )
        assert selected["ok"] is True
        assert selected["data"]["verified"] is True

        combo_location = _location(title, control_type="ComboBoxControl")
        combo = await _tool(session, "find_control", {"location": combo_location})
        assert combo["ok"] is True
        false_success = await _tool(
            session,
            "perform_action",
            {
                "action": "选择",
                "reference": combo["data"]["ref"],
                "parameters": {"选择项": "Beta"},
            },
        )
        assert false_success["ok"] is False
        assert false_success["data"]["verified"] is False

        root_location = _location(title, "Root", "TreeItemControl")
        root = await _tool(session, "find_control", {"location": root_location})
        assert root["ok"] is True
        false_expand = await _tool(
            session,
            "perform_action",
            {
                "action": "展开折叠",
                "reference": root["data"]["ref"],
                "parameters": {"展开": True},
            },
        )
        assert false_expand["ok"] is False
        assert false_expand["data"]["verified"] is False

    try:
        anyio.run(_run_with_mcp, scenario)
    finally:
        _stop_fixture(process)


@pytest.mark.windows_integration
@pytest.mark.skipif(
    not RUN_WINDOWS_INTEGRATION,
    reason="set EASY_UIAUTO_RUN_WINDOWS_TESTS=1 to run desktop integration tests",
)
def test_mcp_stdio_recording_session_lifecycle_without_input_injection():
    async def scenario(session: ClientSession):
        started = await _tool(session, "start_recording", {})
        assert started["ok"] is True, started
        session_id = started["data"]["session_id"]
        status = await _tool(
            session,
            "recording_status",
            {"session_id": session_id, "include_actions": False},
        )
        stopped = await _tool(
            session,
            "stop_recording",
            {"session_id": session_id},
        )
        assert status["ok"] is True
        assert status["data"]["running"] is True
        assert stopped["ok"] is True, stopped
        assert stopped["data"]["running"] is False
        assert stopped["data"]["cleanup_errors"] == []

    environment = {**os.environ, "EASY_UIAUTO_MCP_ALLOW_RECORDING": "1"}
    anyio.run(_run_with_mcp, scenario, environment)


@pytest.mark.windows_integration
@pytest.mark.skipif(
    not RUN_SIMUNPS_INTEGRATION,
    reason="set EASY_UIAUTO_RUN_SIMUNPS_TESTS=1 for reversible SimuNPS smoke tests",
)
def test_mcp_stdio_simunps_search_and_filter_are_reversible():
    window = uiautomation.WindowControl(
        Name="SimuNPS",
        ClassName="EmtMainUI",
        AutomationId="SimuNPS",
        searchDepth=1,
    )
    if not window.Exists(2):
        pytest.skip("SimuNPS is not running")

    async def scenario(session: ClientSession):
        search_location = _location(
            "SimuNPS",
            "model_serach",
            "EditControl",
            class_name="SearchLineEdit",
        )
        filter_location = _location(
            "SimuNPS",
            "消息",
            "CheckBoxControl",
            class_name="QPushButton",
        )
        search = await _tool(session, "find_control", {"location": search_location})
        message_filter = await _tool(session, "find_control", {"location": filter_location})
        assert search["ok"] is True
        assert message_filter["ok"] is True
        original_value = search["data"]["patterns"]["value"]["value"]
        original_toggle = int(message_filter["data"]["patterns"]["toggle"]["state"])
        probe = f"easy-uiauto-mcp-{uuid.uuid4().hex[:8]}"
        try:
            changed_search = await _tool(
                session,
                "perform_action",
                {
                    "action": "输入文本",
                    "reference": search["data"]["ref"],
                    "parameters": {"输入文本": probe},
                },
            )
            changed_filter = await _tool(
                session,
                "perform_action",
                {
                    "action": "切换状态",
                    "reference": message_filter["data"]["ref"],
                    "parameters": {"选中": original_toggle == 0},
                },
            )
            assert changed_search["ok"] is True
            assert changed_search["data"]["verified"] is True
            assert changed_filter["ok"] is True
            assert changed_filter["data"]["verified"] is True
        finally:
            restored_search = await _tool(
                session,
                "perform_action",
                {
                    "action": "输入文本",
                    "reference": search["data"]["ref"],
                    "parameters": {"输入文本": original_value},
                },
            )
            restored_filter = await _tool(
                session,
                "perform_action",
                {
                    "action": "切换状态",
                    "reference": message_filter["data"]["ref"],
                    "parameters": {"选中": bool(original_toggle)},
                },
            )
            assert restored_search["ok"] is True
            assert restored_filter["ok"] is True

        final_search = await _tool(
            session,
            "inspect_control",
            {"reference": search["data"]["ref"]},
        )
        final_filter = await _tool(
            session,
            "inspect_control",
            {"reference": message_filter["data"]["ref"]},
        )
        assert final_search["data"]["patterns"]["value"]["value"] == original_value
        assert int(final_filter["data"]["patterns"]["toggle"]["state"]) == original_toggle

    anyio.run(_run_with_mcp, scenario)
