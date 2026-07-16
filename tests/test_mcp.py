from __future__ import annotations

import sys
import threading

import pytest

pytest.importorskip("mcp")

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

from conftest import FakeControl
from easy_uiauto.mcp import backend, executor
from easy_uiauto.mcp.backend import UIAutomationBackend
from easy_uiauto.mcp.policy import MCPPolicy


def test_executor_serializes_calls_on_dedicated_thread(monkeypatch):
    initialized = []
    monkeypatch.setattr(
        executor,
        "ensure_uiautomation_thread",
        lambda: initialized.append(threading.get_ident()),
    )

    class FakeBackend:
        def thread(self):
            return threading.get_ident()

    with executor.UIAutomationExecutor(FakeBackend) as worker:
        first = worker.call("thread")
        second = worker.call("thread")

    assert initialized == [first]
    assert first == second
    assert first != threading.get_ident()


def test_backend_returns_json_snapshot_and_refreshable_reference(monkeypatch):
    control = FakeControl(name="Apply", control_type="ButtonControl")
    xpath = [
        {"ControlType": "WindowControl", "Name": "Fixture"},
        {"ControlType": "ButtonControl", "Name": "Apply"},
    ]
    monkeypatch.setattr(backend, "find_control", lambda _location: control)
    monkeypatch.setattr(backend, "get_control_xpath", lambda _control: xpath)
    service = UIAutomationBackend(policy=MCPPolicy(), reference_ttl=10)

    found = service.find({"WindowName": "Fixture", "Name": "Apply"})
    inspected = service.inspect(found.data["ref"])

    assert found.ok
    assert found.data["name"] == "Apply"
    assert found.data["xpath"] == xpath
    assert inspected.ok
    assert inspected.data["name"] == "Apply"
    assert inspected.data["ref"] != found.data["ref"]


def test_backend_rejects_stale_reference(monkeypatch):
    control = FakeControl()
    monkeypatch.setattr(backend, "get_control_xpath", lambda _control: [])
    service = UIAutomationBackend(policy=MCPPolicy())
    reference = service._store_reference(control, [])
    control.Exists = lambda *_args, **_kwargs: False

    result = service.inspect(reference)

    assert not result.ok
    assert "stale" in result.message


def test_mcp_policy_blocks_high_risk_actions(monkeypatch):
    calls = []
    monkeypatch.setattr(backend, "run_action", lambda action: calls.append(action) or "ok")
    service = UIAutomationBackend(policy=MCPPolicy(allow_high_risk=False))

    result = service.perform_action(
        "组合键",
        location={"WindowName": "", "PARAMETERS": {}},
        parameters={"组合键": "alt+f4"},
    )

    assert not result.ok
    assert calls == []


def test_backend_action_result_uses_structured_message_state(monkeypatch):
    monkeypatch.setattr(backend, "run_action", lambda _action: "provider failed")
    monkeypatch.setattr(backend, "get_message_type", lambda: 0)
    service = UIAutomationBackend(policy=MCPPolicy())

    result = service.perform_action(
        "点击",
        location={
            "WindowName": "Fixture",
            "Name": "Apply",
            "ClassName": "Button",
            "ControlType": "ButtonControl",
            "foundIndex": 1,
            "AutomationId": "",
            "Xpath": [],
            "Img": "",
            "PARAMETERS": {},
        },
    )

    assert not result.ok
    assert result.message == "provider failed"


def test_mcp_stdio_server_lists_expected_tools():
    async def scenario():
        server = StdioServerParameters(
            command=sys.executable,
            args=["-m", "easy_uiauto.mcp"],
        )
        async with stdio_client(server) as (read_stream, write_stream):
            async with ClientSession(read_stream, write_stream) as session:
                await session.initialize()
                tools = await session.list_tools()
                windows = await session.call_tool("list_windows", {"limit": 2})
                return {tool.name for tool in tools.tools}, windows.structuredContent

    names, windows = anyio.run(scenario)

    assert names == {"list_windows", "find_control", "inspect_control", "perform_action"}
    assert windows["ok"] is True
    assert len(windows["data"]) <= 2
