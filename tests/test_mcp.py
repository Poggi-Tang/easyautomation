from __future__ import annotations

import sys
import threading
import time

import pytest

pytest.importorskip("mcp")

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

from conftest import FakeControl
from easy_uiauto.mcp import backend, executor, sessions
from easy_uiauto.mcp.backend import UIAutomationBackend
from easy_uiauto.mcp.policy import MCPPolicy
from easy_uiauto.mcp.sessions import HighlightSession


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


def test_executor_does_not_return_while_timed_out_mutation_is_running(monkeypatch):
    monkeypatch.setattr(executor, "ensure_uiautomation_thread", lambda: None)
    mutations = []

    class FakeBackend:
        def mutate(self):
            time.sleep(0.05)
            mutations.append("done")
            return "complete"

    with executor.UIAutomationExecutor(FakeBackend) as worker:
        started = time.monotonic()
        result = worker.call("mutate", timeout=0.01)
        elapsed = time.monotonic() - started

    assert result == "complete"
    assert elapsed >= 0.04
    assert mutations == ["done"]


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
    assert inspected.data["ref"] == found.data["ref"]


def test_backend_reference_survives_mutable_name_change(monkeypatch):
    control = FakeControl(name="before")
    monkeypatch.setattr(backend, "get_control_xpath", lambda _control: [])
    service = UIAutomationBackend(policy=MCPPolicy())
    reference = service._store_reference(control, [])
    control.Name = "after"

    result = service.inspect(reference)

    assert result.ok
    assert result.data["name"] == "after"


def test_backend_rejects_stale_reference(monkeypatch):
    control = FakeControl()
    monkeypatch.setattr(backend, "get_control_xpath", lambda _control: [])
    service = UIAutomationBackend(policy=MCPPolicy())
    reference = service._store_reference(control, [])
    control.Exists = lambda *_args, **_kwargs: False

    result = service.inspect(reference)

    assert not result.ok
    assert "stale" in result.message


def test_backend_lists_children_and_serializes_bounded_tree(monkeypatch):
    root = FakeControl(name="root", control_type="WindowControl")
    FakeControl(name="first", parent=root)
    FakeControl(name="second", parent=root)
    monkeypatch.setattr(
        backend,
        "get_control_xpath",
        lambda control: [{"ControlType": control.ControlTypeName, "Name": control.Name}],
    )
    service = UIAutomationBackend(policy=MCPPolicy())
    reference = service._store_reference(root, [])

    children = service.list_children(reference, offset=1, limit=1)
    tree = service.get_control_tree(reference, max_depth=1, max_nodes=2)

    assert children.ok
    assert [item["name"] for item in children.data["items"]] == ["second"]
    assert children.data["has_more"] is False
    assert tree.ok
    assert tree.data["node_count"] == 2
    assert tree.data["truncated"] is True
    assert tree.data["root"]["name"] == "root"


def test_lightweight_reference_resolves_xpath_lazily_for_action(monkeypatch):
    control = FakeControl(name="Apply")
    xpath = [
        {"ControlType": "WindowControl", "Name": "Fixture"},
        {"ControlType": "ButtonControl", "Name": "Apply"},
    ]
    monkeypatch.setattr(backend, "get_control_xpath", lambda _control: xpath)
    monkeypatch.setattr(backend, "find_control", lambda _location: None)
    monkeypatch.setattr(backend, "run_action", lambda _action: "clicked")
    monkeypatch.setattr(backend, "get_message_type", lambda: 2)
    service = UIAutomationBackend(policy=MCPPolicy())
    reference = service._store_reference(control, [])

    result = service.perform_action("点击", reference=reference)

    assert result.ok
    assert result.status == "warning"
    assert service._resolve_reference(reference).xpath == xpath


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


def test_mcp_policy_requires_confirmation_for_high_impact_target(monkeypatch):
    calls = []
    monkeypatch.setattr(backend, "find_control", lambda _location: None)
    monkeypatch.setattr(backend, "run_action", lambda action: calls.append(action) or "ok")
    service = UIAutomationBackend(policy=MCPPolicy())
    location = {
        "WindowName": "SimuNPS",
        "Name": "Run Simulation",
        "ControlType": "ButtonControl",
        "PARAMETERS": {},
    }

    blocked = service.perform_action("点击", location=location)
    dry_run = service.perform_action("点击", location=location, dry_run=True)

    assert not blocked.ok
    assert dry_run.ok
    assert dry_run.data["would_require_high_impact_confirmation"] is True
    assert calls == []


@pytest.mark.parametrize("name", ["Runtime Status", "Clearance"])
def test_mcp_policy_does_not_match_high_impact_terms_inside_unrelated_words(name):
    policy = MCPPolicy()

    result = policy.validate_action("点击", {"Name": name})

    assert result is None


def test_mcp_policy_detects_camel_case_high_impact_automation_id():
    policy = MCPPolicy()

    result = policy.validate_action("点击", {"AutomationId": "runSimulationButton"})

    assert result is not None


def test_recording_is_disabled_by_default():
    service = UIAutomationBackend(policy=MCPPolicy())

    result = service.start_recording()

    assert not result.ok
    assert "EASY_UIAUTO_MCP_ALLOW_RECORDING" in result.message


def test_recording_start_failure_stops_and_joins_recorder(monkeypatch):
    stopped = []
    joined = []

    class FailingRecorder:
        def __init__(self):
            self.started = threading.Event()
            self.run_error = OSError("listener failed")
            self.cleanup_errors = []

        def start(self):
            self.started.set()

        def stop(self):
            stopped.append(True)

        def join(self, timeout):
            joined.append(timeout)

        def is_alive(self):
            return False

    monkeypatch.setattr(sessions, "RecordThread", FailingRecorder)

    with pytest.raises(RuntimeError, match="recording listener failed to start"):
        sessions.RecordingSession.start()

    assert stopped == [True]
    assert joined == [5]


def test_highlight_session_starts_updates_and_stops_without_leaking_thread():
    rect = {"left": 20, "top": 20, "right": 120, "bottom": 80, "width": 100, "height": 60}
    session = HighlightSession(rect)

    session.start()
    assert session.status()["running"] is True
    session.update_rect({**rect, "left": 30, "right": 130})
    session.update_style(color="#00FF00", alpha=0.5)
    status = session.stop()

    assert status["running"] is False
    assert status["stopped"] is True


def test_backend_verifies_value_postcondition(monkeypatch):
    control = FakeControl(control_type="EditControl")
    snapshots = iter(
        [
            {"patterns": {"value": {"value": "before"}}},
            {"patterns": {"value": {"value": "after"}}},
        ]
    )
    monkeypatch.setattr(backend, "find_control", lambda _location: control)
    monkeypatch.setattr(backend, "run_action", lambda _action: "input complete")
    monkeypatch.setattr(backend, "get_message_type", lambda: 2)
    service = UIAutomationBackend(policy=MCPPolicy())
    monkeypatch.setattr(service, "_snapshot", lambda *_args, **_kwargs: next(snapshots))

    result = service.perform_action(
        "输入文本",
        location={"WindowName": "Fixture", "PARAMETERS": {}},
        parameters={"输入文本": "after"},
    )

    assert result.ok
    assert result.data["verified"] is True
    assert result.data["after"]["patterns"]["value"]["value"] == "after"


@pytest.mark.parametrize(
    ("action", "parameters", "pattern_path", "state"),
    [
        ("切换状态", {"选中": "false"}, "toggle", 0),
        ("展开折叠", {"展开": "false"}, "expand_collapse", 0),
    ],
)
def test_backend_normalizes_false_string_for_postcondition(
    monkeypatch, action, parameters, pattern_path, state
):
    control = FakeControl()
    snapshots = iter(
        [
            {"patterns": {pattern_path: {"state": 1}}},
            {"patterns": {pattern_path: {"state": state}}},
        ]
    )
    monkeypatch.setattr(backend, "find_control", lambda _location: control)
    monkeypatch.setattr(backend, "run_action", lambda _action: "complete")
    monkeypatch.setattr(backend, "get_message_type", lambda: 2)
    service = UIAutomationBackend(policy=MCPPolicy())
    monkeypatch.setattr(service, "_snapshot", lambda *_args, **_kwargs: next(snapshots))

    result = service.perform_action(
        action,
        location={"WindowName": "Fixture", "PARAMETERS": {}},
        parameters=parameters,
    )

    assert result.ok
    assert result.data["verified"] is True


def test_snapshot_does_not_attach_stale_reference_to_different_control(monkeypatch):
    first = FakeControl(name="first")
    second = FakeControl(name="second")
    monkeypatch.setattr(backend, "get_control_xpath", lambda _control: [])
    service = UIAutomationBackend(policy=MCPPolicy())
    reference = service._store_reference(first, [])

    snapshot = service._snapshot(second, reference=reference)

    assert snapshot["ref"] != reference
    assert service._resolve_reference(reference).control is first


def test_backend_action_result_uses_structured_message_state(monkeypatch):
    monkeypatch.setattr(backend, "find_control", lambda _location: None)
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


def test_backend_does_not_execute_when_pre_action_lookup_raises(monkeypatch):
    calls = []
    monkeypatch.setattr(
        backend,
        "find_control",
        lambda _location: (_ for _ in ()).throw(OSError("COM unavailable")),
    )
    monkeypatch.setattr(backend, "run_action", lambda action: calls.append(action) or "ok")
    service = UIAutomationBackend(policy=MCPPolicy())

    result = service.perform_action(
        "点击",
        location={"WindowName": "Fixture", "PARAMETERS": {}},
    )

    assert not result.ok
    assert "pre-action target lookup failed" in result.message
    assert calls == []


def test_backend_preserves_action_result_when_post_action_lookup_raises(monkeypatch):
    control = FakeControl()
    lookups = iter([control, OSError("transient COM failure")])

    def find(_location):
        value = next(lookups)
        if isinstance(value, BaseException):
            raise value
        return value

    monkeypatch.setattr(backend, "find_control", find)
    monkeypatch.setattr(backend, "run_action", lambda _action: "clicked")
    monkeypatch.setattr(backend, "get_message_type", lambda: 2)
    service = UIAutomationBackend(policy=MCPPolicy())
    monkeypatch.setattr(
        service,
        "_snapshot",
        lambda *_args, **_kwargs: {"patterns": {"invoke": {"supported": True}}},
    )

    result = service.perform_action(
        "点击",
        location={"WindowName": "Fixture", "PARAMETERS": {}},
    )

    assert result.ok
    assert result.status == "warning"
    assert result.message == "clicked"
    assert result.data["verified"] is None
    assert result.warnings == ["post-action target lookup failed: transient COM failure"]


def test_backend_close_attempts_every_session_even_when_one_stop_fails():
    stopped = []

    class Session:
        def __init__(self, name, fails=False):
            self.name = name
            self.fails = fails

        def stop(self):
            stopped.append(self.name)
            if self.fails:
                raise RuntimeError("cleanup failed")

    service = UIAutomationBackend(policy=MCPPolicy())
    service._recordings = {
        "first": Session("recording-first", fails=True),
        "second": Session("recording-second"),
    }
    service._highlights = {
        "first": Session("highlight-first", fails=True),
        "second": Session("highlight-second"),
    }

    service.close()

    assert stopped == [
        "recording-first",
        "recording-second",
        "highlight-first",
        "highlight-second",
    ]
    assert service._recordings == {}
    assert service._highlights == {}


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

    assert names == {
        "list_windows",
        "find_control",
        "inspect_control",
        "list_children",
        "get_control_tree",
        "cache_stats",
        "clear_caches",
        "invalidate_control",
        "start_recording",
        "recording_status",
        "stop_recording",
        "start_highlight",
        "update_highlight",
        "stop_highlight",
        "perform_action",
    }
    assert windows["ok"] is True
    assert len(windows["data"]) <= 2
