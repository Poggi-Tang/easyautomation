"""TCP is an adapter over shared strict lookup and bounded tree inspection."""

import json
import socket
import threading
from types import SimpleNamespace
from unittest.mock import Mock, call

import pytest

from easy_uiauto.mcp import client, service
from easy_uiauto.mcp.results import ActionError


@pytest.fixture
def adapter():
    instance = service.UIAutomationService(port=0)
    yield instance
    instance._sel.close()


def test_tree_handler_delegates_limits_and_partial_failures(adapter, monkeypatch):
    partial = {"ok": False, "code": "timeout", "data": {"tree": {"Name": "Partial"},
               "inspection": {"worker_stopped": True}}}
    inspect = Mock(return_value=partial)
    monkeypatch.setattr(service.control_tree, "inspect_window", inspect)
    value = adapter._handle_get_control_tree({"window_name": "Example", "max_depth": 3,
        "max_nodes": 17, "timeout_seconds": 0.2}, "request")
    assert value is partial
    inspect.assert_called_once_with("Example", 3, max_nodes=17, timeout_seconds=0.2)


def test_compile_compatibility_uses_bounded_read(adapter, monkeypatch):
    inspect = Mock(return_value={"Name": "Example"})
    monkeypatch.setattr(service.control_tree, "inspect_window", inspect)
    assert adapter._handle_compile_window({"window_name": "Example"}, "r") == "Compiled 'Example' (depth=8)"
    inspect.assert_called_once_with("Example", 8, max_nodes=500, timeout_seconds=15.0)


def test_client_restores_timeout_even_when_transport_fails(monkeypatch):
    instance = client.UIAutomationClient(timeout=30)
    instance._sock = Mock(gettimeout=Mock(return_value=30.0))
    transport = Mock(side_effect=socket.timeout("transport"))
    monkeypatch.setattr(instance, "_call", transport)
    with pytest.raises(socket.timeout):
        instance.get_control_tree("Example", 2, max_nodes=7, timeout_seconds=60)
    assert instance._sock.settimeout.call_args_list == [call(65.0), call(30.0)]
    transport.assert_called_once_with("get_control_tree", window_name="Example", max_depth=2,
                                 max_nodes=7, timeout_seconds=60)


def test_smart_click_does_not_compile_or_click_cached_coordinates(adapter, monkeypatch):
    first = SimpleNamespace(Name="Save", BoundingRectangle=SimpleNamespace(left=10, top=10, right=30, bottom=30))
    second = SimpleNamespace(Name="Save", BoundingRectangle=SimpleNamespace(left=110, top=10, right=130, bottom=30))
    lookup = Mock(side_effect=[first, second])
    execute, activate = Mock(), Mock()
    monkeypatch.setattr(service, "_find_control_utils", lookup)
    monkeypatch.setattr(service.Controller, "_activate_target", activate)
    monkeypatch.setattr(service.Controller, "execute_resolved", execute)
    for _ in range(2):
        adapter._handle_smart_click({"window_name": "Example", "name": "Save"}, "r")
    assert lookup.call_count == 2
    assert execute.call_args_list == [(("click",), {"control": first}), (("click",), {"control": second})]
    assert activate.call_count == 2


def test_smart_click_ambiguity_performs_no_mutation(adapter, monkeypatch):
    execute, activate = Mock(), Mock()
    monkeypatch.setattr(service, "_find_control_utils", Mock(side_effect=ActionError("ambiguous_match", "two")))
    monkeypatch.setattr(service.Controller, "_activate_target", activate)
    monkeypatch.setattr(service.Controller, "execute_resolved", execute)
    with pytest.raises(ActionError, match="ambiguous_match"):
        adapter._handle_smart_click({"window_name": "Example", "name": "Save"}, "r")
    execute.assert_not_called()
    activate.assert_not_called()


@pytest.mark.parametrize("button,expected", [("left", "double-click"),
    ("right", "double-right-click"), ("middle", "double-middle-click")])
def test_double_click_buttons_retain_tcp_compatibility(adapter, monkeypatch, button, expected):
    execute = Mock()
    monkeypatch.setattr(service.Controller, "execute_resolved", execute)
    adapter._handle_click_at_position({"x": 12, "y": 34, "button": button, "click_type": "double"}, "r")
    execute.assert_called_once_with(expected, x=12, y=34)


def test_service_starts_without_uia_warmup_and_responds_after_tree_timeout(monkeypatch, tmp_path):
    instance = service.UIAutomationService(port=0)
    ready = threading.Event()
    monkeypatch.setattr(service, "PID_FILE", str(tmp_path / "service.pid"))
    monkeypatch.setattr(service, "_log", lambda message: ready.set() if "Service ready" in message else None)
    monkeypatch.setattr(service.uiautomation, "GetRootControl", Mock(side_effect=AssertionError("startup queried UIA")))
    monkeypatch.setattr(service.control_tree, "inspect_window", Mock(return_value={
        "ok": False, "code": "timeout", "data": {"tree": {"Name": "Partial"}}}))
    monkeypatch.setattr(service.pyautogui, "getAllWindows", lambda: [])
    thread = threading.Thread(target=instance.start, daemon=True)
    thread.start()
    try:
        assert ready.wait(5)
        port = instance._server_sock.getsockname()[1]
        with client.UIAutomationClient(port=port, timeout=3) as connection:
            result = connection.get_control_tree("Example")
            assert result["code"] == "timeout" and result["data"]["tree"]["Name"] == "Partial"
            assert connection.list_windows() == []
            assert connection._sock.gettimeout() == 3
            connection._call("shutdown")
    finally:
        instance._running = False
        thread.join(timeout=5)
    assert not thread.is_alive()
