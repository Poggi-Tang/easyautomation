"""Inspection bounds, native selection and termination of blocked UIA workers."""

from __future__ import annotations

import ast
import importlib
import importlib.util
import json
import subprocess
import sys
import time
from contextlib import nullcontext
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

_ROOT = Path(__file__).resolve().parents[1] / "src/easy_uiauto/mcp"
_NAME = "_easy_control_tree_tests"
_SPEC = importlib.util.spec_from_file_location(
    _NAME, _ROOT / "__init__.py", submodule_search_locations=[str(_ROOT)]
)
_PACKAGE = importlib.util.module_from_spec(_SPEC)
sys.modules[_NAME] = _PACKAGE
_SPEC.loader.exec_module(_PACKAGE)
inspection = importlib.import_module(f"{_NAME}.control_tree")
native = importlib.import_module(f"{_NAME}.native_windows")


class Control:
    def __init__(self, name="Window", children=()):
        self.Name = name
        self.ControlTypeName = "PaneControl"
        self.ClassName, self.AutomationId = "Example", name
        self.BoundingRectangle = SimpleNamespace(left=-20, top=10, right=80, bottom=110)
        self.children = list(children)
        self.first = Mock(return_value=self.children[0] if self.children else None)
        self.next = Mock(return_value=None)
        for left, right in zip(self.children, self.children[1:], strict=False):
            left.next.return_value = right

    def GetFirstChildControl(self):
        return self.first()

    def GetNextSiblingControl(self):
        return self.next()

    def GetChildren(self):
        pytest.fail("Eager GetChildren enumeration is forbidden")


def collector(**kwargs):
    return inspection.TreeCollector(
        max_depth=kwargs.pop("max_depth", 5),
        max_nodes=kwargs.pop("max_nodes", 500),
        deadline=kwargs.pop("deadline", time.monotonic() + 10),
        checkpoint=lambda _state: None,
        **kwargs,
    )


def test_depth_zero_does_not_query_children_and_preserves_properties():
    root = Control()
    tree = collector(max_depth=0)
    tree.visit(root)
    assert tree.tree["Name"] == "Window"
    assert tree.tree["bounds"]["left"] == -20
    root.first.assert_not_called()
    assert tree.nodes == 1 and "max_depth" in tree.limits


def test_global_node_limit_stops_before_another_provider_query():
    children = [Control(str(number)) for number in range(25)]
    root = Control(children=children)
    tree = collector(max_nodes=3)
    tree.visit(root)
    assert tree.nodes == 3
    assert [item["Name"] for item in tree.tree["children"]] == ["0", "1"]
    assert "max_nodes" in tree.limits
    children[1].first.assert_not_called()
    children[1].next.assert_not_called()


def test_sibling_limit_uses_lazy_iteration():
    root = Control(children=[Control(str(number)) for number in range(50)])
    tree = collector(max_depth=1)
    tree.visit(root)
    assert len(tree.tree["children"]) == 30
    assert tree.nodes == 31 and "children_per_node" in tree.limits
    root.children[29].next.assert_not_called()


def test_provider_error_is_visible_and_does_not_discard_the_root():
    root = Control()
    root.first.side_effect = RuntimeError("Provider unavailable")
    tree = collector()
    tree.visit(root)
    assert tree.tree["Name"] == "Window"
    assert "Provider unavailable" in tree.state()["inspection"]["errors"][0]["message"]
    assert tree.state()["inspection"]["complete"] is False


def test_expired_budget_prevents_the_next_uia_query():
    root = Control()
    tree = collector(deadline=1, clock=lambda: 2)
    with pytest.raises(TimeoutError):
        tree.visit(root)
    root.first.assert_not_called()


def entry(title, *, foreground=False, handle=8):
    return dict(title=title, visible=True, foreground=foreground, handle=handle)


def test_native_selection_prefers_exact_title_without_uia_root(monkeypatch):
    enumerate_windows = Mock(return_value=[entry("Editor - other"), entry("Editor", handle=9)])
    monkeypatch.setattr(native, "enumerate_windows", enumerate_windows)
    assert inspection._select_window("Editor")["handle"] == 9
    enumerate_windows.assert_called_once_with(visible_only=True)


def test_native_selection_rejects_ambiguous_visible_windows(monkeypatch):
    monkeypatch.setattr(
        native, "enumerate_windows", lambda **_kw: [entry("Editor"), entry("Editor")]
    )
    with pytest.raises(inspection.ActionError, match="ambiguous_match"):
        inspection._select_window("Editor")


def test_native_selection_uses_foreground_duplicate_and_reports_missing(monkeypatch):
    monkeypatch.setattr(
        native,
        "enumerate_windows",
        lambda **_kw: [
            entry("Editor", handle=1),
            entry("Editor", foreground=True, handle=2),
        ],
    )
    assert inspection._select_window("Editor")["handle"] == 2
    with pytest.raises(inspection.ActionError, match="window_invalid"):
        inspection._select_window("Missing")


def test_worker_binds_only_the_native_handle_and_never_searches_desktop(monkeypatch, tmp_path):
    (tmp_path / "request.json").write_text(
        json.dumps(
            dict(
                window_name="Editor",
                max_depth=2,
                max_nodes=20,
                deadline=time.monotonic() + 10,
            )
        ),
        encoding="utf-8",
    )
    root = Control("Editor")
    bind = Mock(return_value=root)
    forbidden = Mock(side_effect=AssertionError("Desktop root search is forbidden"))
    monkeypatch.setattr(inspection, "_select_window", lambda _title: entry("Editor", handle=42))
    monkeypatch.setitem(
        sys.modules,
        "uiautomation",
        SimpleNamespace(
            UIAutomationInitializerInThread=lambda: nullcontext(),
            ControlFromHandle=bind,
            GetRootControl=forbidden,
            WindowControl=forbidden,
        ),
    )
    inspection._worker(tmp_path)
    state = inspection._read_state(tmp_path)
    assert state["done"] and state["outcome"]["Name"] == "Editor"
    bind.assert_called_once_with(42)
    forbidden.assert_not_called()


@pytest.mark.parametrize(
    "options",
    [
        {"window_name": " "},
        {"max_depth": -1},
        {"max_depth": True},
        {"max_nodes": 0},
        {"max_nodes": 5001},
        {"timeout_seconds": 0},
        {"timeout_seconds": float("nan")},
        {"timeout_seconds": 61},
    ],
)
def test_invalid_options_do_not_launch_a_worker(monkeypatch, options):
    launch = Mock(side_effect=AssertionError("Worker must not start"))
    monkeypatch.setattr(inspection.subprocess, "Popen", launch)
    with pytest.raises(ValueError):
        inspection.inspect_window(**{"window_name": "Editor", **options})
    launch.assert_not_called()


def test_blocked_process_is_killed_and_partial_checkpoint_is_returned(monkeypatch):
    launched = []
    worker_folders = []
    original = subprocess.Popen

    def launch(*args, **kwargs):
        process = original(*args, **kwargs)
        launched.append(process)
        return process

    def command(folder):
        worker_folders.append(folder)
        # A provider-equivalent permanent block after recording partial evidence.
        inspection._write_state(
            folder,
            dict(
                tree={"Name": "Partial"},
                inspection=dict(
                    stage="node-2.Name",
                    nodes=2,
                ),
            ),
        )
        return [sys.executable, "-c", "import time; time.sleep(60)"]

    monkeypatch.setattr(inspection, "_worker_command", command)
    monkeypatch.setattr(inspection.subprocess, "Popen", launch)
    started = time.monotonic()
    outcome = inspection.inspect_window("Editor", timeout_seconds=0.2)
    assert time.monotonic() - started < 3
    assert outcome["ok"] is False and outcome["code"] == "timeout"
    assert outcome["data"]["tree"]["Name"] == "Partial"
    assert outcome["data"]["inspection"]["stage"] == "node-2.Name"
    assert outcome["data"]["inspection"]["worker_stopped"] is True
    assert len(launched) == 1 and launched[0].poll() is not None
    assert not worker_folders[0].exists()


def test_worker_crash_is_a_failure_and_next_call_can_succeed(monkeypatch):
    monkeypatch.setattr(
        inspection,
        "_worker_command",
        lambda _folder: [
            sys.executable,
            "-c",
            "raise SystemExit(3)",
        ],
    )
    failed = inspection.inspect_window("Editor", timeout_seconds=5)
    assert failed["code"] == "action_failed" and failed["ok"] is False

    def command(folder):
        inspection._write_state(
            folder,
            dict(
                done=True,
                outcome=dict(
                    Name="Editor",
                    ControlType="WindowControl",
                    _inspection=dict(complete=True),
                ),
            ),
        )
        return [sys.executable, "-c", "pass"]

    monkeypatch.setattr(inspection, "_worker_command", command)
    assert inspection.inspect_window("Editor", timeout_seconds=5)["_inspection"]["complete"]


def test_cooperative_worker_timeout_uses_same_stop_metadata(monkeypatch):
    def command(folder):
        inspection._write_state(
            folder,
            dict(
                done=True,
                outcome=inspection.result(
                    False,
                    "timeout",
                    "Deadline reached",
                    data=dict(tree={}, inspection=dict(complete=False)),
                ),
            ),
        )
        return [sys.executable, "-c", "pass"]

    monkeypatch.setattr(inspection, "_worker_command", command)
    value = inspection.inspect_window("Editor", timeout_seconds=5)
    assert value["data"]["inspection"]["timed_out"] is True
    assert value["data"]["inspection"]["worker_stopped"] is True


def test_mcp_entry_delegates_without_activation_or_root_search():
    source = ast.parse((_ROOT / "server.py").read_text(encoding="utf-8"))
    function = next(
        node
        for node in source.body
        if isinstance(node, ast.FunctionDef) and node.name == "get_control_tree"
    )
    inspect = Mock(return_value=dict(Name="Editor"))
    namespace = dict(
        json=json,
        control_tree=SimpleNamespace(inspect_window=inspect),
        normalize_result=lambda error: dict(ok=False, message=str(error)),
        mcp=SimpleNamespace(tool=lambda: lambda function: function),
    )
    exec(
        compile(ast.Module(body=[function], type_ignores=[]), str(_ROOT / "server.py"), "exec"),
        namespace,
    )
    assert json.loads(namespace["get_control_tree"]("Editor", 2, 12, 4)) == {"Name": "Editor"}
    inspect.assert_called_once_with("Editor", 2, max_nodes=12, timeout_seconds=4)
