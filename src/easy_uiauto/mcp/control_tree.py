"""Read-only, bounded UIA inspection in a disposable worker process."""

from __future__ import annotations

import json
import math
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from .results import ActionError, result


def _write_state(folder: Path, value: dict):
    temporary = folder / "state.tmp"
    temporary.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    temporary.replace(folder / "state.json")


def _read_state(folder: Path) -> dict:
    try:
        return json.loads((folder / "state.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


class TreeCollector:
    """Never enumerate children beyond a depth/node limit; checkpoint before UIA."""

    def __init__(self, *, max_depth, max_nodes, deadline, checkpoint, clock=time.monotonic):
        self.max_depth, self.max_nodes = max_depth, max_nodes
        self.deadline, self.checkpoint, self.clock = deadline, checkpoint, clock
        self.tree = None
        self.nodes = 0
        self.limits = set()
        self.errors = []
        self.stage = "bind_window"
        self.last_checkpoint = float("-inf")

    def state(self):
        return dict(
            tree=self.tree,
            inspection=dict(
                stage=self.stage,
                nodes=self.nodes,
                max_depth=self.max_depth,
                max_nodes=self.max_nodes,
                limits_reached=sorted(self.limits),
                errors=self.errors,
                complete=not self.limits and not self.errors,
            ),
        )

    def step(self, stage, *, force=False):
        self.stage = stage
        now = self.clock()
        if force or now - self.last_checkpoint >= 0.2:
            self.checkpoint(self.state())
            self.last_checkpoint = now
        if now >= self.deadline:
            raise TimeoutError("UIA inspection reached its deadline")

    def error(self, error, stage):
        if len(self.errors) < 30:
            self.errors.append(dict(stage=stage, message=str(error)[:500]))

    def visit(self, control, depth=0, children=None):
        if self.nodes >= self.max_nodes:
            self.limits.add("max_nodes")
            return
        node = {}
        if children is None:
            self.tree = node
        else:
            children.append(node)
        self.nodes += 1
        for name in ("ControlTypeName", "Name", "ClassName", "AutomationId", "BoundingRectangle"):
            stage = f"node-{self.nodes}.depth-{depth}.{name}"
            self.step(stage)
            try:
                value = getattr(control, name)
                if name == "BoundingRectangle":
                    node["bounds"] = (
                        {key: getattr(value, key) for key in ("left", "top", "right", "bottom")}
                        if value is not None
                        else None
                    )
                else:
                    text = str(value or "")
                    node["ControlType" if name == "ControlTypeName" else name] = text[:2048]
                    if len(text) > 2048:
                        node.setdefault("_truncated_properties", []).append(name)
            except Exception as error:
                self.error(error, stage)
        if depth >= self.max_depth:
            self.limits.add("max_depth")
            return
        if self.nodes >= self.max_nodes:
            self.limits.add("max_nodes")
            return
        self.step(f"node-{self.nodes}.first_child")
        try:
            child = control.GetFirstChildControl()
        except Exception as error:
            self.error(error, self.stage)
            return
        child_nodes = []
        if child is not None:
            node["children"] = child_nodes
        for count in range(30):
            if child is None:
                break
            self.visit(child, depth + 1, child_nodes)
            if self.nodes >= self.max_nodes:
                self.limits.add("max_nodes")
                break
            if count == 29:
                self.limits.add("children_per_node")
                break
            self.step(f"depth-{depth}.next_sibling")
            try:
                child = child.GetNextSiblingControl()
            except Exception as error:
                self.error(error, self.stage)
                break


def _select_window(title):
    from . import native_windows

    windows = native_windows.enumerate_windows(visible_only=True)
    exact = [window for window in windows if window["title"] == title]
    candidates = exact or [
        window
        for window in windows
        if window["title"] and title.casefold() in window["title"].casefold()
    ]
    if not candidates:
        raise ActionError("window_invalid", f"No visible window matches {title!r}")
    if len(candidates) > 1:
        foreground = [window for window in candidates if window["foreground"]]
        if len(foreground) != 1:
            raise ActionError(
                "ambiguous_match", "Multiple visible windows match; use an exact title"
            )
        return foreground[0]
    return candidates[0]


def _worker(folder: Path):
    request = json.loads((folder / "request.json").read_text(encoding="utf-8"))
    collector = TreeCollector(
        max_depth=request["max_depth"],
        max_nodes=request["max_nodes"],
        deadline=request["deadline"],
        checkpoint=lambda state: _write_state(folder, state),
    )
    try:
        collector.step("win32_window_lookup", force=True)
        window = _select_window(request["window_name"])
        collector.step("uia_initialization", force=True)
        import uiautomation

        with uiautomation.UIAutomationInitializerInThread():
            collector.step("bind_window", force=True)
            control = uiautomation.ControlFromHandle(window["handle"])
            if control is None:
                raise ActionError("window_invalid", "Could not bind UIA to the selected window")
            collector.visit(control)
            collector.step("uia_cleanup", force=True)
        state = collector.state()
        state["inspection"].update(stage="finished", window_title=window["title"])
        tree = state["tree"]
        tree["_inspection"] = state["inspection"]
        _write_state(folder, dict(done=True, outcome=tree))
    except Exception as error:
        state = collector.state()
        state["inspection"]["complete"] = False
        code = (
            "timeout"
            if isinstance(error, TimeoutError)
            else getattr(error, "code", "action_failed")
        )
        _write_state(
            folder,
            dict(
                done=True,
                outcome=result(
                    False,
                    code,
                    str(error),
                    retryable=isinstance(error, TimeoutError),
                    data=state,
                ),
            ),
        )


def _worker_command(folder):
    return [sys.executable, "-m", "easy_uiauto.mcp.control_tree", "--worker", str(folder)]


def inspect_window(
    window_name: str, max_depth: int = 5, *, max_nodes: int = 500, timeout_seconds: float = 15.0
) -> dict:
    """Bound even a blocked COM call; never activate, restore or click a window."""
    if not isinstance(window_name, str) or not window_name.strip():
        raise ValueError("window_name must not be empty")
    if isinstance(max_depth, bool) or not isinstance(max_depth, int) or max_depth < 0:
        raise ValueError("max_depth must be a nonnegative integer (capped at 10)")
    if isinstance(max_nodes, bool) or not isinstance(max_nodes, int) or not 1 <= max_nodes <= 5000:
        raise ValueError("max_nodes must be an integer from 1 to 5000")
    if (
        isinstance(timeout_seconds, bool)
        or not isinstance(timeout_seconds, (int, float))
        or not math.isfinite(timeout_seconds)
        or not 0.1 <= timeout_seconds <= 60
    ):
        raise ValueError("timeout_seconds must be between 0.1 and 60")
    started = time.monotonic()
    deadline = started + timeout_seconds
    request = dict(
        window_name=window_name.strip(),
        max_depth=min(max_depth, 10),
        max_nodes=max_nodes,
        deadline=deadline,
    )
    with tempfile.TemporaryDirectory(prefix="easy-uiauto-tree-") as temporary:
        folder = Path(temporary)
        (folder / "request.json").write_text(json.dumps(request), encoding="utf-8")
        process = subprocess.Popen(
            _worker_command(folder),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        timed_out = False
        try:
            try:
                process.wait(timeout=max(0.001, deadline - time.monotonic()))
            except subprocess.TimeoutExpired:
                timed_out = True
            finally:
                if process.poll() is None:
                    process.kill()  # Stop only this disposable inspection worker.
                process.wait(timeout=2)
            state = _read_state(folder)
            if not timed_out and state.get("done"):
                outcome = state["outcome"]
                inspection = outcome.get("_inspection") or outcome.get("data", {}).get("inspection")
                if inspection is not None:
                    inspection["elapsed_seconds"] = round(time.monotonic() - started, 3)
                    if outcome.get("code") == "timeout":
                        inspection.update(timed_out=True, worker_stopped=True)
                return outcome
            if state.get("done"):
                outcome = state["outcome"]
                state = (
                    dict(tree=outcome, inspection=outcome.get("_inspection", {}))
                    if "ControlType" in outcome
                    else outcome.get("data", {})
                )
            state.setdefault("tree", None)
            state.setdefault("inspection", {})
            state["inspection"].setdefault("stage", "worker_startup")
            state["inspection"].update(
                complete=False,
                timed_out=timed_out,
                elapsed_seconds=round(time.monotonic() - started, 3),
                worker_stopped=True,
            )
            return result(
                False,
                "timeout" if timed_out else "action_failed",
                "UIA inspection timed out; its worker was stopped. Check the partial checkpoint."
                if timed_out
                else f"UIA inspection worker exited without a result ({process.returncode})",
                retryable=timed_out,
                data=state,
            )
        finally:
            if process.poll() is None:
                process.kill()
                process.wait(timeout=2)


if __name__ == "__main__":
    if len(sys.argv) != 3 or sys.argv[1] != "--worker":
        raise SystemExit("This module is an internal UI inspection worker")
    _worker(Path(sys.argv[2]))
