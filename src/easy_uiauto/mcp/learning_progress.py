"""Thread-safe status for long-running UI learning calls."""

from __future__ import annotations

import re
import threading
import time
import uuid
from datetime import UTC, datetime

_LOCK = threading.Lock()
_TASKS: dict[str, dict] = {}
_LATEST_TASK_ID = ""
_MAX_TASKS = 20


def _now() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds")


def _progress_for_message(message: str, previous: float) -> float:
    fixed = (
        ("Starting", 1),
        ("Capturing", 5),
        ("Segmenting", 10),
        ("Traversing", 25),
        ("Resolving", 45),
        ("Understanding", 45),
        ("Saving", 95),
        ("Finalizing", 98),
    )
    for prefix, value in fixed:
        if message.startswith(prefix):
            return max(previous, float(value))
    match = re.search(r"Verifying controls? (\d+)/(\d+)", message)
    if match:
        current = int(match.group(1))
        total = max(1, int(match.group(2)))
        return max(previous, 55 + 38 * min(current, total) / total)
    match = re.search(r"Exploring actions? (\d+)/(\d+)", message)
    if match:
        current = int(match.group(1))
        total = max(1, int(match.group(2)))
        return max(previous, 10 + 80 * min(current, total) / total)
    return previous


def begin(window_name: str, strategy: str, task_type: str = "scan") -> str:
    global _LATEST_TASK_ID
    task_id = f"learn-{uuid.uuid4().hex[:12]}"
    created = {
        "task_id": task_id,
        "task_type": task_type,
        "state": "running",
        "window_name": window_name,
        "strategy": strategy,
        "stage": "Starting UI learning",
        "progress_percent": 1.0,
        "started_at": _now(),
        "updated_at": _now(),
        "started_monotonic": time.monotonic(),
        "result": None,
        "error": "",
    }
    with _LOCK:
        duplicate = next(
            (
                task
                for task in _TASKS.values()
                if task["state"] == "running"
                and task.get("task_type") == task_type
                and task["window_name"].casefold() == window_name.casefold()
            ),
            None,
        )
        if duplicate is not None:
            elapsed = round(time.monotonic() - duplicate["started_monotonic"], 1)
            raise RuntimeError(
                f"UI {task_type} is already running for {window_name}: "
                f"task_id={duplicate['task_id']}, stage={duplicate['stage']}, "
                f"elapsed_seconds={elapsed}. Query get_ui_learning_status instead of "
                "starting a duplicate scan."
            )
        _TASKS[task_id] = created
        _LATEST_TASK_ID = task_id
        while len(_TASKS) > _MAX_TASKS:
            oldest = next(iter(_TASKS))
            if oldest == task_id:
                break
            _TASKS.pop(oldest, None)
    return task_id


def update(task_id: str, message: str) -> dict:
    with _LOCK:
        task = _TASKS[task_id]
        task["stage"] = message
        task["progress_percent"] = round(
            _progress_for_message(message, float(task["progress_percent"])),
            1,
        )
        task["updated_at"] = _now()
        return _public(task)


def complete(task_id: str, result: dict) -> dict:
    with _LOCK:
        task = _TASKS[task_id]
        public_result = (
            dict(result)
            if task.get("task_type") == "scan"
            else {
                key: result.get(key)
                for key in (
                    "actions_considered",
                    "actions_learned",
                    "stopped",
                    "elapsed_seconds",
                )
            }
        )
        task.update(
            {
                "state": "completed",
                "stage": f"UI {task.get('task_type', 'learning')} completed",
                "progress_percent": 100.0,
                "updated_at": _now(),
                "result": public_result,
            }
        )
        return _public(task)


def fail(task_id: str, error: BaseException) -> dict:
    with _LOCK:
        task = _TASKS[task_id]
        task.update(
            {
                "state": "failed",
                "stage": f"UI {task.get('task_type', 'learning')} failed",
                "updated_at": _now(),
                "error": f"{type(error).__name__}: {error}",
            }
        )
        return _public(task)


def get(task_id: str = "") -> dict:
    with _LOCK:
        selected = task_id or _LATEST_TASK_ID
        if not selected:
            return {"state": "idle", "detail": "No UI learning task has started."}
        task = _TASKS.get(selected)
        if task is None:
            raise KeyError(f"Unknown UI learning task: {selected}")
        return _public(task)


def _public(task: dict) -> dict:
    return {key: value for key, value in task.items() if key != "started_monotonic"} | {
        "elapsed_seconds": round(time.monotonic() - task["started_monotonic"], 1),
    }
