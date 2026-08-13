"""Tests for visible and recoverable MCP learning progress."""

from __future__ import annotations

import asyncio
import json
import time

from easy_uiauto.mcp import learning_progress, server


class FakeContext:
    def __init__(self):
        self.progress = []
        self.logs = []

    async def report_progress(self, progress, total=None, message=None):
        self.progress.append((progress, total, message))

    async def info(self, message, **_extra):
        self.logs.append(message)


def test_learning_progress_reports_control_counts_and_completion() -> None:
    task_id = learning_progress.begin("Example", "visual-first")

    status = learning_progress.update(task_id, "Verifying controls 5/10")
    completed = learning_progress.complete(
        task_id,
        {
            "app_id": "example",
            "page_id": "main",
            "controls_saved_this_scan": 10,
            "commands": 8,
            "annotated_page_image": "main.annotated.png",
            "elapsed_seconds": 12.3,
        },
    )

    assert status["state"] == "running"
    assert status["progress_percent"] == 74.0
    assert completed["state"] == "completed"
    assert completed["progress_percent"] == 100.0
    assert learning_progress.get(task_id)["result"]["commands"] == 8


def test_learning_progress_rejects_duplicate_window_scan() -> None:
    task_id = learning_progress.begin("Duplicate Example", "visual-first")

    try:
        learning_progress.begin("duplicate example", "visual-first")
    except RuntimeError as error:
        assert task_id in str(error)
        assert "get_ui_learning_status" in str(error)
    else:
        raise AssertionError("duplicate scan was accepted")
    finally:
        learning_progress.fail(task_id, RuntimeError("test cleanup"))


def test_scan_tool_emits_progress_and_preserves_task_status(monkeypatch) -> None:
    context = FakeContext()
    monkeypatch.setattr(
        server,
        "_vision_api_settings",
        lambda _model: ("https://api.example", "secret", "vision-model"),
    )

    def scan_window(**kwargs):
        kwargs["progress"]("Capturing window screenshot")
        kwargs["progress"]("Segmenting interface with remote AI")
        kwargs["progress"]("Verifying controls 2/2")
        return {
            "ok": True,
            "app_id": "example",
            "page_id": "main",
            "controls_saved_this_scan": 2,
            "commands": 1,
            "annotated_page_image": "main.annotated.png",
            "elapsed_seconds": 0.1,
        }

    monkeypatch.setattr(server.scanner, "scan_window", scan_window)
    monkeypatch.setattr(server, "LEARNING_HEARTBEAT_SECONDS", 0.01)

    result = json.loads(
        asyncio.run(server.scan_window_knowledge("Example", context, show_overlay=False))
    )
    status = learning_progress.get(result["learning_task_id"])

    assert status["state"] == "completed"
    assert status["result"]["controls_saved_this_scan"] == 2
    assert context.progress[0][1] == 100
    assert context.progress[-1][0] == 100.0
    assert result["learning_task_id"] in context.progress[-1][2]


def test_long_scan_sends_progress_heartbeats(monkeypatch) -> None:
    context = FakeContext()
    monkeypatch.setattr(
        server,
        "_vision_api_settings",
        lambda _model: ("https://api.example", "secret", "vision-model"),
    )
    monkeypatch.setattr(server, "LEARNING_HEARTBEAT_SECONDS", 0.01)

    def scan_window(**kwargs):
        kwargs["progress"]("Segmenting interface with remote AI")
        time.sleep(0.04)
        return {
            "app_id": "example",
            "page_id": "main",
            "controls_saved_this_scan": 1,
            "commands": 1,
            "annotated_page_image": "main.annotated.png",
            "elapsed_seconds": 0.04,
        }

    monkeypatch.setattr(server.scanner, "scan_window", scan_window)

    asyncio.run(server.scan_window_knowledge("Example", context, show_overlay=False))

    assert len(context.progress) >= 4
    assert any("Segmenting interface" in item[2] for item in context.progress)


def test_scan_continues_after_client_wait_is_cancelled(monkeypatch) -> None:
    context = FakeContext()
    monkeypatch.setattr(
        server,
        "_vision_api_settings",
        lambda _model: ("https://api.example", "secret", "vision-model"),
    )

    def scan_window(**kwargs):
        kwargs["progress"]("Segmenting interface with remote AI")
        time.sleep(0.08)
        return {
            "ok": True,
            "app_id": "example",
            "page_id": "main",
            "controls_saved_this_scan": 1,
            "commands": 1,
            "annotated_page_image": "main.annotated.png",
            "elapsed_seconds": 0.08,
        }

    monkeypatch.setattr(server.scanner, "scan_window", scan_window)

    async def cancel_wait():
        call = asyncio.create_task(
            server.scan_window_knowledge("Example", context, show_overlay=False)
        )
        await asyncio.sleep(0.01)
        call.cancel()
        try:
            await call
        except asyncio.CancelledError:
            pass
        await asyncio.sleep(0.12)

    asyncio.run(cancel_wait())
    status = json.loads(server.get_ui_learning_status())

    assert status["state"] == "completed"
    assert status["result"]["app_id"] == "example"


def test_exploration_tool_runs_off_event_loop_and_reports_status(monkeypatch) -> None:
    context = FakeContext()
    monkeypatch.setattr(server, "_mode_blocks_operation", lambda: "")
    monkeypatch.setattr(
        server,
        "_vision_api_settings",
        lambda _model: ("https://api.example", "secret", "vision-model"),
    )
    monkeypatch.setattr(server, "LEARNING_HEARTBEAT_SECONDS", 0.01)

    def explore(_directory, *_args, progress, **_kwargs):
        progress("Exploring action 1/2: main.help.click")
        time.sleep(0.02)
        progress("Exploring action 2/2: help.close.click")
        return {
            "ok": True,
            "actions_considered": 2,
            "actions_learned": 2,
            "stopped": "",
        }

    monkeypatch.setattr(server.interaction_learning, "explore_application", explore)

    result = json.loads(asyncio.run(server.explore_ui_workflows("example", context)))
    status = learning_progress.get(result["learning_task_id"])

    assert status["state"] == "completed"
    assert status["task_type"] == "exploration"
    assert status["result"]["actions_learned"] == 2
    assert any("Exploring action" in item[2] for item in context.progress)
