"""Execute verified semantic UI commands from an easy_uiauto knowledge vault."""

from __future__ import annotations

import json
from pathlib import Path
from time import perf_counter

import pyautogui

from . import knowledge
from .scanner import (
    _crop_control,
    _find_window,
    _image_similarity,
    _rect,
    _screenshot_window,
    locate_template,
    resolve_location,
)


def _quarantine(directory: Path, record: dict, reason: str) -> None:
    record = {
        **record,
        "status": "quarantined",
        "notes": f"{record.get('notes', '')}\nQuarantined: {reason}".strip(),
        "verification": {
            **record.get("verification", {}),
            "runtime": "failed",
            "runtime_detail": reason,
            "verified_at": knowledge.utc_now(),
        },
    }
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    knowledge.write_command_catalog(directory)


def _resolve_verified_control(directory: Path, record: dict):
    window = _find_window(record["app_name"])
    window_rect = _rect(window)
    current_screen = _screenshot_window(window_rect)
    return _resolve_verified_control_in_context(
        directory,
        record,
        window,
        window_rect,
        current_screen,
    )


def _resolve_verified_control_in_context(
    directory: Path,
    record: dict,
    window,
    window_rect: dict,
    current_screen,
):
    """Resolve one control against a shared window screenshot."""
    from PIL import Image

    expected_image = directory / record["image"]
    if not expected_image.is_file():
        raise RuntimeError("control reference image is missing")
    try:
        control = resolve_location(record["location"], window=window)
    except Exception as error:
        control = None
        location_error = str(error)
    else:
        location_error = "LOCATION did not resolve"
    with Image.open(expected_image) as reference:
        if control and control is not False:
            current_rect = _rect(control)
            current_crop = _crop_control(current_screen, current_rect, window_rect)
            similarity = _image_similarity(reference, current_crop)
            if similarity >= 0.72:
                return control, similarity, current_rect, "location"
        template = locate_template(reference, current_screen, window_rect)
    if template["found"]:
        return None, template["score"], template["rect"], "image"
    raise RuntimeError(
        f"{location_error}; image fallback failed: {template['detail']}"
    )


def _perform_action(
    action: str,
    text: str,
    control,
    rectangle: dict,
    fast: bool = False,
) -> None:
    if action == "click":
        if control is not None:
            if fast:
                control.Click(simulateMove=False, waitTime=0)
            else:
                control.Click()
        else:
            pyautogui.click(
                rectangle["left"] + rectangle["width"] // 2,
                rectangle["top"] + rectangle["height"] // 2,
            )
    elif action == "double-click":
        pyautogui.doubleClick(
            rectangle["left"] + rectangle["width"] // 2,
            rectangle["top"] + rectangle["height"] // 2,
        )
    elif action == "set-text":
        if not text:
            raise ValueError("text is required for set-text commands")
        try:
            if control is None:
                raise RuntimeError("set-text resolved by image")
            control.GetValuePattern().SetValue(text)
        except Exception:
            if control is not None:
                if fast:
                    control.Click(simulateMove=False, waitTime=0)
                else:
                    control.Click()
            else:
                pyautogui.click(
                    rectangle["left"] + rectangle["width"] // 2,
                    rectangle["top"] + rectangle["height"] // 2,
                )
            pyautogui.hotkey("ctrl", "a")
            pyautogui.write(text)
    else:
        raise ValueError(f"Unsupported UI command action: {action}")


def _record_execution(record: dict, count: int = 1) -> None:
    verification = record.get("function_verification", {})
    record["function_verification"] = {
        **verification,
        "executed": True,
        "execution_count": int(verification.get("execution_count", 0)) + count,
        "last_executed_at": knowledge.utc_now(),
    }


def execute(directory: Path, command: str, text: str = "", confirm: bool = False) -> dict:
    """Execute one verified UI CLI command and quarantine stale knowledge."""
    record, action = knowledge.resolve_command(directory, command)
    if record.get("requires_confirmation") and not confirm:
        raise RuntimeError(
            f"Command requires explicit confirmation because risk is "
            f"{record.get('risk', 'unknown')}: {command}"
        )
    try:
        control, similarity, rectangle, resolved_by = _resolve_verified_control(directory, record)
    except Exception as error:
        reason = f"{type(error).__name__}: {error}"
        _quarantine(directory, record, reason)
        raise RuntimeError(
            f"Knowledge verification failed and the control was quarantined: {reason}. "
            "Rescan the current page before retrying."
        ) from error

    _perform_action(action, text, control, rectangle)
    _record_execution(record)
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)

    return {
        "ok": True,
        "command": command,
        "control_id": record["id"],
        "action": action,
        "image_similarity": similarity,
        "resolved_by": resolved_by,
        "location": record["location"],
    }


def execute_json(directory: Path, command: str, text: str = "", confirm: bool = False) -> str:
    return json.dumps(execute(directory, command, text, confirm), ensure_ascii=False, indent=2)


def _normalize_steps(steps: list[str | dict]) -> list[dict]:
    if not isinstance(steps, list) or not steps:
        raise ValueError("steps must be a non-empty JSON array")
    if len(steps) > 500:
        raise ValueError("a batch is limited to 500 steps")
    normalized = []
    for index, item in enumerate(steps):
        if isinstance(item, str):
            step = {"command": item, "text": ""}
        elif isinstance(item, dict):
            step = {
                "command": str(item.get("command", "")).strip(),
                "text": str(item.get("text", "")),
            }
        else:
            raise ValueError(f"step {index} must be a command string or object")
        if not step["command"]:
            raise ValueError(f"step {index} has no command")
        normalized.append(step)
    return normalized


def execute_many(
    directory: Path,
    steps: list[str | dict],
    confirm: bool = False,
) -> dict:
    """Preflight and execute a same-page command sequence in one process."""
    started = perf_counter()
    normalized = _normalize_steps(steps)
    command_catalog = {
        item["command"]: item for item in knowledge.available_commands(directory)
    }
    records: dict[str, dict] = {}
    prepared = []
    for index, step in enumerate(normalized):
        item = command_catalog.get(step["command"])
        if item is None:
            raise KeyError(f"Unknown or unverified UI command: {step['command']}")
        control_id = item["control_id"]
        if control_id not in records:
            found = knowledge.find_control_record(directory, control_id)
            if found is None:
                raise KeyError(f"Missing control knowledge: {control_id}")
            records[control_id] = found[1]
        record = records[control_id]
        if record.get("requires_confirmation") and not confirm:
            raise RuntimeError(
                "Command requires explicit confirmation because risk is "
                f"{record.get('risk', 'unknown')}: {step['command']}"
            )
        prepared.append(
            {
                "index": index,
                **step,
                "action": item["action"],
                "control_id": control_id,
                "record": record,
            }
        )

    pages = {item["record"].get("page_id", "unknown") for item in prepared}
    if len(pages) != 1:
        raise ValueError(
            "batch commands must belong to one page; split navigation into separate batches"
        )
    app_names = {item["record"].get("app_name", "") for item in prepared}
    if len(app_names) != 1:
        raise ValueError("batch commands must belong to one application window")

    window = _find_window(prepared[0]["record"]["app_name"])
    window_rect = _rect(window)
    current_screen = _screenshot_window(window_rect)
    resolved: dict[str, tuple] = {}
    for control_id, record in records.items():
        try:
            resolved[control_id] = _resolve_verified_control_in_context(
                directory,
                record,
                window,
                window_rect,
                current_screen,
            )
        except Exception as error:
            reason = f"{type(error).__name__}: {error}"
            _quarantine(directory, record, reason)
            raise RuntimeError(
                "Batch preflight failed before any action and the control was quarantined: "
                f"{reason}. Rescan the current page before retrying."
            ) from error
    preflight_finished = perf_counter()

    results = []
    execution_counts: dict[str, int] = {}
    failed_step = None
    for item in prepared:
        control, similarity, rectangle, resolved_by = resolved[item["control_id"]]
        try:
            _perform_action(item["action"], item["text"], control, rectangle, fast=True)
        except Exception as error:
            failed_step = {
                "index": item["index"],
                "command": item["command"],
                "error": f"{type(error).__name__}: {error}",
            }
            break
        execution_counts[item["control_id"]] = (
            execution_counts.get(item["control_id"], 0) + 1
        )
        results.append(
            {
                "index": item["index"],
                "command": item["command"],
                "control_id": item["control_id"],
                "action": item["action"],
                "image_similarity": similarity,
                "resolved_by": resolved_by,
            }
        )

    for control_id, count in execution_counts.items():
        record = records[control_id]
        _record_execution(record, count)
        knowledge.save_control(directory, record)
    if execution_counts:
        knowledge.rebuild_index(directory)

    finished = perf_counter()
    return {
        "ok": failed_step is None,
        "page": next(iter(pages)),
        "requested_steps": len(prepared),
        "completed_steps": len(results),
        "unique_controls_verified": len(resolved),
        "knowledge_writes": len(execution_counts),
        "preflight_seconds": round(preflight_finished - started, 3),
        "execution_seconds": round(finished - preflight_finished, 3),
        "elapsed_seconds": round(finished - started, 3),
        "steps": results,
        "failed_step": failed_step,
    }


def execute_many_json(
    directory: Path,
    steps: list[str | dict],
    confirm: bool = False,
) -> str:
    return json.dumps(execute_many(directory, steps, confirm), ensure_ascii=False, indent=2)
