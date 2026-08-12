"""Execute verified semantic UI commands from an easy_uiauto knowledge vault."""

from __future__ import annotations

import json
from pathlib import Path

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
    from PIL import Image

    expected_image = directory / record["image"]
    if not expected_image.is_file():
        raise RuntimeError("control reference image is missing")
    window = _find_window(record["app_name"])
    window_rect = _rect(window)
    current_screen = _screenshot_window(window_rect)
    try:
        control = resolve_location(record["location"])
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


def execute(directory: Path, command: str, text: str = "") -> dict:
    """Execute one verified UI CLI command and quarantine stale knowledge."""
    record, action = knowledge.resolve_command(directory, command)
    try:
        control, similarity, rectangle, resolved_by = _resolve_verified_control(directory, record)
    except Exception as error:
        reason = f"{type(error).__name__}: {error}"
        _quarantine(directory, record, reason)
        raise RuntimeError(
            f"Knowledge verification failed and the control was quarantined: {reason}. "
            "Rescan the current page before retrying."
        ) from error

    if action == "click":
        if control is not None:
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

    return {
        "ok": True,
        "command": command,
        "control_id": record["id"],
        "action": action,
        "image_similarity": similarity,
        "resolved_by": resolved_by,
        "location": record["location"],
    }


def execute_json(directory: Path, command: str, text: str = "") -> str:
    return json.dumps(execute(directory, command, text), ensure_ascii=False, indent=2)
