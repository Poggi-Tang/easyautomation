"""Execute verified semantic UI commands from an easy_uiauto knowledge vault."""

from __future__ import annotations

import json
import os
from pathlib import Path
from time import perf_counter, sleep
from urllib import request as urlrequest
from uuid import uuid4

import pyautogui
import pyperclip

from . import configuration, knowledge, visualization
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


def _resolve_verified_control(
    directory: Path,
    record: dict,
    allow_vision_fallback: bool = False,
):
    window = _find_window(record["app_name"])
    window_rect = _rect(window)
    current_screen = _screenshot_window(window_rect)
    return _resolve_verified_control_in_context(
        directory,
        record,
        window,
        window_rect,
        current_screen,
        allow_vision_fallback,
    )


def _resolve_verified_control_in_context(
    directory: Path,
    record: dict,
    window,
    window_rect: dict,
    current_screen,
    allow_vision_fallback: bool = False,
):
    """Resolve through LOCATION, templates, OCR, then optional remote vision."""
    from PIL import Image

    image_names = list(record.get("image_variants", []))
    if record.get("image"):
        image_names.append(record["image"])
    expected_images = [directory / value for value in dict.fromkeys(image_names)]
    expected_images = [path for path in expected_images if path.is_file()]
    if not expected_images:
        raise RuntimeError("control reference image is missing")
    try:
        control = resolve_location(record["location"], window=window)
    except Exception as error:
        control = None
        location_error = str(error)
    else:
        location_error = "LOCATION did not resolve"
    best_template = None
    best_similarity = 0.0
    for expected_image in expected_images:
        with Image.open(expected_image) as reference:
            if control and control is not False:
                current_rect = _rect(control)
                current_crop = _crop_control(current_screen, current_rect, window_rect)
                similarity = _image_similarity(reference, current_crop)
                best_similarity = max(best_similarity, similarity)
                if similarity >= 0.72:
                    return control, similarity, current_rect, "location"
            template = locate_template(reference, current_screen, window_rect)
        if best_template is None or template["score"] > best_template["score"]:
            best_template = template
    if best_template and best_template["found"]:
        return None, best_template["score"], best_template["rect"], "image"
    ocr = _locate_record_text(record, current_screen, window_rect)
    if ocr:
        return None, ocr["confidence"], ocr["rect"], "ocr"
    if allow_vision_fallback:
        visual = _locate_record_by_vision(record, current_screen, window_rect)
        if visual:
            return None, visual["confidence"], visual["rect"], "ai-vision"
    template_detail = best_template["detail"] if best_template else "not checked"
    raise RuntimeError(
        f"{location_error}; image fallback failed: {template_detail}; "
        "OCR fallback failed"
    )


def _locate_record_text(record: dict, screenshot, window_rect: dict) -> dict | None:
    try:
        import pytesseract
        from pytesseract import Output
    except ImportError:
        return None
    candidates = [
        record.get("name", ""),
        record.get("semantic_name", ""),
        *record.get("aliases", []),
    ]
    candidates = {
        " ".join(str(value).casefold().split())
        for value in candidates
        if len(str(value).strip()) >= 1
    }
    if not candidates:
        return None
    language = os.environ.get("EASY_UIAUTO_OCR_LANGUAGE", "eng")
    try:
        data = pytesseract.image_to_data(screenshot, lang=language, output_type=Output.DICT)
    except Exception:
        return None
    lines = {}
    for index, text in enumerate(data.get("text", [])):
        value = str(text).strip()
        if not value:
            continue
        key = (
            data.get("block_num", [0])[index],
            data.get("par_num", [0])[index],
            data.get("line_num", [0])[index],
        )
        lines.setdefault(key, []).append(index)
    matches = []
    for indexes in lines.values():
        value = " ".join(str(data["text"][index]).strip() for index in indexes)
        normalized = " ".join(value.casefold().split())
        if normalized not in candidates:
            continue
        confidences = []
        for index in indexes:
            try:
                confidences.append(float(data["conf"][index]))
            except (TypeError, ValueError):
                pass
        confidence = (sum(confidences) / len(confidences) / 100) if confidences else 0
        if confidence < 0.5:
            continue
        left = min(int(data["left"][index]) for index in indexes)
        top = min(int(data["top"][index]) for index in indexes)
        right = max(
            int(data["left"][index]) + int(data["width"][index]) for index in indexes
        )
        bottom = max(
            int(data["top"][index]) + int(data["height"][index]) for index in indexes
        )
        matches.append(
            {
                "confidence": round(confidence, 4),
                "rect": {
                    "left": window_rect["left"] + left,
                    "top": window_rect["top"] + top,
                    "right": window_rect["left"] + right,
                    "bottom": window_rect["top"] + bottom,
                    "width": right - left,
                    "height": bottom - top,
                },
            }
        )
    return matches[0] if len(matches) == 1 else None


def _locate_record_by_vision(record: dict, screenshot, window_rect: dict) -> dict | None:
    from .scanner import _image_data_url, _parse_json_object

    api_url = configuration._existing_vision_value(configuration.VISION_API_URL).strip()
    api_key = configuration._existing_vision_value(configuration.VISION_API_KEY).strip()
    model = configuration._existing_vision_value(configuration.VISION_MODEL).strip()
    if not api_url or not api_key or not model:
        return None
    description = record.get("description") or record.get("semantic_name") or record.get("name")
    payload = {
        "model": model,
        "temperature": 0,
        "messages": [
            {
                "role": "system",
                "content": "Locate one desktop UI control and return strict JSON only.",
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            f"Locate this control: {description}. Return "
                            '{"found":true,"left":0,"top":0,"width":10,'
                            '"height":10,"confidence":0.95}. Coordinates use image pixels. '
                            "Return found=false if ambiguous or absent."
                        ),
                    },
                    {"type": "image_url", "image_url": {"url": _image_data_url(screenshot)}},
                ],
            },
        ],
    }
    try:
        request = urlrequest.Request(
            api_url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with urlrequest.urlopen(request, timeout=30) as response:
            body = json.loads(response.read().decode("utf-8"))
        content = body["choices"][0]["message"]["content"]
        if isinstance(content, list):
            content = "".join(
                item.get("text", "") for item in content if isinstance(item, dict)
            )
        result = _parse_json_object(str(content))
        confidence = float(result.get("confidence", 0))
        left = int(result.get("left", 0))
        top = int(result.get("top", 0))
        width = int(result.get("width", 0))
        height = int(result.get("height", 0))
    except Exception:
        return None
    if not result.get("found") or confidence < 0.8 or width <= 1 or height <= 1:
        return None
    if left < 0 or top < 0 or left + width > screenshot.width or top + height > screenshot.height:
        return None
    return {
        "confidence": round(confidence, 4),
        "rect": {
            "left": window_rect["left"] + left,
            "top": window_rect["top"] + top,
            "right": window_rect["left"] + left + width,
            "bottom": window_rect["top"] + top + height,
            "width": width,
            "height": height,
        },
    }


def _read_control_text(control) -> tuple[bool, str]:
    if control is None:
        return False, ""
    readers = (
        lambda: control.GetValuePattern().Value,
        lambda: control.GetLegacyIAccessiblePattern().Value,
        lambda: control.GetTextPattern().DocumentRange.GetText(-1),
    )
    for reader in readers:
        try:
            value = reader()
        except Exception:
            continue
        if value is not None:
            return True, str(value)
    return False, ""


def _capture_action_image(rectangle: dict):
    from PIL import ImageGrab

    bounds = (
        rectangle["left"],
        rectangle["top"],
        rectangle["left"] + rectangle["width"],
        rectangle["top"] + rectangle["height"],
    )
    try:
        return ImageGrab.grab(bbox=bounds, all_screens=True)
    except Exception:
        return pyautogui.screenshot(
            region=(
                rectangle["left"],
                rectangle["top"],
                rectangle["width"],
                rectangle["height"],
            )
        )


def _visual_change_ratio(before, after) -> float:
    from PIL import ImageChops

    if before.size != after.size:
        return 1.0
    difference = ImageChops.difference(before.convert("RGB"), after.convert("RGB")).convert("L")
    changed = difference.point(lambda value: 255 if value > 12 else 0)
    histogram = changed.histogram()
    pixels = max(1, before.width * before.height)
    return float(histogram[255]) / pixels


def _text_write_evidence(control, expected: str, before, after) -> dict:
    readable, value = _read_control_text(control)
    normalized_value = value.replace("\r\n", "\n")
    normalized_expected = expected.replace("\r\n", "\n")
    if readable and normalized_value == normalized_expected:
        return {
            "verified": True,
            "evidence": "accessible-value",
            "observed_value": value,
            "visual_change_ratio": 0.0,
        }
    ratio = _visual_change_ratio(before, after)
    return {
        "verified": False,
        "evidence": "visual-change-only" if ratio >= 0.0005 else "no-observable-change",
        "observed_value": value if readable else None,
        "visual_change_ratio": round(ratio, 6),
    }


def _paste_text_with_clipboard(text: str, control, rectangle: dict, fast: bool) -> dict:
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
    sleep(0.05)
    before = _capture_action_image(rectangle)
    try:
        previous_clipboard = pyperclip.paste()
    except Exception:
        previous_clipboard = None
    try:
        pyperclip.copy(text)
        pyautogui.hotkey("ctrl", "v")
        sleep(0.12)
        after = _capture_action_image(rectangle)
        evidence = _text_write_evidence(control, text, before, after)
        if not evidence["verified"]:
            sentinel = f"__easy_uiauto_verify_{uuid4().hex}__"
            pyperclip.copy(sentinel)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.hotkey("ctrl", "c")
            sleep(0.05)
            try:
                copied_back = pyperclip.paste()
            except Exception:
                copied_back = sentinel
            pyautogui.press("end")
            if copied_back != sentinel and copied_back.replace("\r\n", "\n") == text.replace(
                "\r\n", "\n"
            ):
                evidence.update(
                    {
                        "verified": True,
                        "evidence": "clipboard-readback",
                        "observed_value": copied_back,
                    }
                )
    finally:
        if previous_clipboard is not None:
            sleep(0.05)
            try:
                pyperclip.copy(previous_clipboard)
            except Exception:
                pass
    if not evidence["verified"]:
        raise RuntimeError(
            "Text paste could not be read back from the target control; "
            "the command was not recorded as successful"
        )
    return {"method": "clipboard-paste", **evidence}


def _perform_action(
    action: str,
    text: str,
    control,
    rectangle: dict,
    fast: bool = False,
) -> dict:
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
        return {"method": "uia-click" if control is not None else "coordinate-click"}
    if action == "double-click":
        pyautogui.doubleClick(
            rectangle["left"] + rectangle["width"] // 2,
            rectangle["top"] + rectangle["height"] // 2,
        )
        return {"method": "coordinate-double-click"}
    if action == "right-click":
        pyautogui.rightClick(
            rectangle["left"] + rectangle["width"] // 2,
            rectangle["top"] + rectangle["height"] // 2,
        )
        return {"method": "coordinate-right-click"}
    if action == "hover":
        pyautogui.moveTo(
            rectangle["left"] + rectangle["width"] // 2,
            rectangle["top"] + rectangle["height"] // 2,
            duration=0 if fast else 0.1,
        )
        return {"method": "coordinate-hover"}
    if action == "set-text":
        if not text:
            raise ValueError("text is required for set-text commands")
        if control is not None:
            before = _capture_action_image(rectangle)
            try:
                control.GetValuePattern().SetValue(text)
                sleep(0.1)
                after = _capture_action_image(rectangle)
                evidence = _text_write_evidence(control, text, before, after)
                if evidence["verified"]:
                    return {"method": "uia-set-value", **evidence}
            except Exception:
                pass
        return _paste_text_with_clipboard(text, control, rectangle, fast)
    raise ValueError(f"Unsupported UI command action: {action}")


def _record_execution(record: dict, count: int = 1, evidence: dict | None = None) -> None:
    verification = record.get("function_verification", {})
    record["function_verification"] = {
        **verification,
        "executed": True,
        "execution_count": int(verification.get("execution_count", 0)) + count,
        "last_executed_at": knowledge.utc_now(),
        "last_action_verification": evidence or {},
    }


def _show_execution_preview(
    records: list[dict],
    rectangles: dict[str, dict],
    duration_ms: int,
    wait_ms: int,
) -> dict:
    markers = []
    for record in records:
        rectangle = rectangles.get(record["id"])
        if not rectangle:
            continue
        markers.append(
            {
                "index": len(markers) + 1,
                "control_id": record["id"],
                "label": record.get("semantic_name") or record.get("name") or record["id"],
                "rect": rectangle,
                "status": record.get("status", "verified"),
                "color": visualization.STATUS_COLORS["verified"],
                "target": True,
            }
        )
    return visualization.show_markers(markers, duration_ms, wait_ms)


def execute(
    directory: Path,
    command: str,
    text: str = "",
    confirm: bool = False,
    allow_vision_fallback: bool = False,
    highlight: bool = False,
    highlight_duration_ms: int = 900,
    highlight_wait_ms: int = 100,
) -> dict:
    """Execute one verified UI CLI command and quarantine stale knowledge."""
    record, action = knowledge.resolve_command(directory, command)
    if record.get("requires_confirmation") and not confirm:
        raise RuntimeError(
            f"Command requires explicit confirmation because risk is "
            f"{record.get('risk', 'unknown')}: {command}"
        )
    try:
        control, similarity, rectangle, resolved_by = _resolve_verified_control(
            directory,
            record,
            allow_vision_fallback,
        )
    except Exception as error:
        reason = f"{type(error).__name__}: {error}"
        _quarantine(directory, record, reason)
        raise RuntimeError(
            f"Knowledge verification failed and the control was quarantined: {reason}. "
            "Rescan the current page before retrying."
        ) from error

    overlay = {"shown": False, "controls": 0, "detail": "disabled"}
    if highlight:
        overlay = _show_execution_preview(
            [record],
            {record["id"]: rectangle},
            highlight_duration_ms,
            highlight_wait_ms,
        )
    if control is not None and action != "hover" and not bool(
        getattr(control, "IsEnabled", True)
    ):
        condition = record.get("enabling_condition") or "the control must become enabled"
        raise RuntimeError(f"Control is currently disabled; required condition: {condition}")
    action_verification = _perform_action(action, text, control, rectangle)
    _record_execution(record, evidence=action_verification)
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
        "overlay": overlay,
        "action_verification": action_verification,
    }


def execute_json(
    directory: Path,
    command: str,
    text: str = "",
    confirm: bool = False,
    allow_vision_fallback: bool = False,
    highlight: bool = False,
    highlight_duration_ms: int = 900,
    highlight_wait_ms: int = 100,
) -> str:
    return json.dumps(
        execute(
            directory,
            command,
            text,
            confirm,
            allow_vision_fallback,
            highlight,
            highlight_duration_ms,
            highlight_wait_ms,
        ),
        ensure_ascii=False,
        indent=2,
    )


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
    allow_vision_fallback: bool = False,
    highlight: bool = False,
    highlight_duration_ms: int = 1200,
    highlight_wait_ms: int = 100,
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
            resolver_args = (directory, record, window, window_rect, current_screen)
            resolved[control_id] = (
                _resolve_verified_control_in_context(*resolver_args, True)
                if allow_vision_fallback
                else _resolve_verified_control_in_context(*resolver_args)
            )
        except Exception as error:
            reason = f"{type(error).__name__}: {error}"
            _quarantine(directory, record, reason)
            raise RuntimeError(
                "Batch preflight failed before any action and the control was quarantined: "
                f"{reason}. Rescan the current page before retrying."
            ) from error
    preflight_finished = perf_counter()

    overlay = {"shown": False, "controls": 0, "detail": "disabled"}
    if highlight:
        overlay = _show_execution_preview(
            list(records.values()),
            {control_id: value[2] for control_id, value in resolved.items()},
            highlight_duration_ms,
            highlight_wait_ms,
        )

    results = []
    execution_counts: dict[str, int] = {}
    failed_step = None
    for item in prepared:
        control, similarity, rectangle, resolved_by = resolved[item["control_id"]]
        try:
            if control is not None and item["action"] != "hover" and not bool(
                getattr(control, "IsEnabled", True)
            ):
                condition = item["record"].get("enabling_condition") or (
                    "the control must become enabled"
                )
                raise RuntimeError(
                    f"Control is currently disabled; required condition: {condition}"
                )
            action_verification = _perform_action(
                item["action"], item["text"], control, rectangle, fast=True
            )
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
                "action_verification": action_verification,
            }
        )

    for control_id, count in execution_counts.items():
        record = records[control_id]
        evidence = next(
            (
                item["action_verification"]
                for item in reversed(results)
                if item["control_id"] == control_id
            ),
            {},
        )
        _record_execution(record, count, evidence)
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
        "overlay": overlay,
    }


def execute_many_json(
    directory: Path,
    steps: list[str | dict],
    confirm: bool = False,
    allow_vision_fallback: bool = False,
    highlight: bool = False,
    highlight_duration_ms: int = 1200,
    highlight_wait_ms: int = 100,
) -> str:
    return json.dumps(
        execute_many(
            directory,
            steps,
            confirm,
            allow_vision_fallback,
            highlight,
            highlight_duration_ms,
            highlight_wait_ms,
        ),
        ensure_ascii=False,
        indent=2,
    )
