"""Diff-driven learning of UI command effects and safe application exploration."""

from __future__ import annotations

import base64
import ctypes
import hashlib
import io
import json
import re
import time
from collections.abc import Callable
from pathlib import Path
from urllib import error as urlerror
from urllib import request as urlrequest

import pyautogui
import uiautomation

from . import knowledge, scanner, ui_cli

BLOCKED_INTENT = re.compile(
    r"(send|publish|post|purchase|pay|delete|remove|upload|logout|log-out|close-window|exit)",
    re.IGNORECASE,
)
SAFE_ACTIONS = {"click", "double-click", "right-click", "hover"}


def _owner_handle(handle: int) -> int:
    if not handle:
        return 0
    try:
        return int(ctypes.windll.user32.GetWindow(handle, 4) or 0)
    except Exception:
        return 0


def snapshot_top_level_windows() -> list[dict]:
    """Return a cheap desktop-level window inventory for popup detection."""
    foreground = int(ctypes.windll.user32.GetForegroundWindow() or 0)
    windows = []
    for z_order, control in enumerate(uiautomation.GetRootControl().GetChildren()):
        try:
            rectangle = scanner._rect(control)
            handle = int(getattr(control, "NativeWindowHandle", 0) or 0)
            if rectangle["width"] <= 1 or rectangle["height"] <= 1:
                continue
            windows.append(
                {
                    "handle": handle,
                    "title": str(getattr(control, "Name", "") or ""),
                    "class_name": str(getattr(control, "ClassName", "") or ""),
                    "control_type": str(getattr(control, "ControlTypeName", "") or ""),
                    "process_id": int(getattr(control, "ProcessId", 0) or 0),
                    "owner_handle": _owner_handle(handle),
                    "rect": rectangle,
                    "z_order": z_order,
                    "foreground": handle == foreground,
                }
            )
        except Exception:
            continue
    return windows


def _capture_virtual_desktop():
    from PIL import ImageGrab

    try:
        image = ImageGrab.grab(all_screens=True)
        origin = {
            "left": int(ctypes.windll.user32.GetSystemMetrics(76)),
            "top": int(ctypes.windll.user32.GetSystemMetrics(77)),
        }
        return image, origin
    except Exception:
        return pyautogui.screenshot(), {"left": 0, "top": 0}


def _image_fingerprint(image) -> str:
    small = image.convert("L").resize((32, 32))
    return hashlib.sha256(bytes(small.getdata())).hexdigest()[:20]


def _control_properties(control) -> dict:
    if not control:
        return {}
    result = {
        "name": str(getattr(control, "Name", "") or ""),
        "automation_id": str(getattr(control, "AutomationId", "") or ""),
        "class_name": str(getattr(control, "ClassName", "") or ""),
        "control_type": str(getattr(control, "ControlTypeName", "") or ""),
        "enabled": bool(getattr(control, "IsEnabled", True)),
        "focusable": bool(getattr(control, "IsKeyboardFocusable", False)),
        "rect": scanner._rect(control),
    }
    for pattern_name, property_name in (
        ("GetTogglePattern", "toggle_state"),
        ("GetSelectionItemPattern", "selected"),
        ("GetValuePattern", "value"),
    ):
        try:
            pattern = getattr(control, pattern_name)()
            if property_name == "toggle_state":
                result[property_name] = int(pattern.ToggleState)
            elif property_name == "selected":
                result[property_name] = bool(pattern.IsSelected)
            else:
                result[property_name] = str(pattern.Value or "")
        except Exception:
            continue
    return result


def capture_snapshot(window_name: str, action_location: dict | None = None) -> dict:
    """Capture full pixels, top-level windows, and the operated control state."""
    window = scanner._find_window(window_name)
    window_rect = scanner._rect(window)
    target_image = scanner._screenshot_window(window_rect)
    desktop_image, desktop_origin = _capture_virtual_desktop()
    action_control = None
    if action_location:
        try:
            action_control = scanner.resolve_location(action_location, window=window)
        except Exception:
            action_control = None
    windows = snapshot_top_level_windows()
    fingerprint_payload = {
        "image": _image_fingerprint(target_image),
        "windows": [
            (item["title"], item["class_name"], item["process_id"], item["rect"])
            for item in windows
        ],
    }
    return {
        "captured_at": knowledge.utc_now(),
        "window_name": str(getattr(window, "Name", "") or window_name),
        "target_handle": int(getattr(window, "NativeWindowHandle", 0) or 0),
        "window_rect": window_rect,
        "windows": windows,
        "desktop_origin": desktop_origin,
        "action_control": _control_properties(action_control),
        "state_id": knowledge.stable_id("state", fingerprint_payload),
        "target_fingerprint": fingerprint_payload["image"],
        "_target_image": target_image,
        "_desktop_image": desktop_image,
    }


def image_change_ratio(before, after) -> float:
    import cv2
    import numpy

    if before.size != after.size:
        return 1.0
    first = cv2.cvtColor(numpy.array(before.convert("RGB")), cv2.COLOR_RGB2GRAY)
    second = cv2.cvtColor(numpy.array(after.convert("RGB")), cv2.COLOR_RGB2GRAY)
    changed = cv2.absdiff(first, second) > 18
    return float(changed.mean())


def _merge_regions(regions: list[dict], gap: int = 12) -> list[dict]:
    merged = []
    for item in sorted(regions, key=lambda value: (value["top"], value["left"])):
        match = None
        for existing in merged:
            if not (
                item["left"] > existing["right"] + gap
                or item["right"] < existing["left"] - gap
                or item["top"] > existing["bottom"] + gap
                or item["bottom"] < existing["top"] - gap
            ):
                match = existing
                break
        if match is None:
            merged.append(dict(item))
            continue
        match["left"] = min(match["left"], item["left"])
        match["top"] = min(match["top"], item["top"])
        match["right"] = max(match["right"], item["right"])
        match["bottom"] = max(match["bottom"], item["bottom"])
        match["width"] = match["right"] - match["left"]
        match["height"] = match["bottom"] - match["top"]
    return merged


def changed_regions(before, after, min_area: int = 24) -> list[dict]:
    """Find merged local pixel-difference boxes while suppressing tiny noise."""
    import cv2
    import numpy

    if before.size != after.size:
        return [
            {
                "left": 0,
                "top": 0,
                "right": after.width,
                "bottom": after.height,
                "width": after.width,
                "height": after.height,
            }
        ]
    first = cv2.cvtColor(numpy.array(before.convert("RGB")), cv2.COLOR_RGB2GRAY)
    second = cv2.cvtColor(numpy.array(after.convert("RGB")), cv2.COLOR_RGB2GRAY)
    mask = (cv2.absdiff(first, second) > 18).astype("uint8") * 255
    kernel = numpy.ones((3, 3), dtype="uint8")
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    contours, _hierarchy = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    result = []
    for contour in contours:
        x, y, width, height = cv2.boundingRect(contour)
        if width * height < min_area:
            continue
        padding = 4
        left = max(0, x - padding)
        top = max(0, y - padding)
        right = min(after.width, x + width + padding)
        bottom = min(after.height, y + height + padding)
        result.append(
            {
                "left": left,
                "top": top,
                "right": right,
                "bottom": bottom,
                "width": right - left,
                "height": bottom - top,
            }
        )
    return _merge_regions(result)


def _window_differences(before: list[dict], after: list[dict]) -> dict:
    before_by_handle = {item["handle"]: item for item in before if item["handle"]}
    after_by_handle = {item["handle"]: item for item in after if item["handle"]}
    return {
        "added": [after_by_handle[key] for key in after_by_handle.keys() - before_by_handle],
        "removed": [before_by_handle[key] for key in before_by_handle.keys() - after_by_handle],
        "changed": [
            {"before": before_by_handle[key], "after": after_by_handle[key]}
            for key in before_by_handle.keys() & after_by_handle
            if before_by_handle[key]["rect"] != after_by_handle[key]["rect"]
            or before_by_handle[key]["title"] != after_by_handle[key]["title"]
        ],
    }


def wait_for_stability(
    window_name: str,
    action_location: dict,
    minimum_seconds: float = 0.2,
    maximum_seconds: float = 3.0,
    interval_seconds: float = 0.1,
) -> tuple[dict, dict]:
    """Observe delayed and transient changes until two consecutive stable frames."""
    started = time.perf_counter()
    previous = None
    stable_frames = 0
    observed_handles = {}
    samples = 0
    last = None
    dynamic_regions = []
    while time.perf_counter() - started < max(minimum_seconds, maximum_seconds):
        last = capture_snapshot(window_name, action_location)
        samples += 1
        for item in last["windows"]:
            if item["handle"]:
                observed_handles[item["handle"]] = item
        if previous is not None:
            ratio = image_change_ratio(previous["_target_image"], last["_target_image"])
            dynamic_regions = changed_regions(
                previous["_target_image"],
                last["_target_image"],
                min_area=12,
            )
            window_same = previous["windows"] == last["windows"]
            stable_frames = stable_frames + 1 if ratio < 0.0005 and window_same else 0
        previous = last
        elapsed = time.perf_counter() - started
        if elapsed >= minimum_seconds and stable_frames >= 2:
            return last, {
                "stable": True,
                "elapsed_seconds": round(elapsed, 3),
                "samples": samples,
                "observed_windows": list(observed_handles.values()),
                "dynamic_regions": [],
            }
        time.sleep(max(0.02, interval_seconds))
    return last, {
        "stable": False,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
        "samples": samples,
        "observed_windows": list(observed_handles.values()),
        "dynamic_regions": dynamic_regions,
        "detail": "visual state remained dynamic until timeout",
    }


def _controls_in_changed_regions(
    window_name: str,
    window_rect: dict,
    regions: list[dict],
) -> list[dict]:
    window = scanner._find_window(window_name)
    controls = []
    seen = set()
    for index, rectangle in enumerate(regions, start=1):
        target = {
            "id": index,
            "role": "control",
            "actions": [],
            "relative_rect": rectangle,
        }
        control = scanner.control_from_visual_target(window, window_rect, target)
        if control is None:
            continue
        try:
            location = scanner.location_from_xpath(scanner.get_control_xpath(control))
            key = json.dumps(location, ensure_ascii=False, sort_keys=True)
        except Exception:
            continue
        if key in seen:
            continue
        seen.add(key)
        controls.append({**_control_properties(control), "location": location})
    return controls


def _popup_controls(window_changes: dict, max_controls: int = 100) -> list[dict]:
    """Inspect only newly added top-level windows, never the whole desktop tree."""
    added = window_changes.get("added", [])
    if not added:
        return []
    roots = {
        int(getattr(item, "NativeWindowHandle", 0) or 0): item
        for item in uiautomation.GetRootControl().GetChildren()
    }
    result = []
    for window_info in added:
        root = roots.get(window_info.get("handle", 0))
        if root is None:
            continue
        controls = []
        for control, depth in scanner._walk_controls(root, 5, max_controls):
            try:
                control_type = str(getattr(control, "ControlTypeName", "") or "")
                name = str(getattr(control, "Name", "") or "")
                automation_id = str(getattr(control, "AutomationId", "") or "")
                actions = scanner._control_actions(
                    control_type,
                    bool(getattr(control, "IsEnabled", True)),
                )
                if not actions and not name and not automation_id:
                    continue
                location = scanner.location_from_xpath(scanner.get_control_xpath(control))
                controls.append(
                    {
                        **_control_properties(control),
                        "depth": depth,
                        "actions": actions,
                        "location": location,
                    }
                )
            except Exception:
                continue
        result.append({"window": window_info, "controls": controls})
    return result


def _image_data_url(image) -> str:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")


def _popup_images(snapshot: dict, window_changes: dict) -> list:
    """Crop newly opened windows from the local virtual-desktop snapshot."""
    desktop = snapshot["_desktop_image"]
    origin = snapshot.get("desktop_origin", {"left": 0, "top": 0})
    images = []
    for item in window_changes.get("added", [])[:3]:
        rectangle = item.get("rect", {})
        left = max(0, int(rectangle.get("left", 0)) - int(origin.get("left", 0)))
        top = max(0, int(rectangle.get("top", 0)) - int(origin.get("top", 0)))
        right = min(
            desktop.width,
            int(rectangle.get("right", 0)) - int(origin.get("left", 0)),
        )
        bottom = min(
            desktop.height,
            int(rectangle.get("bottom", 0)) - int(origin.get("top", 0)),
        )
        if right - left > 1 and bottom - top > 1:
            images.append(desktop.crop((left, top, right, bottom)))
    return images


def _interpret_effects(
    before_image,
    after_image,
    diff_regions: list[dict],
    window_changes: dict,
    changed_controls: list[dict],
    popup_controls: list[dict],
    popup_images: list,
    dynamic_regions: list[dict],
    command: str,
    api_url: str,
    api_key: str,
    model: str,
    version: str,
) -> dict:
    prompt = (
        "Compare these before and after screenshots of one desktop UI operation. Only describe "
        "changes supported by the images and metadata. Generalize dynamic text as placeholders "
        "such as {result}; do not hard-code one observed value as the rule. Return strict JSON: "
        '{"summary":"","effects":[{"type":"property_changed","description":"",'
        '"control_hint":"","property":"Name","value_pattern":""}],'
        '"success_condition":"","response_regions":[1],"confidence":0.95}. '
        f"Command: {command}. Difference boxes: {json.dumps(diff_regions)}. "
        f"Top-level window changes: {json.dumps(window_changes, ensure_ascii=False)}. "
        f"UIA controls at changed regions: {json.dumps(changed_controls, ensure_ascii=False)}. "
        f"Controls in new popup windows: {json.dumps(popup_controls, ensure_ascii=False)}. "
        f"Continuously changing regions to treat as unstable rather than a deterministic effect: "
        f"{json.dumps(dynamic_regions)}. Any additional attached images are crops of newly opened "
        "windows related to this operation."
    )
    payload = {
        "model": model,
        "temperature": 0,
        "stream": True,
        "messages": [
            {
                "role": "system",
                "content": "You analyze observable desktop operation effects as strict JSON.",
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": _image_data_url(before_image)}},
                    {"type": "image_url", "image_url": {"url": _image_data_url(after_image)}},
                    *[
                        {"type": "image_url", "image_url": {"url": _image_data_url(image)}}
                        for image in popup_images
                    ],
                ],
            },
        ],
    }
    request = urlrequest.Request(
        api_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": f"easy-uiauto/{version}",
        },
        method="POST",
    )
    try:
        with urlrequest.urlopen(request, timeout=60) as response:
            content = scanner._read_completion_content(response)
        result = scanner._parse_json_object(str(content))
    except (OSError, KeyError, IndexError, TypeError, ValueError, urlerror.URLError) as error:
        return {
            "summary": f"Effect interpretation unavailable: {type(error).__name__}: {error}",
            "effects": [],
            "success_condition": "",
            "confidence": 0.0,
        }
    effects = result.get("effects", [])
    return {
        "summary": str(result.get("summary") or "").strip()[:500],
        "effects": [item for item in effects[:30] if isinstance(item, dict)]
        if isinstance(effects, list)
        else [],
        "success_condition": str(result.get("success_condition") or "").strip()[:500],
        "response_regions": result.get("response_regions", []),
        "confidence": max(0.0, min(1.0, float(result.get("confidence", 0) or 0))),
    }


def _public_snapshot(snapshot: dict) -> dict:
    return {key: value for key, value in snapshot.items() if not key.startswith("_")}


def _save_interaction_images(
    directory: Path,
    interaction_id: str,
    before: dict,
    after: dict,
) -> dict:
    image_dir = directory / "images" / "interactions" / interaction_id
    image_dir.mkdir(parents=True, exist_ok=True)
    paths = {}
    for name, image in (
        ("before", before["_target_image"]),
        ("after", after["_target_image"]),
        ("desktop-before", before["_desktop_image"]),
        ("desktop-after", after["_desktop_image"]),
    ):
        path = image_dir / f"{name}.png"
        image.save(path)
        paths[name.replace("-", "_")] = str(path.relative_to(directory)).replace("\\", "/")
    return paths


def _attempt_recovery(
    window_name: str,
    before: dict,
    action_location: dict,
    action: str,
) -> dict:
    attempts = []
    methods = [("escape", lambda: pyautogui.press("esc"))]
    if action == "hover":
        rectangle = before["window_rect"]
        methods.insert(
            0,
            (
                "move-pointer",
                lambda: pyautogui.moveTo(
                    rectangle["left"] + rectangle["width"] // 2,
                    rectangle["top"] + 8,
                    duration=0,
                ),
            ),
        )
    methods.append(("browser-back", lambda: pyautogui.hotkey("alt", "left")))
    for method, operation in methods:
        operation()
        time.sleep(0.15)
        try:
            recovered, stability = wait_for_stability(
                window_name,
                action_location,
                minimum_seconds=0.1,
                maximum_seconds=1.5,
            )
        except Exception as error:
            attempts.append({"method": method, "error": str(error)})
            continue
        ratio = image_change_ratio(before["_target_image"], recovered["_target_image"])
        attempts.append({"method": method, "image_change_ratio": round(ratio, 6)})
        if ratio < 0.01:
            return {
                "ok": True,
                "method": method,
                "image_change_ratio": round(ratio, 6),
                "stability": stability,
                "attempts": attempts,
                "detail": "returned to the previous visual state",
            }
    return {
        "ok": False,
        "method": attempts[-1].get("method", "none") if attempts else "none",
        "attempts": attempts,
        "detail": "state did not recover",
    }


def _navigation_likely(context: dict) -> bool:
    if context["window_changes"].get("added"):
        return True
    navigation_terms = ("page", "navigate", "navigation", "dialog", "window", "view")
    return any(
        any(term in str(effect.get("type", "")).casefold() for term in navigation_terms)
        for effect in context.get("effects", [])
        if isinstance(effect, dict)
    )


def learn_command_effect(
    directory: Path,
    command: str,
    api_url: str,
    api_key: str,
    model: str,
    version: str,
    text: str = "",
    confirm: bool = False,
    recover: bool = False,
    maximum_wait_seconds: float = 3.0,
    post_effect: Callable[[dict], dict] | None = None,
    progress: Callable[[str], None] | None = None,
) -> dict:
    """Execute one verified command and persist its observed before/after effect."""
    progress = progress or (lambda _message: None)
    record, _action = knowledge.resolve_command(directory, command)
    if record.get("requires_confirmation") and not confirm:
        raise RuntimeError(f"Command requires explicit confirmation: {command}")
    interaction_id = knowledge.stable_id(
        "interaction",
        record["app_id"],
        command,
        knowledge.utc_now(),
    )
    window = scanner._find_window(record["app_name"])
    scanner._activate_window(window)
    time.sleep(0.05)
    progress(f"Capturing before state for {command}")
    before = capture_snapshot(record["app_name"], record["location"])
    progress(f"Executing reversible command {command}")
    execution = ui_cli.execute(directory, command, text, confirm)
    progress(f"Waiting for UI stability after {command}")
    after, stability = wait_for_stability(
        record["app_name"],
        record["location"],
        maximum_seconds=max(0.5, min(float(maximum_wait_seconds), 15.0)),
    )
    progress(f"Inspecting changed regions after {command}")
    regions = changed_regions(before["_target_image"], after["_target_image"])
    window_changes = _window_differences(before["windows"], after["windows"])
    changed_controls = _controls_in_changed_regions(
        after["window_name"],
        after["window_rect"],
        regions,
    )
    popup_controls = _popup_controls(window_changes)
    popup_images = _popup_images(after, window_changes)
    action_property_changes = {
        key: {
            "before": before["action_control"].get(key),
            "after": after["action_control"].get(key),
        }
        for key in set(before["action_control"]) | set(after["action_control"])
        if before["action_control"].get(key) != after["action_control"].get(key)
    }
    progress(f"Understanding operation effect for {command} with remote AI")
    interpretation = _interpret_effects(
        before["_target_image"],
        after["_target_image"],
        regions,
        window_changes,
        changed_controls,
        popup_controls,
        popup_images,
        stability.get("dynamic_regions", []),
        command,
        api_url,
        api_key,
        model,
        version,
    )
    change_ratio = image_change_ratio(before["_target_image"], after["_target_image"])
    effect_context = {
        "command": command,
        "page_id": record.get("page_id", "unknown"),
        "before_state_id": before["state_id"],
        "after_state_id": after["state_id"],
        "window_name": after["window_name"],
        "target_handle": after.get("target_handle", 0),
        "window_changes": window_changes,
        "effects": interpretation["effects"],
        "image_change_ratio": round(change_ratio, 6),
    }
    discovery = {"attempted": False}
    if post_effect and _navigation_likely(effect_context):
        try:
            discovery = {"attempted": True, **post_effect(effect_context)}
        except Exception as error:
            discovery = {
                "attempted": True,
                "ok": False,
                "error": f"{type(error).__name__}: {error}",
            }
    images = _save_interaction_images(directory, interaction_id, before, after)
    changed = bool(regions or any(window_changes.values()) or action_property_changes)
    observed_handles = {
        item.get("handle")
        for item in stability.get("observed_windows", [])
        if item.get("handle")
    }
    before_handles = {item["handle"] for item in before["windows"] if item["handle"]}
    after_handles = {item["handle"] for item in after["windows"] if item["handle"]}
    transient_handles = observed_handles - before_handles - after_handles
    transient_windows = [
        item
        for item in stability.get("observed_windows", [])
        if item.get("handle") in transient_handles
    ]
    foreground_after = next(
        (item.get("handle") for item in after["windows"] if item.get("foreground")),
        0,
    )
    allowed_foreground = {after.get("target_handle", 0)} | {
        item.get("handle", 0) for item in window_changes.get("added", [])
    }
    interference = {
        "detected": bool(foreground_after and foreground_after not in allowed_foreground),
        "detail": (
            "foreground moved to an unrelated window during observation"
            if foreground_after and foreground_after not in allowed_foreground
            else ""
        ),
    }
    if recover:
        progress(f"Recovering previous UI state after {command}")
    recovery = _attempt_recovery(
        record["app_name"],
        before,
        record["location"],
        _action,
    ) if recover else {
        "ok": None,
        "detail": "not requested",
    }
    interaction = {
        "id": interaction_id,
        "app_id": record["app_id"],
        "app_name": record["app_name"],
        "page_id": record.get("page_id", "unknown"),
        "semantic_name": record.get("semantic_name") or command,
        "command": command,
        "control_id": record["id"],
        "status": (
            "interference-detected"
            if interference["detected"]
            else "effect-observed"
            if changed
            else "no-observable-change"
        ),
        "before_state_id": before["state_id"],
        "after_state_id": after["state_id"],
        "before": _public_snapshot(before),
        "after": _public_snapshot(after),
        "images": images,
        "changed_regions": regions,
        "image_change_ratio": round(change_ratio, 6),
        "window_changes": window_changes,
        "transient_windows": transient_windows,
        "changed_controls": changed_controls,
        "popup_controls": popup_controls,
        "action_property_changes": action_property_changes,
        "effects": interpretation["effects"],
        "effect_summary": interpretation["summary"],
        "effect_confidence": interpretation["confidence"],
        "success_condition": interpretation["success_condition"],
        "stability": stability,
        "execution": execution,
        "recovery": recovery,
        "discovery": discovery,
        "interference": interference,
        "model": model,
    }
    progress(f"Saving interaction knowledge for {command}")
    path = knowledge.save_interaction(directory, interaction)
    found = knowledge.find_control_record(directory, record["id"])
    if found:
        _path, updated = found
        updated["function_verification"] = {
            **updated.get("function_verification", {}),
            "status": (
                "interference-detected"
                if interference["detected"]
                else "effect-observed"
                if changed
                else "executed-no-visible-effect"
            ),
            "effect_record": str(path.relative_to(directory)).replace("\\", "/"),
            "last_effect_observed_at": knowledge.utc_now(),
        }
        knowledge.save_control(directory, updated)
    knowledge.rebuild_index(directory)
    return {
        "ok": True,
        "interaction_id": interaction_id,
        "command": command,
        "status": interaction["status"],
        "before_state_id": before["state_id"],
        "after_state_id": after["state_id"],
        "changed_regions": regions,
        "image_change_ratio": round(change_ratio, 6),
        "window_changes": window_changes,
        "changed_controls": changed_controls,
        "popup_controls": popup_controls,
        "effects": interpretation["effects"],
        "success_condition": interpretation["success_condition"],
        "stability": stability,
        "recovery": recovery,
        "discovery": discovery,
        "interference": interference,
        "record": str(path),
    }


def explore_application(
    directory: Path,
    api_url: str,
    api_key: str,
    model: str,
    version: str,
    policy: str = "safe",
    max_actions: int = 10,
    confirm: bool = False,
    max_depth: int = 3,
    progress: Callable[[str], None] | None = None,
) -> dict:
    """Recursively explore known low-risk commands and learn their operation effects."""
    progress = progress or (lambda _message: None)
    policy = policy.strip().lower()
    if policy not in {"safe", "supervised"}:
        raise ValueError("policy must be safe or supervised")
    pending = []
    learned_pairs = {
        (item.get("page_id"), item.get("command"))
        for item in knowledge.list_interactions(directory, limit=500)
        if item.get("status") == "effect-observed"
    }
    results = []
    stopped = []
    visited = set()
    action_limit = max(1, min(int(max_actions), 100))
    depth_limit = max(1, min(int(max_depth), 8))

    def candidates_for_page(page_id: str) -> list[dict]:
        candidates = []
        for item in knowledge.available_commands(directory, page_id):
            if item["action"] not in SAFE_ACTIONS or BLOCKED_INTENT.search(
                item.get("intent", "")
            ):
                continue
            if item.get("requires_confirmation"):
                pending.append(item["command"])
                continue
            if policy == "safe" and item.get("risk") != "safe":
                continue
            if policy == "supervised" and item.get("risk") not in {
                "safe",
                "state-changing",
            }:
                pending.append(item["command"])
                continue
            pair = (item.get("page"), item.get("command"))
            if pair in learned_pairs or pair in visited:
                continue
            candidates.append(item)
        return candidates

    index = knowledge.load_index(directory)
    app_name = index.get("application", {}).get("name", "")
    if not app_name:
        raise RuntimeError("application knowledge has no window name")
    progress("Scanning current page before interactive exploration")
    initial_scan = scanner.scan_window(
        app_name,
        api_url,
        api_key,
        model,
        version,
        strategy="visual-first",
        root=directory.parent.parent,
        progress=lambda message: progress(f"Initial scan: {message}"),
    )
    if initial_scan.get("app_id") != directory.name:
        raise RuntimeError("visible window does not match the selected application knowledge")
    initial_pages = [initial_scan["page_id"]]
    considered = sum(len(candidates_for_page(page)) for page in initial_pages)
    planned = min(considered, action_limit)
    progress(f"Exploring actions 0/{planned}")

    def explore_page(page_id: str, depth: int) -> list[dict]:
        nested_results = []
        if depth > depth_limit or len(results) >= action_limit or stopped:
            return nested_results
        for item in candidates_for_page(page_id):
            if len(results) >= action_limit or stopped:
                break
            pair = (item.get("page"), item.get("command"))
            visited.add(pair)
            progress(
                f"Exploring action {len(results) + 1}/{planned}: {item['command']}"
            )

            def discover(context: dict) -> dict:
                if depth >= depth_limit:
                    return {"ok": True, "depth_limit_reached": True, "scans": []}
                titles = []
                for window_info in context["window_changes"].get("added", []):
                    if window_info.get("title"):
                        titles.append(window_info["title"])
                if not titles:
                    titles.append(context["window_name"])
                scans = []
                for title in dict.fromkeys(titles):
                    scan = scanner.scan_window(
                        title,
                        api_url,
                        api_key,
                        model,
                        version,
                        strategy="visual-first",
                        root=directory.parent.parent,
                        progress=lambda message, title=title: progress(
                            f"Nested page scan ({title}): {message}"
                        ),
                    )
                    if scan.get("app_id") != directory.name:
                        continue
                    scans.append(
                        {
                            key: scan.get(key)
                            for key in (
                                "app_id",
                                "page_id",
                                "strategy",
                                "visual_targets",
                                "controls_saved_this_scan",
                                "commands",
                            )
                        }
                    )
                    nested_results.extend(explore_page(scan["page_id"], depth + 1))
                return {"ok": True, "scans": scans, "nested_actions": len(nested_results)}

            try:
                learned = learn_command_effect(
                    directory,
                    item["command"],
                    api_url,
                    api_key,
                    model,
                    version,
                    confirm=confirm,
                    recover=True,
                    post_effect=discover,
                    progress=progress,
                )
            except Exception as error:
                stopped.append(f"{item['command']}: {type(error).__name__}: {error}")
                break
            results.append(learned)
            nested_results.append(learned)
            if learned["recovery"].get("ok") is False:
                stopped.append(f"recovery failed after {item['command']}")
                break
        return nested_results

    for page_id in initial_pages:
        explore_page(page_id, 1)
        if len(results) >= action_limit or stopped:
            break
    return {
        "ok": not stopped,
        "policy": policy,
        "max_depth": depth_limit,
        "initial_scan": {
            key: initial_scan.get(key)
            for key in (
                "app_id",
                "page_id",
                "strategy",
                "visual_targets",
                "controls_saved_this_scan",
                "commands",
            )
        },
        "actions_considered": considered,
        "actions_learned": len(results),
        "interactions": results,
        "pending_confirmation": list(dict.fromkeys(pending)),
        "stopped": stopped[0] if stopped else "",
    }
