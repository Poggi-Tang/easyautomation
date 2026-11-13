"""Scan a Windows application into an Obsidian-compatible UI knowledge vault."""

from __future__ import annotations

import base64
import ctypes
import hashlib
import io
import json
import math
import re
import time
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor, as_completed
from ctypes import wintypes
from pathlib import Path
from urllib import error as urlerror
from urllib import request as urlrequest
from uuid import uuid4

import pyautogui
import uiautomation

from easy_uiauto.utils import get_control_xpath
from easy_uiauto.ctrl import Controller

from . import (
    configuration,
    image_storage,
    knowledge,
    native_windows,
    page_identity,
    semantic_identity,
    visualization,
    window_identity,
)
from .protocol import location_from_xpath, normalize_location, stable_location, with_match_mode

ACTIONABLE_TYPES = {
    "ButtonControl": ["click", "hover"],
    "CheckBoxControl": ["click", "hover"],
    "ComboBoxControl": ["click", "hover"],
    "EditControl": ["click", "set-text", "right-click", "hover"],
    "HyperlinkControl": ["click", "right-click", "hover"],
    "ListItemControl": ["click", "double-click", "right-click", "hover"],
    "MenuItemControl": ["click", "hover"],
    "RadioButtonControl": ["click", "hover"],
    "TabItemControl": ["click", "right-click", "hover"],
    "TreeItemControl": ["click", "double-click", "right-click", "hover"],
}

SEMANTIC_CONFIDENCE_THRESHOLD = 0.72
SEMANTIC_BATCH_SIZE = 60
SEMANTIC_RISKS = {"safe", "state-changing", "external", "destructive", "unknown"}
SCAN_STRATEGIES = {"visual-first", "local-visual", "full-uia"}
VISUAL_RESOLUTION_ERRORS = {
    "no_uia_candidate": "UIA point lookup returned no candidate controls",
    "uia_lookup_failed": "UIA lookup failed before a readable candidate was collected",
    "geometry_rejected": "UIA candidates were found but rejected by geometry filters",
    "semantic_rejected": "Geometrically matching UIA candidates were rejected by semantic checks",
    "score_rejected": "UIA candidates passed geometry and semantics but scored below the minimum",
    "resolution_failed": "UIA target resolution failed without diagnostic details",
}
GENERIC_CONTAINER_TYPES = {
    "CustomControl",
    "GroupControl",
    "PaneControl",
    "WindowControl",
}


def _rect(control) -> dict:
    rectangle = control.BoundingRectangle
    left = int(getattr(rectangle, "left", 0))
    top = int(getattr(rectangle, "top", 0))
    right = int(getattr(rectangle, "right", 0))
    bottom = int(getattr(rectangle, "bottom", 0))
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "width": max(0, right - left),
        "height": max(0, bottom - top),
    }


def _window_rect(window, bounds: dict | None = None) -> dict:
    """Use the same visible frame for capture, normalized hints and coordinate mapping."""
    handle = int(getattr(window, "NativeWindowHandle", 0) or 0)
    if bounds is None and handle:
        # Some Qt providers omit a frame pixel from UIA BoundingRectangle.
        # Capture and GUI picking must both use the current physical HWND frame.
        try:
            user, _callback = native_windows._api()
            rectangle = wintypes.RECT()
            if user.IsWindow(handle) and user.GetWindowRect(handle, ctypes.byref(rectangle)):
                bounds = dict(left=rectangle.left, top=rectangle.top, right=rectangle.right,
                              bottom=rectangle.bottom, width=rectangle.right-rectangle.left,
                              height=rectangle.bottom-rectangle.top)
        except (OSError, AttributeError, TypeError):
            pass
    return native_windows.visible_window_bounds(
        handle,
        _rect(window) if bounds is None else bounds,
        class_name=str(getattr(window, "ClassName", "") or ""),
    )


def _trim_window_capture(image, bounds: dict, window):
    """Keep image pixels and screen origin together after optional local edge cropping."""
    if (not hasattr(image, "size") or str(getattr(window, "ClassName", "")) in {
        "Progman", "WorkerW", "Shell_TrayWnd", "Shell_SecondaryTrayWnd",
    }):
        return image, bounds
    from .window_crop import crop_window_image

    try:
        cropped = crop_window_image(image, bounds)
    except (ImportError, RuntimeError):
        # The local visual dependency is optional; native visible bounds still work.
        return image, bounds
    return cropped["image"], cropped["screen_rect"]


def _valid_rect(rectangle: dict, window_rect: dict) -> bool:
    return (
        rectangle["width"] > 1
        and rectangle["height"] > 1
        and rectangle["right"] > window_rect["left"]
        and rectangle["bottom"] > window_rect["top"]
        and rectangle["left"] < window_rect["right"]
        and rectangle["top"] < window_rect["bottom"]
    )


def _clip_rect(rectangle: dict, window_rect: dict) -> dict:
    left = max(rectangle["left"], window_rect["left"])
    top = max(rectangle["top"], window_rect["top"])
    right = min(rectangle["right"], window_rect["right"])
    bottom = min(rectangle["bottom"], window_rect["bottom"])
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "width": max(0, right - left),
        "height": max(0, bottom - top),
    }


def _find_window(title: str | dict):
    return window_identity.resolve(title)


def _activate_window(window) -> None:
    """Activate a UIA window without relying on pygetwindow's stale handles."""
    handle = int(getattr(window, "NativeWindowHandle", 0) or 0)
    user32, _callback = native_windows._api()
    if not handle or not user32.IsWindow(handle):
        raise RuntimeError("UI Automation returned an invalid window handle")
    Controller.activate_window_handle(handle, control=window)


def _process_name(process_id: int) -> str:
    if not process_id:
        return ""
    process_query_limited_information = 0x1000
    handle = ctypes.windll.kernel32.OpenProcess(
        process_query_limited_information,
        False,
        process_id,
    )
    if not handle:
        return ""
    try:
        capacity = wintypes.DWORD(32768)
        buffer = ctypes.create_unicode_buffer(capacity.value)
        if ctypes.windll.kernel32.QueryFullProcessImageNameW(
            handle,
            0,
            buffer,
            ctypes.byref(capacity),
        ):
            return Path(buffer.value).stem
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)
    return ""


def _walk_controls(window, max_depth: int, max_controls: int) -> list[tuple[object, int]]:
    collected = []
    stack = [(window, 0)]
    seen = set()
    while stack and len(collected) < max_controls:
        control, depth = stack.pop()
        identity = (
            getattr(control, "NativeWindowHandle", 0),
            getattr(control, "ControlType", 0),
            tuple(_rect(control).values()),
            getattr(control, "Name", ""),
        )
        if identity in seen:
            continue
        seen.add(identity)
        collected.append((control, depth))
        if depth >= max_depth:
            continue
        try:
            children = control.GetChildren()
        except Exception:
            children = []
        stack.extend((child, depth + 1) for child in reversed(children))
    return collected


def _matches_step(control, step: dict) -> bool:
    expected = {
        "ControlType": getattr(control, "ControlTypeName", "") or "",
        "Name": getattr(control, "Name", "") or "",
        "ClassName": getattr(control, "ClassName", "") or "",
        "AutomationId": getattr(control, "AutomationId", "") or "",
    }
    return all(not step.get(key) or step[key] == value for key, value in expected.items())










def resolve_location(location: dict, window=None, prefix_cache: dict | None = None):
    """Adapt shell contexts to the shared strict, side-effect-free resolver."""
    from easy_uiauto.utils import observe_control
    from .results import ActionError

    try:
        normalized = normalize_location(location)
    except (TypeError, ValueError, AttributeError) as error:
        raise ActionError("invalid_argument", str(error)) from error
    xpath = normalized["Xpath"]
    if window is None:
        window_name = normalized.get("WindowName") or (xpath[0].get("Name", "") if xpath else "")
        if not window_name:
            return observe_control(normalized, prefix_cache=prefix_cache)
        try:
            # Bind through the shared root selector, including class/ID/index.
            # Keep the title-only adapter for partial-title scanning and shell discovery.
            from easy_uiauto.utils import _native_window_scope
            first = xpath[0] if xpath and xpath[0].get("ControlType") == "WindowControl" else {}
            try:
                window = _native_window_scope(window_name, first)
            except ActionError as error:
                if (error.code != "window_invalid" or error.reason
                        or any(first.get(key) for key in ("Name", "ClassName", "AutomationId", "foundIndex"))):
                    raise
                window = _find_window(window_name)
        except ActionError:
            raise
        except Exception as error:
            raise ActionError("action_failed", f"UIA window lookup failed: {error}") from error
    if xpath and _is_desktop_window(window) and not _matches_step(window, xpath[0]):
        matches = {}
        for shell in _desktop_shell_roots(window):
            if not _matches_step(shell, xpath[0]):
                continue
            try:
                candidate = observe_control(normalized, window=shell)
            except ActionError as error:
                if error.code != "control_not_found" or error.reason == "index_out_of_range":
                    raise
                continue
            matches[_control_key(candidate)] = candidate
        if len(matches) > 1:
            raise ActionError("ambiguous_match", "Multiple shell controls match the location")
        if not matches:
            raise ActionError("control_not_found", "Shell control not found", retryable=True)
        return next(iter(matches.values()))
    return observe_control(normalized, window=window, prefix_cache=prefix_cache)


def _screenshot_window(window_rect: dict):
    with Controller.suspend_highlights():
        return _screenshot_window_unmarked(window_rect)


def _screenshot_window_unmarked(window_rect: dict):
    from PIL import ImageGrab

    bounds = (
        window_rect["left"],
        window_rect["top"],
        window_rect["right"],
        window_rect["bottom"],
    )
    try:
        return ImageGrab.grab(bbox=bounds, all_screens=True)
    except Exception:
        return pyautogui.screenshot(
            region=(
                window_rect["left"],
                window_rect["top"],
                window_rect["width"],
                window_rect["height"],
            )
        )


def _image_sha256(image) -> str:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return hashlib.sha256(buffer.getvalue()).hexdigest()


def _image_similarity(first, second) -> float:
    import cv2
    import numpy

    first_array = cv2.cvtColor(numpy.array(first.convert("RGB")), cv2.COLOR_RGB2GRAY)
    second_array = cv2.cvtColor(numpy.array(second.convert("RGB")), cv2.COLOR_RGB2GRAY)
    first_array = cv2.resize(first_array, (64, 64), interpolation=cv2.INTER_AREA)
    second_array = cv2.resize(second_array, (64, 64), interpolation=cv2.INTER_AREA)
    difference = numpy.mean(numpy.abs(first_array.astype("float32") - second_array))
    return round(max(0.0, 1.0 - float(difference) / 255.0), 4)


def _reuse_page_identity(directory: Path, screenshot, page: dict) -> dict:
    return page_identity.select(directory, screenshot, page)


def _rect_iou(first: dict, second: dict) -> float:
    intersection_width = max(
        0,
        min(first["right"], second["right"]) - max(first["left"], second["left"]),
    )
    intersection_height = max(
        0,
        min(first["bottom"], second["bottom"]) - max(first["top"], second["top"]),
    )
    intersection = intersection_width * intersection_height
    first_area = (first["right"] - first["left"]) * (first["bottom"] - first["top"])
    second_area = (second["right"] - second["left"]) * (second["bottom"] - second["top"])
    union = first_area + second_area - intersection
    return intersection / union if union else 0.0


def _reuse_region_identities(
    directory: Path,
    page_id: str,
    regions: list[dict],
    window_rect: dict,
) -> list[dict]:
    previous_regions = [
        record
        for _path, record in knowledge.iter_records(directory, "region")
        if record.get("page_id") == page_id
    ]
    reused = []
    claimed = set()
    for region in regions:
        current_rect = knowledge.normalized_rect(region["rect"], window_rect, relative=True)
        candidates = []
        for previous in previous_regions:
            previous_id = str(previous.get("id", "")).split(".", 1)[-1]
            if previous_id in claimed:
                continue
            previous_rect = previous.get("normalized_rect")
            if not isinstance(previous_rect, dict) and isinstance(previous.get("rect"), dict):
                previous_rect = knowledge.normalized_rect(
                    previous["rect"], window_rect, relative=True
                )
            if not isinstance(previous_rect, dict):
                continue
            candidates.append((_rect_iou(current_rect, previous_rect), previous_id))
        candidates.sort(reverse=True)
        if candidates and candidates[0][0] >= 0.6:
            region = {**region, "id": candidates[0][1]}
            claimed.add(candidates[0][1])
        reused.append(region)
    return reused


def _template_verification(template, screenshot, expected_rect: dict, window_rect: dict) -> dict:
    result = locate_template(template, screenshot, window_rect)
    if not result["found"]:
        return {
            "ok": False,
            "score": result["score"],
            "second_score": result["second_score"],
            "detail": result["detail"],
        }
    location = result["location"]
    expected_x = expected_rect["left"] - window_rect["left"]
    expected_y = expected_rect["top"] - window_rect["top"]
    distance = ((location[0] - expected_x) ** 2 + (location[1] - expected_y) ** 2) ** 0.5
    tolerance = max(6.0, min(template.width, template.height) * 0.15)
    result["ok"] = distance <= tolerance
    result["detail"] = (
        f"best match distance={distance:.1f}px, tolerance={tolerance:.1f}px, "
        f"uniqueness={result['uniqueness']:.3f}"
    )
    return result


def locate_template(template, screenshot, window_rect: dict) -> dict:
    """Find one unique control image in a current window screenshot."""
    import cv2
    import numpy

    if template.width < 8 or template.height < 8:
        return {
            "found": False,
            "score": 0.0,
            "second_score": 0.0,
            "detail": "control image is too small",
        }
    if template.width >= screenshot.width or template.height >= screenshot.height:
        return {
            "found": False,
            "score": 1.0,
            "second_score": 1.0,
            "detail": "control fills the scan area and is not a usable template",
        }
    template_array = cv2.cvtColor(numpy.array(template.convert("RGB")), cv2.COLOR_RGB2GRAY)
    screenshot_array = cv2.cvtColor(
        numpy.array(screenshot.convert("RGB")),
        cv2.COLOR_RGB2GRAY,
    )
    if float(template_array.std()) < 2.0:
        return {
            "found": False,
            "score": 0.0,
            "second_score": 0.0,
            "detail": "control image has no visual detail",
        }
    matches = cv2.matchTemplate(screenshot_array, template_array, cv2.TM_CCOEFF_NORMED)
    _minimum, score, _minimum_location, location = cv2.minMaxLoc(matches)
    suppressed = matches.copy()
    margin_x = max(2, template.width // 2)
    margin_y = max(2, template.height // 2)
    left = max(0, location[0] - margin_x)
    top = max(0, location[1] - margin_y)
    right = min(suppressed.shape[1], location[0] + margin_x + 1)
    bottom = min(suppressed.shape[0], location[1] + margin_y + 1)
    suppressed[top:bottom, left:right] = -1.0
    _second_minimum, second_score, _second_location, _second_maximum = cv2.minMaxLoc(
        suppressed
    )
    uniqueness = float(score - second_score)
    found = score >= 0.75 and uniqueness >= 0.02
    absolute_rect = {
        "left": window_rect["left"] + location[0],
        "top": window_rect["top"] + location[1],
        "right": window_rect["left"] + location[0] + template.width,
        "bottom": window_rect["top"] + location[1] + template.height,
        "width": template.width,
        "height": template.height,
    }
    return {
        "found": found,
        "ok": found,
        "score": round(float(score), 4),
        "second_score": round(float(second_score), 4),
        "uniqueness": uniqueness,
        "location": location,
        "rect": absolute_rect,
        "detail": f"score={score:.3f}, uniqueness={uniqueness:.3f}",
    }


def _crop_control(screenshot, control_rect: dict, window_rect: dict):
    clipped = _clip_rect(control_rect, window_rect)
    box = (
        clipped["left"] - window_rect["left"],
        clipped["top"] - window_rect["top"],
        clipped["right"] - window_rect["left"],
        clipped["bottom"] - window_rect["top"],
    )
    return screenshot.crop(box)


def _save_control_images(
    directory: Path,
    control_id: str,
    crop,
    previous: dict,
) -> tuple[Path, list[str], str]:
    """Preserve visual-state templates instead of replacing the only reference."""
    image_path = directory / "images" / "controls" / f"{control_id}.png"
    variant_dir = directory / "images" / "controls" / control_id
    variant_dir.mkdir(parents=True, exist_ok=True)
    variants = list(previous.get("image_variants", []))
    old_relative = previous.get("image", "")
    old_hash = previous.get("image_sha256", "")
    old_path = directory / old_relative if old_relative else None
    if old_path and old_hash and old_path.is_file():
        old_variant = variant_dir / f"{old_hash}.png"
        if not old_variant.is_file():
            knowledge.replace_image(old_path, old_variant)
        variants.append(str(old_variant.relative_to(directory)).replace("\\", "/"))
    image_hash = _image_sha256(crop)
    crop.save(image_path)
    new_variant = variant_dir / f"{image_hash}.png"
    if not new_variant.is_file():
        crop.save(new_variant)
    newest = str(new_variant.relative_to(directory)).replace("\\", "/")
    variants = [value for value in dict.fromkeys(variants) if value != newest] + [newest]
    return image_path, variants[-image_storage.control_variant_limit():], image_hash


def _parse_json_object(content: str) -> dict:
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content, flags=re.IGNORECASE)
    try:
        result = json.loads(content)
    except json.JSONDecodeError as error:
        raise RuntimeError("Vision response did not contain a valid JSON object") from error
    if not isinstance(result, dict):
        raise RuntimeError("AI segmentation response must be a JSON object")
    return result


def _read_completion_content(response) -> str:
    """Read either an OpenAI JSON response or an SSE streaming completion."""
    content_type = ""
    try:
        content_type = response.headers.get_content_type()
    except (AttributeError, TypeError):
        pass
    if content_type == "text/event-stream":
        return _read_sse_content(response)
    raw = response.read().decode("utf-8")
    if not raw.lstrip().startswith("data:"):
        body = json.loads(raw)
        if body.get("error"):
            raise RuntimeError("Vision API returned an error response")
        choices = body.get("choices") or []
        if not choices:
            raise RuntimeError("Vision response did not contain a completion")
        if choices[0].get("finish_reason") == "length":
            raise RuntimeError("Vision completion was truncated by the provider")
        content = choices[0].get("message", {}).get("content")
        if isinstance(content, list):
            content = "".join(
                item.get("text", "") for item in content if isinstance(item, dict)
            )
        if not isinstance(content, str) or not content.strip():
            raise RuntimeError("Vision response contained no completion content")
        return content
    return _parse_sse_lines(raw.splitlines())


def _read_sse_content(response) -> str:
    return _parse_sse_lines(
        raw_line.decode("utf-8", errors="replace") for raw_line in response
    )


def _parse_sse_lines(lines) -> str:
    pieces = []
    for line in lines:
        line = line.strip()
        if not line.startswith("data:"):
            continue
        data = line[5:].strip()
        if not data or data == "[DONE]":
            continue
        chunk = json.loads(data)
        if chunk.get("error"):
            raise RuntimeError("Vision API returned a streaming error response")
        choices = chunk.get("choices", [])
        if not choices:
            continue
        choice = choices[0]
        if choice.get("finish_reason") == "length":
            raise RuntimeError("Vision completion was truncated by the provider")
        content = choice.get("delta", {}).get("content")
        if content is None:
            content = choice.get("message", {}).get("content")
        if isinstance(content, str):
            pieces.append(content)
        elif isinstance(content, list):
            pieces.extend(
                item.get("text", "") for item in content if isinstance(item, dict)
            )
    if not pieces:
        raise RuntimeError("AI streaming response did not contain completion content")
    return "".join(pieces)


def _tile_origins(length: int, size: int) -> list[int]:
    """Cover an axis with evenly spaced tiles and at least 80 pixels of overlap."""
    if length <= size:
        return [0]
    count = math.ceil((length - 80) / (size - 80))
    return [round(i * (length - size) / (count - 1)) for i in range(count)]


def _merge_tile_target(controls: list[dict], target: dict) -> None:
    """Fuse repeated pixel observations while retaining conflicting caption evidence."""
    from . import diagram_geometry

    def same_entity(existing):
        if diagram_geometry.is_text(existing) != diagram_geometry.is_text(target):
            return False
        attachments = [any(rel.get("type") == "attached-to"
                           for rel in item.get("relationships", []))
                       for item in (existing, target)]
        if any(attachments) and not all(attachments):
            return False
        if _same_visual_target(existing, target):
            return True
        if set(existing.get("source_tiles", [])) & set(target.get("source_tiles", [])):
            return False
        # Roles, captions and classes are model predictions, not pixel identity.
        # In particular, ui-element/readonly often replaces subsystem/entity in
        # a second tile. Do not let that prevent strong geometric deduplication.
        a, b = existing["relative_rect"], target["relative_rect"]
        if _rect_iou(a, b) >= 0.82:
            return True
        first_name = str(existing.get("semantic_name", "")).casefold()
        second_name = str(target.get("semantic_name", "")).casefold()
        if not first_name or not second_name or not (
                first_name in second_name or second_name in first_name):
            return False
        if _rect_iou(a, b) >= 0.85:
            return True
        overlap = (max(0, min(a["right"], b["right"]) - max(a["left"], b["left"]))
                   * max(0, min(a["bottom"], b["bottom"]) - max(a["top"], b["top"])))
        smaller = min(a["width"] * a["height"], b["width"] * b["height"])
        return (bool(existing.get("_tile_clipped")) != bool(target.get("_tile_clipped"))
                and overlap / max(1, smaller) >= 0.75)

    position = next((i for i, existing in enumerate(controls) if same_entity(existing)), None)
    if position is None:
        controls.append({**target, "id": len(controls) + 1})
        return
    existing = controls[position]
    old_clipped = bool(existing.get("_tile_clipped"))
    new_clipped = bool(target.get("_tile_clipped"))
    prefer_new = ((old_clipped and not new_clipped)
                  or (old_clipped == new_clipped
                      and target.get("confidence", 0) > existing.get("confidence", 0)))
    merged = dict(target if prefer_new else existing)
    # Keep a real observation, rather than unioning disagreeing boxes into a
    # larger rectangle that no tile predicted. Preserve all competing evidence.
    observations = []
    for item in (existing, target):
        observations.extend(item.get("tile_observations") or [{
            key: item.get(key) for key in ("relative_rect", "semantic_name", "entity_type",
                                          "role", "observed_value", "confidence", "source_tiles")
        }])
    merged["tile_observations"] = observations
    variants = list(dict.fromkeys(
        value for item in (existing, target)
        for value in [*item.get("observed_value_variants", []), item.get("observed_value", "")]
        if value))
    if len(variants) > 1:
        merged["observed_value_variants"] = variants
        merged["caption_ambiguity"] = "Conflicting tile captions refer to the same pixel target"
    merged["source_tiles"] = sorted(set(existing.get("source_tiles", []))
                                    | set(target.get("source_tiles", [])))
    controls[position] = {**merged, "id": existing["id"]}


def segment_interface(
    screenshot,
    api_url: str,
    api_key: str,
    model: str,
    version: str,
    progress: Callable[[str], None] | None = None,
    analysis_context: str = "",
) -> dict:
    """Analyze overlapping native-resolution tiles instead of shrinking a desktop."""
    width, height = screenshot.size
    tiles = [
        (x, y, min(width, x + 1280), min(height, y + 960))
        for y in _tile_origins(height, 960)
        for x in _tile_origins(width, 1280)
    ]
    report = progress or (lambda _message: None)
    context_args = (analysis_context,) if analysis_context else ()
    if len(tiles) == 1:
        result = _segment_image(screenshot, api_url, api_key, model, version, *context_args)
        result["coverage"] = {"visual_tiles": 1}
        return result
    results = {}
    with ThreadPoolExecutor(max_workers=3) as pool:
        pending = {
            pool.submit(
                _segment_image, screenshot.crop(bounds), api_url, api_key, model, version,
                *context_args,
            ): index
            for index, bounds in enumerate(tiles)
        }
        for future in as_completed(pending):
            try:
                results[pending[future]] = future.result()
            except Exception as error:
                for job in pending:
                    job.cancel()
                raise RuntimeError(
                    f"Vision tile {pending[future] + 1}/{len(tiles)} failed: {error}"
                ) from error
            report(f"Segmenting interface tiles {len(results)}/{len(tiles)}")
    regions, controls = [], []
    for index, (x, y, tile_right, tile_bottom) in enumerate(tiles):
        result = results[index]
        prefix = f"tile-{index + 1}-"
        for region in result["regions"]:
            rect = dict(region["rect"])
            rect.update(left=rect["left"] + x, right=rect["right"] + x,
                        top=rect["top"] + y, bottom=rect["bottom"] + y)
            regions.append({**region, "id": prefix + region["id"], "rect": rect})
        for target in result["controls"]:
            rect = dict(target["relative_rect"])
            clipped = ((x > 0 and rect["left"] <= 2)
                       or (y > 0 and rect["top"] <= 2)
                       or (tile_right < width and rect["right"] >= tile_right - x - 2)
                       or (tile_bottom < height and rect["bottom"] >= tile_bottom - y - 2))
            rect.update(left=rect["left"] + x, right=rect["right"] + x,
                        top=rect["top"] + y, bottom=rect["bottom"] + y)
            translated = {**target, "region_id": prefix + target["region_id"],
                          "relative_rect": rect, "_tile_clipped": clipped,
                          "source_tiles": [index + 1]}
            _merge_tile_target(controls, translated)
    return {"page": results[0]["page"], "regions": regions, "controls": controls,
            "coverage": {"visual_tiles": len(tiles), "tile_bounds": tiles,
                         "raw_observations": sum(len(r["controls"]) for r in results.values()),
                         "merged_observations": sum(len(r["controls"]) for r in results.values())
                                                - len(controls),
                         "edge_clipped_targets": sum(c.get("_tile_clipped", False)
                                                     for c in controls)}}


def _segment_image(
    screenshot,
    api_url: str,
    api_key: str,
    model: str,
    version: str,
    analysis_context: str = "",
) -> dict:
    """Ask the multimodal endpoint for every meaningful visible UI element."""
    buffer = io.BytesIO()
    screenshot.save(buffer, format="PNG")
    image_url = "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")
    width, height = screenshot.size
    prompt = (
        f"Analyze this {width}x{height} desktop application screenshot. "
        "Identify the current page, divide it into functional regions, and enumerate every "
        "visible actionable item and meaningful read-only element. Do not sample representative "
        "items: enumerate each desktop shortcut, file, list row, toolbar button and tray icon "
        "separately, including repeated-looking items. Exclude decorative, layout-only, "
        "inaccessible background elements and duplicate detections of the same item. Use IDs in "
        "kebab-case. Return JSON only with this shape: "
        '{"page":{"id":"main","name":"Main"},'
        '"regions":[{"id":"navigation","name":"Navigation","role":"navigation",'
        '"left":0,"top":0,"width":100,"height":100}],'
        '"controls":[{"id":1,"region":"navigation","semantic_name":"Open menu",'
        '"intent":"open-menu","role":"action","actions":["click"],"risk":"safe",'
        '"confidence":0.97,"left":10,"top":10,"width":40,"height":40}]}. '
        "Use compact JSON. Omit optional empty/default fields; do not repeat boilerplate "
        "descriptions for every icon. Optional fields are entity_type, observed_value, "
        "dynamic_context, current_state, state_reason, enabling_condition, relationships, "
        "aliases, requires_confirmation, description, ambiguity and evidence. Include these "
        "when they carry visible information; keep descriptions under twelve words. "
        "Every control must include numeric confidence from 0 to 1. Use low confidence and "
        "ambiguity when an icon's function cannot be established from visible evidence. "
        "Keep a visually distinct notification badge as its own element even when attached "
        "to another icon; use a relationship with type='attached-to' and target_intent set "
        "to the parent icon's intent. "
        "Visible text is also a control: enumerate standalone labels, captions, values, "
        "titles and pin/port text as entity_type='text-label', with their literal content "
        "in observed_value. Use one rectangle per independent label or line, not per glyph. "
        "Text and its owning graphic have separate IDs and may legitimately overlap. "
        "Use a relationship type='text-of' or 'label-of' for that association; do not use "
        "attached-to for an external caption or enlarge the graphic box to include it. "
        "Do not split intrinsic icon artwork, decorative accents or highlights into extra "
        "UI elements. A colored detail alone is not evidence of a separate notification badge. "
        "Coordinates must use screenshot pixels. This image may be one tile of a larger window; "
        "report only elements visible in this tile. Never replace a group of distinct items "
        "with one generic group rectangle. Control rectangles must tightly cover each target. "
        "Include "
        "read-only status and result fields because they are needed to verify operation effects. "
        "Include meaningful visible entities even when they are not directly actionable: names, "
        "avatars, list-row identity, message/content areas, inputs, badges, status values, and "
        "disabled actions. Separate stable function from the currently observed value. Use a "
        "generic semantic_name such as 'Current conversation contact name' and put the visible "
        "name in observed_value with dynamic_context=true. Infer current_state, state_reason, "
        "enabling_condition, and relationships between controls when visually grounded. For "
        "example, an empty message input can explain why Send is disabled and non-empty draft "
        "text is its enabling condition. "
        "Actions may be click, double-click, right-click, hover, or set-text; never propose "
        "scroll or drag. "
        "Risk must be safe, state-changing, external, destructive, or unknown. Sending, "
        "publishing, purchasing, deleting, uploading, logging out, and closing unsaved work must "
        "require confirmation."
    )
    if analysis_context:
        prompt += (" Analysis scope and grouping guidance (takes precedence over generic "
                   "UI examples and enumeration rules above): " + analysis_context[:4000])
    payload = {
        "model": model,
        **configuration.vision_request_options(model),
        "temperature": 0,
        # Structured output is validated at the provider as well as locally.
        # Some compatible gateways return empty/non-JSON streaming completions
        # for large images even when small streaming probes work.
        "stream": False,
        "response_format": {"type": "json_object"},
        "messages": [
            {
                "role": "system",
                "content": "You segment application interfaces and return strict JSON only.",
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": image_url}},
                ],
            },
        ],
    }
    if model.lower().startswith(("gpt-5", "gpt-6", "o1", "o3", "o4")):
        payload["reasoning_effort"] = "low"
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
        with urlrequest.urlopen(request, timeout=120) as response:
            content = _read_completion_content(response)
    except urlerror.HTTPError as error:
        raise RuntimeError(f"AI segmentation failed with HTTP {error.code}") from error
    except (urlerror.URLError, TimeoutError) as error:
        reason = getattr(error, "reason", str(error))
        raise RuntimeError(f"AI segmentation failed: {reason}") from error
    result = _parse_json_object(str(content))
    if (not isinstance(result.get("page"), dict)
            or not isinstance(result.get("regions"), list)
            or not isinstance(result.get("controls"), list)):
        raise RuntimeError("Vision response is missing page, regions or controls")
    return _validate_segments(result, width, height)


def _annotate_controls(screenshot, candidates: list[dict]):
    """Draw stable numeric labels without changing the source screenshot."""
    from PIL import ImageDraw

    annotated = screenshot.convert("RGB").copy()
    draw = ImageDraw.Draw(annotated)
    for candidate in candidates:
        rectangle = candidate["relative_rect"]
        left = rectangle["left"]
        top = rectangle["top"]
        right = rectangle["right"]
        bottom = rectangle["bottom"]
        label = str(candidate["id"])
        draw.rectangle((left, top, right, bottom), outline="#ff1744", width=3)
        label_width = max(18, 8 * len(label) + 8)
        label_top = max(0, top - 18)
        draw.rectangle(
            (left, label_top, min(annotated.width, left + label_width), label_top + 18),
            fill="#ff1744",
        )
        draw.text((left + 4, label_top + 2), label, fill="white")
    return annotated


def _image_data_url(image) -> str:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")


def _semantic_context(raw: dict) -> dict:
    relationships = []
    raw_relationships = raw.get("relationships", [])
    if not isinstance(raw_relationships, list):
        raw_relationships = []
    for item in raw_relationships[:12]:
        if not isinstance(item, dict):
            continue
        relationship = {
            "type": knowledge.slugify(item.get("type") or "related-to"),
            "target_intent": knowledge.slugify(item.get("target_intent") or "", ""),
            "target_entity": knowledge.slugify(item.get("target_entity") or "", ""),
            "description": str(item.get("description") or "").strip()[:300],
        }
        relationships.append({key: value for key, value in relationship.items() if value})
    return {
        "entity_type": knowledge.slugify(raw.get("entity_type") or "ui-element"),
        "observed_value": str(raw.get("observed_value") or "").strip()[:500],
        "dynamic_context": bool(raw.get("dynamic_context", False)),
        "current_state": knowledge.slugify(raw.get("current_state") or "unknown"),
        "state_reason": str(raw.get("state_reason") or "").strip()[:500],
        "enabling_condition": str(raw.get("enabling_condition") or "").strip()[:500],
        "relationships": relationships,
    }


def _visual_component(target: dict, window_rect: dict) -> dict:
    rectangle = target.get("relative_rect")
    component = {
        "semantic_name": target.get("semantic_name", ""),
        "intent": target.get("intent", ""),
        "description": target.get("description", ""),
        "role": target.get("role", "control"),
        **_semantic_context(target),
    }
    if target.get("native_evidence"):
        component["native_evidence"] = True
    if isinstance(rectangle, dict):
        component["normalized_rect"] = knowledge.normalized_rect(
            rectangle, window_rect, relative=True
        )
        component["geometry_role"] = "visual-hint-only"
    return component


def _request_control_semantics(
    screenshot,
    candidates: list[dict],
    page: dict,
    regions: list[dict],
    api_url: str,
    api_key: str,
    model: str,
    version: str,
) -> dict[int, dict]:
    annotated = _annotate_controls(screenshot, candidates)
    compact_candidates = [
        {
            "id": item["id"],
            "uia_name": item["uia_name"],
            "automation_id": item["automation_id"],
            "control_type": item["control_type"],
            "localized_control_type": item["localized_control_type"],
            "class_name": item["class_name"],
            "help_text": item["help_text"],
            "is_keyboard_focusable": item["is_keyboard_focusable"],
            "is_enabled": item["is_enabled"],
            "region": item["region_id"],
            "supported_actions": item["supported_actions"],
            "rect": item["relative_rect"],
        }
        for item in candidates
    ]
    region_context = [
        {
            "id": item["id"],
            "name": item["name"],
            "role": item["role"],
            "description": item.get("description", ""),
        }
        for item in regions
    ]
    prompt = (
        "Understand the real function of every numbered desktop UI control. "
        "Use the original screenshot for visual context and the annotated screenshot to map "
        "numeric IDs. Use the UI Automation metadata below too. Infer meaning from labels, icons, "
        "nearby text, region, layout, and application context together. Do not merely repeat the "
        "control type or invent a purpose when evidence is weak. For ambiguous controls, explain "
        "the ambiguity and set confidence "
        "below 0.72. Return every input ID exactly once as strict JSON only: "
        '{"controls":[{"id":1,"semantic_name":"Save document","intent":"save-document",'
        '"description":"Save the current document to disk","role":"action",'
        '"entity_type":"document-action","observed_value":"",'
        '"dynamic_context":false,"current_state":"enabled","state_reason":"",'
        '"enabling_condition":"",'
        '"relationships":[{"type":"acts-on","target_intent":"current-document"}],'
        '"actions":["click"],"aliases":["save","write file"],'
        '"risk":"state-changing","requires_confirmation":false,"confidence":0.96,'
        '"evidence":["floppy-disk icon","toolbar context"],"ambiguity":""}]}. '
        "semantic_name must be concise and user-facing. intent must be a stable generic kebab-case "
        "verb phrase suitable for a CLI. actions must be a subset of supported_actions. risk must "
        "be one of safe, state-changing, external, destructive, unknown. Sending, publishing, "
        "purchasing, deleting, "
        "closing without saving, or changing remote state must require confirmation. "
        "Separate stable function from volatile screen content. For example, a chat composer is "
        "'Current conversation message input', while a visible contact name such as Alice is "
        "observed_value='Alice' with dynamic_context=true. Identify entity_type for contact names, "
        "avatars, conversation rows, message history, inputs, status fields, and action buttons. "
        "Infer current_state, state_reason, enabling_condition, and relationships only when the "
        "screenshot, UIA enabled state, and surrounding controls support them. A disabled Send "
        "button beside an empty message input should state that non-empty draft text enables it. "
        f"Page: {json.dumps(page, ensure_ascii=False)}. "
        f"Regions: {json.dumps(region_context, ensure_ascii=False)}. "
        f"Controls: {json.dumps(compact_candidates, ensure_ascii=False)}."
    )
    payload = {
        "model": model,
        **configuration.vision_request_options(model),
        "temperature": 0,
        "stream": True,
        "messages": [
            {
                "role": "system",
                "content": (
                    "You are a desktop UI analyst. Return grounded control semantics "
                    "as strict JSON."
                ),
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": _image_data_url(screenshot)}},
                    {"type": "image_url", "image_url": {"url": _image_data_url(annotated)}},
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
        with urlrequest.urlopen(request, timeout=90) as response:
            content = _read_completion_content(response)
    except urlerror.HTTPError as error:
        raise RuntimeError(f"AI control semantics failed with HTTP {error.code}") from error
    except (urlerror.URLError, TimeoutError) as error:
        reason = getattr(error, "reason", str(error))
        raise RuntimeError(f"AI control semantics failed: {reason}") from error
    result = _parse_json_object(str(content))
    expected = {item["id"]: item for item in candidates}
    semantics = {}
    for raw in result.get("controls", []):
        if not isinstance(raw, dict):
            continue
        try:
            item_id = int(raw.get("id"))
        except (TypeError, ValueError):
            continue
        candidate = expected.get(item_id)
        if candidate is None or item_id in semantics:
            continue
        supported = candidate["supported_actions"]
        requested_actions = raw.get("actions", [])
        if not isinstance(requested_actions, list):
            requested_actions = []
        actions = [action for action in supported if action in requested_actions]
        try:
            confidence = max(0.0, min(1.0, float(raw.get("confidence", 0))))
        except (TypeError, ValueError):
            confidence = 0.0
        semantic_name = str(raw.get("semantic_name") or "").strip()[:120]
        intent = knowledge.slugify(raw.get("intent") or "", "")
        risk = str(raw.get("risk") or "unknown").strip().lower()
        if risk not in SEMANTIC_RISKS:
            risk = "unknown"
        aliases = raw.get("aliases", [])
        evidence = raw.get("evidence", [])
        semantics[item_id] = {
            "semantic_name": semantic_name,
            "intent": intent,
            "description": str(raw.get("description") or "").strip()[:500],
            "role": knowledge.slugify(raw.get("role") or "control"),
            "actions": actions,
            "aliases": [str(value).strip() for value in aliases[:12] if str(value).strip()]
            if isinstance(aliases, list)
            else [],
            "risk": risk,
            "requires_confirmation": bool(raw.get("requires_confirmation"))
            or risk in {"external", "destructive"},
            "confidence": round(confidence, 4),
            "evidence": [str(value).strip() for value in evidence[:12] if str(value).strip()]
            if isinstance(evidence, list)
            else [],
            "ambiguity": str(raw.get("ambiguity") or "").strip()[:500],
            "source": "ai-vision-context",
            **_semantic_context(raw),
        }
    return semantics


def analyze_control_semantics(
    screenshot,
    candidates: list[dict],
    page: dict,
    regions: list[dict],
    api_url: str,
    api_key: str,
    model: str,
    version: str,
    batch_size: int = SEMANTIC_BATCH_SIZE,
) -> dict[int, dict]:
    """Batch controls through the multimodal model and require one result per control."""
    semantics = {}
    batch_size = max(1, min(int(batch_size), SEMANTIC_BATCH_SIZE))
    for offset in range(0, len(candidates), batch_size):
        batch = candidates[offset : offset + batch_size]
        semantics.update(
            _request_control_semantics(
                screenshot,
                batch,
                page,
                regions,
                api_url,
                api_key,
                model,
                version,
            )
        )
    missing = [item["id"] for item in candidates if item["id"] not in semantics]
    if missing:
        raise RuntimeError(
            "AI control semantics omitted control IDs: " + ", ".join(map(str, missing[:20]))
        )
    return semantics


def _validate_segments(result: dict, width: int, height: int) -> dict:
    page = result.get("page") if isinstance(result.get("page"), dict) else {}
    page = {
        "id": knowledge.slugify(page.get("id") or page.get("name") or "main", "main"),
        "name": str(page.get("name") or "Main"),
        "description": str(page.get("description") or ""),
    }
    regions = []
    for index, region in enumerate(result.get("regions", []), start=1):
        if not isinstance(region, dict):
            continue
        try:
            left = max(0, min(width, int(region.get("left", 0))))
            top = max(0, min(height, int(region.get("top", 0))))
            right = max(left, min(width, left + int(region.get("width", 0))))
            bottom = max(top, min(height, top + int(region.get("height", 0))))
        except (TypeError, ValueError):
            continue
        if right - left < 2 or bottom - top < 2:
            continue
        regions.append(
            {
                "id": knowledge.slugify(
                    region.get("id") or region.get("name") or f"region-{index}"
                ),
                "name": str(region.get("name") or f"Region {index}"),
                "role": knowledge.slugify(region.get("role") or "content"),
                "description": str(region.get("description") or ""),
                "rect": {
                    "left": left,
                    "top": top,
                    "right": right,
                    "bottom": bottom,
                    "width": right - left,
                    "height": bottom - top,
                },
            }
        )
    if not regions:
        regions.append(
            {
                "id": "content",
                "name": "Content",
                "role": "content",
                "description": "Fallback full-window region",
                "rect": {
                    "left": 0,
                    "top": 0,
                    "right": width,
                    "bottom": height,
                    "width": width,
                    "height": height,
                },
            }
        )
    controls = []
    for index, raw in enumerate(result.get("controls", []), start=1):
        if not isinstance(raw, dict):
            continue
        try:
            left = max(0, min(width, int(raw.get("left", 0))))
            top = max(0, min(height, int(raw.get("top", 0))))
            right = max(left, min(width, left + int(raw.get("width", 0))))
            bottom = max(top, min(height, top + int(raw.get("height", 0))))
            confidence = max(0.0, min(1.0, float(raw.get("confidence", 0))))
        except (TypeError, ValueError):
            continue
        if right - left < 2 or bottom - top < 2:
            continue
        risk = str(raw.get("risk") or "unknown").strip().lower()
        if risk not in SEMANTIC_RISKS:
            risk = "unknown"
        actions = raw.get("actions", [])
        aliases = raw.get("aliases", [])
        evidence = raw.get("evidence", [])
        controls.append(
            {
                "id": index,
                "region_id": knowledge.slugify(raw.get("region") or "unassigned"),
                "semantic_name": str(raw.get("semantic_name") or "").strip()[:120],
                "intent": knowledge.slugify(raw.get("intent") or "", ""),
                "description": str(raw.get("description") or "").strip()[:500],
                "role": knowledge.slugify(raw.get("role") or "control"),
                "actions": [
                    str(value).strip()
                    for value in actions[:8]
                    if str(value).strip()
                ]
                if isinstance(actions, list)
                else [],
                "aliases": [
                    str(value).strip()
                    for value in aliases[:12]
                    if str(value).strip()
                ]
                if isinstance(aliases, list)
                else [],
                "risk": risk,
                "requires_confirmation": bool(raw.get("requires_confirmation"))
                or risk in {"external", "destructive"},
                "confidence": round(confidence, 4),
                "evidence": [
                    str(value).strip()
                    for value in evidence[:12]
                    if str(value).strip()
                ]
                if isinstance(evidence, list)
                else [],
                "ambiguity": str(raw.get("ambiguity") or "").strip()[:500],
                "source": "ai-vision-target",
                **_semantic_context(raw),
                "relative_rect": {
                    "left": left,
                    "top": top,
                    "right": right,
                    "bottom": bottom,
                    "width": right - left,
                    "height": bottom - top,
                },
            }
        )
    return {"page": page, "regions": regions, "controls": controls}


def _region_for_control(control_rect: dict, window_rect: dict, regions: list[dict]) -> str:
    center_x = (control_rect["left"] + control_rect["right"]) / 2 - window_rect["left"]
    center_y = (control_rect["top"] + control_rect["bottom"]) / 2 - window_rect["top"]
    containing = []
    for region in regions:
        rectangle = region["rect"]
        if (
            rectangle["left"] <= center_x <= rectangle["right"]
            and rectangle["top"] <= center_y <= rectangle["bottom"]
        ):
            containing.append(region)
    if not containing:
        return "unassigned"
    containing.sort(key=lambda region: region["rect"]["width"] * region["rect"]["height"])
    return containing[0]["id"]


def _control_actions(control_type: str, _is_enabled: bool) -> list[str]:
    """Return potential actions; current enabled state is a separate learned precondition."""
    return ACTIONABLE_TYPES.get(control_type, [])


def _control_semantic_name(control, index: int) -> str:
    name = (getattr(control, "Name", "") or "").strip()
    automation_id = (getattr(control, "AutomationId", "") or "").strip()
    control_type = (getattr(control, "ControlTypeName", "") or "control").removesuffix("Control")
    return name or automation_id or f"{control_type} {index}"


def _command_name(page_id: str, region_id: str, semantic_name: str, control_id: str) -> str:
    control_slug = knowledge.slugify(semantic_name, control_id)
    return ".".join((knowledge.slugify(page_id), knowledge.slugify(region_id), control_slug))


def _previous_control_by_unique_automation_id(
    records: list[dict], automation_id: str, control_type: str
) -> dict:
    if not automation_id:
        return {}
    matches = [
        record
        for record in records
        if record.get("automation_id") == automation_id
        and record.get("control_type") == control_type
    ]
    return matches[0] if len(matches) == 1 else {}


def _relative_rect(control_rect: dict, window_rect: dict) -> dict:
    return {
        "left": control_rect["left"] - window_rect["left"],
        "top": control_rect["top"] - window_rect["top"],
        "right": control_rect["right"] - window_rect["left"],
        "bottom": control_rect["bottom"] - window_rect["top"],
        "width": control_rect["width"],
        "height": control_rect["height"],
    }


def _verify_location(
    location: dict,
    expected_rect: dict,
    window=None,
    prefix_cache: dict | None = None,
) -> tuple[bool, dict, str]:
    try:
        located = resolve_location(location, window=window, prefix_cache=prefix_cache)
        if not located or located is False:
            return False, {}, "LOCATION did not resolve"
        actual_rect = _rect(located)
    except Exception as error:
        return False, {}, f"LOCATION resolution failed: {error}"
    intersection_width = max(
        0,
        min(expected_rect["right"], actual_rect["right"])
        - max(expected_rect["left"], actual_rect["left"]),
    )
    intersection_height = max(
        0,
        min(expected_rect["bottom"], actual_rect["bottom"])
        - max(expected_rect["top"], actual_rect["top"]),
    )
    intersection = intersection_width * intersection_height
    union = (
        expected_rect["width"] * expected_rect["height"]
        + actual_rect["width"] * actual_rect["height"]
        - intersection
    )
    iou = intersection / union if union else 0.0
    return iou >= 0.8, actual_rect, f"bounds IoU={iou:.3f}"


def _visual_sample_points(target: dict, window_rect: dict) -> list[tuple[int, int]]:
    rectangle = target["relative_rect"]
    left = window_rect["left"] + rectangle["left"]
    top = window_rect["top"] + rectangle["top"]
    right = window_rect["left"] + rectangle["right"]
    bottom = window_rect["top"] + rectangle["bottom"]
    center_x = (left + right) // 2
    center_y = (top + bottom) // 2
    inset_x = max(1, (right - left) // 4)
    inset_y = max(1, (bottom - top) // 4)
    return list(
        dict.fromkeys(
            [
                (center_x, center_y),
                (left + inset_x, center_y),
                (right - inset_x, center_y),
                (center_x, top + inset_y),
                (center_x, bottom - inset_y),
            ]
        )
    )


def _control_key(control) -> tuple:
    rectangle = _rect(control)
    return (
        getattr(control, "NativeWindowHandle", 0) or 0,
        getattr(control, "ControlTypeName", "") or "",
        getattr(control, "AutomationId", "") or "",
        getattr(control, "Name", "") or "",
        rectangle["left"],
        rectangle["top"],
        rectangle["right"],
        rectangle["bottom"],
    )


def _same_visual_target(first: dict, second: dict) -> bool:
    """Only deduplicate overlapping detections with compatible entity identities."""
    if _rect_iou(first["relative_rect"], second["relative_rect"]) < 0.65:
        return False
    for key in ("role", "entity_type", "observed_value"):
        a, b = str(first.get(key) or "").casefold(), str(second.get(key) or "").casefold()
        if a and b and a != b:
            return False
    a, b = first.get("intent"), second.get("intent")
    same_value = bool(first.get("observed_value")) and (
        first.get("observed_value") == second.get("observed_value")
    )
    return not (a and b and a != b and not same_value)


def _is_desktop_window(window) -> bool:
    return (getattr(window, "ClassName", "") in {"Progman", "WorkerW"}
            or getattr(window, "Name", "") == "Program Manager")


def _desktop_shell_roots(window) -> list:
    """Inspect Explorer shell roots only, never other applications' UI trees."""
    roots = {_control_key(window): window}
    for entry in native_windows.enumerate_windows(visible_only=True):
        if entry["class_name"] in {
            "Progman", "WorkerW", "Shell_TrayWnd", "Shell_SecondaryTrayWnd"
        }:
            if entry["handle"] == int(getattr(window, "NativeWindowHandle", 0) or 0):
                continue
            control = uiautomation.ControlFromHandle(entry["handle"])
            if control is not None:
                roots.setdefault(_control_key(control), control)
    return list(roots.values())


def _desktop_item_is_visible(control) -> bool:
    """UIA IsOffscreen does not account for other windows covering an icon."""
    rect = _rect(control)
    expected = _control_key(control)
    try:
        current = uiautomation.ControlFromPoint(
            (rect["left"] + rect["right"]) // 2,
            (rect["top"] + rect["bottom"]) // 2,
        )
        for _ in range(8):
            if not current:
                break
            if _control_key(current) == expected:
                return True
            current = current.GetParentControl()
    except Exception:
        pass
    return False


def _supplement_desktop_targets(window, window_rect: dict, segmentation: dict,
                                max_controls: int) -> None:
    """Retain every exposed visible desktop item even when vision omits it."""
    targets = segmentation.setdefault("controls", [])
    seen, icons, added, occluded = set(), 0, 0, 0
    queue = [(shell, 0) for shell in _desktop_shell_roots(window)]
    limited = False
    while queue:
        control, depth = queue.pop(0)
        key = _control_key(control)
        if key in seen:
            continue
        seen.add(key)
        control_type = getattr(control, "ControlTypeName", "")
        if control_type == "ListItemControl":
            rect = _rect(control)
            if (not _valid_rect(rect, window_rect)
                    or bool(getattr(control, "IsOffscreen", False))):
                continue
            if not _desktop_item_is_visible(control):
                occluded += 1
                continue
            icons += 1
            if icons > max_controls:
                limited = True
                break
            relative = _relative_rect(_clip_rect(rect, window_rect), window_rect)
            name = str(getattr(control, "Name", "") or "").strip()
            matched = None
            for target in targets:
                visual = target["relative_rect"]
                cx, cy = ((visual["left"] + visual["right"]) / 2,
                          (visual["top"] + visual["bottom"]) / 2)
                if (relative["left"] <= cx <= relative["right"]
                        and relative["top"] <= cy <= relative["bottom"]
                        and _rect_iou(relative, visual) >= 0.2
                        and (not target.get("observed_value")
                             or str(target["observed_value"]).casefold() == name.casefold())
                        and not target.get("_uia_control")):
                    matched = target
                    break
            if matched is not None:
                matched["_uia_control"] = control
                continue
            targets.append({
                "id": len(targets) + 1, "region_id": "desktop-items",
                "semantic_name": "Desktop item", "intent": "open-desktop-item",
                "description": "Open the named desktop item; its effect has not been tested.",
                "role": "action", "entity_type": "desktop-item", "observed_value": name,
                "dynamic_context": True, "actions": ["double-click"], "aliases": [name],
                "risk": "unknown", "requires_confirmation": True, "confidence": 1.0,
                "source": "uia-desktop-item", "relative_rect": relative,
                "_uia_control": control,
            })
            added += 1
            continue
        if depth < 4 and getattr(control, "ClassName", "") in {
            "Progman", "WorkerW", "SHELLDLL_DefView", "SysListView32"
        }:
            queue.extend((child, depth + 1) for child in control.GetChildren())
    if added:
        segmentation["regions"].append({
            "id": "desktop-items", "name": "Desktop items", "role": "desktop-workspace",
            "description": "Individually enumerated Windows desktop items",
            "rect": _relative_rect(window_rect, window_rect),
        })
    segmentation.setdefault("coverage", {}).update(
        desktop_items_seen=icons, desktop_items_added=added, desktop_limit_reached=limited,
        desktop_items_occluded=occluded,
    )


def _segment_desktop(screenshot, window, window_rect: dict, api_url: str, api_key: str,
                     model: str, version: str, max_controls: int, progress) -> dict:
    """Use exact native desktop items and visual analysis of the taskbars only."""
    width, height = screenshot.size
    segmentation = _validate_segments({"page": {
        "id": "desktop", "name": "Windows Desktop",
        "description": "Native desktop items with visually analyzed taskbars",
    }}, width, height)
    _supplement_desktop_targets(window, window_rect, segmentation, max_controls)
    bars = []
    for shell in _desktop_shell_roots(window):
        if getattr(shell, "ClassName", "") in {"Shell_TrayWnd", "Shell_SecondaryTrayWnd"}:
            rect = _clip_rect(_rect(shell), window_rect)
            if _valid_rect(rect, window_rect):
                bars.append(_relative_rect(rect, window_rect))
    for index, rect in enumerate(bars, start=1):
        progress(f"Segmenting desktop taskbars {index}/{len(bars)}")
        crop = screenshot.crop((rect["left"], rect["top"], rect["right"], rect["bottom"]))
        result = _segment_image(crop, api_url, api_key, model, version)
        prefix = f"taskbar-{index}-"
        for key, rectangle_key in (("regions", "rect"), ("controls", "relative_rect")):
            for item in result[key]:
                translated = dict(item)
                local = dict(item[rectangle_key])
                local.update(left=local["left"] + rect["left"],
                             right=local["right"] + rect["left"],
                             top=local["top"] + rect["top"],
                             bottom=local["bottom"] + rect["top"])
                translated[rectangle_key] = local
                if key == "regions":
                    translated["id"] = prefix + item["id"]
                else:
                    translated["id"] = len(segmentation["controls"]) + 1
                    translated["region_id"] = prefix + item["region_id"]
                segmentation[key].append(translated)
    segmentation["coverage"].update(visual_tiles=len(bars),
                                     desktop_strategy="native-items-and-visual-taskbars")
    return segmentation


def _candidate_fits_target(
    control, absolute_target: dict, window, diagnostics: dict | None = None,
) -> bool:
    rectangle = _rect(control)
    target_area = max(1, absolute_target["width"] * absolute_target["height"])
    area_ratio = rectangle["width"] * rectangle["height"] / target_area
    generic = getattr(control, "ControlTypeName", "") in GENERIC_CONTAINER_TYPES
    max_area_ratio = 4 if generic else 16
    intersection = (max(0, min(rectangle["right"], absolute_target["right"])
                        - max(rectangle["left"], absolute_target["left"]))
                    * max(0, min(rectangle["bottom"], absolute_target["bottom"])
                          - max(rectangle["top"], absolute_target["top"])))
    overlap_ratio = intersection / target_area
    reason = ""
    if _control_key(control) == _control_key(window):
        reason = "window_itself"
    elif area_ratio > max_area_ratio:
        reason = "area_too_large"
    elif area_ratio < 0.1:
        reason = "area_too_small"
    elif overlap_ratio < 0.5:
        reason = "insufficient_overlap"
    if diagnostics is not None:
        diagnostics.update(
            rect=rectangle, reason=reason, area_ratio=round(area_ratio, 4),
            max_area_ratio=max_area_ratio, overlap_ratio=round(overlap_ratio, 4),
        )
    return not reason


def _candidate_score(control, target: dict, absolute_target: dict, window) -> tuple:
    control_type = getattr(control, "ControlTypeName", "") or "Control"
    actions = _control_actions(control_type, bool(getattr(control, "IsEnabled", True)))
    rectangle = _rect(control)
    target_actions = set(target.get("actions", []))
    score = 0.0
    if target_actions and target_actions.intersection(actions):
        score += 8
    elif actions:
        score += 4
    if getattr(control, "AutomationId", ""):
        score += 4
    if getattr(control, "Name", ""):
        score += 2
    if control_type not in GENERIC_CONTAINER_TYPES:
        score += 2
    elif not actions:
        score -= 4
    if target.get("role") in {"status", "result", "label", "value"} and control_type in {
        "TextControl",
        "EditControl",
        "DocumentControl",
    }:
        score += 5
    center_x = (absolute_target["left"] + absolute_target["right"]) // 2
    center_y = (absolute_target["top"] + absolute_target["bottom"]) // 2
    if rectangle["left"] <= center_x <= rectangle["right"]:
        score += 1
    if rectangle["top"] <= center_y <= rectangle["bottom"]:
        score += 1
    overlap = _rect_iou(rectangle, absolute_target)
    score += overlap * 4
    if control is window or control_type == "DesktopControl":
        score -= 20
    area = max(1, rectangle["width"] * rectangle["height"])
    return score, -area


def control_from_visual_target(
    window,
    window_rect: dict,
    target: dict,
    point_lookup: Callable[[int, int], object] | None = None,
    *,
    diagnostics: dict | None = None,
):
    """Choose a stable UIA target, optionally explaining lookup/filter failures.

    The return value remains a control or None. Diagnostics describe this target's
    point samples and candidates, never the target application's full UIA tree.
    """
    point_lookup = point_lookup or uiautomation.ControlFromPoint
    relative = target["relative_rect"]
    absolute_target = {
        "left": window_rect["left"] + relative["left"],
        "top": window_rect["top"] + relative["top"],
        "right": window_rect["left"] + relative["right"],
        "bottom": window_rect["top"] + relative["bottom"],
        "width": relative["width"],
        "height": relative["height"],
    }
    details = diagnostics if diagnostics is not None else {}
    details.clear()
    details.update(
        reason="", target_rect=absolute_target, sample_points=[], candidate_count=0,
        geometry_matched=0, semantics_matched=0, rejected_candidates=[],
        rejection_count=0, lookup_errors=[], lookup_error_count=0,
    )
    assessments = {}

    def qualifies(key, control):
        if key in assessments:
            return assessments[key]
        geometry = {}
        fits = _candidate_fits_target(control, absolute_target, window, geometry)
        rejection = {"name": key[3], "control_type": key[1], "automation_id": key[2],
                     **geometry}
        if fits:
            details["geometry_matched"] += 1
            binding = semantic_identity.check(target, semantic_identity.native_metadata(control))
            fits = binding["status"] not in {"conflict", "unresolved"}
            if fits:
                details["semantics_matched"] += 1
            else:
                rejection.update(reason="semantic_" + binding["status"], semantic_binding=binding)
        assessments[key] = fits
        if not fits:
            details["rejection_count"] += 1
            if len(details["rejected_candidates"]) < 8:
                details["rejected_candidates"].append(rejection)
        return fits

    def lookup_error(stage, x, y, error):
        details["lookup_error_count"] += 1
        if len(details["lookup_errors"]) < 8:
            details["lookup_errors"].append(
                {"stage": stage, "x": x, "y": y,
                 "error": f"{type(error).__name__}: {error}"}
            )

    candidates = {}
    direct = target.get("_uia_control")
    if direct is not None:
        key = _control_key(direct)
        candidates[key] = direct
        details["candidate_count"] = 1
        if qualifies(key, direct):
            details["reason"] = "resolved"
            return direct
    for x, y in _visual_sample_points(target, window_rect):
        sample = {"x": x, "y": y, "status": "empty"}
        details["sample_points"].append(sample)
        try:
            current = point_lookup(x, y)
        except Exception as error:
            sample["status"] = "error"
            lookup_error("point_lookup", x, y, error)
            continue
        if current:
            sample["status"] = "hit"
        for _depth in range(16):
            if not current:
                break
            try:
                candidates.setdefault(_control_key(current), current)
                if current is window or getattr(current, "ControlTypeName", "") == "DesktopControl":
                    break
                current = current.GetParentControl()
            except Exception as error:
                lookup_error("candidate_read", x, y, error)
                break
    details["candidate_count"] = len(candidates)
    if not candidates:
        details["reason"] = (
            "uia_lookup_failed" if details["lookup_error_count"] else "no_uia_candidate"
        )
        return None
    ranked = sorted(
        ((_candidate_score(control, target, absolute_target, window), control)
         for key, control in candidates.items() if qualifies(key, control)),
        key=lambda item: item[0],
        reverse=True,
    )
    if not ranked:
        details["reason"] = (
            "semantic_rejected" if details["geometry_matched"] else "geometry_rejected"
        )
        return None
    score, selected = ranked[0]
    details["best_score"] = score[0]
    if score[0] < 0:
        details["reason"] = "score_rejected"
        return None
    details["reason"] = "resolved"
    return selected


def _native_action_controls(window, limit: int = 200):
    """Read only shallow native menu/titlebar containers of the selected window."""
    queue = [(window, 0)]
    containers = {"MenuBarControl", "MenuControl", "TitleBarControl", "PaneControl", "GroupControl"}
    seen = set()
    while queue and len(seen) < limit:
        control, depth = queue.pop(0)
        key = _control_key(control)
        if key in seen:
            continue
        seen.add(key)
        if semantic_identity.literal_semantics(semantic_identity.native_metadata(control)):
            yield control
        if depth < 3 and (depth == 0 or getattr(control, "ControlTypeName", "") in containers):
            try:
                children = control.GetChildren()[:limit - len(seen)]
                queue.extend((child, depth + 1) for child in children)
            except Exception:
                continue


def controls_from_visual_targets(
    window,
    window_rect: dict,
    targets: list[dict],
    max_controls: int,
) -> tuple[list[tuple[object, int]], dict[int, dict], list[dict]]:
    """Resolve visual targets to unique UIA controls without walking the full tree."""
    controls = []
    semantics = {}
    failures = []
    seen: dict[tuple, int] = {}
    for target in targets[: max(1, max_controls)]:
        diagnostics = {}
        control = control_from_visual_target(window, window_rect, target, diagnostics=diagnostics)
        if control is None:
            reason = diagnostics.get("reason", "resolution_failed")
            failures.append(
                {
                    "visual_target": target.get("semantic_name") or target.get("id"),
                    "error": VISUAL_RESOLUTION_ERRORS[reason],
                    "reason": reason,
                    "diagnostics": diagnostics,
                    "observation": target,
                }
            )
            continue
        try:
            xpath = get_control_xpath(control)
            key = (json.dumps(xpath, ensure_ascii=False, sort_keys=True), _control_key(control))
        except Exception as error:
            failures.append(
                {
                    "visual_target": target.get("semantic_name") or target.get("id"),
                    "error": f"XPath failed: {error}",
                    "observation": target,
                }
            )
            continue
        if key in seen:
            index = seen[key]
            primary = semantics[index]
            if _same_visual_target(primary, target):
                continue
            # Shared UIA ownership is not evidence of a shared visual identity.
            # Keep the distinct entity independently visible, with no executable locator.
            visual = target
            if target.get("actions") and not primary.get("actions"):
                visual = primary
                semantics[index] = {**target, "visual_components": []}
            failures.append({"visual_target": visual.get("semantic_name"),
                             "error": "distinct visual entity shares a UIA control",
                             "observation": visual})
            continue
        index = len(controls) + 1
        seen[key] = index
        controls.append((control, max(0, len(xpath) - 1)))
        semantics[index] = {**target, "visual_components": []}
    # Literal menu labels are independent native evidence. Never repair a model
    # box by silently assigning its meaning to the adjacent UIA menu item.
    for control in _native_action_controls(window):
        if len(controls) >= max_controls:
            break
        rectangle = _rect(control)
        if (rectangle["width"] <= 0 or rectangle["height"] <= 0
                or rectangle["left"] < window_rect["left"]
                or rectangle["top"] < window_rect["top"]
                or rectangle["right"] > window_rect["right"]
                or rectangle["bottom"] > window_rect["bottom"]):
            continue
        xpath = get_control_xpath(control)
        key = (json.dumps(xpath, ensure_ascii=False, sort_keys=True), _control_key(control))
        if key in seen:
            continue
        index = len(controls) + 1
        seen[key] = index
        controls.append((control, max(0, len(xpath) - 1)))
        literal = semantic_identity.literal_semantics(semantic_identity.native_metadata(control))
        semantics[index] = {**literal,
                            "relative_rect": _relative_rect(rectangle, window_rect),
                            "visual_components": []}
    return controls, semantics, failures


@image_storage.serialized_learning
def scan_window(
    window_name: str,
    api_url: str,
    api_key: str,
    model: str,
    version: str,
    max_depth: int = 12,
    max_controls: int = 3000,
    verify_limit: int = 500,
    strategy: str = "visual-first",
    root: Path | None = None,
    progress: Callable[[str], None] | None = None,
    show_overlay: bool = False,
    overlay_duration_ms: int = 3000,
    manual_layout: dict | None = None,
    semantic_mode: str = "api",
    navigation_context: dict | None = None,
    observation_owner: dict | None = None,
) -> dict:
    """Scan one visible application window and persist its UI command knowledge."""
    started = time.perf_counter()
    external_progress = progress or (lambda _message: None)
    stage_timings = []

    def progress(message: str) -> None:
        stage_timings.append(
            {
                "stage": message,
                "elapsed_seconds": round(time.perf_counter() - started, 2),
            }
        )
        external_progress(message)
    strategy = strategy.strip().lower()
    if strategy not in SCAN_STRATEGIES:
        raise ValueError(f"strategy must be one of: {', '.join(sorted(SCAN_STRATEGIES))}")
    if semantic_mode not in {"api", "native"}:
        raise ValueError("semantic_mode must be api or native")
    if navigation_context is not None and manual_layout is None:
        raise ValueError("展开控件学习必须绑定人工布局的所选窗口")
    if manual_layout is not None:
        from . import manual_layout as layout_rules

        layout_rules.validate_regions(manual_layout.get("regions", []))
        if strategy == "visual-first":
            strategy = "local-visual"
    elif strategy == "local-visual":
        raise ValueError("local-visual requires defined regions and a bound window")
    progress("Finding target window through Win32")
    if manual_layout is not None:
        candidates = native_windows.matching_windows(
            window_name, handle=int(manual_layout.get("window_handle", 0)))
        candidates = [item for item in candidates if item["visible"]]
        if len(candidates) != 1:
            raise ValueError("所选软件窗口已关闭或不可见，请重新选择软件")
        selected = candidates[0]
        if ((manual_layout.get("window_class") and selected["class_name"] != manual_layout["window_class"])
                or (manual_layout.get("window_process_id") and selected["process_id"] != manual_layout["window_process_id"])):
            raise ValueError("所选窗口已经失效，请重新选择软件")
        window = uiautomation.ControlFromHandle(selected["handle"])
        if window is None:
            raise ValueError("无法读取所选软件的 UI Automation 窗口")
    else:
        window = _find_window(window_name)
    progress("Reading target window through UI Automation")
    identity = window_identity.describe(window, window_name)
    actual_name = (identity["title"] or identity["uia_name"] or window_name).strip()
    access = native_windows.check_uia_access(int(getattr(window, "ProcessId", 0) or 0))
    if observation_owner is not None:
        from ..library_actions import bind_window, owned_windows
        from .control_observation import _check_foreground
        owner = bind_window(observation_owner)
        allowed = {observation_owner["handle"], *(w["handle"] for w in owned_windows(observation_owner))}
        if int(window.NativeWindowHandle) not in allowed:raise ValueError("扫描窗口不属于所选软件")
        _check_foreground(owner)
    elif navigation_context and strategy == "local-visual":
        from .control_observation import _check_foreground
        _check_foreground(window)
    else:
        _activate_window(window)
    window_rect = _window_rect(window)
    if window_rect["width"] <= 1 or window_rect["height"] <= 1:
        raise RuntimeError(f"Window has no visible bounds: {actual_name}")
    manual_segmentation = (layout_rules.segmentation(manual_layout, window_rect)
                           if manual_layout is not None else None)

    process_id = int(getattr(window, "ProcessId", 0) or 0)
    process_name = _process_name(process_id)
    app_id = knowledge.slugify(
        (manual_layout or {}).get("app_id") or process_name or actual_name, "application")
    directory = knowledge.initialize_app(
        app_id, actual_name, root, preserve_name=True)
    knowledge.load_index(directory)
    progress("Capturing window screenshot")
    screenshot = _screenshot_window(window_rect)
    capture_bounds = dict(window_rect)
    if manual_layout is None:
        screenshot, window_rect = _trim_window_capture(screenshot, window_rect, window)
    page_image = directory / "images" / "pages" / "latest.png"
    local_reports, local_paths, local_frames = [], {}, []
    if manual_layout is not None:
        progress("Using human-defined regions")
        segmentation = manual_segmentation
        if strategy == "local-visual":
            from . import local_learning

            segmentation["controls"], local_reports = local_learning.detect(
                screenshot, window_rect, segmentation, directory, progress,
                expanded=bool(navigation_context))
    elif strategy == "visual-first" and _is_desktop_window(window):
        progress("Resolving native desktop items")
        segmentation = _segment_desktop(
            screenshot, window, window_rect, api_url, api_key, model, version,
            max(1, max_controls), progress,
        )
    else:
        progress("Segmenting interface with remote AI")
        segmentation = segment_interface(screenshot, api_url, api_key, model, version, progress)
    if observation_owner is not None:
        _check_foreground(owner)
    elif strategy == "local-visual":
        from .control_observation import _check_foreground
        _check_foreground(window)
    else:
        _activate_window(window)
    refreshed_rect = _window_rect(window)
    if refreshed_rect != capture_bounds:
        raise RuntimeError("Window moved or resized during visual analysis; scan again when stable; "
                           f"captured={capture_bounds}, current={refreshed_rect}, "
                           f"window_handle={getattr(window, 'NativeWindowHandle', 0)}")
    screenshot.save(page_image)
    proposed_page = {**segmentation["page"], "window_identity": identity,
                     "layout_fingerprint": page_identity.layout_fingerprint(segmentation, screenshot.size)}
    page = (proposed_page if manual_layout is not None
            else _reuse_page_identity(directory, screenshot, proposed_page))
    page_id = page["id"]
    if manual_layout is None:
        segmentation["regions"] = _reuse_region_identities(
            directory, page_id, segmentation["regions"], window_rect)
    scan_id = knowledge.stable_id("scan", app_id, page_id, uuid4().hex)
    previous_page_records = [
        record
        for _path, record in knowledge.iter_records(directory, "control")
        if record.get("page_id") == page_id
    ]
    page_digest = _image_sha256(screenshot)
    page_image_named = directory / "images" / "pages" / f"{page_id}-{page_digest}.png"
    if (not page_image_named.is_file()
            or hashlib.sha256(page_image_named.read_bytes()).hexdigest() != page_digest):
        temporary_image = page_image_named.with_suffix(".tmp")
        screenshot.save(temporary_image, format="PNG")
        temporary_image.replace(page_image_named)
    page_record = {
        **page,
        "app_id": app_id,
        "window_name": actual_name,
        "image": str(page_image_named.relative_to(directory)).replace("\\", "/"),
        "model": model,
        "scan_id": scan_id,
        "capture_geometry": "visible-window-bounds+local-edge-crop",
    }
    if manual_layout is not None:
        existing_page = directory / "pages" / f"{knowledge.slugify(page_id)}.md"
        if existing_page.is_file():
            previous_page, _body = knowledge.read_markdown(existing_page)
            page_record = {**previous_page, **page_record}
        page_record.update(region_source="manual", observed_window_size={
            "width": window_rect["width"], "height": window_rect["height"]},
            capture_geometry="visible-window-bounds", window_class=selected["class_name"])
        if strategy == "local-visual":
            page_record["local_annotations"] = local_reports
    knowledge.save_page(directory, page_record)
    previous_regions = {r["id"]: r for _p, r in knowledge.iter_records(directory, "region")} if manual_layout is not None else {}
    for region in segmentation["regions"]:
        region_record = {key: value for key, value in region.items() if key != "rect"}
        knowledge.save_region(
            directory,
            {
                **previous_regions.get(f"{page_id}.{region['id']}", {}),
                **region_record,
                "id": f"{page_id}.{region['id']}",
                "page_id": page_id,
                "app_id": app_id,
                "normalized_rect": knowledge.normalized_rect(
                    region["rect"], window_rect, relative=True
                ),
                "geometry_role": "visual-hint-only",
                **({"source": "manual"} if manual_layout is not None else {}),
            },
        )

    semantic_candidates = []
    control_scopes = {}
    visual_failures = []
    visual_targets = segmentation.get("controls", [])
    effective_strategy = ("local-visual" if strategy == "local-visual" else
                          "manual-regions" if manual_layout is not None else strategy)

    def region_of(rectangle: dict) -> str:
        if manual_layout is not None:
            return layout_rules.assign_region(
                knowledge.normalized_rect(rectangle, window_rect), manual_layout["regions"])[0]
        return _region_for_control(rectangle, window_rect, segmentation["regions"])

    if strategy == "local-visual":
        controls, semantics, visual_failures, local_paths = local_learning.resolve(
            window, window_rect, visual_targets, max_controls, progress)
        if navigation_context:
            from ..library_actions import owned_windows

            allowed = {w["handle"]: w for w in owned_windows(selected)}
            for handle in navigation_context.get("popup_handles", []):
                if handle not in allowed:
                    raise ValueError("展开窗口已失效或不属于所选软件")
                popup = uiautomation.ControlFromHandle(handle)
                if popup is None: continue
                bounds = _window_rect(popup)
                image = _screenshot_window(bounds)
                targets, reports = local_learning.detect(image, bounds, segmentation, directory, progress, expanded=True)
                found, meanings, missing, paths = local_learning.resolve(
                    popup, bounds, targets, max(0,max_controls-len(controls)), progress)
                dx, dy = bounds["left"]-window_rect["left"], bounds["top"]-window_rect["top"]
                for target in targets:
                    r = target["relative_rect"]
                    r.update(left=r["left"]+dx, right=r["right"]+dx, top=r["top"]+dy, bottom=r["bottom"]+dy)
                metadata = allowed[handle]
                scope = {"kind":"owned", "class_name":metadata["class_name"], "title":metadata["title"]}
                for index, (child, depth) in enumerate(found, 1):
                    control_scopes[id(child)] = (popup, scope)
                    semantics[len(controls)+1] = meanings[index]
                    controls.append((child,depth))
                local_paths.update(paths); visual_targets.extend(targets); visual_failures.extend(missing)
                local_frames.append((bounds,image)); local_reports.extend(reports)
        if manual_layout.get("window_scope",{}).get("kind") == "owned":
            for control, _depth in controls:
                control_scopes[id(control)] = (window,manual_layout["window_scope"])
        segmentation["coverage"].update(detector="YOLO+RapidOCR", point_bindings=len(controls),
                                          uia_enumerated=0, local_annotations=local_reports)
    elif strategy == "visual-first":
        if visual_targets:
            progress(f"Resolving {len(visual_targets)} visual targets through local UIA")
            controls, semantics, visual_failures = controls_from_visual_targets(
                window,
                window_rect,
                visual_targets,
                max_controls,
            )
        else:
            controls, semantics = [], {}
            effective_strategy = "visual-first-empty"
            visual_failures.append(
                {
                    "error": (
                        "vision returned no key controls; no full-tree or second-AI fallback "
                        "was run"
                    )
                }
            )
    else:
        progress("Traversing UI Automation controls")
        controls = _walk_controls(window, max(1, min(max_depth, 30)), max(1, max_controls))
        if navigation_context:
            from ..library_actions import owned_windows

            expected = {"handle": selected["handle"], "process_id": selected["process_id"]}
            allowed = {w["handle"]: w for w in owned_windows(expected)}
            seen = {_control_key(c) for c, _d in controls}
            for handle in navigation_context.get("popup_handles", []):
                if handle not in allowed:
                    raise ValueError("展开窗口已失效或不属于所选软件")
                popup = uiautomation.ControlFromHandle(handle)
                if popup is None:
                    continue
                metadata = allowed[handle]
                for child, child_depth in _walk_controls(popup, max_depth, max(1, max_controls-len(controls))):
                    if child is popup or _control_key(child) in seen or len(controls) >= max_controls:
                        continue
                    seen.add(_control_key(child))
                    controls.append((child, child_depth))
                    control_scopes[id(child)] = (popup, {
                        "kind": "owned", "class_name": metadata["class_name"], "title": metadata["title"]})
        if manual_layout is not None:
            segmentation["coverage"]["uia_enumerated"] = len(controls)
            controls = [(control, depth) for control, depth in controls if control is not window]
            if manual_layout.get("restrict_regions"):
                controls = [(control, depth) for control, depth in controls
                            if region_of(_rect(control)) != "unassigned"]
            if manual_layout.get("window_scope", {}).get("kind") == "owned":
                for control, _depth in controls:
                    control_scopes[id(control)] = (window, manual_layout["window_scope"])
        for index, (control, _depth) in enumerate(controls, start=1):
            raw_rect = _rect(control)
            scoped_popup = id(control) in control_scopes
            control_rect = raw_rect if scoped_popup else _clip_rect(raw_rect, window_rect)
            if (bool(getattr(control, "IsOffscreen", False))
                    or control_rect["width"] <= 1 or control_rect["height"] <= 1
                    or (not scoped_popup and not _valid_rect(control_rect, window_rect))):
                continue
            control_type = (getattr(control, "ControlTypeName", "") or "Control")
            supported_actions = _control_actions(
                control_type,
                bool(getattr(control, "IsEnabled", True)),
            )
            uia_name = (getattr(control, "Name", "") or "").strip()
            automation_id = (getattr(control, "AutomationId", "") or "").strip()
            if not supported_actions and not uia_name and not automation_id:
                continue
            semantic_candidates.append(
                {
                    "id": index,
                    "uia_name": uia_name,
                    "automation_id": automation_id,
                    "control_type": control_type,
                    "localized_control_type": str(
                        getattr(control, "LocalizedControlType", "") or ""
                    ).strip(),
                    "class_name": (getattr(control, "ClassName", "") or "").strip(),
                    "help_text": str(getattr(control, "HelpText", "") or "").strip(),
                    "is_keyboard_focusable": bool(
                        getattr(control, "IsKeyboardFocusable", False)
                    ),
                    "is_enabled": bool(getattr(control, "IsEnabled", True)),
                    "region_id": region_of(control_rect),
                    "supported_actions": supported_actions,
                    "relative_rect": _relative_rect(control_rect, window_rect),
                }
            )
        if semantic_mode == "native":
            from .manual_layout import native_semantics

            progress(f"Reading {len(semantic_candidates)} native control labels")
            semantics = {item["id"]: native_semantics(item) for item in semantic_candidates}
        else:
            progress(f"Understanding {len(semantic_candidates)} contextual controls with remote AI")
            semantics = analyze_control_semantics(
                screenshot, semantic_candidates, page, segmentation["regions"],
                api_url, api_key, model, version)
    if strategy == "local-visual" and semantic_mode == "api" and controls:
        candidates=[]
        for index,(control,_depth) in enumerate(controls,1):
            candidates.append(dict(id=index,uia_name=str(getattr(control,"Name","") or ""),
                automation_id=str(getattr(control,"AutomationId","") or ""),control_type=control.ControlTypeName,
                localized_control_type=str(getattr(control,"LocalizedControlType","") or ""),
                class_name=str(getattr(control,"ClassName","") or ""),help_text=str(getattr(control,"HelpText","") or ""),
                is_keyboard_focusable=bool(getattr(control,"IsKeyboardFocusable",False)),
                is_enabled=bool(getattr(control,"IsEnabled",True)),region_id=semantics[index].get("region_id","unassigned"),
                supported_actions=_control_actions(control.ControlTypeName,True),relative_rect=semantics[index]["relative_rect"]))
        progress(f"分析 {len(candidates)} 个控件的语义（不执行交互）")
        meanings=analyze_control_semantics(screenshot,candidates,page,segmentation["regions"],api_url,api_key,model,version)
        for identifier,meaning in meanings.items():semantics[identifier].update(meaning)
        semantic_candidates=candidates
    from .. import reference_library
    reference_description_error = ""
    reference_description_report = {}
    try:
        reference_settings = (dict(api_url=api_url,api_key=api_key,api_model=model)
                              if api_url and api_key and model else None)
        additions = reference_library.descriptions(directory,semantics,progress=progress,settings=reference_settings,
                                                  report=reference_description_report)
        for identifier, fields in additions.items():
            semantics[identifier].update(fields)
        for item in reference_description_report.get("unmatched",[]):
            semantics[item["id"]].update(description_status="unmatched",description_reason=item["reason"])
    except (ValueError, RuntimeError) as error:
        reference_description_error = str(error)
        for semantic in semantics.values():
            semantic.update(description_status="error",description_reason=reference_description_error)
        progress("参考资料说明未能补充，继续保存定位和图片："+str(error))
    verification_screenshot = _screenshot_window(window_rect)
    counts = {"observed": 0, "verified": 0, "suspect": 0, "quarantined": 0}
    key_controls = 0
    in_bounds = 0
    saved_current = 0
    preserved_current = 0
    verified_actions = 0
    failures = list(visual_failures)
    command_counts: dict[str, int] = {}
    observed_ids = set()
    identity_counts: dict[str, int] = {}
    previous_by_id = {record.get("id"): record for record in previous_page_records}
    runtime_blockers = {knowledge.control_anchor(record): record for record in previous_page_records
                        if knowledge.runtime_quarantined(record)
                        and knowledge.control_anchor(record)}
    semantic_counts = {"verified": 0, "uncertain": 0, "manual": 0, "not-required": 0}
    verification_interval = max(1, len(controls) // 20)
    location_prefix_cache: dict = {}
    from . import control_observation

    observation_frames = {}

    def observation_frame(rectangle):
        frame = control_observation.union_bounds(window_rect, rectangle)
        key = tuple(frame[k] for k in ("left", "top", "right", "bottom"))
        if key not in observation_frames:
            image = screenshot if frame == window_rect else _screenshot_window(frame)
            for popup_bounds, popup_image in local_frames:
                if (_clip_rect(rectangle,popup_bounds) == rectangle):
                    image = image.copy()
                    image.paste(popup_image,(popup_bounds["left"]-frame["left"],popup_bounds["top"]-frame["top"]))
                    break
            reference = image_storage.save_interaction_image(directory, image)
            observation_frames[key] = (frame, image, reference)
        return observation_frames[key]
    for index, (control, depth) in enumerate(controls, start=1):
        if index == 1 or index == len(controls) or index % verification_interval == 0:
            progress(f"Verifying controls {index}/{len(controls)}")
        try:
            control_rect = _rect(control)
            scoped_popup = id(control) in control_scopes
            if (bool(getattr(control, "IsOffscreen", False))
                    or control_rect["width"] <= 1 or control_rect["height"] <= 1
                    or (not scoped_popup and not _valid_rect(control_rect, window_rect))):
                continue
            in_bounds += 1
            if not scoped_popup:
                control_rect = _clip_rect(control_rect, window_rect)
            xpath = local_paths.get(id(control)) or get_control_xpath(control)
            location = stable_location(location_from_xpath(xpath))
            control_type = (getattr(control, "ControlTypeName", "") or "Control")
            region_id = (semantics.get(index,{}).get("region_id") if strategy == "local-visual" else None)
            if not region_id or region_id == "unassigned":
                region_id = region_of(control_rect)
            supported_actions = _control_actions(
                control_type,
                bool(getattr(control, "IsEnabled", True)),
            )
            is_key = bool(
                supported_actions
                or getattr(control, "AutomationId", "")
                or getattr(control, "Name", "")
            )
            identity = (
                app_id,
                page_id,
                location["Xpath"],
                control_type,
                getattr(control, "AutomationId", ""),
                "" if getattr(control, "AutomationId", "") else getattr(control, "Name", ""),
            )
            control_id = knowledge.stable_id("control", *identity)
            identity_counts[control_id] = identity_counts.get(control_id, 0) + 1
            if identity_counts[control_id] > 1:
                control_id = knowledge.stable_id("control", *identity, identity_counts[control_id])
            semantic = semantics.get(index, {})
            previous = previous_by_id.get(control_id, {})
            if not previous:
                previous = _previous_control_by_unique_automation_id(
                    previous_page_records,
                    str(getattr(control, "AutomationId", "") or ""),
                    control_type,
                )
                if previous:
                    control_id = previous["id"]
            observed_ids.add(control_id)
            if previous and navigation_context and control_id in navigation_context["path"]:
                # The clicked entry may now be gone while its UIA proxy lingers.
                # Its image was saved before input; never replace it with the result.
                preserved_current += 1
                continue
            if previous.get("window_scope", {}).get("kind") == "owned" and not scoped_popup:
                # A closed native popup cannot be recaptured from the main frame.
                preserved_current += 1
                continue
            if manual_layout is not None and previous.get("location"):
                previous_location = normalize_location(previous["location"])
                previous_mode = previous_location.get("PARAMETERS", {}).get("match_mode")
                if previous_mode in {"identity", "position"}:
                    location = with_match_mode(location, previous_mode)
            if previous.get("semantic_source") == "manual":
                semantic = {
                    "semantic_name": previous.get("semantic_name", ""),
                    "intent": previous.get("intent", ""),
                    "description": previous.get("description", ""),
                    "role": previous.get("semantic_role", "control"),
                    "actions": previous.get("actions", supported_actions),
                    "aliases": previous.get("aliases", []),
                    "risk": previous.get("risk", "unknown"),
                    "requires_confirmation": previous.get("requires_confirmation", False),
                    "confidence": 1.0,
                    "evidence": previous.get("semantic_evidence", ["human teaching"]),
                    "ambiguity": "",
                    "source": "manual",
                    "entity_type": previous.get("entity_type", "ui-element"),
                    "observed_value": previous.get("observed_value", ""),
                    "dynamic_context": previous.get("dynamic_context", False),
                    "current_state": previous.get("current_state", "unknown"),
                    "state_reason": previous.get("state_reason", ""),
                    "enabling_condition": previous.get("enabling_condition", ""),
                    "relationships": previous.get("relationships", []),
                    "visual_components": previous.get("visual_components", []),
                    "description_source": "manual",
                }
            native_identity = semantic_identity.native_metadata(control)
            binding = semantic_identity.check(semantic, native_identity)
            if binding["status"] in {"conflict", "unresolved"}:
                semantic = {**semantic, "ambiguity": f"Native UIA/semantic conflict: {binding}"}
            native_semantic = semantic_identity.literal_semantics(native_identity)
            if (native_semantic and binding["status"] == "consistent"
                    and semantic.get("source") != "manual"
                    and float(semantic.get("confidence", 0)) < SEMANTIC_CONFIDENCE_THRESHOLD):
                prediction = {key: semantic.get(key) for key in (
                    "semantic_name", "intent", "confidence", "source")}
                semantic = {**semantic, **native_semantic, "visual_semantic_prediction": prediction}
            if native_semantic and binding["status"] not in {"conflict", "unresolved"}:
                risk_order = {"unknown": 0, "safe": 1, "state-changing": 2,
                              "external": 3, "destructive": 4}
                if (risk_order.get(native_semantic["risk"], 0)
                        > risk_order.get(semantic.get("risk"), 0)):
                    semantic = {**semantic, "risk": native_semantic["risk"],
                                "requires_confirmation": native_semantic["requires_confirmation"]}
            semantic_name = semantic.get("semantic_name") or _control_semantic_name(control, index)
            intent = semantic.get("intent") or knowledge.slugify(semantic_name, control_id)
            actions = [
                action for action in semantic.get("actions", []) if action in supported_actions
            ]
            semantic_confidence = float(semantic.get("confidence", 0.0))
            semantic_status = "not-required"
            if supported_actions:
                if semantic.get("source") == "manual":
                    semantic_status = "manual"
                elif (
                    semantic_confidence >= SEMANTIC_CONFIDENCE_THRESHOLD
                    and semantic.get("semantic_name")
                    and semantic.get("intent")
                    and actions
                ):
                    semantic_status = "verified"
                else:
                    semantic_status = "uncertain"
            elif semantic.get("source") == "manual":
                semantic_status = "manual"
            elif semantic_confidence >= SEMANTIC_CONFIDENCE_THRESHOLD:
                semantic_status = "verified"
            if binding["status"] in {"conflict", "unresolved"}:
                semantic_status = "uncertain"
            semantic_counts[semantic_status] += 1
            base_command = _command_name(page_id, region_id, intent, control_id)
            command = base_command
            if supported_actions:
                command_counts[base_command] = command_counts.get(base_command, 0) + 1
                if command_counts[base_command] > 1:
                    command = f"{base_command}-{command_counts[base_command]}"
            # Owned dialogs/menus can extend outside the main frame. Keep their
            # complete UIA crop instead of clipping it to the main screenshot.
            observed_bounds, observed_image, observed_reference = observation_frame(control_rect)
            crop = _crop_control(observed_image, control_rect, observed_bounds)
            image_path, image_variants, image_sha256 = _save_control_images(
                directory,
                control_id,
                crop,
                previous,
            )

            location_ok = False
            verification_attempted = False
            location_detail = "not required for structural control"
            similarity = _image_similarity(
                crop,
                _screenshot_window(control_rect) if scoped_popup
                else _crop_control(verification_screenshot, control_rect, window_rect),
            )
            template_result = {"ok": False, "score": 0.0, "detail": "not checked"}
            if supported_actions and verified_actions < max(0, verify_limit):
                verification_attempted = True
                location_ok, verified_rect, location_detail = _verify_location(
                    location,
                    control_rect,
                    control_scopes.get(id(control), (window,))[0],
                    location_prefix_cache,
                )
                if location_ok:
                    verification_crop = (_screenshot_window(verified_rect) if scoped_popup else _crop_control(
                        verification_screenshot, verified_rect, window_rect))
                    similarity = _image_similarity(crop, verification_crop)
                template_result = ({"ok": False, "score": 0.0, "detail": "Scoped popup uses UIA only"}
                                   if scoped_popup else _template_verification(
                    crop,
                    verification_screenshot,
                    verified_rect if location_ok else control_rect,
                    window_rect,
                ))
                verified_actions += 1
            elif supported_actions:
                location_detail = "verification limit reached"
            image_ok = similarity >= 0.75
            if supported_actions:
                if not verification_attempted:
                    status = "suspect"
                else:
                    status = (
                        "verified"
                        if image_ok
                        and (location_ok or template_result["ok"])
                        and semantic_status in {"verified", "manual"}
                        else "quarantined"
                    )
            else:
                status = "observed"
            if status == "quarantined":
                failures.append(
                    {
                        "id": control_id,
                        "name": semantic_name,
                        "location": location_detail,
                        "image_similarity": similarity,
                        "semantic_status": semantic_status,
                        "semantic_confidence": semantic_confidence,
                        "semantic_ambiguity": semantic.get("ambiguity", ""),
                    }
                )
            previous_function = previous.get("function_verification", {})
            function_verification = (
                previous_function
                if previous_function.get("status") == "human-confirmed"
                else {
                    "status": "inferred" if supported_actions else "not-applicable",
                    "method": semantic.get("source", "uia-metadata"),
                    "executed": False,
                    "verified_at": knowledge.utc_now(),
                }
            )
            record = {
                "id": control_id,
                "app_id": app_id,
                "app_name": actual_name,
                "page_id": page_id,
                "region_id": region_id,
                "semantic_name": semantic_name,
                "intent": intent,
                "description": semantic.get("description", ""),
                "semantic_role": semantic.get("role", "control"),
                "semantic_confidence": semantic_confidence,
                "semantic_status": semantic_status,
                "semantic_source": semantic.get("source", "uia-metadata"),
                "semantic_binding": binding,
                "visual_semantic_prediction": semantic.get("visual_semantic_prediction"),
                "semantic_evidence": semantic.get("evidence", []),
                "semantic_ambiguity": semantic.get("ambiguity", ""),
                "entity_type": semantic.get("entity_type", "ui-element"),
                "observed_value": semantic.get("observed_value", ""),
                "dynamic_context": bool(semantic.get("dynamic_context", False)),
                "current_state": (
                    "disabled"
                    if not bool(getattr(control, "IsEnabled", True))
                    else semantic.get("current_state", "enabled")
                ),
                "state_reason": semantic.get("state_reason", ""),
                "enabling_condition": semantic.get("enabling_condition", ""),
                "relationships": semantic.get("relationships", []),
                "visual_components": [
                    _visual_component(item, window_rect)
                    for item in semantic.get("visual_components", [])
                    if isinstance(item, dict)
                ],
                "aliases": semantic.get("aliases", []),
                "risk": semantic.get("risk", "unknown"),
                "requires_confirmation": bool(semantic.get("requires_confirmation", False)),
                "command": command,
                "name": getattr(control, "Name", "") or "",
                "class_name": getattr(control, "ClassName", "") or "",
                "control_type": control_type,
                "automation_id": getattr(control, "AutomationId", "") or "",
                "accelerator_key": native_identity["accelerator_key"],
                "access_key": native_identity["access_key"],
                "framework_id": getattr(control, "FrameworkId", "") or "",
                "localized_control_type": str(
                    getattr(control, "LocalizedControlType", "") or ""
                ),
                "help_text": str(getattr(control, "HelpText", "") or ""),
                "is_keyboard_focusable": bool(
                    getattr(control, "IsKeyboardFocusable", False)
                ),
                "is_enabled_at_scan": bool(getattr(control, "IsEnabled", True)),
                "depth": depth,
                "normalized_rect": knowledge.normalized_rect(control_rect, window_rect),
                "geometry_role": "visual-hint-only",
                "location": location,
                "actions": actions,
                "supported_actions": supported_actions,
                "is_key": is_key,
                "status": status,
                "image": str(image_path.relative_to(directory)).replace("\\", "/"),
                "image_variants": image_variants,
                "image_sha256": image_sha256,
                "image_similarity": similarity,
                "template_score": template_result["score"],
                "visual_fallback_ready": template_result["ok"],
                "verification": {
                    "location": (
                        "passed" if location_ok else "failed" if actions else "not-required"
                    ),
                    "location_detail": location_detail,
                    "image": "passed" if image_ok else "failed" if actions else "observed",
                    "template_detail": template_result["detail"],
                    "verified_at": knowledge.utc_now(),
                },
                "function_verification": function_verification,
                "model": model,
                "tags": list(
                    dict.fromkeys(
                        [
                            control_type,
                            region_id,
                            "key" if is_key else "structural",
                            intent,
                            *semantic.get("aliases", []),
                        ]
                    )
                ),
                "notes": "Generated by scan_window_knowledge.",
                "scan_id": scan_id,
                "scan_strategy": effective_strategy,
            }
            if manual_layout is not None:
                record["region_source"] = "manual"
                record["region_assignment"] = layout_rules.assign_region(
                    record["normalized_rect"], manual_layout["regions"])[1]
                record["visual_observation"] = control_observation.describe(
                    observed_reference, observed_bounds, control_rect, window_rect, phase="located")
                record["visual_observation"]["purpose"] = "learning-reference"
                if strategy == "local-visual":
                    record["detection"] = {"method":"local-yolo-rapidocr+get_control_info",
                        "analysis_result":semantics[index].get("analysis_result", ""),
                        "normalized_rect":knowledge.normalized_rect(semantics[index]["relative_rect"],window_rect,relative=True)}
            if navigation_context and (_control_key(control) not in navigation_context.get("before_keys", set())
                    and control_id not in navigation_context.get("before_ids", set())):
                path = navigation_context["path"]
                if control_id not in path:
                    record.update(navigation_path=list(path), navigation_source="expanded-learning",
                                  window_scope=control_scopes.get(id(control), (None, {"kind": "main"}))[1])
            for key in ("description_source","description_status","description_reason","description_method","description_basis","reference_evidence"):
                if key in semantic: record[key]=semantic[key]
            record = knowledge.merge_control_state(previous, record)
            blocker = runtime_blockers.get(knowledge.control_anchor(record))
            if blocker:
                record = knowledge.merge_control_state(blocker, record)
            status = record["status"]
            knowledge.save_control(directory, record)
            counts[status] += 1
            key_controls += int(is_key)
            saved_current += 1
        except Exception as error:
            counts["suspect"] += 1
            failures.append({"index": index, "error": f"{type(error).__name__}: {error}"})

    visual_observations = 0
    visual_identity_counts: dict[str, int] = {}
    for failure in visual_failures:
        target = failure.get("observation")
        if not isinstance(target, dict) or not isinstance(target.get("relative_rect"), dict):
            continue
        rectangle = target["relative_rect"]
        semantic_name = target.get("semantic_name") or "Visual element"
        intent = target.get("intent") or knowledge.slugify(semantic_name, "visual-element")
        region_id = target.get("region_id", "unassigned")
        control_id = knowledge.stable_id(
            "visual",
            app_id,
            page_id,
            region_id,
            intent,
            target.get("observed_value", ""),
        )
        visual_identity_counts[control_id] = visual_identity_counts.get(control_id, 0) + 1
        if visual_identity_counts[control_id] > 1:
            control_id = knowledge.stable_id(
                "visual-instance", control_id, visual_identity_counts[control_id]
            )
        observed_ids.add(control_id)
        absolute = {**rectangle, "left":window_rect["left"]+rectangle["left"],
                    "right":window_rect["left"]+rectangle["right"],
                    "top":window_rect["top"]+rectangle["top"], "bottom":window_rect["top"]+rectangle["bottom"]}
        bounds, image, reference = observation_frame(absolute)
        crop = _crop_control(image, absolute, bounds)
        previous = previous_by_id.get(control_id, {})
        image_path, image_variants, image_sha256 = _save_control_images(
            directory, control_id, crop, previous
        )
        confidence = float(target.get("confidence", 0.0))
        semantic_status = (
            "verified" if confidence >= SEMANTIC_CONFIDENCE_THRESHOLD else "uncertain"
        )
        visual_record = {
                "id": control_id,
                "app_id": app_id,
                "app_name": actual_name,
                "page_id": page_id,
                "region_id": region_id,
                "semantic_name": semantic_name,
                "intent": intent,
                "description": target.get("description", ""),
                "semantic_role": target.get("role", "visual-entity"),
                "semantic_confidence": confidence,
                "semantic_status": semantic_status,
                "semantic_source": target.get("source", "ai-vision-target"),
                "semantic_evidence": target.get("evidence", []),
                "semantic_ambiguity": target.get("ambiguity", ""),
                "aliases": target.get("aliases", []),
                **_semantic_context(target),
                "risk": target.get("risk", "unknown"),
                "requires_confirmation": False,
                "command": "",
                "name": "",
                "class_name": "",
                "control_type": "VisualElement",
                "automation_id": "",
                "depth": 0,
                "normalized_rect": knowledge.normalized_rect(
                    rectangle, window_rect, relative=True
                ),
                "geometry_role": "visual-hint-only",
                "location": ({**location_from_xpath([], parameters={"region_id":region_id,
                    "normalized_rect":knowledge.normalized_rect(rectangle,window_rect,relative=True),
                    "locator_kind":"visual-observation"}), "WindowName":actual_name,
                    "Img":str(image_path.relative_to(directory)).replace("\\", "/")}
                    if strategy == "local-visual" else {}),
                **({"visual_observation":{**control_observation.describe(reference,bounds,absolute,window_rect,phase="located"),
                      "purpose":"learning-reference"}, "analysis_result":target.get("analysis_result", "")}
                   if strategy == "local-visual" else {}),
                "actions": [],
                "supported_actions": [],
                "is_key": True,
                "status": "observed",
                "image": str(image_path.relative_to(directory)).replace("\\", "/"),
                "image_variants": image_variants,
                "image_sha256": image_sha256,
                "verification": {
                    "location": "not-available",
                    "image": "observed",
                    "detail": failure.get("error", "visual-only observation"),
                    "reason": failure.get("reason", ""),
                    "diagnostics": failure.get("diagnostics", {}),
                    "verified_at": knowledge.utc_now(),
                },
                "function_verification": {
                    "status": "not-applicable",
                    "method": "visual-observation",
                    "executed": False,
                    "verified_at": knowledge.utc_now(),
                },
                "model": model,
                "tags": ["visual-only", region_id, intent],
                "notes": "Visual entity retained even though no distinct UIA node resolved.",
                "scan_id": scan_id,
                "scan_strategy": effective_strategy,
            }
        if strategy == "local-visual":
            if previous.get("semantic_source") == "manual":
                for key in ("semantic_name","intent","description","aliases","risk","semantic_source",
                            "semantic_confidence","semantic_status"):
                    if key in previous: visual_record[key]=previous[key]
            visual_record = knowledge.merge_control_state(previous,visual_record)
        knowledge.save_control(directory,visual_record)
        counts["observed"] += 1
        semantic_counts[semantic_status] += 1
        key_controls += 1
        saved_current += 1
        visual_observations += 1

    if (controls or visual_observations) and saved_current == 0 and preserved_current == 0:
        raise RuntimeError(
            "UIA returned controls but none remained inside the refreshed window bounds; "
            "the window likely moved or changed during scanning"
        )

    retained_unobserved = 0
    archived_layout_records = 0
    if strategy == "local-visual" and saved_current:
        archived_layout_records = local_learning.archive_layout_records(directory,previous_page_records,observed_ids,
            {r["id"] for r in manual_layout["regions"]},scan_id)
        if navigation_context:
            local_learning.expanded_regions(directory,controls,navigation_context,window_rect,
                                            observation_frame,app_id,page_id)
    for previous in previous_page_records:
        if previous.get("id") in observed_ids:
            continue
        if effective_strategy in {"visual-first", "local-visual"} or previous.get("navigation_path") or navigation_context:
            retained_unobserved += 1
            continue
        if manual_layout is not None and manual_layout.get("restrict_regions") and previous.get("region_id") not in {
                r["id"] for r in manual_layout["regions"]}:
            retained_unobserved += 1
            continue
        previous["status"] = "quarantined"
        previous["notes"] = (
            f"{previous.get('notes', '')}\nQuarantined because the control was not present "
            f"in rescan {scan_id}."
        ).strip()
        previous["verification"] = {
            **previous.get("verification", {}),
            "rescan": "missing",
            "verified_at": knowledge.utc_now(),
        }
        knowledge.save_control(directory, previous)
        counts["quarantined"] += 1

    progress("Saving learned controls and annotated page image")
    scanned_records = [
        record
        for _path, record in knowledge.iter_records(directory, "control")
        if record.get("scan_id") == scan_id
    ]
    if manual_layout is not None and any(r.get("region_id") == "unassigned" for r in scanned_records):
        knowledge.save_region(directory, {
            "id": f"{page_id}.unassigned", "page_id": page_id, "app_id": app_id,
            "name": "待归档", "description": "未被区域覆盖或跨区归属不明确的控件",
            "source": "manual", "classification_only": True,
            "geometry_role": "visual-hint-only",
        })
    if manual_layout is not None:
        from .dynamic_regions import synchronize

        synchronize(directory, [r for _p, r in knowledge.iter_records(directory, "control")],
                    {r["id"] for r in scanned_records})
    annotated_path, annotation_markers = visualization.save_scan_annotation(
        directory,
        page_id,
        screenshot,
        window_rect,
        scanned_records,
    )
    page_record["annotated_image"] = str(annotated_path.relative_to(directory)).replace(
        "\\", "/"
    )
    knowledge.save_page(directory, page_record)
    overlay = {"shown": False, "controls": 0, "detail": "disabled"}
    if show_overlay:
        live_markers = visualization.control_markers(
            scanned_records,
            window_rect,
            window_rect,
        )
        overlay = visualization.show_markers(live_markers, overlay_duration_ms)

    progress("Pruning unreferenced knowledge images")
    try:
        storage_result = image_storage.cleanup_images(directory, dry_run=False)
    except (OSError, ValueError) as error:
        storage_result = {"ok": False, "detail": f"Image cleanup deferred: {error}"}
    progress("Finalizing knowledge index and command catalog")
    index = knowledge.rebuild_index(directory)
    command_items = knowledge.available_commands(directory)
    catalog_path = knowledge.write_command_catalog(directory)
    elapsed = round(time.perf_counter() - started, 2)
    return {
        "ok": True,
        "app_id": app_id,
        "app_name": actual_name,
        "page_id": page_id,
        "scan_id": scan_id,
        "observed_control_ids": [record["id"] for record in scanned_records],
        "vault": str(directory),
        "page_image": str(page_image_named),
        "uia_access": access,
        "page_identity": {key: page[key] for key in ("identity_decision", "previous_ai_id", "identity_similarity")
                          if key in page},
        "annotated_page_image": str(annotated_path),
        "annotated_controls": len(annotation_markers),
        "overlay": overlay,
        "regions": len(segmentation["regions"]),
        "strategy": effective_strategy,
        "reference_description_error": reference_description_error,
        "reference_descriptions": reference_description_report,
        "archived_layout_records": archived_layout_records,
        "visual_targets": len(visual_targets),
        "coverage": segmentation.get("coverage", {}),
        **({"unassigned_controls": sum(r.get("region_id") == "unassigned" for r in scanned_records),
            "region_source": "manual", "semantic_mode": semantic_mode}
           if manual_layout is not None else {}),
        "visual_only_observations": visual_observations,
        "uia_controls_seen": len(controls),
        "controls_in_bounds": in_bounds,
        "controls_saved_this_scan": saved_current,
        "controls_retained_unobserved": retained_unobserved,
        "knowledge_controls_total": len(index["controls"]),
        "key_controls": key_controls,
        "commands": len(command_items),
        "command_items": command_items,
        "quarantine_summary": [
            {
                "control_id": item.get("id"),
                "semantic_name": item.get("semantic_name") or item.get("name"),
                "semantic_ambiguity": item.get("semantic_ambiguity", ""),
                "location_detail": item.get("verification", {}).get("location_detail", ""),
                "runtime_quarantine": knowledge.runtime_quarantined(item),
                "runtime_failure": item.get("quarantine", {}).get("reason") or item.get(
                    "verification", {}).get("runtime_detail", ""),
            }
            for item in index["controls"]
            if item.get("status") == "quarantined"
        ],
        "command_catalog": str(catalog_path),
        "status_counts": counts,
        "semantic_counts": semantic_counts,
        "semantic_controls_analyzed": (
            len(visual_targets)
            if effective_strategy == "visual-first"
            else len(semantic_candidates)
        ),
        "failures": failures[:50],
        "truncated": (len(controls) >= max_controls or len(visual_targets) > max_controls
                      or any(report.get("truncated") for report in local_reports)
                      or segmentation.get("coverage", {}).get("uia_enumerated", 0) >= max_controls
                      or segmentation.get("coverage", {}).get("desktop_limit_reached", False)),
        "elapsed_seconds": elapsed,
        "stage_timings": stage_timings,
        "image_storage": storage_result,
    }
