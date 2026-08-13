"""Scan a Windows application into an Obsidian-compatible UI knowledge vault."""

from __future__ import annotations

import base64
import ctypes
import hashlib
import io
import json
import re
import time
from collections.abc import Callable
from ctypes import wintypes
from pathlib import Path
from urllib import error as urlerror
from urllib import request as urlrequest

import pyautogui
import uiautomation

from easy_uiauto.utils import get_control_xpath

from . import knowledge, visualization
from .protocol import location_from_xpath, normalize_location, stable_location

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
SCAN_STRATEGIES = {"visual-first", "full-uia"}
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


def _find_window(title: str):
    root = uiautomation.GetRootControl()
    candidates = []
    for control in root.GetChildren():
        name = (getattr(control, "Name", "") or "").strip()
        if not name:
            continue
        if name == title:
            return control
        if title.casefold() in name.casefold():
            candidates.append(control)
    if not candidates:
        raise RuntimeError(f"Window not found: {title}")
    return candidates[0]


def _activate_window(window) -> None:
    """Activate a UIA window without relying on pygetwindow's stale handles."""
    handle = int(getattr(window, "NativeWindowHandle", 0) or 0)
    if not handle or not ctypes.windll.user32.IsWindow(handle):
        raise RuntimeError("UI Automation returned an invalid window handle")
    ctypes.windll.user32.ShowWindow(handle, 9)  # SW_RESTORE
    if not ctypes.windll.user32.SetForegroundWindow(handle):
        try:
            window.SetFocus()
        except Exception as error:
            raise RuntimeError("Could not activate the target window") from error
    time.sleep(0.2)


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


def _descendant_matches(parent, step: dict, max_depth: int) -> list:
    matches = []
    queue = [(parent, 0)]
    while queue:
        current, depth = queue.pop(0)
        if depth >= max_depth:
            continue
        try:
            children = current.GetChildren()
        except Exception:
            children = []
        for child in children:
            if _matches_step(child, step):
                matches.append(child)
            queue.append((child, depth + 1))
    return matches


def _automation_id_step(location: dict) -> dict | None:
    xpath = location.get("Xpath", [])
    for step in reversed(xpath):
        if step.get("AutomationId"):
            return step
    if location.get("AutomationId"):
        return {
            "AutomationId": location["AutomationId"],
            "ControlType": location.get("ControlType", ""),
            "ClassName": location.get("ClassName", ""),
            "foundIndex": location.get("foundIndex", ""),
        }
    return None


def _resolve_unique_automation_id(window, location: dict):
    """Use UIA's native property search when an AutomationId is unique in the window."""
    step = _automation_id_step(location)
    if not step:
        return None
    control_type = getattr(uiautomation.ControlType, step.get("ControlType", ""), None)
    search_depth = max(1, min(30, int(step.get("searchDepth", 12) or 12)))
    query = {
        "AutomationId": step["AutomationId"],
        "searchDepth": search_depth,
        "searchInterval": 0.05,
    }
    if control_type is not None:
        query["ControlType"] = control_type
    if step.get("ClassName"):
        query["ClassName"] = step["ClassName"]
    try:
        first = uiautomation.Control(searchFromControl=window, foundIndex=1, **query)
        if not first.Exists(0.25, 0.05):
            return None
        second = uiautomation.Control(searchFromControl=window, foundIndex=2, **query)
        if second.Exists(0.15, 0.05):
            return None
        return first
    except Exception:
        return None


def resolve_location(location: dict, window=None, prefix_cache: dict | None = None):
    """Resolve a canonical LOCATION through UIA without logging or window side effects."""
    normalized = normalize_location(location)
    xpath = normalized.get("Xpath", [])
    window_name = normalized.get("WindowName", "")
    if not window_name and xpath:
        window_name = xpath[0].get("Name", "")
    if window is None:
        window = _find_window(window_name)
    current = window
    prefix_cache = prefix_cache if prefix_cache is not None else {}
    automation_key = (
        "automation-id",
        normalized.get("AutomationId", ""),
        normalized.get("ControlType", ""),
        normalized.get("ClassName", ""),
    )
    if automation_key in prefix_cache:
        return prefix_cache[automation_key]
    fast_control = _resolve_unique_automation_id(window, normalized)
    if fast_control is not None:
        prefix_cache[automation_key] = fast_control
        return fast_control
    previous_search_depth = 1
    steps = xpath[1:] if xpath and _matches_step(window, xpath[0]) else xpath
    prefix = []
    for step in steps:
        prefix.append(json.dumps(step, ensure_ascii=False, sort_keys=True))
        cache_key = tuple(prefix)
        cached = prefix_cache.get(cache_key)
        if cached is not None:
            current = cached
            previous_search_depth = int(step.get("searchDepth", previous_search_depth + 1) or 1)
            continue
        search_depth = int(step.get("searchDepth", previous_search_depth + 1) or 1)
        relative_depth = max(1, min(12, search_depth - previous_search_depth))
        candidates = _descendant_matches(current, step, relative_depth)
        if not candidates and step.get("AutomationId") and step.get("Name"):
            relaxed_step = {key: value for key, value in step.items() if key != "Name"}
            candidates = _descendant_matches(current, relaxed_step, relative_depth)
        if not candidates and relative_depth < 4:
            candidates = _descendant_matches(current, step, 4)
        if (
            not candidates
            and relative_depth < 4
            and step.get("AutomationId")
            and step.get("Name")
        ):
            relaxed_step = {key: value for key, value in step.items() if key != "Name"}
            candidates = _descendant_matches(current, relaxed_step, 4)
        if not candidates:
            return None
        found_index = step.get("foundIndex")
        try:
            candidate_index = max(0, int(found_index) - 1) if found_index not in (None, "") else 0
        except (TypeError, ValueError):
            candidate_index = 0
        current = candidates[min(candidate_index, len(candidates) - 1)]
        prefix_cache[cache_key] = current
        previous_search_depth = search_depth
    return current


def _screenshot_window(window_rect: dict):
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
    """Keep a stable page ID when AI naming changes for the same visual page."""
    from PIL import Image

    best = None
    for _path, previous in knowledge.iter_records(directory, "page"):
        image_name = previous.get("image", "")
        image_path = directory / image_name if image_name else None
        if image_path is None or not image_path.is_file():
            continue
        try:
            with Image.open(image_path) as previous_image:
                similarity = _image_similarity(previous_image, screenshot)
        except OSError:
            continue
        if best is None or similarity > best[0]:
            best = (similarity, previous)
    if best is not None and best[0] >= 0.94:
        previous = best[1]
        return {
            **page,
            "id": previous["id"],
            "name": previous.get("name") or page.get("name"),
            "previous_ai_id": page.get("id"),
            "identity_similarity": best[0],
        }
    return page


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
    variants.append(str(new_variant.relative_to(directory)).replace("\\", "/"))
    return image_path, list(dict.fromkeys(variants))[-8:], image_hash


def _parse_json_object(content: str) -> dict:
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content, flags=re.IGNORECASE)
    result = json.loads(content)
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
        content = body["choices"][0]["message"]["content"]
        if isinstance(content, list):
            return "".join(
                item.get("text", "") for item in content if isinstance(item, dict)
            )
        return str(content)
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
        choices = chunk.get("choices", [])
        if not choices:
            continue
        choice = choices[0]
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


def segment_interface(
    screenshot,
    api_url: str,
    api_key: str,
    model: str,
    version: str,
) -> dict:
    """Ask the multimodal endpoint for page, regions, and key visual controls."""
    buffer = io.BytesIO()
    screenshot.save(buffer, format="PNG")
    image_url = "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")
    width, height = screenshot.size
    prompt = (
        f"Analyze this {width}x{height} desktop application screenshot. "
        "Identify the current page, divide it into functional regions, and identify only controls "
        "that a user would need to operate or read to complete tasks. Exclude decorative, "
        "duplicate, layout-only, and inaccessible background elements. Use concise generic IDs in "
        "kebab-case. Return JSON only with this shape: "
        '{"page":{"id":"main","name":"Main","description":""},'
        '"regions":[{"id":"navigation","name":"Navigation","role":"navigation",'
        '"description":"","left":0,"top":0,"width":100,"height":100}],'
        '"controls":[{"id":1,"region":"navigation","semantic_name":"Open menu",'
        '"intent":"open-navigation-menu","description":"Open the application navigation",'
        '"role":"action","entity_type":"navigation-action","observed_value":"",'
        '"dynamic_context":false,"current_state":"enabled","state_reason":"",'
        '"enabling_condition":"","relationships":[],"actions":["click"],'
        '"aliases":["menu"],"risk":"safe",'
        '"requires_confirmation":false,"confidence":0.97,"evidence":["menu icon"],'
        '"ambiguity":"","left":10,"top":10,"width":40,"height":40}]}. '
        "Coordinates must use screenshot pixels. Include all major functional areas, "
        "not each control. Control rectangles must tightly cover the visible target. Include key "
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
    payload = {
        "model": model,
        "temperature": 0,
        "stream": True,
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
            content = _read_completion_content(response)
    except urlerror.HTTPError as error:
        raise RuntimeError(f"AI segmentation failed with HTTP {error.code}") from error
    except (urlerror.URLError, TimeoutError) as error:
        reason = getattr(error, "reason", str(error))
        raise RuntimeError(f"AI segmentation failed: {reason}") from error
    result = _parse_json_object(str(content))
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
):
    """Choose the smallest stable UIA control represented by a visual target."""
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
    candidates = {}
    for x, y in _visual_sample_points(target, window_rect):
        try:
            current = point_lookup(x, y)
        except Exception:
            continue
        for _depth in range(16):
            if not current:
                break
            try:
                candidates.setdefault(_control_key(current), current)
                if current is window or getattr(current, "ControlTypeName", "") == "DesktopControl":
                    break
                current = current.GetParentControl()
            except Exception:
                break
    if not candidates:
        return None
    ranked = sorted(
        candidates.values(),
        key=lambda control: _candidate_score(control, target, absolute_target, window),
        reverse=True,
    )
    selected = ranked[0]
    if _candidate_score(selected, target, absolute_target, window)[0] < 0:
        return None
    return selected


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
    seen: dict[str, int] = {}
    for target in targets[: max(1, max_controls)]:
        control = control_from_visual_target(window, window_rect, target)
        if control is None:
            failures.append(
                {
                    "visual_target": target.get("semantic_name") or target.get("id"),
                    "error": "no UIA control at visual target",
                    "observation": target,
                }
            )
            continue
        try:
            xpath = get_control_xpath(control)
            key = json.dumps(xpath, ensure_ascii=False, sort_keys=True)
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
            components = list(primary.get("visual_components", []))
            if target.get("actions") and not primary.get("actions"):
                components.append(primary)
                semantics[index] = {**target, "visual_components": components}
            else:
                components.append(target)
                primary["visual_components"] = components
            continue
        index = len(controls) + 1
        seen[key] = index
        controls.append((control, max(0, len(xpath) - 1)))
        semantics[index] = {**target, "visual_components": []}
    return controls, semantics, failures


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
    window = _find_window(window_name)
    actual_name = (getattr(window, "Name", "") or window_name).strip()
    _activate_window(window)
    window_rect = _rect(window)
    if window_rect["width"] <= 1 or window_rect["height"] <= 1:
        raise RuntimeError(f"Window has no visible bounds: {actual_name}")

    process_id = int(getattr(window, "ProcessId", 0) or 0)
    process_name = _process_name(process_id)
    app_id = knowledge.slugify(process_name or actual_name, "application")
    directory = knowledge.initialize_app(app_id, actual_name, root)
    knowledge.load_index(directory)
    progress("Capturing window screenshot")
    screenshot = _screenshot_window(window_rect)
    page_image = directory / "images" / "pages" / "latest.png"
    screenshot.save(page_image)

    progress("Segmenting interface with remote AI")
    segmentation = segment_interface(screenshot, api_url, api_key, model, version)
    _activate_window(window)
    refreshed_rect = _rect(window)
    if refreshed_rect != window_rect:
        window_rect = refreshed_rect
        screenshot = _screenshot_window(window_rect)
    page = _reuse_page_identity(directory, screenshot, segmentation["page"])
    page_id = page["id"]
    segmentation["regions"] = _reuse_region_identities(
        directory,
        page_id,
        segmentation["regions"],
        window_rect,
    )
    scan_id = knowledge.stable_id("scan", app_id, page_id, knowledge.utc_now())
    previous_page_records = [
        record
        for _path, record in knowledge.iter_records(directory, "control")
        if record.get("page_id") == page_id
    ]
    page_image_named = directory / "images" / "pages" / f"{page_id}.png"
    screenshot.save(page_image_named)
    page_record = {
        **page,
        "app_id": app_id,
        "window_name": actual_name,
        "image": str(page_image_named.relative_to(directory)).replace("\\", "/"),
        "model": model,
    }
    knowledge.save_page(directory, page_record)
    for region in segmentation["regions"]:
        region_record = {key: value for key, value in region.items() if key != "rect"}
        knowledge.save_region(
            directory,
            {
                **region_record,
                "id": f"{page_id}.{region['id']}",
                "page_id": page_id,
                "app_id": app_id,
                "normalized_rect": knowledge.normalized_rect(
                    region["rect"], window_rect, relative=True
                ),
                "geometry_role": "visual-hint-only",
            },
        )

    semantic_candidates = []
    visual_failures = []
    visual_targets = segmentation.get("controls", [])
    effective_strategy = strategy
    if strategy == "visual-first":
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
        for index, (control, _depth) in enumerate(controls, start=1):
            control_rect = _clip_rect(_rect(control), window_rect)
            if not _valid_rect(control_rect, window_rect):
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
                    "region_id": _region_for_control(
                        control_rect,
                        window_rect,
                        segmentation["regions"],
                    ),
                    "supported_actions": supported_actions,
                    "relative_rect": _relative_rect(control_rect, window_rect),
                }
            )
        progress(
            f"Understanding {len(semantic_candidates)} contextual controls with remote AI"
        )
        semantics = analyze_control_semantics(
            screenshot,
            semantic_candidates,
            page,
            segmentation["regions"],
            api_url,
            api_key,
            model,
            version,
        )
    verification_screenshot = _screenshot_window(window_rect)
    counts = {"observed": 0, "verified": 0, "suspect": 0, "quarantined": 0}
    key_controls = 0
    in_bounds = 0
    saved_current = 0
    verified_actions = 0
    failures = list(visual_failures)
    command_counts: dict[str, int] = {}
    observed_ids = set()
    previous_by_id = {record.get("id"): record for record in previous_page_records}
    semantic_counts = {"verified": 0, "uncertain": 0, "manual": 0, "not-required": 0}
    verification_interval = max(1, len(controls) // 20)
    location_prefix_cache: dict = {}
    for index, (control, depth) in enumerate(controls, start=1):
        if index == 1 or index == len(controls) or index % verification_interval == 0:
            progress(f"Verifying controls {index}/{len(controls)}")
        try:
            control_rect = _rect(control)
            if not _valid_rect(control_rect, window_rect):
                continue
            in_bounds += 1
            control_rect = _clip_rect(control_rect, window_rect)
            xpath = get_control_xpath(control)
            location = stable_location(location_from_xpath(xpath))
            control_type = (getattr(control, "ControlTypeName", "") or "Control")
            region_id = _region_for_control(control_rect, window_rect, segmentation["regions"])
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
                }
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
            semantic_counts[semantic_status] += 1
            base_command = _command_name(page_id, region_id, intent, control_id)
            command = base_command
            if supported_actions:
                command_counts[base_command] = command_counts.get(base_command, 0) + 1
                if command_counts[base_command] > 1:
                    command = f"{base_command}-{command_counts[base_command]}"
            crop = _crop_control(screenshot, control_rect, window_rect)
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
                _crop_control(verification_screenshot, control_rect, window_rect),
            )
            template_result = {"ok": False, "score": 0.0, "detail": "not checked"}
            if supported_actions and verified_actions < max(0, verify_limit):
                verification_attempted = True
                location_ok, verified_rect, location_detail = _verify_location(
                    location,
                    control_rect,
                    window,
                    location_prefix_cache,
                )
                if location_ok:
                    verification_crop = _crop_control(
                        verification_screenshot,
                        verified_rect,
                        window_rect,
                    )
                    similarity = _image_similarity(crop, verification_crop)
                template_result = _template_verification(
                    crop,
                    verification_screenshot,
                    verified_rect if location_ok else control_rect,
                    window_rect,
                )
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
            knowledge.save_control(directory, record)
            counts[status] += 1
            key_controls += int(is_key)
            saved_current += 1
        except Exception as error:
            counts["suspect"] += 1
            failures.append({"index": index, "error": f"{type(error).__name__}: {error}"})

    visual_observations = 0
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
        observed_ids.add(control_id)
        crop = screenshot.crop(
            (
                rectangle["left"],
                rectangle["top"],
                rectangle["right"],
                rectangle["bottom"],
            )
        )
        previous = previous_by_id.get(control_id, {})
        image_path, image_variants, image_sha256 = _save_control_images(
            directory, control_id, crop, previous
        )
        confidence = float(target.get("confidence", 0.0))
        semantic_status = (
            "verified" if confidence >= SEMANTIC_CONFIDENCE_THRESHOLD else "uncertain"
        )
        knowledge.save_control(
            directory,
            {
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
                "location": {},
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
            },
        )
        counts["observed"] += 1
        semantic_counts[semantic_status] += 1
        key_controls += 1
        saved_current += 1
        visual_observations += 1

    if (controls or visual_observations) and saved_current == 0:
        raise RuntimeError(
            "UIA returned controls but none remained inside the refreshed window bounds; "
            "the window likely moved or changed during scanning"
        )

    retained_unobserved = 0
    for previous in previous_page_records:
        if previous.get("id") in observed_ids:
            continue
        if effective_strategy == "visual-first":
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
        "vault": str(directory),
        "page_image": str(page_image_named),
        "annotated_page_image": str(annotated_path),
        "annotated_controls": len(annotation_markers),
        "overlay": overlay,
        "regions": len(segmentation["regions"]),
        "strategy": effective_strategy,
        "visual_targets": len(visual_targets),
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
        "truncated": len(controls) >= max_controls,
        "elapsed_seconds": elapsed,
        "stage_timings": stage_timings,
    }
