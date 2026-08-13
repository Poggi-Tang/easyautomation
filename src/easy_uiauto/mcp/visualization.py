"""Fast visual feedback for learned and executable UI controls."""

from __future__ import annotations

from pathlib import Path

from easy_uiauto.draw import show_control_overlay

from . import knowledge

STATUS_COLORS = {
    "verified": "#00c853",
    "observed": "#00b8d4",
    "suspect": "#ffab00",
    "quarantined": "#ff6d00",
}
PAGE_MATCH_THRESHOLD = 0.82


def _rect_size(rectangle: dict) -> tuple[int, int]:
    width = rectangle.get("width")
    if width is None:
        width = rectangle["right"] - rectangle["left"]
    height = rectangle.get("height")
    if height is None:
        height = rectangle["bottom"] - rectangle["top"]
    return max(1, int(width)), max(1, int(height))


def _scaled_rect(rectangle: dict, source: dict, destination: dict) -> dict:
    source_width, source_height = _rect_size(source)
    destination_width, destination_height = _rect_size(destination)
    scale_x = destination_width / source_width
    scale_y = destination_height / source_height
    left = destination["left"] + round((rectangle["left"] - source["left"]) * scale_x)
    top = destination["top"] + round((rectangle["top"] - source["top"]) * scale_y)
    right = destination["left"] + round((rectangle["right"] - source["left"]) * scale_x)
    bottom = destination["top"] + round((rectangle["bottom"] - source["top"]) * scale_y)
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "width": max(0, right - left),
        "height": max(0, bottom - top),
    }


def _rect_from_normalized(rectangle: dict, destination: dict) -> dict:
    destination_width, destination_height = _rect_size(destination)
    left = destination["left"] + round(float(rectangle["left"]) * destination_width)
    top = destination["top"] + round(float(rectangle["top"]) * destination_height)
    right = destination["left"] + round(float(rectangle["right"]) * destination_width)
    bottom = destination["top"] + round(float(rectangle["bottom"]) * destination_height)
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "width": max(0, right - left),
        "height": max(0, bottom - top),
    }


def control_markers(
    records: list[dict],
    source_window_rect: dict | None,
    destination_window_rect: dict,
    target_ids: set[str] | None = None,
) -> list[dict]:
    """Build markers from stable normalized hints, with legacy-pixel compatibility."""
    target_ids = target_ids or set()
    sorted_records = sorted(
        records,
        key=lambda item: (
            item.get("normalized_rect", item.get("rect", {})).get("top", 0),
            item.get("normalized_rect", item.get("rect", {})).get("left", 0),
            item.get("semantic_name", item.get("id", "")),
        ),
    )
    markers = []
    for record in sorted_records:
        try:
            normalized = record.get("normalized_rect")
            if isinstance(normalized, dict):
                current_rect = _rect_from_normalized(normalized, destination_window_rect)
            elif isinstance(record.get("rect"), dict) and isinstance(source_window_rect, dict):
                current_rect = _scaled_rect(
                    record["rect"], source_window_rect, destination_window_rect
                )
            else:
                continue
        except (KeyError, TypeError, ValueError):
            continue
        if current_rect["width"] <= 1 or current_rect["height"] <= 1:
            continue
        markers.append(
            {
                "index": len(markers) + 1,
                "control_id": record.get("id", ""),
                "label": record.get("semantic_name") or record.get("name") or record.get("id"),
                "rect": current_rect,
                "status": record.get("status", "observed"),
                "color": STATUS_COLORS.get(record.get("status", ""), "#00b8d4"),
                "target": record.get("id") in target_ids,
            }
        )
    return markers


def annotate_screenshot(screenshot, markers: list[dict]):
    """Return a numbered page image matching the live overlay legend."""
    from PIL import ImageDraw

    annotated = screenshot.convert("RGB").copy()
    draw = ImageDraw.Draw(annotated)
    for marker in markers:
        rectangle = marker["rect"]
        color = "#ff1744" if marker.get("target") else marker["color"]
        left = rectangle["left"]
        top = rectangle["top"]
        right = rectangle["right"]
        bottom = rectangle["bottom"]
        draw.rectangle((left, top, right, bottom), outline=color, width=3)
        label = str(marker["index"])
        label_top = max(0, top - 18)
        label_width = max(18, 8 * len(label) + 8)
        draw.rectangle((left, label_top, left + label_width, label_top + 18), fill=color)
        draw.text((left + 4, label_top + 2), label, fill="white")
    return annotated


def save_scan_annotation(
    directory: Path,
    page_id: str,
    screenshot,
    window_rect: dict,
    records: list[dict],
) -> tuple[Path, list[dict]]:
    relative_window = {
        "left": 0,
        "top": 0,
        "right": screenshot.width,
        "bottom": screenshot.height,
        "width": screenshot.width,
        "height": screenshot.height,
    }
    markers = control_markers(records, window_rect, relative_window)
    path = directory / "images" / "pages" / f"{page_id}.annotated.png"
    annotate_screenshot(screenshot, markers).save(path)
    return path, markers


def show_markers(markers: list[dict], duration_ms: int = 3000, wait_ms: int = 100) -> dict:
    return show_control_overlay(
        markers,
        show_time=max(300, min(int(duration_ms), 30000)),
        wait_ms=max(0, min(int(wait_ms), 1000)),
    )


def _match_current_page(directory: Path, requested_page_id: str = "") -> dict:
    from PIL import Image

    from .scanner import _find_window, _image_similarity, _rect, _screenshot_window

    index = knowledge.load_index(directory)
    pages = index.get("pages", [])
    if requested_page_id:
        pages = [item for item in pages if item.get("id") == requested_page_id]
        if not pages:
            raise KeyError(f"Unknown page: {requested_page_id}")
    app_name = index.get("application", {}).get("name", "")
    if not app_name:
        raise RuntimeError("Application knowledge has no window name")
    window = _find_window(app_name)
    window_rect = _rect(window)
    screenshot = _screenshot_window(window_rect)
    matches = []
    for page in pages:
        image_name = page.get("image", "")
        image_path = directory / image_name if image_name else None
        if image_path is None or not image_path.is_file():
            continue
        try:
            with Image.open(image_path) as reference:
                score = _image_similarity(reference, screenshot)
        except OSError:
            continue
        matches.append((score, page))
    if not matches:
        raise RuntimeError("No saved page image is available for current-page matching")
    score, page = max(matches, key=lambda item: item[0])
    if score < PAGE_MATCH_THRESHOLD:
        raise RuntimeError(
            f"Current window does not reliably match a learned page (best={score:.3f}); "
            "scan this page before showing controls"
        )
    return {
        "index": index,
        "page": page,
        "similarity": score,
        "window_rect": window_rect,
        "window": window,
        "screenshot": screenshot,
    }


def _runtime_control_markers(
    directory: Path,
    records: list[dict],
    window,
    window_rect: dict,
    screenshot,
) -> tuple[list[dict], list[dict]]:
    """Resolve current rectangles without trusting persisted geometry hints."""
    from PIL import Image

    from . import ui_cli
    from .scanner import _clip_rect, _rect, locate_template, resolve_location

    prefix_cache = {}
    resolved = []
    unresolved = []
    for record in records:
        rectangle = None
        method = ""
        try:
            control = resolve_location(
                record.get("location", {}), window=window, prefix_cache=prefix_cache
            )
            if control and control is not False:
                candidate = _clip_rect(_rect(control), window_rect)
                if candidate["width"] > 1 and candidate["height"] > 1:
                    rectangle = candidate
                    method = "location"
        except Exception:
            pass
        if rectangle is None:
            image_names = [*record.get("image_variants", [])]
            if record.get("image"):
                image_names.append(record["image"])
            best_template = None
            for image_name in dict.fromkeys(image_names):
                image_path = directory / image_name
                if not image_path.is_file():
                    continue
                try:
                    with Image.open(image_path) as reference:
                        result = locate_template(reference, screenshot, window_rect)
                except OSError:
                    continue
                if best_template is None or result["score"] > best_template["score"]:
                    best_template = result
            if best_template and best_template["found"]:
                rectangle = best_template["rect"]
                method = "image"
        if rectangle is None:
            ocr = ui_cli._locate_record_text(record, screenshot, window_rect)
            if ocr:
                rectangle = ocr["rect"]
                method = "ocr"
        if rectangle is None:
            unresolved.append(
                {
                    "control_id": record.get("id", ""),
                    "semantic_name": record.get("semantic_name") or record.get("name"),
                    "reason": "current LOCATION, image template, and OCR did not resolve",
                }
            )
            continue
        runtime_record = {key: value for key, value in record.items() if key != "normalized_rect"}
        resolved.append({**runtime_record, "rect": rectangle, "resolution_method": method})

    markers = control_markers(resolved, window_rect, window_rect)
    methods = {item.get("id"): item.get("resolution_method") for item in resolved}
    for marker in markers:
        marker["resolution_method"] = methods.get(marker["control_id"], "")
    return markers, unresolved


def show_page_controls(
    directory: Path,
    page_id: str = "",
    include: str = "executable",
    duration_ms: int = 5000,
) -> dict:
    """Match the visible page, list its controls, and draw them in one overlay."""
    include = include.strip().lower()
    if include not in {"executable", "known"}:
        raise ValueError("include must be executable or known")
    context = _match_current_page(directory, page_id)
    page = context["page"]
    records = [
        item
        for item in context["index"].get("controls", [])
        if item.get("page_id") == page.get("id")
    ]
    if include == "executable":
        records = [
            item
            for item in records
            if item.get("status") == "verified"
            and item.get("semantic_status") in {"verified", "manual"}
            and item.get("actions")
        ]
    markers, unresolved = _runtime_control_markers(
        directory,
        records,
        context["window"],
        context["window_rect"],
        context["screenshot"],
    )
    overlay = show_markers(markers, duration_ms)
    commands_by_control: dict[str, list[str]] = {}
    for item in knowledge.available_commands(directory, page.get("id", "")):
        commands_by_control.setdefault(item["control_id"], []).append(item["command"])
    controls = [
        {
            "index": marker["index"],
            "control_id": marker["control_id"],
            "semantic_name": marker["label"],
            "status": marker["status"],
            "commands": commands_by_control.get(marker["control_id"], []),
            "rect": marker["rect"],
            "resolution_method": marker["resolution_method"],
        }
        for marker in markers
    ]
    return {
        "ok": True,
        "page_id": page.get("id"),
        "page_name": page.get("name", page.get("id")),
        "page_similarity": context["similarity"],
        "include": include,
        "controls": controls,
        "unresolved_controls": unresolved,
        "overlay": overlay,
    }
