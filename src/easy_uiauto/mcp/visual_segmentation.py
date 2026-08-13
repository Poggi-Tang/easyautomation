"""Fast class-agnostic rectangle detection inside one screen region."""

from __future__ import annotations

import time
from collections import defaultdict
from heapq import heappop, heappush
from pathlib import Path
from tempfile import gettempdir
from uuid import uuid4


def normalize_screen_rect(rect: dict) -> dict:
    """Validate a screen rectangle and fill derived right/bottom/size fields."""
    if not isinstance(rect, dict):
        raise ValueError("rect must be a JSON object")
    for wrapper in ("rect", "bounds", "rectangle"):
        if isinstance(rect.get(wrapper), dict):
            rect = rect[wrapper]
            break
    try:
        left = int(rect.get("left", rect.get("x", 0)))
        top = int(rect.get("top", rect.get("y", 0)))
        width = int(rect.get("width", 0))
        height = int(rect.get("height", 0))
        right = int(rect.get("right", left + width))
        bottom = int(rect.get("bottom", top + height))
    except (TypeError, ValueError) as error:
        raise ValueError("rect coordinates must be integers") from error
    if width <= 0:
        width = right - left
    if height <= 0:
        height = bottom - top
    if width <= 1 or height <= 1:
        raise ValueError("rect width and height must be greater than 1")
    return {
        "left": left,
        "top": top,
        "right": left + width,
        "bottom": top + height,
        "width": width,
        "height": height,
    }


def capture_rect(rect: dict):
    """Capture exactly one screen rectangle."""
    from PIL import ImageGrab

    normalized = normalize_screen_rect(rect)
    bounds = (
        normalized["left"],
        normalized["top"],
        normalized["right"],
        normalized["bottom"],
    )
    try:
        return ImageGrab.grab(bbox=bounds, all_screens=True)
    except Exception:
        import pyautogui

        return pyautogui.screenshot(
            region=(
                normalized["left"],
                normalized["top"],
                normalized["width"],
                normalized["height"],
            )
        )


def _iou(first: dict, second: dict) -> float:
    left = max(first["left"], second["left"])
    top = max(first["top"], second["top"])
    right = min(first["right"], second["right"])
    bottom = min(first["bottom"], second["bottom"])
    intersection = max(0, right - left) * max(0, bottom - top)
    if not intersection:
        return 0.0
    union = first["width"] * first["height"] + second["width"] * second["height"]
    return intersection / max(1, union - intersection)


def _contains(outer: dict, inner: dict, margin: int = 1) -> bool:
    return (
        outer["left"] <= inner["left"] + margin
        and outer["top"] <= inner["top"] + margin
        and outer["right"] >= inner["right"] - margin
        and outer["bottom"] >= inner["bottom"] - margin
        and outer["width"] * outer["height"] > inner["width"] * inner["height"]
    )


def _candidate(x: int, y: int, width: int, height: int, source: str, image_size) -> dict:
    image_width, image_height = image_size
    area_ratio = width * height / max(1, image_width * image_height)
    return {
        "left": int(x),
        "top": int(y),
        "right": int(x + width),
        "bottom": int(y + height),
        "width": int(width),
        "height": int(height),
        "area_ratio": round(area_ratio, 6),
        "sources": [source],
    }


def _contour_candidates(mask, source: str, image_size, min_width: int, min_height: int):
    import cv2

    image_width, image_height = image_size
    contours, _hierarchy = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    candidates = []
    for contour in contours:
        x, y, width, height = cv2.boundingRect(contour)
        if width < min_width or height < min_height:
            continue
        area_ratio = width * height / max(1, image_width * image_height)
        if area_ratio < 0.00008 or area_ratio > 0.98:
            continue
        contour_area = float(cv2.contourArea(contour))
        fill_ratio = contour_area / max(1, width * height)
        if source == "outline" and fill_ratio < 0.025:
            continue
        item = _candidate(x, y, width, height, source, image_size)
        item["fill_ratio"] = round(fill_ratio, 4)
        candidates.append(item)
    return candidates


def _candidate_masks(image):
    """Build complementary masks for borders, filled regions, and grouped text strokes."""
    try:
        import cv2
        import numpy
    except ModuleNotFoundError as error:
        raise RuntimeError(
            'Visual rectangle detection requires pip install "easy-uiauto[mcp,vision]"'
        ) from error

    rgb = numpy.array(image.convert("RGB"))
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    denoised = cv2.GaussianBlur(gray, (3, 3), 0)
    edges = cv2.Canny(denoised, 35, 110)
    outline = cv2.morphologyEx(
        edges,
        cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)),
        iterations=1,
    )

    text_width = max(3, min(11, image.width // 80))
    grouped = cv2.morphologyEx(
        edges,
        cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (text_width, 3)),
        iterations=1,
    )

    _threshold, dark = cv2.threshold(
        denoised, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU
    )
    dark = cv2.morphologyEx(
        dark,
        cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (text_width, 2)),
        iterations=1,
    )

    border_pixels = numpy.concatenate(
        (rgb[0, :, :], rgb[-1, :, :], rgb[:, 0, :], rgb[:, -1, :]), axis=0
    )
    background = numpy.median(border_pixels, axis=0)
    color_distance = numpy.max(
        numpy.abs(rgb.astype("int16") - background.astype("int16")), axis=2
    )
    foreground = (color_distance >= 6).astype("uint8") * 255
    foreground = cv2.morphologyEx(
        foreground,
        cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)),
        iterations=1,
    )
    return (
        (foreground, "background-difference"),
        (outline, "outline"),
        (grouped, "grouped-edge"),
        (dark, "dark-region"),
    )


def _merge_candidates(candidates: list[dict]) -> list[dict]:
    """Merge near-identical boxes while preserving genuinely nested boxes."""
    merged = []
    ordered = sorted(
        candidates,
        key=lambda value: value["width"] * value["height"],
        reverse=True,
    )
    for item in ordered:
        duplicate = None
        for existing in merged:
            same_edges = max(
                abs(item["left"] - existing["left"]),
                abs(item["top"] - existing["top"]),
                abs(item["right"] - existing["right"]),
                abs(item["bottom"] - existing["bottom"]),
            ) <= 3
            if same_edges or _iou(item, existing) >= 0.86:
                duplicate = existing
                break
        if duplicate is None:
            merged.append(item)
            continue
        duplicate["sources"] = list(dict.fromkeys([*duplicate["sources"], *item["sources"]]))
        duplicate["fill_ratio"] = max(
            float(duplicate.get("fill_ratio", 0)), float(item.get("fill_ratio", 0))
        )
    return merged


def _remove_internal_fragments(boxes: list[dict]) -> list[dict]:
    """Drop single-mask fragments already represented by a strong enclosing candidate."""
    retained = []
    for item in boxes:
        item_area = item["width"] * item["height"]
        fragment = False
        for outer in boxes:
            if not _contains(outer, item, margin=0):
                continue
            outer_area = outer["width"] * outer["height"]
            ratio = item_area / max(1, outer_area)
            if ratio <= 0.15 or (len(item["sources"]) == 1 and len(outer["sources"]) >= 2):
                if ratio <= 0.6:
                    fragment = True
                    break
        if not fragment:
            retained.append(item)
    return retained


def _assign_hierarchy(boxes: list[dict]) -> None:
    for index, item in enumerate(boxes, start=1):
        item["id"] = index
    for item in boxes:
        parents = [candidate for candidate in boxes if _contains(candidate, item)]
        if parents:
            parent = min(parents, key=lambda value: value["width"] * value["height"])
            item["parent_id"] = parent["id"]
        else:
            item["parent_id"] = None
    child_counts = {
        item["id"]: sum(candidate.get("parent_id") == item["id"] for candidate in boxes)
        for item in boxes
    }
    for item in boxes:
        item["child_count"] = child_counts[item["id"]]
        item["kind"] = "container" if item["child_count"] else "element"
        item["size_policy"] = "variable"
        source_score = min(0.32, 0.08 * len(item["sources"]))
        fill_score = 0.06 if float(item.get("fill_ratio", 0)) >= 0.1 else 0.0
        item["confidence"] = round(min(0.99, 0.55 + source_score + fill_score), 3)


def _overlap(first_start: int, first_end: int, second_start: int, second_end: int) -> int:
    return max(0, min(first_end, second_end) - max(first_start, second_start))


def _axis_overlap_ratio(first: dict, second: dict, axis: str) -> float:
    if axis == "x":
        overlap = _overlap(first["left"], first["right"], second["left"], second["right"])
        denominator = min(first["width"], second["width"])
    else:
        overlap = _overlap(first["top"], first["bottom"], second["top"], second["bottom"])
        denominator = min(first["height"], second["height"])
    return overlap / max(1, denominator)


def _center(item: dict, axis: str) -> float:
    if axis == "x":
        return item["left"] + item["width"] / 2
    return item["top"] + item["height"] / 2


def _horizontal_alignment_tolerance(first: dict, second: dict) -> int:
    return max(3, round(min(first["height"], second["height"]) * 0.35))


def _vertical_alignment_tolerance(first: dict, second: dict) -> int:
    return max(4, round(min(first["width"], second["width"]) * 0.2))


def _relation(first: dict, second: dict) -> dict | None:
    """Describe a nearby spatial relation without using equal width or height."""
    x_overlap = _axis_overlap_ratio(first, second, "x")
    y_overlap = _axis_overlap_ratio(first, second, "y")
    horizontal_tolerance = _horizontal_alignment_tolerance(first, second)
    vertical_tolerance = _vertical_alignment_tolerance(first, second)
    horizontal_gap = max(first["left"], second["left"]) - min(first["right"], second["right"])
    vertical_gap = max(first["top"], second["top"]) - min(first["bottom"], second["bottom"])
    horizontal_gap = max(0, horizontal_gap)
    vertical_gap = max(0, vertical_gap)
    near_x = horizontal_gap <= max(12, round(min(first["height"], second["height"]) * 1.5))
    near_y = vertical_gap <= max(12, round(min(first["height"], second["height"]) * 1.5))
    horizontal_alignment = [
        edge
        for edge in ("top", "bottom")
        if abs(first[edge] - second[edge]) <= horizontal_tolerance
    ]
    vertical_alignment = [
        edge
        for edge in ("left", "right")
        if abs(first[edge] - second[edge]) <= vertical_tolerance
    ]
    if y_overlap >= 0.45 and near_x and horizontal_alignment:
        left, right = (
            (first, second)
            if _center(first, "x") <= _center(second, "x")
            else (second, first)
        )
        return {
            "type": "right-of",
            "source_id": left["id"],
            "target_id": right["id"],
            "gap": horizontal_gap,
            "overlap": round(y_overlap, 4),
            "aligned": horizontal_alignment,
        }
    if x_overlap >= 0.45 and near_y and vertical_alignment:
        top, bottom = (
            (first, second)
            if _center(first, "y") <= _center(second, "y")
            else (second, first)
        )
        return {
            "type": "below",
            "source_id": top["id"],
            "target_id": bottom["id"],
            "gap": vertical_gap,
            "overlap": round(x_overlap, 4),
            "aligned": vertical_alignment,
        }
    return None


def _layout_graph(boxes: list[dict]) -> list[dict]:
    """Build local sibling relations; nested boxes are represented by parent_id instead."""
    siblings = defaultdict(list)
    for item in boxes:
        siblings[item.get("parent_id")].append(item)
    relations = []
    for items in siblings.values():
        for index, first in enumerate(items):
            for second in items[index + 1 :]:
                relation = _relation(first, second)
                if relation:
                    relations.append(relation)
    for item in boxes:
        item["relation_ids"] = [
            index
            for index, relation in enumerate(relations, start=1)
            if item["id"] in {relation["source_id"], relation["target_id"]}
        ]
    return [{"id": index, **relation} for index, relation in enumerate(relations, start=1)]


def _components(node_ids: list[int], relations: list[dict]) -> list[list[int]]:
    adjacency = {node_id: set() for node_id in node_ids}
    for relation in relations:
        source = relation["source_id"]
        target = relation["target_id"]
        if source in adjacency and target in adjacency:
            adjacency[source].add(target)
            adjacency[target].add(source)
    groups = []
    unseen = set(node_ids)
    while unseen:
        start = unseen.pop()
        stack = [start]
        group = [start]
        while stack:
            current = stack.pop()
            for neighbor in adjacency[current] & unseen:
                unseen.remove(neighbor)
                stack.append(neighbor)
                group.append(neighbor)
        if len(group) > 1:
            groups.append(sorted(group))
    return groups


def _group_signature(member_ids: list[int], relations: list[dict]) -> str:
    member_set = set(member_ids)
    relation_types = sorted(
        relation["type"]
        for relation in relations
        if {relation["source_id"], relation["target_id"]} <= member_set
    )
    return f"members:{len(member_ids)}|relations:{','.join(relation_types)}"


def _attach_vertical_singletons(
    node_ids: list[int],
    horizontal_groups: list[list[int]],
    relations: list[dict],
) -> list[list[int]]:
    """Attach vertical singleton content to its nearest horizontal row seed."""
    allowed = set(node_ids)
    adjacency = defaultdict(list)
    for relation in relations:
        if relation["type"] != "below":
            continue
        source = relation["source_id"]
        target = relation["target_id"]
        if source not in allowed or target not in allowed:
            continue
        weight = max(1, int(relation.get("gap", 0)) + 1)
        adjacency[source].append((target, weight))
        adjacency[target].append((source, weight))

    owner = {}
    distance = {}
    queue = []
    seeds = set()
    for group_index, members in enumerate(horizontal_groups):
        for member_id in members:
            seeds.add(member_id)
            owner[member_id] = group_index
            distance[member_id] = 0
            heappush(queue, (0, group_index, member_id))

    while queue:
        current_distance, group_index, current = heappop(queue)
        if distance.get(current) != current_distance or owner.get(current) != group_index:
            continue
        for neighbor, weight in adjacency[current]:
            if neighbor in seeds:
                continue
            candidate_distance = current_distance + weight
            existing_distance = distance.get(neighbor)
            if existing_distance is not None and candidate_distance >= existing_distance:
                continue
            distance[neighbor] = candidate_distance
            owner[neighbor] = group_index
            heappush(queue, (candidate_distance, group_index, neighbor))

    attached = [set(members) for members in horizontal_groups]
    for node_id, group_index in owner.items():
        attached[group_index].add(node_id)
    return [sorted(members) for members in attached]


def _layout_groups(
    boxes: list[dict],
    relations: list[dict],
    image_size: tuple[int, int],
    padding: int = 4,
) -> list[dict]:
    """Create size-independent groups without joining adjacent repeated rows."""
    by_id = {item["id"]: item for item in boxes}
    sibling_ids = defaultdict(list)
    for item in boxes:
        sibling_ids[item.get("parent_id")].append(item["id"])
    groups = []
    image_width, image_height = image_size
    for parent_id, node_ids in sibling_ids.items():
        horizontal_relations = [
            relation for relation in relations if relation["type"] == "right-of"
        ]
        horizontal_groups = _components(node_ids, horizontal_relations)
        attached_groups = _attach_vertical_singletons(
            node_ids, horizontal_groups, relations
        )
        for members in attached_groups:
            items = [by_id[member_id] for member_id in members]
            left = max(0, min(item["left"] for item in items) - padding)
            top = max(0, min(item["top"] for item in items) - padding)
            right = min(image_width, max(item["right"] for item in items) + padding)
            bottom = min(image_height, max(item["bottom"] for item in items) + padding)
            groups.append(
                {
                    "id": len(groups) + 1,
                    "parent_id": parent_id,
                    "member_ids": members,
                    "rect_in_input": {
                        "left": left,
                        "top": top,
                        "right": right,
                        "bottom": bottom,
                        "width": right - left,
                        "height": bottom - top,
                    },
                    "layout_signature": _group_signature(members, relations),
                    "size_policy": "variable",
                }
            )
    return groups


def _group_with_coordinates(group: dict, screen_rect: dict) -> dict:
    rectangle = group["rect_in_input"]
    return {
        **group,
        "rect": {
            **rectangle,
            "left": screen_rect["left"] + rectangle["left"],
            "top": screen_rect["top"] + rectangle["top"],
            "right": screen_rect["left"] + rectangle["right"],
            "bottom": screen_rect["top"] + rectangle["bottom"],
        },
        "rect_relative_to_input": {
            key: round(
                rectangle[key]
                / screen_rect[
                    "width" if key in {"left", "right", "width"} else "height"
                ],
                6,
            )
            for key in ("left", "top", "right", "bottom", "width", "height")
        },
    }


def _with_coordinates(box: dict, screen_rect: dict) -> dict:
    width = screen_rect["width"]
    height = screen_rect["height"]
    local_rect = {
        "left": box["left"],
        "top": box["top"],
        "right": box["right"],
        "bottom": box["bottom"],
        "width": box["width"],
        "height": box["height"],
    }
    metadata = {
        key: value
        for key, value in box.items()
        if key not in {"left", "top", "right", "bottom", "width", "height"}
    }
    return {
        **metadata,
        "rect": {
            "left": screen_rect["left"] + box["left"],
            "top": screen_rect["top"] + box["top"],
            "right": screen_rect["left"] + box["right"],
            "bottom": screen_rect["top"] + box["bottom"],
            "width": box["width"],
            "height": box["height"],
        },
        "rect_in_input": local_rect,
        "rect_relative_to_input": {
            "left": round(box["left"] / width, 6),
            "top": round(box["top"] / height, 6),
            "right": round(box["right"] / width, 6),
            "bottom": round(box["bottom"] / height, 6),
            "width": round(box["width"] / width, 6),
            "height": round(box["height"] / height, 6),
        },
    }


def _annotate(
    image,
    boxes: list[dict],
    destination: Path,
    groups: list[dict] | None = None,
) -> None:
    from PIL import ImageDraw

    annotated = image.convert("RGB").copy()
    draw = ImageDraw.Draw(annotated)
    for group in groups or []:
        rectangle = group["rect_in_input"]
        draw.rectangle(
            (
                rectangle["left"],
                rectangle["top"],
                rectangle["right"],
                rectangle["bottom"],
            ),
            outline="#00a152",
            width=2,
        )
        draw.text(
            (rectangle["left"] + 2, max(0, rectangle["top"] - 12)),
            f"G{group['id']}",
            fill="#00a152",
        )
    for item in boxes:
        color = "#e53935" if item.get("parent_id") is None else "#1976d2"
        draw.rectangle(
            (item["left"], item["top"], item["right"], item["bottom"]),
            outline=color,
            width=2,
        )
        draw.text((item["left"] + 2, item["top"] + 1), str(item["id"]), fill=color)
    destination.parent.mkdir(parents=True, exist_ok=True)
    annotated.save(destination)


def detect_rectangles(
    image,
    screen_rect: dict,
    min_width: int = 6,
    min_height: int = 6,
    max_boxes: int = 300,
) -> dict:
    """Detect class-agnostic visual rectangles within an already captured region."""
    started = time.perf_counter()
    normalized = normalize_screen_rect(screen_rect)
    if image.size != (normalized["width"], normalized["height"]):
        raise ValueError("image size must match the input rect width and height")
    min_width = max(2, int(min_width))
    min_height = max(2, int(min_height))
    max_boxes = max(1, min(int(max_boxes), 2000))
    candidates = []
    for mask, source in _candidate_masks(image):
        candidates.extend(
            _contour_candidates(
                mask,
                source,
                image.size,
                min_width,
                min_height,
            )
        )
    boxes = _remove_internal_fragments(_merge_candidates(candidates))
    truncated = len(boxes) > max_boxes
    if truncated:
        boxes.sort(
            key=lambda value: (
                len(value["sources"]),
                min(float(value["area_ratio"]), 0.1),
                float(value.get("fill_ratio", 0)),
            ),
            reverse=True,
        )
        boxes = boxes[:max_boxes]
    boxes.sort(key=lambda value: (value["top"], value["left"], -value["area_ratio"]))
    _assign_hierarchy(boxes)
    relations = _layout_graph(boxes)
    groups = _layout_groups(boxes, relations, image.size)
    return {
        "ok": True,
        "input_rect": normalized,
        "boxes": [_with_coordinates(box, normalized) for box in boxes],
        "box_count": len(boxes),
        "relations": relations,
        "relation_count": len(relations),
        "groups": [_group_with_coordinates(group, normalized) for group in groups],
        "group_count": len(groups),
        "truncated": truncated,
        "detector": "local-class-agnostic-rectangles-v2",
        "layout_model": "size-independent-spatial-relations-v1",
        "semantic_analysis": False,
        "timing_ms": round((time.perf_counter() - started) * 1000, 2),
    }


def detect_screen_rectangles(
    rect: dict,
    min_width: int = 6,
    min_height: int = 6,
    max_boxes: int = 300,
    save_annotated: bool = True,
) -> dict:
    """Capture and detect visual rectangles in exactly one supplied screen rect."""
    normalized = normalize_screen_rect(rect)
    image = capture_rect(normalized)
    result = detect_rectangles(image, normalized, min_width, min_height, max_boxes)
    if save_annotated:
        destination = (
            Path(gettempdir())
            / "easy_uiauto"
            / f"visual-segmentation-{uuid4().hex[:12]}.png"
        )
        local_boxes = [
            {
                **item,
                **item["rect_in_input"],
            }
            for item in result["boxes"]
        ]
        _annotate(image, local_boxes, destination, result["groups"])
        result["annotated_image"] = str(destination)
    return result
