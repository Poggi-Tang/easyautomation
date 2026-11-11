"""Human-defined, window-relative regions shared by the GUI and scanner."""

from __future__ import annotations

import math


def validate_regions(regions: list[dict]) -> list[dict]:
    """Accept flat, named, non-overlapping normalized rectangles."""
    if not isinstance(regions, list) or not regions:
        raise ValueError("请先划分至少一个区域")
    result, ids = [], set()
    for item in regions:
        identifier, name = str(item.get("id", "")).strip(), str(item.get("name", "")).strip()
        if not identifier or identifier in ids or not name:
            raise ValueError("区域必须有不重复的 ID 和名称")
        rectangle = item.get("normalized_rect", {})
        values = [rectangle.get(key) for key in ("left", "top", "right", "bottom")]
        if any(isinstance(v, bool) or not isinstance(v, (float, int))
               or not math.isfinite(v) for v in values):
            raise ValueError("区域坐标必须是有限数值")
        left, top, right, bottom = values
        if not (0 <= left < right <= 1 and 0 <= top < bottom <= 1):
            raise ValueError("区域必须位于截图内，并且宽高大于零")
        for previous in result:
            r = previous["normalized_rect"]
            overlap = min(right, r["right"]) - max(left, r["left"])
            overlap_y = min(bottom, r["bottom"]) - max(top, r["top"])
            if overlap > 1e-6 and overlap_y > 1e-6:
                raise ValueError(f"区域“{name}”与“{previous['name']}”重叠，请调整边界")
        ids.add(identifier)
        result.append({**item, "id": identifier, "name": name,
                       "normalized_rect": dict(zip(("left", "top", "right", "bottom"), values))})
    return result


def assign_region(rectangle: dict, regions: list[dict]) -> tuple[str, str]:
    """Assign by visible coverage; ambiguous/cross-region containers remain explicit."""
    left, top, right, bottom = [float(rectangle[k]) for k in ("left", "top", "right", "bottom")]
    area = (right-left)*(bottom-top)
    if area <= 0:
        return "unassigned", "invalid-bounds"
    matches = []
    for region in regions:
        r = region["normalized_rect"]
        intersection = max(0, min(right, r["right"])-max(left, r["left"])) * max(
            0, min(bottom, r["bottom"])-max(top, r["top"]))
        if intersection:
            matches.append((intersection/area, region["id"]))
    if not matches:
        return "unassigned", "outside-defined-regions"
    matches.sort(key=lambda item: (-item[0], item[1]))
    if matches[0][0] < .5 or (len(matches) > 1 and abs(matches[0][0]-matches[1][0]) < 1e-6):
        return "unassigned", "cross-region-ambiguous"
    return matches[0][1], "contained" if matches[0][0] >= .999 else "largest-overlap"


def segmentation(layout: dict, window_rect: dict) -> dict:
    """Map user rectangles to a same-size screenshot; never infer a new layout."""
    size = layout.get("image_size", {})
    if (size.get("width"), size.get("height")) != (window_rect["width"], window_rect["height"]):
        raise ValueError("软件窗口尺寸已改变，请刷新截图并确认区域后再学习")
    regions = validate_regions(layout.get("regions", []))
    width, height = window_rect["width"], window_rect["height"]
    output = []
    for region in regions:
        n = region["normalized_rect"]
        left, top, right, bottom = round(n["left"]*width), round(n["top"]*height), round(n["right"]*width), round(n["bottom"]*height)
        if right-left < 2 or bottom-top < 2:
            raise ValueError("区域在当前截图中不足 2 像素，请放大后重新划分")
        output.append({"id": region["id"], "name": region["name"], "role": "manual-region",
                       "description": region.get("description", ""),
                       "rect": dict(left=left, top=top, right=right, bottom=bottom,
                                    width=right-left, height=bottom-top)})
    return {"page": {"id": layout.get("page_id", "manual-main"),
                     "name": layout.get("page_name", "人工区域布局")},
            "regions": output, "controls": [],
            "coverage": {"region_source": "manual", "visual_tiles": 0}}


def native_semantics(candidate: dict) -> dict:
    """Keep literal UIA labels distinct from inferred or human-taught purposes."""
    name = str(candidate.get("uia_name") or "").strip()
    label = name or candidate.get("automation_id") or candidate.get("control_type", "Control")
    return {"semantic_name": label, "intent": label, "role": "control",
            "name":name, "automation_id":candidate.get("automation_id", ""),
            "help_text":candidate.get("help_text", ""), "control_type":candidate.get("control_type", ""),
            "description": candidate.get("help_text", ""),
            "actions": candidate.get("supported_actions", []), "aliases": [],
            "risk": "unknown", "requires_confirmation": True,
            "confidence": 1.0 if name else .4, "source": "uia-metadata",
            "evidence": ["literal UIA Name" if name else "UIA attributes only"],
            "ambiguity": "" if name else "控件缺少可读名称，尚未确认具体功能",
            "entity_type": "ui-element", "observed_value": name}
