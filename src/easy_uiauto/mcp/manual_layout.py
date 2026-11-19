"""Human-defined, window-relative regions shared by the GUI and scanner."""

from __future__ import annotations

import math

from .semantic_identity import generated_description


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


def adapt_to_window(layout: dict, window_rect: dict, *, window=None, owner=None) -> dict:
    """Rebind picked regions on resize; keep manual crop geometry independent of identity."""
    from copy import deepcopy
    size = layout.get("image_size", {})
    if (size.get("width"), size.get("height")) == (window_rect["width"], window_rect["height"]):return layout
    result = deepcopy(layout)
    if window is not None:
        from ..library_actions import resolve_record
        from . import knowledge, scanner
        for region in result.get("regions", []):
            if not region.get("location", {}).get("Xpath"):
                break
            control = resolve_record({**region, "window_scope":layout.get("window_scope", {"kind":"main"})},
                                     owner or dict(handle=layout["window_handle"],process_id=layout["window_process_id"],
                                                   class_name=layout["window_class"]), window)
            for key, attribute in (("automation_id","AutomationId"),("class_name","ClassName"),("control_type","ControlTypeName")):
                if region.get(key) and region[key] != getattr(control, attribute, ""):
                    raise ValueError("区域定位命中了不同控件，请重新选择区域；已有控件属性保留")
            rectangle = scanner._rect(control)
            if any(rectangle[k] < window_rect[k]-2 for k in ("left","top")) or any(
                    rectangle[k] > window_rect[k]+2 for k in ("right","bottom")):
                raise ValueError("区域已超出当前窗口，未扩大扫描范围；请检查区域边界")
            region["normalized_rect"] = knowledge.normalized_rect(rectangle, window_rect)
        else:
            validate_regions(result.get("regions", []))
            result["image_size"] = dict(width=window_rect["width"],height=window_rect["height"])
            result["geometry_adapted"] = "native-location"
            return result
    # A drawn crop has no native identity. Only reuse it when the reference frame
    # is visually the same after scaling; never blindly stretch a changed layout.
    reference = layout.get("reference_image")
    if window is not None and reference:
        from PIL import Image
        from ..state_matching import same_image
        from . import scanner
        with Image.open(reference) as old:
            same = same_image(old, scanner._screenshot_window(window_rect))
        if same:
            result = deepcopy(layout)
            result["image_size"] = dict(width=window_rect["width"],height=window_rect["height"])
            result["geometry_adapted"] = "scaled-reference"
            return result
    return layout


def resolve_learning_region(region, screenshot, window_rect, *, window, owner, image_path):
    """Obtain a live crop from identity or a unique template, never from a saved box."""
    from PIL import Image

    from ..library_actions import resolve_record
    from ..utils import CONTROL_TYPE_IDS
    from . import scanner
    from .protocol import normalize_location
    from .results import ActionError

    location = normalize_location(region.get("location") or {})
    # Position-mode picker paths contain sibling ordinals. They are not identities.
    identifier = region.get("automation_id") or location["AutomationId"]
    name = region.get("native_name") or location["Name"]
    class_name = region.get("class_name") or location["ClassName"]
    control_type = region.get("control_type") or location["ControlType"]
    if identifier or name:
        selector = {"AutomationId": identifier} if identifier else {"Name": name}
        if class_name:
            selector["ClassName"] = class_name
        identity = dict(selector, ControlType=control_type)
        root_matches = all(getattr(window, key, "") == value for key, value in selector.items()) and (
            not control_type or getattr(window, "ControlTypeName", "") == control_type
        )
        try:
            control = window if root_matches else resolve_record({"location": identity}, owner, window)
        except ActionError as error:
            if error.code != "control_not_found" or error.reason:
                raise
        else:
            if not getattr(control, "IsOffscreen", False) and all(
                getattr(control, key, "") == value for key, value in selector.items()
            ) and (
                not control_type or control.ControlTypeName == control_type
            ):
                if control_type:
                    selector["ControlType"] = CONTROL_TYPE_IDS[control_type]
                duplicate = window.Control(foundIndex=1 if root_matches else 2, **selector)
                if duplicate.Refind(maxSearchSeconds=0, raiseException=False):
                    raise ActionError("ambiguous_match", "区域属性匹配到多个控件，未扫描")
                rectangle = scanner._rect(control)
                if not (window_rect["left"] <= rectangle["left"] < rectangle["right"]
                        <= window_rect["right"] and window_rect["top"] <= rectangle["top"]
                        < rectangle["bottom"] <= window_rect["bottom"]):
                    return None
                return rectangle, "native-location", control
    if image_path:
        with Image.open(image_path) as image:
            match = scanner.locate_template(image, screenshot, window_rect)
        if match["found"]:
            return match["rect"], "image-location", None
    return None


def segmentation(layout: dict, window_rect: dict) -> dict:
    """Map user rectangles to a same-size screenshot; never infer a new layout."""
    size = layout.get("image_size", {})
    if (size.get("width"), size.get("height")) != (window_rect["width"], window_rect["height"]):
        raise ValueError(
            "软件窗口尺寸已改变，尚未确认区域在新尺寸下的边界；请刷新截图确认区域，已有控件属性和学习记录保留；"
            f"记录尺寸={size.get('width')}×{size.get('height')}，"
            f"当前可见帧={window_rect['width']}×{window_rect['height']}，"
            f"window_handle={layout.get('window_handle', 0)}，坐标来源=visible-window-bounds")
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
            "description": generated_description(candidate.get("help_text")),
            "actions": candidate.get("supported_actions", []), "aliases": [],
            "risk": "unknown", "requires_confirmation": True,
            "confidence": 1.0 if name else .4, "source": "uia-metadata",
            "evidence": ["literal UIA Name" if name else "UIA attributes only"],
            "ambiguity": "" if name else "控件缺少可读名称，尚未确认具体功能",
            "entity_type": "ui-element", "observed_value": name}
