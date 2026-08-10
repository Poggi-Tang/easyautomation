"""Protocol helpers for the UI automation TCP service.

Shared between service.py and client.py. Zero dependencies beyond stdlib.
"""

import json
from typing import Any

LOCATION_KEYS = (
    "WindowName",
    "Name",
    "ClassName",
    "ControlType",
    "foundIndex",
    "AutomationId",
    "Xpath",
    "Img",
    "PARAMETERS",
)


def _normalize_xpath_step(step: dict) -> dict:
    """Convert recorded, MCP, or vector-store XPath keys to the core schema."""
    normalized = {
        "ControlType": step.get("ControlType", step.get("control_type", "")),
    }
    aliases = (
        ("Name", "name"),
        ("ClassName", "class_name"),
        ("AutomationId", "automation_id"),
        ("foundIndex", "found_index"),
        ("searchDepth", "search_depth"),
    )
    for canonical, alternate in aliases:
        value = step.get(canonical, step.get(alternate))
        if value not in (None, ""):
            normalized[canonical] = value
    return normalized


def location_from_xpath(
    xpath: list[dict],
    parameters: dict | None = None,
    image: str = "",
) -> dict:
    """Build the canonical easy_uiauto LOCATION object from a control XPath."""
    normalized_xpath = [_normalize_xpath_step(step) for step in xpath]
    if not normalized_xpath:
        return build_location(parameters=json.dumps(parameters or {}, ensure_ascii=False))
    first = normalized_xpath[0]
    last = normalized_xpath[-1]
    return {
        "WindowName": first.get("Name", ""),
        "Name": last.get("Name", ""),
        "ClassName": last.get("ClassName", ""),
        "ControlType": last.get("ControlType", ""),
        "foundIndex": last.get("foundIndex", ""),
        "AutomationId": last.get("AutomationId", ""),
        "Xpath": normalized_xpath,
        "Img": image,
        "PARAMETERS": parameters or {},
    }


def normalize_location(value: dict) -> dict:
    """Accept a LOCATION, recorded action, coordinate result, or vector record."""
    if not isinstance(value, dict):
        raise ValueError("location must be a JSON object")

    candidate = value
    for wrapper_key in ("LOCATION", "location"):
        wrapped = candidate.get(wrapper_key)
        if isinstance(wrapped, dict):
            candidate = wrapped
            break

    raw_xpath = candidate.get("Xpath", candidate.get("xpath", []))
    if isinstance(raw_xpath, str):
        raw_xpath = json.loads(raw_xpath)
    if not isinstance(raw_xpath, list):
        raise ValueError("location Xpath must be a JSON array")

    derived = location_from_xpath(raw_xpath)
    aliases = {
        "WindowName": "window_name",
        "Name": "name",
        "ClassName": "class_name",
        "ControlType": "control_type",
        "foundIndex": "found_index",
        "AutomationId": "automation_id",
        "Img": "image",
        "PARAMETERS": "parameters",
    }
    for key, alternate in aliases.items():
        if key in candidate:
            derived[key] = candidate[key]
        elif alternate in candidate:
            derived[key] = candidate[alternate]
    derived["Xpath"] = [_normalize_xpath_step(step) for step in raw_xpath]
    return {key: derived.get(key, "" if key != "PARAMETERS" else {}) for key in LOCATION_KEYS}


def build_location(
    window_name: str = "",
    name: str = "",
    class_name: str = "",
    control_type: str = "",
    automation_id: str = "",
    found_index: int = 0,
    xpath: str = "",
    parameters: str = "",
) -> dict:
    """Build a LOCATION dict from individual parameters.

    Mirrors server.py's _build_location but lives here so both
    the MCP server and TCP service can reuse it without importing MCP.
    """
    return {
        "WindowName": window_name or "",
        "Name": name or "",
        "ClassName": class_name or "",
        "ControlType": control_type or "",
        "foundIndex": found_index or "",
        "AutomationId": automation_id or "",
        "Xpath": json.loads(xpath) if xpath else [],
        "Img": "",
        "PARAMETERS": json.loads(parameters) if parameters else {},
    }


def pack_response(
    req_id: str,
    result: Any = None,
    error: str | None = None,
    timing: float | None = None,
    elapsed: float | None = None,
) -> dict:
    """Build a JSON-compatible response dict."""
    resp: dict = {"id": req_id, "result": result, "error": error}
    if timing is not None:
        resp["timing_ms"] = round(timing * 1000, 1)
    if elapsed is not None:
        resp["timing_ms"] = round(elapsed * 1000, 1)
    return resp


def parse_params(params: dict, *keys: str) -> list:
    """Extract named keys from params dict, returning defaults if missing."""
    defaults = {
        "window_name": "",
        "name": "",
        "class_name": "",
        "control_type": "",
        "automation_id": "",
        "found_index": 0,
        "xpath": "",
        "button": "left",
        "click_type": "single",
        "x_offset": -1,
        "y_offset": -1,
        "text": "",
        "method": "clipboard",
        "max_depth": 5,
        "x": -1,
        "y": -1,
        "amount": 3,
        "direction": "down",
        "key": "",
        "keys": "",
        "region_x": 0,
        "region_y": 0,
        "region_width": 0,
        "region_height": 0,
    }
    return [params.get(k, defaults.get(k)) for k in keys]
