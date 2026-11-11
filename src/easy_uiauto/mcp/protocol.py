"""Protocol helpers for the UI automation TCP service.

Shared between service.py and client.py. Zero dependencies beyond stdlib.
"""

import ast
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


def normalize_match_mode(value: str) -> str:
    """Identity selects the first matching name; position selects a type ordinal."""
    if not isinstance(value, str) or value not in {"identity", "position"}:
        raise ValueError("match_mode must be 'identity' or 'position'")
    return value


def with_match_mode(value: dict, match_mode: str) -> dict:
    """Override a whole recorded path without changing legacy action signatures."""
    mode = normalize_match_mode(match_mode)
    location = normalize_location(value)
    parameters = location.get("PARAMETERS") or {}
    if isinstance(parameters, str):
        parameters = json.loads(parameters)
    location["PARAMETERS"] = {**parameters, "match_mode": mode}
    location["Xpath"] = [{**step, "match_mode": mode} for step in location["Xpath"]]
    return location


CONTROL_IDENTITY_KEYS = {"automation_id":"AutomationId", "name":"Name",
                         "class_name":"ClassName", "control_type":"ControlType"}


def with_control_identity(value: dict, identity: dict) -> dict:
    """Edit a target's criteria without rewriting its ancestors or match mode."""
    location = normalize_location(value)
    if not location["Xpath"]:
        raise ValueError("控件没有 XPath，请先在高级中补充定位路径")
    for field, text in identity.items():
        if field not in CONTROL_IDENTITY_KEYS or not isinstance(text,str):
            raise ValueError("控件属性必须为文本")
        if field=="control_type" and not text.strip():
            raise ValueError("ControlType 不能为空")
        key=CONTROL_IDENTITY_KEYS[field]
        location[key]=text
        if text:location["Xpath"][-1][key]=text
        else:location["Xpath"][-1].pop(key,None)
        if key=="Name" and len(location["Xpath"])==1:location["WindowName"]=text
    return location


def _normalize_xpath_step(step: dict) -> dict:
    """Convert recorded, MCP, or vector-store XPath keys to the core schema."""
    if not isinstance(step, dict):
        raise ValueError("Each XPath step must be an object")
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
    if "match_mode" in step or "MatchMode" in step:
        normalized["match_mode"] = normalize_match_mode(step.get("match_mode", step.get("MatchMode")))
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


def stable_location(value: dict) -> dict:
    """Preserve the recorded identity and explicit mode when storing knowledge."""
    return normalize_location(value)


def parse_xpath(value) -> list[dict]:
    """Accept JSON and legacy Python literals, without executing recorded text."""
    if value in (None, ""):
        return []
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            try:
                value = ast.literal_eval(value)
            except (ValueError, SyntaxError) as error:
                raise ValueError("location Xpath must be an array of objects") from error
    if not isinstance(value, list):
        raise ValueError("location Xpath must be an array of objects")
    return [_normalize_xpath_step(step) for step in value]


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

    raw_xpath = parse_xpath(candidate.get("Xpath", candidate.get("xpath", [])))

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
    mode = candidate.get("match_mode", candidate.get("MatchMode"))
    parameters = derived.get("PARAMETERS") or {}
    if isinstance(parameters, str):
        parameters = json.loads(parameters)
    if not isinstance(parameters, dict):
        raise ValueError("location PARAMETERS must be an object")
    if mode is not None:
        mode = normalize_match_mode(mode)
        parameters = {**parameters, "match_mode": mode}
        derived["Xpath"] = [{**step, "match_mode": mode} for step in derived["Xpath"]]
    elif "match_mode" in parameters:
        normalize_match_mode(parameters["match_mode"])
    derived["PARAMETERS"] = parameters
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
    match_mode: str | None = None,
) -> dict:
    """Build a LOCATION dict from individual parameters.

    Mirrors server.py's _build_location but lives here so both
    the MCP server and TCP service can reuse it without importing MCP.
    """
    result = {
        "WindowName": window_name or "",
        "Name": name or "",
        "ClassName": class_name or "",
        "ControlType": control_type or "",
        "foundIndex": found_index if isinstance(found_index, bool) or found_index else "",
        "AutomationId": automation_id or "",
        "Xpath": parse_xpath(xpath),
        "Img": "",
        "PARAMETERS": json.loads(parameters) if parameters else {},
    }
    return with_match_mode(result, match_mode) if match_mode is not None else result


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
