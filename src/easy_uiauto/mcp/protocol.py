"""Protocol helpers for the UI automation TCP service.

Shared between service.py and client.py. Zero dependencies beyond stdlib.
"""

import json
import time
from typing import Any, Optional


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
    error: Optional[str] = None,
    timing: Optional[float] = None,
    elapsed: Optional[float] = None,
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
