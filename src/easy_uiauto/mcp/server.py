"""easy_uiauto MCP server for Windows desktop UI automation."""

import argparse
import base64
import ctypes
import io
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from typing import Optional
from urllib import error as urlerror
from urllib import request as urlrequest

import pyautogui
import uiautomation
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp import Image as MCPImage

# easy-uiauto imports
from easy_uiauto.ctrl import Controller, run_action as _run_action_easy
from easy_uiauto.utils import (
    auto_scroll,
    find_control as _find_control,
    get_control_info,
    get_control_xpath,
    set_top_window,
)

from .. import __version__
from . import configuration, interaction_learning, knowledge, scanner, skill_installation, ui_cli
from .protocol import location_from_xpath, normalize_location

# Disable pyautogui failsafe for automation (moving mouse to corner won't crash)
pyautogui.FAILSAFE = False
# Disable pyautogui pause between actions for faster automation
pyautogui.PAUSE = 0.1

mcp = FastMCP(
    "easy_uiauto",
    instructions=(
        "Windows desktop UI automation server. "
        "Use these tools to inspect and interact with desktop applications. "
        "Start with list_windows to see open windows, "
        "then get_control_tree to explore a window's UI hierarchy, "
        "then use click/type_text/etc to interact with controls."
    ),
)

VALID_MODES = {"operate", "learn", "mixed"}
CURRENT_MODE = "operate"
LEARNING_LOG_PATH = os.environ.get(
    "EASY_UIAUTO_LEARNING_LOG",
    os.path.join(os.environ.get("TEMP", os.path.expanduser("~")), "easy_uiauto_learning.log"),
)
LEARNING_CONSOLE_STARTED = False
MUTATING_RECORDED_ACTIONS = {
    "点击", "按下鼠标左键", "释放鼠标左键", "右击", "按下鼠标右键", "释放鼠标右键",
    "中击", "双击", "设置文本", "输入文本", "键盘点击", "键盘按下", "键盘释放",
    "拖拽", "组合键",
}
MUTATING_METHODS = {
    "click", "click_at_position", "drag_control", "type_text", "set_text",
    "press_key", "hotkey", "run_action", "scroll", "mouse_scroll",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _build_location(
    window_name: str = "",
    name: str = "",
    class_name: str = "",
    control_type: str = "",
    automation_id: str = "",
    found_index: int = 0,
    xpath: str = "",
    parameters: str = "",
) -> dict:
    """Build a LOCATION dict from tool parameters."""
    location = {
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
    return location


def _screen_region(
    region_x: int,
    region_y: int,
    region_width: int,
    region_height: int,
) -> Optional[tuple[int, int, int, int]]:
    """Return a valid PyAutoGUI region, or None for the entire screen."""
    values = (region_x, region_y, region_width, region_height)
    if region_width == 0 and region_height == 0:
        return None
    if region_width <= 0 or region_height <= 0:
        raise ValueError("region_width and region_height must both be positive")
    return values


def _box_to_dict(box) -> dict:
    left, top, width, height = int(box.left), int(box.top), int(box.width), int(box.height)
    return {
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "center_x": left + width // 2,
        "center_y": top + height // 2,
    }


def _locate_image(
    image_path: str,
    confidence: float,
    grayscale: bool,
    region: Optional[tuple[int, int, int, int]],
) -> Optional[dict]:
    """Locate an image template on screen and return its bounds, if present."""
    if not os.path.isfile(image_path):
        raise ValueError(f"Image template does not exist: {image_path}")
    if not 0 < confidence <= 1:
        raise ValueError("confidence must be greater than 0 and no greater than 1")

    try:
        box = pyautogui.locateOnScreen(
            image_path,
            confidence=confidence,
            grayscale=grayscale,
            region=region,
        )
    except (TypeError, NotImplementedError) as error:
        raise RuntimeError(
            "Image matching with confidence requires the vision extra: "
            'pip install "easy-uiauto[mcp,vision]"'
        ) from error
    except Exception as error:
        image_not_found = getattr(pyautogui, "ImageNotFoundException", ())
        if image_not_found and isinstance(error, image_not_found):
            return None
        raise RuntimeError(f"Image matching failed: {error}") from error

    return _box_to_dict(box) if box else None


def _require_pytesseract():
    try:
        import pytesseract
    except ImportError as error:
        raise RuntimeError(
            "OCR requires the vision extra: pip install \"easy-uiauto[mcp,vision]\""
        ) from error
    return pytesseract


def _find_text_on_screen(
    text: str,
    language: str,
    match_mode: str,
    min_confidence: float,
    region: Optional[tuple[int, int, int, int]],
) -> Optional[dict]:
    """Locate a single OCR text box on screen."""
    if not text.strip():
        raise ValueError("text must not be empty")
    if match_mode not in {"contains", "exact"}:
        raise ValueError("match_mode must be 'contains' or 'exact'")
    if not 0 <= min_confidence <= 100:
        raise ValueError("min_confidence must be between 0 and 100")

    pytesseract = _require_pytesseract()
    try:
        screenshot = pyautogui.screenshot(region=region)
        data = pytesseract.image_to_data(
            screenshot,
            lang=language,
            output_type=pytesseract.Output.DICT,
        )
    except Exception as error:
        raise RuntimeError(
            "OCR failed. Install Tesseract OCR and ensure its executable is on PATH: "
            f"{error}"
        ) from error

    needle = text.casefold().strip()
    offset_x, offset_y = (region[0], region[1]) if region else (0, 0)
    for index, candidate in enumerate(data.get("text", [])):
        candidate = (candidate or "").strip()
        if not candidate:
            continue
        try:
            confidence = float(data["conf"][index])
        except (KeyError, TypeError, ValueError):
            continue
        haystack = candidate.casefold()
        matched = haystack == needle if match_mode == "exact" else needle in haystack
        if not matched or confidence < min_confidence:
            continue
        left = int(data["left"][index]) + offset_x
        top = int(data["top"][index]) + offset_y
        width = int(data["width"][index])
        height = int(data["height"][index])
        return {
            "text": candidate,
            "confidence": confidence,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "center_x": left + width // 2,
            "center_y": top + height // 2,
        }
    return None


def _vision_api_settings(model: str) -> tuple[str, str, str]:
    """Read current user settings so a running MCP sees setup changes immediately."""
    url = configuration._existing_vision_value(configuration.VISION_API_URL).strip()
    api_key = configuration._existing_vision_value(configuration.VISION_API_KEY).strip()
    configured_model = model.strip() or configuration._existing_vision_value(
        configuration.VISION_MODEL
    ).strip()
    if not url or not api_key or not configured_model:
        raise RuntimeError(
            "Remote vision requires EASY_UIAUTO_VISION_API_URL, "
            "EASY_UIAUTO_VISION_API_KEY, and EASY_UIAUTO_VISION_MODEL. Run "
            "easy_uiauto --quick-setup-codex --vision-url URL --vision-model MODEL. "
            "Do not retry with full-uia because it requires the same configuration."
        )
    return url, api_key, configured_model


def _vision_response_json(content: str) -> dict:
    """Extract the single JSON object required from a vision model response."""
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content, flags=re.IGNORECASE)
    try:
        result = json.loads(content)
    except json.JSONDecodeError as error:
        raise RuntimeError("Vision API did not return a JSON object") from error
    if not isinstance(result, dict):
        raise RuntimeError("Vision API response must be a JSON object")
    return result


def _find_control_by_vision(
    description: str,
    model: str,
    region: Optional[tuple[int, int, int, int]],
) -> dict:
    """Use a remote OpenAI-compatible vision API to locate one described control."""
    if not description.strip():
        raise ValueError("description must not be empty")
    url, api_key, configured_model = _vision_api_settings(model)
    screenshot = pyautogui.screenshot(region=region)
    width, height = screenshot.size
    buf = io.BytesIO()
    screenshot.save(buf, format="PNG")
    image_url = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
    prompt = (
        "Locate exactly one UI target in this screenshot. Target: "
        f"{description}\nThe image is {width}x{height} pixels. Return JSON only: "
        '{"found":true,"left":0,"top":0,"width":0,"height":0,"confidence":0.0}. '
        "Use image-pixel coordinates. If absent, return {\"found\":false}."
    )
    payload = {
        "model": configured_model,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": "You locate UI elements precisely and return JSON only."},
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
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": f"easy-uiauto/{__version__}",
        },
        method="POST",
    )
    try:
        with urlrequest.urlopen(request, timeout=45) as response:
            body = json.loads(response.read().decode("utf-8"))
    except urlerror.HTTPError as error:
        try:
            detail = json.loads(error.read().decode("utf-8")).get("error", {}).get("message", "")
        except Exception:
            detail = ""
        suffix = f": {detail}" if detail else ""
        raise RuntimeError(f"Vision API request failed with HTTP {error.code}{suffix}") from error
    except urlerror.URLError as error:
        raise RuntimeError(f"Vision API request failed: {error.reason}") from error

    try:
        content = body["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as error:
        raise RuntimeError("Vision API response did not contain a chat completion") from error
    if isinstance(content, list):
        content = "".join(item.get("text", "") for item in content if isinstance(item, dict))
    result = _vision_response_json(str(content))
    if not result.get("found", False):
        return {"found": False}

    try:
        left = int(result["left"])
        top = int(result["top"])
        box_width = int(result["width"])
        box_height = int(result["height"])
    except (KeyError, TypeError, ValueError) as error:
        raise RuntimeError("Vision API match must include integer left, top, width, and height") from error
    if left < 0 or top < 0 or box_width <= 0 or box_height <= 0 or left + box_width > width or top + box_height > height:
        raise RuntimeError("Vision API returned bounds outside the screenshot")
    offset_x, offset_y = (region[0], region[1]) if region else (0, 0)
    return {
        "found": True,
        "left": left + offset_x,
        "top": top + offset_y,
        "width": box_width,
        "height": box_height,
        "center_x": left + offset_x + box_width // 2,
        "center_y": top + offset_y + box_height // 2,
        "confidence": result.get("confidence"),
        "model": configured_model,
    }


def _mode_blocks_operation() -> Optional[str]:
    if CURRENT_MODE == "learn":
        return "Error: Current mode is learn; mutating UI operations are blocked"
    return None


def _learning_log(message: str) -> None:
    try:
        os.makedirs(os.path.dirname(LEARNING_LOG_PATH), exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LEARNING_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception:
        return


def _ensure_learning_console() -> None:
    global LEARNING_CONSOLE_STARTED
    if LEARNING_CONSOLE_STARTED:
        return
    try:
        os.makedirs(os.path.dirname(LEARNING_LOG_PATH), exist_ok=True)
        if not os.path.exists(LEARNING_LOG_PATH):
            with open(LEARNING_LOG_PATH, "w", encoding="utf-8") as f:
                f.write("")
        ps_cmd = (
            "$Host.UI.RawUI.WindowTitle='easy_uiauto learning console'; "
            f"Write-Host 'Learning log: {LEARNING_LOG_PATH}'; "
            f"Get-Content -LiteralPath '{LEARNING_LOG_PATH}' -Tail 40 -Wait"
        )
        subprocess.Popen(
            ["powershell.exe", "-NoExit", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd],
            creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
        )
        LEARNING_CONSOLE_STARTED = True
        _learning_log("learning console opened")
    except Exception as e:
        _learning_log(f"failed to open learning console: {e}")


def _result_ok(result) -> bool:
    return not _is_error_result(str(result))


def _auto_learn_location(action_title: str, location: dict, notes: str = "") -> None:
    if CURRENT_MODE != "mixed":
        return
    try:
        if not any(location.get(k) for k in ("Name", "ClassName", "ControlType", "AutomationId", "Xpath")):
            return
        ctrl = _find_control(location, debug=False)
        if not ctrl or ctrl is False:
            return
        xpath = location.get("Xpath") or get_control_xpath(ctrl)
        record = _record_from_control(
            ctrl,
            xpath,
            alias="",
            app_name="",
            notes=notes or f"Auto-learned from successful action: {action_title}",
            tags="auto_learn,mixed",
            source="mixed_auto_learn",
        )
        store = _load_control_vector_store()
        rid = store.upsert_control(record)
        _learning_log(
            f"auto-upsert id={rid} alias={record.get('alias')} "
            f"window={record.get('window_name')} type={record.get('control_type')} "
            f"name={record.get('name')} class={record.get('class_name')}"
        )
    except Exception:
        # Auto-learning must never make the UI action fail.
        return


def _auto_learn_point(action_title: str, x: int, y: int, notes: str = "") -> None:
    if CURRENT_MODE != "mixed":
        return
    try:
        xpath_list, ctrl = get_control_info(x, y)
        if not ctrl:
            return
        record = _record_from_control(
            ctrl,
            xpath_list,
            alias="",
            app_name="",
            notes=notes or f"Auto-learned from successful point action: {action_title}",
            tags="auto_learn,mixed,point",
            source="mixed_auto_learn",
        )
        record["capture_point"] = {"x": x, "y": y}
        store = _load_control_vector_store()
        rid = store.upsert_control(record)
        _learning_log(
            f"auto-upsert-point id={rid} alias={record.get('alias')} "
            f"point=({x},{y}) window={record.get('window_name')} "
            f"type={record.get('control_type')} name={record.get('name')}"
        )
    except Exception:
        return


def _recorded_action_is_mutating(action: dict) -> bool:
    return action.get("ACTION") in MUTATING_RECORDED_ACTIONS


def _auto_learn_recorded_action(action: dict) -> None:
    if CURRENT_MODE != "mixed":
        return
    try:
        location = action.get("LOCATION") or {}
        if isinstance(location, dict):
            _auto_learn_location(action.get("ACTION", "recorded action"), location)
    except Exception:
        return


def _rect_to_dict(rect) -> dict:
    if not rect:
        return {"left": 0, "top": 0, "right": 0, "bottom": 0, "width": 0, "height": 0}
    left = getattr(rect, "left", 0)
    top = getattr(rect, "top", 0)
    right = getattr(rect, "right", 0)
    bottom = getattr(rect, "bottom", 0)
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "width": max(0, right - left),
        "height": max(0, bottom - top),
    }


def _normalize_xpath_step(step: dict) -> dict:
    normalized = {
        "control_type": step.get("control_type") or step.get("ControlType") or "",
        "name": step.get("name") or step.get("Name") or "",
        "class_name": step.get("class_name") or step.get("ClassName") or "",
        "automation_id": step.get("automation_id") or step.get("AutomationId") or "",
    }
    found_index = step.get("found_index", step.get("foundIndex"))
    search_depth = step.get("search_depth", step.get("searchDepth"))
    if found_index not in (None, ""):
        normalized["found_index"] = found_index
    if search_depth not in (None, ""):
        normalized["search_depth"] = search_depth
    return normalized


def _record_from_control(ctrl, xpath_list: list, alias: str = "", app_name: str = "",
                         notes: str = "", tags: str = "", source: str = "") -> dict:
    xpath = [_normalize_xpath_step(step) for step in (xpath_list or [])]
    window_name = ""
    for step in xpath:
        if step["control_type"] == "WindowControl":
            window_name = step["name"]
            break

    name = ctrl.Name or ""
    class_name = ctrl.ClassName or ""
    control_type = ctrl.ControlTypeName if hasattr(ctrl, "ControlTypeName") else ""
    automation_id = ctrl.AutomationId or ""
    if not app_name:
        app_name = window_name or ""

    if not alias:
        parts = [app_name or "App", control_type or "Control", automation_id or name or class_name or "Unnamed"]
        alias = ".".join(_safe_alias_part(p) for p in parts if p)

    try:
        rect = _rect_to_dict(ctrl.BoundingRectangle)
    except Exception:
        rect = _rect_to_dict(None)

    return {
        "alias": alias,
        "name": name,
        "class_name": class_name,
        "control_type": control_type,
        "automation_id": automation_id,
        "framework_id": getattr(ctrl, "FrameworkId", "") or "",
        "window_name": window_name,
        "app_name": app_name,
        "notes": notes,
        "source": source or "easy_uiauto",
        "xpath": xpath,
        "rect": rect,
        "patterns": [],
        "tags": [t.strip() for t in tags.split(",") if t.strip()],
        "LOCATION": location_from_xpath(xpath_list),
    }


def _safe_alias_part(text: str) -> str:
    text = re.sub(r"\s+", "_", str(text).strip())
    text = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "_", text)
    return text.strip("_.") or "Control"


def _load_control_vector_store():
    module_dir = os.environ.get("EASY_UIAUTO_CONTROL_VECTOR_DB_DIR", "")
    if not module_dir:
        raise RuntimeError(
            "Control-vector storage is not configured. Set EASY_UIAUTO_CONTROL_VECTOR_DB_DIR "
            "to a directory containing control_vector_store.py."
        )
    if module_dir not in sys.path:
        sys.path.insert(0, module_dir)
    import control_vector_store  # type: ignore
    return control_vector_store


def _ctrl_pressed() -> bool:
    # VK_CONTROL=0x11, VK_LCONTROL=0xA2, VK_RCONTROL=0xA3.
    for vk in (0x11, 0xA2, 0xA3):
        if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
            return True
    return False


def _safe_call(func, *args, **kwargs) -> str:
    """Call a function and return result as string, catching exceptions."""
    try:
        result = func(*args, **kwargs)
        return str(result) if result is not None else "Success"
    except Exception as e:
        return f"Error: {type(e).__name__}: {e}"


def _is_error_result(result) -> bool:
    """Best-effort detection for handlers that report failures as strings."""
    return isinstance(result, str) and (
        result.lower().startswith("error")
        or "异常" in result
        or "未找到" in result
    )


def _normalize_batch_step(step: dict) -> tuple[str, dict]:
    """Normalize one batch step into (method, params).

    Supported forms:
      {"method": "click", "params": {...}}
      {"action": "click", ...params}
      {"ACTION": "点击", "LOCATION": {...}}  # easy-uiauto recorded action
    """
    if not isinstance(step, dict):
        raise ValueError("Each step must be a JSON object")

    if "ACTION" in step and "LOCATION" in step:
        return "run_action", {"action": step}

    method = step.get("method") or step.get("action") or step.get("tool")
    if not method:
        raise ValueError("Step is missing 'method' or 'action'")

    params = step.get("params")
    if params is None:
        params = {k: v for k, v in step.items() if k not in {"method", "action", "tool"}}
    if not isinstance(params, dict):
        raise ValueError("Step 'params' must be a JSON object")

    aliases = {
        "type": "type_text",
        "input_text": "type_text",
        "send_keys": "hotkey",
        "key": "press_key",
        "screenshot": "take_screenshot",
    }
    return aliases.get(str(method), str(method)), params


# ---------------------------------------------------------------------------
# Window Management
# ---------------------------------------------------------------------------

@mcp.tool()
def get_mode() -> str:
    """Return the current UI automation mode: operate, learn, or mixed."""
    return json.dumps({"mode": CURRENT_MODE}, ensure_ascii=False)


@mcp.tool()
def set_mode(mode: str) -> str:
    """Set the UI automation mode.

    Modes:
      - operate: allow UI operations, no automatic control-library writes.
      - learn: allow observation/capture/upsert only; block mutating UI operations.
      - mixed: allow UI operations and best-effort auto-learning of successful controls.
    """
    global CURRENT_MODE
    mode = (mode or "").strip().lower()
    if mode not in VALID_MODES:
        return f"Error: mode must be one of {sorted(VALID_MODES)}"
    CURRENT_MODE = mode
    _learning_log(f"mode set to {CURRENT_MODE}")
    if CURRENT_MODE in {"learn", "mixed"}:
        _ensure_learning_console()
    return json.dumps({"ok": True, "mode": CURRENT_MODE}, ensure_ascii=False)


@mcp.tool()
def list_windows() -> str:
    """List all open windows with their titles, positions, and sizes.

    Returns a JSON array of window objects, each with: title, left, top, width, height.
    Use this first to discover what windows are available.
    """
    try:
        windows = pyautogui.getAllWindows()
        result = []
        for w in windows:
            if w.title and w.visible:
                result.append({
                    "title": w.title,
                    "left": w.left,
                    "top": w.top,
                    "width": w.width,
                    "height": w.height,
                })
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return f"Error listing windows: {e}"


@mcp.tool()
def activate_window(title: str) -> str:
    """Bring a window to the front and maximize it.

    Args:
        title: The window title (or partial title match).
    """
    try:
        set_top_window(title)
        return f"Activated window: {title}"
    except Exception as e:
        return f"Error activating window: {e}"


# ---------------------------------------------------------------------------
# UI Inspection
# ---------------------------------------------------------------------------

@mcp.tool()
def get_control_tree(
    window_name: str,
    max_depth: int = 5,
) -> str:
    """Get the UI control tree of a window. Shows the hierarchy of UI elements.

    This is the primary way to explore what controls exist in a window.
    Returns a nested JSON structure of controls with their properties.

    Args:
        window_name: Title of the window to inspect.
        max_depth: How deep to explore the tree (default 5, max 10). Deeper = slower.
    """
    max_depth = min(max_depth, 10)

    try:
        set_top_window(window_name)

        # Find the top-level window control
        window = uiautomation.WindowControl(Name=window_name, searchDepth=1)
        if not window.Exists(maxSearchSeconds=3):
            # Try partial match
            window = uiautomation.WindowControl(searchFromControl=uiautomation.GetRootControl(), Name=window_name, searchDepth=5)
            if not window.Exists(maxSearchSeconds=3):
                return f"Error: Window '{window_name}' not found"

        def _control_to_dict(ctrl, depth=0):
            """Recursively build control tree dict."""
            if depth > max_depth:
                return None

            try:
                rect = ctrl.BoundingRectangle
                info = {
                    "ControlType": ctrl.ControlTypeName if hasattr(ctrl, 'ControlTypeName') else str(type(ctrl).__name__),
                    "Name": ctrl.Name or "",
                    "ClassName": ctrl.ClassName or "",
                    "AutomationId": ctrl.AutomationId or "",
                    "bounds": {
                        "left": rect.left if rect else 0,
                        "top": rect.top if rect else 0,
                        "right": rect.right if rect else 0,
                        "bottom": rect.bottom if rect else 0,
                    },
                }

                # Add children (limit to avoid explosion)
                children = []
                try:
                    for child in ctrl.GetChildren():
                        child_info = _control_to_dict(child, depth + 1)
                        if child_info:
                            children.append(child_info)
                        if len(children) >= 30:  # Limit children per node
                            children.append({"_note": f"... truncated, more children exist"})
                            break
                except Exception:
                    pass

                if children:
                    info["children"] = children

                return info
            except Exception:
                return None

        tree = _control_to_dict(window)
        return json.dumps(tree, ensure_ascii=False, indent=2)

    except Exception as e:
        return f"Error getting control tree: {e}"


@mcp.tool()
def find_control(
    location: Optional[dict] = None,
    window_name: str = "",
    name: str = "",
    class_name: str = "",
    control_type: str = "",
    automation_id: str = "",
    found_index: int = 0,
    xpath: str = "",
) -> str:
    """Find a control from a complete easy_uiauto LOCATION object.

    Prefer ``location`` from an easy_uiauto recorded action or from
    ``get_control_at_position``. The complete XPath retains hierarchy,
    duplicate indexes, and search depth. Legacy individual selector arguments
    remain available for compatibility.

    Args:
        location: Complete LOCATION object, a recorded action containing
            LOCATION, or a coordinate result containing LOCATION.
        window_name: Parent window title.
        name: Control name (UIA Name property).
        class_name: Control class name.
        control_type: Control type (e.g., "ButtonControl", "EditControl").
        automation_id: UIA AutomationId.
        found_index: 1-based index when multiple controls match (0 = first).
        xpath: JSON string of XPath path list, e.g. '[{"ControlType":"WindowControl","Name":"Notepad"}]'.
    """
    try:
        if location is not None:
            resolved_location = normalize_location(location)
        else:
            resolved_location = _build_location(
                window_name=window_name,
                name=name,
                class_name=class_name,
                control_type=control_type,
                automation_id=automation_id,
                found_index=found_index,
                xpath=xpath,
            )

        ctrl = _find_control(resolved_location, debug=False)
        if not ctrl or ctrl is False:
            return f"Error: Control not found with the given parameters"

        try:
            rect = ctrl.BoundingRectangle
            bounds = {
                "left": rect.left if rect else 0,
                "top": rect.top if rect else 0,
                "right": rect.right if rect else 0,
                "bottom": rect.bottom if rect else 0,
            }
        except Exception:
            bounds = {}

        result = {
            "Name": ctrl.Name or "",
            "ClassName": ctrl.ClassName or "",
            "ControlType": ctrl.ControlTypeName if hasattr(ctrl, 'ControlTypeName') else "",
            "AutomationId": ctrl.AutomationId or "",
            "bounds": bounds,
            "IsEnabled": getattr(ctrl, 'IsEnabled', True),
            "IsVisible": getattr(ctrl, 'IsVisible', True),
            "LOCATION": resolved_location,
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return f"Error finding control: {e}"


@mcp.tool()
def get_control_at_position(x: int, y: int) -> str:
    """Get information about the UI control at the given screen coordinates.

    Useful for exploring the UI by pointing at elements.

    Args:
        x: Screen X coordinate.
        y: Screen Y coordinate.
    """
    try:
        xpath_list, ctrl = get_control_info(x, y)
        if not ctrl:
            return f"No control found at ({x}, {y})"

        try:
            rect = ctrl.BoundingRectangle
            bounds = {
                "left": rect.left if rect else 0,
                "top": rect.top if rect else 0,
                "right": rect.right if rect else 0,
                "bottom": rect.bottom if rect else 0,
            }
        except Exception:
            bounds = {}

        location = location_from_xpath(xpath_list)
        result = {
            "Name": ctrl.Name or "",
            "ClassName": ctrl.ClassName or "",
            "ControlType": ctrl.ControlTypeName if hasattr(ctrl, 'ControlTypeName') else "",
            "AutomationId": ctrl.AutomationId or "",
            "bounds": bounds,
            "xpath": xpath_list,
            "LOCATION": location,
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return f"Error getting control at position: {e}"


@mcp.tool()
def capture_control_record_at_position(
    x: int,
    y: int,
    alias: str = "",
    app_name: str = "",
    notes: str = "",
    tags: str = "",
    source: str = "easy_uiauto",
) -> str:
    """Capture a control at screen coordinates as a control-vector-db record.

    The returned JSON can be reviewed manually, saved with upsert_control_record,
    or used by an AI workflow to build future easy_uiauto actions.

    Args:
        x: Screen X coordinate.
        y: Screen Y coordinate.
        alias: Stable control alias to store, e.g. MyApplication.NewProjectDialog.FinishOk.
        app_name: Optional app/domain name. Defaults to the owning window name.
        notes: Human-readable purpose or caveats.
        tags: Comma-separated search tags.
        source: Source label stored with the record.
    """
    try:
        xpath_list, ctrl = get_control_info(x, y)
        if not ctrl:
            return f"Error: No control found at ({x}, {y})"
        record = _record_from_control(ctrl, xpath_list, alias, app_name, notes, tags, source)
        _learning_log(
            f"captured point=({x},{y}) alias={record.get('alias')} "
            f"window={record.get('window_name')} type={record.get('control_type')} "
            f"name={record.get('name')} class={record.get('class_name')}"
        )
        return json.dumps(record, ensure_ascii=False, indent=2)
    except Exception as e:
        return f"Error capturing control record: {e}"


@mcp.tool()
def capture_control_record_on_ctrl(
    alias: str = "",
    app_name: str = "",
    notes: str = "",
    tags: str = "",
    source: str = "manual_ctrl_capture",
    timeout_ms: int = 10000,
) -> str:
    """Wait for Ctrl, then capture the control under the current mouse pointer.

    Usage: call this tool, move the mouse over a target control, press Ctrl,
    then review the returned control-vector-db JSON record.

    Args:
        alias: Stable control alias to store.
        app_name: Optional app/domain name. Defaults to the owning window name.
        notes: Human-readable purpose or caveats.
        tags: Comma-separated search tags.
        source: Source label stored with the record.
        timeout_ms: Maximum wait time for Ctrl.
    """
    try:
        deadline = time.time() + max(timeout_ms, 1) / 1000
        was_pressed = _ctrl_pressed()
        while time.time() < deadline:
            is_pressed = _ctrl_pressed()
            if is_pressed and not was_pressed:
                x, y = pyautogui.position()
                xpath_list, ctrl = get_control_info(x, y)
                if not ctrl:
                    return f"Error: No control found at current pointer ({x}, {y})"
                record = _record_from_control(ctrl, xpath_list, alias, app_name, notes, tags, source)
                record["capture_point"] = {"x": x, "y": y}
                _learning_log(
                    f"captured-on-ctrl point=({x},{y}) alias={record.get('alias')} "
                    f"window={record.get('window_name')} type={record.get('control_type')} "
                    f"name={record.get('name')} class={record.get('class_name')}"
                )
                return json.dumps(record, ensure_ascii=False, indent=2)
            was_pressed = is_pressed
            time.sleep(0.05)
        return f"Error: Timed out waiting for Ctrl after {timeout_ms} ms"
    except Exception as e:
        return f"Error capturing control record on Ctrl: {e}"


@mcp.tool()
def upsert_control_record(record_json: str) -> str:
    """Insert or update one control-vector-db record.

    The input must match the JSON shape returned by capture_control_record_at_position.
    It writes both SQLite and ChromaDB through control_vector_store.upsert_control.

    Args:
        record_json: Control record JSON object.
    """
    try:
        record = json.loads(record_json)
        if not isinstance(record, dict):
            return "Error: record_json must be a JSON object"
        if not record.get("alias"):
            return "Error: record_json must include alias"
        store = _load_control_vector_store()
        rid = store.upsert_control(record)
        _learning_log(
            f"manual-upsert id={rid} alias={record.get('alias')} "
            f"window={record.get('window_name')} type={record.get('control_type')} "
            f"name={record.get('name')} class={record.get('class_name')}"
        )
        return json.dumps({"ok": True, "id": rid, "alias": record.get("alias")}, ensure_ascii=False)
    except Exception as e:
        return f"Error upserting control record: {e}"


# ---------------------------------------------------------------------------
# Application knowledge and semantic UI CLI
# ---------------------------------------------------------------------------

@mcp.tool()
def get_ui_learning_readiness() -> str:
    """Check UI-learning configuration without networking or exposing credentials.

    Call this before scanning. Settings are read from the current Windows user
    environment on every request, so a running MCP process can see configuration
    changes without another restart.
    """
    try:
        status = configuration.vision_configuration_status()
        status.update(
            {
                "version": __version__,
                "knowledge_vault": str(knowledge.vault_root()),
                "known_applications": len(knowledge.list_apps()),
            }
        )
        return json.dumps(status, ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error checking UI learning readiness: {error}"


@mcp.tool()
def scan_window_knowledge(
    window_name: str,
    max_depth: int = 12,
    max_controls: int = 3000,
    verify_limit: int = 500,
    strategy: str = "visual-first",
) -> str:
    """Scan one application window into an Obsidian-compatible UI knowledge vault.

    The default visual-first strategy asks AI for only important controls and
    regions, resolves their pixel targets through UIA, selects a stable ancestor,
    and verifies LOCATION plus control images. Use full-uia only for diagnostics.

    Args:
        window_name: Exact or partial title of the visible application window.
        max_depth: Maximum UIA tree depth, from 1 to 30.
        max_controls: Hard cap for controls visited; the result reports truncation.
        verify_limit: Maximum actionable controls to verify during this scan.
        strategy: ``visual-first`` (default) or diagnostic ``full-uia``.
    """
    try:
        api_url, api_key, model = _vision_api_settings("")
        result = scanner.scan_window(
            window_name=window_name,
            api_url=api_url,
            api_key=api_key,
            model=model,
            version=__version__,
            max_depth=max_depth,
            max_controls=max_controls,
            verify_limit=verify_limit,
            strategy=strategy,
        )
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error scanning application knowledge: {error}"


@mcp.tool()
def list_ui_knowledge_apps() -> str:
    """List applications and locator/semantic verification counts."""
    try:
        return json.dumps(knowledge.list_apps(), ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error listing UI knowledge applications: {error}"


@mcp.tool()
def search_ui_knowledge(
    app_id: str,
    query: str = "",
    include_quarantine: bool = False,
    limit: int = 50,
) -> str:
    """Search an application's Markdown control knowledge.

    Args:
        app_id: Application identifier returned by scan_window_knowledge.
        query: Terms matching names, IDs, pages, regions, types, tags, or actions.
        include_quarantine: Include failed records for inspection and repair.
        limit: Maximum records returned.
    """
    try:
        directory = knowledge.app_dir(app_id)
        statuses = None
        if include_quarantine:
            statuses = {"verified", "observed", "suspect", "quarantined"}
        controls = knowledge.search_controls(directory, query, statuses, limit)
        return json.dumps(controls, ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error searching UI knowledge: {error}"


@mcp.tool()
def list_ui_commands(app_id: str, page_id: str = "") -> str:
    """List the verified semantic UI CLI commands for an application or page."""
    try:
        commands = knowledge.available_commands(knowledge.app_dir(app_id), page_id)
        return json.dumps(commands, ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error listing UI commands: {error}"


@mcp.tool()
def run_ui_command(
    app_id: str,
    command: str,
    text: str = "",
    confirm: bool = False,
    allow_vision_fallback: bool = False,
) -> str:
    """Run one verified application-specific UI command.

    Before execution, the saved LOCATION and control PNG are checked against
    the current UI. A stale or mismatched record is quarantined instead of being
        clicked. Use ``text`` for commands ending in ``.set-text``. Commands
        marked external or destructive require ``confirm=true``.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        return ui_cli.execute_json(
            knowledge.app_dir(app_id),
            command,
            text,
            confirm,
            allow_vision_fallback,
        )
    except Exception as error:
        return f"Error running UI command: {error}"


@mcp.tool()
def run_ui_commands(
    app_id: str,
    steps: list[dict | str],
    confirm: bool = False,
    allow_vision_fallback: bool = False,
) -> str:
    """Run a verified same-page UI command sequence with one shared preflight.

    Args:
        app_id: Application identifier returned by scan_window_knowledge.
        steps: Ordered command strings or objects with ``command`` and optional
            ``text``. The window is found and captured once, each unique control
            is verified before any action, and knowledge is written once per
            used control after execution.
        confirm: Explicit approval for all external or destructive commands.

    Split navigation or page-changing workflows into separate batches. If any
    preflight check fails, no action is executed. Execution stops on the first
    action error and reports the completed prefix.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        return ui_cli.execute_many_json(
            knowledge.app_dir(app_id),
            steps,
            confirm,
            allow_vision_fallback,
        )
    except Exception as error:
        return f"Error running UI command batch: {error}"


@mcp.tool()
def learn_ui_command_effect(
    app_id: str,
    command: str,
    text: str = "",
    confirm: bool = False,
    recover: bool = False,
    maximum_wait_seconds: float = 3.0,
) -> str:
    """Learn one command's before/after visual, window, and local UIA effects.

    Full target/desktop screenshots and top-level window inventories stay in the
    local vault. Only target-window before/after images are sent to the configured
    vision API. Difference boxes scope post-operation UIA inspection. Set
    ``recover=true`` only for reversible exploration; recovery currently uses Escape.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        api_url, api_key, model = _vision_api_settings("")
        result = interaction_learning.learn_command_effect(
            knowledge.app_dir(app_id),
            command,
            api_url,
            api_key,
            model,
            __version__,
            text=text,
            confirm=confirm,
            recover=recover,
            maximum_wait_seconds=maximum_wait_seconds,
        )
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error learning UI command effect: {error}"


@mcp.tool()
def explore_ui_workflows(
    app_id: str,
    policy: str = "safe",
    max_actions: int = 10,
    confirm: bool = False,
    max_depth: int = 3,
) -> str:
    """Interact with known reversible controls and learn their direct responses.

    ``safe`` executes only commands classified safe. ``supervised`` may also
    execute reversible state-changing commands, but external/destructive and
    confirmation-required commands are returned as pending and never executed.
    New pages and dialogs are scanned recursively up to ``max_depth`` before
    recovery. Scrolling and dragging are intentionally excluded.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        api_url, api_key, model = _vision_api_settings("")
        result = interaction_learning.explore_application(
            knowledge.app_dir(app_id),
            api_url,
            api_key,
            model,
            __version__,
            policy=policy,
            max_actions=max_actions,
            confirm=confirm,
            max_depth=max_depth,
        )
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error exploring UI workflows: {error}"


@mcp.tool()
def list_ui_interactions(app_id: str, command: str = "", limit: int = 50) -> str:
    """List learned before/after operation-effect records."""
    try:
        records = knowledge.list_interactions(knowledge.app_dir(app_id), command, limit)
        return json.dumps(records, ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error listing UI interactions: {error}"


@mcp.tool()
def teach_ui_control(
    app_id: str,
    control_id: str,
    semantic_name: str,
    intent: str,
    description: str,
    actions: str = "",
    aliases: str = "",
    risk: str = "safe",
    requires_confirmation: bool = False,
) -> str:
    """Teach or correct one control's real function using human-confirmed semantics.

    This updates semantic fields only. It never bypasses LOCATION/image checks;
    controls with invalid positioning remain quarantined. ``actions`` and
    ``aliases`` are comma-separated values.
    """
    try:
        selected_actions = [value.strip() for value in actions.split(",") if value.strip()]
        record = knowledge.teach_control(
            knowledge.app_dir(app_id),
            control_id,
            semantic_name,
            intent,
            description,
            selected_actions or None,
            [value.strip() for value in aliases.split(",") if value.strip()],
            risk,
            requires_confirmation,
        )
        return json.dumps(record, ensure_ascii=False, indent=2)
    except Exception as error:
        return f"Error teaching UI control: {error}"


@mcp.tool()
def rebuild_ui_knowledge_index(app_id: str) -> str:
    """Rebuild the disposable JSON index and UI CLI catalog from Markdown."""
    try:
        directory = knowledge.app_dir(app_id)
        index = knowledge.rebuild_index(directory)
        catalog = knowledge.write_command_catalog(directory)
        return json.dumps(
            {
                "ok": True,
                "app_id": app_id,
                "controls": len(index["controls"]),
                "commands": len(knowledge.available_commands(directory)),
                "catalog": str(catalog),
            },
            ensure_ascii=False,
            indent=2,
        )
    except Exception as error:
        return f"Error rebuilding UI knowledge index: {error}"


# ---------------------------------------------------------------------------
# Mouse Operations
# ---------------------------------------------------------------------------

@mcp.tool()
def click(
    window_name: str = "",
    name: str = "",
    class_name: str = "",
    control_type: str = "",
    automation_id: str = "",
    found_index: int = 0,
    xpath: str = "",
    button: str = "left",
    click_type: str = "single",
    x_offset: int = -1,
    y_offset: int = -1,
) -> str:
    """Click on a UI control.

    Locate a control by its properties and click on it. Use get_control_tree
    first to discover available controls.

    Args:
        window_name: Parent window title.
        name: Control name.
        class_name: Control class name.
        control_type: Control type (e.g., "ButtonControl").
        automation_id: UIA AutomationId.
        found_index: 1-based index for duplicate controls (0 = first).
        xpath: JSON XPath path string.
        button: Mouse button - "left", "right", or "middle".
        click_type: Click type - "single", "double".
        x_offset: X offset within control (-1 = center).
        y_offset: Y offset within control (-1 = center).
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        action_title = f"Click {name or control_type or 'control'}"
        params = json.dumps({"x": x_offset, "y": y_offset})
        location = _build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath, parameters=params,
        )

        if click_type == "double":
            result = Controller.double_click(
                ActionTitle=action_title, **{k: location[k] for k in location}
            )
        elif button == "right":
            result = Controller.right_click(
                ActionTitle=action_title, **{k: location[k] for k in location}
            )
        elif button == "middle":
            result = Controller.centre_click(
                ActionTitle=action_title, **{k: location[k] for k in location}
            )
        else:
            result = Controller.left_click(
                ActionTitle=action_title, **{k: location[k] for k in location}
            )

        if _result_ok(result):
            _auto_learn_location(action_title, location)
        return str(result)

    except Exception as e:
        return f"Error clicking: {e}"


@mcp.tool()
def click_at_position(
    x: int,
    y: int,
    button: str = "left",
    click_type: str = "single",
) -> str:
    """Click at absolute screen coordinates.

    Args:
        x: Screen X coordinate.
        y: Screen Y coordinate.
        button: "left", "right", or "middle".
        click_type: "single" or "double".
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        if click_type == "double":
            pyautogui.doubleClick(x, y, button=button)
        else:
            pyautogui.click(x, y, button=button)
        _auto_learn_point(f"Click at ({x}, {y})", x, y)
        return f"Clicked ({button}/{click_type}) at ({x}, {y})"
    except Exception as e:
        return f"Error clicking at position: {e}"


@mcp.tool()
def move_mouse(x: int, y: int) -> str:
    """Move the mouse cursor to the given screen coordinates.

    Args:
        x: Screen X coordinate.
        y: Screen Y coordinate.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        pyautogui.moveTo(x, y)
        return f"Mouse moved to ({x}, {y})"
    except Exception as e:
        return f"Error moving mouse: {e}"


@mcp.tool()
def drag_control(
    source_window: str,
    source_name: str = "",
    source_class: str = "",
    source_type: str = "",
    source_automation_id: str = "",
    source_xpath: str = "",
    target_window: str = "",
    target_name: str = "",
    target_class: str = "",
    target_type: str = "",
    target_automation_id: str = "",
    target_xpath: str = "",
) -> str:
    """Drag from one control to another.

    Args:
        source_window: Source control's window title.
        source_name: Source control name.
        source_class: Source control class name.
        source_type: Source control type.
        source_automation_id: Source AutomationId.
        source_xpath: Source XPath JSON string.
        target_window: Target control's window title (defaults to source_window).
        target_name: Target control name.
        target_class: Target control class name.
        target_type: Target control type.
        target_automation_id: Target AutomationId.
        target_xpath: Target XPath JSON string.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        params = json.dumps({
            "目的控件父窗口名称": target_window or source_window,
            "目的控件Name": target_name,
            "目的控件ClassName": target_class,
            "目的控件ControlType": target_type,
            "目的控件foundIndex": "",
            "目的控件AutomationId": target_automation_id,
            "目的控件Xpath": json.loads(target_xpath) if target_xpath else "",
        })

        result = Controller.drag_control(
            ActionTitle=f"Drag {source_name} to {target_name}",
            WindowName=source_window,
            Name=source_name,
            ClassName=source_class,
            ControlType=source_type,
            foundIndex="",
            AutomationId=source_automation_id,
            Xpath=json.loads(source_xpath) if source_xpath else "",
            Img="",
            PARAMETERS=json.loads(params),
        )
        if _result_ok(result):
            source_location = _build_location(
                window_name=source_window,
                name=source_name,
                class_name=source_class,
                control_type=source_type,
                automation_id=source_automation_id,
                xpath=source_xpath,
            )
            _auto_learn_location(f"Drag {source_name}", source_location)
        return str(result)
    except Exception as e:
        return f"Error dragging: {e}"


@mcp.tool()
def scroll(
    amount: int = 3,
    direction: str = "down",
    x: int = -1,
    y: int = -1,
) -> str:
    """Scroll the mouse wheel.

    Args:
        amount: Number of scroll ticks (default 3).
        direction: "up" or "down".
        x: Screen X coordinate to scroll at (-1 = current position).
        y: Screen Y coordinate to scroll at (-1 = current position).
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        if x >= 0 and y >= 0:
            pyautogui.moveTo(x, y)
        auto_scroll(amount, direction=direction)
        return f"Scrolled {direction} {amount} ticks"
    except Exception as e:
        return f"Error scrolling: {e}"


# ---------------------------------------------------------------------------
# Keyboard Operations
# ---------------------------------------------------------------------------

@mcp.tool()
def type_text(
    text: str,
    window_name: str = "",
    name: str = "",
    class_name: str = "",
    control_type: str = "",
    automation_id: str = "",
    found_index: int = 0,
    xpath: str = "",
    method: str = "clipboard",
) -> str:
    """Type text into a control.

    If control parameters are provided, focuses the control first.
    Uses clipboard paste by default (works with most apps including rich text).

    Args:
        text: The text to type.
        window_name: Target window title (optional, focuses window first).
        name: Target control name (optional).
        class_name: Target control class name.
        control_type: Target control type (e.g., "EditControl").
        automation_id: Target AutomationId.
        found_index: 1-based index for duplicate controls.
        xpath: Target XPath JSON string.
        method: "clipboard" (default, paste via Ctrl+V) or "sendkeys" (direct SendKeys).
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        if method == "clipboard":
            location = _build_location(
                window_name=window_name, name=name, class_name=class_name,
                control_type=control_type, automation_id=automation_id,
                found_index=found_index, xpath=xpath,
            )
            result = Controller.input_text(
                ActionTitle=f"Type text into {name or control_type or 'control'}",
                WindowName=window_name,
                Name=name,
                ClassName=class_name,
                ControlType=control_type,
                foundIndex=found_index or "",
                AutomationId=automation_id,
                Xpath=location["Xpath"],
                Img="",
                PARAMETERS={"输入文本": text},
            )
        else:
            location = _build_location(
                window_name=window_name, name=name, class_name=class_name,
                control_type=control_type, automation_id=automation_id,
                found_index=found_index, xpath=xpath,
            )
            result = Controller.set_text(
                ActionTitle=f"Set text on {name or control_type or 'control'}",
                WindowName=window_name,
                Name=name,
                ClassName=class_name,
                ControlType=control_type,
                foundIndex=found_index or "",
                AutomationId=automation_id,
                Xpath=location["Xpath"],
                Img="",
                PARAMETERS={"设置文本": text},
            )
        if _result_ok(result):
            _auto_learn_location(f"Type text into {name or control_type or 'control'}", location)
        return str(result)
    except Exception as e:
        return f"Error typing text: {e}"


@mcp.tool()
def set_text(
    text: str,
    window_name: str = "",
    name: str = "",
    class_name: str = "",
    control_type: str = "EditControl",
    automation_id: str = "",
    found_index: int = 0,
    xpath: str = "",
) -> str:
    """Set text directly on a control via SendKeys (replaces existing text).

    Unlike type_text which pastes from clipboard, this sends keys directly.
    Works better for simple text fields.

    Args:
        text: The text to set.
        window_name: Target window title.
        name: Target control name.
        class_name: Target control class name.
        control_type: Target control type.
        automation_id: Target AutomationId.
        found_index: 1-based index for duplicate controls.
        xpath: Target XPath JSON string.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        location = _build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath,
        )
        result = Controller.set_text(
            ActionTitle=f"Set text on {name or control_type}",
            WindowName=window_name,
            Name=name,
            ClassName=class_name,
            ControlType=control_type,
            foundIndex=found_index or "",
            AutomationId=automation_id,
            Xpath=location["Xpath"],
            Img="",
            PARAMETERS={"设置文本": text},
        )
        if _result_ok(result):
            _auto_learn_location(f"Set text on {name or control_type}", location)
        return str(result)
    except Exception as e:
        return f"Error setting text: {e}"


@mcp.tool()
def press_key(
    key: str,
    window_name: str = "",
) -> str:
    """Press a keyboard key.

    Args:
        key: Key name, e.g. "enter", "tab", "escape", "f1", "backspace",
             "delete", "up", "down", "left", "right", "home", "end",
             "page_up", "page_down", "space", "a"-"z", "0"-"9".
        window_name: Optional window to activate first.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        if window_name:
            set_top_window(window_name)

        result = Controller.key_click(
            ActionTitle=f"Press {key}",
            WindowName=window_name or "",
            Name="",
            ClassName="",
            ControlType="",
            foundIndex="",
            AutomationId="",
            Xpath=[],
            Img="",
            PARAMETERS={"键盘按键": key},
        )
        return str(result)
    except Exception as e:
        return f"Error pressing key: {e}"


@mcp.tool()
def hotkey(
    keys: str,
    window_name: str = "",
) -> str:
    """Send a keyboard shortcut / hotkey combination.

    Args:
        keys: Key combination with '+' separator, e.g. "ctrl+s", "ctrl+shift+n",
              "alt+f4", "ctrl+c", "ctrl+v". Use "ctrl_l"/"ctrl_r" for left/right ctrl,
              "alt_l"/"alt_gr" for left/right alt, "shift_l"/"shift_r" for left/right shift.
        window_name: Optional window to activate first.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        if window_name:
            set_top_window(window_name)

        result = Controller.key_group(
            ActionTitle=f"Hotkey {keys}",
            WindowName=window_name or "",
            Name="",
            ClassName="",
            ControlType="",
            foundIndex="",
            AutomationId="",
            Xpath=[],
            Img="",
            PARAMETERS={"组合键": keys},
        )
        return str(result)
    except Exception as e:
        return f"Error sending hotkey: {e}"


# ---------------------------------------------------------------------------
# Batch Operations
# ---------------------------------------------------------------------------

@mcp.tool()
def run_action(action_json: str) -> str:
    """Execute one easy-uiauto recorded action.

    Args:
        action_json: JSON object with easy-uiauto recorded action fields,
            e.g. {"TEST_ID":"1","ACTION":"点击","LOCATION":{...}}.
    """
    try:
        action = json.loads(action_json)
        if not isinstance(action, dict):
            return "Error: action_json must be a JSON object"
        if CURRENT_MODE == "learn" and _recorded_action_is_mutating(action):
            return "Error: Current mode is learn; mutating recorded actions are blocked"
        result = _run_action_easy(action)
        if _result_ok(result):
            _auto_learn_recorded_action(action)
        return str(result)
    except Exception as e:
        return f"Error running action: {e}"


@mcp.tool()
def run_actions(
    actions_json: str,
    stop_on_error: bool = True,
    delay_ms: int = 0,
) -> str:
    """Run multiple UI automation actions in one MCP call.

    Each item can be either:
      - {"method":"click","params":{"window_name":"Notepad","name":"OK"}}
      - {"action":"type_text","text":"hello","window_name":"Notepad"}
      - {"ACTION":"点击","LOCATION":{...}} for easy-uiauto recorded actions

    Args:
        actions_json: JSON array of action objects.
        stop_on_error: Stop after the first failed step (default true).
        delay_ms: Optional delay between steps.
    """
    try:
        actions = json.loads(actions_json)
    except Exception as e:
        return f"Error: actions_json is not valid JSON: {e}"

    if not isinstance(actions, list):
        return "Error: actions_json must be a JSON array"

    handlers = {
        "get_mode": get_mode,
        "set_mode": set_mode,
        "activate_window": activate_window,
        "click": click,
        "click_at_position": click_at_position,
        "drag_control": drag_control,
        "find_control": find_control,
        "get_control_at_position": get_control_at_position,
        "get_control_tree": get_control_tree,
        "hotkey": hotkey,
        "list_windows": list_windows,
        "mouse_scroll": mouse_scroll,
        "move_mouse": move_mouse,
        "press_key": press_key,
        "scroll": scroll,
        "set_text": set_text,
        "type_text": type_text,
    }

    results = []
    for index, step in enumerate(actions, start=1):
        t0 = time.perf_counter()
        try:
            method, params = _normalize_batch_step(step)
            if CURRENT_MODE == "learn" and (
                method in MUTATING_METHODS
                or (method == "run_action" and _recorded_action_is_mutating(params["action"]))
            ):
                elapsed = round((time.perf_counter() - t0) * 1000, 1)
                results.append({
                    "index": index,
                    "method": method,
                    "ok": False,
                    "result": "Error: Current mode is learn; mutating UI operations are blocked",
                    "timing_ms": elapsed,
                })
                if stop_on_error:
                    break
                continue
            if method == "run_action":
                result = _run_action_easy(params["action"])
            else:
                handler = handlers.get(method)
                if handler is None:
                    raise ValueError(f"Unsupported batch method: {method}")
                result = handler(**params)

            elapsed = round((time.perf_counter() - t0) * 1000, 1)
            ok = not _is_error_result(result)
            item = {
                "index": index,
                "method": method,
                "ok": ok,
                "result": str(result) if result is not None else "Success",
                "timing_ms": elapsed,
            }
            results.append(item)

            if method == "run_action" and ok:
                _auto_learn_recorded_action(params["action"])

            if delay_ms > 0 and index < len(actions):
                time.sleep(delay_ms / 1000)
            if stop_on_error and not ok:
                break
        except Exception as e:
            elapsed = round((time.perf_counter() - t0) * 1000, 1)
            results.append({
                "index": index,
                "method": step.get("method") or step.get("action") or step.get("ACTION", ""),
                "ok": False,
                "error": f"{type(e).__name__}: {e}",
                "timing_ms": elapsed,
            })
            if stop_on_error:
                break

    summary = {
        "ok": all(item.get("ok", False) for item in results) and len(results) == len(actions),
        "total": len(actions),
        "executed": len(results),
        "results": results,
    }
    return json.dumps(summary, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Screenshot
# ---------------------------------------------------------------------------

@mcp.tool()
def take_screenshot(
    region_x: int = 0,
    region_y: int = 0,
    region_width: int = 0,
    region_height: int = 0,
) -> MCPImage:
    """Take a screenshot of the current screen.

    Returns the screenshot as an image. With no arguments, captures the full screen.
    Optionally capture a specific region.

    Args:
        region_x: X coordinate of region top-left (0 = left edge).
        region_y: Y coordinate of region top-left (0 = top edge).
        region_width: Width of region (0 = full screen width).
        region_height: Height of region (0 = full screen height).
    """
    try:
        if region_width > 0 and region_height > 0:
            screenshot = pyautogui.screenshot(region=(region_x, region_y, region_width, region_height))
        else:
            screenshot = pyautogui.screenshot()

        # Convert to bytes
        buf = io.BytesIO()
        screenshot.save(buf, format="PNG")
        return MCPImage(data=buf.getvalue(), format="png")

    except Exception as e:
        # Return error as text - but MCPImage doesn't support that, so raise
        raise RuntimeError(f"Error taking screenshot: {e}")


# ---------------------------------------------------------------------------
# Vision fallback: image templates and OCR text
# ---------------------------------------------------------------------------

@mcp.tool()
def find_control_by_image(
    image_path: str,
    confidence: float = 0.85,
    grayscale: bool = False,
    region_x: int = 0,
    region_y: int = 0,
    region_width: int = 0,
    region_height: int = 0,
) -> str:
    """Find an on-screen control from an image template.

    Requires ``pip install \"easy-uiauto[mcp,vision]\"``. Use this only when
    UI Automation properties cannot identify the target. Returns matching bounds
    and center coordinates, or ``found: false`` when no match is present.
    """
    try:
        region = _screen_region(region_x, region_y, region_width, region_height)
        match = _locate_image(image_path, confidence, grayscale, region)
        return json.dumps({"found": match is not None, "match": match}, ensure_ascii=False)
    except Exception as error:
        return f"Error finding image: {error}"


@mcp.tool()
def click_by_image(
    image_path: str,
    confidence: float = 0.85,
    grayscale: bool = False,
    region_x: int = 0,
    region_y: int = 0,
    region_width: int = 0,
    region_height: int = 0,
    button: str = "left",
    click_type: str = "single",
) -> str:
    """Find an image template on screen and click its center.

    Requires ``pip install \"easy-uiauto[mcp,vision]\"``. Prefer UIA control
    selection when available; image matching is a visual fallback.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        region = _screen_region(region_x, region_y, region_width, region_height)
        match = _locate_image(image_path, confidence, grayscale, region)
        if match is None:
            return "Error: Image template was not found on screen"
        if click_type == "double":
            pyautogui.doubleClick(match["center_x"], match["center_y"], button=button)
        else:
            pyautogui.click(match["center_x"], match["center_y"], button=button)
        _auto_learn_point("Click image template", match["center_x"], match["center_y"])
        return json.dumps({"ok": True, "match": match}, ensure_ascii=False)
    except Exception as error:
        return f"Error clicking image: {error}"


@mcp.tool()
def find_text_on_screen(
    text: str,
    language: str = "eng",
    match_mode: str = "contains",
    min_confidence: float = 60.0,
    region_x: int = 0,
    region_y: int = 0,
    region_width: int = 0,
    region_height: int = 0,
) -> str:
    """Locate visible text using OCR.

    Requires ``pip install \"easy-uiauto[mcp,vision]\"`` and a local Tesseract
    OCR installation. ``language`` is a Tesseract language code, such as ``eng``
    or ``chi_sim``. Returns matching bounds and center coordinates.
    """
    try:
        region = _screen_region(region_x, region_y, region_width, region_height)
        match = _find_text_on_screen(text, language, match_mode, min_confidence, region)
        return json.dumps({"found": match is not None, "match": match}, ensure_ascii=False)
    except Exception as error:
        return f"Error finding text with OCR: {error}"


@mcp.tool()
def click_text_on_screen(
    text: str,
    language: str = "eng",
    match_mode: str = "contains",
    min_confidence: float = 60.0,
    region_x: int = 0,
    region_y: int = 0,
    region_width: int = 0,
    region_height: int = 0,
    button: str = "left",
    click_type: str = "single",
) -> str:
    """Locate visible text through OCR and click its center.

    Requires the same vision dependencies as :func:`find_text_on_screen`.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        region = _screen_region(region_x, region_y, region_width, region_height)
        match = _find_text_on_screen(text, language, match_mode, min_confidence, region)
        if match is None:
            return f"Error: OCR text was not found: {text}"
        if click_type == "double":
            pyautogui.doubleClick(match["center_x"], match["center_y"], button=button)
        else:
            pyautogui.click(match["center_x"], match["center_y"], button=button)
        _auto_learn_point("Click OCR text", match["center_x"], match["center_y"])
        return json.dumps({"ok": True, "match": match}, ensure_ascii=False)
    except Exception as error:
        return f"Error clicking OCR text: {error}"


@mcp.tool()
def find_control_by_vision(
    description: str,
    model: str = "",
    region_x: int = 0,
    region_y: int = 0,
    region_width: int = 0,
    region_height: int = 0,
) -> str:
    """Locate a control with a remote multimodal API.

    The screenshot is sent only to the API configured through
    ``EASY_UIAUTO_VISION_API_URL``, ``EASY_UIAUTO_VISION_API_KEY``, and
    ``EASY_UIAUTO_VISION_MODEL``. The endpoint must support OpenAI-compatible
    chat completions with image URLs. Use UIA, OCR, or template matching first
    whenever a deterministic selector is available.
    """
    try:
        region = _screen_region(region_x, region_y, region_width, region_height)
        return json.dumps(_find_control_by_vision(description, model, region), ensure_ascii=False)
    except Exception as error:
        return f"Error finding control with vision API: {error}"


@mcp.tool()
def click_by_vision(
    description: str,
    model: str = "",
    region_x: int = 0,
    region_y: int = 0,
    region_width: int = 0,
    region_height: int = 0,
    button: str = "left",
    click_type: str = "single",
) -> str:
    """Locate a control with a remote multimodal API and click its center.

    This is a final visual fallback. It sends a screenshot to the configured
    remote API, so do not use it for sensitive screens unless that API is approved.
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        region = _screen_region(region_x, region_y, region_width, region_height)
        match = _find_control_by_vision(description, model, region)
        if not match["found"]:
            return "Error: Vision API did not find the requested control"
        if click_type == "double":
            pyautogui.doubleClick(match["center_x"], match["center_y"], button=button)
        else:
            pyautogui.click(match["center_x"], match["center_y"], button=button)
        _auto_learn_point("Click vision API match", match["center_x"], match["center_y"])
        return json.dumps({"ok": True, "match": match}, ensure_ascii=False)
    except Exception as error:
        return f"Error clicking control with vision API: {error}"


# ---------------------------------------------------------------------------
# Scroll helper (raw mouse scroll without control targeting)
# ---------------------------------------------------------------------------

@mcp.tool()
def mouse_scroll(amount: int = 3, direction: str = "down") -> str:
    """Scroll the mouse wheel at the current cursor position.

    Args:
        amount: Number of scroll ticks (default 3).
        direction: "up" or "down".
    """
    try:
        blocked = _mode_blocks_operation()
        if blocked:
            return blocked
        if direction == "up":
            pyautogui.scroll(amount)
        else:
            pyautogui.scroll(-amount)
        return f"Scrolled {direction} {amount} ticks"
    except Exception as e:
        return f"Error scrolling: {e}"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="easy_uiauto",
        description="easy_uiauto MCP server for Windows desktop UI automation.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Start the MCP server:\n"
            "  easy_uiauto\n"
            "  python -m easy_uiauto.mcp.server\n\n"
            "Control location workflow (preferred):\n"
            "  1. Call get_control_at_position(x, y), or record an action with run_record().\n"
            "  2. Read the returned or recorded LOCATION object.\n"
            "  3. Call find_control(location=LOCATION).\n"
            "  LOCATION fields: WindowName, Name, ClassName, ControlType, foundIndex,\n"
            "                   AutomationId, Xpath, Img, PARAMETERS.\n"
            "  Keep the complete Xpath for duplicate or deeply nested controls.\n\n"
            "Record and replay:\n"
            "  Python: from easy_uiauto.record import run_record\n"
            "          run_record(write_file=True), then press ESC to stop.\n"
            "  MCP:    run_action(action_json='<recorded action JSON>')\n"
            "  Batch:  run_actions(actions_json='[<action JSON>, ...]')\n\n"
            "Operating modes:\n"
            "  operate: execute UI operations\n"
            "  learn:   capture and store controls without mutating the UI\n"
            "  mixed:   execute operations and learn successful controls\n\n"
            "Visual fallback order:\n"
            "  UIA LOCATION -> multi-state image templates -> OCR -> opt-in remote AI vision\n"
            "  OCR requires Tesseract and easy-uiauto[mcp,vision].\n"
            "  Remote AI vision requires EASY_UIAUTO_VISION_API_URL,\n"
            "  EASY_UIAUTO_VISION_API_KEY, and EASY_UIAUTO_VISION_MODEL.\n"
            "  Remote vision uploads the selected screenshot to the configured API.\n\n"
            "Client configuration:\n"
            "  Fast AI setup (prompts securely for a missing API key):\n"
            "    easy_uiauto --quick-setup-codex --vision-url <URL> --vision-model <MODEL>\n"
            "  Full setup and validation (Python vision deps, Tesseract, UIA, OCR, AI):\n"
            "    easy_uiauto --full-setup-codex --vision-url <URL> --vision-model <MODEL>\n"
            "  easy_uiauto --install-codex | --show-codex-config | --uninstall-codex\n"
            "  easy_uiauto --install-codex-skills | --uninstall-codex-skills\n"
            "  easy_uiauto --install-claude-code | --show-claude-code-config\n"
            "              | --uninstall-claude-code\n"
            "  Restart the client after changing its MCP configuration.\n\n"
            "Testing and diagnostics:\n"
            "  Installation: easy_uiauto --version\n"
            "                easy_uiauto --help\n"
            "  MCP config:   easy_uiauto --show-codex-config\n"
            "                easy_uiauto --show-claude-code-config\n"
            "  Read-only MCP smoke test:\n"
            "    list_windows -> get_control_at_position -> find_control(location=LOCATION)\n"
            "  Source checkout test suite:\n"
            "    python -m pip install -e \".[dev,mcp,vision]\"\n"
            "    python -m pytest -q\n"
            "  OCR smoke test: find_text_on_screen(text=..., language='eng')\n"
            "  AI vision test: find_control_by_vision(description=...)\n"
            "  AI vision sends the selected screenshot to the configured remote API.\n\n"
            "Application knowledge CLI:\n"
            "  easy_uiauto_ui scan <window> [--strategy visual-first|full-uia]\n"
            "  easy_uiauto_ui apps | commands <app-id> | search <app-id> [query]\n"
            "  easy_uiauto_ui run <app-id> <command> [--text TEXT] [--confirm]\n"
            "  easy_uiauto_ui batch <app-id> '<steps-json>' [--confirm]\n"
            "  easy_uiauto_ui learn-effect <app-id> <command> [--recover]\n"
            "  easy_uiauto_ui explore <app-id> [--policy safe|supervised] [--max-depth N]\n"
            "  easy_uiauto_ui interactions <app-id>\n"
            "  Runtime remote vision is disabled unless --allow-vision-fallback is set.\n"
            "  Automated learning intentionally excludes scrolling and dragging.\n\n"
            "MCP tools include:\n"
            "  modes: get_mode, set_mode\n"
            "  windows: list_windows, activate_window\n"
            "  inspection: get_control_tree, find_control, get_control_at_position\n"
            "  learning: capture_control_record_at_position, capture_control_record_on_ctrl, upsert_control_record\n"
            "  mouse: click, click_at_position, move_mouse, drag_control, scroll, mouse_scroll\n"
            "  keyboard: type_text, set_text, press_key, hotkey\n"
            "  batch: run_action, run_actions\n"
            "  screenshot: take_screenshot\n"
            "  knowledge: get_ui_learning_readiness, scan_window_knowledge,\n"
            "             list_ui_knowledge_apps, search_ui_knowledge,\n"
            "             list_ui_commands, run_ui_command, run_ui_commands,\n"
            "             learn_ui_command_effect, explore_ui_workflows, list_ui_interactions,\n"
            "             teach_ui_control,\n"
            "             rebuild_ui_knowledge_index\n"
            "  vision: find_control_by_image, click_by_image, find_text_on_screen, click_text_on_screen,\n"
            "          find_control_by_vision, click_by_vision\n\n"
            "Full documentation: https://github.com/Poggi-Tang/easyautomation"
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"easy_uiauto {__version__}",
    )
    actions = parser.add_mutually_exclusive_group()
    actions.add_argument(
        "--quick-setup-codex",
        action="store_true",
        help="Configure remote vision and replace the global Codex MCP entry.",
    )
    actions.add_argument(
        "--full-setup-codex",
        action="store_true",
        help="Install vision requirements, configure Codex, and run UIA/OCR/AI checks.",
    )
    actions.add_argument(
        "--install-codex",
        action="store_true",
        help="Add easy_uiauto to the global Codex MCP configuration.",
    )
    actions.add_argument(
        "--install-codex-skills",
        action="store_true",
        help="Install or update the bundled learning and operation Codex skills.",
    )
    actions.add_argument(
        "--uninstall-codex-skills",
        action="store_true",
        help="Remove the bundled easy_uiauto skills from Codex.",
    )
    actions.add_argument(
        "--uninstall-codex",
        action="store_true",
        help="Remove easy_uiauto from the global Codex MCP configuration.",
    )
    actions.add_argument(
        "--show-codex-config",
        action="store_true",
        help="Show the easy_uiauto MCP entry configured in Codex.",
    )
    actions.add_argument(
        "--install-claude-code",
        action="store_true",
        help="Add easy_uiauto to the user-scoped Claude Code MCP configuration.",
    )
    actions.add_argument(
        "--uninstall-claude-code",
        action="store_true",
        help="Remove easy_uiauto from the user-scoped Claude Code MCP configuration.",
    )
    actions.add_argument(
        "--show-claude-code-config",
        action="store_true",
        help="Show the easy_uiauto MCP entry configured in Claude Code.",
    )
    parser.add_argument(
        "--vision-url",
        default="",
        metavar="URL",
        help="Remote vision API URL used by the quick and full Codex setup actions.",
    )
    parser.add_argument(
        "--vision-model",
        default="",
        metavar="MODEL",
        help="Remote vision model used by the quick and full Codex setup actions.",
    )
    return parser


def main(argv: Optional[list[str]] = None):
    """Run the MCP server."""
    parser = _build_arg_parser()
    args = parser.parse_args(argv)
    if args.full_setup_codex:
        try:
            output = configuration.full_setup_codex(
                api_url=args.vision_url,
                model=args.vision_model,
                version=__version__,
            )
        except (RuntimeError, ValueError) as error:
            parser.error(str(error))
        print(output)
        return
    if args.quick_setup_codex:
        try:
            output = configuration.quick_setup_codex(
                api_url=args.vision_url,
                model=args.vision_model,
                version=__version__,
            )
        except (RuntimeError, ValueError) as error:
            parser.error(str(error))
        print(output)
        return
    actions = {
        "install_codex": configuration.install_codex,
        "uninstall_codex": configuration.uninstall_codex,
        "show_codex_config": configuration.show_codex,
        "install_codex_skills": skill_installation.install_codex_skills,
        "uninstall_codex_skills": skill_installation.uninstall_codex_skills,
        "install_claude_code": configuration.install_claude_code,
        "uninstall_claude_code": configuration.uninstall_claude_code,
        "show_claude_code_config": configuration.show_claude_code,
    }
    for option, action in actions.items():
        if getattr(args, option):
            try:
                output = action()
            except RuntimeError as error:
                parser.error(str(error))
            if output:
                print(output)
            return
    mcp.run()


if __name__ == "__main__":
    main()
