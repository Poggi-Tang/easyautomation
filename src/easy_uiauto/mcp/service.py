"""TCP UI automation adapter using shared strict actions and bounded tree reads.

Start once, then send JSON commands via localhost:9876.

Usage:
    python -m easy_uiauto.mcp.service [--port 9876] [--cache-depth 8]
"""
from __future__ import annotations

import argparse
import base64
import io
import json
import os
import selectors
import socket
import time
from typing import Any, Callable, Optional

# ---------------------------------------------------------------------------
# Configure pyautogui early
# ---------------------------------------------------------------------------
import pyautogui

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.05  # faster than MCP server's 0.1

# ---------------------------------------------------------------------------
# easy-uiauto imports — careful to avoid name shadowing
# ---------------------------------------------------------------------------
from easy_uiauto.ctrl import Controller, run_action as _run_action_easy
from easy_uiauto.utils import (  # noqa: E402
    auto_scroll,
    find_control as _find_control_utils,  # unambiguous name
    get_control_info,
    get_control_xpath,
)
import uiautomation

# ---------------------------------------------------------------------------
# Local protocol helpers
# ---------------------------------------------------------------------------
from . import control_tree
from .protocol import (
    build_location,
    location_from_xpath,
    normalize_location,
    pack_response,
    parse_params,
    parse_xpath,
)
from .. import __version__

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DEFAULT_PORT = 9876
DEFAULT_CACHE_DEPTH = 8
PID_FILE = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")), "easy_uiauto_service.pid")
LOG_FILE = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")), "easy_uiauto_service.log")
RECV_BUF = 65536


def _log(msg: str) -> None:
    """Write a timestamped message to the log file."""
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)  # also visible in the terminal that launched the service
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# The service
# ---------------------------------------------------------------------------


class UIAutomationService:
    """Long-running TCP service that accepts JSON commands and performs UI actions."""

    @staticmethod
    def _pointer_action(button, click_type):
        if button not in {"left", "right", "middle"} or click_type not in {"single", "double"}:
            raise ValueError("Unsupported pointer action")
        if click_type == "double":
            return {"left": "double-click", "right": "double-right-click",
                    "middle": "double-middle-click"}[button]
        return {"left": "click", "right": "right-click", "middle": "middle-click"}[button]

    def __init__(self, host: str = "127.0.0.1", port: int = DEFAULT_PORT, cache_depth: int = DEFAULT_CACHE_DEPTH):
        self.host = host
        self.port = port
        self.cache_depth = cache_depth
        self._running = False
        self._sel = selectors.DefaultSelector()
        self._server_sock: Optional[socket.socket] = None

        # Handler dispatch table
        self._handlers: dict[str, Callable[[dict, str], Any]] = {
            "activate_window": self._handle_activate_window,
            "compile_window": self._handle_compile_window,
            "run_action": self._handle_run_action,
            "run_actions": self._handle_run_actions,
            "smart_click": self._handle_smart_click,
            "smart_find": self._handle_smart_find,
            "click": self._handle_click,
            "click_at_position": self._handle_click_at_position,
            "drag_control": self._handle_drag_control,
            "find_control": self._handle_find_control,
            "get_control_at_position": self._handle_get_control_at_position,
            "get_control_tree": self._handle_get_control_tree,
            "hotkey": self._handle_hotkey,
            "list_windows": self._handle_list_windows,
            "mouse_scroll": self._handle_mouse_scroll,
            "move_mouse": self._handle_move_mouse,
            "press_key": self._handle_press_key,
            "screenshot": self._handle_screenshot,
            "scroll": self._handle_scroll,
            "set_text": self._handle_set_text,
            "shutdown": self._handle_shutdown,
            "type_text": self._handle_type_text,
        }

    # ---- lifecycle ---------------------------------------------------------

    def start(self) -> None:
        """Bind, warm cache, and enter the accept loop."""
        _log(f"UI Automation Service starting on {self.host}:{self.port}")
        _log(f"Explicit inspection depth: {self.cache_depth}, PID: {os.getpid()}")

        # Write PID file
        try:
            with open(PID_FILE, "w") as f:
                json.dump({"pid": os.getpid(), "port": self.port}, f)
        except Exception:
            pass

        # Set up server socket
        self._server_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._server_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._server_sock.bind((self.host, self.port))
        self._server_sock.listen(5)
        self._server_sock.setblocking(False)

        self._sel.register(self._server_sock, selectors.EVENT_READ, data=None)
        self._running = True

        _log(f"Service ready — listening on {self.host}:{self.port}")

        try:
            while self._running:
                events = self._sel.select(timeout=1.0)
                for key, mask in events:
                    if key.data is None:
                        # New connection
                        self._accept_connection(key.fileobj)
                    else:
                        # Client data
                        self._handle_client(key.fileobj, key.data)
        except KeyboardInterrupt:
            _log("Received interrupt, shutting down")
        finally:
            self._cleanup()

    def _cleanup(self) -> None:
        """Release resources."""
        self._running = False
        if self._server_sock:
            self._sel.unregister(self._server_sock)
            self._server_sock.close()
        self._sel.close()
        try:
            os.remove(PID_FILE)
        except Exception:
            pass
        _log("Service stopped")


    # ---- networking --------------------------------------------------------

    def _accept_connection(self, sock: socket.socket) -> None:
        """Accept a new client connection."""
        conn, addr = sock.accept()
        conn.setblocking(False)
        buf = bytearray()
        self._sel.register(conn, selectors.EVENT_READ, data=buf)
        _log(f"Client connected: {addr}")

    def _handle_client(self, conn: socket.socket, buf: bytearray) -> None:
        """Read available data, extract complete JSON lines, dispatch."""
        try:
            data = conn.recv(RECV_BUF)
        except (ConnectionResetError, ConnectionAbortedError):
            self._close_connection(conn)
            return

        if not data:
            self._close_connection(conn)
            return

        buf.extend(data)

        # Process complete lines (newline-delimited JSON)
        while True:
            idx = buf.find(b"\n")
            if idx == -1:
                break
            line = bytes(buf[:idx]).decode("utf-8", errors="replace").strip()
            del buf[: idx + 1]

            if not line:
                continue

            try:
                cmd = json.loads(line)
            except json.JSONDecodeError:
                self._send_error(conn, "", f"Invalid JSON: {line[:100]}")
                continue

            req_id = cmd.get("id", "")
            method = cmd.get("method", "")
            params = cmd.get("params", {})

            t0 = time.perf_counter()
            try:
                handler = self._handlers.get(method)
                if handler is None:
                    resp = pack_response(req_id, error=f"Unknown method: {method}")
                else:
                    result = handler(params, req_id)
                    elapsed = time.perf_counter() - t0
                    resp = pack_response(req_id, result=result, elapsed=elapsed)
            except Exception as e:
                elapsed = time.perf_counter() - t0
                resp = pack_response(req_id, error=f"{type(e).__name__}: {e}", elapsed=elapsed)

            try:
                payload = json.dumps(resp, ensure_ascii=False) + "\n"
                conn.sendall(payload.encode("utf-8"))
            except (BrokenPipeError, ConnectionResetError):
                self._close_connection(conn)
                return

    def _close_connection(self, conn: socket.socket) -> None:
        """Unregister and close a client connection."""
        try:
            self._sel.unregister(conn)
        except Exception:
            pass
        try:
            conn.close()
        except Exception:
            pass
        _log("Client disconnected")

    def _send_error(self, conn: socket.socket, req_id: str, msg: str) -> None:
        """Send an error response immediately."""
        resp = pack_response(req_id, error=msg)
        try:
            conn.sendall(json.dumps(resp, ensure_ascii=False).encode("utf-8") + b"\n")
        except Exception:
            self._close_connection(conn)

    # ---- window activation caching ----------------------------------------

    def _ensure_window(self, window_name: str) -> None:
        """Resolve a visible unique window before explicitly activating its handle."""
        if window_name:
            control = _find_control_utils(build_location(window_name=window_name))
            Controller._activate_target(control)


    # ---- handler: list_windows --------------------------------------------

    def _handle_list_windows(self, params: dict, req_id: str) -> list[dict]:
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
        return result

    # ---- handler: compile_window ------------------------------------------

    def _handle_compile_window(self, params: dict, req_id: str) -> dict | str:
        """Compatibility entry: perform a bounded read, without persistent UIA caches."""
        window_name = params.get("window_name", "")
        depth = params.get("max_depth", self.cache_depth)
        result = self._handle_get_control_tree({**params, "max_depth": depth}, req_id)
        if result.get("ok") is False:
            return result
        return f"Compiled '{window_name}' (depth={depth})"

    # ---- handler: run_action (recorded/replay format) ---------------------

    def _handle_run_action(self, params: dict, req_id: str) -> str:
        """Execute a recorded-style action dict directly via easy-uiauto's run_action.

        params: {"action": {"TEST_ID": "1", "ACTION": "点击", "LOCATION": {...}}}
        """
        action = params.get("action", params)  # support both wrapped and direct
        if params.get("match_mode") is not None:
            from .protocol import with_match_mode
            action = {**action, "LOCATION": with_match_mode(action["LOCATION"], params["match_mode"])}
        result = _run_action_easy(action)
        return str(result)

    def _handle_run_actions(self, params: dict, req_id: str) -> dict:
        """Execute multiple service commands or recorded actions in one request.

        params:
          {
            "actions": [
              {"method": "click", "params": {...}},
              {"action": "type_text", "text": "hello"},
              {"ACTION": "点击", "LOCATION": {...}}
            ],
            "stop_on_error": true,
            "delay_ms": 0
          }
        """
        actions = params.get("actions", [])
        if isinstance(actions, str):
            actions = json.loads(actions)
        if not isinstance(actions, list):
            raise ValueError("'actions' must be a list")

        stop_on_error = bool(params.get("stop_on_error", True))
        delay_ms = int(params.get("delay_ms", 0) or 0)
        results = []

        aliases = {
            "type": "type_text",
            "input_text": "type_text",
            "send_keys": "hotkey",
            "key": "press_key",
            "screenshot": "screenshot",
        }
        unsupported = {"run_actions", "shutdown"}

        for index, step in enumerate(actions, start=1):
            t0 = time.perf_counter()
            try:
                if not isinstance(step, dict):
                    raise ValueError("Each action must be an object")

                if "ACTION" in step and "LOCATION" in step:
                    method = "run_action"
                    step_params = {"action": step}
                else:
                    method = str(step.get("method") or step.get("action") or step.get("tool") or "")
                    if not method:
                        raise ValueError("Action is missing 'method' or 'action'")
                    method = aliases.get(method, method)
                    if method in unsupported:
                        raise ValueError(f"Unsupported batch method: {method}")
                    step_params = step.get("params")
                    if step_params is None:
                        step_params = {
                            k: v for k, v in step.items()
                            if k not in {"method", "action", "tool"}
                        }
                    if not isinstance(step_params, dict):
                        raise ValueError("Action 'params' must be an object")

                handler = self._handlers.get(method)
                if handler is None:
                    raise ValueError(f"Unknown method: {method}")

                result = handler(step_params, f"{req_id}:{index}")
                elapsed = time.perf_counter() - t0
                ok = not (
                    isinstance(result, str)
                    and (
                        result.lower().startswith("error")
                        or "异常" in result
                        or "未找到" in result
                    )
                )
                results.append({
                    "index": index,
                    "method": method,
                    "ok": ok,
                    "result": result,
                    "timing_ms": round(elapsed * 1000, 1),
                })

                if isinstance(result, dict) and result.get("ok") is False:
                    ok = False
                    results[-1]["ok"] = False

                if delay_ms > 0 and index < len(actions):
                    time.sleep(delay_ms / 1000)
                if stop_on_error and not ok:
                    break
            except Exception as e:
                elapsed = time.perf_counter() - t0
                results.append({
                    "index": index,
                    "method": step.get("method") or step.get("action") or step.get("ACTION", "") if isinstance(step, dict) else "",
                    "ok": False,
                    "error": f"{type(e).__name__}: {e}",
                    "timing_ms": round(elapsed * 1000, 1),
                })
                if stop_on_error:
                    break

        return {
            "ok": all(item.get("ok", False) for item in results) and len(results) == len(actions),
            "total": len(actions),
            "executed": len(results),
            "results": results,
        }

    # ---- handler: activate_window -----------------------------------------

    def _handle_activate_window(self, params: dict, req_id: str) -> str:
        [title] = parse_params(params, "window_name")
        self._ensure_window(title)
        return f"Activated: {title}"

    # ---- handler: click ---------------------------------------------------

    def _handle_click(self, params: dict, req_id: str) -> str:
        [window_name, name, class_name, control_type, automation_id,
         found_index, xpath, button, click_type, x_offset, y_offset] = parse_params(
            params, "window_name", "name", "class_name", "control_type",
            "automation_id", "found_index", "xpath", "button", "click_type",
            "x_offset", "y_offset",
        )

        loc = build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath,
            parameters=json.dumps({"x": x_offset, "y": y_offset}),
            match_mode=params.get("match_mode"),
        )
        action_title = f"Click {name or control_type or 'control'}"

        if click_type == "double":
            return Controller.double_click(ActionTitle=action_title, **loc)
        elif button == "right":
            return Controller.right_click(ActionTitle=action_title, **loc)
        elif button == "middle":
            return Controller.centre_click(ActionTitle=action_title, **loc)
        else:
            return Controller.left_click(ActionTitle=action_title, **loc)

    # ---- handler: smart_click / smart_find ---------------------------------

    def _smart_find_control(self, window_name: str, name: str = "",
                            control_type: str = "", class_name: str = "",
                            max_depth: int = 6, *, found_index: int = 0,
                            match_mode: str | None = None, xpath: str = "") -> dict:
        """Read a fresh control through the shared mode rules; never cache coordinates."""
        from easy_uiauto.utils import observe_control

        location = build_location(window_name=window_name, name=name,
                                  control_type=control_type, class_name=class_name,
                                  found_index=found_index, match_mode=match_mode, xpath=xpath)
        control = observe_control(location, search_depth=max_depth)
        # A mixed positional path resolved exactly; retain its parent-selection intent.
        path = location["Xpath"]
        if match_mode is not None or not any(step.get("match_mode") == "position" for step in path):
            path = get_control_xpath(control, match_mode=match_mode or "identity")
        rect = control.BoundingRectangle
        return {"Name": control.Name or "", "ControlType": control.ControlTypeName,
                "ClassName": control.ClassName or "", "AutomationId": control.AutomationId or "",
                "bounds": {"left": rect.left, "top": rect.top, "right": rect.right, "bottom": rect.bottom},
                "center": {"x": (rect.left + rect.right) // 2, "y": (rect.top + rect.bottom) // 2},
                "xpath": path}

    def _handle_smart_find(self, params: dict, req_id: str) -> dict | str:
        [window_name, name, control_type, class_name, max_depth] = parse_params(
            params, "window_name", "name", "control_type", "class_name", "max_depth",
        )
        return self._smart_find_control(
            window_name, name or "", control_type or "", class_name or "",
            int(max_depth) if max_depth else 6,
            found_index=params.get("found_index", 0), match_mode=params.get("match_mode"),
            xpath=params.get("xpath", ""),
        )

    def _handle_smart_click(self, params: dict, req_id: str) -> str:
        [window_name, name, control_type, class_name, button, click_type] = parse_params(
            params, "window_name", "name", "control_type", "class_name", "button", "click_type")
        action = self._pointer_action(button, click_type)
        location = build_location(window_name=window_name, name=name, class_name=class_name,
                                  control_type=control_type, automation_id=params.get("automation_id", ""),
                                  found_index=params.get("found_index", 0), xpath=params.get("xpath", ""),
                                  match_mode=params.get("match_mode"))
        control = _find_control_utils(location)
        Controller._activate_target(control)
        rect = control.BoundingRectangle
        cx, cy = (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2
        Controller.execute_resolved(action, control=control)
        return f"步骤：SmartClick {name or 'control'} 控件【{control.Name}】 点击 ({cx},{cy}) 成功！"

    # ---- handler: click_at_position ---------------------------------------

    def _handle_click_at_position(self, params: dict, req_id: str) -> str:
        [x, y, button, click_type] = parse_params(params, "x", "y", "button", "click_type")
        Controller.execute_resolved(self._pointer_action(button, click_type), x=x, y=y)
        return f"Clicked ({button}/{click_type}) at ({x}, {y})"

    # ---- handler: type_text -----------------------------------------------

    def _handle_type_text(self, params: dict, req_id: str) -> str:
        [window_name, text, name, class_name, control_type, automation_id,
         found_index, xpath, method] = parse_params(
            params, "window_name", "text", "name", "class_name", "control_type",
            "automation_id", "found_index", "xpath", "method",
        )
        loc = build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath,
            parameters=json.dumps({"输入文本": text}),
            match_mode=params.get("match_mode"),
        )
        if method == "sendkeys":
            loc["PARAMETERS"] = {**loc["PARAMETERS"], "设置文本": text}
            return Controller.set_text(ActionTitle=f"Set text on {name or control_type}", **loc)
        else:
            return Controller.input_text(ActionTitle=f"Type into {name or control_type}", **loc)

    # ---- handler: set_text ------------------------------------------------

    def _handle_set_text(self, params: dict, req_id: str) -> str:
        [window_name, text, name, class_name, control_type, automation_id,
         found_index, xpath] = parse_params(
            params, "window_name", "text", "name", "class_name", "control_type",
            "automation_id", "found_index", "xpath",
        )
        loc = build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath,
            parameters=json.dumps({"设置文本": text}),
            match_mode=params.get("match_mode"),
        )
        return Controller.set_text(ActionTitle=f"Set text on {name or control_type}", **loc)

    # ---- handler: press_key -----------------------------------------------

    def _handle_press_key(self, params: dict, req_id: str) -> str:
        [key, window_name] = parse_params(params, "key", "window_name")
        return Controller.key_click(
            ActionTitle=f"Press {key}",
            WindowName=window_name or "",
            **{"Name": "", "ClassName": "", "ControlType": "", "foundIndex": "",
               "AutomationId": "", "Xpath": [], "Img": "",
               "PARAMETERS": json.dumps({"键盘按键": key})},
        )

    # ---- handler: hotkey --------------------------------------------------

    def _handle_hotkey(self, params: dict, req_id: str) -> str:
        [keys, window_name] = parse_params(params, "keys", "window_name")
        return Controller.key_group(
            ActionTitle=f"Hotkey {keys}",
            WindowName=window_name or "",
            **{"Name": "", "ClassName": "", "ControlType": "", "foundIndex": "",
               "AutomationId": "", "Xpath": [], "Img": "",
               "PARAMETERS": json.dumps({"组合键": keys})},
        )

    # ---- handler: screenshot ----------------------------------------------

    def _handle_screenshot(self, params: dict, req_id: str) -> dict:
        [rx, ry, rw, rh] = parse_params(
            params, "region_x", "region_y", "region_width", "region_height",
        )
        if rw > 0 and rh > 0:
            img = pyautogui.screenshot(region=(rx, ry, rw, rh))
        else:
            img = pyautogui.screenshot()

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return {
            "format": "png",
            "width": img.width,
            "height": img.height,
            "data": base64.b64encode(buf.getvalue()).decode("ascii"),
        }

    # ---- handler: get_control_tree -----------------------------------------

    def _handle_get_control_tree(self, params: dict, req_id: str) -> dict:
        return control_tree.inspect_window(
            params.get("window_name", ""), params.get("max_depth", 5),
            max_nodes=params.get("max_nodes", 500),
            timeout_seconds=params.get("timeout_seconds", 15.0))

    # ---- handler: find_control --------------------------------------------

    def _handle_find_control(self, params: dict, req_id: str) -> dict | str:
        [window_name, name, class_name, control_type, automation_id,
         found_index, xpath] = parse_params(
            params, "window_name", "name", "class_name", "control_type",
            "automation_id", "found_index", "xpath",
        )
        if params.get("location") is not None:
            location = normalize_location(params["location"])
            if params.get("match_mode") is not None:
                from .protocol import with_match_mode
                location = with_match_mode(location, params["match_mode"])
        else:
            location = build_location(
                window_name=window_name, name=name, class_name=class_name,
                control_type=control_type, automation_id=automation_id,
                found_index=found_index, xpath=xpath, match_mode=params.get("match_mode"),
            )
        ctrl = _find_control_utils(location, debug=False)
        if not ctrl or ctrl is False:
            return f"Error: Control not found"
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
        return {
            "Name": ctrl.Name or "",
            "ClassName": ctrl.ClassName or "",
            "ControlType": ctrl.ControlTypeName if hasattr(ctrl, "ControlTypeName") else "",
            "AutomationId": ctrl.AutomationId or "",
            "bounds": bounds,
            "IsEnabled": getattr(ctrl, "IsEnabled", True),
            "IsVisible": getattr(ctrl, "IsVisible", True),
            "LOCATION": location,
        }

    # ---- handler: get_control_at_position ---------------------------------

    def _handle_get_control_at_position(self, params: dict, req_id: str) -> dict | str:
        [x, y] = parse_params(params, "x", "y")
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
        return {
            "Name": ctrl.Name or "",
            "ClassName": ctrl.ClassName or "",
            "ControlType": ctrl.ControlTypeName if hasattr(ctrl, "ControlTypeName") else "",
            "AutomationId": ctrl.AutomationId or "",
            "bounds": bounds,
            "xpath": xpath_list,
            "LOCATION": location_from_xpath(xpath_list),
        }

    # ---- handler: drag_control --------------------------------------------

    def _handle_drag_control(self, params: dict, req_id: str) -> str:
        [source_window, source_name, source_class, source_type, source_automation_id,
         source_xpath, target_window, target_name, target_class, target_type,
         target_automation_id, target_xpath] = parse_params(
            params, "source_window", "source_name", "source_class", "source_type",
            "source_automation_id", "source_xpath", "target_window", "target_name",
            "target_class", "target_type", "target_automation_id", "target_xpath",
        )
        target_params = json.dumps({
            "目的控件父窗口名称": target_window or source_window,
            "目的控件Name": target_name,
            "目的控件ClassName": target_class,
            "目的控件ControlType": target_type,
            "目的控件foundIndex": "",
            "目的控件AutomationId": target_automation_id,
            "目的控件Xpath": parse_xpath(target_xpath),
        })
        return Controller.drag_control(
            ActionTitle=f"Drag {source_name} to {target_name}",
            WindowName=source_window,
            Name=source_name,
            ClassName=source_class,
            ControlType=source_type,
            foundIndex="",
            AutomationId=source_automation_id,
            Xpath=parse_xpath(source_xpath),
            Img="",
            PARAMETERS=target_params,
        )

    # ---- handler: scroll / mouse_scroll -----------------------------------

    def _handle_scroll(self, params: dict, req_id: str) -> str:
        [amount, direction, x, y] = parse_params(params, "amount", "direction", "x", "y")
        if x >= 0 and y >= 0:
            pyautogui.moveTo(x, y)
        auto_scroll(amount, direction=direction)
        return f"Scrolled {direction} {amount} ticks"

    def _handle_mouse_scroll(self, params: dict, req_id: str) -> str:
        [amount, direction] = parse_params(params, "amount", "direction")
        if direction == "up":
            pyautogui.scroll(amount)
        else:
            pyautogui.scroll(-amount)
        return f"Scrolled {direction} {amount} ticks"

    # ---- handler: move_mouse ----------------------------------------------

    def _handle_move_mouse(self, params: dict, req_id: str) -> str:
        [x, y] = parse_params(params, "x", "y")
        pyautogui.moveTo(x, y)
        return f"Mouse moved to ({x}, {y})"

    # ---- handler: shutdown ------------------------------------------------

    def _handle_shutdown(self, params: dict, req_id: str) -> str:
        _log("Shutdown requested by client")
        self._running = False
        return "Service shutting down"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        prog="easy_uiauto_service",
        description="easy_uiauto TCP service for long-running Windows UI automation.",
    )
    parser.add_argument("--version", action="version", version=f"easy_uiauto {__version__}")
    parser.add_argument("--host", default="127.0.0.1", help="Bind address (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help=f"TCP port (default: {DEFAULT_PORT})")
    parser.add_argument("--cache-depth", type=int, default=DEFAULT_CACHE_DEPTH,
                        help=f"Default depth for explicit compile_window reads; no startup warming (default: {DEFAULT_CACHE_DEPTH})")
    args = parser.parse_args()

    svc = UIAutomationService(host=args.host, port=args.port, cache_depth=args.cache_depth)
    svc.start()


if __name__ == "__main__":
    main()
