"""TCP-based UI automation service — long-running, fast, with control-cache pre-warming.

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
import sys
import time
import traceback
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
    compile_controls,
    correct_ctrl_position,
    find_control as _find_control_utils,  # unambiguous name
    get_control_info,
    get_control_xpath,
    package_location,
    push_message,
    set_top_window,
)
import uiautomation

# ---------------------------------------------------------------------------
# Local protocol helpers
# ---------------------------------------------------------------------------
from .protocol import build_location, pack_response, parse_params
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

    def __init__(self, host: str = "127.0.0.1", port: int = DEFAULT_PORT, cache_depth: int = DEFAULT_CACHE_DEPTH):
        self.host = host
        self.port = port
        self.cache_depth = cache_depth
        self.current_window: str = ""  # cached for _ensure_window
        self._running = False
        self._sel = selectors.DefaultSelector()
        self._server_sock: Optional[socket.socket] = None
        self._tree_cache: dict = {}  # cached control info by (window_name, name, control_type)
        self._cached_window: str = ""  # which window the tree cache is for
        self._compiled_windows: set = set()  # track which windows have been compiled

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
        _log(f"Cache depth: {self.cache_depth}, PID: {os.getpid()}")

        # Write PID file
        try:
            with open(PID_FILE, "w") as f:
                json.dump({"pid": os.getpid(), "port": self.port}, f)
        except Exception:
            pass

        # Warm the control cache
        self._warm_cache()

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

    def _warm_cache(self) -> None:
        """Pre-warm CONTROL_CACHE by walking the desktop control tree."""
        if self.cache_depth <= 0:
            _log("Cache warming disabled (depth <= 0)")
            return
        _log(f"Warming control cache (max_depth={self.cache_depth})...")
        t0 = time.perf_counter()
        try:
            compile_controls(control=None, max_depth=self.cache_depth, compile_log=False)
            elapsed = time.perf_counter() - t0
            _log(f"Cache warmed in {elapsed:.1f}s")
        except Exception as e:
            _log(f"Cache warming failed (non-fatal): {e}")

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
        """Activate window only if it differs from the cached current window."""
        if window_name and window_name != self.current_window:
            set_top_window(window_name)
            self.current_window = window_name
            # Invalidate tree cache on window switch
            self._tree_cache.clear()
            self._cached_window = ""

    def _ensure_compiled(self, window_name: str, max_depth: int = 8) -> None:
        """Compile a window's control tree so find_control hits the cache.

        This is THE key optimization: after compilation, CONTROL_CACHE has
        proper xpath entries (with searchDepth/ClassName), so find_control
        works instantly instead of timing out.
        Does NOT require window activation — works on background windows too.
        """
        if not window_name or window_name in self._compiled_windows:
            return
        try:
            import uiautomation as _ua
            _ua.SetGlobalSearchTimeout(3)
            win = _ua.WindowControl(Name=window_name, searchDepth=1)
            if not win.Exists(maxSearchSeconds=2):
                win = _ua.WindowControl(
                    searchFromControl=_ua.GetRootControl(),
                    Name=window_name, searchDepth=5,
                )
                if not win.Exists(maxSearchSeconds=2):
                    _log(f"Window '{window_name}' not found for compilation")
                    return
            _log(f"Compiling controls for '{window_name}' (depth={max_depth})...")
            t0 = time.perf_counter()
            compile_controls(control=win, max_depth=max_depth, compile_log=False)
            elapsed = time.perf_counter() - t0
            self._compiled_windows.add(window_name)
            _log(f"Compiled '{window_name}' in {elapsed:.1f}s")
        except Exception as e:
            _log(f"Compile failed for '{window_name}': {e}")

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

    def _handle_compile_window(self, params: dict, req_id: str) -> str:
        [window_name, max_depth] = parse_params(params, "window_name", "max_depth")
        depth = int(max_depth) if max_depth else 8
        # Force re-compile by clearing the flag
        self._compiled_windows.discard(window_name)
        self._ensure_compiled(window_name, depth)
        return f"Compiled '{window_name}' (depth={depth})"

    # ---- handler: run_action (recorded/replay format) ---------------------

    def _handle_run_action(self, params: dict, req_id: str) -> str:
        """Execute a recorded-style action dict directly via easy-uiauto's run_action.

        params: {"action": {"TEST_ID": "1", "ACTION": "点击", "LOCATION": {...}}}
        """
        action = params.get("action", params)  # support both wrapped and direct
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

        self._ensure_window(window_name)
        self._ensure_compiled(window_name)  # compile first so find_control hits cache

        # If no xpath provided, auto-generate from compiled cache via tree walk
        if not xpath and (name or control_type):
            ctrl_info = self._smart_find_control(window_name, name, control_type, class_name)
            if ctrl_info:
                xpath = json.dumps(ctrl_info.get("xpath", []), ensure_ascii=False)
                _log(f"Auto-generated xpath for '{name}': {len(ctrl_info.get('xpath',[]))} levels")

        loc = build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath,
            parameters=json.dumps({"x": x_offset, "y": y_offset}),
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
                            max_depth: int = 6) -> dict | None:
        """Use get_control_tree to find a control and return its bounds + center.

        This bypasses easy-uiauto's find_control (which has Qt menu bugs) by
        using the reliable uiautomation.WindowControl directly.
        Results are cached per-window for fast repeated lookups.
        """
        # Check cache first
        cache_key = f"{window_name}|{name}|{control_type}|{class_name}"
        if self._cached_window == window_name and cache_key in self._tree_cache:
            return self._tree_cache[cache_key]

        try:
            import uiautomation as _ua_inner
            # Only set_top_window if window not already active
            if window_name != self._cached_window:
                set_top_window(window_name)
            win = _ua_inner.WindowControl(Name=window_name, searchDepth=1)
            if not win.Exists(maxSearchSeconds=2):
                win = _ua_inner.WindowControl(
                    searchFromControl=_ua_inner.GetRootControl(),
                    Name=window_name, searchDepth=5,
                )
                if not win.Exists(maxSearchSeconds=2):
                    return None

            def _search(ctrl, path=None, depth=0) -> dict | None:
                if depth > max_depth:
                    return None
                if path is None:
                    path = []
                try:
                    cname = ctrl.Name or ""
                    ctype = ctrl.ControlTypeName if hasattr(ctrl, "ControlTypeName") else ""
                    cclass = ctrl.ClassName or ""
                    caid = ctrl.AutomationId or ""

                    # Build xpath node for this control
                    node = {"ControlType": ctype}
                    if cname:
                        node["Name"] = cname
                    if cclass:
                        node["ClassName"] = cclass
                    if caid:
                        node["AutomationId"] = caid
                    node["searchDepth"] = depth + 1
                    current_path = path + [node]

                    match = True
                    if name and name != cname:
                        match = False
                    if control_type and control_type != ctype:
                        match = False
                    if class_name and class_name != cclass:
                        match = False
                    if match and (name or control_type):
                        rect = ctrl.BoundingRectangle
                        if rect:
                            cx = (rect.left + rect.right) // 2
                            cy = (rect.top + rect.bottom) // 2
                            return {
                                "Name": cname,
                                "ControlType": ctype,
                                "ClassName": cclass,
                                "AutomationId": caid,
                                "bounds": {
                                    "left": rect.left, "top": rect.top,
                                    "right": rect.right, "bottom": rect.bottom,
                                },
                                "center": {"x": cx, "y": cy},
                                "xpath": current_path,
                            }
                    for child in ctrl.GetChildren():
                        found = _search(child, current_path, depth + 1)
                        if found:
                            return found
                except Exception:
                    pass
                return None

            result = _search(win)
            if result:
                self._tree_cache[cache_key] = result
                self._cached_window = window_name
            return result
        except Exception:
            return None

    def _handle_smart_find(self, params: dict, req_id: str) -> dict | str:
        [window_name, name, control_type, class_name, max_depth] = parse_params(
            params, "window_name", "name", "control_type", "class_name", "max_depth",
        )
        result = self._smart_find_control(
            window_name, name or "", control_type or "", class_name or "",
            int(max_depth) if max_depth else 6,
        )
        if result:
            return result
        return f"Error: Control not found (name='{name}', type='{control_type}')"

    def _handle_smart_click(self, params: dict, req_id: str) -> str:
        [window_name, name, control_type, class_name, button, click_type] = parse_params(
            params, "window_name", "name", "control_type", "class_name", "button", "click_type",
        )
        self._ensure_window(window_name)
        self._ensure_compiled(window_name)  # compile first for cache

        ctrl_info = self._smart_find_control(window_name, name or "", control_type or "", class_name or "")
        if not ctrl_info:
            return f"步骤：SmartClick {name or control_type} 异常：未找到控件"

        cx = ctrl_info["center"]["x"]
        cy = ctrl_info["center"]["y"]

        if click_type == "double":
            pyautogui.doubleClick(cx, cy, button=button)
        else:
            pyautogui.click(cx, cy, button=button)

        return f"步骤：SmartClick {name or 'control'} 控件【{ctrl_info['Name']}】 点击 ({cx},{cy}) 成功！"

    # ---- handler: click_at_position ---------------------------------------

    def _handle_click_at_position(self, params: dict, req_id: str) -> str:
        [x, y, button, click_type] = parse_params(
            params, "x", "y", "button", "click_type",
        )
        if click_type == "double":
            pyautogui.doubleClick(x, y, button=button)
        else:
            pyautogui.click(x, y, button=button)
        return f"Clicked ({button}/{click_type}) at ({x}, {y})"

    # ---- handler: type_text -----------------------------------------------

    def _handle_type_text(self, params: dict, req_id: str) -> str:
        [window_name, text, name, class_name, control_type, automation_id,
         found_index, xpath, method] = parse_params(
            params, "window_name", "text", "name", "class_name", "control_type",
            "automation_id", "found_index", "xpath", "method",
        )
        self._ensure_window(window_name)
        self._ensure_compiled(window_name) if window_name else None
        loc = build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath,
            parameters=json.dumps({"输入文本": text}),
        )
        if method == "sendkeys":
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
        self._ensure_window(window_name)
        self._ensure_compiled(window_name) if window_name else None
        loc = build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath,
            parameters=json.dumps({"设置文本": text}),
        )
        return Controller.set_text(ActionTitle=f"Set text on {name or control_type}", **loc)

    # ---- handler: press_key -----------------------------------------------

    def _handle_press_key(self, params: dict, req_id: str) -> str:
        [key, window_name] = parse_params(params, "key", "window_name")
        self._ensure_window(window_name)
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
        self._ensure_window(window_name)
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

    def _handle_get_control_tree(self, params: dict, req_id: str) -> dict | str:
        [window_name, max_depth] = parse_params(params, "window_name", "max_depth")
        max_depth = min(int(max_depth), 10)

        try:
            set_top_window(window_name)
            window = uiautomation.WindowControl(Name=window_name, searchDepth=1)
            if not window.Exists(maxSearchSeconds=3):
                window = uiautomation.WindowControl(
                    searchFromControl=uiautomation.GetRootControl(),
                    Name=window_name, searchDepth=5,
                )
                if not window.Exists(maxSearchSeconds=3):
                    return f"Error: Window '{window_name}' not found"

            def _ctrl_to_dict(ctrl, depth=0):
                if depth > max_depth:
                    return None
                try:
                    rect = ctrl.BoundingRectangle
                    info = {
                        "ControlType": ctrl.ControlTypeName if hasattr(ctrl, "ControlTypeName") else "",
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
                    children = []
                    try:
                        for child in ctrl.GetChildren():
                            ci = _ctrl_to_dict(child, depth + 1)
                            if ci:
                                children.append(ci)
                            if len(children) >= 30:
                                children.append({"_note": "...truncated"})
                                break
                    except Exception:
                        pass
                    if children:
                        info["children"] = children
                    return info
                except Exception:
                    return None

            tree = _ctrl_to_dict(window)
            return tree if tree else f"Error: Could not read control tree for '{window_name}'"
        except Exception as e:
            return f"Error: {e}"

    # ---- handler: find_control --------------------------------------------

    def _handle_find_control(self, params: dict, req_id: str) -> dict | str:
        [window_name, name, class_name, control_type, automation_id,
         found_index, xpath] = parse_params(
            params, "window_name", "name", "class_name", "control_type",
            "automation_id", "found_index", "xpath",
        )
        location = build_location(
            window_name=window_name, name=name, class_name=class_name,
            control_type=control_type, automation_id=automation_id,
            found_index=found_index, xpath=xpath,
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
            "目的控件Xpath": json.loads(target_xpath) if target_xpath else "",
        })
        return Controller.drag_control(
            ActionTitle=f"Drag {source_name} to {target_name}",
            WindowName=source_window,
            Name=source_name,
            ClassName=source_class,
            ControlType=source_type,
            foundIndex="",
            AutomationId=source_automation_id,
            Xpath=json.loads(source_xpath) if source_xpath else [],
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
                        help=f"Control tree cache depth, 0 to disable (default: {DEFAULT_CACHE_DEPTH})")
    args = parser.parse_args()

    svc = UIAutomationService(host=args.host, port=args.port, cache_depth=args.cache_depth)
    svc.start()


if __name__ == "__main__":
    main()
