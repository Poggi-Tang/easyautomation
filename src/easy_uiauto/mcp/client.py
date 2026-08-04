"""Python client for the UI automation TCP service.

Usage as context manager:
    with UIAutomationClient() as client:
        client.activate_window("SimuLab")
        client.click(name="File", control_type="MenuItemControl")

Usage as standalone:
    client = UIAutomationClient()
    client.connect()
    client.click(name="File", control_type="MenuItemControl")
    client.close()
"""
from __future__ import annotations

import base64
import json
import socket
import uuid
from typing import Any, Optional


class UIAutomationError(Exception):
    """Raised when the service returns an error."""
    pass


class UIAutomationClient:
    """Thin TCP client for the UI automation service."""

    def __init__(self, host: str = "127.0.0.1", port: int = 9876, timeout: float = 30.0):
        self.host = host
        self.port = port
        self.timeout = timeout
        self._sock: Optional[socket.socket] = None
        self._buf = b""

    # ---- lifecycle ---------------------------------------------------------

    def connect(self) -> None:
        """Connect to the service. Raises ConnectionRefusedError if not running."""
        self._sock = socket.create_connection((self.host, self.port), timeout=self.timeout)

    def close(self) -> None:
        """Close the connection."""
        if self._sock:
            try:
                self._sock.close()
            except Exception:
                pass
            self._sock = None

    def __enter__(self) -> "UIAutomationClient":
        self.connect()
        return self

    def __exit__(self, *args) -> None:
        self.close()

    # ---- internal ----------------------------------------------------------

    def _call(self, method: str, **params) -> Any:
        """Send a command and return the result. Raises UIAutomationError on failure."""
        if self._sock is None:
            raise UIAutomationError("Not connected. Call connect() first.")

        req_id = str(uuid.uuid4())[:8]
        cmd = json.dumps({"id": req_id, "method": method, "params": params}, ensure_ascii=False)
        self._sock.sendall(cmd.encode("utf-8") + b"\n")

        # Read response (line-delimited JSON)
        while b"\n" not in self._buf:
            chunk = self._sock.recv(65536)
            if not chunk:
                raise UIAutomationError("Connection closed by service")
            self._buf += chunk

        line, self._buf = self._buf.split(b"\n", 1)
        resp = json.loads(line.decode("utf-8"))

        if resp.get("error"):
            raise UIAutomationError(resp["error"])
        return resp.get("result")

    # ---- convenience methods -----------------------------------------------

    def list_windows(self) -> list[dict]:
        return self._call("list_windows")

    def activate_window(self, title: str) -> str:
        return self._call("activate_window", window_name=title)

    def get_control_tree(self, window_name: str, max_depth: int = 5) -> dict | str:
        return self._call("get_control_tree", window_name=window_name, max_depth=max_depth)

    def find_control(self, **kwargs) -> dict | str:
        return self._call("find_control", **kwargs)

    def get_control_at_position(self, x: int, y: int) -> dict | str:
        return self._call("get_control_at_position", x=x, y=y)

    def click(self, **kwargs) -> str:
        return self._call("click", **kwargs)

    def smart_find(self, window_name: str, name: str = "", control_type: str = "",
                   class_name: str = "", max_depth: int = 6) -> dict | str:
        """Find a control using control-tree walk (bypasses easy-uiauto find_control bug)."""
        return self._call("smart_find", window_name=window_name, name=name,
                         control_type=control_type, class_name=class_name, max_depth=max_depth)

    def smart_click(self, window_name: str = "", name: str = "", control_type: str = "",
                    class_name: str = "", button: str = "left", click_type: str = "single") -> str:
        """Click a control by walking the control tree + pyautogui coordinate click."""
        return self._call("smart_click", window_name=window_name, name=name,
                         control_type=control_type, class_name=class_name,
                         button=button, click_type=click_type)

    def run_action(self, action: dict) -> str:
        """Execute a recorded-style action dict directly via easy-uiauto run_action."""
        return self._call("run_action", action=action)

    def run_actions(self, actions: list[dict], stop_on_error: bool = True,
                    delay_ms: int = 0) -> dict:
        """Execute multiple commands or recorded actions in one service request."""
        return self._call(
            "run_actions",
            actions=actions,
            stop_on_error=stop_on_error,
            delay_ms=delay_ms,
        )

    def compile_window(self, window_name: str, max_depth: int = 8) -> str:
        """Pre-compile a window's control tree so find_control hits cache instantly."""
        return self._call("compile_window", window_name=window_name, max_depth=max_depth)

    def click_at_position(self, x: int, y: int, button: str = "left", click_type: str = "single") -> str:
        return self._call("click_at_position", x=x, y=y, button=button, click_type=click_type)

    def move_mouse(self, x: int, y: int) -> str:
        return self._call("move_mouse", x=x, y=y)

    def drag_control(self, **kwargs) -> str:
        return self._call("drag_control", **kwargs)

    def scroll(self, amount: int = 3, direction: str = "down", x: int = -1, y: int = -1) -> str:
        return self._call("scroll", amount=amount, direction=direction, x=x, y=y)

    def mouse_scroll(self, amount: int = 3, direction: str = "down") -> str:
        return self._call("mouse_scroll", amount=amount, direction=direction)

    def type_text(self, text: str, **kwargs) -> str:
        return self._call("type_text", text=text, **kwargs)

    def set_text(self, text: str, **kwargs) -> str:
        return self._call("set_text", text=text, **kwargs)

    def press_key(self, key: str, window_name: str = "") -> str:
        return self._call("press_key", key=key, window_name=window_name)

    def hotkey(self, keys: str, window_name: str = "") -> str:
        return self._call("hotkey", keys=keys, window_name=window_name)

    def screenshot(self, region_x: int = 0, region_y: int = 0,
                   region_width: int = 0, region_height: int = 0) -> bytes:
        """Take a screenshot, returning raw PNG bytes."""
        result = self._call("screenshot", region_x=region_x, region_y=region_y,
                            region_width=region_width, region_height=region_height)
        return base64.b64decode(result["data"])

    def shutdown(self) -> str:
        return self._call("shutdown")
