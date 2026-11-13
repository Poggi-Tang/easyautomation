"""Live UI Automation region picking without injected clicks or wheel scrolling."""
from __future__ import annotations

import ctypes
from ctypes import wintypes
from pathlib import Path
from queue import Empty, Queue
import threading
import time


class ControlPath:
    """A point's leaf-to-window path; scrolling never chooses an arbitrary sibling."""
    def __init__(self, root, same):
        self.root, self.same = root, same
        self.nodes, self.level = [], 0

    def reset(self, leaf):
        nodes, node = [], leaf
        for _ in range(64):
            if node is None: break
            nodes.append(node)
            if self.same(node, self.root):
                self.nodes, self.level = nodes, 0
                return
            node = node.GetParentControl()
        self.nodes, self.level = [], 0
        raise ValueError("鼠标位置不属于所选软件窗口")

    @property
    def current(self):
        return self.nodes[self.level] if self.nodes else None

    def scroll(self, delta):
        if self.nodes:
            self.level = max(0, min(len(self.nodes)-1, self.level + (1 if delta > 0 else -1)))
        return self.current


def selection_metadata(control):
    from .mcp.scanner import _rect

    rect = _rect(control)
    if rect["width"] < 2 or rect["height"] < 2 or bool(getattr(control, "IsOffscreen", False)):
        raise ValueError("当前控件没有可见区域，请移动鼠标或上滚选择父节点")
    return {"automation_id": str(control.AutomationId or ""), "native_name": str(control.Name or ""),
            "class_name": str(control.ClassName or ""), "control_type": str(control.ControlTypeName),
            "rect": rect}


def normalized_selection(rect, bounds):
    """Clip resize margins, while rejecting controls from outside this window."""
    left, top = max(rect["left"], bounds["left"]), max(rect["top"], bounds["top"])
    right, bottom = min(rect["right"], bounds["right"]), min(rect["bottom"], bounds["bottom"])
    if right-left < 2 or bottom-top < 2 or bounds["width"] < 2 or bounds["height"] < 2:
        raise ValueError("控件不在所选窗口的可见范围内")
    return dict(left=(left-bounds["left"])/bounds["width"], top=(top-bounds["top"])/bounds["height"],
                right=(right-bounds["left"])/bounds["width"], bottom=(bottom-bounds["top"])/bounds["height"])


class PickerKeys:
    """Ignore held modifiers and key-repeat; consume only Ctrl and Escape."""
    def __init__(self, held=()):
        self.held = set(held)
        if self.held & {0xA2, 0xA3}: self.held.discard(0x11)
        self.initial = set(self.held)

    def event(self, key, down):
        if key not in {0x11, 0xA2, 0xA3, 0x1B}: return False, None
        if key in {0xA2, 0xA3} and 0x11 in self.initial: key = 0x11
        if not down:
            initially_down = key in self.initial
            self.held.discard(key); self.initial.discard(key)
            # A down event sent before our hook must receive its corresponding up.
            return not initially_down, None
        if key in self.held: return True, None
        self.held.add(key)
        if key == 0x1B: return True, "cancel"
        if self.initial: return True, None
        return True, "confirm"


def _native():
    user = ctypes.WinDLL("user32", use_last_error=True)
    for name, args, result in (
        ("GetCursorPos", [ctypes.POINTER(wintypes.POINT)], wintypes.BOOL),
        ("WindowFromPoint", [wintypes.POINT], wintypes.HWND),
        ("GetAncestor", [wintypes.HWND, wintypes.UINT], wintypes.HWND),
        ("GetAsyncKeyState", [ctypes.c_int], ctypes.c_short),
        ("GetWindowLongW", [wintypes.HWND, ctypes.c_int], wintypes.LONG),
        ("SetWindowLongW", [wintypes.HWND, ctypes.c_int, wintypes.LONG], wintypes.LONG),
        ("MonitorFromWindow", [wintypes.HWND, wintypes.DWORD], wintypes.HANDLE),
        ("SetWindowPos", [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                          ctypes.c_int, ctypes.c_int, wintypes.UINT], wintypes.BOOL),
    ):
        function = getattr(user, name); function.argtypes, function.restype = args, result
    return user


def pick_region(expected: dict, folder: Path, progress, *, preview=None, match_mode="position", title="添加 / 修改区域"):
    """Run inside a cancellable LibraryTask process; UIA never blocks the Tk loop."""
    import tkinter as tk
    from tkinter import ttk
    import uiautomation as auto
    from pynput import keyboard, mouse
    from .ctrl import Controller
    from .draw import ScreenLineBox
    from .library_actions import activate_application, bind_window, diagnose_uia_error, owned_windows, resolve_record
    from .mcp import native_windows, scanner
    from .mcp.gui_icons import apply_window_icon
    from .mcp.protocol import location_from_xpath
    from .utils import get_control_xpath, observe_control

    user = _native()
    updates, commands = Queue(), Queue()
    stopped, cancelled, capture_ready = threading.Event(), threading.Event(), threading.Event()
    displayed = {"token": 0}
    allowed_windows = {expected["handle"]: expected}
    result = {}
    root = tk.Tk(); root.withdraw(); apply_window_icon(root)
    root.title("easy-uiauto · "+title)
    root.resizable(False, False); root.attributes("-topmost", True)
    body = ttk.Frame(root, padding=12); body.pack(fill="both", expand=True)
    ttk.Label(body, text="悬浮选择控件 · 滚轮切换层级", font=("Microsoft YaHei UI", 11, "bold")).grid(
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
    values = {}
    for row, (label, key) in enumerate((("AutomationID", "automation_id"), ("Name", "native_name"),
                                       ("ClassName", "class_name"), ("ControlType", "control_type")), 1):
        ttk.Label(body, text=label, width=15).grid(row=row, column=0, sticky="nw", pady=4)
        values[key] = tk.StringVar(root, value="等待鼠标移入所选软件")
        ttk.Label(body, textvariable=values[key], wraplength=365, width=48).grid(row=row, column=1, sticky="w", pady=4)
    status = tk.StringVar(root, value="上滚：父节点；下滚：子节点。Ctrl 确认，Esc 取消。")
    ttk.Label(body, textvariable=status, wraplength=480).grid(row=5, column=0, columnspan=2, sticky="w", pady=10)
    cancel = ttk.Button(body, text="取消选择（Esc）", command=cancelled.set)
    cancel.grid(row=6, column=0, columnspan=2, sticky="e")
    box = ScreenLineBox(root=root, mode="manual", control=None, line_width=2, color="#ff1744")

    # Place the inspector in the target monitor's top-right work area.
    class MonitorInfo(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("monitor", wintypes.RECT),
                    ("work", wintypes.RECT), ("flags", wintypes.DWORD)]
    info = MonitorInfo(); info.cbSize = ctypes.sizeof(info)
    user.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MonitorInfo)]
    if not user.GetMonitorInfoW(user.MonitorFromWindow(expected["handle"], 2), ctypes.byref(info)):
        root.destroy(); raise ctypes.WinError(ctypes.get_last_error())
    root.update_idletasks()
    root.geometry(f"+{max(info.work.left, info.work.right-root.winfo_reqwidth()-16)}+{info.work.top+16}")
    root.update_idletasks()
    hwnd = user.GetAncestor(root.winfo_id(), 2)
    # Showing an inspector must not steal activation from the selected software.
    user.SetWindowLongW(hwnd, -20, user.GetWindowLongW(hwnd, -20) | 0x08000000 | 0x80)

    def point():
        p = wintypes.POINT()
        if not user.GetCursorPos(ctypes.byref(p)): raise ctypes.WinError(ctypes.get_last_error())
        return p.x, p.y

    def over_target(position):
        handle = user.WindowFromPoint(wintypes.POINT(*position))
        return bool(handle and int(user.GetAncestor(handle, 2) or 0) in allowed_windows)

    keys = PickerKeys(k for k in (0x11, 0xA2, 0xA3, 0x1B) if user.GetAsyncKeyState(k) & 0x8000)

    def key_filter(message, data):
        consume, action = keys.event(int(data.vkCode), message in {0x100, 0x104})
        if action == "cancel": cancelled.set()
        elif action: commands.put((action, displayed["token"]))
        if consume: keyboard_listener.suppress_event()

    def mouse_filter(message, data):
        if message == 0x20A and over_target((data.pt.x, data.pt.y)):
            delta = ctypes.c_short((data.mouseData >> 16) & 0xFFFF).value
            if delta: commands.put(("wheel", delta))
            mouse_listener.suppress_event()

    keyboard_listener = keyboard.Listener(win32_event_filter=key_filter)
    mouse_listener = mouse.Listener(win32_event_filter=mouse_filter)
    listeners = []

    def inspect():
        try:
            with auto.UIAutomationInitializerInThread():
                main = bind_window(expected)
                activate_application(expected)
                path = ControlPath(main, auto.ControlsAreSame)
                last, token, current = None, 0, None
                selected_window, scope = expected, {"kind":"main"}
                inventory_at=0
                updates.put(("ready", None))
                while not stopped.is_set():
                    if cancelled.is_set() or (folder / "cancel").exists(): raise InterruptedError("已取消区域选择")
                    if time.monotonic()-inventory_at>.5:
                        allowed_windows.clear();allowed_windows.update({expected["handle"]:expected,
                            **{w["handle"]:w for w in owned_windows(expected)}})
                        inventory_at=time.monotonic()
                    position = point()
                    if not over_target(position):
                        if current is not None:
                            current = None; displayed["token"] = 0
                            updates.put(("clear", "请将鼠标移回所选软件窗口"))
                        last = None
                    elif last is None or abs(position[0]-last[0])+abs(position[1]-last[1]) > 2:
                        current = None; displayed["token"] = 0
                        updates.put(("clear", "正在读取当前位置控件…"))
                        try:
                            handle=int(user.GetAncestor(user.WindowFromPoint(wintypes.POINT(*position)),2) or 0)
                            selected_window=allowed_windows[handle]
                            scope=({"kind":"main"} if handle==expected["handle"] else
                                   {"kind":"owned","class_name":selected_window["class_name"],"title":selected_window["title"]})
                            path=ControlPath(main if handle==expected["handle"] else auto.ControlFromHandle(handle),auto.ControlsAreSame)
                            path.reset(auto.ControlFromPoint(*position))
                            current = selection_metadata(path.current)
                            last = position
                        except Exception as error:
                            updates.put(("clear", str(diagnose_uia_error(error,expected=expected,uia_operation=True)))); last = position
                        if current:
                            token += 1; updates.put(("candidate", {**current, "token": token, "level": path.level}))
                    try: action, value = commands.get(timeout=.04)
                    except Empty: continue
                    if action == "cancel": raise InterruptedError("已取消区域选择")
                    if action == "wheel" and current:
                        path.scroll(value)
                        current = selection_metadata(path.current)
                        token += 1; updates.put(("candidate", {**current, "token": token, "level": path.level}))
                    elif action == "confirm":
                        if not current or value != token or not over_target(point()):
                            updates.put(("notice", "尚无稳定选中的控件，请等待属性显示后再按 Ctrl")); continue
                        # Reuse the recorder's XPath generator for the selected ancestor itself.
                        metadata = selection_metadata(path.current)
                        xpath = get_control_xpath(path.current, match_mode=match_mode)
                        location = location_from_xpath(xpath, parameters={"match_mode": match_mode})
                        if not auto.ControlsAreSame(resolve_record({"location":location,"window_scope":scope},expected,main), path.current):
                            raise ValueError("区域 XPath 未能唯一回到当前控件，请重新选择")
                        bind_window(expected)
                        bounds = native_windows.matching_windows("", handle=selected_window["handle"])[0]["bounds"]
                        bounds = native_windows.visible_window_bounds(selected_window["handle"], bounds, class_name=selected_window["class_name"])
                        rectangle = normalized_selection(metadata["rect"], bounds)
                        if metadata != current:
                            updates.put(("notice", "控件属性或边界已变化，请移动鼠标重新选择")); continue
                        updates.put(("capture", None))
                        if not capture_ready.wait(3): raise TimeoutError("无法隐藏选择浮窗，已停止截图")
                        image = scanner._screenshot_window(bounds)
                        if selection_metadata(path.current) != metadata:
                            raise ValueError("截图时控件发生变化，请重新选择区域")
                        image_path = folder / "region-window.png"; image.save(image_path)
                        result.update(selection={**metadata, "location": location, "normalized_rect": rectangle,"window_scope":scope},
                                      capture={"image": str(image_path), "width": image.width, "height": image.height,
                                               "window_name": selected_window["title"], "window_class": selected_window["class_name"],
                                               "window_scope":scope})
                        updates.put(("done", None)); return
        except BaseException as error:
            result["error"] = error; updates.put(("done", None))

    def pump():
        if cancelled.is_set() or (folder / "cancel").exists():
            result["error"] = InterruptedError("已取消区域选择"); root.quit(); return
        for _ in range(40):
            try: event, data = updates.get_nowait()
            except Empty: break
            if event == "ready":
                for listener in (keyboard_listener, mouse_listener):
                    listener.start(); listeners.append(listener)
                root.update_idletasks()
                user.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0010 | 0x0001 | 0x0002 | 0x0040)
                progress("鼠标悬浮选择区域；上滚父级、下滚子级，Ctrl 确认，Esc 取消")
            elif event == "candidate":
                for key, variable in values.items(): variable.set(data[key] or "（空）")
                box.set_rect(data["rect"])
                status.set(f"红框为当前选择：从鼠标命中控件上溯 {data['level']} 层。上滚父级 / 下滚子级；Ctrl 确认，Esc 取消。")
                root.update_idletasks(); displayed["token"] = data["token"]
                if preview: preview(data)
            elif event == "clear":
                box.set_visible(False); displayed["token"] = 0
                for variable in values.values(): variable.set("—")
                status.set(data)
                if preview: preview({"token": 0})
            elif event == "notice": status.set(data)
            elif event == "capture":
                box.set_visible(False); root.withdraw(); root.update_idletasks()
                root.after(120, capture_ready.set)
            elif event == "done": root.quit(); return
        root.after(30, pump)

    def callback_error(kind, error, traceback):
        result["error"] = error; root.quit()
    root.report_callback_exception = callback_error
    root.protocol("WM_DELETE_WINDOW", cancelled.set)
    thread = threading.Thread(target=inspect, daemon=True)
    try:
        root.deiconify(); root.update_idletasks()
        hwnd = user.GetAncestor(root.winfo_id(), 2)
        user.SetWindowLongW(hwnd, -20, user.GetWindowLongW(hwnd, -20) | 0x08000000 | 0x80)
        user.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0010 | 0x0001 | 0x0002 | 0x0040)
        thread.start(); root.after(30, pump); root.mainloop()
    finally:
        stopped.set(); capture_ready.set()
        for listener in listeners:
            try: listener.stop(); listener.join(timeout=.5)
            except Exception as error: result.setdefault("error", error)
        box.destroy(); root.destroy()
    if "error" in result: raise result["error"]
    if "selection" not in result: raise InterruptedError("已取消区域选择")
    return result
