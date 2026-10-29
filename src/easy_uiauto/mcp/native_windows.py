"""Top-level window inventory without querying the desktop UIA provider tree."""

from __future__ import annotations

import ctypes
from ctypes import wintypes


def _dwm_api():
    dwm = ctypes.WinDLL("dwmapi", use_last_error=True)
    get = dwm.DwmGetWindowAttribute
    get.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD]
    get.restype = ctypes.c_long
    return get


def visible_window_bounds(handle: int, bounds: dict, *, class_name: str = "") -> dict:
    """Exclude invisible resize margins using DWM's physical screen coordinates.

    Only accept a nonempty rectangle contained in the supplied physical window
    bounds. This prevents a stale handle or a DPI coordinate mismatch from
    moving the crop. Desktop and taskbar surfaces keep their original bounds.
    """
    original = dict(bounds)
    if not handle or class_name in {"Progman", "WorkerW", "Shell_TrayWnd", "Shell_SecondaryTrayWnd"}:
        return original
    try:
        left, top, right, bottom = (int(bounds[key]) for key in ("left", "top", "right", "bottom"))
        if right <= left or bottom <= top:
            return original
        frame = wintypes.RECT()
        if _dwm_api()(int(handle), 9, ctypes.byref(frame), ctypes.sizeof(frame)) != 0:
            return original
        if not (left <= frame.left < frame.right <= right
                and top <= frame.top < frame.bottom <= bottom):
            return original
        return dict(left=frame.left, top=frame.top, right=frame.right, bottom=frame.bottom,
                    width=frame.right - frame.left, height=frame.bottom - frame.top)
    except (OSError, AttributeError, TypeError, ValueError, KeyError):
        return original


def _api():
    # Private DLL signatures must not change pygetwindow's EnumWindows callback
    # types on the shared ctypes.windll.user32 instance.
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    callback = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    signatures = {
        "EnumWindows": ((callback, wintypes.LPARAM), wintypes.BOOL),
        "GetWindowTextLengthW": ((wintypes.HWND,), ctypes.c_int),
        "GetWindowTextW": ((wintypes.HWND, wintypes.LPWSTR, ctypes.c_int), ctypes.c_int),
        "GetClassNameW": ((wintypes.HWND, wintypes.LPWSTR, ctypes.c_int), ctypes.c_int),
        "IsWindowVisible": ((wintypes.HWND,), wintypes.BOOL),
        "IsWindowEnabled": ((wintypes.HWND,), wintypes.BOOL),
        "IsWindow": ((wintypes.HWND,), wintypes.BOOL),
        "IsIconic": ((wintypes.HWND,), wintypes.BOOL),
        "ShowWindow": ((wintypes.HWND, ctypes.c_int), wintypes.BOOL),
        "SetForegroundWindow": ((wintypes.HWND,), wintypes.BOOL),
        "GetWindowRect": ((wintypes.HWND, ctypes.POINTER(wintypes.RECT)), wintypes.BOOL),
        "GetWindowThreadProcessId": ((wintypes.HWND, ctypes.POINTER(wintypes.DWORD)),
                                     wintypes.DWORD),
        "GetWindow": ((wintypes.HWND, ctypes.c_uint), wintypes.HWND),
        "GetForegroundWindow": ((), wintypes.HWND),
    }
    for name, (arguments, result) in signatures.items():
        function = getattr(user32, name)
        function.argtypes, function.restype = arguments, result
    return user32, callback


def enumerate_windows(*, visible_only=False):
    """Read native metadata in z-order; never bind UIA to unrelated applications.

    Hidden/untitled windows remain available for exact-handle observations and
    Explorer shell discovery. Windows that disappear during enumeration are
    skipped. Failure of EnumWindows itself is an error, not an empty desktop.
    """
    user32, callback_type = _api()
    foreground = int(user32.GetForegroundWindow() or 0)
    found, callback_errors = [], []

    @callback_type
    def visit(handle, _parameter):
        try:
            if not user32.IsWindow(handle):
                return True
            visible = bool(user32.IsWindowVisible(handle))
            if visible_only and not visible:
                return True
            length = user32.GetWindowTextLengthW(handle)
            title = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(handle, title, len(title))
            class_name = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(handle, class_name, len(class_name))
            rectangle, process = wintypes.RECT(), wintypes.DWORD()
            if not user32.GetWindowRect(handle, ctypes.byref(rectangle)):
                return True
            user32.GetWindowThreadProcessId(handle, ctypes.byref(process))
            if not user32.IsWindow(handle):
                return True
            bounds = dict(left=rectangle.left, top=rectangle.top,
                          right=rectangle.right, bottom=rectangle.bottom,
                          width=max(0, rectangle.right - rectangle.left),
                          height=max(0, rectangle.bottom - rectangle.top))
            found.append(dict(handle=int(handle), title=title.value, class_name=class_name.value,
                              process_id=int(process.value),
                              enabled=bool(user32.IsWindowEnabled(handle)),
                              visible=visible, bounds=bounds, z_order=len(found),
                              owner_handle=int(user32.GetWindow(handle, 4) or 0),
                              foreground=int(handle) == foreground))
            return True
        except Exception as error:
            # Never let ctypes swallow a Python callback error as a successful
            # partial enumeration (which could falsely confirm disappearance).
            callback_errors.append(error)
            return False

    completed = user32.EnumWindows(visit, 0)
    if callback_errors:
        raise RuntimeError("Could not enumerate native windows") from callback_errors[0]
    if not completed:
        raise ctypes.WinError(ctypes.get_last_error())
    return found


def matching_windows(title, *, handle=0, partial=False):
    windows = enumerate_windows()
    if handle:
        return [window for window in windows if window["handle"] == handle]
    exact = [window for window in windows if window["title"] == title]
    if exact or not partial:
        return exact
    return [window for window in windows if window["visible"] and window["title"]
            and title.casefold() in window["title"].casefold()]
