"""Top-level window inventory without querying the desktop UIA provider tree."""

from __future__ import annotations

import ctypes
from ctypes import wintypes


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
