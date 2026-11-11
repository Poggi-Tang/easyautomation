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


def menu_owner(handle: int) -> int:
    """Native popup menus may have no GW_OWNER; query their active menu loop."""
    class GUIThreadInfo(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD),
                    ("hwndActive", wintypes.HWND), ("hwndFocus", wintypes.HWND),
                    ("hwndCapture", wintypes.HWND), ("hwndMenuOwner", wintypes.HWND),
                    ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND),
                    ("rcCaret", wintypes.RECT)]
    user32, _callback = _api()
    info = GUIThreadInfo(cbSize=ctypes.sizeof(GUIThreadInfo))
    query = user32.GetGUIThreadInfo
    query.argtypes, query.restype = [wintypes.DWORD, ctypes.POINTER(GUIThreadInfo)], wintypes.BOOL
    thread = user32.GetWindowThreadProcessId(handle, None)
    if thread and query(thread, ctypes.byref(info)) and info.flags & 4:
        return int(info.hwndMenuOwner or 0)
    return 0


def allow_foreground(process_id: int) -> bool:
    """The foreground GUI grants only its own worker permission to activate a window."""
    try:
        user32=ctypes.WinDLL("user32",use_last_error=True)
        call=user32.AllowSetForegroundWindow
        call.argtypes,call.restype=[wintypes.DWORD],wintypes.BOOL
        return bool(call(int(process_id)))
    except (OSError,ValueError):return False


def request_close_window(handle: int) -> bool:
    """Ask a selected window to close asynchronously, honoring application prompts."""
    user32=ctypes.WinDLL('user32',use_last_error=True)
    post=user32.PostMessageW
    post.argtypes,post.restype=[wintypes.HWND,wintypes.UINT,wintypes.WPARAM,wintypes.LPARAM],wintypes.BOOL
    return bool(post(int(handle),0x0010,0,0))


def process_elevated(process_id: int) -> bool:
    """Read TokenElevation; access denial is unknown, never proof of elevation."""
    kernel=ctypes.WinDLL("kernel32",use_last_error=True)
    security=ctypes.WinDLL("advapi32",use_last_error=True)
    kernel.OpenProcess.argtypes=[wintypes.DWORD,wintypes.BOOL,wintypes.DWORD]
    kernel.OpenProcess.restype=wintypes.HANDLE
    kernel.CloseHandle.argtypes=[wintypes.HANDLE];kernel.CloseHandle.restype=wintypes.BOOL
    security.OpenProcessToken.argtypes=[wintypes.HANDLE,wintypes.DWORD,ctypes.POINTER(wintypes.HANDLE)]
    security.OpenProcessToken.restype=wintypes.BOOL
    security.GetTokenInformation.argtypes=[wintypes.HANDLE,ctypes.c_int,wintypes.LPVOID,
                                           wintypes.DWORD,ctypes.POINTER(wintypes.DWORD)]
    security.GetTokenInformation.restype=wintypes.BOOL
    process=kernel.OpenProcess(0x1000,False,process_id)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not process:raise ctypes.WinError(ctypes.get_last_error())
    token=wintypes.HANDLE()
    try:
        if not security.OpenProcessToken(process,8,ctypes.byref(token)):  # TOKEN_QUERY
            raise ctypes.WinError(ctypes.get_last_error())
        value,length=wintypes.DWORD(),wintypes.DWORD()
        if not security.GetTokenInformation(token,20,ctypes.byref(value),ctypes.sizeof(value),ctypes.byref(length)):
            raise ctypes.WinError(ctypes.get_last_error())
        return bool(value.value)
    finally:
        if token:kernel.CloseHandle(token)
        kernel.CloseHandle(process)
