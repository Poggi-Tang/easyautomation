"""Disposable native Windows target for manual-layout GUI acceptance checks."""

import ctypes
from ctypes import wintypes
import argparse
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--menu", action="store_true")
    parser.add_argument("--workflow", action="store_true")
    parser.add_argument("--picker", action="store_true", help="Log target input to verify picker event suppression")
    parser.add_argument("--events", type=Path)
    args = parser.parse_args()
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    try:
        user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except (AttributeError, OSError):
        pass
    callback_type = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT,
                                      wintypes.WPARAM, wintypes.LPARAM)

    class WindowClass(ctypes.Structure):
        _fields_ = [("style", wintypes.UINT), ("proc", callback_type),
                    ("class_extra", ctypes.c_int), ("window_extra", ctypes.c_int),
                    ("instance", wintypes.HINSTANCE), ("icon", wintypes.HICON),
                    ("cursor", wintypes.HANDLE), ("brush", wintypes.HBRUSH),
                    ("menu", wintypes.LPCWSTR), ("name", wintypes.LPCWSTR)]

    user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.DefWindowProcW.restype = ctypes.c_ssize_t
    user32.CreateWindowExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
                                     wintypes.DWORD, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                     wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, ctypes.c_void_p]
    user32.CreateWindowExW.restype = wintypes.HWND
    user32.DestroyWindow.argtypes = [wintypes.HWND]
    user32.RegisterClassW.argtypes = [ctypes.POINTER(WindowClass)]
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.UpdateWindow.argtypes = [wintypes.HWND]
    user32.SetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPCWSTR]
    kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
    kernel32.GetModuleHandleW.restype = wintypes.HMODULE
    status = None
    file_menu = None
    dialog = None
    events = []
    if args.picker and args.events:
        args.events.parent.mkdir(parents=True, exist_ok=True)
        args.events.write_text("[]", encoding="utf-8")

    def event(name):
        events.append(name)
        if args.events:
            args.events.parent.mkdir(parents=True, exist_ok=True)
            temporary = args.events.with_suffix(".tmp")
            temporary.write_text(json.dumps(events), encoding="utf-8")
            temporary.replace(args.events)

    @callback_type
    def procedure(handle, message, wparam, lparam):
        nonlocal dialog
        if message == 2 and handle == main_handle:
            user32.PostQuitMessage(0)
            return 0
        if args.workflow and message == 0x007B and handle == main_handle:
            event("context-open")
            point = wintypes.POINT()
            user32.GetCursorPos(ctypes.byref(point))
            user32.TrackPopupMenu(file_menu, 0, point.x, point.y, 0, main_handle, None)
            return 0
        if args.workflow and message == 273:
            command = int(wparam) & 0xffff
            if command == 301:
                event("new-project-dialog")
                dialog = user32.CreateWindowExW(0, name, "新建项目", 0x00CF0000,
                    150, 180, 510, 210, main_handle, None, instance, None)
                for cls_name, text, x, y, width, height, identifier, style in (
                    ("STATIC", "选择存储路径", 20, 25, 110, 24, 401, 0),
                    ("EDIT", "C:/Example", 140, 22, 230, 26, 402, 0x00800000 | 0x0800),
                    ("BUTTON", "选择", 380, 22, 70, 26, 403, 0),
                    ("STATIC", "设置项目名称", 20, 65, 110, 24, 404, 0),
                    ("EDIT", "project", 140, 62, 310, 26, 405, 0x00800000),
                    ("BUTTON", "取消", 290, 110, 70, 28, 406, 0),
                    ("BUTTON", "确定", 380, 110, 70, 28, 407, 0),
                ):
                    user32.CreateWindowExW(0, cls_name, text, 0x50000000 | style,
                        x, y, width, height, dialog, identifier, instance, None)
                user32.ShowWindow(dialog, 5); user32.UpdateWindow(dialog)
            elif command in {406, 407}:
                event("cancel" if command == 406 else "confirm")
                user32.DestroyWindow(dialog)
            return 0
        if args.menu and message == 0x117 and file_menu and int(wparam) == file_menu:
            event("file-open")
        if message == 273 and status:
            event({102: "save", 201: "new", 202: "open"}.get(int(wparam) & 0xffff, "other"))
            user32.SetWindowTextW(status, "A button was clicked")
        return user32.DefWindowProcW(handle, message, wparam, lparam)

    instance = kernel32.GetModuleHandleW(None)
    name = "EasyUIAutoWorkflowFixture" if args.workflow else "EasyUIAutoMenuFixture" if args.menu else "EasyUIAutoKnowledgeFixture"
    cls = WindowClass(0, procedure, 0, 0, instance, None, None, 6, None, name)
    if not user32.RegisterClassW(ctypes.byref(cls)):
        raise ctypes.WinError(ctypes.get_last_error())
    title = "Knowledge Workflow Target" if args.workflow else "Knowledge Menu Replay Target" if args.menu else "Knowledge Library Acceptance Target"
    main_handle = None
    handle = user32.CreateWindowExW(0, name, title, 0x00CF0000,
                                    70, 80, 720, 500, None, None, instance, None)
    if not handle: raise ctypes.WinError(ctypes.get_last_error())
    main_handle = handle
    if args.menu or args.workflow:
        user32.CreateMenu.restype = wintypes.HMENU
        user32.CreatePopupMenu.restype = wintypes.HMENU
        user32.AppendMenuW.argtypes = [wintypes.HMENU, wintypes.UINT, ctypes.c_size_t, wintypes.LPCWSTR]
        user32.SetMenu.argtypes = [wintypes.HWND, wintypes.HMENU]
        menu, file_menu = int(user32.CreateMenu()), int(user32.CreatePopupMenu())
        if args.workflow:
            user32.TrackPopupMenu.argtypes = [wintypes.HMENU, wintypes.UINT, ctypes.c_int, ctypes.c_int,
                                             ctypes.c_int, wintypes.HWND, ctypes.c_void_p]
            for n, label in enumerate(("新建项目", "打开项目", "导入", "关闭所有项目", "粘贴", "展开所有", "折叠所有", "重新加载"), 301):
                user32.AppendMenuW(file_menu, 0, n, label)
        else:
            user32.AppendMenuW(file_menu, 0, 201, "&New")
            user32.AppendMenuW(file_menu, 0, 202, "&Open")
            user32.AppendMenuW(menu, 0x10, file_menu, "&File")
            user32.SetMenu(handle, menu)

    def child(class_name, text, x, y, width, height, identifier, style=0):
        return user32.CreateWindowExW(0, class_name, text, 0x50000000 | style,
                                      x, y, width, height, handle, identifier, instance, None)

    if args.workflow:
        child("STATIC", "项目管理（右键空白区域）", 20, 18, 320, 30, 100)
    else:
        child("STATIC", "Toolbar", 18, 12, 650, 25, 100)
        child("BUTTON", "Open", 20, 45, 120, 38, 101)
        child("BUTTON", "Save", 155, 45, 120, 38, 102)
        child("BUTTON", "Options", 290, 45, 120, 38, 103)
        child("STATIC", "Content", 18, 128, 650, 25, 104)
        child("EDIT", "Sample input", 20, 164, 620, 150, 105, 0x00800000 | 0x0004)
        status = child("STATIC", "No actions executed", 20, 380, 640, 30, 106)
    user32.ShowWindow(handle, 5); user32.UpdateWindow(handle)
    print(f"fixture_ready hwnd={int(handle)}", flush=True)
    message = wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(message), None, 0, 0) > 0:
        if args.picker and message.message == 0x20A: event("wheel-reached-target")
        if args.picker and message.message in {0x100, 0x104} and message.wParam in {0x11, 0xA2, 0xA3, 0x1B}:
            event("picker-key-reached-target")
        user32.TranslateMessage(ctypes.byref(message))
        user32.DispatchMessageW(ctypes.byref(message))


if __name__ == "__main__":
    main()
