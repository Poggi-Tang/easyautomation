from __future__ import annotations

import ctypes
import sys
from ctypes import wintypes

import win32api
import win32con
import win32gui

handles = {}


# Common-controls constants are not exported by pywin32's win32con module.
ICC_LISTVIEW_CLASSES = 0x00000001
ICC_TREEVIEW_CLASSES = 0x00000002
ICC_TAB_CLASSES = 0x00000008
LVS_REPORT = 0x0001
LVS_SINGLESEL = 0x0004
LVS_SHOWSELALWAYS = 0x0008
TVS_HASBUTTONS = 0x0001
TVS_HASLINES = 0x0002
TVS_LINESATROOT = 0x0004
TVS_SHOWSELALWAYS = 0x0020
LVM_FIRST = 0x1000
LVM_INSERTITEMW = LVM_FIRST + 77
LVM_INSERTCOLUMNW = LVM_FIRST + 97
TVM_FIRST = 0x1100
TVM_INSERTITEMW = TVM_FIRST + 50
TCM_FIRST = 0x1300
TCM_INSERTITEMW = TCM_FIRST + 62
LVIF_TEXT = 0x0001
LVCF_FMT = 0x0001
LVCF_WIDTH = 0x0002
LVCF_TEXT = 0x0004
LVCFMT_LEFT = 0x0000
TVIF_TEXT = 0x0001
TCIF_TEXT = 0x0001
TVI_ROOT = ctypes.c_void_p(-0x10000 & ((1 << (ctypes.sizeof(ctypes.c_void_p) * 8)) - 1))
TVI_LAST = ctypes.c_void_p(-0x0FFFE & ((1 << (ctypes.sizeof(ctypes.c_void_p) * 8)) - 1))


class INITCOMMONCONTROLSEX(ctypes.Structure):
    _fields_ = [("dwSize", wintypes.DWORD), ("dwICC", wintypes.DWORD)]


class LVCOLUMNW(ctypes.Structure):
    _fields_ = [
        ("mask", wintypes.UINT),
        ("fmt", ctypes.c_int),
        ("cx", ctypes.c_int),
        ("pszText", wintypes.LPWSTR),
        ("cchTextMax", ctypes.c_int),
        ("iSubItem", ctypes.c_int),
        ("iImage", ctypes.c_int),
        ("iOrder", ctypes.c_int),
        ("cxMin", ctypes.c_int),
        ("cxDefault", ctypes.c_int),
        ("cxIdeal", ctypes.c_int),
    ]


class LVITEMW(ctypes.Structure):
    _fields_ = [
        ("mask", wintypes.UINT),
        ("iItem", ctypes.c_int),
        ("iSubItem", ctypes.c_int),
        ("state", wintypes.UINT),
        ("stateMask", wintypes.UINT),
        ("pszText", wintypes.LPWSTR),
        ("cchTextMax", ctypes.c_int),
        ("iImage", ctypes.c_int),
        ("lParam", wintypes.LPARAM),
        ("iIndent", ctypes.c_int),
        ("iGroupId", ctypes.c_int),
        ("cColumns", wintypes.UINT),
        ("puColumns", ctypes.POINTER(wintypes.UINT)),
        ("piColFmt", ctypes.POINTER(ctypes.c_int)),
        ("iGroup", ctypes.c_int),
    ]


class TVITEMW(ctypes.Structure):
    _fields_ = [
        ("mask", wintypes.UINT),
        ("hItem", wintypes.HANDLE),
        ("state", wintypes.UINT),
        ("stateMask", wintypes.UINT),
        ("pszText", wintypes.LPWSTR),
        ("cchTextMax", ctypes.c_int),
        ("iImage", ctypes.c_int),
        ("iSelectedImage", ctypes.c_int),
        ("cChildren", ctypes.c_int),
        ("lParam", wintypes.LPARAM),
    ]


class TVINSERTSTRUCTW(ctypes.Structure):
    _fields_ = [("hParent", wintypes.HANDLE), ("hInsertAfter", wintypes.HANDLE), ("item", TVITEMW)]


class TCITEMW(ctypes.Structure):
    _fields_ = [
        ("mask", wintypes.UINT),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("pszText", wintypes.LPWSTR),
        ("cchTextMax", ctypes.c_int),
        ("iImage", ctypes.c_int),
        ("lParam", wintypes.LPARAM),
    ]


user32 = ctypes.windll.user32
comctl32 = ctypes.windll.comctl32
user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.SendMessageW.restype = ctypes.c_ssize_t


def _send_struct(hwnd, message, wparam, value):
    return user32.SendMessageW(hwnd, message, wparam, ctypes.addressof(value))


def initialize_common_controls():
    controls = INITCOMMONCONTROLSEX()
    controls.dwSize = ctypes.sizeof(controls)
    controls.dwICC = ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES | ICC_TAB_CLASSES
    if not comctl32.InitCommonControlsEx(ctypes.byref(controls)):
        raise ctypes.WinError()


def populate_list_view(list_view):
    header_text = ctypes.create_unicode_buffer("Items")
    column = LVCOLUMNW(
        mask=LVCF_FMT | LVCF_WIDTH | LVCF_TEXT,
        fmt=LVCFMT_LEFT,
        cx=230,
        pszText=ctypes.cast(header_text, wintypes.LPWSTR),
    )
    _send_struct(list_view, LVM_INSERTCOLUMNW, 0, column)
    for index in range(40):
        item_text = ctypes.create_unicode_buffer(f"List item {index:02d}")
        item = LVITEMW(
            mask=LVIF_TEXT,
            iItem=index,
            iSubItem=0,
            pszText=ctypes.cast(item_text, wintypes.LPWSTR),
        )
        if _send_struct(list_view, LVM_INSERTITEMW, 0, item) == -1:
            raise ctypes.WinError()


def insert_tree_item(tree_view, text, parent):
    item_text = ctypes.create_unicode_buffer(text)
    insert = TVINSERTSTRUCTW()
    insert.hParent = parent
    insert.hInsertAfter = TVI_LAST
    insert.item.mask = TVIF_TEXT
    insert.item.pszText = ctypes.cast(item_text, wintypes.LPWSTR)
    return _send_struct(tree_view, TVM_INSERTITEMW, 0, insert)


def populate_tree_view(tree_view):
    root = insert_tree_item(tree_view, "Root node", TVI_ROOT)
    if not root:
        raise ctypes.WinError()
    insert_tree_item(tree_view, "Child node A", root)
    insert_tree_item(tree_view, "Child node B", root)


def populate_tab_control(tab_control):
    for index, text in enumerate(("General", "Advanced", "Diagnostics")):
        item_text = ctypes.create_unicode_buffer(text)
        item = TCITEMW(mask=TCIF_TEXT, pszText=ctypes.cast(item_text, wintypes.LPWSTR))
        if _send_struct(tab_control, TCM_INSERTITEMW, index, item) == -1:
            raise ctypes.WinError()


def window_proc(hwnd, message, wparam, lparam):
    if message == win32con.WM_COMMAND and win32api.LOWORD(wparam) == 102:
        value = win32gui.GetWindowText(handles["edit"])
        win32gui.SetWindowText(handles["status"], f"applied:{value}")
        return 0
    if message == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(hwnd, message, wparam, lparam)


def create_control(class_name, text, style, x, y, width, height, parent, control_id):
    return win32gui.CreateWindow(
        class_name,
        text,
        win32con.WS_CHILD | win32con.WS_VISIBLE | style,
        x,
        y,
        width,
        height,
        parent,
        control_id,
        win32api.GetModuleHandle(None),
        None,
    )


def main():
    title = sys.argv[1] if len(sys.argv) > 1 else "Easy UIAuto Win32 Fixture"
    initialize_common_controls()
    instance = win32api.GetModuleHandle(None)
    window_class = win32gui.WNDCLASS()
    window_class.lpfnWndProc = window_proc
    window_class.hInstance = instance
    window_class.hCursor = win32gui.LoadCursor(None, win32con.IDC_ARROW)
    window_class.hbrBackground = win32con.COLOR_WINDOW + 1
    window_class.lpszClassName = "EasyUIAutoWin32Fixture"
    win32gui.RegisterClass(window_class)

    window = win32gui.CreateWindow(
        window_class.lpszClassName,
        title,
        win32con.WS_OVERLAPPEDWINDOW,
        win32con.CW_USEDEFAULT,
        win32con.CW_USEDEFAULT,
        850,
        640,
        None,
        None,
        instance,
        None,
    )
    create_control("STATIC", "Input", 0, 20, 22, 60, 24, window, 100)
    handles["edit"] = create_control(
        "EDIT", "", win32con.WS_BORDER | win32con.ES_AUTOHSCROLL,
        90, 20, 360, 26, window, 101
    )
    create_control("BUTTON", "Apply", 0, 20, 65, 90, 30, window, 102)
    create_control(
        "BUTTON", "Enabled", win32con.BS_AUTOCHECKBOX, 130, 65, 100, 30, window, 103
    )
    handles["status"] = create_control("STATIC", "idle", 0, 20, 110, 430, 26, window, 104)
    combo = create_control(
        "COMBOBOX", "", win32con.CBS_DROPDOWNLIST, 20, 155, 210, 120, window, 105
    )
    win32gui.SendMessage(combo, win32con.CB_ADDSTRING, 0, "Alpha")
    win32gui.SendMessage(combo, win32con.CB_ADDSTRING, 0, "Beta")
    win32gui.SendMessage(combo, win32con.CB_SETCURSEL, 0, 0)

    handles["list"] = create_control(
        "SysListView32",
        "",
        LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | win32con.WS_BORDER | win32con.WS_TABSTOP,
        270,
        65,
        260,
        220,
        window,
        106,
    )
    populate_list_view(handles["list"])

    handles["tree"] = create_control(
        "SysTreeView32",
        "",
        TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS | win32con.WS_BORDER | win32con.WS_TABSTOP,
        550,
        65,
        250,
        220,
        window,
        107,
    )
    populate_tree_view(handles["tree"])

    handles["tabs"] = create_control(
        "SysTabControl32",
        "",
        win32con.WS_TABSTOP,
        20,
        315,
        780,
        80,
        window,
        108,
    )
    populate_tab_control(handles["tabs"])

    win32gui.ShowWindow(window, win32con.SW_SHOW)
    win32gui.UpdateWindow(window)
    print("READY", flush=True)
    win32gui.PumpMessages()


if __name__ == "__main__":
    main()
