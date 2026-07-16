from __future__ import annotations

import sys

import win32api
import win32con
import win32gui

handles = {}


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
        520,
        320,
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

    win32gui.ShowWindow(window, win32con.SW_SHOW)
    win32gui.UpdateWindow(window)
    print("READY", flush=True)
    win32gui.PumpMessages()


if __name__ == "__main__":
    main()
