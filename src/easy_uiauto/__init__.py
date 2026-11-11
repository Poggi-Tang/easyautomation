# -*- coding: utf-8 -*-
"""easy-uiauto - UI automation toolkit based on pyautogui and uiautomation"""

__version__ = "0.9.13"
__author__ = "Poggi-Tang"

from importlib import import_module
from typing import TYPE_CHECKING

# Preserve the public objects, but load their backends only when requested.
# Opening the knowledge manager must not initialize recording, UIA or OpenCV.
_EXPORT_MODULES = {
    **{name: ".ctrl" for name in (
        "Controller", "left_click", "right_click", "double_click", "centre_click",
        "mouse_left_press", "mouse_left_release", "mouse_right_press", "mouse_right_release",
        "mouse_move_pos", "mouse_move_control", "drag_control", "drag_control_by_control",
        "set_text", "input_text", "key_click", "key_press", "key_release", "key_group",
        "activate_window", "run_action")},
    "run_record": ".record", "RecordThread": ".record",
    **{name: ".utils" for name in (
        "find_control", "find_control_by_xpath", "get_control_info", "get_control_xpath",
        "compile_controls", "push_message", "set_top_window", "auto_scroll")},
    **{name: ".draw" for name in ("ScreenControlOverlay", "ScreenLineBox", "show_control_overlay")},
    **{name: ".visual" for name in ("crop_window_image", "segment_image", "segment_screen_rect")},
}

if TYPE_CHECKING:
    from .ctrl import (Controller, left_click, right_click, double_click, centre_click,
        mouse_left_press, mouse_left_release, mouse_right_press, mouse_right_release,
        mouse_move_pos, mouse_move_control, drag_control, drag_control_by_control,
        set_text, input_text, key_click, key_press, key_release, key_group, activate_window, run_action)
    from .record import run_record, RecordThread
    from .utils import (find_control, find_control_by_xpath, get_control_info, get_control_xpath,
        compile_controls, push_message, set_top_window, auto_scroll)
    from .draw import ScreenControlOverlay, ScreenLineBox, show_control_overlay
    from .visual import crop_window_image, segment_image, segment_screen_rect

__all__ = [
    "Controller",
    "ScreenControlOverlay",
    "ScreenLineBox",
    "show_control_overlay",
    "segment_image",
    "segment_screen_rect",
    "crop_window_image",
    "left_click",
    "right_click",
    "double_click",
    "centre_click",
    "mouse_left_press",
    "mouse_left_release",
    "mouse_right_press",
    "mouse_right_release",
    "mouse_move_pos",
    "mouse_move_control",
    "drag_control",
    "drag_control_by_control",
    "set_text",
    "input_text",
    "key_click",
    "key_press",
    "key_release",
    "key_group",
    "activate_window",
    "run_action",
    "run_record",
    "RecordThread",
    "find_control",
    "find_control_by_xpath",
    "get_control_info",
    "compile_controls",
    "push_message",
    "set_top_window",
    "auto_scroll",
]


def __getattr__(name):
    module = _EXPORT_MODULES.get(name)
    if module is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    value = getattr(import_module(module, __name__), name)
    globals()[name] = value
    return value


def __dir__():
    return sorted(set(globals()) | set(_EXPORT_MODULES))
