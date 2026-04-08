# -*- coding: utf-8 -*-
"""easy-uiauto - UI automation toolkit based on pyautogui and uiautomation"""

__version__ = "0.1.2"
__author__ = "Poggi-Tang"

from .ctrl import (
    Controller,
    left_click,
    right_click,
    double_click,
    centre_click,
    mouse_left_press,
    mouse_left_release,
    mouse_right_press,
    mouse_right_release,
    mouse_move_pos,
    mouse_move_control,
    drag_control,
    drag_control_by_control,
    set_text,
    input_text,
    key_click,
    key_press,
    key_release,
    key_group,
    activate_window,
    run_action,
)

from .record import run_record, RecordThread

from .utils import (
    find_control,
    find_control_by_xpath,
    get_control_info,
    get_control_xpath,
    compile_controls,
    push_message,
    set_top_window,
    auto_scroll,
)

from .draw import ScreenLineBox

__all__ = [
    "Controller",
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
    "ScreenLineBox",
]
