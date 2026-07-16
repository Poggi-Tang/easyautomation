# -*- coding: utf-8 -*-
# @Name:      ui_ctrl.py
# @Author:    tang
# @Date:      2025/9/23-9:55
# @depict:    控制
import ast
import ctypes
import threading
import time
from collections.abc import Callable
from ctypes import wintypes

import pyautogui
import pyperclip
import uiautomation

from .draw import get_visible_rect_map_by_control
from .highlight import HighlightProcess
from .utils import (
    auto_scroll,
    correct_ctrl_position,
    find_control,
    get_control_coordinates,
    package_location,
    push_message,
    set_top_window,
)


def _show_thread(rect_map, show_time):
    """在独立解释器进程中显示一次性 overlay，不在工作线程创建 Tk。"""
    worker = HighlightProcess(rect_map, name="easy-uiauto-action-highlight")
    try:
        worker.start(timeout=5)
        time.sleep(max(0, show_time) / 1000)
    except Exception as e:
        try:
            push_message(f"show_ctrl_area 异常: {e}")
        except Exception:
            pass
    finally:
        worker.stop(timeout=2)


def show_ctrl_area(control_obj, show_time=300):
    """非阻塞地显示控件区域：在后台线程启动 Tk 主循环。"""
    try:
        rect_map = get_visible_rect_map_by_control(control_obj)
        if not rect_map or rect_map.get("width", 0) <= 0 or rect_map.get("height", 0) <= 0:
            return
        t = threading.Thread(target=_show_thread, args=(rect_map, show_time), daemon=True)
        t.start()
    except Exception as e:
        try:
            push_message(f"启动 show_ctrl_area 线程失败: {e}")
        except Exception:
            pass


def locate_and_prepare(LOCATION, PARAMETERS, timeout=3.0, interval=0.2):
    """尝试在限定时间内定位控件并返回 (control, x, y)。失败则返回 (None, None, None) 并记录日志。

    LOCATION: dict 或其他 find_control 可接受的定位参数
    PARAMETERS: dict，用于覆盖坐标
    timeout: 最大等待秒数
    interval: 重试间隔秒数
    """
    PARAMETERS = _parse_parameters(PARAMETERS)
    start = time.monotonic()
    last_exc = None
    while time.monotonic() - start <= timeout:
        try:
            control = find_control(LOCATION)
            if control:
                try:
                    correct_x, correct_y = correct_ctrl_position(control)
                except Exception as e:
                    correct_x, correct_y = None, None
                    last_exc = e
                if correct_x is None or correct_y is None:
                    time.sleep(interval)
                    continue
                x = int(PARAMETERS.get("x", correct_x))
                y = int(PARAMETERS.get("y", correct_y))
                x = correct_x if x == -1 else x
                y = correct_y if y == -1 else y
                return control, x, y
        except Exception as e:
            last_exc = e
        time.sleep(interval)
    # 超时未找到
    try:
        push_message(f"定位控件超时: {LOCATION} 错误: {last_exc}")
    except Exception:
        pass
    return None, None, None


def prepare_control(WindowName, Name, ClassName, ControlType, foundIndex,
                    AutomationId, Xpath, Img, PARAMETERS):
    """兼容旧接口的封装：使用 package_location 构造 LOCATION 并调用 locate_and_prepare。失败抛出异常。"""
    LOCATION = package_location(WindowName, Name, ClassName, ControlType,
                                foundIndex, AutomationId, Xpath, Img, PARAMETERS)
    control, x, y = locate_and_prepare(LOCATION, PARAMETERS)
    if not control:
        raise Exception('未找到控件或超时')
    # 非阻塞显示控件区域
    try:
        show_ctrl_area(control)
    except Exception as e:
        try:
            push_message(f"显示控件区域失败: {e}")
        except Exception:
            pass
    return control, x, y


# 恢复消息类型定义与更新函数（此前被替换时意外移除）
class _MESSAGE:
    """
    日志级别
    Args：
        - ERROR:错误
        - WARNING:警告
        - INFO:正确
        - OTHER:其他
    """
    ERROR = 0
    WARNING = 1
    INFO = 2
    OTHER = 3

_message_state = threading.local()
_action_mechanism_state = threading.local()


def get_message_type():
    """
    获取当前消息类型

    :return
        int: 当前消息类型的值
    """
    return getattr(_message_state, "value", _MESSAGE.OTHER)


def update_message_type(current_message_type):
    """
    更新当前消息类型

    :param
        current_message_type (int): 要设置的消息类型值
    """
    _message_state.value = current_message_type


def get_action_mechanism():
    return getattr(_action_mechanism_state, "value", None)


def _set_action_mechanism(value):
    _action_mechanism_state.value = value


def _parse_parameters(parameters):
    if parameters is None:
        return {}
    if isinstance(parameters, dict):
        return parameters
    if isinstance(parameters, str):
        try:
            value = ast.literal_eval(parameters)
        except (SyntaxError, ValueError) as exc:
            raise ValueError("PARAMETERS 必须是字典或字典字符串") from exc
        if isinstance(value, dict):
            return value
    raise ValueError("PARAMETERS 必须是字典或字典字符串")


def _absolute_control_position(control, x, y):
    point = control.MoveCursorToInnerPos(x=x, y=y, simulateMove=False)
    if not point:
        raise ValueError("控件没有可用的点击区域")
    return int(point[0]), int(point[1])


def _require_control(location):
    control = find_control(location)
    if not control:
        raise LookupError("未找到控件")
    return control


def _activate_target_window(window_name):
    if window_name and not set_top_window(window_name):
        raise RuntimeError(f"无法激活目标窗口：{window_name}")


def _focus_location(location):
    control = _resolve_location(location)
    if control is not None:
        _activate_target_window(location.get("WindowName"))
        control.Click(waitTime=0)
    return control


def _resolve_location(location):
    locator_fields = (
        location.get("WindowName"),
        location.get("Name"),
        location.get("ClassName"),
        location.get("ControlType"),
        location.get("AutomationId"),
        location.get("Xpath"),
        location.get("Img"),
    )
    if not any(locator_fields):
        return None
    return _require_control(location)


def _invoke_semantic_action(control):
    control_type = getattr(control, "ControlTypeName", "")
    try:
        if control_type == "CheckBoxControl":
            pattern = _get_pattern(control, "GetTogglePattern", uiautomation.PatternId.TogglePattern)
            return bool(pattern and pattern.Toggle())
        if control_type in {"RadioButtonControl", "ListItemControl", "TabItemControl"}:
            pattern = _get_pattern(
                control,
                "GetSelectionItemPattern",
                uiautomation.PatternId.SelectionItemPattern,
            )
            return bool(pattern and pattern.Select())
        if control_type in {"ButtonControl", "MenuItemControl", "HyperlinkControl"}:
            pattern = _get_pattern(control, "GetInvokePattern", uiautomation.PatternId.InvokePattern)
            return bool(pattern and pattern.Invoke())
    except Exception:
        return False
    return False


def _set_control_value(control, value):
    if control is None:
        return False
    try:
        pattern = _get_pattern(control, "GetValuePattern", uiautomation.PatternId.ValuePattern)
        if pattern is None:
            return False
        if getattr(pattern, "IsReadOnly", False):
            return False
        expected = str(value)
        if not pattern.SetValue(expected):
            return False
        deadline = time.monotonic() + 0.25
        while time.monotonic() <= deadline:
            if pattern.Value == expected:
                return True
            time.sleep(0.01)
        return False
    except Exception:
        return False


def _get_pattern(control, convenience_method, pattern_id):
    method = getattr(control, convenience_method, None)
    if callable(method):
        return method()
    return control.GetPattern(pattern_id)


def _coerce_optional_bool(value, parameter_name):
    if value is None or isinstance(value, bool):
        return value
    if value in (0, 1):
        return bool(value)
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "1", "yes", "on"}:
            return True
        if normalized in {"false", "0", "no", "off"}:
            return False
    raise ValueError(f"{parameter_name} 必须是布尔值")


def _safe_control_attr(control, name, default=""):
    try:
        return getattr(control, name, default)
    except Exception:
        return default


def _is_qt_control(control):
    framework = str(_safe_control_attr(control, "FrameworkId", "")).lower()
    class_name = str(_safe_control_attr(control, "ClassName", ""))
    automation_id = str(_safe_control_attr(control, "AutomationId", ""))
    return (
        framework == "qt"
        or automation_id.startswith("QApplication.")
        or (not framework and class_name.startswith("Q"))
    )


def _message_click_control(control, *, double=False):
    """Deliver a click to the owning window without moving the physical cursor."""
    try:
        top_level = control.GetTopLevelControl()
        hwnd = int(_safe_control_attr(top_level, "NativeWindowHandle", 0) or 0)
        rect = control.BoundingRectangle
        if not hwnd or rect.right <= rect.left or rect.bottom <= rect.top:
            return False
        point = wintypes.POINT(
            int((rect.left + rect.right) // 2),
            int((rect.top + rect.bottom) // 2),
        )
        user32 = ctypes.windll.user32
        user32.ScreenToClient.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.POINT)]
        user32.ScreenToClient.restype = wintypes.BOOL
        user32.SendMessageTimeoutW.argtypes = [
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
            wintypes.UINT,
            wintypes.UINT,
            ctypes.POINTER(ctypes.c_size_t),
        ]
        user32.SendMessageTimeoutW.restype = ctypes.c_size_t
        if not user32.ScreenToClient(hwnd, ctypes.byref(point)):
            return False
        lparam = (int(point.y) & 0xFFFF) << 16 | (int(point.x) & 0xFFFF)

        def send(message, wparam=0):
            result = ctypes.c_size_t()
            return bool(
                user32.SendMessageTimeoutW(
                    hwnd,
                    message,
                    wparam,
                    lparam,
                    0x0002,  # SMTO_ABORTIFHUNG
                    1000,
                    ctypes.byref(result),
                )
            )

        if not send(0x0200):  # WM_MOUSEMOVE
            return False
        for _ in range(2 if double else 1):
            if not send(0x0201, 0x0001):  # WM_LBUTTONDOWN / MK_LBUTTON
                return False
            time.sleep(0.04)
            if not send(0x0202):  # WM_LBUTTONUP
                return False
            time.sleep(0.06)
        return True
    except Exception:
        return False


def _named_item_candidates(control, item_name):
    candidates = []
    sources = [control]
    try:
        top_level = control.GetTopLevelControl()
        if top_level is not control:
            sources.append(top_level)
    except Exception:
        pass
    for source in sources:
        if hasattr(source, "ListItemControl"):
            try:
                candidates.append(source.ListItemControl(Name=item_name, searchDepth=10))
            except Exception:
                continue
    return candidates


def _wait_for_control_value(control, expected, timeout=0.75):
    deadline = time.monotonic() + timeout
    while time.monotonic() <= deadline:
        try:
            control.Refind(maxSearchSeconds=0, raiseException=False)
        except Exception:
            pass
        try:
            pattern = _get_pattern(
                control, "GetValuePattern", uiautomation.PatternId.ValuePattern
            )
            if pattern is not None and pattern.Value == expected:
                return True
        except Exception:
            pass
        time.sleep(0.02)
    return False


def _select_named_item_by_message(control, item_name, candidates):
    if not _is_qt_control(control):
        return False

    def visible_items():
        items = []
        expected_process = int(_safe_control_attr(control, "ProcessId", 0) or 0)
        for candidate in candidates + _named_item_candidates(control, item_name):
            try:
                if not candidate.Exists(0):
                    continue
                process_id = int(_safe_control_attr(candidate, "ProcessId", 0) or 0)
                if expected_process and process_id and process_id != expected_process:
                    continue
                rect = candidate.BoundingRectangle
                if rect.right > rect.left and rect.bottom > rect.top and not bool(
                    _safe_control_attr(candidate, "IsOffscreen", False)
                ):
                    items.append(candidate)
            except Exception:
                continue
        return items

    items = visible_items()
    if not items:
        if not _message_click_control(control):
            return False
        deadline = time.monotonic() + 1.0
        while time.monotonic() <= deadline and not items:
            items = visible_items()
            if not items:
                time.sleep(0.03)
    for candidate in items:
        if _message_click_control(candidate) and _wait_for_control_value(control, item_name):
            return True
    return False


def _wait_for_expand_state(control, target_state, timeout=0.35):
    deadline = time.monotonic() + timeout
    while time.monotonic() <= deadline:
        try:
            control.Refind(maxSearchSeconds=0, raiseException=False)
        except Exception:
            pass
        try:
            pattern = _get_pattern(
                control,
                "GetExpandCollapsePattern",
                uiautomation.PatternId.ExpandCollapsePattern,
            )
            if pattern is not None and int(pattern.ExpandCollapseState) == target_state:
                return True
        except Exception:
            pass
        time.sleep(0.02)
    return False


def _select_named_item(control, item_name):
    try:
        expand = _get_pattern(
            control,
            "GetExpandCollapsePattern",
            uiautomation.PatternId.ExpandCollapsePattern,
        )
        if expand is not None:
            expand.Expand(waitTime=0)
    except Exception:
        pass

    candidates = _named_item_candidates(control, item_name)
    for candidate in candidates:
        try:
            if not candidate.Exists(1):
                continue
            pattern = _get_pattern(
                candidate,
                "GetSelectionItemPattern",
                uiautomation.PatternId.SelectionItemPattern,
            )
            if pattern is not None and pattern.Select(waitTime=0):
                is_selected = False
                try:
                    is_selected = bool(pattern.IsSelected)
                except Exception:
                    pass
                try:
                    value_pattern = _get_pattern(
                        control, "GetValuePattern", uiautomation.PatternId.ValuePattern
                    )
                except Exception:
                    value_pattern = None
                if value_pattern is not None:
                    if value_pattern.Value == item_name:
                        _set_action_mechanism("selection_pattern")
                        return True
                    continue
                if is_selected:
                    _set_action_mechanism("selection_pattern")
                    return True
        except Exception:
            continue
    if hasattr(control, "Select"):
        if control.Select(item_name):
            try:
                value_pattern = _get_pattern(
                    control, "GetValuePattern", uiautomation.PatternId.ValuePattern
                )
            except Exception:
                value_pattern = None
            if value_pattern is None or value_pattern.Value == item_name:
                _set_action_mechanism("provider_select")
                return True
    selected = _select_named_item_by_message(control, item_name, candidates)
    if selected:
        _set_action_mechanism("qt_window_message")
    return selected


class Controller:
    """将原有顶层控制函数迁移为类方法，保持行为不变并复用公共构造逻辑。"""

    @classmethod
    def _make_location(cls, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img, PARAMETERS):
        """构造 LOCATION 字典，集中管理调用点。"""
        return package_location(WindowName, Name, ClassName, ControlType,
                                foundIndex, AutomationId, Xpath, Img, PARAMETERS)

    @classmethod
    def left_click(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                      AutomationId, Xpath, Img, PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""

        try:
            # 准备控件
            control, x, y = prepare_control(WindowName, Name, ClassName, ControlType,
                                            foundIndex, AutomationId, Xpath, Img, PARAMETERS)
            if not getattr(control, "IsEnabled", True):
                raise Exception('控件不可点击')
            parameters = _parse_parameters(PARAMETERS)
            force_coordinates = _coerce_optional_bool(
                parameters.get("强制坐标", False), "强制坐标"
            )
            if force_coordinates or not _invoke_semantic_action(control):
                _activate_target_window(WindowName)
                control.Click(x, y)

            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 点击 成功！"

        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 点击 异常：{e}"

        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def mouse_left_press(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                            PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            control = _require_control(LOCATION)
            _activate_target_window(WindowName)
            x, y = get_pos(control, PARAMETERS)
            absolute_x, absolute_y = _absolute_control_position(control, x, y)
            uiautomation.PressMouse(absolute_x, absolute_y)
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标左键按下 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标左键按下 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def mouse_left_release(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                              PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            positioning_error = None
            try:
                control = _require_control(LOCATION)
                x, y = get_pos(control, PARAMETERS)
                _absolute_control_position(control, x, y)
            except Exception as exc:
                positioning_error = exc
            uiautomation.ReleaseMouse()
            if positioning_error:
                _MESSAGE_TYPE = _MESSAGE.WARNING
                MESSAGE = (
                    f"步骤：{ActionTitle} 控件【{Name}】 定位异常：{positioning_error}，"
                    "已在当前位置释放鼠标左键"
                )
            else:
                MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标左键释放 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标左键释放 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def right_click(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                       PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            control = _require_control(LOCATION)
            _activate_target_window(WindowName)
            x, y = get_pos(control, PARAMETERS)
            control.RightClick(x, y)
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 右键点击 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 右键点击 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def mouse_right_press(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                             PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            control = _require_control(LOCATION)
            _activate_target_window(WindowName)
            x, y = get_pos(control, PARAMETERS)
            absolute_x, absolute_y = _absolute_control_position(control, x, y)
            uiautomation.RightPressMouse(absolute_x, absolute_y)
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标右键按下 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标右键按下 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def mouse_right_release(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                              PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            positioning_error = None
            try:
                control = _require_control(LOCATION)
                x, y = get_pos(control, PARAMETERS)
                _absolute_control_position(control, x, y)
            except Exception as exc:
                positioning_error = exc
            uiautomation.RightReleaseMouse()
            if positioning_error:
                _MESSAGE_TYPE = _MESSAGE.WARNING
                MESSAGE = (
                    f"步骤：{ActionTitle} 控件【{Name}】 定位异常：{positioning_error}，"
                    "已在当前位置释放鼠标右键"
                )
            else:
                MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标右键释放 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标右键释放 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def centre_click(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                        PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            control = _require_control(LOCATION)
            _activate_target_window(WindowName)
            x, y = get_pos(control, PARAMETERS)
            control.MiddleClick(x, y)
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 中键点击 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 中键点击 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def double_click(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                        PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            control = _require_control(LOCATION)
            _activate_target_window(WindowName)
            x, y = get_pos(control, PARAMETERS)
            control.DoubleClick(x, y)

            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 双击 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 双击 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def mouse_move_pos(cls, ActionTitle, x, y):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            uiautomation.MoveTo(x, y)
            MESSAGE = f"步骤：{ActionTitle} 位置【{x}{y}】 执行动作 移动鼠标到坐标位置 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 位置【{x}{y}】 执行动作 移动鼠标到坐标位置 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def mouse_move_control(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                              PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            control = _require_control(LOCATION)
            _activate_target_window(WindowName)
            x, y = get_pos(control, PARAMETERS)
            control.MoveCursorToInnerPos(x, y)
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 移动鼠标到坐标位置 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 移动鼠标到坐标位置 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def drag_control_by_control(
        cls,
        current_control,
        target_control,
        current_offset=None,
        target_offset=None,
    ):
        current_coord = get_control_coordinates(current_control)
        target_coord = get_control_coordinates(target_control)
        current_coord_x, current_coord_y = cls._drag_point(current_coord, current_offset)
        target_coord_x, target_coord_y = cls._drag_point(target_coord, target_offset)
        uiautomation.DragDrop(current_coord_x, current_coord_y, target_coord_x, target_coord_y)

    @staticmethod
    def _drag_point(coordinates, offset):
        left, top, right, bottom = coordinates
        if offset and offset[0] != -1 and offset[1] != -1:
            return left + int(offset[0]), top + int(offset[1])
        return (left + right) // 2, (top + bottom) // 2

    @classmethod
    def drag_control(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                        PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            PARAMETERS = _parse_parameters(PARAMETERS)
            LOCATION = {'WindowName': WindowName,
                        'Name': Name,
                        'ClassName': ClassName,
                        'ControlType': ControlType,
                        'foundIndex': foundIndex,
                        'AutomationId': AutomationId,
                        'Xpath': Xpath,
                        'Img': Img,
                        'PARAMETERS': {}
                        }
            DestCarl_Locators = {
                'WindowName': PARAMETERS['目的控件父窗口名称'],
                'Name': PARAMETERS['目的控件Name'],
                'ClassName': PARAMETERS['目的控件ClassName'],
                'ControlType': PARAMETERS['目的控件ControlType'],
                'foundIndex': PARAMETERS['目的控件foundIndex'],
                'AutomationId': PARAMETERS['目的控件AutomationId'],
                'Xpath': PARAMETERS['目的控件Xpath'],
            }
            control = _require_control(LOCATION)
            DestCtrl = _require_control(DestCarl_Locators)
            _activate_target_window(WindowName)
            cls.drag_control_by_control(
                control,
                DestCtrl,
                current_offset=(PARAMETERS.get("x", -1), PARAMETERS.get("y", -1)),
                target_offset=(
                    PARAMETERS.get("目的控件x", -1),
                    PARAMETERS.get("目的控件y", -1),
                ),
            )

            MESSAGE = f"步骤：{ActionTitle} 执行动作 拖动控件 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 执行动作 拖动控件 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def set_text(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                    PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)

            control = _require_control(LOCATION)
            PARAMETERS = _parse_parameters(PARAMETERS)
            value = PARAMETERS["设置文本"]
            if not _set_control_value(control, value):
                _activate_target_window(WindowName)
                control.Click(waitTime=0)
                control.SendKeys("{Ctrl}a")
                control.SendKeys(value)
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 设置文本 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 设置文本 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def input_text(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                      PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)
            PARAMETERS = _parse_parameters(PARAMETERS)
            control = _resolve_location(LOCATION)
            if _set_control_value(control, PARAMETERS["输入文本"]):
                MESSAGE = f"步骤：{ActionTitle}执行动作 输入 成功"
                update_message_type(_MESSAGE_TYPE)
                return MESSAGE
            if control is not None:
                _activate_target_window(WindowName)
                control.Click(waitTime=0)
            previous_clipboard = None
            try:
                previous_clipboard = pyperclip.paste()
            except Exception:
                pass
            try:
                pyperclip.copy(PARAMETERS["输入文本"])
                pyautogui.hotkey("ctrl", "v", interval=0)
                time.sleep(0.05)
            finally:
                if previous_clipboard is not None:
                    pyperclip.copy(previous_clipboard)
            MESSAGE = f"步骤：{ActionTitle}执行动作 输入 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 输入 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def key_click(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                     PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)
            _focus_location(LOCATION)
            PARAMETERS = _parse_parameters(PARAMETERS)
            pyautogui.press(PARAMETERS["键盘按键"])
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘操作 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘操作 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def key_press(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                     PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)
            _focus_location(LOCATION)
            PARAMETERS = _parse_parameters(PARAMETERS)
            pyautogui.keyDown(PARAMETERS["键盘按键"])
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘按下 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘按下 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def key_release(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                       PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)
            _focus_location(LOCATION)
            PARAMETERS = _parse_parameters(PARAMETERS)
            pyautogui.keyUp(PARAMETERS["键盘按键"])
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘释放 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘释放 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def key_group(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                     PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
            LOCATION = cls._make_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                           PARAMETERS)
            key_mapping = {
                'ctrl_l': 'ctrlleft',
                'ctrl_r': 'ctrlright',
                'alt_l': 'altleft',
                'alt_r': 'altright',
                'shift_l': 'shiftleft',
                'shift_r': 'shiftright',
                'alt_gr': 'altright',
            }
            _focus_location(LOCATION)
            PARAMETERS = _parse_parameters(PARAMETERS)
            raw_keys = PARAMETERS["组合键"].lower().split("+")
            mapped_keys = [key_mapping.get(k.strip(), k.strip()) for k in raw_keys]
            pyautogui.hotkey(*mapped_keys)
            MESSAGE = f"步骤：{ActionTitle}执行动作 组合键 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 组合键 异常：{e}"
        update_message_type(_MESSAGE_TYPE)
        return MESSAGE

    @classmethod
    def scroll(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
               AutomationId, Xpath, Img, PARAMETERS):
        message_type = _MESSAGE.INFO
        try:
            parameters = _parse_parameters(PARAMETERS)
            location = cls._make_location(
                WindowName,
                Name,
                ClassName,
                ControlType,
                foundIndex,
                AutomationId,
                Xpath,
                Img,
                parameters,
            )
            if any((WindowName, Name, ClassName, ControlType, AutomationId, Xpath, Img)):
                control = _require_control(location)
                _activate_target_window(WindowName)
                if not control.MoveCursorToInnerPos(simulateMove=False):
                    raise ValueError("控件没有可用的滚动区域")
            auto_scroll(
                int(parameters.get("滚动距离", 1)),
                direction=parameters.get("滚动方向", "down"),
            )
            message = f"步骤：{ActionTitle} 执行动作 滚动 成功"
        except Exception as exc:
            message_type = _MESSAGE.ERROR
            message = f"步骤：{ActionTitle} 执行动作 滚动 异常：{exc}"
        update_message_type(message_type)
        return message

    @classmethod
    def toggle_control(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                       AutomationId, Xpath, Img, PARAMETERS):
        message_type = _MESSAGE.INFO
        try:
            parameters = _parse_parameters(PARAMETERS)
            location = cls._make_location(
                WindowName, Name, ClassName, ControlType, foundIndex,
                AutomationId, Xpath, Img, parameters
            )
            control = _require_control(location)
            pattern = _get_pattern(
                control, "GetTogglePattern", uiautomation.PatternId.TogglePattern
            )
            if pattern is None:
                raise ValueError("控件不支持 TogglePattern")
            desired = _coerce_optional_bool(parameters.get("选中"), "选中")
            if desired is None:
                if not pattern.Toggle():
                    raise RuntimeError("切换控件失败")
            else:
                target_state = 1 if desired else 0
                for _ in range(3):
                    if int(pattern.ToggleState) == target_state:
                        break
                    previous_state = int(pattern.ToggleState)
                    if not pattern.Toggle() or int(pattern.ToggleState) == previous_state:
                        raise RuntimeError("切换控件失败")
                else:
                    raise RuntimeError("切换控件未达到目标状态")
                if int(pattern.ToggleState) != target_state:
                    raise RuntimeError("切换控件未达到目标状态")
            message = f"步骤：{ActionTitle} 执行动作 切换状态 成功"
        except Exception as exc:
            message_type = _MESSAGE.ERROR
            message = f"步骤：{ActionTitle} 执行动作 切换状态 异常：{exc}"
        update_message_type(message_type)
        return message

    @classmethod
    def select_control(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                       AutomationId, Xpath, Img, PARAMETERS):
        message_type = _MESSAGE.INFO
        try:
            parameters = _parse_parameters(PARAMETERS)
            location = cls._make_location(
                WindowName, Name, ClassName, ControlType, foundIndex,
                AutomationId, Xpath, Img, parameters
            )
            control = _require_control(location)
            item_name = parameters.get("选择项")
            concrete = uiautomation.Control.CreateControlFromControl(control) or control
            if item_name is not None:
                success = _select_named_item(concrete, str(item_name))
            else:
                pattern = _get_pattern(
                    control,
                    "GetSelectionItemPattern",
                    uiautomation.PatternId.SelectionItemPattern,
                )
                success = pattern is not None and pattern.Select()
            if not success:
                raise RuntimeError("选择控件失败")
            message = f"步骤：{ActionTitle} 执行动作 选择 成功"
        except Exception as exc:
            message_type = _MESSAGE.ERROR
            message = f"步骤：{ActionTitle} 执行动作 选择 异常：{exc}"
        update_message_type(message_type)
        return message

    @classmethod
    def expand_collapse_control(
        cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
        AutomationId, Xpath, Img, PARAMETERS
    ):
        message_type = _MESSAGE.INFO
        try:
            parameters = _parse_parameters(PARAMETERS)
            location = cls._make_location(
                WindowName, Name, ClassName, ControlType, foundIndex,
                AutomationId, Xpath, Img, parameters
            )
            control = _require_control(location)
            pattern = _get_pattern(
                control,
                "GetExpandCollapsePattern",
                uiautomation.PatternId.ExpandCollapsePattern,
            )
            if pattern is None:
                raise ValueError("控件不支持 ExpandCollapsePattern")
            should_expand = _coerce_optional_bool(parameters.get("展开", True), "展开")
            if should_expand is None:
                should_expand = True
            target_state = 1 if should_expand else 0
            try:
                current_state = int(pattern.ExpandCollapseState)
            except Exception:
                current_state = None
            if current_state != target_state:
                success = pattern.Expand() if should_expand else pattern.Collapse()
                reached_target = bool(success) and _wait_for_expand_state(
                    control, target_state
                )
                if reached_target:
                    _set_action_mechanism("expand_collapse_pattern")
                if (
                    not reached_target
                    and _is_qt_control(control)
                    and _message_click_control(control, double=True)
                ):
                    reached_target = _wait_for_expand_state(
                        control, target_state, timeout=0.75
                    )
                    if reached_target:
                        _set_action_mechanism("qt_window_message")
                if not reached_target:
                    raise RuntimeError("展开/折叠控件未达到目标状态")
            else:
                _set_action_mechanism("already_in_target_state")
            message = f"步骤：{ActionTitle} 执行动作 展开折叠 成功"
        except Exception as exc:
            message_type = _MESSAGE.ERROR
            message = f"步骤：{ActionTitle} 执行动作 展开折叠 异常：{exc}"
        update_message_type(message_type)
        return message

    @classmethod
    def activate_window(cls, WindowTitle):
        return set_top_window(WindowTitle)


# ==============================鼠标动作==============================


def left_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                  AutomationId, Xpath, Img, PARAMETERS):
    """
    鼠标左键点击控件

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.left_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                                 AutomationId, Xpath, Img, PARAMETERS)


def get_pos(control, PARAMETERS: dict):
    """
    获取控件位置坐标

    :param
        control: 控件对象
        PARAMETERS (dict): 参数字典

    :return
        tuple: (x, y) 坐标值
    """
    PARAMETERS = _parse_parameters(PARAMETERS)
    correct_x, correct_y = correct_ctrl_position(control)
    show_ctrl_area(control)
    x, y = (int(PARAMETERS.get("x", correct_x)), int(PARAMETERS.get("y", correct_y)))
    if x == -1:
        x = correct_x
    if y == -1:
        y = correct_y
    return x, y


def mouse_left_press(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                        PARAMETERS):
    """
    鼠标左键按下

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.mouse_left_press(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                       PARAMETERS)


def mouse_left_release(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                          PARAMETERS):
    """
    鼠标左键释放

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.mouse_left_release(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                         PARAMETERS)


def right_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                   PARAMETERS):
    """
    鼠标右键点击

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.right_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                  PARAMETERS)


def mouse_right_press(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                         PARAMETERS):
    """
    鼠标右键按下

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.mouse_right_press(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                        PARAMETERS)


def mouse_right_release(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                           PARAMETERS):
    """
    鼠标右键释放

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.mouse_right_release(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                          PARAMETERS)


def centre_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                    PARAMETERS):
    """
    中键点击

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.centre_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                   PARAMETERS)


def double_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                    PARAMETERS):
    """
    鼠标左键双击

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.double_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                   PARAMETERS)


def mouse_move_pos(ActionTitle, x, y):
    """
    移动鼠标到坐标位置

    :param
        ActionTitle (str): 动作标题
        x (int): X坐标
        y (int): Y坐标

    :return
        str: 执行结果信息
    """
    return Controller.mouse_move_pos(ActionTitle, x, y)


def mouse_move_control(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                          PARAMETERS):
    """
    移动鼠标到控件位置

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.mouse_move_control(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                         PARAMETERS)


def drag_control_by_control(current_control, target_control):
    """
    通过控件进行拖拽

    :param current_control: 当前控件对象
    :param target_control: 目标控件对象
    """
    return Controller.drag_control_by_control(current_control, target_control)


def drag_control(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                    PARAMETERS):
    """
    通过控件拖拽

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典

    :return
        str: 执行结果信息
    """
    return Controller.drag_control(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                   PARAMETERS)


# ==============================键盘动作==============================
def set_text(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                PARAMETERS):
    """
    设置文本（直接设置控件文本内容）

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典，必须包含"设置文本"键

    :return
        str: 执行结果信息
    """
    return Controller.set_text(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                               PARAMETERS)


def input_text(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                  PARAMETERS):
    """
    输入文本（通过剪贴板粘贴方式输入文本）

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典，必须包含"输入文本"键

    :return
        str: 执行结果信息
    """
    return Controller.input_text(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                 PARAMETERS)


def key_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                 PARAMETERS):
    """
    键盘按键点击（按下并释放）

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典，必须包含"键盘按键"键

    :return
        str: 执行结果信息
    """
    return Controller.key_click(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                PARAMETERS)


def key_press(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                 PARAMETERS):
    """
    键盘按键按下

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典，必须包含"键盘按键"键

    :return
        str: 执行结果信息
    """
    return Controller.key_press(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                PARAMETERS)


def key_release(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                   PARAMETERS):
    """
    键盘按键释放

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典，必须包含"键盘按键"键

    :return
        str: 执行结果信息
    """
    return Controller.key_release(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                  PARAMETERS)


def key_group(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                 PARAMETERS):
    """
    组合键操作

    :param
        ActionTitle (str): 动作标题
        WindowName (str): 窗口名称
        Name (str): 控件名称
        ClassName (str): 控件类名
        ControlType (str): 控件类型
        foundIndex (int): 控件索引
        AutomationId (str): 自动化ID
        Xpath (str): XPath路径
        Img (str): 图像路径
        PARAMETERS (dict): 参数字典，必须包含"组合键"键

    :return
        str: 执行结果信息
    """
    return Controller.key_group(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                                PARAMETERS)


def activate_window(WindowTitle):
    """
    激活指定标题的窗口

    :param
        WindowTitle (str): 窗口标题
    """
    return Controller.activate_window(WindowTitle)


def scroll(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
           AutomationId, Xpath, Img, PARAMETERS):
    """滚动当前鼠标所在区域。"""
    return Controller.scroll(
        ActionTitle,
        WindowName,
        Name,
        ClassName,
        ControlType,
        foundIndex,
        AutomationId,
        Xpath,
        Img,
        PARAMETERS,
    )


def toggle_control(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                   AutomationId, Xpath, Img, PARAMETERS):
    return Controller.toggle_control(
        ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
        AutomationId, Xpath, Img, PARAMETERS
    )


def select_control(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                   AutomationId, Xpath, Img, PARAMETERS):
    return Controller.select_control(
        ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
        AutomationId, Xpath, Img, PARAMETERS
    )


def expand_collapse_control(ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                            AutomationId, Xpath, Img, PARAMETERS):
    return Controller.expand_collapse_control(
        ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
        AutomationId, Xpath, Img, PARAMETERS
    )


def generate_action_title(record_info):
    """
    生成动作标题

    :param
        record_info (dict): 录制的控件信息

    :return
        str: 动作标题
    """
    ACTION = record_info["ACTION"]
    Name = record_info["LOCATION"].get("Name", "")
    ControlType = record_info["LOCATION"].get("ControlType", "")
    PARAMETERS = record_info["LOCATION"]["PARAMETERS"]

    if '输入文本' == ACTION:
        ActionTitle = f'{ACTION}【{PARAMETERS["输入文本"]}】'
    elif '设置文本' == ACTION:
        if Name:
            ActionTitle = f'{ACTION}【{Name}{PARAMETERS["设置文本"]}】'
        else:
            ActionTitle = f'{ACTION}【{ControlType}{PARAMETERS["设置文本"]}】'
    elif '键盘' in ACTION:
        ActionTitle = f"键盘按键【{PARAMETERS['键盘按键']}】"
    elif '组合键' in ACTION:
        ActionTitle = f"组合键【{PARAMETERS['组合键']}】"
    elif Name:
        ActionTitle = f'{ACTION}【{Name}】'
    else:
        ActionTitle = f'{ACTION}【{ControlType}】'
    return ActionTitle


# @timeit
def run_action(record_info):
    """
    执行动作

    :param
        record_info (dict): 录制信息，包含ACTION和LOCATION字段

    :return
        str: 执行结果信息
    """
    _set_action_mechanism(None)
    if not isinstance(record_info, dict) or not isinstance(record_info.get("LOCATION"), dict):
        update_message_type(_MESSAGE.ERROR)
        return "动作数据格式错误"

    action = record_info.get("ACTION", "")
    func = Execute_Function.get(action)
    if func is None:
        update_message_type(_MESSAGE.ERROR)
        return f"不支持的动作：{action}"

    normalized = dict(record_info)
    normalized["LOCATION"] = dict(record_info["LOCATION"])
    normalized["LOCATION"].setdefault("PARAMETERS", {})
    try:
        action_title = generate_action_title(normalized)
        return func(action_title, **normalized["LOCATION"])
    except (KeyError, TypeError, ValueError) as exc:
        update_message_type(_MESSAGE.ERROR)
        return f"动作数据格式错误：{exc}"


Execute_Function: dict[str, Callable[..., str]] = {
    '点击': left_click,
    '按下鼠标左键': mouse_left_press,
    '释放鼠标左键': mouse_left_release,
    '右击': right_click,
    '按下鼠标右键': mouse_right_press,
    '释放鼠标右键': mouse_right_release,
    '中击': centre_click,
    '双击': double_click,
    '移动鼠标到坐标': mouse_move_pos,
    '移动鼠标到控件': mouse_move_control,
    '设置文本': set_text,
    '输入文本': input_text,
    '键盘点击': key_click,
    '键盘按下': key_press,
    '键盘释放': key_release,
    '拖拽': drag_control,
    '组合键': key_group,
    '滚动': scroll,
    '切换状态': toggle_control,
    '选择': select_control,
    '展开折叠': expand_collapse_control,
}
