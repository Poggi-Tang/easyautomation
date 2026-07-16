# -*- coding: utf-8 -*-
# @Name:            utils.py
# @Author:          tang
# @LastEditDate:    2025/9/30-9:44
# @depict:          通用字典、函数
import ast
import ctypes
import logging
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from functools import wraps
from pathlib import Path
from typing import Any

import pyautogui
import pygetwindow as gw
import uiautomation

_UIA_THREAD_STATE = threading.local()
LOGGER = logging.getLogger("easy_uiauto")
if not LOGGER.handlers:
    _handler = logging.StreamHandler(sys.stderr)
    _handler.setFormatter(logging.Formatter("%(message)s"))
    LOGGER.addHandler(_handler)
LOGGER.setLevel(logging.INFO)
LOGGER.propagate = False


def ensure_uiautomation_thread():
    """Initialize UIAutomation once for the current thread."""
    initializer = getattr(_UIA_THREAD_STATE, "initializer", None)
    if initializer is None:
        initializer = uiautomation.UIAutomationInitializerInThread()
        _UIA_THREAD_STATE.initializer = initializer
    return initializer

# 虚拟按键
MODIFIER_VK = {
    91, 92, 160, 161, 162, 163, 164, 165  # Win、Shift、Ctrl、Alt
}
VK_KEY_NAME = {
    # ===================数字=====================#
    49: '1',
    50: '2',
    51: '3',
    52: '4',
    53: '5',
    54: '6',
    55: '7',
    56: '8',
    57: '9',
    48: '0',
    # ===================字母=====================#
    65: 'a',
    66: 'b',
    67: 'c',
    68: 'd',
    69: 'e',
    70: 'f',
    71: 'g',
    72: 'h',
    73: 'i',
    74: 'j',
    75: 'k',
    76: 'l',
    77: 'm',
    78: 'n',
    79: 'o',
    80: 'p',
    81: 'q',
    82: 'r',
    83: 's',
    84: 't',
    85: 'u',
    86: 'v',
    87: 'w',
    88: 'x',
    89: 'y',
    90: 'z',
    # ===================一级符号=====================#
    192: '`',
    189: '-',
    187: '=',
    219: '[',
    221: ']',
    220: '\\',
    186: ';',
    222: "'",
    188: ',',
    190: '.',
    191: '/',
    # ===================二级符号（此处只做查看，在组合键中体现）=====================#
    # 192:'~',
    # 49:'!',
    # 50:'@',
    # 51:'#',
    # 52:'$',
    # 53:'%',
    # 54:'^',
    # 55:'&',
    # 56:'*',
    # 57:'(',
    # 48:')',
    # 189:'_',
    # 187:'+',
    # 219:'{',
    # 221:'}',
    # 220:'|',
    # 186:':',
    # 222:'"',
    # 188:'<',
    # 190:'>',
    # 191:'?',
    # ==================特殊按键=====================#
    # 44: 'push_message_screen',
    36: 'home',
    35: 'end',
    45: 'insert',
    46: 'delete',
    8: 'backspace',
    13: 'enter',
    9: 'tab',
    20: 'caps_lock',
    160: 'shift',
    161: 'shift_r',
    162: 'ctrl_l',
    163: 'ctrl_r',
    164: 'alt_l',
    165: 'alt_gr',
    112: 'f1',
    113: 'f2',
    114: 'f3',
    115: 'f4',
    116: 'f5',
    117: 'f6',
    118: 'f7',
    119: 'f8',
    120: 'f9',
    121: 'f10',
    122: 'f11',
    123: 'f12',
    255: 'Fn',
    91: 'cmd',
    38: 'up',
    40: 'down',
    37: 'left',
    39: 'right',
    33: 'page_up',
    34: 'page_down',
    # ===================小键盘数字=====================#
    97: '1',
    98: '2',
    99: '3',
    100: '4',
    101: '5',
    102: '6',
    103: '7',
    104: '8',
    105: '9',
    96: '0',
    # ===================小键盘符号=====================#
    111: '/',
    106: '*',
    109: '-',
    107: '+',
    110: '.',

}
INPUT_VK_KEY = [  # 可输入的虚拟键码列表
    48, 49, 50, 51, 52, 53, 54, 55, 56, 57,
    65, 66, 67, 68, 69, 70, 71, 72, 73, 74,
    75, 76, 77, 78, 79, 80, 81, 82, 83, 84,
    85, 86, 87, 88, 89, 90, 96, 97, 98, 99,
    100, 101, 102, 103, 104, 105, 106, 107,
    109, 110, 111, 186, 187, 188, 189, 190,
    191, 192, 219, 220, 221, 222]
# 当前软件名称
CURRENT_APP_NAME = None
CONTROL_TYPE_IDS = {  # 控件类型名称到ID的映射字典
    'AppBarControl': 50040,
    'ButtonControl': 50000,
    'CalendarControl': 50001,
    'CheckBoxControl': 50002,
    'ComboBoxControl': 50003,
    'CustomControl': 50025,
    'DataGridControl': 50028,
    'DataItemControl': 50029,
    'DocumentControl': 50030,
    'EditControl': 50004,
    'GroupControl': 50026,
    'HeaderControl': 50034,
    'HeaderItemControl': 50035,
    'HyperlinkControl': 50005,
    'ImageControl': 50006,
    'ListControl': 50008,
    'ListItemControl': 50007,
    'MenuBarControl': 50010,
    'MenuControl': 50009,
    'MenuItemControl': 50011,
    'PaneControl': 50033,
    'ProgressBarControl': 50012,
    'RadioButtonControl': 50013,
    'ScrollBarControl': 50014,
    'SemanticZoomControl': 50039,
    'SeparatorControl': 50038,
    'SliderControl': 50015,
    'SpinnerControl': 50016,
    'SplitButtonControl': 50031,
    'StatusBarControl': 50017,
    'TabControl': 50018,
    'TabItemControl': 50019,
    'TableControl': 50036,
    'TextControl': 50020,
    'ThumbControl': 50027,
    'TitleBarControl': 50037,
    'ToolBarControl': 50021,
    'ToolTipControl': 50022,
    'TreeControl': 50023,
    'TreeItemControl': 50024,
    'WindowControl': 50032
}
ITEM_TYPE_NAMES = [  # 项目类型控件名称列表
    # 表格类
    'HeaderControl',
    'DataItemControl',
    'HeaderItemControl',
    # 列表类
    'ListItemControl',
    # 树类
    'TreeItemControl',
    # 标签类
    'TabItemControl',
]
CONTAINER_TYPE_NAMES = [  # 容器类型控件名称列表
    'TableControl',
    'ListControl',
    'TreeControl',
    'TabControl',
]
# 控件缓存
CONTROL_CACHE: dict[str, object] = {}  # 控件缓存字典，存储XPath到控件的映射
CACHE_METADATA: dict[str, dict[str, object]] = {}  # 缓存元数据字典
CACHE_QUEUE: list[tuple[str, object]] = []  # 缓存队列，用于异步处理
CACHE_PROCESSING = False  # 缓存处理状态标志
CONTROL_CACHE_TIMEOUT = 5.0  # 控件缓存超时时间（秒）
CACHE_LOCK = threading.RLock()
# 编译缓存
APP: list[str] = []  # 应用程序列表
COMPILE_COUNT = 0  # 编译计数器
COMPILE_PROGRESS = 120  # 编译进度基准值


def timeit(func):
    """测量函数执行时间"""

    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()  # 高精度计时开始
        result = func(*args, **kwargs)  # 执行目标函数
        elapsed = time.perf_counter() - start_time  # 计算耗时
        LOGGER.debug("[计时] %s 执行耗时: %.3f 毫秒", func.__name__, elapsed * 1000)
        return result  # 返回原函数结果

    return wrapper


def push_message(log, log_path=''):
    """
    此函数用于替代print()

    :param log: 日志信息
    :param log_path: 保存文件路径
    :return:
    """
    message = str(log)
    LOGGER.info(message)
    if log_path:
        with open(log_path, 'a', encoding='utf-8') as file:
            file.write(message + "\n")
    return


def push_message_run_time(func, param, count=1):
    """打印函数执行时间"""
    push_message("=" * 80)
    control = None
    all_start = int(datetime.now().timestamp() * 1000)
    for i in range(count):
        start = int(datetime.now().timestamp() * 1000)
        control = func(param)
        if i == 0:
            end = int(datetime.now().timestamp() * 1000)
            push_message(f"首次耗时：{end - start}ms")
        if i == count - 1:
            end = int(datetime.now().timestamp() * 1000)
            push_message(f"最后一次耗时：{end - start}ms")
    push_message(control)
    # push_message(f'控件类型:{control.GetTopLevelControl().FrameworkId}')
    # push_message("期望控件：", param[-1]['ControlType'], param[-1].get('Name', ''), param[-1].get('ClassName', ''))
    # push_message("实际控件：", control.ControlTypeName, control.Name, control.ClassName)
    all_run_time = int(datetime.now().timestamp() * 1000) - all_start
    push_message(f"总耗时：{all_run_time}ms 平均耗时{all_run_time / count}ms")


def get_window_title_by_handle(hwnd):
    """通过窗口句柄获取真实标题"""
    length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return ""

    buffer = ctypes.create_unicode_buffer(length + 1)
    ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


@timeit
def set_top_window(window_title):
    """
    激活窗口
    :param window_title:
    :return: 激活结果 通过True 失败False
    """
    if not window_title:
        return False
    try:
        ensure_uiautomation_thread()
        window = uiautomation.WindowControl(Name=window_title, searchDepth=1)
        if not window.Exists(0):
            return False
        handle = int(window.NativeWindowHandle or 0)
        user32 = ctypes.windll.user32
        if handle:
            if user32.IsIconic(handle):
                user32.ShowWindow(handle, 9)
            try:
                window.SetActive()
            except Exception:
                user32.SetForegroundWindow(handle)
            if int(user32.GetForegroundWindow() or 0) == handle:
                return True

        real_title = get_window_title_by_handle(handle) or window_title
        for candidate in gw.getWindowsWithTitle(real_title):
            if candidate.title != real_title:
                continue
            candidate_handle = int(getattr(candidate, "_hWnd", 0) or 0)
            if handle and candidate_handle != handle:
                continue
            try:
                candidate.activate()
            except Exception:
                continue
            if int(user32.GetForegroundWindow() or 0) == (handle or candidate_handle):
                return True
        return False
    except Exception:
        return False


def get_control_info(x, y):
    """
    获取坐标点控件Xpath、控件
    :param x: 横坐标
    :param y: 纵坐标
    :return: 控件Xpath、控件
    """
    ensure_uiautomation_thread()
    try:
        control = uiautomation.ControlFromPoint(x, y)
        if control is None:
            return [], None
        child = control.GetFirstChildControl()
        while child:
            # print(child)
            if child.Exists(0):
                boundingRect = child.BoundingRectangle
                if boundingRect.left <= x <= boundingRect.right and boundingRect.top <= y <= boundingRect.bottom:
                    control = child
                    break
                child = child.GetNextSiblingControl()
            else:
                break
        xpath = get_control_xpath(control, x, y)
        return xpath, control

    except Exception as e:
        push_message(f"获取控件信息异常: {e}")
        return [], None


def get_focused_control_info():
    """Return the focused control's XPath and Control in the current thread."""
    ensure_uiautomation_thread()
    try:
        control = uiautomation.GetFocusedControl()
        if control is None:
            return [], None
        return get_control_xpath(control), control
    except Exception as exc:
        push_message(f"获取焦点控件异常: {exc}")
        return [], None


def auto_scroll(pixels, duration=0, direction='down'):
    """
    滚动鼠标
    :param pixels: 滚动距离(像素)
    :param duration: 滚动持续时间(秒)
    :param direction: 滚动方向('up'、'down'、'left'或'right')
    """
    normalized = direction.lower()
    if normalized not in {'up', 'down', 'left', 'right'}:
        raise ValueError("方向参数必须是'up'、'down'、'left'或'right'")
    amount = abs(int(pixels))
    if normalized in {'up', 'down'}:
        pyautogui.scroll(amount if normalized == 'up' else -amount, _pause=False)
    else:
        pyautogui.hscroll(amount if normalized == 'right' else -amount, _pause=False)
    if duration > 0:
        time.sleep(duration)


# #######################编译#############################
def set_global_control_cache_timeout(seconds):
    """设置全局控件缓存超时时间"""
    global CONTROL_CACHE_TIMEOUT
    seconds = float(seconds)
    if seconds < 0:
        raise ValueError("缓存超时时间不能为负数")
    CONTROL_CACHE_TIMEOUT = seconds


def get_control_cache_stats():
    """Return JSON-safe cache statistics without exposing cached COM objects."""
    with CACHE_LOCK:
        thread_counts: dict[str, int] = {}
        for metadata in CACHE_METADATA.values():
            thread_id = str(metadata.get("thread_id", "unknown"))
            thread_counts[thread_id] = thread_counts.get(thread_id, 0) + 1
        return {
            "entries": len(CONTROL_CACHE),
            "metadata_entries": len(CACHE_METADATA),
            "queue_entries": len(CACHE_QUEUE),
            "timeout_seconds": CONTROL_CACHE_TIMEOUT,
            "entries_by_thread": thread_counts,
        }


def clear_control_cache():
    """Release all cached control references and pending cache work."""
    with CACHE_LOCK:
        removed = len(CONTROL_CACHE)
        CONTROL_CACHE.clear()
        CACHE_METADATA.clear()
        CACHE_QUEUE.clear()
    return removed


# @timeit
def compile_controls(control=None, max_depth=15, compile_log=False):
    """
    编译控件
    :param control: 控件
    :param max_depth: 最大遍历深度
    :param compile_log: 编译日志
    """
    if max_depth < 1:
        raise ValueError("max_depth 必须大于 0")
    push_message(f"\n编译开始预计耗时：{max_depth * 0.8}秒")
    ensure_uiautomation_thread()
    global COMPILE_COUNT, COMPILE_PROGRESS
    COMPILE_COUNT = 0
    COMPILE_PROGRESS = max(1, 120 * max_depth)
    APP.clear()
    if control is None:
        for app in uiautomation.GetRootControl().GetChildren():
            _recursive_cache_controls(app, [], max_depth, 1, compile_log)
    else:
        parent = control.GetParentControl()
        parent_xpath = get_control_xpath(parent) if parent else []
        _recursive_cache_controls(control, parent_xpath, max_depth, 1, compile_log)
    if compile_log:
        push_message("编译进度: 100%")
    push_message(f"\n编译完成，已添加 {len(CONTROL_CACHE)} 个控件")


def generate_cache_keys(xpath):
    """生成带索引和不带索引的稳定 XPath 缓存键。"""
    return _build_cache_key(xpath, True), _build_cache_key(xpath, False)


def _build_cache_key(xpath, include_found_index):
    parts = []
    for item in xpath or []:
        control_type = item.get("ControlType", "")
        if isinstance(control_type, int):
            control_type = uiautomation.ControlTypeNames.get(control_type, str(control_type))
        fields = [
            ("ControlType", control_type),
            ("Name", item.get("Name", "")),
            ("ClassName", item.get("ClassName", "")),
            ("AutomationId", item.get("AutomationId", "")),
            ("searchDepth", item.get("searchDepth", 1)),
        ]
        if include_found_index:
            fields.append(("foundIndex", item.get("foundIndex", 1)))
        encoded = ",".join(f"{name}:{value!r}" for name, value in fields)
        parts.append(encoded)
    return "|".join(parts)


def _store_cached_control(cache_key, control):
    with CACHE_LOCK:
        CONTROL_CACHE[cache_key] = control
        CACHE_METADATA[cache_key] = {
            "last_verified": time.monotonic(),
            "window_handle": getattr(control, "NativeWindowHandle", 0),
            "thread_id": threading.get_ident(),
            "signature": _control_signature(control),
        }


def _get_cached_control(cache_key):
    with CACHE_LOCK:
        control = CONTROL_CACHE.get(cache_key)
        metadata = CACHE_METADATA.get(cache_key)
        if control is None or metadata is None:
            return None
        if metadata.get("thread_id") != threading.get_ident():
            return None
        last_verified = metadata.get("last_verified", 0.0)
        if not isinstance(last_verified, (int, float)):
            return None
        age = time.monotonic() - last_verified
        if CONTROL_CACHE_TIMEOUT and age > CONTROL_CACHE_TIMEOUT:
            CONTROL_CACHE.pop(cache_key, None)
            CACHE_METADATA.pop(cache_key, None)
            return None
    try:
        if hasattr(control, "Exists") and not control.Exists(0):
            raise LookupError("cached control no longer exists")
        if metadata.get("signature") != _control_signature(control):
            raise LookupError("cached control identity changed")
        get_control_coordinates(control)
    except Exception:
        with CACHE_LOCK:
            CONTROL_CACHE.pop(cache_key, None)
            CACHE_METADATA.pop(cache_key, None)
        return None
    with CACHE_LOCK:
        CACHE_METADATA[cache_key]["last_verified"] = time.monotonic()
    return control


def _control_signature(control):
    return tuple(
        getattr(control, attr, "")
        for attr in ("ControlTypeName", "Name", "ClassName", "AutomationId", "NativeWindowHandle")
    )


def _recursive_cache_controls(ctrl, parent_xpath, max_depth, current_depth, compile_log=False):
    """递归缓存控件及其子控件"""
    global COMPILE_COUNT, COMPILE_PROGRESS
    progress = min(100, (COMPILE_COUNT / COMPILE_PROGRESS) * 100)
    if compile_log:
        push_message(f"编译进度: {int(progress)}%")
    if current_depth > max_depth:
        return

    try:
        app = ctrl.GetTopLevelControl()
        if app:
            app_name = app.Name
            if app_name not in APP:
                APP.append(app_name)
                current_depth = 1
        else:
            app_name = '桌面'
            current_depth = 2
        current_info = {
            "ControlType": ctrl.ControlTypeName,
            "searchDepth": current_depth,
        }
        # 构建当前控件的xpath信息

        if ctrl.Name:
            current_info["Name"] = ctrl.Name
        if ctrl.ClassName:
            current_info["ClassName"] = ctrl.ClassName
        if ctrl.AutomationId:
            current_info["AutomationId"] = ctrl.AutomationId

        if ctrl.ControlTypeName in ITEM_TYPE_NAMES:
            foundIndex = 0
            for child in ctrl.GetParentControl().GetChildren():
                if child.ControlType == ctrl.ControlType:
                    foundIndex += 1
                if child.Name == ctrl.Name:
                    # 查看是否已经缓存过
                    current_info['foundIndex'] = foundIndex
                    ctrl_xpath = parent_xpath + [current_info]
                    key_with, key_without = generate_cache_keys(ctrl_xpath)
                    if key_with not in CONTROL_CACHE:
                        break

        ctrl_xpath = parent_xpath + [current_info]
        key_with, key_without = generate_cache_keys(ctrl_xpath)
        if compile_log:
            push_message(f"编译控件：{current_info}")
        # 同时缓存两个key指向同一个控件
        _store_cached_control(key_with, ctrl)
        _store_cached_control(key_without, ctrl)
        COMPILE_COUNT += 1
        if current_depth < max_depth:
            try:
                children = ctrl.GetChildren()
                for child in children:
                    _recursive_cache_controls(child, ctrl_xpath, max_depth, current_depth + 1, compile_log)
            except Exception:
                # 获取子控件失败时不中断流程
                pass

    except Exception as e:
        # 单个控件缓存失败不影响整体
        push_message(f"编译控件失败：{e}")
        pass


# #########################Xpath定位#####################################
def _cache_sibling_controls(parent_ctrl, parent_xpath):
    """缓存兄弟控件"""
    try:
        children = parent_ctrl.GetChildren()
        for idx, child in enumerate(children):
            child_xpath = parent_xpath + [{
                'ControlType': child.ControlTypeName,
                'foundIndex': idx,
                'ClassName': getattr(child, 'ClassName', ''),
                'AutomationId': getattr(child, 'AutomationId', ''),
                'Name': getattr(child, 'Name', '')
            }]
            cache_key = generate_cache_key(child_xpath)
            _async_cache_control(cache_key, child)
    except Exception:
        pass


def _async_cache_control(cache_key, control):
    """异步添加控件到缓存"""
    _store_cached_control(cache_key, control)


def _process_cache_queue():
    """处理缓存队列"""
    global CACHE_PROCESSING
    try:
        while True:
            with CACHE_LOCK:
                if not CACHE_QUEUE:
                    break
                cache_key, control = CACHE_QUEUE.pop(0)
            if _get_cached_control(cache_key) is None:
                _store_cached_control(cache_key, control)
            time.sleep(0.001)
    finally:
        with CACHE_LOCK:
            CACHE_PROCESSING = False


def generate_cache_key(xpath):
    """生成XPath路径的缓存键"""
    return _build_cache_key(xpath, True)

def find_control_by_xpath(xpath, debug=False, use_cache=True):
    """通过完整 XPath 顺序定位控件，并使用带过期校验的缓存。"""
    ensure_uiautomation_thread()
    xpath = _parse_xpath(xpath)
    if not xpath:
        return None
    key_with, key_without = generate_cache_keys(xpath)
    cached = _get_cached_control(key_with) if use_cache else None
    if (
        use_cache
        and cached is None
        and all(int(item.get("foundIndex", 1) or 1) <= 1 for item in xpath)
    ):
        cached = _get_cached_control(key_without)
    if cached is not None:
        if any("foundIndex" in item for item in xpath):
            fresh = _find_control_from_root(xpath, debug=debug)
            try:
                same = fresh is not None and uiautomation.ControlsAreSame(cached, fresh)
            except Exception:
                same = cached is fresh
            if not same:
                with CACHE_LOCK:
                    CONTROL_CACHE.pop(key_with, None)
                    CACHE_METADATA.pop(key_with, None)
                    CONTROL_CACHE.pop(key_without, None)
                    CACHE_METADATA.pop(key_without, None)
                cached = None
        if cached is None:
            return find_control_by_xpath(xpath, debug=debug, use_cache=False)
        if debug:
            push_message(f"使用缓存: {cached}")
        return cached

    control = _find_control_from_root(xpath, debug=debug)
    if use_cache and control is not None:
        _store_cached_control(key_with, control)
        if all(int(item.get("foundIndex", 1) or 1) <= 1 for item in xpath):
            _store_cached_control(key_without, control)
    return control


def _find_control_from_root(xpath, debug=False):
    current: Any = uiautomation.GetRootControl()
    try:
        for index, item in enumerate(xpath):
            if not isinstance(item, dict):
                raise ValueError("XPath 节点必须是字典")
            control_type_name = item.get("ControlType")
            if control_type_name not in CONTROL_TYPE_IDS:
                raise ValueError(f"未知控件类型: {control_type_name}")
            conditions: dict[str, Any] = {
                "ControlType": CONTROL_TYPE_IDS[control_type_name],
                "foundIndex": max(1, int(item.get("foundIndex", 1) or 1)),
            }
            depth = max(1, int(item.get("searchDepth", 1) or 1))
            conditions["searchDepth"] = depth
            conditions["Depth"] = depth
            for field in ("Name", "ClassName", "AutomationId"):
                value = item.get(field)
                if value not in (None, ""):
                    conditions[field] = value
            current = current.Control(**conditions)
            if not current.Refind(maxSearchSeconds=0, raiseException=False):
                return None
            if debug:
                push_message(f"第【{index}】层: {current}")
        return current
    except (AttributeError, LookupError, TypeError, ValueError) as exc:
        if debug:
            push_message(f"XPath 定位失败: {exc}")
        return None


def get_control_xpath(control, x=None, y=None):
    """
    获取控件Xpath
    :param x:横坐标
    :param y:纵坐标
    :param control: 当前控件对象
    :return: Xpath
    """

    if control is None:
        return []

    raw_path = []
    current = control
    seen = set()
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        raw_path.append(current)
        if _safe_control_attr(current, "ControlTypeName") == "DesktopControl":
            break
        try:
            current = current.GetParentControl()
        except Exception:
            break
    raw_path.reverse()

    xpath = []
    depth_since_last = 0
    for current in raw_path:
        control_type = _safe_control_attr(current, "ControlTypeName")
        class_name = _safe_control_attr(current, "ClassName")
        if control_type == "DesktopControl" or class_name == "#32769":
            continue
        depth_since_last += 1
        if control_type in {"CustomControl", "GroupControl"} or class_name in {
            "QWidget",
            "GroupControl",
            "CustomControl",
            "WindowControl",
        }:
            continue

        info = {"ControlType": control_type, "searchDepth": depth_since_last}
        for attr, key in (
            ("Name", "Name"),
            ("ClassName", "ClassName"),
            ("AutomationId", "AutomationId"),
        ):
            value = _safe_control_attr(current, attr)
            if value:
                info[key] = value
        found_index = _control_found_index(current, x=x, y=y)
        if found_index > 1 or control_type in ITEM_TYPE_NAMES:
            info["foundIndex"] = found_index
        xpath.append(info)
        depth_since_last = 0
    return xpath


def _safe_control_attr(control, attribute, default=""):
    try:
        return getattr(control, attribute, default)
    except Exception:
        return default


def _control_found_index(control, x=None, y=None):
    try:
        parent = control.GetParentControl()
        siblings = parent.GetChildren() if parent else []
    except Exception:
        return 1
    matched_index = 0
    identity_attributes = ["ControlTypeName"]
    identity_attributes.extend(
        attr
        for attr in ("Name", "ClassName", "AutomationId")
        if _safe_control_attr(control, attr) not in (None, "")
    )
    target_identity = {
        attr: _safe_control_attr(control, attr) for attr in identity_attributes
    }
    for sibling in siblings:
        sibling_identity = {
            attr: _safe_control_attr(sibling, attr, default=None)
            for attr in identity_attributes
        }
        if any(sibling_identity[attr] != value for attr, value in target_identity.items()):
            continue
        matched_index += 1
        try:
            if uiautomation.ControlsAreSame(sibling, control):
                return matched_index
        except Exception:
            if sibling is control:
                return matched_index
        if x is not None and y is not None:
            try:
                rect = sibling.BoundingRectangle
                if rect.left <= x <= rect.right and rect.top <= y <= rect.bottom:
                    return matched_index
            except Exception:
                continue
    return max(1, matched_index)


def redirect_control(name, ctrl):
    if ctrl:
        Children = []
        Repeat = []
        if ctrl.ControlTypeName in ITEM_TYPE_NAMES:
            ctrl = ctrl.GetParentControl()
            Children = ctrl.GetChildren()
        elif ctrl.ControlTypeName in CONTAINER_TYPE_NAMES:
            Children = ctrl.GetChildren()
        # 添加所有重名控件
        for child_ctrl in Children:
            if child_ctrl.Name == name:
                Repeat.append(child_ctrl)
        # 检测多个重名就返回源控件（根据index定位得出）
        if len(Repeat) > 1:
            # return ctrl
            return Repeat[-1]
        # 检测一个重名就返回此控件
        elif len(Repeat) == 1:
            # push_message(f'策略Xpath定位:定位成功')
            ctrl = Repeat[0]
    return ctrl


# ###################禁用uiautomation日志#################
# def log_decorator(func):
#     """日志装饰器，用于包装原始日志方法"""
#     pass
# auto.Logger.Write = log_decorator(auto.Logger.Write)
# ########################################################
def strategy_xpath(name, class_name, control_type, Xpath, debug=False, automation_id=""):
    """使用 XPath 定位；虚拟化列表节点会进行有限次数滚动重试。"""
    if not Xpath:
        return None
    try:
        control = find_control_by_xpath(Xpath, debug)
        if _control_matches(control, name, class_name, control_type, automation_id):
            return control
        if control_type not in ITEM_TYPE_NAMES or len(Xpath) < 2:
            return None

        container = find_control_by_xpath(Xpath[:-1], debug)
        if container is None:
            return None
        try:
            scroll_pattern = container.GetScrollPattern()
            if scroll_pattern is None:
                return None
            original_scroll = (
                scroll_pattern.HorizontalScrollPercent,
                scroll_pattern.VerticalScrollPercent,
            )
        except Exception:
            return None
        found = False
        try:
            for direction in ("up", "down"):
                scroll_pattern.SetScrollPercent(*original_scroll)
                unchanged_count = 0
                previous_signature = None
                amount = (
                    uiautomation.ScrollAmount.SmallDecrement
                    if direction == "up"
                    else uiautomation.ScrollAmount.SmallIncrement
                )
                for _ in range(20):
                    if not scroll_pattern.Scroll(
                        uiautomation.ScrollAmount.NoAmount, amount, waitTime=0
                    ):
                        break
                    control = _find_control_from_root(Xpath, debug=debug)
                    if _control_matches(
                        control, name, class_name, control_type, automation_id
                    ):
                        key_with, _ = generate_cache_keys(Xpath)
                        _store_cached_control(key_with, control)
                        found = True
                        return control
                    children = container.GetChildren()
                    signature = tuple(
                        (
                            getattr(child, "ControlTypeName", ""),
                            getattr(child, "Name", ""),
                            getattr(child, "AutomationId", ""),
                        )
                        for child in children[:20]
                    )
                    unchanged_count = (
                        unchanged_count + 1 if signature == previous_signature else 0
                    )
                    previous_signature = signature
                    if unchanged_count >= 3:
                        break
        finally:
            if not found:
                try:
                    scroll_pattern.SetScrollPercent(*original_scroll)
                except Exception:
                    pass
        return None
    except Exception as exc:
        if debug:
            push_message(f"策略 XPath 错误: {exc}")
        return None


def _control_matches(control, name, class_name, control_type, automation_id=""):
    if control is None:
        return False
    expected = {
        "ControlTypeName": control_type,
        "Name": name,
        "ClassName": class_name,
        "AutomationId": automation_id,
    }
    return all(not value or getattr(control, attr, "") == value for attr, value in expected.items())


@timeit
def find_control(LOCATION, debug=False):
    """
    根据配置列表获取控件
    此版本支持Name、ClassName、Type、Xpath、foundIndex、AutomationId
    :param
         debug: 是否打印调试日志
         LOCATION: 定位器字典，包含：
            - WindowName: 父窗口名称
            - Name: 控件名称
            - ClassName: 控件类名
            - Type: 控件类型
            - Xpath: 层级路径
            - foundIndex: 索引位置
    :return: 控件对象或None
    """
    ensure_uiautomation_thread()
    if not isinstance(LOCATION, dict):
        raise TypeError("LOCATION 必须是字典")

    global CURRENT_APP_NAME
    window_name = LOCATION.get("WindowName", "") or ""
    name = LOCATION.get("Name", "") or ""
    class_name = LOCATION.get("ClassName", "") or ""
    control_type = LOCATION.get("ControlType", "") or ""
    automation_id = LOCATION.get("AutomationId", "") or ""
    xpath = _parse_xpath(LOCATION.get("Xpath", []))
    if xpath is None:
        return None

    if debug:
        push_message(
            f"定位控件: window={window_name!r}, name={name!r}, "
            f"class={class_name!r}, type={control_type!r}"
        )

    first_node = xpath[0] if xpath else {}
    is_desktop = first_node.get("ClassName") == "WorkerW"
    is_taskbar = (
        window_name == "任务栏"
        or first_node.get("Name") == "任务栏"
        or first_node.get("ClassName") == "Shell_TrayWnd"
    )
    if window_name and not is_desktop and not is_taskbar:
        CURRENT_APP_NAME = window_name

    if xpath:
        control = strategy_xpath(
            name, class_name, control_type, xpath, debug, automation_id=automation_id
        )
        if control is not None:
            return control

    parent = None
    if window_name:
        for parent_type in ("WindowControl", "PaneControl"):
            parent = strategy_dictionary({"ControlType": parent_type, "Name": window_name})
            if parent is not None:
                break

    target_requested = any((name, class_name, control_type, automation_id))
    if target_requested and control_type:
        dictionary = {
            "ControlType": control_type,
            "Name": name,
            "ClassName": class_name,
            "AutomationId": automation_id,
            "searchDepth": LOCATION.get("searchDepth", 0xFFFFFFFF),
            "foundIndex": LOCATION.get("foundIndex", 1),
        }
        if parent is not None or not window_name:
            control = strategy_dictionary(dictionary, parent)
            if control is not None:
                return control
    elif parent is not None and not target_requested and not LOCATION.get("Img"):
        return parent

    image_control = _find_image_control(LOCATION)
    if image_control is not None:
        return image_control
    if debug:
        push_message("所有定位策略均失败，未找到控件")
    return None


def _parse_xpath(value):
    if value in (None, ""):
        return []
    if isinstance(value, list):
        parsed = value
    elif not isinstance(value, str):
        return None
    else:
        try:
            parsed = ast.literal_eval(value)
        except (SyntaxError, ValueError):
            return None
    if not isinstance(parsed, list):
        return None
    for item in parsed:
        if not isinstance(item, dict):
            return None
        if item.get("ControlType") not in CONTROL_TYPE_IDS:
            return None
        try:
            if int(item.get("searchDepth", 1) or 1) < 1:
                return None
            if int(item.get("foundIndex", 1) or 1) < 1:
                return None
        except (TypeError, ValueError):
            return None
    return parsed


def strategy_dictionary(dictionary, parent=None, timeout=0):
    """
    uiautomation.Control重构
    :param timeout: 查找超时时间
    :param Prant: 父控件
    :param dictionary[Dict]
        dictionary['ControlType']   控件类型名称
        dictionary['ClassName']     控件类名
        dictionary['AutomationId']  控件自动化ID
        dictionary['Name']          控件名称
        dictionary['searchDepth']   搜索层级
        dictionary['foundIndex']    索引
    :return:
    """
    try:
        ensure_uiautomation_thread()
        control_type = dictionary.get("ControlType")
        if control_type not in CONTROL_TYPE_IDS:
            return None
        parent = parent or uiautomation.GetRootControl()
        conditions = {
            "ControlType": CONTROL_TYPE_IDS[control_type],
            "searchDepth": max(1, int(dictionary.get("searchDepth", 0xFFFFFFFF) or 1)),
            "foundIndex": max(1, int(dictionary.get("foundIndex", 1) or 1)),
        }
        for field in ("Name", "ClassName", "AutomationId"):
            value = dictionary.get(field)
            if value not in (None, ""):
                conditions[field] = value
        control = parent.Control(**conditions)
        state = control.Refind(maxSearchSeconds=max(0, float(timeout)), raiseException=False)
        if state:
            return control
        return None
    except (AttributeError, KeyError, TypeError, ValueError):
        return None


@dataclass
class RectStruct:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return max(0, self.right - self.left)

    @property
    def height(self) -> int:
        return max(0, self.bottom - self.top)

    @property
    def check_pos(self) -> tuple[int, int]:
        x = self.left + (self.right - self.left) // 2
        y = self.top + (self.bottom - self.top) // 2
        return x, y

    def is_valid(self) -> bool:
        return self.width > 0 and self.height > 0


class _ImageParentControl:
    ControlTypeName = "PaneControl"
    Name = "Screen"
    ClassName = "Screen"
    AutomationId = ""

    @property
    def BoundingRectangle(self):
        return get_system_rect()[0]

    def GetParentControl(self):
        return None


class ImageControl:
    """Minimal UIA-compatible adapter around an on-screen image match."""

    ControlTypeName = "ImageControl"
    ControlType = CONTROL_TYPE_IDS["ImageControl"]
    ClassName = "ImageMatch"
    AutomationId = ""
    IsEnabled = True
    NativeWindowHandle = 0
    FrameworkId = "Image"

    def __init__(self, image_path, left, top, width, height):
        self.image_path = str(image_path)
        self.Name = Path(image_path).name
        self.BoundingRectangle = RectStruct(
            int(left), int(top), int(left + width), int(top + height)
        )
        self._parent = _ImageParentControl()

    def Exists(self, *_args, **_kwargs):
        return self.BoundingRectangle.is_valid()

    def GetParentControl(self):
        return self._parent

    def GetTopLevelControl(self):
        return self._parent

    def MoveCursorToInnerPos(self, x=None, y=None, ratioX=0.5, ratioY=0.5, **_kwargs):
        rect = self.BoundingRectangle
        x = int(rect.width * ratioX) if x is None else int(x)
        y = int(rect.height * ratioY) if y is None else int(y)
        absolute_x = (rect.left if x >= 0 else rect.right) + x
        absolute_y = (rect.top if y >= 0 else rect.bottom) + y
        pyautogui.moveTo(absolute_x, absolute_y)
        return absolute_x, absolute_y

    def Click(self, x=None, y=None, **_kwargs):
        point = self.MoveCursorToInnerPos(x, y)
        pyautogui.click(*point)

    def RightClick(self, x=None, y=None, **_kwargs):
        point = self.MoveCursorToInnerPos(x, y)
        pyautogui.rightClick(*point)

    def MiddleClick(self, x=None, y=None, **_kwargs):
        point = self.MoveCursorToInnerPos(x, y)
        pyautogui.middleClick(*point)

    def DoubleClick(self, x=None, y=None, **_kwargs):
        point = self.MoveCursorToInnerPos(x, y)
        pyautogui.doubleClick(*point)


def _find_image_control(location):
    image_path = location.get("Img")
    if not image_path:
        return None
    path = Path(image_path).expanduser()
    if not path.is_file():
        return None
    parameters = location.get("PARAMETERS") or {}
    locate_kwargs = {}
    try:
        if isinstance(parameters, dict) and parameters.get("confidence") is not None:
            locate_kwargs["confidence"] = float(parameters["confidence"])
        match = pyautogui.locateOnScreen(str(path), **locate_kwargs)
    except Exception as exc:
        push_message(f"图像定位失败: {exc}")
        return None
    if match is None:
        return None
    if hasattr(match, "left"):
        left, top, width, height = match.left, match.top, match.width, match.height
    else:
        left, top, width, height = match
    return ImageControl(path, left, top, width, height)


def to_rect(rect_like) -> RectStruct:
    """
    兼容 tuple/list 或 Rect，统一转为 Rect
    """
    if isinstance(rect_like, (tuple, list)) and len(rect_like) == 4:
        return RectStruct(*rect_like)
    if isinstance(rect_like, RectStruct):
        return rect_like
    if all(hasattr(rect_like, name) for name in ("left", "top", "right", "bottom")):
        return RectStruct(rect_like.left, rect_like.top, rect_like.right, rect_like.bottom)

    raise TypeError(f"不支持的 rect 类型: {type(rect_like)} -> {rect_like}")


def get_system_rect():
    user32 = ctypes.windll.user32
    left = user32.GetSystemMetrics(76)
    top = user32.GetSystemMetrics(77)
    width = user32.GetSystemMetrics(78)
    height = user32.GetSystemMetrics(79)
    menu_height = user32.GetSystemMetrics(15)
    system_rect = RectStruct(left, top, left + width, top + height)
    system_menu_rect = RectStruct(left, top, left + width, top + menu_height)
    system_window_max_rect = RectStruct(
        left,
        system_menu_rect.bottom,
        left + width,
        system_rect.bottom
    )
    return system_rect, system_menu_rect, system_window_max_rect


def get_check_point_by_point(window_rect, control_rect) -> RectStruct | tuple[RectStruct, int, int]:
    """
    计算控件在窗口内的可见区域（可点击区域）
    """
    window_rect = to_rect(window_rect)
    control_rect = to_rect(control_rect)

    check_left = max(control_rect.left, window_rect.left)
    check_top = max(control_rect.top, window_rect.top)
    check_right = min(control_rect.right, window_rect.right)
    check_bottom = min(control_rect.bottom, window_rect.bottom)
    visible_rect = RectStruct(check_left, check_top, check_right, check_bottom)

    # 如果不可见，返回一个空的矩形
    if not visible_rect.is_valid():
        return RectStruct(0, 0, 0, 0)

    return visible_rect


# @timeit
def get_check_point_by_control(control):
    """
    根据控件获取可见区域坐标
    """
    # print(control)
    parent = control.GetParentControl()
    if parent is None:
        return to_rect(get_control_coordinates(control))
    parent_rect = to_rect(get_control_coordinates(parent))
    # push_message(f"父控件坐标: {parent_rect}")

    control_rect = to_rect(get_control_coordinates(control))
    # push_message(f"控件坐标: {control_rect}")
    return get_check_point_by_point(parent_rect, control_rect)


def get_control_coordinates(control):
    """
    获取控件的坐标信息。

    :param control: 控件对象
    :return: 控件的左、顶、右、底坐标
    """
    bounding_rect = control.BoundingRectangle
    left = bounding_rect.left
    top = bounding_rect.top
    right = bounding_rect.right
    bottom = bounding_rect.bottom

    return left, top, right, bottom


def correct_ctrl_position(control):
    """
    校正控件位置，将鼠标移动到控件的中心。

    :param:
        -control(obj): 控件对象
    :return:
        -x: 控件中心点的x坐标
        -y: 控件中心点的y坐标
    """
    bounding_rect = control.BoundingRectangle
    left = bounding_rect.left
    top = bounding_rect.top
    right = bounding_rect.right
    bottom = bounding_rect.bottom
    x = (right - left) // 2
    y = (bottom - top) // 2

    return x, y


def package_location(WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img, PARAMETERS):
    """
    将定位器进行封装。

    :param WindowName: 窗口名称
    :param Name: 控件名称
    :param ClassName: 控件类名
    :param ControlType: 控件类型
    :param foundIndex: 控件索引
    :param AutomationId: 控件AutomationId
    :param Xpath: 控件Xpath
    :param Img: 控件图片
    :param PARAMETERS: 控件参数
    :return LOCATION 定位器
    """
    LOCATION = {
        'WindowName': WindowName,
        'Name': Name,
        'ClassName': ClassName,
        'ControlType': ControlType,
        'foundIndex': foundIndex,
        'AutomationId': AutomationId,
        'Xpath': Xpath,
        'Img': Img,
        'PARAMETERS': PARAMETERS
    }
    return LOCATION


def disassemble_location(LOCATION):
    """
    将定位器进行分解。

    :param LOCATION: 定位器
    :return:
        - WindowName: 窗口名称
        - Name: 控件名称
        - ClassName: 控件类名
        - Type: 控件类型
        - foundIndex: 控件索引
        - AutomationId: 控件AutomationId
        - Xpath: 控件Xpath
        - Img: 控件图片
        - PARAMETERS: 控件参数
    """
    WindowName = LOCATION.get('WindowName', '')
    Name = LOCATION.get('Name', '')
    ClassName = LOCATION.get('ClassName', '')
    ControlType = LOCATION.get('ControlType', '')
    foundIndex = LOCATION.get('foundIndex', 0)
    AutomationId = LOCATION.get('AutomationId', '')
    Xpath = LOCATION.get('Xpath', '[]')
    Img = LOCATION.get('Img', '')
    PARAMETERS = LOCATION.get('PARAMETERS', {})
    return WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img, PARAMETERS


def format_action_data_by_xpath(action_name, test_id, Xpath, PARAMETERS, img):
    """
    格式化动作数据字典
    参数：
    Xpath (list): 通过get_all_parent_controls获取到的控件完整路径

    返回：
    str: 转换后的动作数据字典。
    """
    # 判断是否为字符串
    if Xpath == [] or Xpath is None or len(Xpath) == 0:
        LOCATION = package_location("", "", "", "", "", "", [], "", PARAMETERS)
    else:
        LOCATION = package_location(
            WindowName=Xpath[0].get('Name', ''),
            Name=Xpath[-1].get('Name', ''),
            ClassName=Xpath[-1].get('ClassName', ''),
            ControlType=Xpath[-1].get('ControlType', ''),
            foundIndex=Xpath[-1].get('foundIndex', ''),
            AutomationId=Xpath[-1].get('AutomationId', ''),
            Xpath=Xpath,
            Img=img,
            PARAMETERS=PARAMETERS
        )

    ACTION_DATA = {
        'TEST_ID': test_id,
        'ACTION': action_name,
        'LOCATION': LOCATION,
    }
    return ACTION_DATA
