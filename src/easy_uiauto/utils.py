# -*- coding: utf-8 -*-
# @Name:            utils.py
# @Author:          tang
# @LastEditDate:    2025/9/30-9:44
# @depict:          通用字典、函数
import ctypes
import sys
import threading
import time
from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime
from functools import wraps

import pyautogui
import uiautomation
from comtypes import CoUninitialize, CoInitialize

# 虚拟按键
MODIFIER_VK = {
    160, 161, 162, 163, 164, 165  # 修饰键（Shift、Ctrl、Alt等）的虚拟键码集合
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
CURRENT_APP_NAME = 'uiautomation'
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
CONTROL_CACHE = {}  # 控件缓存字典，存储XPath到控件的映射
CACHE_METADATA = {}  # 缓存元数据字典
CACHE_QUEUE = []  # 缓存队列，用于异步处理
CACHE_PROCESSING = False  # 缓存处理状态标志
CONTROL_CACHE_TIMEOUT = 5000  # 控件缓存超时时间（毫秒）
# 编译缓存
APP = []  # 应用程序列表
COMPILE_COUNT = 0  # 编译计数器
COMPILE_PROGRESS = 120  # 编译进度基准值


def timeit(func):
    """测量函数执行时间"""

    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()  # 高精度计时开始
        result = func(*args, **kwargs)  # 执行目标函数
        elapsed = time.perf_counter() - start_time  # 计算耗时
        push_message(f"[计时] {func.__name__} 执行耗时: {elapsed * 1000:.3f} 毫秒")
        return result  # 返回原函数结果

    return wrapper


def push_message(log, log_path=''):
    """
    此函数用于替代print()

    :param log: 日志信息
    :param log_path: 保存文件路径
    :return:
    """
    print(log)
    if log_path:
        with open(log_path, 'a', encoding='utf-8') as file:
            file.write(log)
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
    user32 = _get_user32()
    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return ""

    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def _get_user32():
    """Configure pointer-sized HWND arguments for 64-bit Windows."""
    # Keep ctypes signatures private: pygetwindow uses its own EnumWindows
    # callback type on ctypes.windll.user32 in the same MCP process.
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    for name in ('IsWindowVisible', 'IsIconic', 'SetForegroundWindow'):
        getattr(user32, name).argtypes = (ctypes.c_void_p,)
    user32.ShowWindow.argtypes = (ctypes.c_void_p, ctypes.c_int)
    user32.GetWindowTextLengthW.argtypes = (ctypes.c_void_p,)
    user32.GetWindowTextW.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int)
    return user32


def _find_visible_window_handle(window_title):
    """Enumerate top-level Win32 windows and stop at the first exact title match."""
    user32 = _get_user32()
    current = user32.GetForegroundWindow()
    if current and user32.IsWindowVisible(current) and get_window_title_by_handle(current) == window_title:
        return current

    match = None

    callback_type = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)

    @callback_type
    def check_window(hwnd, _):
        nonlocal match
        if user32.IsWindowVisible(hwnd) and get_window_title_by_handle(hwnd) == window_title:
            match = hwnd
            return False
        return True

    user32.EnumWindows.argtypes = (callback_type, ctypes.c_void_p)
    user32.EnumWindows(check_window, None)
    return match


@timeit
def set_top_window(window_title):
    """Restore and activate a visible top-level window by its exact title."""
    if not window_title:
        return False
    try:
        user32 = _get_user32()
        hwnd = _find_visible_window_handle(window_title)
        if not hwnd:
            push_message(f"未找到窗口：{window_title}")
            return False
        if user32.GetForegroundWindow() == hwnd:
            return True
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.005)
        activated = user32.GetForegroundWindow() == hwnd
        if not activated:
            push_message(f"窗口激活失败：{window_title}")
        return activated
    except Exception as e:
        push_message(f"激活异常：{e}")
        return False


def get_control_info(x, y, *, match_mode="identity", raise_errors=False):
    """
    获取坐标点控件Xpath、控件
    :param x: 横坐标
    :param y: 纵坐标
    :return: 控件Xpath、控件
    """
    uiautomation.SetGlobalSearchTimeout(3)
    CoInitialize()

    try:
        control = uiautomation.ControlFromPoint(x, y)
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
        xpath = get_control_xpath(control, x, y, match_mode=match_mode)
        return xpath, control

    except Exception as e:
        if raise_errors:
            raise
        push_message(f"获取控件信息异常: {e}")
        return [], None
    finally:
        CoUninitialize()


def auto_scroll(pixels, duration=0, direction='down'):
    """
    滚动鼠标
    :param pixels: 滚动距离(像素)
    :param duration: 滚动持续时间(秒)
    :param direction: 滚动方向('up'或'down')
    """
    try:
        if direction.lower() not in ['up', 'down']:
            raise ValueError("方向参数必须是'up'或'down'")

        scroll = -pixels if direction == 'up' else pixels
        pyautogui.scroll(scroll, _pause=False)
        if duration > 0:
            time.sleep(duration)
        # push_message(f"已滚动{scroll}像素")
    except Exception as e:
        push_message(f"发生错误: {e}")


# #######################编译#############################
def set_global_control_cache_timeout(seconds):
    """设置全局控件缓存超时时间"""
    global CONTROL_CACHE_TIMEOUT
    CONTROL_CACHE_TIMEOUT = seconds


# @timeit
def compile_controls(control=None, max_depth=15, compile_log=False):
    """
    编译控件
    :param control: 控件
    :param max_depth: 最大遍历深度
    :param compile_log: 编译日志
    """
    push_message(f"\n编译开始预计耗时：{max_depth * 0.8}秒")
    uiautomation.SetGlobalSearchTimeout(1)
    global COMPILE_PROGRESS
    COMPILE_PROGRESS = COMPILE_PROGRESS * max_depth
    if control is None:
        compile_app = uiautomation.GetRootControl().GetChildren()
        for app in compile_app:
            xpath = get_control_xpath(app)[1:]
            _recursive_cache_controls(app, xpath, max_depth, 1, compile_log)
    else:
        xpath = get_control_xpath(control)[1:]
        _recursive_cache_controls(control, xpath, max_depth, 1, compile_log)

    uiautomation.SetGlobalSearchTimeout(5)
    sys.stdout.write("\r编译进度: |%s%s| %d%%" % ('█' * int(100), ' ' * (100 - int(100)), 100))
    sys.stdout.flush()
    push_message(f"\n编译完成，已添加 {len(CONTROL_CACHE)} 个控件")


def generate_cache_keys(xpath):
    """
    生成XPath路径的缓存键，每个控件生成两个key：
    1. 带foundIndex的key（用于精确匹配）
    2. 不带foundIndex的key（用于备选匹配）
    """
    keys_with_foundindex = []
    keys_without_foundindex = []

    for item in xpath:
        if item.get('ControlType') not in {'DesktopControl', 'CustomControl', 'GroupControl'} and \
                item.get('ClassName') not in {'QWidget',
                                              'GroupControl',
                                              'CustomControl',
                                              'WindowControl', '#32769'}:
            # 生成带foundIndex的key
            item_keys_with = []
            if item.get('ControlType'):
                ControlType = uiautomation.ControlTypeNames.get(item['ControlType'], item['ControlType'])
                item_keys_with.append(ControlType)
            if item.get('Name'):
                item_keys_with.append(f"Name:{item['Name']}")
            if item.get('ClassName'):
                item_keys_with.append(f"ClassName:{item['ClassName']}")
            if item.get('AutomationId'):
                item_keys_with.append(f"AutomationId:{item['AutomationId']}")
            if 'searchDepth' in item:
                item_keys_with.append(f"searchDepth:{str(item['searchDepth'])}")
            # 只有当foundIndex存在且大于0时才添加
            if item.get('foundIndex'):
                item_keys_with.append(f"foundIndex:{str(item['foundIndex'])}")

            key_with = ",".join(item_keys_with)
            keys_with_foundindex.append(key_with)

            # 生成不带foundIndex的key（移除foundIndex字段）
            item_keys_without = item_keys_with[:]
            # 如果有foundIndex字段，则从不带foundIndex的key中移除
            for i, key_part in enumerate(item_keys_without):
                if key_part.startswith("foundIndex:"):
                    item_keys_without.pop(i)
                    break

            key_without = ",".join(item_keys_without)
            keys_without_foundindex.append(key_without)

    # 返回两个完整路径键
    return "|".join(keys_with_foundindex), "|".join(keys_without_foundindex)


def _recursive_cache_controls(ctrl, parent_xpath, max_depth, current_depth, compile_log=False):
    """递归缓存控件及其子控件"""
    global COMPILE_COUNT, COMPILE_PROGRESS
    progress = (COMPILE_COUNT / COMPILE_PROGRESS) * 100
    sys.stdout.write("\r编译进度: |%s%s| %d%%" % ('█' * int(progress), ' ' * (100 - int(progress)), progress))
    sys.stdout.flush()
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
        CONTROL_CACHE[key_with] = ctrl
        CONTROL_CACHE[key_without] = ctrl
        COMPILE_COUNT += 1
        if current_depth < max_depth:
            try:
                children = ctrl.GetChildren()
                for idx, child in enumerate(children):
                    _recursive_cache_controls(child, ctrl_xpath, max_depth, current_depth + 1, compile_log)
            except Exception as e:
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
    global CACHE_PROCESSING
    CACHE_QUEUE.append((cache_key, control))
    if not CACHE_PROCESSING:
        CACHE_PROCESSING = True
        threading.Thread(target=_process_cache_queue, daemon=True).start()


def _process_cache_queue():
    """处理缓存队列"""
    global CACHE_PROCESSING
    while CACHE_QUEUE:
        cache_key, control = CACHE_QUEUE.pop(0)
        if cache_key not in CONTROL_CACHE:
            CONTROL_CACHE[cache_key] = control
            CACHE_METADATA[cache_key] = {
                'last_verified': time.time(),
                'window_handle': getattr(control, 'NativeWindowHandle', 0)
            }
        time.sleep(0.001)
    CACHE_PROCESSING = False


def generate_cache_key(xpath):
    """生成XPath路径的缓存键"""
    key_parts = []
    for item in xpath:
        item_key = (
            f"ControlType:{item.get('ControlType', '')},"
            f"Name:{item.get('Name', '')},"
            f"ClassName:{item.get('ClassName', '')},"
            f"AutomationId:{item.get('AutomationId', '')},"
            f"searchDepth:{item.get('searchDepth', 0)}"
        )
        key_parts.append(item_key)
    return "|".join(key_parts)

def _brothers(ctrl):
    while 10:
        children = ctrl.GetChildren()
        if len( children)>1:
            return  children
        if len(children)==1:
            ctrl = children[0]
        else:
            return  ctrl

# @timeit
def find_control_by_xpath(xpath, debug=False, *, match_mode=None):
    """Resolve XPath; an explicit mode overrides the recorded per-step modes."""
    return find_control({"Xpath": xpath}, debug=debug, match_mode=match_mode)


def get_control_xpath(control, x=None, y=None, *, match_mode="identity"):
    """
    获取控件Xpath
    :param x:横坐标
    :param y:纵坐标
    :param control: 当前控件对象
    :return: Xpath
    """

    from .mcp.protocol import normalize_match_mode

    mode = normalize_match_mode(match_mode)
    lineage = []
    while control:
        if control.ControlTypeName == 'DesktopControl' or control.ClassName == '#32769':
            break
        parent = control.GetParentControl()
        lineage.append((control, parent))
        if control.ControlTypeName == 'WindowControl' and (
                parent is None or parent.ControlTypeName not in {
                    'TableControl', 'ListControl', 'TreeControl', 'TabControl',
                    'WindowControl', 'GroupControl', 'CustomControl'}):
            break
        control = parent

    xpath = []
    for depth, (node, parent) in enumerate(reversed(lineage), 1):
        step = {"ControlType": node.ControlTypeName, "searchDepth": depth,
                "match_mode": mode}
        for key in ("Name", "ClassName", "AutomationId"):
            value = getattr(node, key, "")
            if value:
                step[key] = value
        # Retain real ancestors so a type ordinal is always relative to its parent.
        # Name equality cannot establish which same-named sibling was recorded.
        if depth > 1 and parent is not None:
            ordinal = 0
            for sibling in parent.GetChildren():
                if sibling.ControlType != node.ControlType:
                    continue
                ordinal += 1
                same = sibling is node
                if not same:
                    try:
                        same = uiautomation.ControlsAreSame(sibling, node)
                    except (AttributeError, TypeError):
                        same = False
                if same:
                    step["foundIndex"] = ordinal
                    break
            else:
                raise ValueError("Recorded control is no longer a child of its parent")
        xpath.append(step)
    return xpath


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
def strategy_xpath(name, class_name, control_type, Xpath,debug=False):
    try:
        ctrl = find_control_by_xpath(Xpath, debug)

        if ctrl == uiautomation:
            return False
        if ctrl and ctrl.ControlTypeName == control_type and ctrl.Name == name and ctrl.ClassName == class_name:
            return ctrl

        elif ctrl and control_type in ITEM_TYPE_NAMES:
            if ctrl.ControlTypeName in ITEM_TYPE_NAMES:
                ctrl = ctrl.GetParentControl()
            ctrl.MoveCursorToInnerPos(simulateMove=False)
            Children = ctrl.GetChildren()
            ctrl_hight = ctrl.BoundingRectangle.height()
            child_hight = Children[0].BoundingRectangle.height()
            lastchild_name = Children[0].Name

            if ctrl_hight * 0.8 < child_hight * len(Children):
                for i in range(50):
                    auto_scroll(500, direction='down')
            rollsize = ctrl_hight * 0.5
            find_count = 0
            while True:
                current_ctrl = find_control_by_xpath(Xpath, debug)
                if current_ctrl and current_ctrl.ControlTypeName == control_type and current_ctrl.Name == name and current_ctrl.ClassName == class_name:
                    ctrl = current_ctrl
                    break
                current_ctrl = redirect_control(name, current_ctrl)
                if debug:
                    push_message('重定向当前控件：{}'.format(current_ctrl))

                if current_ctrl and current_ctrl.ControlTypeName == control_type and current_ctrl.Name == name and current_ctrl.ClassName == class_name:
                    ctrl = current_ctrl
                    break
                auto_scroll(int(rollsize), direction='up')
                # 校验是否滚动到底

                current_lastchild_name = ctrl.GetChildren()[0].Name
                if current_lastchild_name == lastchild_name:
                    find_count += 1
                    if find_count == 4:
                        ctrl = current_ctrl
                        break
                else:
                    find_count = 0
                lastchild_name = current_lastchild_name
            if debug:
                push_message(f'策略 Xpath 滚动定位成功')
            return ctrl
        return False
    except Exception as e:
        if debug:
            push_message(f'策略 Xpath 错误: {str(e)}')
        return False


def _find_window_scope(window_name, xpath=None):
    if not window_name:
        return uiautomation.GetRootControl()
    from .mcp.results import ActionError

    try:
        return _native_window_scope(window_name, xpath[0] if xpath else {})
    except ActionError as error:
        if error.code == "window_invalid":
            return False
        raise


def _selector_index(value):
    """Validate a recorded ordinal once, consistently for windows and controls."""
    from .mcp.results import ActionError

    try:
        if isinstance(value, bool) or (value not in (None, "") and not isinstance(value, (str, int))):
            raise ValueError
        index = int(value) if value not in (None, "") else 0
        if index < 0:
            raise ValueError
        return index
    except (TypeError, ValueError) as error:
        raise ActionError("invalid_argument", "foundIndex must be zero or a positive 1-based integer") from error


def _native_window_scope(window_name, first_step, *, native_title=False):
    """Enumerate HWNDs natively, then match the recorded UIA window identity.

    Native titles/classes only narrow the initial candidates. A provider's Name
    need not equal GetWindowTextW. Explicit title activation uses native_title.
    """
    from .mcp import native_windows
    from .mcp.results import ActionError

    try:
        index = _selector_index(first_step.get("foundIndex"))
        windows = native_windows.enumerate_windows(visible_only=True)
        bound = {}
        if native_title:
            candidates = [w for w in windows if w["title"] == window_name]
        else:
            expected = {key: first_step[key] for key in ("Name", "ClassName", "AutomationId")
                        if first_step.get(key)}
            if window_name and not expected.get("Name"): expected["Name"] = window_name
            expected_type = first_step.get("ControlType")
            expected_class = expected.get("ClassName")
            titles = {n for n in (window_name, expected.get("Name")) if n}
            preferred = [w for w in windows if w["title"] in titles
                         or expected_class and w.get("class_name") == expected_class]

            def match(entries):
                matches = []
                for entry in entries:
                    control = uiautomation.ControlFromHandle(entry["handle"])
                    if control is None:
                        raise RuntimeError("Could not bind UIA to a candidate window")
                    if (all(getattr(control, key) == value for key, value in expected.items())
                            and (not expected_type or control.ControlTypeName == expected_type)):
                        matches.append(entry)
                        bound[entry["handle"]] = control
                return matches

            candidates = match(preferred or windows)
            if not candidates and preferred:
                tried = {w["handle"] for w in preferred}
                candidates = match([w for w in windows if w["handle"] not in tried])
        if not candidates or index > len(candidates):
            identity_mismatch = (not native_title and not candidates and window_name
                                 and any(w["title"] == window_name for w in windows))
            raise ActionError(
                "window_invalid",
                "Native title matches but the recorded UIA window identity does not match" if identity_mismatch
                else "Window not found or no longer visible",
                retryable=not bool(index) and not bool(identity_mismatch),
                reason="index_out_of_range" if index else "window_identity_mismatch" if identity_mismatch else "",
            )
        if not index and len(candidates) != 1:
            raise ActionError("ambiguous_match", "Multiple visible windows match the location")
        selected = candidates[(index or 1) - 1]
        control = bound.get(selected["handle"])
        if control is None:
            control = uiautomation.ControlFromHandle(selected["handle"])
        if control is None:
            raise RuntimeError("Could not bind UIA to the selected window")
        return control
    except ActionError:
        raise
    except Exception as error:
        raise ActionError("action_failed", f"Native window lookup failed: {error}") from error


def _has_second_match(window, selector):
    """Check whether a window-scoped selector matches more than one control."""
    conditions = {
        key: selector[key]
        for key in ('Name', 'ClassName', 'AutomationId')
        if selector.get(key)
    }
    control_type = selector.get('ControlType')
    if control_type:
        conditions['ControlType'] = CONTROL_TYPE_IDS[control_type]
    try:
        second = window.Control(foundIndex=2, **conditions)
        return bool(second.Refind(maxSearchSeconds=0, raiseException=False))
    except Exception:
        # If uniqueness cannot be checked, do not trust the first match.
        return True


_STRICT_CONTROL_RESOLUTION = ContextVar("strict_control_resolution", default=False)
_ACTIVATE_RESOLVED_WINDOW = ContextVar("activate_resolved_window", default=False)


@contextmanager
def strict_control_resolution(*, activate=False):
    """Use typed lookup; only mutation scopes may activate a resolved window."""
    token = _STRICT_CONTROL_RESOLUTION.set(True)
    activation_token = _ACTIVATE_RESOLVED_WINDOW.set(activate)
    try:
        yield
    finally:
        _ACTIVATE_RESOLVED_WINDOW.reset(activation_token)
        _STRICT_CONTROL_RESOLUTION.reset(token)


def _observe_match(parent, selector, *, search_depth=0xFFFFFFFF, absent_code="control_not_found"):
    """Apply one explicit selection rule without mixing identity and type ordinals."""
    from .mcp.results import ActionError
    from .mcp.protocol import normalize_match_mode

    mode = normalize_match_mode(selector.get("match_mode", "identity"))
    index = _selector_index(selector.get("foundIndex"))
    if mode == "position":
        if not selector.get("ControlType") or not index:
            raise ActionError("invalid_argument", "Position mode requires ControlType and a positive foundIndex")
        conditions = {}
        search_depth = 1
    else:
        keys = ("Name", "ClassName") if selector.get("Name") else ("AutomationId", "ClassName")
        conditions = {key: selector[key] for key in keys if selector.get(key) not in (None, "")}
        index = 1
    control_type = selector.get("ControlType")
    if control_type:
        if control_type not in CONTROL_TYPE_IDS:
            raise ActionError("invalid_argument", f"Unknown ControlType: {control_type}")
        conditions["ControlType"] = CONTROL_TYPE_IDS[control_type]
    if not conditions:
        raise ActionError("invalid_argument", "A control selector must include a name, class, type, or AutomationId")
    conditions["searchDepth"] = search_depth

    def lookup(found_index):
        try:
            candidate = parent.Control(foundIndex=found_index, **conditions)
            return candidate if candidate.Refind(maxSearchSeconds=0, raiseException=False) else None
        except ActionError:
            raise
        except Exception as error:
            raise ActionError("action_failed", f"UIA lookup failed: {error}") from error

    control = lookup(index)
    if control is None:
        message = "Window not found or no longer available" if absent_code == "window_invalid" else "Control not found"
        raise ActionError(absent_code, message, retryable=mode == "identity",
                          reason="index_out_of_range" if mode == "position" else "")
    return control


def observe_control(location, *, window=None, prefix_cache=None, search_depth=0xFFFFFFFF, match_mode=None):
    """Read a LOCATION without activation, scrolling, or persistent caches.

    A supplied window and prefix cache belong to one non-mutating inspection only.
    Identity defaults to the first matching name; position uses a 1-based type ordinal.
    """
    import json
    from .mcp.protocol import normalize_location, normalize_match_mode, with_match_mode
    from .mcp.results import ActionError

    if isinstance(search_depth, bool) or not isinstance(search_depth, int) or search_depth < 1:
        raise ActionError("invalid_argument", "search_depth must be a positive integer")
    try:
        normalized = (with_match_mode(location, match_mode) if match_mode is not None
                      else normalize_location(location))
    except (TypeError, ValueError, AttributeError) as error:
        raise ActionError("invalid_argument", str(error)) from error
    xpath = normalized["Xpath"]
    default_mode = normalized["PARAMETERS"].get("match_mode", "identity")
    leaf_mode = xpath[-1].get("match_mode", default_mode) if xpath else default_mode
    if leaf_mode == "position" and xpath and normalized.get("foundIndex") not in (None, "", 0, "0"):
        leaf_index = xpath[-1].get("foundIndex")
        if leaf_index not in (None, "", 0, "0") and str(leaf_index) != str(normalized["foundIndex"]):
            raise ActionError("invalid_argument", "LOCATION and XPath specify different leaf indices")
        xpath[-1] = {**xpath[-1], "foundIndex": normalized["foundIndex"]}
    first_step = xpath[0] if xpath else {}
    window_name = normalized["WindowName"]
    window_step = first_step.get("ControlType") in {"WindowControl", "DesktopControl"}
    # Validate all selectors before any backend call or shortcut can bypass them.
    for selector in [normalized, *xpath]:
        _selector_index(selector.get("foundIndex"))
        normalize_match_mode(selector.get("match_mode", default_mode))
        try:
            raw_depth = selector.get("searchDepth")
            if raw_depth is not None and (isinstance(raw_depth, bool) or not isinstance(raw_depth, (str, int)) or int(raw_depth) < 0):
                raise ValueError
        except (TypeError, ValueError) as error:
            raise ActionError("invalid_argument", "foundIndex and searchDepth must be nonnegative integers") from error
        desktop_root = (selector.get("ControlType") == "DesktopControl"
                        and (selector is first_step or len(xpath) == 1))
        if selector.get("ControlType") and selector["ControlType"] not in CONTROL_TYPE_IDS and not desktop_root:
            raise ActionError("invalid_argument", f"Unknown ControlType: {selector['ControlType']}")
    if window is None:
        if first_step.get("ClassName") in {"Progman", "WorkerW", "Shell_TrayWnd", "Shell_SecondaryTrayWnd"}:
            window = _native_window_scope("", first_step)
        elif first_step.get("ControlType") != "DesktopControl" and (window_name or window_step):
            window = _native_window_scope(window_name, first_step if window_step else {})
        else:
            try:
                window = uiautomation.GetRootControl()
                if window is None:
                    raise RuntimeError("UIA returned no desktop root")
            except Exception as error:
                raise ActionError("action_failed", f"UIA desktop lookup failed: {error}") from error

    failures = []
    cache = prefix_cache if prefix_cache is not None else {}

    def attempt(parent, selector, search_depth=search_depth):
        try:
            return _observe_match(parent, selector, search_depth=search_depth)
        except ActionError as error:
            if error.reason == "index_out_of_range" or error.code not in {"control_not_found", "ambiguous_match"}:
                raise
            failures.append(error)
            return None

    if xpath:
        # A bound shell root can itself be a PaneControl rather than WindowControl.
        root_matches = all(not first_step.get(key) or first_step[key] == getattr(window, attr, None)
                           for key, attr in (("Name", "Name"), ("ClassName", "ClassName"),
                                             ("AutomationId", "AutomationId")))
        type_name = getattr(window, "ControlTypeName", None)
        if type_name is not None:
            root_matches = root_matches and first_step.get("ControlType") in (None, "", type_name)
        else:
            root_matches = root_matches and CONTROL_TYPE_IDS.get(first_step.get("ControlType")) == getattr(window, "ControlType", None)
        skip_root = window_step or root_matches
        steps = xpath[1:] if skip_root else xpath
        current = window
        previous_depth = int(first_step.get("searchDepth", 1) or 1) if skip_root else 0
        prefix = []
        for step in steps:
            depth = int(step.get("searchDepth", previous_depth + 1) or 1)
            selector = dict(step)
            selector["match_mode"] = step.get("match_mode", default_mode)
            prefix.append(json.dumps(selector, sort_keys=True, ensure_ascii=False))
            cache_key = (id(window), *prefix)
            cached = cache.get(cache_key) if step is not steps[-1] else None
            if cached is not None:
                try:
                    if not cached.Refind(maxSearchSeconds=0, raiseException=False):
                        cached = None
                except Exception as error:
                    raise ActionError("action_failed", f"Cached UIA lookup failed: {error}") from error
            current = cached if cached is not None else attempt(current, selector, max(1, depth - previous_depth))
            if current is None:
                break
            if step is not steps[-1]:
                cache[cache_key] = current
            previous_depth = depth
        if current is not None:
            return current
        # A positional path is parent-relative; a global leaf search could change it.
        if any(step.get("match_mode", default_mode) == "position" for step in steps):
            raise ActionError("control_not_found", "Positional XPath parent was not found",
                              reason="position_context_missing")
    # A failed ambiguous hierarchy cannot be bypassed by a leaf-only fallback.
    ambiguous = next((error for error in failures if error.code == "ambiguous_match"), None)
    if ambiguous:
        raise ambiguous
    target_selector = {key: normalized.get(key) for key in ("Name", "AutomationId", "ClassName", "ControlType", "foundIndex")}
    target_selector["match_mode"] = default_mode
    if any(target_selector.get(key) for key in ("Name", "AutomationId", "ClassName", "ControlType")):
        control = attempt(window, target_selector)
        if control is not None:
            return control
    elif not normalized.get("AutomationId") and not xpath:
        if window_name:
            return window
        raise ActionError("invalid_argument", "A location must identify a window or a control")
    ambiguous = next((error for error in failures if error.code == "ambiguous_match"), None)
    if ambiguous:
        raise ambiguous
    raise ActionError("control_not_found", "Control not found", retryable=True)


@timeit
def find_control(LOCATION, debug=False, strict=True, *, match_mode=None):
    """Use the shared mode rules; strict remains a compatibility-only flag."""
    control = (observe_control(LOCATION, match_mode=match_mode) if match_mode is not None
               else observe_control(LOCATION))
    if _ACTIVATE_RESOLVED_WINDOW.get():
        from .ctrl import Controller
        from .mcp.results import ActionError

        try:
            Controller._activate_target(control)
        except ActionError:
            raise
        except Exception as error:
            raise ActionError("window_invalid", f"Resolved window could not be activated: {error}") from error
    return control


def strategy_dictionary(dictionary, Prant=uiautomation.GetRootControl(), timeout=10):
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
    uiautomation.SetGlobalSearchTimeout(timeout)
    conditions = {
        'Name': None if dictionary.get('Name') == '' else dictionary.get('Name', None),
        'ClassName': None if dictionary.get('ClassName') == '' else dictionary.get('ClassName', None),
        'AutomationId': None if dictionary.get('AutomationId') == '' else dictionary.get('AutomationId', None),
        # 'searchDepth': dictionary.get('searchDepth', 0xFFFFFFFF),
        # 'foundIndex': dictionary.get('foundIndex', 1)
    }
    control_type_name = dictionary.get('ControlType', '')
    if control_type_name:
        control_type = CONTROL_TYPE_IDS.get(control_type_name)
        if control_type is None:
            return False
        conditions['ControlType'] = control_type
    try:
        control = Prant.Control(**conditions)
        state = control.Refind(maxSearchSeconds=0, raiseException=False)
        if state:
            return control
        return state
    except:
        return False


from dataclasses import dataclass


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
        if x < 0 or y < 0:
            return 0, 0
        return x, y

    def is_valid(self) -> bool:
        return self.width > 0 and self.height > 0


def to_rect(rect_like) -> RectStruct:
    """
    兼容 tuple/list 或 Rect，统一转为 Rect
    """
    if isinstance(rect_like, (tuple, list)) and len(rect_like) == 4:
        return RectStruct(*rect_like)
    if isinstance(rect_like, RectStruct):
        return rect_like

    raise TypeError(f"不支持的 rect 类型: {type(rect_like)} -> {rect_like}")


def get_system_rect():
    # 系统（上，顶，右，下）最大尺寸
    system_rect = RectStruct(0, 0, 1920, 1080)
    system_menu_rect = RectStruct(0, 0, 1920, 40)
    system_window_max_rect = RectStruct(
        0,
        system_menu_rect.bottom,
        1920,
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

    # 移动光标到控件中心
    control.MoveCursorToInnerPos(x=x, y=y, simulateMove=False)
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
    ControlType = LOCATION.get('ControlType', LOCATION.get('Type', ''))
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
