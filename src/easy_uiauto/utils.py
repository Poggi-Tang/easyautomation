# -*- coding: utf-8 -*-
# @Name:            utils.py
# @Author:          tang
# @LastEditDate:    2025/9/30-9:44
# @depict:          通用字典、函数
import ctypes
import sys
import threading
import time
from datetime import datetime
from functools import wraps

import pyautogui
import pygetwindow as gw
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
    try:
        # 使用uiautomation检查
        window = uiautomation.WindowControl(Name=window_title, searchDepth=1)
        if window.Exists(0):
            window_title = get_window_title_by_handle(window.NativeWindowHandle)
            if window.FrameworkId == 'Win32':
                windows = gw.getWindowsWithTitle(window_title)
                if windows:
                    for window in windows:
                        if window.title == window_title:
                            window.activate()
                            # time.sleep(0.2)
                            active_window = gw.getActiveWindow()
                            if active_window and active_window.title == window_title:
                                push_message(f"使用pygetwindow激活成功")
                                return True
                push_message(f"未找到窗口")
                return False
            elif 'dialog' in window.Name or window.ClassName == 'Window':
                push_message("对话框不支持激活")
                return False
            elif 'Menu' in window.Name or 'Menu' in window.ClassName:
                push_message("菜单不支持激活")
                return False
            elif window.GetTopLevelControl() != window:
                push_message("非主窗口不支持激活")
                return False
            window.SetActive()
            # time.sleep(0.2)
            active_window = gw.getActiveWindow()
            if active_window and active_window.title in window_title:
                push_message(f"使用uiautomation激活成功")
                return True
            push_message(f"激活失败，尝试通过pygetwindow激活")
        windows = gw.getWindowsWithTitle(window_title)
        if windows:
            for window in windows:
                if window.title == window_title:
                    window.activate()
                    time.sleep(0.2)
                    active_window = gw.getActiveWindow()
                    if active_window and active_window.title == window_title:
                        push_message(f"使用pygetwindow激活成功")
                        return True
        push_message(f"未找到窗口")
        return False
    except Exception as e:
        push_message(f"激活异常{e}")
        return False


def get_control_info(x, y):
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
        xpath = get_control_xpath(control, x, y)
        CoUninitialize()
        return xpath, control

    except Exception as e:
        push_message(f"获取控件信息异常: {e}")
        CoUninitialize()
        return [], None


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


# @timeit
def find_control_by_xpath(xpath):
    """获取控件的Xpath路径"""
    key_with, key_without = generate_cache_keys(xpath)
    try:
        # current_ctrl = CONTROL_CACHE.get(key1,CONTROL_CACHE.get(key2,None))
        current_ctrl = CONTROL_CACHE.get(key_with, None)
        if current_ctrl:
            # 快速检测
            isExists = current_ctrl.ControlType
            # 计算控件可见区域
            rect = get_check_point_by_control(current_ctrl)
            if rect.check_pos != (0, 0):
                # 检测成功
                # push_message("使用缓存", current_ctrl)
                return current_ctrl
        raise Exception("使用缓存失败，重新建立缓存")
    except Exception as e:
        # push_message(e)
        try:
            current_ctrl = uiautomation.GetRootControl()
            last_ctrl = current_ctrl
            # 创建xpath的副本，避免修改原始数据
            xpath_copy = [item.copy() for item in xpath]
            for i, item in enumerate(xpath_copy):
                item_name = item.get('Name', '')
                item_class = item.get('ClassName', '')
                item_type = item.get('ControlType', '')
                if i != len(xpath_copy) - 1 and item_type in {"CustomControl", "GroupControl"} or item_class in {'#32769'}:
                    continue
                if "Area" in item_name or "Area" in item_class or "Splitter" in item_name or "Splitter" in item_class:
                    continue
                item_foundIndex = item.get('foundIndex', -1)
                if item_foundIndex in [0, 1]:
                    del xpath_copy[i]['foundIndex']
                item_automation = item.get('AutomationId', -1)
                if item_automation == '':
                    del xpath_copy[i]['AutomationId']
                # ===============节点类型直接匹配===============
                if item_type in ITEM_TYPE_NAMES or item_foundIndex not in [-1, 0, 1]:
                    brothers = current_ctrl.GetChildren()
                    found_index = 1
                    matched = {}
                    for brother in brothers:
                        if brother.ControlTypeName == item['ControlType']:
                            if brother.Name == item.get('Name', '') and \
                                    brother.ClassName == item.get('ClassName', '') \
                                    and brother.AutomationId == item.get('AutomationId', ''):
                                matched[found_index] = brother
                            found_index += 1
                    matched_keys = list(matched.keys())
                    matched_keys_count = len(matched_keys)
                    if matched_keys_count == 1:
                        current_ctrl = matched[matched_keys[0]]
                        CONTROL_CACHE[key_with] = current_ctrl
                        CONTROL_CACHE[key_without] = current_ctrl
                        # push_message(f'第【{i}】层{current_ctrl}')
                        return current_ctrl
                    elif matched_keys_count > 1:
                        index = item.get('foundIndex', 0)
                        if index <= 1:
                            current_ctrl = matched[matched_keys[0]]
                        elif index in matched_keys:
                            current_ctrl = matched[index]
                        else:
                            current_ctrl = matched[matched_keys[-1]]
                        if current_ctrl is None:
                            # 均未找到匹配的index就返回第一个
                            current_ctrl = matched[matched_keys[0]]
                        CONTROL_CACHE[key_with] = current_ctrl
                        CONTROL_CACHE[key_without] = current_ctrl
                        # push_message('添加缓存2', current_ctrl)
                        current_ctrl = CONTROL_CACHE.get(key_with, CONTROL_CACHE.get(key_without, None))
                        # push_message("重新使用缓存", current_ctrl)
                        # push_message(f'第【{i}】层{current_ctrl}')
                        return current_ctrl
                    else:
                        # push_message(f'第【{i}】层{last_ctrl}')
                        return last_ctrl
                # ===============匹配程序（root、Win32/Qt）===============
                abs_model = True  # qt
                TopLeve = last_ctrl.GetTopLevelControl()
                if TopLeve:
                    if TopLeve.FrameworkId == 'Win32':
                        abs_model = False
                else:  # root
                    abs_model = False

                if abs_model:
                    xpath_copy[i]['ControlType'] = CONTROL_TYPE_IDS[xpath_copy[i]['ControlType']]
                    current_ctrl = current_ctrl.Control(**xpath_copy[i])
                else:
                    xpath_copy[i]['ControlType'] = CONTROL_TYPE_IDS[xpath_copy[i]['ControlType']]
                    if 'foundIndex' in xpath_copy[i]:
                        del xpath_copy[i]['foundIndex']
                    current_ctrl = current_ctrl.Control(**xpath_copy[i])
                # debug
                # push_message(f'第【{i}】层{current_ctrl}')
                last_ctrl = current_ctrl
            CONTROL_CACHE[key_with] = current_ctrl
            CONTROL_CACHE[key_without] = current_ctrl
            return current_ctrl
        except Exception as e:
            push_message(e)


def get_control_xpath(control, x=None, y=None):
    """
    获取控件Xpath
    :param x:横坐标
    :param y:纵坐标
    :param control: 当前控件对象
    :return: Xpath
    """

    xpath = []
    over_tag = False
    search_depth = 0
    search_depth_list = []
    # desktop = True

    while control:
        foundIndex = 0
        search_depth += 1
        if control.ControlTypeName not in {'DesktopControl'} and \
                control.ClassName not in {'QWidget',
                                          'GroupControl',
                                          'CustomControl',
                                          'WindowControl', '#32769'}:
            if control.ControlTypeName in ITEM_TYPE_NAMES:
                for child in control.GetParentControl().GetChildren():
                    if child.ControlType == control.ControlType:
                        foundIndex += 1
                    if child.Name == control.Name:
                        if x and y:
                            boundingRect = child.BoundingRectangle
                            if boundingRect.left <= x <= boundingRect.right and boundingRect.top <= y <= boundingRect.bottom:
                                control = child
                                break
                        else:
                            control = child
            current_control_info = {
                "ControlType": control.ControlTypeName,
            }

            if control.Name:
                current_control_info['Name'] = control.Name
            if control.ClassName:
                current_control_info['ClassName'] = control.ClassName
            if control.AutomationId:
                current_control_info['AutomationId'] = control.AutomationId
            if foundIndex != 0:
                current_control_info['foundIndex'] = foundIndex
            # else:
            #     current_control_info['foundIndex'] = 1

            xpath.append(current_control_info)
            search_depth_list.append(search_depth)

        control = control.GetParentControl()
        if over_tag:
            break
        parent = control.GetParentControl()
        if control and parent:
            if control.ControlTypeName == 'WindowControl' and \
                    parent.ControlTypeName not in [
                'TableControl', 'ListControl', 'TreeControl', 'TabControl', 'WindowControl', 'GroupControl', 'CustomControl']:
                # print(control.GetParentControl())
                over_tag = True
        else:
            over_tag = True

    # if xpath and xpath[0].get('ClassName', '') == '#32769' and desktop:
    #     xpath = xpath[1:]
    xpath = xpath[::-1]
    for index in range(len(xpath)):
        xpath[index]['searchDepth'] = search_depth_list[index]

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
def strategy_xpath(name, class_name, control_type, Xpath):
    try:
        ctrl = find_control_by_xpath(Xpath)

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
                current_ctrl = find_control_by_xpath(Xpath)
                if current_ctrl and current_ctrl.ControlTypeName == control_type and current_ctrl.Name == name and current_ctrl.ClassName == class_name:
                    ctrl = current_ctrl
                    break
                current_ctrl = redirect_control(name, current_ctrl)
                # push_message('重定向当前控件：{}'.format(current_ctrl))

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
            push_message(f'策略 Xpath 滚动定位成功')
            return ctrl
        return False
    except Exception as e:
        push_message(f'策略 Xpath 错误: {str(e)}')
        return False


@timeit
def find_control(LOCATION):
    """
    根据配置列表获取控件
    此版本支持Name、ClassName、Type、Xpath、foundIndex、AutomationId
    :param LOCATION: 定位器字典，包含：
        - WindowName: 父窗口名称
        - Name: 控件名称
        - ClassName: 控件类名
        - Type: 控件类型
        - Xpath: 层级路径
        - foundIndex: 索引位置
    :return: 控件对象或None
    """
    global CURRENT_APP_NAME

    WindowName, Name, ClassName, ControlTypeName, AutomationId, searchDepth, foundIndex = \
        '', '', '', '', '', 0xFFFFFFFF, 0
    # 参数提取与初始化
    WindowName = LOCATION.get('WindowName', '')
    if LOCATION.get('Name'): Name = LOCATION['Name']
    if LOCATION.get('ClassName'): ClassName = LOCATION['ClassName']
    if LOCATION.get('ControlType'): ControlTypeName = LOCATION['ControlType']
    if LOCATION.get('Xpath'): xpath = LOCATION['Xpath']
    if LOCATION.get('AutomationId'): AutomationId = LOCATION['AutomationId']
    searchDepth = LOCATION.get('searchDepth', 0xFFFFFFFF)
    foundIndex = LOCATION.get('foundIndex', 0)
    push_message(f"\033[1;32m{'-' * 20}开始进行控件定位{'-' * 20}\033[0m")
    push_message(f"\033[1;32m控件名称：{Name}\033[0m")
    push_message(f"\033[1;32m控件类名：{ClassName}\033[0m")
    push_message(f"\033[1;32m控件类型：{ControlTypeName}\033[0m")
    push_message(f"\033[1;32m控件所在窗口：{WindowName}\033[0m")

    # 处理Xpath格式（字符串转列表）
    if isinstance(LOCATION.get('Xpath'), str):
        xpath = eval(LOCATION['Xpath'])
    else:
        xpath = LOCATION.get('Xpath', [])
    if 'WorkerW' == xpath[0].get('ClassName'):
        uiautomation.SendKeys('{Win}d', 0, 0)
        push_message("桌面已激活")
    if WindowName and WindowName != CURRENT_APP_NAME:
        # 激活窗口
        if WindowName == '任务栏' or xpath[0].get('Name') == '任务栏' or xpath[0].get('ClassName') == 'Shell_TrayWnd':
            push_message(f"Windos任务栏无需激活")
            return True
        if set_top_window(window_title=WindowName):
            CURRENT_APP_NAME = WindowName

    start_time = int(datetime.now().timestamp() * 1000)
    control = strategy_xpath(Name, ClassName, ControlTypeName, xpath)
    if control:
        current_time_ms = int(datetime.now().timestamp() * 1000)
        push_message(f"[计时] strategy_dictionary 执行耗时: {current_time_ms - start_time}ms")
        return control
    start_time = int(datetime.now().timestamp() * 1000)
    # 精准匹配
    if xpath:
        window = strategy_dictionary(xpath[0])
        if window:
            control = strategy_dictionary(xpath[-1], window)
            if control:
                current_time_ms = int(datetime.now().timestamp() * 1000)
                push_message(f"[计时] strategy_dictionary 执行耗时: {current_time_ms - start_time}ms")
                return control
        push_message(f"\n所有定位策略均失败，未找到控件")
        return False
    else:
        window_dictionary = {'ControlType': 'WindowControl', 'Name': WindowName}
        pane_dictionary = {'ControlType': 'PaneControl', 'Name': WindowName}
        # control_dictionary = {'ControlType': ControlTypeName, 'Name': Name,
        #                       'ClassName': ClassName, 'AutomationId': AutomationId,
        #                       'searchDepth':searchDepth,'foundIndex': foundIndex}
        control_dictionary = {'ControlType': ControlTypeName, 'Name': Name,
                              'ClassName': ClassName, 'AutomationId': AutomationId}
        window = strategy_dictionary(window_dictionary)
        if window:
            control = strategy_dictionary(control_dictionary, window)
        else:
            pane = strategy_dictionary(pane_dictionary)
            if pane:
                control = strategy_dictionary(control_dictionary, pane)
        current_time_ms = int(datetime.now().timestamp() * 1000)
        push_message(f"[计时] find_control_dictionary 执行耗时: {current_time_ms - start_time}ms")
        if control:
            return control
        else:
            push_message(f"\n所有定位策略均失败，未找到控件")
            return False


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
        'ControlType': CONTROL_TYPE_IDS[dictionary['ControlType']],
        'Name': None if dictionary.get('Name') == '' else dictionary.get('Name', None),
        'ClassName': None if dictionary.get('ClassName') == '' else dictionary.get('ClassName', None),
        'AutomationId': None if dictionary.get('AutomationId') == '' else dictionary.get('AutomationId', None),
        # 'searchDepth': dictionary.get('searchDepth', 0xFFFFFFFF),
        # 'foundIndex': dictionary.get('foundIndex', 1)
    }
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
    ControlType = LOCATION.get('Type', '')
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
