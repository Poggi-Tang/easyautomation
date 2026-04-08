# -*- coding: utf-8 -*-
# @Name:      ui_ctrl.py
# @Author:    tang
# @Date:      2025/9/23-9:55
# @depict:    控制
import tkinter
import threading
import time

import pyautogui
import pyperclip
import uiautomation

from .draw import ScreenLineBox
from .utils import find_control, package_location, get_control_coordinates, correct_ctrl_position, push_message


def _show_thread(control_obj, show_time):
    """在独立线程中创建 Tk 窗口并显示一次性 overlay（不会阻塞调用线程）。"""
    try:
        root = tkinter.Tk()
        root.withdraw()
        ScreenLineBox(root, control_obj, mode="once", show_time=show_time)
        root.mainloop()
    except Exception as e:
        try:
            push_message(f"show_ctrl_area 异常: {e}")
        except Exception:
            pass


def show_ctrl_area(control_obj, show_time=300):
    """非阻塞地显示控件区域：在后台线程启动 Tk 主循环。"""
    try:
        t = threading.Thread(target=_show_thread, args=(control_obj, show_time), daemon=True)
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
    start = time.time()
    last_exc = None
    while time.time() - start <= timeout:
        try:
            control = find_control(LOCATION)
            if control:
                try:
                    correct_x, correct_y = correct_ctrl_position(control)
                except Exception as e:
                    correct_x, correct_y = 0, 0
                    last_exc = e
                x = int(PARAMETERS.get('x', correct_x))
                y = int(PARAMETERS.get('y', correct_y))
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

_message_type = _MESSAGE.OTHER


def get_message_type():
    """
    获取当前消息类型

    :return
        int: 当前消息类型的值
    """
    global _message_type
    return _message_type


def update_message_type(current_message_type):
    """
    更新当前消息类型

    :param
        current_message_type (int): 要设置的消息类型值
    """
    global _message_type
    _message_type = current_message_type


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
            # 对于特定控件类型，使用可点击点
            if control.ControlTypeName in {'ButtonControl', 'MenuItemControl', 'CheckBoxControl'}:
                if not control.IsEnabled:
                    raise Exception('控件不可点击')
                control.Click(x, y)
            else:
                control.Click(x, y)

            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 点击 成功！"

        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 点击 异常：{e}"

        finally:
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

            control = find_control(LOCATION)
            if control:
                x, y = get_pos(control, PARAMETERS)
                uiautomation.PressMouse(x, y)
            else:
                raise Exception("未找到控件！")
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标左键按下 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标左键按下 异常：{e}"
        finally:
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

            result = cls.mouse_move_control(ActionTitle, **LOCATION)
            if result[0] == _MESSAGE.INFO:
                uiautomation.ReleaseMouse()
                MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标左键释放 成功！"
            else:
                raise Exception("未找到控件！")
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标左键释放 异常：{e}"
        finally:
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

            control = find_control(LOCATION)
            if control:
                x, y = get_pos(control, PARAMETERS)
                control.RightClick(x, y)
            else:
                raise Exception("未找到控件！")
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 右键点击 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 右键点击 异常：{e}"
        finally:
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

            control = find_control(LOCATION)
            if control:
                x, y = get_pos(control, PARAMETERS)
                uiautomation.RightPressMouse(x, y)
            else:
                raise Exception("未找到控件！")
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标右键按下 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标右键按下 异常：{e}"
        finally:
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

            result = cls.mouse_move_control(ActionTitle, **LOCATION)
            if result[0] == _MESSAGE.INFO:
                uiautomation.RightReleaseMouse()
                MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标右键释放 成功！"
            else:
                raise Exception("未找到控件！")
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 鼠标右键释放 异常：{e}"
        finally:
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

            control = find_control(LOCATION)
            if control:
                x, y = get_pos(control, PARAMETERS)
                control.MiddleClick(x, y)
            else:
                raise Exception("未找到控件！")
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 中键点击 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 中键点击 异常：{e}"
        finally:
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

            control = find_control(LOCATION)
            if control:
                x, y = get_pos(control, PARAMETERS)
                control.DoubleClick(x, y)
            else:
                raise Exception("未找到控件！")

            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 双击 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 双击 异常：{e}"
        finally:
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
        finally:
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

            control = find_control(LOCATION)
            if control:
                x, y = get_pos(control, PARAMETERS)
                control.MoveCursorToInnerPos(x, y)
            else:
                raise Exception("未找到控件！")
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 移动鼠标到坐标位置 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 移动鼠标到坐标位置 异常：{e}"
        finally:
            update_message_type(_MESSAGE_TYPE)
            return MESSAGE

    @classmethod
    def drag_control_by_control(cls, current_control, target_control):
        current_coord = get_control_coordinates(current_control)
        target_coord = get_control_coordinates(target_control)
        current_coord_x = (current_coord[0] + current_coord[2]) // 2
        current_coord_y = (current_coord[1] + current_coord[3]) // 2
        target_coord_x = (target_coord[0] + target_coord[2]) // 2
        target_coord_y = (target_coord[1] + target_coord[3]) // 2
        uiautomation.DragDrop(current_coord_x, current_coord_y, target_coord_x, target_coord_y)

    @classmethod
    def drag_control(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex, AutomationId, Xpath, Img,
                        PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""
        try:
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
            control = find_control(LOCATION)
            DestCtrl = find_control(DestCarl_Locators)
            cls.drag_control_by_control(control, DestCtrl)

            MESSAGE = f"步骤：{ActionTitle} 执行动作 拖动控件 成功"
        except BaseException as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 执行动作 拖动控件 异常：{e}"
        finally:
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

            control = find_control(LOCATION)
            if control:
                correct_ctrl_position(control)
            control.SendKeys(PARAMETERS["设置文本"])
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 设置文本 成功！"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle} 控件【{Name}】 执行动作 设置文本 异常：{e}"
        finally:
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
            if WindowName:
                control = find_control(LOCATION)
                if control:
                    control.Click(waitTime=0)
            pyperclip.copy(PARAMETERS['输入文本'])
            pyautogui.hotkey('ctrl', 'v', interval=0)
            MESSAGE = f"步骤：{ActionTitle}执行动作 输入 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 输入 异常：{e}"
        finally:
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
            if WindowName:
                control = find_control(LOCATION)
                if control:
                    control.Click(waitTime=0)
            pyautogui.press(PARAMETERS['键盘按键'])
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘操作 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘操作 异常：{e}"
        finally:
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
            if WindowName:
                control = find_control(LOCATION)
                if control:
                    control.Click(waitTime=0)
            pyautogui.keyDown(PARAMETERS['键盘按键'])
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘按下 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘按下 异常：{e}"
        finally:
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
            if WindowName:
                control = find_control(LOCATION)
                if control:
                    control.Click(waitTime=0)
            pyautogui.keyUp(PARAMETERS['键盘按键'])
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘释放 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 键盘释放 异常：{e}"
        finally:
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
            }
            if WindowName:
                control = find_control(LOCATION)
                if control:
                    control.Click(waitTime=0)
            raw_keys = PARAMETERS['组合键'].lower().split('+')
            mapped_keys = [key_mapping.get(k.strip(), k.strip()) for k in raw_keys]
            pyautogui.hotkey(*mapped_keys)
            MESSAGE = f"步骤：{ActionTitle}执行动作 组合键 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 组合键 异常：{e}"
        finally:
            update_message_type(_MESSAGE_TYPE)
            return MESSAGE

    @classmethod
    def activate_window(cls, WindowTitle):
        if pyautogui.getActiveWindowTitle() != WindowTitle:
            windows = pyautogui.getAllWindows()
            for window in windows:
                if window.title == WindowTitle:
                    window.activate()
                    window.maximize()


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
    if isinstance(PARAMETERS,str):
        PARAMETERS = eval(PARAMETERS)
    correct_x, correct_y = correct_ctrl_position(control)
    show_ctrl_area(control)
    x, y = (int(PARAMETERS.get("x", correct_x)), int(PARAMETERS.get("y", correct_y)))
    if x == -1: x = correct_x
    if y == -1: y = correct_y
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
    # 兼容LOCATION中可能存在PARAMETERS
    if "PARAMETERS" not in record_info["LOCATION"]:
        record_info["LOCATION"]["PARAMETERS"] = {}
    func = Execute_Function.get(record_info["ACTION"])
    ActionTitle = generate_action_title(record_info)
    result = func(ActionTitle, **record_info["LOCATION"])
    return result


Execute_Function = {
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
}
