# -*- coding: utf-8 -*-
# @Name:      ui_ctrl.py
# @Author:    tang
# @Date:      2025/9/23-9:55
# @depict:    控制
import tkinter
import threading
import time
from contextlib import contextmanager
from queue import Empty, Queue
from weakref import WeakSet

import uiautomation

from .draw import ScreenLineBox
from .utils import find_control, package_location, correct_ctrl_position, push_message, VK_KEY_NAME


_preview_lock = threading.RLock()
_clipboard_lock = threading.RLock()
_previews = WeakSet()


class ControlPatternUnavailable(RuntimeError):
    """The requested pattern could not be obtained before any write was sent."""


class ControlActionError(RuntimeError):
    """False means no input was sent; None means an attempted input is unconfirmed."""

    def __init__(self, message, *, executed):
        self.executed = executed
        code = "action_failed" if executed is False else "result_unconfirmed"
        suffix = "" if executed is False else "; input was attempted, do not repeat the action"
        super().__init__(f"[{code}] {message}{suffix}")


def _control_rectangle(control):
    rect = control.BoundingRectangle
    return _checked_rectangle(dict(left=rect.left, top=rect.top,
                                   right=rect.right, bottom=rect.bottom))


def _checked_rectangle(rectangle):
    rect = {key: int(rectangle[key]) for key in ("left", "top")}
    rect["right"] = int(rectangle.get("right", rect["left"] + rectangle.get("width", 0)))
    rect["bottom"] = int(rectangle.get("bottom", rect["top"] + rectangle.get("height", 0)))
    rect["width"], rect["height"] = rect["right"] - rect["left"], rect["bottom"] - rect["top"]
    if rect["width"] <= 0 or rect["height"] <= 0:
        raise ValueError("Target rectangle must have positive width and height")
    return rect


class ControlPreview:
    """Own Tk on one thread; exchange only rectangles and acknowledged commands."""

    def __init__(self, rectangles, show_time=300):
        self.rectangles = [_checked_rectangle(rect) for rect in rectangles]
        if not self.rectangles:
            raise ValueError("At least one target rectangle is required")
        self.show_time = max(0, int(show_time))
        self.ready = threading.Event()
        self.closed = threading.Event()
        self._stop = threading.Event()
        self._commands = Queue()
        self.error = None
        self.thread = threading.Thread(target=self._run, name="easy-uiauto-preview", daemon=True)

    def start(self, timeout=5):
        self.thread.start()
        if not self.ready.wait(timeout):
            self._stop.set()
            raise TimeoutError("Control preview did not become ready")
        if self.error:
            raise RuntimeError(f"Control preview failed: {self.error}")
        if self.closed.is_set():
            raise RuntimeError("Control preview closed before it became usable")
        with _preview_lock:
            _previews.add(self)
        return self

    def _run(self):
        root = None
        boxes = []
        after_id = None
        pump = None
        try:
            if self._stop.is_set():
                return
            root = tkinter.Tk()
            root.withdraw()
            for rectangle in self.rectangles:
                box = ScreenLineBox(root=root, mode="manual", control=None)
                boxes.append(box)
                box.set_rect(rectangle)
            root.update()
            deadline = time.monotonic() + self.show_time / 1000 if self.show_time else None

            def pump():
                nonlocal after_id
                after_id = None
                try:
                    if self._stop.is_set() or (deadline and time.monotonic() >= deadline):
                        root.quit()
                        return
                    for _ in range(32):
                        try:
                            visible, acknowledged = self._commands.get_nowait()
                        except Empty:
                            break
                        for box in boxes:
                            box.set_visible(visible)
                        root.update_idletasks()
                        acknowledged.set()
                    after_id = root.after(16, pump)
                except Exception as error:
                    self.error = str(error)
                    root.quit()

            after_id = root.after(0, pump)
            self.ready.set()
            root.mainloop()
        except Exception as error:
            self.error = str(error)
        finally:
            self.ready.set()
            if root is not None:
                if after_id is not None:
                    try:
                        root.after_cancel(after_id)
                    except tkinter.TclError:
                        pass
                for box in boxes:
                    try:
                        box.destroy()
                    except Exception as error:
                        self.error = self.error or str(error)
                try:
                    root.destroy()
                except tkinter.TclError:
                    pass
            boxes.clear()
            box = pump = root = None
            with _preview_lock:
                _previews.discard(self)
            self.closed.set()

    def set_visible(self, visible, timeout=2):
        if self.closed.is_set():
            return
        acknowledged = threading.Event()
        self._commands.put((bool(visible), acknowledged))
        deadline = time.monotonic() + timeout
        while not acknowledged.wait(0.02):
            if self.closed.is_set():
                return
            if time.monotonic() >= deadline:
                raise TimeoutError("Control preview visibility change timed out")

    def close(self, timeout=3):
        self._stop.set()
        if self.thread.ident is not None and self.thread.ident != threading.get_ident():
            if not self.closed.wait(timeout):
                raise TimeoutError("Control preview did not close")
        if self.error:
            raise RuntimeError(f"Control preview failed: {self.error}")

    def summary(self):
        return dict(shown=True, controls=len(self.rectangles), detail="Controller/ScreenLineBox")


def show_ctrl_area(control_obj, show_time=300):
    """Preview a snapshot without passing a UIA object to the Tk thread."""
    try:
        return ControlPreview([_control_rectangle(control_obj)], show_time).start()
    except Exception as e:
        try:
            push_message(f"启动 show_ctrl_area 线程失败: {e}")
        except Exception:
            pass


@contextmanager
def suspend_control_previews():
    """Keep controller previews out of verification screenshots."""
    with _preview_lock:
        previews = list(_previews)
    hidden = []
    try:
        for preview in previews:
            preview.set_visible(False)
            hidden.append(preview)
        yield
    finally:
        for preview in hidden:
            preview.set_visible(True)


def locate_and_prepare(LOCATION, PARAMETERS, timeout=3.0, interval=0.2):
    """尝试在限定时间内定位控件并返回 (control, x, y)。失败则返回 (None, None, None) 并记录日志。

    LOCATION: dict 或其他 find_control 可接受的定位参数
    PARAMETERS: dict，用于覆盖坐标
    timeout: 最大等待秒数
    interval: 重试间隔秒数
    """
    from .mcp.results import ActionError

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
        except ActionError:
            raise
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
    def show_targets(cls, rectangles, show_time=900):
        return ControlPreview(rectangles, show_time).start()

    suspend_highlights = staticmethod(suspend_control_previews)

    @staticmethod
    def _key_code(key):
        if isinstance(key, int):
            return key
        name = str(key).strip().removeprefix("Key.").upper()
        aliases = {
            "CTRL_L": "LCTRL", "CTRL_R": "RCTRL", "CTRLLEFT": "LCTRL", "CTRLRIGHT": "RCTRL",
            "ALT_L": "LALT", "ALT_R": "RALT", "ALTLEFT": "LALT", "ALTRIGHT": "RALT",
            "SHIFT_L": "LSHIFT", "SHIFT_R": "RSHIFT", "SHIFTLEFT": "LSHIFT", "SHIFTRIGHT": "RSHIFT",
            "CMD": "WIN", "CMD_L": "LWIN", "CMD_R": "RWIN", "WINLEFT": "LWIN", "WINRIGHT": "RWIN",
            "CONTROL": "CTRL", "RETURN": "ENTER", "ESCAPE": "ESC", "DEL": "DELETE",
            "CAPS_LOCK": "CAPSLOCK", "PAGE_DOWN": "PAGEDOWN", "PAGE_UP": "PAGEUP",
            "PGDN": "PAGEDOWN", "PGUP": "PAGEUP", "PRINT_SCREEN": "PRINTSCREEN",
            "BACKSPACE": "BACK", "CAPSLOCK": "CAPITAL", "NUM_LOCK": "NUMLOCK",
            "SCROLL_LOCK": "SCROLL", "SCROLLLOCK": "SCROLL", "MENU": "APPS",
        }
        name = aliases.get(name, name)
        code = uiautomation.SpecialKeyNames.get(name)
        if isinstance(code, int):
            return code
        if len(name) == 1 and name.isascii() and name.isalnum():
            return ord(name)
        for virtual_key, recorded_name in VK_KEY_NAME.items():
            if str(recorded_name).upper() == name:
                return virtual_key
        code = getattr(uiautomation.Keys, name if name.startswith("VK_") else f"VK_{name}", None)
        if isinstance(code, int):
            return code
        raise ValueError(f"Unknown keyboard key: {key}")

    @classmethod
    def press_key(cls, key):
        code = cls._key_code(key)
        try:
            uiautomation.PressKey(code, waitTime=0)
        finally:
            uiautomation.ReleaseKey(code, waitTime=0)
        return {"executed": True, "method": "uia-key"}

    @classmethod
    def key_down(cls, key):
        uiautomation.PressKey(cls._key_code(key), waitTime=0)

    @classmethod
    def key_up(cls, key):
        uiautomation.ReleaseKey(cls._key_code(key), waitTime=0)

    @classmethod
    def hotkey(cls, *keys):
        codes = [cls._key_code(key) for key in keys]
        if not codes:
            raise ValueError("At least one hotkey key is required")
        pressed = []
        original_error = None
        try:
            for code in codes:
                pressed.append(code)
                uiautomation.PressKey(code, waitTime=0)
            time.sleep(0.02)
        except BaseException as error:
            original_error = error
            raise
        finally:
            cleanup_errors = []
            for code in reversed(pressed):
                try:
                    uiautomation.ReleaseKey(code, waitTime=0)
                except Exception as error:
                    cleanup_errors.append(str(error))
            if cleanup_errors:
                detail = "Could not release hotkey: " + "; ".join(cleanup_errors)
                if original_error is not None:
                    original_error.add_note(detail)
                else:
                    raise RuntimeError(detail)
        return {"executed": True, "method": "uia-hotkey"}

    @staticmethod
    def get_clipboard_text():
        return uiautomation.GetClipboardText()

    @staticmethod
    def set_clipboard_text(text):
        if not uiautomation.SetClipboardText(str(text)):
            raise RuntimeError("Could not write clipboard text")

    @classmethod
    @contextmanager
    def preserve_clipboard_text(cls):
        with _clipboard_lock:
            previous = cls.get_clipboard_text()
            original_error = None
            try:
                yield
            except BaseException as error:
                original_error = error
                raise
            finally:
                try:
                    cls.set_clipboard_text(previous)
                except Exception as error:
                    if original_error is not None:
                        original_error.add_note(f"Clipboard restoration also failed: {error}")
                    else:
                        raise

    @classmethod
    def activate_window_handle(cls, handle, control=None):
        handle = int(handle or 0)
        if not handle:
            raise ValueError("A window handle is required")
        if uiautomation.IsIconic(handle):
            uiautomation.ShowWindow(handle, uiautomation.SW.Restore)
        if uiautomation.GetForegroundWindow() != handle:
            uiautomation.SetForegroundWindow(handle)
            if uiautomation.GetForegroundWindow() != handle and control is not None:
                control.SetFocus()
        deadline = time.monotonic() + 0.5
        while uiautomation.GetForegroundWindow() != handle:
            if time.monotonic() >= deadline:
                raise RuntimeError("[window_invalid] Could not activate the target window")
            time.sleep(0.02)
        return {"executed": True, "method": "uia-window-activate"}

    @classmethod
    def execute_resolved(cls, action, *, control=None, rectangle=None, x=None, y=None,
                         text="", select_all=False, destination=None, wheel_times=1):
        """Execute an already located target. x/y are inner offsets for controls,
        screen coordinates otherwise. This never performs a second XPath lookup.
        """
        state = {"attempted": False}
        try:
            return cls._execute_resolved(
                action, control=control, rectangle=rectangle, x=x, y=y, text=text,
                select_all=select_all, destination=destination, wheel_times=wheel_times,
                state=state,
            )
        except (ControlPatternUnavailable, ControlActionError):
            raise
        except Exception as error:
            raise ControlActionError(str(error), executed=None if state["attempted"] else False) from error

    @classmethod
    def _execute_resolved(cls, action, *, control, rectangle, x, y, text, select_all,
                          destination, wheel_times, state):
        started = time.monotonic()
        pointer_actions = {"click", "double-click", "right-click", "middle-click", "hover",
                           "left-down", "left-up", "right-down", "right-up", "drag",
                           "wheel-up", "wheel-down"}
        if action in pointer_actions:
            rect = _control_rectangle(control) if control is not None else (
                _checked_rectangle(rectangle) if rectangle is not None else None)
            if control is not None and action != "hover" and not bool(control.IsEnabled):
                raise RuntimeError("[control_disabled] 控件不可点击")
            if control is None and rect is None and (x is None or y is None):
                raise ValueError("Pointer actions require a control, rectangle, or screen point")
            if control is None:
                px = int(x) if x is not None else rect["left"] + rect["width"] // 2
                py = int(y) if y is not None else rect["top"] + rect["height"] // 2
            if action == "drag":
                destination = _checked_rectangle(destination)
            if action in {"wheel-up", "wheel-down"} and int(wheel_times) < 0:
                raise ValueError("wheel_times must not be negative")
            state["attempted"] = True
            if action in {"click", "double-click", "right-click", "middle-click", "hover"}:
                methods = {"click": "Click", "double-click": "DoubleClick", "right-click": "RightClick",
                           "middle-click": "MiddleClick", "hover": "MoveCursorToInnerPos"}
                if control is not None:
                    options = dict(x=x, y=y, simulateMove=False)
                    if action != "hover":
                        options["waitTime"] = 0
                    getattr(control, methods[action])(**options)
                elif action == "double-click":
                    uiautomation.Click(px, py, waitTime=uiautomation.GetDoubleClickTime() / 2000)
                    uiautomation.Click(px, py, waitTime=0)
                elif action == "hover":
                    uiautomation.MoveTo(px, py, moveSpeed=0, waitTime=0)
                else:
                    getattr(uiautomation, methods[action])(px, py, waitTime=0)
            else:
                if control is not None:
                    point = control.MoveCursorToInnerPos(x=x, y=y, simulateMove=False)
                    if point is None:
                        raise RuntimeError("Target has no clickable area")
                    px, py = point
                if action == "drag":
                    target = _checked_rectangle(destination)
                    uiautomation.DragDrop(px, py, target["left"] + target["width"] // 2,
                                         target["top"] + target["height"] // 2, moveSpeed=0, waitTime=0)
                elif action in {"wheel-up", "wheel-down"}:
                    uiautomation.MoveTo(px, py, moveSpeed=0, waitTime=0)
                    method = uiautomation.WheelUp if action == "wheel-up" else uiautomation.WheelDown
                    method(wheelTimes=int(wheel_times), waitTime=0)
                elif action.endswith("-up"):
                    method = uiautomation.ReleaseMouse if action == "left-up" else uiautomation.RightReleaseMouse
                    method(waitTime=0)
                else:
                    method = uiautomation.PressMouse if action == "left-down" else uiautomation.RightPressMouse
                    method(px, py, waitTime=0)
            method_name = ("uia-" if control is not None else "coordinate-") + action
        elif action == "set-value":
            if control is None:
                raise ControlPatternUnavailable("ValuePattern requires a UIA control")
            try:
                pattern = control.GetValuePattern()
                if pattern is None:
                    raise RuntimeError("ValuePattern is unavailable")
            except Exception as error:
                raise ControlPatternUnavailable(str(error)) from error
            state["attempted"] = True
            if not pattern.SetValue(text, waitTime=0):
                raise RuntimeError("UIA SetValue returned failure")
            method_name = "uia-set-value"
        elif action == "send-keys":
            state["attempted"] = True
            if control is not None:
                control.SendKeys(text, interval=0, waitTime=0)
            else:
                uiautomation.SendKeys(text, interval=0, waitTime=0)
            method_name = "uia-send-keys"
        elif action == "paste-text":
            state["attempted"] = True
            if control is not None or rectangle is not None:
                cls.execute_resolved("click", control=control, rectangle=rectangle)
            if select_all:
                cls.hotkey("ctrl", "a")
            with cls.preserve_clipboard_text():
                cls.set_clipboard_text(text)
                cls.hotkey("ctrl", "v")
                time.sleep(0.12)
            method_name = "clipboard-paste"
        else:
            raise ValueError(f"Unsupported controller action: {action}")
        return {"executed": True, "method": method_name,
                "duration_seconds": round(time.monotonic() - started, 4)}

    @classmethod
    def left_click(cls, ActionTitle, WindowName, Name, ClassName, ControlType, foundIndex,
                      AutomationId, Xpath, Img, PARAMETERS):
        _MESSAGE_TYPE = _MESSAGE.INFO
        MESSAGE = ""

        try:
            # 准备控件
            control, x, y = prepare_control(WindowName, Name, ClassName, ControlType,
                                            foundIndex, AutomationId, Xpath, Img, PARAMETERS)
            cls.execute_resolved("click", control=control, x=x, y=y)

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
                cls.execute_resolved("left-down", control=control, x=x, y=y)
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

            control = find_control(LOCATION)
            if control:
                x, y = get_pos(control, PARAMETERS)
                cls.execute_resolved("left-up", control=control, x=x, y=y)
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
                cls.execute_resolved("right-click", control=control, x=x, y=y)
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
                cls.execute_resolved("right-down", control=control, x=x, y=y)
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

            control = find_control(LOCATION)
            if control:
                x, y = get_pos(control, PARAMETERS)
                cls.execute_resolved("right-up", control=control, x=x, y=y)
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
                cls.execute_resolved("middle-click", control=control, x=x, y=y)
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
                cls.execute_resolved("double-click", control=control, x=x, y=y)
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
            cls.execute_resolved("hover", x=x, y=y)
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
                cls.execute_resolved("hover", control=control, x=x, y=y)
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
        return cls.execute_resolved("drag", control=current_control,
                                    destination=_control_rectangle(target_control))

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
            if not control:
                raise RuntimeError("未找到控件！")
            cls.execute_resolved("send-keys", control=control, text=PARAMETERS["设置文本"])
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
                    cls.execute_resolved("click", control=control)
                else:
                    raise RuntimeError("未找到控件！")
            cls.execute_resolved("paste-text", text=PARAMETERS['输入文本'])
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
                    cls.execute_resolved("click", control=control)
                else:
                    raise RuntimeError("未找到控件！")
            cls.press_key(PARAMETERS['键盘按键'])
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
                    cls.execute_resolved("click", control=control)
                else:
                    raise RuntimeError("未找到控件！")
            cls.key_down(PARAMETERS['键盘按键'])
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
                    cls.execute_resolved("click", control=control)
                else:
                    raise RuntimeError("未找到控件！")
            cls.key_up(PARAMETERS['键盘按键'])
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
            if WindowName:
                control = find_control(LOCATION)
                if control:
                    cls.execute_resolved("click", control=control)
                else:
                    raise RuntimeError("未找到控件！")
            keys = PARAMETERS['组合键'].replace('++', '+{PLUS}').split('+')
            cls.hotkey(*(key if key != '{PLUS}' else '+' for key in keys))
            MESSAGE = f"步骤：{ActionTitle}执行动作 组合键 成功"
        except Exception as e:
            _MESSAGE_TYPE = _MESSAGE.ERROR
            MESSAGE = f"步骤：{ActionTitle}执行动作 组合键 异常：{e}"
        finally:
            update_message_type(_MESSAGE_TYPE)
            return MESSAGE

    @classmethod
    def activate_window(cls, WindowTitle):
        from .utils import _native_window_scope

        try:
            window = _native_window_scope(WindowTitle, {})
            cls.activate_window_handle(window.NativeWindowHandle, control=window)
            return True
        except Exception as error:
            push_message(f"窗口激活失败：{error}")
            return False


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
