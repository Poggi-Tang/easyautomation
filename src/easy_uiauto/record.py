# -*- coding: utf-8 -*-
# @Name:    record.py
# @Author:  tang
# @Date:    2025/9/23-9:55
# @depict:  录制线程
import threading
import time
from math import hypot

from pynput import keyboard, mouse
from pynput.mouse import Button

from .highlight import HighlightProcess
from .utils import (
    INPUT_VK_KEY,
    MODIFIER_VK,
    VK_KEY_NAME,
    format_action_data_by_xpath,
    get_control_info,
    get_focused_control_info,
    push_message,
)

RUN_TIME = time.strftime("%Y%m%d%H%M%S", time.localtime())
SHORTCUT_MODIFIER_VK = {91, 92, 162, 163, 164, 165}


# ========== 录制线程 ==========
class RecordThread(threading.Thread):
    """
    录制动作线程；高亮窗口运行在独立解释器进程中。
    """

    def __init__(self, action_callback=None, close_callback=None):
        super().__init__()
        self.test_id = 0
        self.combo_handled = False
        self.daemon = True
        self.action_callback = action_callback
        self.close_callback = close_callback
        self.pressed_keys = set()
        self.pressed_key_sequence = []  # 顺序记录组合键顺序
        self.actions_data = []
        self.running = True
        self._state_lock = threading.RLock()
        self._stopped = False
        self._stop_complete = threading.Event()
        self.started = threading.Event()
        self.run_error = None
        self.cleanup_errors = []
        self._mouse_press = None
        self._last_left_click = None
        self._handled_combo_keys = set()
        self.drag_threshold = 4
        self.double_click_interval = 0.5
        self.last_click_time = 0
        self.last_click_position = (0, 0)
        self.dragging = False
        self.last_input = None
        self.last_input_time = 0.0
        self.input_merge_interval = 3  # 秒

        # Tk 必须在创建它的主线程中销毁，因此高亮运行在独立子进程中。
        self.ui_ready = threading.Event()
        self.ui_error = None
        self._ui_highlight = HighlightProcess(
            name="easy-uiauto-record-highlight",
        )

        # 监听器
        self.mouse_listener = mouse.Listener(
            on_click=self.on_click,
            on_scroll=self.on_scroll
        )
        self.keyboard_listener = keyboard.Listener(
            on_press=self.on_press,
            on_release=self.on_release
        )

    def run(self):
        try:
            self.start_ui_thread()
            self.mouse_listener.start()
            self.keyboard_listener.start()
            self.started.set()
            push_message(f"\033[1;32m{'录制开始'.center(52, '-')}\033[0m")
            while self.running:
                time.sleep(0.05)
        except Exception as exc:
            self.run_error = exc
            self.started.set()
            push_message(f"录制启动或运行错误: {exc}")
        finally:
            if not self._stopped:
                self.stop()
            self._stop_complete.wait(timeout=10)

    # 保留原方法名以兼容调用方；实现已改为独立 UI 进程。
    def start_ui_thread(self):
        try:
            self._ui_highlight.start(timeout=5)
        except Exception as exc:
            self.ui_error = str(exc)
            self.ui_ready.set()
            raise
        self.ui_ready.set()

    @property
    def ui_process(self):
        return self._ui_highlight.process

    def _send_ui_command(self, command):
        if command.get("cmd") == "set_style":
            command = {
                "cmd": "style",
                "style": {
                    key: command.get(key)
                    for key in ("color", "line_width", "alpha")
                    if command.get(key) is not None
                },
            }
        self._ui_highlight.send(command)

    def _stop_ui_process(self):
        status = self._ui_highlight.stop(timeout=2)
        self.ui_error = status.get("error")
        if status.get("forced_termination"):
            self.cleanup_errors.append("高亮窗口进程未在超时内退出，已强制终止")

    # ESC 退出时
    def on_esc_exit(self, box_track, root):
        try:
            box_track.destroy()
        except Exception:
            pass
        try:
            root.destroy()
        except Exception:
            pass

    def stop(self):
        with self._state_lock:
            if self._stopped:
                already_stopped = True
            else:
                already_stopped = False
                self.running = False
                self.flush_last_input()
                self._stopped = True
        if already_stopped:
            self._stop_complete.wait(timeout=10)
            return
        try:
            self.mouse_listener.stop()
        except Exception as exc:
            self.cleanup_errors.append(f"鼠标监听器停止失败: {exc}")
        try:
            self.keyboard_listener.stop()
        except Exception as exc:
            self.cleanup_errors.append(f"键盘监听器停止失败: {exc}")
        for listener in (self.mouse_listener, self.keyboard_listener):
            if getattr(listener, "ident", None) == threading.get_ident():
                continue
            try:
                listener.join(timeout=2)
            except Exception as exc:
                push_message(f"监听器停止等待错误: {exc}")
                self.cleanup_errors.append(str(exc))
            try:
                if listener.is_alive():
                    self.cleanup_errors.append("监听器未在超时内退出")
            except Exception as exc:
                self.cleanup_errors.append(str(exc))
        push_message(f"\033[1;32m{'开始推送动作信息'.center(49, '-')}\033[0m")

        try:
            for test_id, action_data in enumerate(self.actions_data, 1):
                action_data["TEST_ID"] = str(test_id)
                if self.action_callback:
                    try:
                        self.action_callback(action_data)
                    except Exception as exc:
                        push_message(f"动作回调错误: {exc}")

            if self.close_callback:
                try:
                    self.close_callback()
                except Exception as exc:
                    push_message(f"关闭回调错误: {exc}")
        finally:
            try:
                try:
                    self._stop_ui_process()
                except Exception as exc:
                    self.cleanup_errors.append(f"高亮窗口进程停止失败: {exc}")
            finally:
                self._stop_complete.set()

    def flush_last_input(self):
        with self._state_lock:
            if self.last_input and not self._stopped:
                value = self.last_input["value"]
                action_data = format_action_data_by_xpath(action_name="输入文本",
                                                          test_id=self.test_id,
                                                          Xpath=self.last_input.get("xpath", []),
                                                          PARAMETERS={"输入文本": value},
                                                          img='')
                self.actions_data.append(action_data)
                self.last_input = None

    def _append_action(self, action_data):
        with self._state_lock:
            if self._stopped:
                return
            self.actions_data.append(action_data)

    # —— 替换原 frame/tracker：用 UI 指令给 ScreenLineBox 设置/跟踪控件 ——
    def track_control(self, xpath):
        """锁定并跟踪指定控件"""
        if not xpath:
            return

        try:
            self._send_ui_command({"cmd": "set_xpath", "xpath": xpath})
            target = xpath[-1]
            push_message(
                f"锁定控件: {target.get('Name', '')} ({target.get('ControlType', '')})"
            )
        except Exception as exc:
            push_message(f"高亮窗口命令错误: {exc}")

    def on_click(self, x, y, button, pressed):
        try:
            with self._state_lock:
                if self._stopped:
                    return
            self.flush_last_input()
            current_time = time.time()
            xpath, control = get_control_info(x, y)
            if control and xpath:
                self.track_control(xpath)

            if pressed:
                offset = self._relative_offset(control, x, y)
                with self._state_lock:
                    self._mouse_press = {
                        "button": button,
                        "position": (x, y),
                        "time": current_time,
                        "xpath": xpath,
                        "offset": offset,
                    }
                return

            with self._state_lock:
                press = self._mouse_press
                self._mouse_press = None
            if not press or press["button"] != button:
                return

            start_x, start_y = press["position"]
            distance = hypot(x - start_x, y - start_y)
            if button == Button.left and distance >= self.drag_threshold:
                parameters = self._destination_parameters(
                    xpath,
                    control,
                    x,
                    y,
                    source_offset=press.get("offset"),
                )
                if press["xpath"] and parameters:
                    self._append_action(
                        format_action_data_by_xpath(
                            action_name="拖拽",
                            test_id=self.test_id,
                            Xpath=press["xpath"],
                            PARAMETERS=parameters,
                            img="",
                        )
                    )
                else:
                    push_message("拖拽事件缺少源控件或目的控件定位信息，已忽略")
                self._last_left_click = None
                return

            action_name = {
                Button.left: "点击",
                Button.right: "右击",
                Button.middle: "中击",
            }.get(button)
            if action_name is None:
                return

            action_data = format_action_data_by_xpath(
                action_name=action_name,
                test_id=self.test_id,
                Xpath=press["xpath"],
                PARAMETERS={
                    "x": press["offset"][0] if press.get("offset") else -1,
                    "y": press["offset"][1] if press.get("offset") else -1,
                },
                img="",
            )
            with self._state_lock:
                if self._stopped:
                    return
                if button == Button.left and self._is_double_click(
                    current_time, press["position"], press["xpath"]
                ):
                    previous = self._last_left_click
                    if previous is None:
                        raise RuntimeError("双击状态不一致")
                    previous_index = previous["index"]
                    action_data["ACTION"] = "双击"
                    self.actions_data[previous_index] = action_data
                    self._last_left_click = None
                else:
                    self.actions_data.append(action_data)
                    if button == Button.left:
                        self._last_left_click = {
                            "time": current_time,
                            "position": press["position"],
                            "xpath": press["xpath"],
                            "index": len(self.actions_data) - 1,
                        }
                    else:
                        self._last_left_click = None
        except Exception as exc:
            push_message(f"鼠标事件处理错误: {exc}")

    def _is_double_click(self, current_time, position, xpath):
        previous = self._last_left_click
        if not previous:
            return False
        return (
            current_time - previous["time"] <= self.double_click_interval
            and hypot(
                position[0] - previous["position"][0],
                position[1] - previous["position"][1],
            ) < self.drag_threshold
            and xpath == previous["xpath"]
            and previous["index"] == len(self.actions_data) - 1
            and self.actions_data[previous["index"]]["ACTION"] == "点击"
        )

    @staticmethod
    def _relative_offset(control, x, y):
        if control is None:
            return None
        try:
            rect = control.BoundingRectangle
            return int(x - rect.left), int(y - rect.top)
        except Exception:
            return None

    @classmethod
    def _destination_parameters(cls, xpath, control, x, y, source_offset=None):
        if not xpath:
            return {}
        destination_offset = cls._relative_offset(control, x, y)
        return {
            "x": source_offset[0] if source_offset else -1,
            "y": source_offset[1] if source_offset else -1,
            "目的控件父窗口名称": xpath[0].get("Name", ""),
            "目的控件Name": xpath[-1].get("Name", ""),
            "目的控件ClassName": xpath[-1].get("ClassName", ""),
            "目的控件ControlType": xpath[-1].get("ControlType", ""),
            "目的控件foundIndex": xpath[-1].get("foundIndex", 1),
            "目的控件AutomationId": xpath[-1].get("AutomationId", ""),
            "目的控件Xpath": xpath,
            "目的控件x": destination_offset[0] if destination_offset else -1,
            "目的控件y": destination_offset[1] if destination_offset else -1,
        }

    def on_scroll(self, x, y, dx, dy):
        if not dy and not dx:
            return
        try:
            with self._state_lock:
                if self._stopped:
                    return
                self._last_left_click = None
            xpath, _ = get_control_info(x, y)
            self.flush_last_input()
            if dy:
                distance = abs(int(dy))
                direction = "up" if dy > 0 else "down"
            else:
                distance = abs(int(dx))
                direction = "right" if dx > 0 else "left"
            self._append_action(
                format_action_data_by_xpath(
                    action_name="滚动",
                    test_id=self.test_id,
                    Xpath=xpath,
                    PARAMETERS={
                        "滚动距离": distance,
                        "滚动方向": direction,
                    },
                    img="",
                )
            )
        except Exception as exc:
            push_message(f"滚轮事件处理错误: {exc}")

    def _pressed_vk_numbers(self):
        return [vk for vk, name in VK_KEY_NAME.items() if name in self.pressed_keys]

    def on_press(self, key):
        try:
            if key == keyboard.Key.esc:
                self.stop()
                return

            vk_number, key_name = self._key_details(key)
            with self._state_lock:
                if self._stopped:
                    return
                self._last_left_click = None
                newly_pressed = key_name not in self.pressed_keys
                if key_name not in self.pressed_keys:
                    self.pressed_keys.add(key_name)
                    self.pressed_key_sequence.append(key_name)
                pressed_vk = self._pressed_vk_numbers()
                modifier_pressed = (
                    "alt_gr" not in self.pressed_keys
                    and any(vk in SHORTCUT_MODIFIER_VK for vk in pressed_vk)
                )
                is_modifier = vk_number in MODIFIER_VK
                if (
                    newly_pressed
                    and modifier_pressed
                    and not is_modifier
                    and len(self.pressed_keys) > 1
                    and key_name not in self._handled_combo_keys
                ):
                    self.flush_last_input()
                    combo = "+".join(self.pressed_key_sequence)
                    self.actions_data.append(
                        format_action_data_by_xpath(
                            action_name="组合键",
                            test_id=self.test_id,
                            Xpath=self._focused_xpath(),
                            PARAMETERS={"组合键": combo},
                            img="",
                        )
                    )
                    self.combo_handled = True
                    self._handled_combo_keys.add(key_name)
        except Exception as e:
            push_message(f"键盘按下事件处理错误: {e}")

    def on_release(self, key):
        try:
            vk_number, key_name = self._key_details(key)
            with self._state_lock:
                if self._stopped:
                    return
                if key_name in self.pressed_keys:
                    self.pressed_keys.remove(key_name)
                if key_name in self.pressed_key_sequence:
                    self.pressed_key_sequence.remove(key_name)

                if key_name in self._handled_combo_keys:
                    self._handled_combo_keys.discard(key_name)
                    if not self._handled_combo_keys:
                        self.combo_handled = False
                    return

                text = self._key_text(key, vk_number)
                if text is not None:
                    now = time.time()
                    focused_xpath = self._focused_xpath()
                    if (
                        self.last_input
                        and self.last_input.get("xpath", []) == focused_xpath
                        and now - self.last_input_time < self.input_merge_interval
                    ):
                        self.last_input["value"] += text
                    else:
                        self.flush_last_input()
                        self.last_input = {"value": text, "xpath": focused_xpath}
                    self.last_input_time = now
                elif vk_number not in MODIFIER_VK:
                    self.flush_last_input()
                    self.actions_data.append(
                        format_action_data_by_xpath(
                            action_name="键盘点击",
                            test_id=self.test_id,
                            Xpath=self._focused_xpath(),
                            PARAMETERS={"键盘按键": key_name},
                            img="",
                        )
                    )
        except Exception as e:
            push_message(f"键盘释放事件处理错误: {e}")

    @staticmethod
    def _key_details(key):
        try:
            vk_number = key.vk
        except AttributeError:
            vk_number = key.value.vk
        return vk_number, VK_KEY_NAME.get(vk_number, str(key).replace("Key.", ""))

    @staticmethod
    def _key_text(key, vk_number):
        if vk_number == 32:
            return " "
        char = getattr(key, "char", None)
        if isinstance(char, str) and char:
            return char
        if vk_number in INPUT_VK_KEY:
            return VK_KEY_NAME.get(vk_number)
        return None

    @staticmethod
    def _focused_xpath():
        xpath, _ = get_focused_control_info()
        return xpath


def on_action_callback(action_data):
    """
    录制结束后置执行

    :param action_data: 依次收到的每个动作的信息
        {
            'TEST_ID': ID
            'ACTION': 名称
            'LOCATION': 控件定位器
        }
    """
    global RUN_TIME
    action_name = f"TEST_INFO_{action_data['TEST_ID']}"
    push_message(f'\n{action_name} = {action_data}')
    push_message(f"push_message(run_action({action_name}))")
    with open(f"Record{RUN_TIME}.py", "a", encoding="utf-8") as f:
        f.write(f'\n{action_name} = {action_data}\n'
                f'push_message(run_action({action_name}))\n'
                f'# print(find_control({action_name}["LOCATION"],debug=True))\n'
                )
    return


def on_close_callback():
    """录制结束时的回调"""
    push_message(f"\033[1;32m{'录制结束'.center(52, '-')}\033[0m")


def record_help():
    """录制帮助"""
    push_message(f"\033[1;33m{'控件录制'.center(52)}\033[0m")
    push_message("\033[1;32m支持录制操作:\033[0m")
    push_message(f"\033[1;34m{'鼠标事件'.center(52, '-')}\033[0m")
    push_message("\033[1;37m1. 鼠标点击 (点击控件时会自动锁定并显示方框)\033[0m")
    push_message("\033[1;37m2. 鼠标双击\033[0m")
    push_message("\033[1;37m3. 鼠标右击\033[0m")
    push_message("\033[1;37m4. 鼠标拖拽\033[0m")
    push_message(f"\033[1;34m{'键盘事件'.center(52, '-')}\033[0m")
    push_message("\033[1;37m1. 键盘输入\033[0m")
    push_message("\033[1;37m2. 键盘按键\033[0m")
    push_message("\033[1;37m3. 键盘组合键\033[0m")
    push_message("\033[1;31mESC键: 停止录制\033[0m")


def run_record(write_file=True):
    """
    录制示例
    :param write_file: 将录制数据写入到文件
    :return:
    """
    if write_file:
        global RUN_TIME
        RUN_TIME = (
            time.strftime("%Y%m%d%H%M%S", time.localtime())
            + f"{time.time_ns() % 1_000_000_000:09d}"
        )
        with open(f"Record{RUN_TIME}.py", "x", encoding="utf-8") as f:
            f.write(f'# -*- coding: utf-8 -*-\n'
                    f'# @Name:      {RUN_TIME}.py\n'
                    f'from easy_uiauto import run_action\n'
                    f'from easy_uiauto import push_message, compile_controls,find_control\n'
                    f'# compile_controls(max_depth=1)\n')
    record_help()
    # 创建录制线程
    record_thread = RecordThread(
        action_callback=on_action_callback if write_file else None,
        close_callback=on_close_callback,
    )
    try:
        # 启动录制
        record_thread.start()

        # 等待录制线程结束
        record_thread.join()

    except KeyboardInterrupt:
        push_message("\033[1;31mESC键: 用户中断录制\033[0m")
        record_thread.stop()
    except Exception as e:
        push_message(f"\033[1;31m录制出错 :{e}\033[0m")
        record_thread.stop()
    return record_thread.actions_data


if __name__ == '__main__':
    run_record(True)

