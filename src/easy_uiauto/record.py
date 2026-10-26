# -*- coding: utf-8 -*-
# @Name:    record.py
# @Author:  tang
# @Date:    2025/9/23-9:55
# @depict:  录制线程
import threading
import time
import tkinter as tk
from queue import Empty, Queue

from comtypes import CoInitialize, CoUninitialize
from pynput import keyboard, mouse
from pynput.mouse import Button

from .draw import ScreenLineBox
from .utils import (
    INPUT_VK_KEY,
    MODIFIER_VK,
    VK_KEY_NAME,
    format_action_data_by_xpath,
    get_control_info,
    push_message,
)

RUN_TIME = time.strftime("%Y%m%d%H%M%S", time.localtime())


# ========== 录制线程 ==========
class RecordThread(threading.Thread):
    """
    录制动作线程；控件高亮由独立 Tk 线程通过队列更新。
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
        self.last_click_time = 0
        self.last_click_position = (0, 0)
        self.dragging = False
        self.last_input = None
        self.last_input_time = 0
        self.input_merge_interval = 3  # 秒

        self._state_lock = threading.RLock()
        self._stopped = False
        self._stop_complete = threading.Event()
        self.started = threading.Event()
        self.run_error = None
        self.cleanup_errors = []
        self.ui_error = None
        self.ui_cmd_queue = Queue()
        self.ui_thread = None
        self.ui_ready = threading.Event()

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
            with self._state_lock:
                if self._stopped:
                    return
                self.mouse_listener.start()
                self.keyboard_listener.start()
                self.started.set()
            push_message(f"\033[1;32m{'录制开始'.center(52, '-')}\033[0m")
            while self.running:
                time.sleep(0.05)
        except Exception as error:
            self.run_error = error
            push_message(f"录制启动或运行错误: {error}")
        finally:
            self.started.set()
            try:
                self.stop()
            except Exception as error:
                self.run_error = self.run_error or error
                push_message(f"录制清理错误: {error}")

    # Keep Tk in one dedicated thread and pass UIA controls through an in-process
    # queue. This matches the working recorder behavior and avoids resolving a
    # fresh XPath in a separate interpreter for every selected control.
    def start_ui_thread(self):
        def ui_thread_func():
            root = None
            box = None
            pump = None
            try:
                root = tk.Tk()
                root.withdraw()
                box = ScreenLineBox(
                    root=root, control=None, mode="track",
                    interval_ms=50, topmost=True,
                )

                def pump():
                    try:
                        while True:
                            command = self.ui_cmd_queue.get_nowait()
                            name = command.get("cmd")
                            if name == "set_control":
                                try:
                                    box.set_control(command.get("control"))
                                except Exception as error:
                                    self.ui_error = str(error)
                                    push_message(f"高亮控件更新失败: {error}")
                            elif name == "set_style":
                                box.set_style(
                                    color=command.get("color"),
                                    line_width=command.get("line_width"),
                                    alpha=command.get("alpha"),
                                )
                            elif name == "close":
                                box.destroy()
                                return
                    except Empty:
                        pass
                    root.after(16, pump)

                root.after(0, pump)
                self.ui_ready.set()
                root.mainloop()
            except Exception as error:
                self.ui_error = str(error)
                self.ui_ready.set()
            finally:
                if root is not None:
                    try:
                        root.destroy()
                    except Exception:
                        pass
                # Break Tk's callback-closure cycle on the thread that owns Tcl.
                pump = None
                box = None
                root = None

        self.ui_thread = threading.Thread(
            target=ui_thread_func, name="easy-uiauto-record-ui", daemon=True,
        )
        self.ui_thread.start()
        if not self.ui_ready.wait(timeout=5):
            raise TimeoutError("录制高亮窗口启动超时")
        if self.ui_error:
            raise RuntimeError(f"录制高亮窗口启动失败: {self.ui_error}")

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
                return
            self.running = False
            self.flush_last_input()
            self._stopped = True
        try:
            for listener in (self.mouse_listener, self.keyboard_listener):
                # pynput's join raises before start; partial startup is expected.
                if listener.ident is None:
                    continue
                try:
                    listener.stop()
                    if listener.ident != threading.get_ident():
                        listener.join(timeout=2)
                        if listener.is_alive():
                            self.cleanup_errors.append("监听器未在超时内退出")
                except Exception as error:
                    self.cleanup_errors.append(str(error))
            push_message(f"\033[1;32m{'开始推送动作信息'.center(49, '-')}\033[0m")
            for test_id, action_data in enumerate(self.actions_data, 1):
                action_data["TEST_ID"] = str(test_id)
                if self.action_callback:
                    self.action_callback(action_data)
            if self.close_callback:
                self.close_callback()
        except Exception as error:
            self.run_error = self.run_error or error
            raise
        finally:
            self.ui_cmd_queue.put({"cmd": "close"})
            if self.ui_thread and self.ui_thread.ident != threading.get_ident():
                self.ui_thread.join(timeout=3)
                if self.ui_thread.is_alive():
                    self.cleanup_errors.append("高亮窗口线程未在超时内退出")
            self._stop_complete.set()

    def flush_last_input(self):
        if self.last_input:
            value = self.last_input["value"]
            action_data = format_action_data_by_xpath(action_name="输入文本",
                                                      test_id=self.test_id,
                                                      Xpath=[],
                                                      PARAMETERS={"输入文本": value},
                                                      img='')
            self.actions_data.append(action_data)
            self.last_input = None

    # —— 将控件对象送回同进程的 Tk 线程跟踪，避免跨进程 XPath 重查 ——
    def track_control(self, control):
        """锁定并跟踪指定控件"""
        if control and getattr(control, "ClassName", None) == 'TkChild':
            # 与原逻辑保持一致的过滤
            print("性能不足，此动作无效")
            return

        try:
            self.ui_cmd_queue.put({"cmd": "set_control", "control": control})
            try:
                push_message(f"锁定控件: {control.Name} ({control.ControlTypeName})")
            except Exception:
                push_message("锁定控件")
        except Exception:
            pass

    def on_click(self, x, y, button, pressed):
        if self._stopped:
            return
        self.flush_last_input()
        current_time = time.time()
        CoInitialize()
        if pressed:
            Xpath, control = get_control_info(x, y)
            if control:
                # 锁定并跟踪该控件（不再 hide/show）
                self.track_control(control)
            if button == Button.left:
                if not self.dragging:
                    if len(Xpath) == 0:
                        # 获取失败，重新获取
                        Xpath, control = get_control_info(x, y)
                        action_data = format_action_data_by_xpath(action_name="点击",
                                                                  test_id=self.test_id,
                                                                  Xpath=Xpath,
                                                                  PARAMETERS={"x": -1, "y": -1},
                                                                  img='')
                        self.actions_data.append(action_data)
                    elif self.last_click_time and current_time - self.last_click_time < 0.2:
                        self.actions_data.pop(-1)
                        action_data = format_action_data_by_xpath(action_name="双击",
                                                                  test_id=self.test_id,
                                                                  Xpath=Xpath,
                                                                  PARAMETERS={"x": -1, "y": -1},
                                                                  img='')
                        self.actions_data.append(action_data)
                    else:
                        action_data = format_action_data_by_xpath(action_name="点击",
                                                                  test_id=self.test_id,
                                                                  Xpath=Xpath,
                                                                  PARAMETERS={"x": -1, "y": -1},
                                                                  img='')
                        self.actions_data.append(action_data)
            elif button == Button.right:
                action_data = format_action_data_by_xpath(action_name="右击",
                                                          test_id=self.test_id,
                                                          Xpath=Xpath,
                                                          PARAMETERS={"x": -1, "y": -1},
                                                          img='')
                self.actions_data.append(action_data)
            elif button == Button.middle:
                action_data = format_action_data_by_xpath(action_name="中击",
                                                          test_id=self.test_id,
                                                          Xpath=Xpath,
                                                          PARAMETERS={"x": -1, "y": -1},
                                                          img='')
                self.actions_data.append(action_data)
        else:
            if self.last_click_time and current_time - self.last_click_time > 0.2 and (
                    x, y) != self.last_click_position:
                src_xpath, src_control = get_control_info(self.last_click_position[0], self.last_click_position[1])
                self.track_control(src_control)
                dest_xpath, dest_control = get_control_info(x, y)
                self.track_control(dest_control)
                if len(dest_xpath) == 0:
                    drag_parameters = {}
                else:
                    drag_parameters = {
                        "目的控件父窗口名称": f"{dest_xpath[0].get('Name', '')}",
                        "目的控件Name": f"{dest_xpath[-1].get('Name', '')}",
                        "目的控件ClassName": f"{dest_xpath[-1].get('ClassName', '')}",
                        "目的控件ControlType": f"{dest_xpath[-1].get('ControlType')}",
                        "目的控件foundIndex": f"{dest_xpath[-1].get('foundIndex', '1')}",
                        '目的控件AutomationId': f"{dest_xpath[-1].get('AutomationId', '')}",
                        "目的控件Xpath": f"{dest_xpath}"
                    }

                if self.actions_data:
                    self.actions_data.pop(-1)
                action_data = format_action_data_by_xpath(action_name="拖拽",
                                                          test_id=self.test_id,
                                                          Xpath=src_xpath,
                                                          PARAMETERS=drag_parameters,
                                                          img='')
                self.actions_data.append(action_data)

        self.last_click_time = current_time
        self.last_click_position = (x, y)
        CoUninitialize()

    def on_scroll(self, x, y, dx, dy):
        pass

    def _pressed_vk_numbers(self):
        return [vk for vk, name in VK_KEY_NAME.items() if name in self.pressed_keys]

    def on_press(self, key):
        if self._stopped:
            return
        try:
            if key == keyboard.Key.esc:
                self.stop()
                return

            try:
                vk_number = key.vk
            except AttributeError:
                vk_number = key.value.vk

            key_name = VK_KEY_NAME.get(vk_number, str(key))

            # 记录按下的键（去重）
            if key_name not in self.pressed_keys:
                self.pressed_keys.add(key_name)
                self.pressed_key_sequence.append(key_name)

            # 判断组合键是否包含修饰键 vk
            if len(self.pressed_keys) > 1 and any(vk in MODIFIER_VK for vk in self._pressed_vk_numbers()):
                self.flush_last_input()
                combo = '+'.join(self.pressed_key_sequence)  # 不排序，保持原始顺序
                action_data = format_action_data_by_xpath(action_name="组合键",
                                                          test_id=self.test_id,
                                                          Xpath=[],
                                                          PARAMETERS={"组合键": combo},
                                                          img='')
                self.actions_data.append(action_data)
                self.combo_handled = True
        except Exception as e:
            push_message(f"键盘按下事件处理错误: {e}")

    def on_release(self, key):
        if self._stopped:
            return
        try:
            try:
                vk_number = key.vk
            except AttributeError:
                vk_number = key.value.vk

            key_name = VK_KEY_NAME.get(vk_number, str(key))

            if key_name in self.pressed_keys:
                self.pressed_keys.remove(key_name)
            if key_name in self.pressed_key_sequence:
                self.pressed_key_sequence.remove(key_name)

            if self.combo_handled:
                if len(self.pressed_keys) == 0:
                    self.combo_handled = False
                    self.pressed_key_sequence.clear()
                return

            if vk_number in INPUT_VK_KEY:
                now = time.time()
                if self.last_input and now - self.last_input_time < self.input_merge_interval:
                    self.last_input["value"] += key_name
                else:
                    self.flush_last_input()
                    self.last_input = {"value": key_name}
                self.last_input_time = now
            else:
                self.flush_last_input()
                action_data = format_action_data_by_xpath(action_name="键盘点击",
                                                          test_id=self.test_id,
                                                          Xpath=[],
                                                          PARAMETERS={"键盘按键": key_name},
                                                          img='')
                self.actions_data.append(action_data)
        except Exception as e:
            push_message(f"键盘释放事件处理错误: {e}")


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
        with open(f"Record{RUN_TIME}.py", "a", encoding="utf-8") as f:
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
        record_thread.join(timeout=5)
    except Exception as e:
        push_message(f"\033[1;31m录制出错 :{e}\033[0m")
        record_thread.stop()
        raise

    record_thread._stop_complete.wait(timeout=5)
    if record_thread.run_error is not None:
        raise RuntimeError(f"录制失败: {record_thread.run_error}") from record_thread.run_error
    if record_thread.cleanup_errors:
        raise RuntimeError("录制清理失败: " + "; ".join(record_thread.cleanup_errors))
    return record_thread.actions_data


if __name__ == '__main__':
    run_record(True)

