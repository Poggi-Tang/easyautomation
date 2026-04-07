# -*- coding: utf-8 -*-
# @Name:    record.py
# @Author:  tang
# @Date:    2025/9/23-9:55
# @depict:  录制线程
import threading
import time
from queue import Queue, Empty
import tkinter as tk
from comtypes import CoInitialize, CoUninitialize
from pynput import mouse, keyboard
from pynput.mouse import Button
from draw import ScreenLineBox
from utils import VK_KEY_NAME, INPUT_VK_KEY, format_action_data_by_xpath, push_message, \
    get_control_info, MODIFIER_VK

RUN_TIME = time.strftime("%Y%m%d%H%M%S", time.localtime())


# ========== 录制线程 ==========
class RecordThread(threading.Thread):
    """
    录制动作线程 —— 已将绘制窗口替换为 ScreenLineBox (独立 UI 线程)
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

        # —— 新的 UI 线程 & 指令队列 ——
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
        # 启动 UI 线程（Tk mainloop + ScreenLineBox）
        self.start_ui_thread()

        self.mouse_listener.start()
        self.keyboard_listener.start()

        push_message(f"\033[1;32m{'录制开始'.center(52, '-')}\033[0m")

        # 主循环 —— 这里基本只保持存活；UI 由独立线程处理
        while self.running:
            time.sleep(0.05)

    # —— UI 线程：创建 Tk root 与 ScreenLineBox，并按队列处理命令 ——
    def start_ui_thread(self):
        def ui_thread_func():
            root = tk.Tk()
            root.withdraw()

            # 初始不指定控件，track 模式；100ms 刷新
            # 注意：使用默认样式：红色、线宽2、alpha=1.0
            box = ScreenLineBox(root=root, control=None, mode="track",
                                interval_ms=50, topmost=True)

            # 处理命令队列
            def pump():
                try:
                    while True:
                        cmd = self.ui_cmd_queue.get_nowait()
                        name = cmd.get("cmd")
                        if name == "set_control":
                            box.set_control(cmd.get("control"))
                        elif name == "set_style":
                            box.set_style(color=cmd.get("color"),
                                          line_width=cmd.get("line_width"),
                                          alpha=cmd.get("alpha"))
                        elif name == "close":
                            try:
                                box.destroy()
                            except Exception:
                                pass
                            root.after(10, root.destroy)
                            return
                except Empty:
                    pass
                root.after(16, pump)  # ~60fps 轮询队列

            pump()
            self.ui_ready.set()
            root.mainloop()

        self.ui_thread = threading.Thread(target=ui_thread_func, daemon=True)
        self.ui_thread.start()
        # 等 UI 就绪
        self.ui_ready.wait(timeout=5.0)

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
        self.flush_last_input()
        self.mouse_listener.stop()
        self.keyboard_listener.stop()
        push_message(f"\033[1;32m{'开始推送动作信息'.center(49, '-')}\033[0m")

        test_id = 0
        for action_data in self.actions_data:
            test_id += 1
            action_data["TEST_ID"] = str(test_id)
            if self.action_callback:
                self.action_callback(action_data)

        if self.close_callback:
            self.close_callback()

        self.running = False

        # 关闭 UI
        try:
            self.ui_cmd_queue.put({"cmd": "close"})
        except Exception:
            pass

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

    # —— 替换原 frame/tracker：用 UI 指令给 ScreenLineBox 设置/跟踪控件 ——
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
            except:
                push_message("锁定控件")
        except Exception:
            pass

    def on_click(self, x, y, button, pressed):
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
    push_message(f"push_message(kl_run_action({action_name}))")
    with open(f"Record{RUN_TIME}.py", "a", encoding="utf-8") as f:
        f.write(f'\n{action_name} = {action_data}\n'
                f'push_message(kl_run_action({action_name}))\n')
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
                    f'from easyuiauto.ui_ctrl import kl_run_action\n'
                    f'from easyuiauto.utils import push_message,compile_controls\n'
                    f'# compile_controls(max_depth=1)\n')
    record_help()
    # 创建录制线程
    record_thread = RecordThread(
        action_callback=on_action_callback,
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


if __name__ == '__main__':
    run_record(True)
