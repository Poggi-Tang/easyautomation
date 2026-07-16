from __future__ import annotations

from types import SimpleNamespace

from pynput import keyboard
from pynput.mouse import Button

from conftest import FakeControl
from easy_uiauto import record


def make_recorder(monkeypatch):
    recorder = record.RecordThread()
    monkeypatch.setattr(recorder, "track_control", lambda _control: None)
    control = FakeControl()
    xpath = [{"ControlType": "ButtonControl", "Name": "control", "searchDepth": 1}]
    monkeypatch.setattr(record, "get_control_info", lambda _x, _y: (xpath, control))
    monkeypatch.setattr(record, "get_focused_control_info", lambda: (xpath, control))
    return recorder


def test_two_clicks_at_same_position_become_double_click(monkeypatch):
    recorder = make_recorder(monkeypatch)
    times = iter([1.0, 1.05, 1.15, 1.20])
    monkeypatch.setattr(record.time, "time", lambda: next(times))

    recorder.on_click(10, 10, Button.left, True)
    recorder.on_click(10, 10, Button.left, False)
    recorder.on_click(10, 10, Button.left, True)
    recorder.on_click(10, 10, Button.left, False)

    assert [item["ACTION"] for item in recorder.actions_data] == ["双击"]


def test_fast_clicks_at_different_positions_stay_separate(monkeypatch):
    recorder = make_recorder(monkeypatch)
    times = iter([1.0, 1.05, 1.10, 1.15])
    monkeypatch.setattr(record.time, "time", lambda: next(times))

    recorder.on_click(10, 10, Button.left, True)
    recorder.on_click(10, 10, Button.left, False)
    recorder.on_click(30, 30, Button.left, True)
    recorder.on_click(30, 30, Button.left, False)

    assert [item["ACTION"] for item in recorder.actions_data] == ["点击", "点击"]


def test_click_records_relative_control_offset(monkeypatch):
    recorder = make_recorder(monkeypatch)
    monkeypatch.setattr(record.time, "time", lambda: 1.0)

    recorder.on_click(110, 205, Button.left, True)
    recorder.on_click(110, 205, Button.left, False)

    assert recorder.actions_data[0]["LOCATION"]["PARAMETERS"] == {"x": 10, "y": 5}


def test_left_mouse_movement_becomes_drag(monkeypatch):
    recorder = make_recorder(monkeypatch)
    monkeypatch.setattr(record.time, "time", lambda: 1.0)

    recorder.on_click(10, 10, Button.left, True)
    recorder.on_click(40, 40, Button.left, False)

    assert [item["ACTION"] for item in recorder.actions_data] == ["拖拽"]
    parameters = recorder.actions_data[0]["LOCATION"]["PARAMETERS"]
    assert parameters["x"] == -90
    assert parameters["y"] == -190
    assert parameters["目的控件x"] == -60
    assert parameters["目的控件y"] == -160


def test_space_is_merged_into_input(monkeypatch):
    recorder = make_recorder(monkeypatch)
    monkeypatch.setattr(record.time, "time", lambda: 1.0)
    recorder.on_release(SimpleNamespace(vk=65, char="a"))
    recorder.on_release(keyboard.Key.space)
    recorder.on_release(SimpleNamespace(vk=66, char="b"))
    recorder.flush_last_input()
    assert recorder.actions_data[-1]["LOCATION"]["PARAMETERS"] == {"输入文本": "a b"}


def test_modifier_combo_is_recorded_once(monkeypatch):
    recorder = make_recorder(monkeypatch)
    ctrl_key = SimpleNamespace(vk=162, char=None)
    shift_key = SimpleNamespace(vk=160, char=None)
    c_key = SimpleNamespace(vk=67, char="c")

    recorder.on_press(ctrl_key)
    recorder.on_press(shift_key)
    recorder.on_press(c_key)

    assert [item["ACTION"] for item in recorder.actions_data] == ["组合键"]
    assert recorder.actions_data[0]["LOCATION"]["PARAMETERS"] == {"组合键": "ctrl_l+shift+c"}


def test_modifier_combo_ignores_key_auto_repeat(monkeypatch):
    recorder = make_recorder(monkeypatch)
    ctrl_key = SimpleNamespace(vk=162, char=None)
    c_key = SimpleNamespace(vk=67, char="c")
    recorder.on_press(ctrl_key)
    recorder.on_press(c_key)
    recorder.on_press(c_key)
    assert [item["ACTION"] for item in recorder.actions_data] == ["组合键"]


def test_held_modifier_records_consecutive_shortcuts(monkeypatch):
    recorder = make_recorder(monkeypatch)
    ctrl_key = SimpleNamespace(vk=162, char=None)
    c_key = SimpleNamespace(vk=67, char="c")
    v_key = SimpleNamespace(vk=86, char="v")

    recorder.on_press(ctrl_key)
    recorder.on_press(c_key)
    recorder.on_release(c_key)
    recorder.on_press(v_key)
    recorder.on_release(v_key)
    recorder.on_release(ctrl_key)

    assert [item["ACTION"] for item in recorder.actions_data] == ["组合键", "组合键"]
    assert [
        item["LOCATION"]["PARAMETERS"]["组合键"] for item in recorder.actions_data
    ] == ["ctrl_l+c", "ctrl_l+v"]


def test_windows_key_is_treated_as_modifier(monkeypatch):
    recorder = make_recorder(monkeypatch)
    win_key = SimpleNamespace(vk=91, char=None)
    r_key = SimpleNamespace(vk=82, char="r")

    recorder.on_press(win_key)
    recorder.on_press(r_key)

    assert recorder.actions_data[0]["ACTION"] == "组合键"
    assert recorder.actions_data[0]["LOCATION"]["PARAMETERS"] == {"组合键": "cmd+r"}


def test_shift_printable_key_is_recorded_as_text(monkeypatch):
    recorder = make_recorder(monkeypatch)
    shift_key = SimpleNamespace(vk=160, char=None)
    a_key = SimpleNamespace(vk=65, char="A")

    recorder.on_press(shift_key)
    recorder.on_press(a_key)
    recorder.on_release(a_key)
    recorder.on_release(shift_key)
    recorder.flush_last_input()

    assert [item["ACTION"] for item in recorder.actions_data] == ["输入文本"]
    assert recorder.actions_data[0]["LOCATION"]["PARAMETERS"] == {"输入文本": "A"}


def test_altgr_printable_key_is_recorded_as_text(monkeypatch):
    recorder = make_recorder(monkeypatch)
    ctrl_key = SimpleNamespace(vk=162, char=None)
    altgr_key = SimpleNamespace(vk=165, char=None)
    q_key = SimpleNamespace(vk=81, char="@")

    recorder.on_press(ctrl_key)
    recorder.on_press(altgr_key)
    recorder.on_press(q_key)
    recorder.on_release(q_key)
    recorder.on_release(altgr_key)
    recorder.on_release(ctrl_key)
    recorder.flush_last_input()

    assert [item["ACTION"] for item in recorder.actions_data] == ["输入文本"]
    assert recorder.actions_data[0]["LOCATION"]["PARAMETERS"] == {"输入文本": "@"}


def test_scroll_records_target_xpath(monkeypatch):
    recorder = make_recorder(monkeypatch)
    recorder.on_scroll(10, 10, 0, -3)
    action = recorder.actions_data[0]
    assert action["ACTION"] == "滚动"
    assert action["LOCATION"]["Xpath"]
    assert action["LOCATION"]["PARAMETERS"] == {"滚动距离": 3, "滚动方向": "down"}


def test_horizontal_scroll_is_recorded(monkeypatch):
    recorder = make_recorder(monkeypatch)
    recorder.on_scroll(10, 10, -2, 0)
    action = recorder.actions_data[0]
    assert action["LOCATION"]["PARAMETERS"] == {"滚动距离": 2, "滚动方向": "left"}


def test_intervening_right_click_prevents_false_double_click(monkeypatch):
    recorder = make_recorder(monkeypatch)
    times = iter([1.0, 1.05, 1.10, 1.15, 1.20, 1.25])
    monkeypatch.setattr(record.time, "time", lambda: next(times))

    recorder.on_click(10, 10, Button.left, True)
    recorder.on_click(10, 10, Button.left, False)
    recorder.on_click(10, 10, Button.right, True)
    recorder.on_click(10, 10, Button.right, False)
    recorder.on_click(10, 10, Button.left, True)
    recorder.on_click(10, 10, Button.left, False)

    assert [item["ACTION"] for item in recorder.actions_data] == ["点击", "右击", "点击"]


def test_drag_without_destination_locator_is_not_recorded(monkeypatch):
    recorder = make_recorder(monkeypatch)
    control = FakeControl()
    source = ([{"ControlType": "ButtonControl", "Name": "source"}], control)
    destination = ([], None)
    results = iter([source, destination])
    monkeypatch.setattr(record, "get_control_info", lambda _x, _y: next(results))
    monkeypatch.setattr(record.time, "time", lambda: 1.0)

    recorder.on_click(10, 10, Button.left, True)
    recorder.on_click(40, 40, Button.left, False)

    assert recorder.actions_data == []


def test_stopped_recorder_rejects_late_events(monkeypatch):
    recorder = make_recorder(monkeypatch)
    recorder._stopped = True

    recorder.on_click(10, 10, Button.left, True)
    recorder.on_scroll(10, 10, 0, -1)
    recorder.on_release(SimpleNamespace(vk=65, char="a"))

    assert recorder.actions_data == []


def test_stop_signals_completion(monkeypatch):
    recorder = make_recorder(monkeypatch)
    monkeypatch.setattr(recorder.mouse_listener, "stop", lambda: None)
    monkeypatch.setattr(recorder.keyboard_listener, "stop", lambda: None)
    recorder.stop()
    assert recorder._stop_complete.is_set()
    assert recorder.running is False


def test_stop_completes_when_listener_join_raises_oserror(monkeypatch):
    recorder = make_recorder(monkeypatch)
    monkeypatch.setattr(recorder.mouse_listener, "stop", lambda: None)
    monkeypatch.setattr(recorder.keyboard_listener, "stop", lambda: None)
    monkeypatch.setattr(
        recorder.mouse_listener,
        "join",
        lambda **_kwargs: (_ for _ in ()).throw(OSError("join failed")),
    )
    monkeypatch.setattr(recorder.keyboard_listener, "join", lambda **_kwargs: None)

    recorder.stop()

    assert recorder._stop_complete.is_set()
    assert recorder.running is False
    assert recorder.cleanup_errors == ["join failed"]


def test_stop_reports_listener_stop_failure(monkeypatch):
    recorder = make_recorder(monkeypatch)
    monkeypatch.setattr(
        recorder.mouse_listener,
        "stop",
        lambda: (_ for _ in ()).throw(OSError("stop failed")),
    )
    monkeypatch.setattr(recorder.keyboard_listener, "stop", lambda: None)
    monkeypatch.setattr(recorder.mouse_listener, "join", lambda **_kwargs: None)
    monkeypatch.setattr(recorder.keyboard_listener, "join", lambda **_kwargs: None)

    recorder.stop()

    assert recorder._stop_complete.is_set()
    assert recorder.cleanup_errors == ["鼠标监听器停止失败: stop failed"]


def test_record_highlight_process_starts_and_stops_cleanly():
    recorder = record.RecordThread()

    recorder.start_ui_thread()
    assert recorder.ui_process is not None
    assert recorder.ui_process.poll() is None
    recorder._stop_ui_process()

    assert recorder.ui_process.poll() is not None
