"""Legacy action signatures/messages share cancellation-safe bookkeeping."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from easy_uiauto import ctrl


CASES = [
    ("left_click", "点击", ["click"]),
    ("mouse_left_press", "鼠标左键按下", ["left-down"]),
    ("mouse_left_release", "鼠标左键释放", ["left-up"]),
    ("right_click", "右键点击", ["right-click"]),
    ("mouse_right_press", "鼠标右键按下", ["right-down"]),
    ("mouse_right_release", "鼠标右键释放", ["right-up"]),
    ("centre_click", "中键点击", ["middle-click"]),
    ("double_click", "双击", ["double-click"]),
    ("mouse_move_pos", "移动鼠标到坐标位置", ["hover"]),
    ("mouse_move_control", "移动鼠标到坐标位置", ["hover"]),
    ("drag_control", "拖动控件", ["drag"]),
    ("set_text", "设置文本", ["send-keys"]),
    ("input_text", "输入", ["click", "paste-text"]),
    ("key_click", "键盘操作", ["click", "press-key"]),
    ("key_press", "键盘按下", ["click", "key-down"]),
    ("key_release", "键盘释放", ["click", "key-up"]),
    ("key_group", "组合键", ["click", "hotkey"]),
]


def arguments(method):
    if method == "mouse_move_pos":
        return ("Step", 4, 5)
    parameters = {"x": 4, "y": 5, "设置文本": "text", "输入文本": "text",
                  "键盘按键": "A", "组合键": "Ctrl+A", "目的控件父窗口名称": "Example",
                  "目的控件Name": "Other", "目的控件ClassName": "", "目的控件ControlType": "ButtonControl",
                  "目的控件foundIndex": "", "目的控件AutomationId": "", "目的控件Xpath": []}
    return ("Step", "Example", "Target", "", "ButtonControl", 0, "", [], "", parameters)


@pytest.fixture
def execution(monkeypatch):
    calls = []
    control = SimpleNamespace(BoundingRectangle=SimpleNamespace(left=0, top=0, right=10, bottom=10))
    monkeypatch.setattr(ctrl, "find_control", lambda _location: control)
    monkeypatch.setattr(ctrl, "prepare_control", lambda *_args: (control, 4, 5))
    monkeypatch.setattr(ctrl, "get_pos", lambda *_args: (4, 5))
    monkeypatch.setattr(ctrl, "correct_ctrl_position", lambda *_args: (4, 5))
    monkeypatch.setattr(ctrl.Controller, "_activate_target", lambda *_args: None)
    monkeypatch.setattr(ctrl.Controller, "execute_resolved", lambda action, **_kw: calls.append(action))
    for name, action in [("press_key", "press-key"), ("key_down", "key-down"),
                         ("key_up", "key-up"), ("hotkey", "hotkey")]:
        monkeypatch.setattr(ctrl.Controller, name, lambda *_args, action=action: calls.append(action))
    states = Mock()
    monkeypatch.setattr(ctrl, "update_message_type", states)
    return calls, states


@pytest.mark.parametrize("method,title,expected", CASES)
def test_success_messages_and_action_sequences_remain_compatible(execution, method, title, expected):
    calls, states = execution
    result = getattr(ctrl.Controller, method)(*arguments(method))
    if method == "mouse_move_pos":
        assert result == "步骤：Step 位置【45】 执行动作 移动鼠标到坐标位置 成功！"
    elif method == "drag_control":
        assert result == "步骤：Step 执行动作 拖动控件 成功"
    elif method in {"input_text", "key_click", "key_press", "key_release", "key_group"}:
        assert result == f"步骤：Step执行动作 {title} 成功"
    else:
        assert result == f"步骤：Step 控件【Target】 执行动作 {title} 成功！"
    assert calls == expected
    states.assert_called_once_with(ctrl._MESSAGE.INFO)


@pytest.mark.parametrize("method,_title,_expected", CASES)
@pytest.mark.parametrize("error_type", [RuntimeError, KeyboardInterrupt, SystemExit])
def test_failure_runs_once_and_cancellation_propagates(monkeypatch, execution, method, _title, _expected, error_type):
    calls, states = execution
    failure = Mock(side_effect=error_type("[action_failed] injected"))
    monkeypatch.setattr(ctrl, "find_control", failure)
    monkeypatch.setattr(ctrl, "prepare_control", failure)
    monkeypatch.setattr(ctrl.Controller, "execute_resolved", failure)
    if error_type is RuntimeError:
        result = getattr(ctrl.Controller, method)(*arguments(method))
        assert "异常" in result and "[action_failed] injected" in result
        states.assert_called_once_with(ctrl._MESSAGE.ERROR)
    else:
        with pytest.raises(error_type, match="injected"):
            getattr(ctrl.Controller, method)(*arguments(method))
        states.assert_called_once()
    assert failure.call_count == 1
    assert calls == []


def test_cleanup_error_does_not_replace_cancellation(monkeypatch, execution):
    monkeypatch.setattr(ctrl.Controller, "execute_resolved", Mock(side_effect=KeyboardInterrupt("cancelled")))
    monkeypatch.setattr(ctrl, "update_message_type", Mock(side_effect=RuntimeError("cleanup")))
    with pytest.raises(KeyboardInterrupt) as raised:
        ctrl.Controller.mouse_move_pos("Step", 1, 2)
    assert any("cleanup" in note for note in raised.value.__notes__)


def test_drag_ambiguity_has_no_activation_or_action(monkeypatch, execution):
    from easy_uiauto.mcp.results import ActionError

    calls, _states = execution
    activation = Mock()
    monkeypatch.setattr(ctrl.Controller, "_activate_target", activation)
    monkeypatch.setattr(ctrl, "find_control", Mock(side_effect=[object(), ActionError("ambiguous_match", "duplicate destination")]))
    assert "[ambiguous_match]" in ctrl.Controller.drag_control(*arguments("drag_control"))
    activation.assert_not_called()
    assert calls == []


def test_attempted_input_marker_survives_legacy_message(monkeypatch, execution):
    monkeypatch.setattr(ctrl.Controller, "execute_resolved", Mock(side_effect=ctrl.ControlActionError("sent", executed=None)))
    value = ctrl.Controller.mouse_move_pos("Step", 1, 2)
    assert "[result_unconfirmed]" in value and "do not repeat" in value


@pytest.mark.parametrize("action,name", [("double-right-click", "RightClick"),
                                        ("double-middle-click", "MiddleClick")])
def test_secondary_button_double_click_uses_core_and_marks_partial_input(monkeypatch, action, name):
    method = Mock(side_effect=[None, RuntimeError("second click failed")])
    monkeypatch.setattr(ctrl.uiautomation, name, method)
    with pytest.raises(ctrl.ControlActionError) as raised:
        ctrl.Controller.execute_resolved(action, x=12, y=34)
    assert method.call_count == 2
    assert raised.value.executed is None
