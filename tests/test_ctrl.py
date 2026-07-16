from __future__ import annotations

import pytest

from conftest import FakeControl
from easy_uiauto import ctrl


def action_args(parameters=None):
    return (
        "test action",
        "window",
        "control",
        "FakeClass",
        "ButtonControl",
        1,
        "",
        [{"ControlType": "ButtonControl", "Name": "control"}],
        "",
        parameters or {},
    )


@pytest.fixture
def prepared_control(monkeypatch):
    control = FakeControl()
    monkeypatch.setattr(ctrl, "find_control", lambda _location: control)
    monkeypatch.setattr(ctrl, "set_top_window", lambda _window_name: True)
    monkeypatch.setattr(ctrl, "show_ctrl_area", lambda *_args, **_kwargs: None)
    return control


def test_mouse_left_press_uses_absolute_screen_coordinates(monkeypatch, prepared_control):
    calls = []
    monkeypatch.setattr(ctrl.uiautomation, "PressMouse", lambda x, y: calls.append((x, y)))

    message = ctrl.Controller.mouse_left_press(*action_args())

    assert calls == [(120, 210)]
    assert "成功" in message


def test_mouse_release_moves_to_control_and_releases(monkeypatch, prepared_control):
    calls = []
    monkeypatch.setattr(ctrl.uiautomation, "ReleaseMouse", lambda: calls.append("release"))

    message = ctrl.Controller.mouse_left_release(*action_args())

    assert calls == ["release"]
    assert prepared_control.events[-1] == ("move", 120, 210)
    assert "成功" in message


def test_mouse_release_is_unconditional_when_control_disappears(monkeypatch):
    calls = []
    monkeypatch.setattr(ctrl, "find_control", lambda _location: None)
    monkeypatch.setattr(ctrl.uiautomation, "ReleaseMouse", lambda: calls.append("release"))

    message = ctrl.Controller.mouse_left_release(*action_args())

    assert calls == ["release"]
    assert "已在当前位置释放" in message


def test_right_mouse_press_and_release(monkeypatch, prepared_control):
    calls = []
    monkeypatch.setattr(
        ctrl.uiautomation,
        "RightPressMouse",
        lambda x, y: calls.append(("press", x, y)),
    )
    monkeypatch.setattr(
        ctrl.uiautomation,
        "RightReleaseMouse",
        lambda: calls.append(("release",)),
    )

    assert "成功" in ctrl.Controller.mouse_right_press(*action_args())
    assert "成功" in ctrl.Controller.mouse_right_release(*action_args())
    assert calls == [("press", 120, 210), ("release",)]


@pytest.mark.parametrize("method_name", ["input_text", "key_click", "key_press", "key_release", "key_group"])
def test_keyboard_actions_do_not_continue_after_locator_failure(monkeypatch, method_name):
    monkeypatch.setattr(ctrl, "find_control", lambda _location: None)
    calls = []
    monkeypatch.setattr(ctrl.pyautogui, "hotkey", lambda *args, **kwargs: calls.append(args))
    monkeypatch.setattr(ctrl.pyautogui, "press", lambda *args, **kwargs: calls.append(args))
    monkeypatch.setattr(ctrl.pyautogui, "keyDown", lambda *args, **kwargs: calls.append(args))
    monkeypatch.setattr(ctrl.pyautogui, "keyUp", lambda *args, **kwargs: calls.append(args))
    monkeypatch.setattr(ctrl.pyperclip, "copy", lambda value: calls.append((value,)))
    parameters = {
        "输入文本": "secret",
        "键盘按键": "a",
        "组合键": "ctrl+s",
    }

    message = getattr(ctrl.Controller, method_name)(*action_args(parameters))

    assert calls == []
    assert "异常" in message


def test_parameter_string_uses_safe_literal_parsing(prepared_control):
    assert ctrl.get_pos(prepared_control, "{'x': 2, 'y': 3}") == (2, 3)
    with pytest.raises(ValueError):
        ctrl.get_pos(prepared_control, "__import__('os').getcwd()")


def test_run_action_rejects_unknown_action():
    message = ctrl.run_action(
        {
            "ACTION": "不存在的动作",
            "LOCATION": {
                "WindowName": "",
                "Name": "",
                "ClassName": "",
                "ControlType": "",
                "foundIndex": 0,
                "AutomationId": "",
                "Xpath": [],
                "Img": "",
                "PARAMETERS": {},
            },
        }
    )
    assert "不支持" in message


def test_scroll_moves_to_recorded_control(monkeypatch, prepared_control):
    calls = []
    monkeypatch.setattr(ctrl, "auto_scroll", lambda amount, direction: calls.append((amount, direction)))

    message = ctrl.Controller.scroll(*action_args({"滚动距离": 3, "滚动方向": "down"}))

    assert calls == [(3, "down")]
    assert prepared_control.events[-1] == ("move", 120, 210)
    assert "成功" in message


def test_drag_uses_recorded_offsets(monkeypatch):
    source = FakeControl()
    target = FakeControl()
    calls = []
    monkeypatch.setattr(ctrl.uiautomation, "DragDrop", lambda *args: calls.append(args))

    ctrl.Controller.drag_control_by_control(source, target, (2, 3), (30, 10))

    assert calls == [(102, 203, 130, 210)]


class FakeTogglePattern:
    def __init__(self, state=0, succeeds=True):
        self.ToggleState = state
        self.succeeds = succeeds
        self.calls = 0

    def Toggle(self):
        self.calls += 1
        if self.succeeds:
            self.ToggleState = 0 if self.ToggleState == 1 else 1
        return self.succeeds


class FakeSelectionPattern:
    def __init__(self, succeeds=True):
        self.succeeds = succeeds
        self.calls = 0
        self.IsSelected = False

    def Select(self, **_kwargs):
        self.calls += 1
        self.IsSelected = self.succeeds
        return self.succeeds


class FakeExpandPattern:
    def __init__(self, state=0):
        self.calls = []
        self.ExpandCollapseState = state

    def Expand(self):
        self.calls.append("expand")
        self.ExpandCollapseState = 1
        return True

    def Collapse(self):
        self.calls.append("collapse")
        self.ExpandCollapseState = 0
        return True


class FakeValuePattern:
    def __init__(self, value="", read_only=False, succeeds=True):
        self.Value = value
        self.IsReadOnly = read_only
        self.succeeds = succeeds
        self.calls = []

    def SetValue(self, value, **_kwargs):
        self.calls.append(value)
        if self.succeeds:
            self.Value = value
        return self.succeeds


def test_get_pattern_falls_back_to_generic_get_pattern():
    expected = object()

    class GenericControl:
        def GetPattern(self, pattern_id):
            assert pattern_id == 123
            return expected

    assert ctrl._get_pattern(GenericControl(), "GetMissingPattern", 123) is expected


def test_set_control_value_uses_value_pattern_and_rejects_read_only():
    writable = FakeValuePattern()
    read_only = FakeValuePattern(read_only=True)

    class GenericControl:
        def __init__(self, pattern):
            self.pattern = pattern

        def GetPattern(self, _pattern_id):
            return self.pattern

    assert ctrl._set_control_value(GenericControl(writable), 42)
    assert writable.Value == "42"
    assert not ctrl._set_control_value(GenericControl(read_only), "blocked")
    assert read_only.calls == []


def test_set_control_value_rejects_provider_false_success(monkeypatch):
    pattern = FakeValuePattern(value="Alpha")
    pattern.SetValue = lambda value, **_kwargs: pattern.calls.append(value) or True

    class GenericControl:
        def GetPattern(self, _pattern_id):
            return pattern

    monotonic_values = iter([0.0, 1.0])
    monkeypatch.setattr(ctrl.time, "monotonic", lambda: next(monotonic_values))

    assert not ctrl._set_control_value(GenericControl(), "Beta")
    assert pattern.Value == "Alpha"


def test_left_click_prefers_invoke_pattern(monkeypatch, prepared_control):
    invoked = []

    class InvokePattern:
        def Invoke(self):
            invoked.append(True)
            return True

    prepared_control.GetPattern = lambda _pattern_id: InvokePattern()

    message = ctrl.Controller.left_click(*action_args())

    assert "成功" in message
    assert invoked == [True]
    assert not any(event[0] == "click" for event in prepared_control.events)


def test_left_click_stops_when_target_window_cannot_be_activated(
    monkeypatch, prepared_control
):
    monkeypatch.setattr(ctrl, "set_top_window", lambda _window_name: False)

    message = ctrl.Controller.left_click(*action_args())

    assert "异常" in message
    assert not any(event[0] == "click" for event in prepared_control.events)


def test_left_click_with_recorded_offset_still_prefers_semantic_action(
    monkeypatch, prepared_control
):
    invoked = []

    class InvokePattern:
        def Invoke(self):
            invoked.append(True)
            return True

    prepared_control.GetPattern = lambda _pattern_id: InvokePattern()

    message = ctrl.Controller.left_click(*action_args({"x": 2, "y": 3}))

    assert "成功" in message
    assert invoked == [True]
    assert not any(event[0] == "click" for event in prepared_control.events)


def test_left_click_can_force_recorded_coordinates(monkeypatch, prepared_control):
    invoked = []

    class InvokePattern:
        def Invoke(self):
            invoked.append(True)
            return True

    prepared_control.GetPattern = lambda _pattern_id: InvokePattern()

    message = ctrl.Controller.left_click(
        *action_args({"x": 2, "y": 3, "强制坐标": True})
    )

    assert "成功" in message
    assert invoked == []
    assert prepared_control.events[-1][:3] == ("click", 2, 3)


def test_toggle_control_honors_desired_state_and_indeterminate(monkeypatch, prepared_control):
    pattern = FakeTogglePattern(state=1)
    prepared_control.GetPattern = lambda _pattern_id: pattern

    assert "成功" in ctrl.Controller.toggle_control(*action_args({"选中": True}))
    assert pattern.calls == 0

    pattern.ToggleState = 2
    assert "成功" in ctrl.Controller.toggle_control(*action_args({"选中": True}))
    assert pattern.calls == 1


def test_toggle_control_reports_pattern_failure(monkeypatch, prepared_control):
    pattern = FakeTogglePattern(state=0, succeeds=False)
    prepared_control.GetPattern = lambda _pattern_id: pattern

    message = ctrl.Controller.toggle_control(*action_args({"选中": True}))

    assert "异常" in message
    assert ctrl.get_message_type() == ctrl._MESSAGE.ERROR


def test_toggle_control_normalizes_false_string(monkeypatch, prepared_control):
    pattern = FakeTogglePattern(state=1)
    prepared_control.GetPattern = lambda _pattern_id: pattern

    message = ctrl.Controller.toggle_control(*action_args({"选中": "false"}))

    assert "成功" in message
    assert pattern.ToggleState == 0


def test_select_control_uses_selection_item_pattern(monkeypatch, prepared_control):
    pattern = FakeSelectionPattern()
    prepared_control.GetPattern = lambda _pattern_id: pattern
    monkeypatch.setattr(
        ctrl.uiautomation.Control,
        "CreateControlFromControl",
        lambda control: control,
    )

    message = ctrl.Controller.select_control(*action_args())

    assert "成功" in message
    assert pattern.calls == 1


@pytest.mark.parametrize(
    ("parameters", "expected"),
    [({"展开": True}, "expand"), ({"展开": False}, "collapse")],
)
def test_expand_collapse_control(monkeypatch, prepared_control, parameters, expected):
    pattern = FakeExpandPattern(state=0 if expected == "expand" else 1)
    prepared_control.GetPattern = lambda _pattern_id: pattern

    message = ctrl.Controller.expand_collapse_control(*action_args(parameters))

    assert "成功" in message
    assert pattern.calls == [expected]


def test_expand_collapse_is_idempotent(monkeypatch, prepared_control):
    pattern = FakeExpandPattern(state=1)
    prepared_control.GetPattern = lambda _pattern_id: pattern

    message = ctrl.Controller.expand_collapse_control(*action_args({"展开": True}))

    assert "成功" in message
    assert pattern.calls == []


def test_select_named_item_is_scoped_to_target_control(monkeypatch):
    selection = FakeSelectionPattern()
    expand = FakeExpandPattern()

    class Item:
        def Exists(self, _timeout):
            return True

        def GetPattern(self, _pattern_id):
            return selection

    class ComboControl:
        def GetPattern(self, pattern_id):
            if pattern_id == ctrl.uiautomation.PatternId.ExpandCollapsePattern:
                return expand
            return None

        def ListItemControl(self, **kwargs):
            assert kwargs["Name"] == "Beta"
            return Item()

        def GetTopLevelControl(self):
            return self

    monkeypatch.setattr(
        ctrl.uiautomation,
        "ListItemControl",
        lambda **_kwargs: pytest.fail("must not search the entire desktop"),
    )
    assert ctrl._select_named_item(ComboControl(), "Beta")
    assert selection.IsSelected


def test_select_named_item_rejects_provider_false_success():
    selection = FakeSelectionPattern()
    value = FakeValuePattern(value="Alpha")
    expand = FakeExpandPattern()

    class Item:
        def Exists(self, _timeout):
            return True

        def GetPattern(self, _pattern_id):
            return selection

    class ComboControl:
        def GetPattern(self, pattern_id):
            if pattern_id == ctrl.uiautomation.PatternId.ExpandCollapsePattern:
                return expand
            if pattern_id == ctrl.uiautomation.PatternId.ValuePattern:
                return value
            return None

        def ListItemControl(self, **_kwargs):
            return Item()

        def GetTopLevelControl(self):
            return self

    assert not ctrl._select_named_item(ComboControl(), "Beta")
    assert selection.IsSelected
    assert value.Value == "Alpha"


def test_set_text_fallback_replaces_existing_text(monkeypatch, prepared_control):
    prepared_control.GetPattern = lambda _pattern_id: None

    message = ctrl.Controller.set_text(*action_args({"设置文本": "new"}))

    assert "成功" in message
    assert prepared_control.events[-3:] == [
        ("click", None, None, {"waitTime": 0}),
        ("send_keys", "{Ctrl}a"),
        ("send_keys", "new"),
    ]
