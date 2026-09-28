"""Action paths should not inherit UIAutomation's half-second post-action wait."""

from types import SimpleNamespace

from easy_uiauto import ctrl


def test_right_click_skips_uiautomation_default_wait(monkeypatch) -> None:
    calls = []
    control = SimpleNamespace(RightClick=lambda *args, **kwargs: calls.append((args, kwargs)))
    monkeypatch.setattr(ctrl, "find_control", lambda _location: control)
    monkeypatch.setattr(ctrl, "get_pos", lambda _control, _parameters: (12, 34))

    result = ctrl.Controller.right_click(
        "right click", "window", "item", "class", "ListItemControl", 0,
        "id", [], "", {},
    )

    assert "成功" in result
    assert calls == [((12, 34), {"simulateMove": False, "waitTime": 0})]


def test_left_click_skips_uiautomation_default_wait(monkeypatch) -> None:
    calls = []
    control = SimpleNamespace(
        ControlTypeName="MenuItemControl",
        IsEnabled=True,
        Click=lambda *args, **kwargs: calls.append((args, kwargs)),
    )
    monkeypatch.setattr(ctrl, "prepare_control", lambda *_args: (control, 12, 34))

    result = ctrl.Controller.left_click(
        "click", "window", "item", "class", "MenuItemControl", 0,
        "id", [], "", {},
    )

    assert "成功" in result
    assert calls == [((12, 34), {"simulateMove": False, "waitTime": 0})]
