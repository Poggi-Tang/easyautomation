"""Tests for semantic UI command execution and quarantine behavior."""

from __future__ import annotations

from copy import deepcopy
from types import SimpleNamespace

import pytest
from PIL import Image, ImageDraw

from easy_uiauto import ctrl
from easy_uiauto.mcp import knowledge, ui_cli


def _control(**properties):
    return SimpleNamespace(**{
        "IsEnabled": True,
        "BoundingRectangle": SimpleNamespace(left=20, top=40, right=120, bottom=80),
        **properties,
    })


@pytest.mark.parametrize("method", ["location", "image"])
def test_explicit_quarantine_revalidation_does_not_execute_action(
        monkeypatch, tmp_path, method):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    knowledge.save_control(directory, record)
    ui_cli._quarantine(directory, record, "old runtime failure")
    monkeypatch.setattr(ui_cli, "_resolve_verified_control", lambda *_a: (None, 0.98, {}, method))
    monkeypatch.setattr(ui_cli, "_perform_action", lambda *_a: (_ for _ in ()).throw(
        AssertionError("Revalidation must not execute")))
    result = ui_cli.revalidate_control(directory, record["id"])
    assert result["status"] == "verified" and result["action_executed"] is False
    saved = knowledge.find_control_record(directory, record["id"])[1]
    assert not saved["quarantine"]["active"]
    assert saved["quarantine"]["reason"] == "old runtime failure"
    assert knowledge.available_commands(directory)


@pytest.mark.parametrize("method", ["ocr", "ai-vision", "failed"])
def test_unconfirmed_quarantine_repair_cannot_release_it(monkeypatch, tmp_path, method):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    knowledge.save_control(directory, record)
    ui_cli._quarantine(directory, record, "score=0.292; uniqueness=0.006")

    def resolve(*_args):
        if method == "failed":
            raise RuntimeError("LOCATION and template failed")
        return None, 0.99, {}, method

    monkeypatch.setattr(ui_cli, "_resolve_verified_control", resolve)
    with pytest.raises(RuntimeError):
        ui_cli.revalidate_control(directory, record["id"])
    saved = knowledge.find_control_record(directory, record["id"], include_quarantine=True)[1]
    assert saved["quarantine"]["active"] and not knowledge.available_commands(directory)


def _record() -> dict:
    return {
        "id": "save",
        "app_id": "example",
        "app_name": "Example",
        "page_id": "main",
        "region_id": "toolbar",
        "semantic_name": "Save",
        "intent": "save-document",
        "description": "Save the current document.",
        "semantic_status": "verified",
        "semantic_confidence": 0.96,
        "semantic_source": "ai-vision-context",
        "command": "main.toolbar.save",
        "name": "Save",
        "control_type": "ButtonControl",
        "automation_id": "save",
        "location": {"WindowName": "Example", "Xpath": []},
        "actions": ["click"],
        "supported_actions": ["click"],
        "is_key": True,
        "status": "verified",
        "image": "images/controls/save.png",
        "tags": [],
        "notes": "",
    }


def test_execute_clicks_verified_control(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    clicks: list[dict] = []
    control = _control(Click=lambda **kwargs: clicks.append(kwargs))
    monkeypatch.setattr(
        ui_cli,
        "_resolve_verified_control",
        lambda *_args: (control, 0.95, {}, "location"),
    )

    result = ui_cli.execute(directory, "main.toolbar.save.click")

    assert result["ok"] is True
    assert clicks == [{"x": None, "y": None, "simulateMove": False, "waitTime": 0}]
    assert result["resolved_by"] == "location"


def test_execute_highlights_resolved_target_before_action(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    events = []
    rectangle = {"left": 20, "top": 40, "right": 120, "bottom": 80, "width": 100, "height": 40}
    control = _control(Click=lambda **_kwargs: events.append("click"))
    monkeypatch.setattr(
        ui_cli,
        "_resolve_verified_control",
        lambda *_args: (control, 0.95, rectangle, "location"),
    )
    monkeypatch.setattr(
        ui_cli.Controller,
        "show_targets",
        lambda rectangles, **_kwargs: events.append(("overlay", rectangles)) or SimpleNamespace(
            summary=lambda: {"shown": True, "controls": len(rectangles)},
            close=lambda: events.append("closed"),
        ),
    )

    result = ui_cli.execute(
        directory,
        "main.toolbar.save.click",
        highlight=True,
    )

    assert events[0][0] == "overlay"
    assert events[0][1] == [rectangle]
    assert events[1] == "click"
    assert events[2] == "closed"
    assert result["overlay"]["shown"] is True
    assert result["overlay"]["closed"] is True


def test_execute_uses_verified_image_fallback(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    knowledge.save_control(directory, _record())
    knowledge.rebuild_index(directory)
    clicks: list[tuple[int, int]] = []
    rectangle = {"left": 20, "top": 40, "width": 100, "height": 40}
    monkeypatch.setattr(
        ui_cli,
        "_resolve_verified_control",
        lambda *_args: (None, 0.91, rectangle, "image"),
    )
    monkeypatch.setattr(ctrl.uiautomation, "Click", lambda x, y, **kwargs: clicks.append(
        (x, y, kwargs["waitTime"])))

    result = ui_cli.execute(directory, "main.toolbar.save.click")

    assert result["resolved_by"] == "image"
    assert clicks == [(70, 60, 0)]


def test_perform_action_supports_right_click_and_hover(monkeypatch) -> None:
    calls = []
    rectangle = {"left": 20, "top": 40, "width": 100, "height": 40}
    monkeypatch.setattr(
        ctrl.uiautomation,
        "RightClick",
        lambda x, y, **kwargs: calls.append(("right-click", x, y, kwargs["waitTime"])),
    )
    monkeypatch.setattr(
        ctrl.uiautomation,
        "MoveTo",
        lambda x, y, **kwargs: calls.append(("hover", x, y, kwargs["waitTime"])),
    )

    ui_cli._perform_action("right-click", "", None, rectangle)
    ui_cli._perform_action("hover", "", None, rectangle, fast=True)

    assert calls == [("right-click", 70, 60, 0), ("hover", 70, 60, 0)]


def test_set_text_verifies_uia_value_without_using_clipboard(monkeypatch) -> None:
    pattern = SimpleNamespace(Value="")

    def set_value(value, *, waitTime):
        assert waitTime == 0
        pattern.Value = value
        return True

    pattern.SetValue = set_value
    control = SimpleNamespace(GetValuePattern=lambda: pattern)
    monkeypatch.setattr(
        ui_cli,
        "_capture_action_image",
        lambda _rectangle: Image.new("RGB", (100, 30), "white"),
    )
    monkeypatch.setattr(
        ui_cli,
        "_paste_text_with_clipboard",
        lambda *_args: (_ for _ in ()).throw(AssertionError("clipboard was not expected")),
    )

    result = ui_cli._perform_action(
        "set-text",
        "你好",
        control,
        {"left": 0, "top": 0, "width": 100, "height": 30},
    )

    assert result["method"] == "uia-set-value"
    assert result["verified"] is True
    assert result["evidence"] == "accessible-value"


@pytest.mark.parametrize("outcome", ["unreadable", "rejected", "exception"])
def test_unconfirmed_value_write_is_not_repeated_as_clipboard_paste(monkeypatch, outcome):
    calls = []

    def set_value(value, *, waitTime):
        calls.append((value, waitTime))
        if outcome == "exception":
            raise RuntimeError("Provider lost the write result")
        return outcome != "rejected"

    control = SimpleNamespace(
        GetValuePattern=lambda: SimpleNamespace(Value="", SetValue=set_value),
    )
    monkeypatch.setattr(ui_cli, "_capture_action_image", lambda _rect: Image.new("RGB", (100, 30)))
    monkeypatch.setattr(ui_cli, "sleep", lambda _seconds: None)
    monkeypatch.setattr(ui_cli, "_paste_text_with_clipboard", lambda *_args: pytest.fail(
        "An attempted ValuePattern write must not be repeated"))

    with pytest.raises(RuntimeError, match="result_unconfirmed"):
        ui_cli._perform_action(
            "set-text", "你好", control, {"left": 0, "top": 0, "width": 100, "height": 30},
        )
    assert calls == [("你好", 0)]


def test_missing_value_pattern_uses_verified_unicode_clipboard_paste(monkeypatch) -> None:
    control = _control(
        GetValuePattern=lambda: None,
        Click=lambda **_kwargs: None,
    )
    unchanged = Image.new("RGB", (100, 30), "white")
    changed = unchanged.copy()
    ImageDraw.Draw(changed).rectangle((5, 5, 40, 20), fill="black")
    images = iter((unchanged, unchanged, changed))
    monkeypatch.setattr(ui_cli, "_capture_action_image", lambda _rectangle: next(images))
    hotkeys = []
    monkeypatch.setattr(ui_cli.Controller, "press_key", lambda key: hotkeys.append((key,)))
    copied = []
    clipboard = ["original clipboard"]
    monkeypatch.setattr(ui_cli.Controller, "get_clipboard_text", lambda: clipboard[0])

    def copy(value):
        copied.append(value)
        clipboard[0] = value

    def hotkey(*keys):
        hotkeys.append(keys)
        if keys == ("ctrl", "c"):
            clipboard[0] = "你好"

    monkeypatch.setattr(ui_cli.Controller, "hotkey", hotkey)
    monkeypatch.setattr(ui_cli.Controller, "set_clipboard_text", copy)
    monkeypatch.setattr(ctrl.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(ui_cli, "sleep", lambda _seconds: None)

    result = ui_cli._perform_action(
        "set-text",
        "你好",
        control,
        {"left": 0, "top": 0, "width": 100, "height": 30},
    )

    assert result["method"] == "clipboard-paste"
    assert result["verified"] is True
    assert result["evidence"] == "clipboard-readback"
    assert hotkeys == [
        ("ctrl", "a"),
        ("ctrl", "v"),
        ("ctrl", "a"),
        ("ctrl", "c"),
        ("end",),
    ]
    assert copied[0] == "你好"
    assert copied[-1] == "original clipboard"


def test_set_text_rejects_clipboard_paste_without_observable_effect(monkeypatch) -> None:
    control = _control(
        GetValuePattern=lambda: None,
        Click=lambda **_kwargs: None,
    )
    monkeypatch.setattr(
        ui_cli,
        "_capture_action_image",
        lambda _rectangle: Image.new("RGB", (100, 30), "white"),
    )
    clipboard = [""]
    monkeypatch.setattr(ui_cli.Controller, "hotkey", lambda *_keys: None)
    monkeypatch.setattr(ui_cli.Controller, "press_key", lambda _key: None)
    monkeypatch.setattr(ui_cli.Controller, "get_clipboard_text", lambda: clipboard[0])
    monkeypatch.setattr(ui_cli.Controller, "set_clipboard_text", lambda value: clipboard.__setitem__(0, value))
    monkeypatch.setattr(ctrl.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(ui_cli, "sleep", lambda _seconds: None)

    try:
        ui_cli._perform_action(
            "set-text",
            "你好",
            control,
            {"left": 0, "top": 0, "width": 100, "height": 30},
        )
    except RuntimeError as error:
        assert "not recorded as successful" in str(error)
    else:
        raise AssertionError("unchanged input was reported as successful")


def test_execute_rejects_currently_disabled_control_with_learned_condition(
    monkeypatch, tmp_path
) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    record["enabling_condition"] = "Enter non-empty draft text"
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    control = SimpleNamespace(IsEnabled=False)
    monkeypatch.setattr(
        ui_cli,
        "_resolve_verified_control",
        lambda *_args: (
            control,
            0.95,
            {"left": 0, "top": 0, "width": 100, "height": 30},
            "automation-id",
        ),
    )

    try:
        ui_cli.execute(directory, "main.toolbar.save.click")
    except RuntimeError as error:
        assert "Enter non-empty draft text" in str(error)
    else:
        raise AssertionError("disabled control was clicked")


def test_resolution_uses_ocr_after_location_and_templates_fail(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    image_path = directory / record["image"]
    image_path.parent.mkdir(parents=True, exist_ok=True)
    Image.new("RGB", (20, 10), "white").save(image_path)
    rectangle = {"left": 20, "top": 40, "right": 120, "bottom": 80, "width": 100, "height": 40}
    monkeypatch.setattr(ui_cli, "resolve_location", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        ui_cli,
        "locate_template",
        lambda *_args: {"found": False, "score": 0.2, "detail": "not unique"},
    )
    monkeypatch.setattr(
        ui_cli,
        "_locate_record_text",
        lambda *_args: {"confidence": 0.91, "rect": rectangle},
    )

    result = ui_cli._resolve_verified_control_in_context(
        directory,
        record,
        object(),
        {"left": 0, "top": 0, "right": 200, "bottom": 100, "width": 200, "height": 100},
        Image.new("RGB", (200, 100), "white"),
    )

    assert result[0] is None
    assert result[1] == 0.91
    assert result[3] == "ocr"


def test_failed_runtime_verification_quarantines_control(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    knowledge.save_control(directory, _record())
    knowledge.rebuild_index(directory)
    monkeypatch.setattr(
        ui_cli,
        "_resolve_verified_control",
        lambda *_args: (_ for _ in ()).throw(RuntimeError("wrong image")),
    )

    try:
        ui_cli.execute(directory, "main.toolbar.save.click")
    except RuntimeError as error:
        assert "quarantined" in str(error)
    else:
        raise AssertionError("invalid runtime control was executed")

    assert (directory / "quarantine" / "save.md").is_file()
    assert not knowledge.available_commands(directory)


def test_dangerous_command_requires_explicit_confirmation(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    record["risk"] = "destructive"
    record["requires_confirmation"] = True
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)

    try:
        ui_cli.execute(directory, "main.toolbar.save.click")
    except RuntimeError as error:
        assert "explicit confirmation" in str(error)
    else:
        raise AssertionError("dangerous command executed without confirmation")

    try:
        ui_cli.execute_many(directory, ["main.toolbar.save.click"])
    except RuntimeError as error:
        assert "explicit confirmation" in str(error)
    else:
        raise AssertionError("dangerous batch executed without confirmation")

    clicks: list[bool] = []
    control = _control(Click=lambda **kwargs: clicks.append(kwargs["simulateMove"] is False))
    monkeypatch.setattr(
        ui_cli,
        "_resolve_verified_control",
        lambda *_args: (control, 0.95, {}, "location"),
    )
    result = ui_cli.execute(directory, "main.toolbar.save.click", confirm=True)
    assert result["ok"] is True
    assert clicks == [True]


def test_execute_many_shares_preflight_and_writes_once(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    save = _record()
    open_record = deepcopy(save)
    open_record.update(
        {
            "id": "open",
            "name": "Open",
            "semantic_name": "Open",
            "intent": "open-document",
            "command": "main.toolbar.open-document",
            "image": "images/controls/open.png",
        }
    )
    knowledge.save_control(directory, save)
    knowledge.save_control(directory, open_record)
    knowledge.rebuild_index(directory)

    calls = {"window": 0, "screenshot": 0, "resolved": [], "saved": [], "rebuilt": 0}
    clicks = []
    controls = {
        "save": _control(Click=lambda **_kwargs: clicks.append("save")),
        "open": _control(Click=lambda **_kwargs: clicks.append("open")),
    }

    def find_window(_name):
        calls["window"] += 1
        return object()

    def screenshot(_rectangle):
        calls["screenshot"] += 1
        return object()

    def resolve(_directory, record, _window, _window_rect, _screen):
        calls["resolved"].append(record["id"])
        return controls[record["id"]], 0.95, {}, "location"

    monkeypatch.setattr(ui_cli, "_find_window", find_window)
    monkeypatch.setattr(ui_cli, "_rect", lambda _control: {"left": 0})
    monkeypatch.setattr(ui_cli, "_screenshot_window", screenshot)
    monkeypatch.setattr(ui_cli, "_resolve_verified_control_in_context", resolve)
    monkeypatch.setattr(
        knowledge,
        "save_control",
        lambda _directory, record: calls["saved"].append(record["id"]),
    )
    monkeypatch.setattr(
        knowledge,
        "rebuild_index",
        lambda _directory: calls.__setitem__("rebuilt", calls["rebuilt"] + 1),
    )

    result = ui_cli.execute_many(
        directory,
        [
            "main.toolbar.save.click",
            "main.toolbar.save.click",
            "main.toolbar.open-document.click",
        ],
    )

    assert result["ok"] is True
    assert result["completed_steps"] == 3
    assert result["unique_controls_verified"] == 2
    assert result["knowledge_writes"] == 2
    assert calls == {
        "window": 1,
        "screenshot": 1,
        "resolved": ["save", "open"],
        "saved": ["save", "open"],
        "rebuilt": 1,
    }
    assert clicks == ["save", "save", "open"]


def test_execute_many_preflight_failure_runs_no_actions(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    save = _record()
    open_record = deepcopy(save)
    open_record.update(
        {
            "id": "open",
            "name": "Open",
            "semantic_name": "Open",
            "intent": "open-document",
            "command": "main.toolbar.open-document",
            "image": "images/controls/open.png",
        }
    )
    knowledge.save_control(directory, save)
    knowledge.save_control(directory, open_record)
    knowledge.rebuild_index(directory)
    clicks = []
    control = SimpleNamespace(Click=lambda **_kwargs: clicks.append(True))

    monkeypatch.setattr(ui_cli, "_find_window", lambda _name: object())
    monkeypatch.setattr(ui_cli, "_rect", lambda _control: {"left": 0})
    monkeypatch.setattr(ui_cli, "_screenshot_window", lambda _rectangle: object())

    def resolve(_directory, record, _window, _window_rect, _screen):
        if record["id"] == "open":
            raise RuntimeError("wrong image")
        return control, 0.95, {}, "location"

    monkeypatch.setattr(ui_cli, "_resolve_verified_control_in_context", resolve)

    try:
        ui_cli.execute_many(
            directory,
            ["main.toolbar.save.click", "main.toolbar.open-document.click"],
        )
    except RuntimeError as error:
        assert "before any action" in str(error)
    else:
        raise AssertionError("invalid batch was executed")

    assert clicks == []
    assert (directory / "quarantine" / "open.md").is_file()


def test_execute_many_rejects_commands_from_different_pages(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    save = _record()
    next_page = deepcopy(save)
    next_page.update(
        {
            "id": "next",
            "page_id": "settings",
            "intent": "next-page",
            "command": "settings.toolbar.next-page",
            "image": "images/controls/next.png",
        }
    )
    knowledge.save_control(directory, save)
    knowledge.save_control(directory, next_page)
    knowledge.rebuild_index(directory)
    monkeypatch.setattr(
        ui_cli,
        "_find_window",
        lambda _name: (_ for _ in ()).throw(AssertionError("window must not be touched")),
    )

    try:
        ui_cli.execute_many(
            directory,
            ["main.toolbar.save.click", "settings.toolbar.next-page.click"],
        )
    except ValueError as error:
        assert "one page" in str(error)
    else:
        raise AssertionError("cross-page batch was accepted")
