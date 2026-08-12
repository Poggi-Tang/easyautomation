"""Tests for semantic UI command execution and quarantine behavior."""

from __future__ import annotations

from copy import deepcopy
from types import SimpleNamespace

from easy_uiauto.mcp import knowledge, ui_cli


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
    clicks: list[bool] = []
    control = SimpleNamespace(Click=lambda: clicks.append(True))
    monkeypatch.setattr(
        ui_cli,
        "_resolve_verified_control",
        lambda *_args: (control, 0.95, {}, "location"),
    )

    result = ui_cli.execute(directory, "main.toolbar.save.click")

    assert result["ok"] is True
    assert clicks == [True]
    assert result["resolved_by"] == "location"


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
    monkeypatch.setattr(ui_cli.pyautogui, "click", lambda x, y: clicks.append((x, y)))

    result = ui_cli.execute(directory, "main.toolbar.save.click")

    assert result["resolved_by"] == "image"
    assert clicks == [(70, 60)]


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
    control = SimpleNamespace(Click=lambda: clicks.append(True))
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
        "save": SimpleNamespace(Click=lambda **_kwargs: clicks.append("save")),
        "open": SimpleNamespace(Click=lambda **_kwargs: clicks.append("open")),
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
