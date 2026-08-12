"""Tests for semantic UI command execution and quarantine behavior."""

from __future__ import annotations

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
        "command": "main.toolbar.save",
        "name": "Save",
        "control_type": "ButtonControl",
        "automation_id": "save",
        "location": {"WindowName": "Example", "Xpath": []},
        "actions": ["click"],
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
