"""Tests for the generated application UI command line."""

from __future__ import annotations

import json

from easy_uiauto.mcp import knowledge, ui_cli_command


def test_cli_lists_apps(monkeypatch, capsys) -> None:
    monkeypatch.setattr(
        knowledge,
        "list_apps",
        lambda: [{"id": "example", "name": "Example", "controls": 3}],
    )

    ui_cli_command.main(["apps"])

    result = json.loads(capsys.readouterr().out)
    assert result[0]["id"] == "example"


def test_cli_lists_verified_commands(monkeypatch, capsys, tmp_path) -> None:
    monkeypatch.setattr(knowledge, "app_dir", lambda _app_id: tmp_path)
    monkeypatch.setattr(
        knowledge,
        "available_commands",
        lambda _directory, _page: [{"command": "main.toolbar.save.click"}],
    )

    ui_cli_command.main(["commands", "example"])

    result = json.loads(capsys.readouterr().out)
    assert result == [{"command": "main.toolbar.save.click"}]


def test_cli_help_lists_teaching_and_confirmation(capsys) -> None:
    try:
        ui_cli_command.main(["--help"])
    except SystemExit as error:
        assert error.code == 0
    output = capsys.readouterr().out
    assert "teach" in output
    assert "batch" in output
    assert "real control meanings" in output

    parser = ui_cli_command._build_parser()
    run = parser.parse_args(["run", "example", "main.toolbar.delete.click", "--confirm"])
    assert run.confirm is True


def test_cli_batch_accepts_command_strings_and_text_objects(monkeypatch, capsys, tmp_path) -> None:
    monkeypatch.setattr(knowledge, "app_dir", lambda _app_id: tmp_path)
    captured = {}

    def execute_many(directory, steps, confirm):
        captured.update(directory=directory, steps=steps, confirm=confirm)
        return {"ok": True, "completed_steps": len(steps)}

    monkeypatch.setattr(ui_cli_command.ui_cli, "execute_many", execute_many)
    steps = '["main.toolbar.save.click", {"command":"main.form.name.set-text","text":"A"}]'

    ui_cli_command.main(["batch", "example", steps, "--confirm"])

    assert json.loads(capsys.readouterr().out)["completed_steps"] == 2
    assert captured == {
        "directory": tmp_path,
        "steps": [
            "main.toolbar.save.click",
            {"command": "main.form.name.set-text", "text": "A"},
        ],
        "confirm": True,
    }
