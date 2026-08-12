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
    assert "real control meanings" in output

    parser = ui_cli_command._build_parser()
    run = parser.parse_args(["run", "example", "main.toolbar.delete.click", "--confirm"])
    assert run.confirm is True
