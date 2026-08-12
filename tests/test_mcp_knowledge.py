"""Tests for the Obsidian-compatible UI knowledge vault."""

from __future__ import annotations

from easy_uiauto.mcp import knowledge


def _control_record(control_id: str = "control-1", status: str = "verified") -> dict:
    return {
        "id": control_id,
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
        "status": status,
        "image": "images/controls/control-1.png",
        "tags": ["save"],
        "notes": "",
    }


def test_markdown_is_source_of_truth_and_index_can_be_rebuilt(tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    knowledge.save_page(directory, {"id": "main", "name": "Main"})
    knowledge.save_region(directory, {"id": "main.toolbar", "name": "Toolbar"})
    knowledge.save_control(directory, _control_record())

    index = knowledge.rebuild_index(directory)
    (directory / ".easy_uiauto" / "index.json").unlink()
    rebuilt = knowledge.load_index(directory)

    assert index["controls"][0]["id"] == "control-1"
    assert rebuilt["controls"][0]["status"] == "verified"
    assert (directory / "controls" / "control-1.md").is_file()


def test_only_verified_controls_produce_ui_commands(tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    knowledge.save_control(directory, _control_record("verified", "verified"))
    quarantined = _control_record("bad", "quarantined")
    quarantined["command"] = "main.toolbar.bad"
    knowledge.save_control(directory, quarantined)
    knowledge.rebuild_index(directory)

    commands = knowledge.available_commands(directory)

    assert [item["command"] for item in commands] == ["main.toolbar.save.click"]
    assert (directory / "quarantine" / "bad.md").is_file()


def test_control_status_move_removes_old_markdown(tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _control_record()
    knowledge.save_control(directory, record)
    record["status"] = "quarantined"
    knowledge.save_control(directory, record)

    assert not (directory / "controls" / "control-1.md").exists()
    assert (directory / "quarantine" / "control-1.md").is_file()


def test_command_catalog_is_derived_from_verified_records(tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    knowledge.save_control(directory, _control_record())
    knowledge.rebuild_index(directory)

    catalog = knowledge.write_command_catalog(directory)

    assert "main.toolbar.save.click" in catalog.read_text(encoding="utf-8")
