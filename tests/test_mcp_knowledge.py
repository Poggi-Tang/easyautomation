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


def test_teach_control_repairs_semantics_but_not_failed_positioning(tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _control_record(status="quarantined")
    record["semantic_status"] = "uncertain"
    record["verification"] = {"location": "failed", "image": "passed"}
    knowledge.save_control(directory, record)

    taught = knowledge.teach_control(
        directory,
        "control-1",
        "Save document",
        "save-document",
        "Save the current document to disk.",
        ["click"],
        ["save", "write file"],
        "state-changing",
    )

    assert taught["semantic_source"] == "manual"
    assert taught["semantic_confidence"] == 1.0
    assert taught["function_verification"]["status"] == "human-confirmed"
    assert taught["status"] == "quarantined"
    assert not knowledge.available_commands(directory)


def test_teach_control_publishes_when_positioning_is_verified(tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _control_record(status="quarantined")
    record["semantic_status"] = "uncertain"
    record["verification"] = {"location": "passed", "image": "passed"}
    knowledge.save_control(directory, record)

    taught = knowledge.teach_control(
        directory,
        "control-1",
        "Delete item",
        "delete-item",
        "Delete the selected item.",
        ["click"],
        ["remove"],
        "destructive",
    )

    assert taught["status"] == "verified"
    assert taught["requires_confirmation"] is True
    commands = knowledge.available_commands(directory)
    assert commands[0]["command"] == "main.toolbar.delete-item.click"
