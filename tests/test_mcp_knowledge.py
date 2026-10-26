"""Tests for the Obsidian-compatible UI knowledge vault."""

from __future__ import annotations

import json

from easy_uiauto.mcp import knowledge, server


def test_owned_dialog_scan_does_not_replace_application_window_name(tmp_path):
    directory = knowledge.initialize_app("example", "Document - Example", tmp_path)
    knowledge.initialize_app("example", "About Example", tmp_path, preserve_name=True)
    assert knowledge.read_markdown(directory / "app.md")[0]["name"] == "Document - Example"


def test_runtime_quarantine_survives_four_rescans_and_rebuilds(tmp_path):
    from easy_uiauto.mcp import ui_cli

    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _control_record()
    knowledge.save_control(directory, record)
    ui_cli._quarantine(directory, record, "XPath step 4 failed; score=0.292; uniqueness=0.006")
    for _ in range(4):
        refreshed = {**record, "status": "verified", "verification": {
            "location": "passed", "image": "passed"}}
        knowledge.save_control(directory, refreshed)
        knowledge.rebuild_index(directory)
        held = knowledge.find_control_record(directory, record["id"], include_quarantine=True)[1]
        assert held["status"] == "quarantined" and held["quarantine"]["active"]
        assert "0.292" in held["quarantine"]["reason"]
        assert not knowledge.available_commands(directory)
    taught = knowledge.teach_control(directory, record["id"], "Save", "save", "Save item")
    assert taught["status"] == "quarantined"
    assert not knowledge.available_commands(directory)


def test_rescan_preserves_effect_but_human_teaching_can_add_new_confirmation(tmp_path):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _control_record()
    evidence = {"status": "effect-observed", "effect_record": "interactions/one.md"}
    knowledge.save_control(directory, {**record, "function_verification": evidence})
    knowledge.save_control(directory, {**record, "function_verification": {
        "status": "inferred", "executed": False}})
    saved = knowledge.find_control_record(directory, record["id"])[1]
    assert saved["function_verification"] == evidence
    knowledge.save_control(directory, {**saved, "function_verification": {
        "status": "human-confirmed", "executed": False}})
    assert knowledge.find_control_record(directory, record["id"])[1][
        "function_verification"]["status"] == "human-confirmed"


def test_runtime_failure_anchor_survives_changed_id_without_copying_old_function_proof():
    old = {**_control_record("old"), "status": "quarantined",
           "verification": {"runtime": "failed", "runtime_detail": "XPath failed"},
           "function_verification": {"status": "effect-observed", "effect_record": "old.md"}}
    new = {**_control_record("new"), "location": {"Xpath": ["different"]},
           "function_verification": {"status": "inferred"}}
    assert knowledge.control_anchor(old) == knowledge.control_anchor(new)
    held = knowledge.merge_control_state(old, new)
    assert held["id"] == "new" and held["status"] == "quarantined"
    assert held["quarantine"]["control_id"] == "old"
    assert held["function_verification"]["status"] == "inferred"


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


def test_batch_search_returns_multiple_intents_without_duplicate_controls(
    monkeypatch, tmp_path
) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    input_record = _control_record("input")
    input_record.update(
        {
            "semantic_name": "Current conversation message input",
            "intent": "compose-message",
            "aliases": ["message input"],
        }
    )
    send_record = _control_record("send")
    send_record.update(
        {
            "semantic_name": "Send message",
            "intent": "send-message",
            "aliases": ["send"],
        }
    )
    knowledge.save_control(directory, input_record)
    knowledge.save_control(directory, send_record)
    knowledge.rebuild_index(directory)
    monkeypatch.setattr(knowledge, "app_dir", lambda _app_id: directory)

    result = server.search_ui_knowledge_batch(
        "example", ["message", "send", "message"], limit_per_query=10
    )
    parsed = json.loads(result)

    assert parsed["queries"] == ["message", "send"]
    assert set(parsed["matches"]) == {"message", "send"}
    assert {item["id"] for item in parsed["controls"]} == {"input", "send"}


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
    record.update(name="Delete selected item", automation_id="delete-item")
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


def test_rebuild_migrates_volatile_geometry_and_process_fields(tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    page_rect = {
        "left": 100,
        "top": 200,
        "right": 500,
        "bottom": 500,
        "width": 400,
        "height": 300,
    }
    knowledge.write_markdown(
        directory / "pages" / "main.md",
        {"kind": "page", "id": "main", "name": "Main", "rect": page_rect},
        "# Main\n",
    )
    control = _control_record()
    control.update({"kind": "control", "rect": {**page_rect}})
    knowledge.write_markdown(directory / "controls" / "control-1.md", control, "# Save\n")
    knowledge.write_markdown(
        directory / "regions" / "main.toolbar.md",
        {
            "kind": "region",
            "id": "main.toolbar",
            "page_id": "main",
            "rect": {
                "left": 0,
                "top": 0,
                "right": 200,
                "bottom": 60,
                "width": 200,
                "height": 60,
            },
        },
        "# Toolbar\n",
    )
    knowledge.write_markdown(
        directory / "interactions" / "effect.md",
        {
            "kind": "interaction",
            "id": "effect",
            "before": {
                "window_rect": page_rect,
                "target_handle": 99,
                "windows": [
                    {
                        "title": "Example",
                        "handle": 99,
                        "owner_handle": 1,
                        "process_id": 11320,
                        "rect": page_rect,
                    }
                ],
            },
            "after": {"window_rect": page_rect, "windows": []},
            "changed_regions": [{"left": 40, "top": 30, "right": 160, "bottom": 90}],
        },
        "# Effect\n",
    )

    first = knowledge.rebuild_index(directory)
    second = knowledge.rebuild_index(directory)

    assert "rect" not in first["pages"][0]
    assert "rect" not in first["regions"][0]
    assert first["regions"][0]["normalized_rect"] == {
        "left": 0.0,
        "top": 0.0,
        "right": 0.5,
        "bottom": 0.2,
    }
    migrated_control = first["controls"][0]
    assert "rect" not in migrated_control
    assert migrated_control["normalized_rect"] == {
        "left": 0.0,
        "top": 0.0,
        "right": 1.0,
        "bottom": 1.0,
    }
    interaction = first["interactions"][0]
    assert "window_rect" not in interaction["before"]
    assert "target_handle" not in interaction["before"]
    assert interaction["before"]["windows"] == [{"title": "Example"}]
    assert interaction["changed_regions"][0] == {
        "left": 0.1,
        "top": 0.1,
        "right": 0.4,
        "bottom": 0.3,
    }
    assert second["interactions"][0]["changed_regions"] == interaction["changed_regions"]
