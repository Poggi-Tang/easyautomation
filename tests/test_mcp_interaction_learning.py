"""Tests for diff-driven UI operation-effect learning."""

from __future__ import annotations

from copy import deepcopy

from PIL import Image, ImageDraw

from easy_uiauto.mcp import interaction_learning, knowledge


def _image(changed: bool = False) -> Image.Image:
    image = Image.new("RGB", (200, 120), "white")
    if changed:
        ImageDraw.Draw(image).rectangle((50, 40, 120, 80), fill="black")
    return image


def _record() -> dict:
    return {
        "id": "save",
        "app_id": "example",
        "app_name": "Example",
        "page_id": "main",
        "region_id": "content",
        "semantic_name": "Save",
        "intent": "save-document",
        "description": "Save the document.",
        "semantic_status": "verified",
        "semantic_confidence": 0.95,
        "semantic_source": "ai-vision-target",
        "command": "main.content.save-document",
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


def _snapshot(image: Image.Image, state: str) -> dict:
    return {
        "captured_at": "2026-01-01T00:00:00+00:00",
        "window_name": "Example",
        "window_rect": {
            "left": 0,
            "top": 0,
            "right": 200,
            "bottom": 120,
            "width": 200,
            "height": 120,
        },
        "windows": [],
        "desktop_origin": {"left": 0, "top": 0},
        "action_control": {"name": "Save"},
        "state_id": state,
        "target_fingerprint": state,
        "_target_image": image,
        "_desktop_image": image,
    }


def test_changed_regions_reduces_pixel_diff_to_local_boxes() -> None:
    regions = interaction_learning.changed_regions(_image(), _image(True))

    assert len(regions) == 1
    assert regions[0]["left"] <= 50
    assert regions[0]["right"] >= 120
    assert regions[0]["width"] < 200


def test_learn_command_effect_persists_before_after_and_success(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    knowledge.save_control(directory, _record())
    knowledge.rebuild_index(directory)
    before = _snapshot(_image(), "before-state")
    after = _snapshot(_image(True), "after-state")
    monkeypatch.setattr(
        interaction_learning,
        "capture_snapshot",
        lambda *_args, **_kwargs: deepcopy(before),
    )
    monkeypatch.setattr(interaction_learning.scanner, "_find_window", lambda *_args: object())
    monkeypatch.setattr(interaction_learning.scanner, "_activate_window", lambda *_args: None)
    monkeypatch.setattr(
        interaction_learning,
        "wait_for_stability",
        lambda *_args, **_kwargs: (deepcopy(after), {"stable": True, "samples": 3}),
    )
    monkeypatch.setattr(
        interaction_learning.ui_cli,
        "execute",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        interaction_learning,
        "_controls_in_changed_regions",
        lambda *_args: [{"name": "Result", "control_type": "TextControl"}],
    )
    monkeypatch.setattr(
        interaction_learning,
        "_interpret_effects",
        lambda *_args: {
            "summary": "Result changed.",
            "effects": [
                {"type": "property_changed", "description": "Result text changed."}
            ],
            "success_condition": "Result value changes.",
            "confidence": 0.96,
        },
    )

    result = interaction_learning.learn_command_effect(
        directory,
        "main.content.save-document.click",
        "https://api.example/v1",
        "secret",
        "vision-model",
        "0.5.0",
    )

    assert result["status"] == "effect-observed"
    assert result["success_condition"] == "Result value changes."
    assert (directory / "interactions" / f"{result['interaction_id']}.md").is_file()
    interactions = knowledge.list_interactions(directory)
    assert interactions[0]["before_state_id"] == "before-state"
    control = knowledge.find_control_record(directory, "save")[1]
    assert control["function_verification"]["status"] == "effect-observed"


def test_safe_exploration_excludes_state_changing_and_blocked_intents(
    monkeypatch,
    tmp_path,
) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    commands = [
        {
            "command": "main.menu.open-help.click",
            "intent": "open-help",
            "action": "click",
            "risk": "safe",
            "requires_confirmation": False,
        },
        {
            "command": "main.form.save.click",
            "intent": "save-document",
            "action": "click",
            "risk": "state-changing",
            "requires_confirmation": False,
        },
        {
            "command": "main.chat.send.click",
            "intent": "send-message",
            "action": "click",
            "risk": "safe",
            "requires_confirmation": False,
        },
    ]
    monkeypatch.setattr(
        knowledge,
        "available_commands",
        lambda _directory, page_id="": [
            item for item in commands if not page_id or item.get("page", "main") == page_id
        ],
    )
    monkeypatch.setattr(
        interaction_learning.scanner,
        "scan_window",
        lambda *_args, **_kwargs: {
            "app_id": "example",
            "page_id": "main",
            "strategy": "visual-first",
            "visual_targets": 3,
            "controls_saved_this_scan": 3,
            "commands": 3,
        },
    )
    learned = []

    def learn(_directory, command, *_args, **_kwargs):
        learned.append(command)
        return {"command": command, "recovery": {"ok": True}}

    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)

    result = interaction_learning.explore_application(
        directory,
        "https://api.example/v1",
        "secret",
        "vision-model",
        "0.5.0",
        policy="safe",
    )

    assert result["ok"] is True
    assert learned == ["main.menu.open-help.click"]


def test_popup_scan_skips_desktop_tree_without_new_windows(monkeypatch) -> None:
    monkeypatch.setattr(
        interaction_learning.uiautomation,
        "GetRootControl",
        lambda: (_ for _ in ()).throw(AssertionError("desktop tree should not be read")),
    )

    assert interaction_learning._popup_controls({"added": []}) == []


def test_exploration_scans_and_recurses_into_new_dialog(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    commands = [
        {
            "command": "main.menu.open-help.click",
            "intent": "open-help",
            "action": "click",
            "risk": "safe",
            "requires_confirmation": False,
            "page": "main",
        },
        {
            "command": "help.content.read-more.click",
            "intent": "read-more",
            "action": "click",
            "risk": "safe",
            "requires_confirmation": False,
            "page": "help",
        },
    ]
    monkeypatch.setattr(
        knowledge,
        "available_commands",
        lambda _directory, page_id="": [
            item for item in commands if not page_id or item["page"] == page_id
        ],
    )
    scans = []

    def scan(title, *_args, **_kwargs):
        page = "main" if not scans else "help"
        scans.append(title)
        return {
            "app_id": "example",
            "page_id": page,
            "strategy": "visual-first",
            "visual_targets": 1,
            "controls_saved_this_scan": 1,
            "commands": 1,
        }

    monkeypatch.setattr(interaction_learning.scanner, "scan_window", scan)
    learned_commands = []

    def learn(_directory, command, *_args, post_effect=None, **_kwargs):
        learned_commands.append(command)
        discovery = {"attempted": False}
        if command.startswith("main.") and post_effect:
            discovery = post_effect(
                {
                    "window_name": "Example",
                    "window_changes": {"added": [{"title": "Help"}]},
                }
            )
        return {"command": command, "recovery": {"ok": True}, "discovery": discovery}

    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)

    result = interaction_learning.explore_application(
        directory,
        "https://api.example/v1",
        "secret",
        "vision-model",
        "0.5.0",
        max_depth=3,
    )

    assert result["ok"] is True
    assert result["actions_learned"] == 2
    assert learned_commands == [
        "main.menu.open-help.click",
        "help.content.read-more.click",
    ]
    assert scans == ["Example", "Help"]
