"""Tests for diff-driven UI operation-effect learning."""

from __future__ import annotations

from copy import deepcopy
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from PIL import Image, ImageDraw

from easy_uiauto.mcp import interaction_learning, knowledge


def test_exploration_ignores_commands_not_observed_in_current_scan(monkeypatch, tmp_path):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = {**_record(), "risk": "safe"}
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    monkeypatch.setattr(interaction_learning.scanner, "scan_window", lambda *_a, **_kw: dict(
        app_id="example", page_id="main", observed_control_ids=[]))
    learn = Mock(side_effect=AssertionError("Closed popup item selected from history"))
    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)
    result = interaction_learning.explore_application(directory, "url", "key", "model", "version")
    assert result["ok"] and result["actions_learned"] == 0
    learn.assert_not_called()


def test_parent_candidate_list_does_not_repeat_action_already_run_in_nested_page(
        monkeypatch, tmp_path):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    commands = [dict(command=f"main.menu.{intent}.click", intent=intent, action="click",
                     risk="safe", page="main") for intent in ("open-help", "read-more")]
    monkeypatch.setattr(knowledge, "available_commands", lambda *_a: commands)
    monkeypatch.setattr(interaction_learning.scanner, "scan_window", lambda *_a, **_kw: dict(
        app_id="example", page_id="main"))
    attempted = []

    def learn(_directory, command, *_a, post_effect=None, **_kw):
        attempted.append(command)
        if "open-help" in command:
            post_effect(dict(window_name="Example", window_changes={"added": []}))
        return dict(command=command, recovery={"ok": True})

    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)
    result = interaction_learning.explore_application(directory, "url", "key", "model", "version")
    assert result["ok"] and result["actions_learned"] == 2
    assert attempted == [item["command"] for item in commands]


@pytest.mark.parametrize("status", [
    "effect-observed", "no-observable-change", "interference-detected"])
@pytest.mark.parametrize("renamed", [False, True, "new-id"])
def test_explore_does_not_repeat_previous_attempt_even_if_scan_renames_command(
        monkeypatch, tmp_path, status, renamed):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = {**_record(), "risk": "safe"}
    knowledge.save_control(directory, record)
    command = record["command"] + ".click"
    knowledge.save_interaction(directory, dict(
        id="prior", app_id="example", page_id="main", command=command,
        control_id=record["id"], status=status))
    if renamed == "new-id":
        knowledge.save_control(directory, {**record, "id": "new-id", "command": "main.other.alias"})
    elif renamed:
        knowledge.save_control(directory, {**record, "command": "main.other.alias"})
    knowledge.rebuild_index(directory)
    monkeypatch.setattr(interaction_learning.scanner, "scan_window", lambda *_a, **_kw: dict(
        app_id="example", page_id="main", commands=1))
    learn = Mock(side_effect=AssertionError("Previously attempted control repeated"))
    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)
    result = interaction_learning.explore_application(directory, "url", "key", "model", "version")
    assert result["ok"] and result["actions_learned"] == 0
    learn.assert_not_called()
    assert len(knowledge.list_interactions(directory)) == 1


def test_explore_deduplication_includes_history_older_than_500_records(monkeypatch, tmp_path):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    command = "main.menu.open-help.click"
    history = [dict(page_id="main", command=command, status="no-observable-change")]
    history += [dict(page_id="other", command=f"other.menu.item-{i}.click",
                     status="effect-observed") for i in range(501)]
    monkeypatch.setattr(knowledge, "load_index", lambda *_a: dict(
        application={"name": "Example"}, interactions=history))
    monkeypatch.setattr(knowledge, "available_commands", lambda *_a: [dict(
        command=command, page="main", action="click", intent="open-help", risk="safe")])
    monkeypatch.setattr(interaction_learning.scanner, "scan_window", lambda *_a, **_kw: dict(
        app_id="example", page_id="main"))
    learn = Mock(side_effect=AssertionError("Old interaction was repeated"))
    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)
    result = interaction_learning.explore_application(directory, "url", "key", "model", "version")
    assert result["actions_learned"] == 0
    learn.assert_not_called()


@pytest.mark.parametrize("automation_id,name", [
    ("Restore", "Generic"), ("Close", "Generic"), ("Maximize", "Generic"),
    ("Minimize", "Generic"), ("Button", "还原"), ("Button", "最大化")])
def test_safe_exploration_excludes_window_state_buttons_even_when_semantics_say_safe(
        monkeypatch, tmp_path, automation_id, name):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = {**_record(), "risk": "safe", "intent": "window-control", "name": name,
              "automation_id": automation_id}
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    monkeypatch.setattr(interaction_learning.scanner, "scan_window", lambda *_a, **_kw: dict(
        app_id="example", page_id="main", commands=1))
    learn = Mock(side_effect=AssertionError("Window restore button must be excluded"))
    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)
    result = interaction_learning.explore_application(directory, "url", "key", "model", "version")
    assert result["actions_learned"] == 0
    learn.assert_not_called()


def test_new_unrelated_foreground_is_interference_not_an_allowed_popup(monkeypatch, tmp_path):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    before, after = _snapshot(_image(), "before"), _snapshot(_image(True), "after")
    before["target_handle"] = after["target_handle"] = 1
    target = dict(handle=1, title="Example", process_id=10, owner_handle=0,
                  rect=before["window_rect"], foreground=True)
    other = {**target, "handle": 2, "title": "Other app", "process_id": 20}
    before["windows"] = [target]
    after["windows"] = [{**target, "foreground": False}, other]
    monkeypatch.setattr(interaction_learning.scanner, "_find_window", lambda *_a: object())
    monkeypatch.setattr(interaction_learning.scanner, "_activate_window", lambda *_a: None)
    monkeypatch.setattr(interaction_learning, "capture_snapshot", lambda *_a: before)
    monkeypatch.setattr(interaction_learning, "wait_for_stability", lambda *_a, **_kw: (
        after, {"stable": True, "observed_windows": []}))
    monkeypatch.setattr(interaction_learning.ui_cli, "execute", lambda *_a, **_kw: {"ok": True})
    forbidden = Mock(side_effect=AssertionError("Interference must stop analysis and recovery"))
    for name in ("_interpret_effects", "_popup_controls", "_attempt_recovery"):
        monkeypatch.setattr(interaction_learning, name, forbidden)
    discover = Mock()
    result = interaction_learning.learn_command_effect(
        directory, record["command"] + ".click", "url", "key", "model", "version",
        recover=True, post_effect=discover)
    assert result["status"] == "interference-detected"
    assert result["recovery"]["ok"] is False
    forbidden.assert_not_called()
    discover.assert_not_called()


@pytest.mark.parametrize("disabled_state", [
    {"is_enabled_at_scan": False}, {"current_state": "disabled"}])
def test_exploration_skips_disabled_controls_and_continues_to_enabled_actions(
        monkeypatch, tmp_path, disabled_state):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    disabled = {**_record(), **disabled_state, "risk": "safe"}
    enabled = {**_record(), "id": "help", "command": "main.menu.open-help",
               "intent": "open-help", "risk": "safe", "is_enabled_at_scan": True}
    knowledge.save_control(directory, disabled)
    knowledge.save_control(directory, enabled)
    knowledge.rebuild_index(directory)
    monkeypatch.setattr(interaction_learning.scanner, "scan_window", lambda *_a, **_kw: dict(
        app_id="example", page_id="main"))
    learn = Mock(return_value={"recovery": {"ok": True}})
    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)
    result = interaction_learning.explore_application(directory, "url", "key", "model", "version")
    assert result["ok"] and result["actions_learned"] == 1
    assert learn.call_args.args[1] == "main.menu.open-help.click"


@pytest.mark.parametrize("intent,name", [
    ("new-window", "New window"), ("new-file", "New"),
    ("new-document", "New document"), ("generic", "新建窗口")])
def test_exploration_excludes_new_documents_without_escape_reversal(
        monkeypatch, tmp_path, intent, name):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    knowledge.save_control(directory, {**_record(), "risk": "safe", "intent": intent,
                                       "name": name})
    knowledge.rebuild_index(directory)
    monkeypatch.setattr(interaction_learning.scanner, "scan_window", lambda *_a, **_kw: dict(
        app_id="example", page_id="main"))
    learn = Mock(side_effect=AssertionError("Cannot recover a new document with Escape"))
    monkeypatch.setattr(interaction_learning, "learn_command_effect", learn)
    result = interaction_learning.explore_application(directory, "url", "key", "model", "version")
    assert result["ok"] and result["actions_learned"] == 0
    learn.assert_not_called()


def test_recovery_withholds_keys_when_target_cannot_be_activated(monkeypatch):
    monkeypatch.setattr(interaction_learning.scanner, "_find_window", lambda *_a: SimpleNamespace(
        NativeWindowHandle=42))
    monkeypatch.setattr(interaction_learning.scanner, "_activate_window", Mock(
        side_effect=RuntimeError("Foreground lock")))
    press, hotkey = Mock(), Mock()
    monkeypatch.setattr(interaction_learning.pyautogui, "press", press)
    monkeypatch.setattr(interaction_learning.pyautogui, "hotkey", hotkey)
    result = interaction_learning._attempt_recovery("Example", {"target_handle": 42}, {}, "click")
    assert result["ok"] is False and "withheld" in result["detail"]
    press.assert_not_called()
    hotkey.assert_not_called()


def test_recovery_activates_owned_modal_instead_of_disabled_owner(monkeypatch):
    owner = SimpleNamespace(NativeWindowHandle=42, IsEnabled=False)
    dialog = SimpleNamespace(NativeWindowHandle=43, IsEnabled=True)
    monkeypatch.setattr(interaction_learning.scanner, "_find_window", lambda *_a: owner)
    monkeypatch.setattr(interaction_learning.native_windows, "enumerate_windows", lambda **_kw: [
        dict(handle=99, title="Unrelated", owner_handle=0, enabled=True,
             bounds=dict(width=100, height=100)),
        dict(handle=43, title="About", owner_handle=42, enabled=True,
             bounds=dict(width=100, height=100))])
    monkeypatch.setattr(interaction_learning.uiautomation, "ControlFromHandle", lambda handle: (
        dialog if handle == 43 else pytest.fail("Unrelated window selected")))
    activation, press, hotkey = Mock(), Mock(), Mock()
    monkeypatch.setattr(interaction_learning.scanner, "_activate_window", activation)
    monkeypatch.setattr(interaction_learning.pyautogui, "press", press)
    monkeypatch.setattr(interaction_learning.pyautogui, "hotkey", hotkey)
    image = _image()
    monkeypatch.setattr(interaction_learning, "wait_for_stability", lambda *_a, **_kw: (
        {"_target_image": image}, {"stable": True}))
    result = interaction_learning._attempt_recovery(
        "Example", dict(target_handle=42, _target_image=image), {}, "click")
    assert result["ok"] is True
    activation.assert_called_once_with(dialog)
    press.assert_called_once_with("esc")
    hotkey.assert_not_called()


def test_nested_discovery_failure_withholds_parent_recovery(monkeypatch, tmp_path):
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    record = _record()
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    before, after = _snapshot(_image(), "before"), _snapshot(_image(True), "after")
    monkeypatch.setattr(interaction_learning.scanner, "_find_window", lambda *_a: object())
    monkeypatch.setattr(interaction_learning.scanner, "_activate_window", lambda *_a: None)
    monkeypatch.setattr(interaction_learning, "capture_snapshot", lambda *_a: before)
    monkeypatch.setattr(interaction_learning, "wait_for_stability", lambda *_a, **_kw: (
        after, {"stable": True, "observed_windows": []}))
    monkeypatch.setattr(interaction_learning.ui_cli, "execute", lambda *_a, **_kw: {"ok": True})
    monkeypatch.setattr(interaction_learning, "_controls_in_changed_regions", lambda *_a: [])
    monkeypatch.setattr(interaction_learning, "_interpret_effects", lambda *_a: dict(
        effects=[], summary="Menu opened", confidence=.9, success_condition="Menu visible"))
    monkeypatch.setattr(interaction_learning, "_navigation_likely", lambda *_a: True)
    recovery = Mock(side_effect=AssertionError("Unsafe parent recovery"))
    monkeypatch.setattr(interaction_learning, "_attempt_recovery", recovery)
    result = interaction_learning.learn_command_effect(
        directory, record["command"] + ".click", "url", "key", "model", "version",
        recover=True, post_effect=lambda *_a: dict(ok=False, stopped="Child interference"))
    assert result["recovery"]["ok"] is False
    assert "Nested discovery stopped" in result["recovery"]["detail"]
    recovery.assert_not_called()


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
