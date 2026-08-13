"""Tests for application scanning and control-image verification."""

from __future__ import annotations

from types import SimpleNamespace

from PIL import Image, ImageDraw, ImageGrab

from easy_uiauto.mcp import knowledge, scanner


class FakeControl:
    def __init__(
        self,
        name: str,
        control_type: str,
        rect: tuple[int, int, int, int],
        children: list[FakeControl] | None = None,
        automation_id: str = "",
        parent: FakeControl | None = None,
    ) -> None:
        self.Name = name
        self.ControlTypeName = control_type
        self.ControlType = hash(control_type)
        self.ClassName = control_type
        self.AutomationId = automation_id
        self.FrameworkId = "Fake"
        self.LocalizedControlType = control_type.removesuffix("Control").lower()
        self.HelpText = ""
        self.IsKeyboardFocusable = control_type in {"ButtonControl", "EditControl"}
        self.IsEnabled = True
        self.NativeWindowHandle = 1
        self.ProcessId = 123
        self.BoundingRectangle = SimpleNamespace(
            left=rect[0], top=rect[1], right=rect[2], bottom=rect[3]
        )
        self._children = children or []
        self._parent = parent
        for child in self._children:
            child._parent = self

    def GetChildren(self):
        return self._children

    def GetParentControl(self):
        return self._parent


def _test_screen() -> Image.Image:
    image = Image.new("RGB", (400, 300), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((40, 30, 160, 90), fill="blue")
    draw.line((45, 35, 155, 85), fill="white", width=4)
    return image


def test_window_screenshot_uses_virtual_desktop_coordinates(monkeypatch) -> None:
    calls = []
    expected = Image.new("RGB", (300, 200), "white")
    monkeypatch.setattr(
        ImageGrab,
        "grab",
        lambda **kwargs: calls.append(kwargs) or expected,
    )

    image = scanner._screenshot_window(
        {
            "left": 2500,
            "top": -100,
            "right": 2800,
            "bottom": 100,
            "width": 300,
            "height": 200,
        }
    )

    assert image is expected
    assert calls == [{"bbox": (2500, -100, 2800, 100), "all_screens": True}]


def test_completion_reader_accepts_json_and_sse() -> None:
    json_response = SimpleNamespace(
        read=lambda: b'{"choices":[{"message":{"content":"{\\"ok\\":true}"}}]}'
    )
    sse_response = SimpleNamespace(
        read=lambda: (
            b'data: {"choices":[{"delta":{"content":"{\\"ok\\":"}}]}\n\n'
            b'data: {"choices":[],"usage":{"total_tokens":12}}\n\n'
            b'data: {"choices":[{"delta":{"content":"true}"}}]}\n\n'
            b"data: [DONE]\n\n"
        )
    )

    assert scanner._read_completion_content(json_response) == '{"ok":true}'
    assert scanner._read_completion_content(sse_response) == '{"ok":true}'


def test_location_resolution_reuses_shared_xpath_prefixes() -> None:
    save = FakeControl("Save", "ButtonControl", (20, 20, 80, 50))
    open_control = FakeControl("Open", "ButtonControl", (90, 20, 150, 50))
    toolbar = FakeControl(
        "Toolbar",
        "GroupControl",
        (0, 0, 200, 80),
        [save, open_control],
    )
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [toolbar])
    calls = {"window": 0}
    original = window.GetChildren

    def counted_children():
        calls["window"] += 1
        return original()

    window.GetChildren = counted_children
    cache = {}
    common = [
        {"ControlType": "WindowControl", "Name": "Example"},
        {"ControlType": "GroupControl", "Name": "Toolbar"},
    ]

    first = scanner.resolve_location(
        {
            "WindowName": "Example",
            "Xpath": [*common, {"ControlType": "ButtonControl", "Name": "Save"}],
        },
        window,
        cache,
    )
    second = scanner.resolve_location(
        {
            "WindowName": "Example",
            "Xpath": [*common, {"ControlType": "ButtonControl", "Name": "Open"}],
        },
        window,
        cache,
    )

    assert first is save
    assert second is open_control
    assert calls["window"] == 1


def test_location_uses_unique_automation_id_before_xpath(monkeypatch) -> None:
    class SearchResult:
        def __init__(self, found_index: int) -> None:
            self.found_index = found_index

        def Exists(self, *_args) -> bool:
            return self.found_index == 1

    def find_control(**kwargs):
        return SearchResult(kwargs["foundIndex"])

    monkeypatch.setattr(scanner.uiautomation, "Control", find_control)
    window = SimpleNamespace(
        GetChildren=lambda: (_ for _ in ()).throw(AssertionError("XPath was traversed"))
    )
    location = {
        "WindowName": "Example",
        "AutomationId": "chat_input_field",
        "ControlType": "EditControl",
        "Xpath": [
            {"ControlType": "WindowControl", "Name": "Example"},
            {
                "ControlType": "EditControl",
                "Name": "Old contact",
                "AutomationId": "chat_input_field",
            },
        ],
    }

    assert scanner.resolve_location(location, window=window).found_index == 1


def test_disabled_control_keeps_potential_actions() -> None:
    assert scanner._control_actions("ButtonControl", False) == ["click", "hover"]


def test_previous_control_identity_reuses_unique_automation_id() -> None:
    previous = {
        "id": "old-control-id",
        "automation_id": "chat_input_field",
        "control_type": "EditControl",
    }

    assert (
        scanner._previous_control_by_unique_automation_id(
            [previous], "chat_input_field", "EditControl"
        )
        == previous
    )
    assert (
        scanner._previous_control_by_unique_automation_id(
            [previous, {**previous, "id": "duplicate"}],
            "chat_input_field",
            "EditControl",
        )
        == {}
    )


def _semantic_result(
    candidates: list[dict],
    confidence: float = 0.96,
    intent: str = "save-document",
) -> dict[int, dict]:
    return {
        item["id"]: {
            "semantic_name": "Save document",
            "intent": intent,
            "description": "Save the current document to disk.",
            "role": "action",
            "actions": ["click"],
            "aliases": ["save", "write file"],
            "risk": "state-changing",
            "requires_confirmation": False,
            "confidence": confidence,
            "evidence": ["Save label", "button in content region"],
            "ambiguity": "" if confidence >= 0.72 else "Icon meaning is unclear.",
            "source": "ai-vision-context",
            "entity_type": "document-action",
            "observed_value": "",
            "dynamic_context": False,
            "current_state": "enabled",
            "state_reason": "",
            "enabling_condition": "",
            "relationships": [],
        }
        for item in candidates
    }


def test_template_verification_requires_correct_unique_location() -> None:
    screen = _test_screen()
    template = screen.crop((40, 30, 160, 90))
    window_rect = {"left": 0, "top": 0, "right": 400, "bottom": 300}
    expected = {
        "left": 40,
        "top": 30,
        "right": 160,
        "bottom": 90,
        "width": 120,
        "height": 60,
    }

    result = scanner._template_verification(template, screen, expected, window_rect)

    assert result["ok"] is True
    assert result["score"] > 0.99


def test_resolve_location_reuses_supplied_window(monkeypatch) -> None:
    button = FakeControl("Save", "ButtonControl", (40, 30, 160, 90), automation_id="save")
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [button])
    location = {
        "WindowName": "Example",
        "Xpath": [
            {"ControlType": "WindowControl", "Name": "Example", "searchDepth": 1},
            {
                "ControlType": "ButtonControl",
                "Name": "Save",
                "AutomationId": "save",
                "searchDepth": 2,
            },
        ],
    }
    monkeypatch.setattr(
        scanner,
        "_find_window",
        lambda _name: (_ for _ in ()).throw(AssertionError("window lookup was repeated")),
    )

    assert scanner.resolve_location(location, window=window) is button


def test_visual_target_selects_actionable_ancestor() -> None:
    label = FakeControl("Save", "TextControl", (50, 40, 110, 70))
    button = FakeControl(
        "Save",
        "ButtonControl",
        (40, 30, 160, 90),
        [label],
        automation_id="save",
    )
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [button])
    target = {
        "role": "action",
        "actions": ["click"],
        "relative_rect": {
            "left": 40,
            "top": 30,
            "right": 160,
            "bottom": 90,
            "width": 120,
            "height": 60,
        },
    }

    selected = scanner.control_from_visual_target(
        window,
        scanner._rect(window),
        target,
        point_lookup=lambda _x, _y: label,
    )

    assert selected is button


def test_visual_targets_on_same_uia_control_keep_semantic_components(monkeypatch) -> None:
    control = FakeControl("圆子", "ListItemControl", (20, 20, 180, 80), automation_id="圆子")
    window = FakeControl("微信", "WindowControl", (0, 0, 400, 300), [control])
    targets = [
        {
            "id": 1,
            "semantic_name": "Open conversation",
            "intent": "open-conversation",
            "actions": ["click"],
            "relative_rect": {
                "left": 20,
                "top": 20,
                "right": 180,
                "bottom": 80,
                "width": 160,
                "height": 60,
            },
        },
        {
            "id": 2,
            "semantic_name": "Contact avatar",
            "intent": "identify-contact-avatar",
            "actions": [],
            "entity_type": "contact-avatar",
            "relative_rect": {
                "left": 25,
                "top": 25,
                "right": 65,
                "bottom": 65,
                "width": 40,
                "height": 40,
            },
        },
    ]
    monkeypatch.setattr(scanner, "control_from_visual_target", lambda *_args: control)
    monkeypatch.setattr(
        scanner,
        "get_control_xpath",
        lambda _control: [
            {"ControlType": "WindowControl", "Name": "微信"},
            {"ControlType": "ListItemControl", "AutomationId": "圆子"},
        ],
    )

    controls, semantics, failures = scanner.controls_from_visual_targets(
        window, scanner._rect(window), targets, 20
    )

    assert len(controls) == 1
    assert failures == []
    assert semantics[1]["intent"] == "open-conversation"
    assert semantics[1]["visual_components"][0]["entity_type"] == "contact-avatar"


def test_segment_validation_keeps_only_valid_key_targets() -> None:
    result = scanner._validate_segments(
        {
            "controls": [
                {
                    "semantic_name": "Save",
                    "intent": "save-document",
                    "role": "action",
                    "actions": ["click"],
                    "risk": "state-changing",
                    "confidence": 0.95,
                    "entity_type": "message-action",
                    "current_state": "disabled",
                    "state_reason": "Message draft is empty",
                    "enabling_condition": "Enter non-empty draft text",
                    "relationships": [
                        {"type": "enabled-by", "target_intent": "compose-message"}
                    ],
                    "left": 10,
                    "top": 20,
                    "width": 80,
                    "height": 30,
                },
                {"semantic_name": "invalid", "left": 1, "top": 1, "width": 0, "height": 0},
            ]
        },
        400,
        300,
    )

    assert len(result["controls"]) == 1
    assert result["controls"][0]["intent"] == "save-document"
    assert result["controls"][0]["relative_rect"]["width"] == 80
    assert result["controls"][0]["entity_type"] == "message-action"
    assert result["controls"][0]["current_state"] == "disabled"
    assert result["controls"][0]["enabling_condition"] == "Enter non-empty draft text"
    assert result["controls"][0]["relationships"] == [
        {"type": "enabled-by", "target_intent": "compose-message"}
    ]


def test_scan_writes_images_markdown_and_verified_command(monkeypatch, tmp_path) -> None:
    button = FakeControl("Save", "ButtonControl", (40, 30, 160, 90), automation_id="save")
    label = FakeControl("Heading", "TextControl", (20, 120, 180, 150))
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [button, label])
    monkeypatch.setattr(scanner, "_activate_window", lambda _window: None)
    monkeypatch.setattr(scanner, "_find_window", lambda _title: window)
    monkeypatch.setattr(scanner, "_process_name", lambda _pid: "example-app")
    monkeypatch.setattr(scanner, "_screenshot_window", lambda _rect: _test_screen())
    monkeypatch.setattr(
        scanner,
        "segment_interface",
        lambda *_args: {
            "page": {"id": "main", "name": "Main", "description": ""},
            "regions": [
                {
                    "id": "content",
                    "name": "Content",
                    "role": "content",
                    "description": "",
                    "rect": {
                        "left": 0,
                        "top": 0,
                        "right": 400,
                        "bottom": 300,
                        "width": 400,
                        "height": 300,
                    },
                }
            ],
        },
    )
    monkeypatch.setattr(
        scanner,
        "get_control_xpath",
        lambda control: [
            {"ControlType": "WindowControl", "Name": "Example"},
            {"ControlType": control.ControlTypeName, "Name": control.Name},
        ],
    )
    monkeypatch.setattr(
        scanner,
        "_verify_location",
        lambda _location, rect, *_args: (True, rect, "bounds IoU=1.000"),
    )
    monkeypatch.setattr(
        scanner,
        "analyze_control_semantics",
        lambda _image, candidates, *_args: _semantic_result(candidates),
    )

    result = scanner.scan_window(
        "Example",
        "https://api.example/v1",
        "secret",
        "vision-model",
        "1.0.0",
        strategy="full-uia",
        root=tmp_path,
    )

    directory = knowledge.app_dir("example-app", tmp_path)
    controls = knowledge.load_index(directory)["controls"]
    commands = knowledge.available_commands(directory)
    assert result["controls_saved_this_scan"] == 3
    assert result["knowledge_controls_total"] == 3
    assert result["commands"] == 1
    assert any(control["status"] == "verified" for control in controls)
    assert commands[0]["command"].endswith(".save-document.click")
    button_record = next(control for control in controls if control["actions"])
    assert button_record["semantic_name"] == "Save document"
    assert button_record["intent"] == "save-document"
    assert button_record["semantic_confidence"] == 0.96
    assert button_record["function_verification"]["status"] == "inferred"
    assert button_record["entity_type"] == "document-action"
    assert button_record["current_state"] == "enabled"
    assert "rect" not in button_record
    assert button_record["normalized_rect"] == {
        "left": 0.1,
        "top": 0.1,
        "right": 0.4,
        "bottom": 0.3,
    }
    assert button_record["geometry_role"] == "visual-hint-only"
    index = knowledge.load_index(directory)
    assert "rect" not in index["pages"][0]
    assert "rect" not in index["regions"][0]
    assert index["regions"][0]["normalized_rect"] == {
        "left": 0.0,
        "top": 0.0,
        "right": 1.0,
        "bottom": 1.0,
    }
    assert list((directory / "images" / "controls").glob("*.png"))
    assert (directory / "images" / "pages" / "main.annotated.png").is_file()
    assert result["annotated_controls"] == 3
    assert result["annotated_page_image"].endswith("main.annotated.png")
    assert (directory / "operations" / "UI-CLI.md").is_file()


def test_visual_first_scan_skips_full_tree_and_second_ai_request(monkeypatch, tmp_path) -> None:
    button = FakeControl("Save", "ButtonControl", (40, 30, 160, 90), automation_id="save")
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [button])
    target = {
        "id": 1,
        "region_id": "content",
        "semantic_name": "Save document",
        "intent": "save-document",
        "description": "Save the current document.",
        "role": "action",
        "actions": ["click"],
        "aliases": ["save"],
        "risk": "state-changing",
        "requires_confirmation": False,
        "confidence": 0.97,
        "evidence": ["Save label"],
        "ambiguity": "",
        "source": "ai-vision-target",
        "relative_rect": {
            "left": 40,
            "top": 30,
            "right": 160,
            "bottom": 90,
            "width": 120,
            "height": 60,
        },
    }
    monkeypatch.setattr(scanner, "_activate_window", lambda _window: None)
    monkeypatch.setattr(scanner, "_find_window", lambda _title: window)
    monkeypatch.setattr(scanner, "_process_name", lambda _pid: "example-app")
    monkeypatch.setattr(scanner, "_screenshot_window", lambda _rect: _test_screen())
    monkeypatch.setattr(
        scanner,
        "segment_interface",
        lambda *_args: {
            "page": {"id": "main", "name": "Main", "description": ""},
            "regions": scanner._validate_segments({}, 400, 300)["regions"],
            "controls": [target],
        },
    )
    monkeypatch.setattr(
        scanner,
        "controls_from_visual_targets",
        lambda *_args: ([(button, 1)], {1: target}, []),
    )
    monkeypatch.setattr(
        scanner,
        "_walk_controls",
        lambda *_args: (_ for _ in ()).throw(AssertionError("full tree was traversed")),
    )
    monkeypatch.setattr(
        scanner,
        "analyze_control_semantics",
        lambda *_args: (_ for _ in ()).throw(AssertionError("second AI request was made")),
    )
    monkeypatch.setattr(
        scanner,
        "get_control_xpath",
        lambda control: [
            {"ControlType": "WindowControl", "Name": "Example"},
            {"ControlType": control.ControlTypeName, "Name": control.Name},
        ],
    )
    monkeypatch.setattr(
        scanner,
        "_verify_location",
        lambda _location, rect, *_args: (True, rect, "bounds IoU=1.000"),
    )
    directory = knowledge.initialize_app("example-app", "Example", tmp_path)
    knowledge.save_control(
        directory,
        {
            "id": "previous-key-control",
            "app_id": "example-app",
            "app_name": "Example",
            "page_id": "main",
            "region_id": "content",
            "semantic_name": "Previously learned control",
            "status": "verified",
            "actions": [],
        },
    )
    knowledge.rebuild_index(directory)

    result = scanner.scan_window(
        "Example",
        "https://api.example/v1",
        "secret",
        "vision-model",
        "0.5.0",
        root=tmp_path,
    )

    assert result["strategy"] == "visual-first"
    assert result["visual_targets"] == 1
    assert result["uia_controls_seen"] == 1
    assert result["semantic_controls_analyzed"] == 1
    assert result["controls_retained_unobserved"] == 1
    assert knowledge.find_control_record(directory, "previous-key-control") is not None


def test_visual_first_scan_retains_visual_entity_without_distinct_uia_node(
    monkeypatch, tmp_path
) -> None:
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300))
    target = {
        "id": 1,
        "region_id": "content",
        "semantic_name": "Current contact avatar",
        "intent": "identify-current-contact-avatar",
        "description": "Avatar for the current conversation contact.",
        "role": "identity",
        "entity_type": "contact-avatar",
        "observed_value": "",
        "dynamic_context": True,
        "current_state": "visible",
        "state_reason": "",
        "enabling_condition": "",
        "relationships": [
            {"type": "identifies", "target_entity": "current-conversation-contact"}
        ],
        "actions": [],
        "aliases": ["avatar"],
        "risk": "safe",
        "requires_confirmation": False,
        "confidence": 0.95,
        "evidence": ["portrait beside conversation title"],
        "ambiguity": "",
        "source": "ai-vision-target",
        "relative_rect": {
            "left": 20,
            "top": 20,
            "right": 70,
            "bottom": 70,
            "width": 50,
            "height": 50,
        },
    }
    monkeypatch.setattr(scanner, "_activate_window", lambda _window: None)
    monkeypatch.setattr(scanner, "_find_window", lambda _title: window)
    monkeypatch.setattr(scanner, "_process_name", lambda _pid: "example-app")
    monkeypatch.setattr(scanner, "_screenshot_window", lambda _rect: _test_screen())
    monkeypatch.setattr(
        scanner,
        "segment_interface",
        lambda *_args: {
            "page": {"id": "main", "name": "Main", "description": ""},
            "regions": scanner._validate_segments({}, 400, 300)["regions"],
            "controls": [target],
        },
    )
    monkeypatch.setattr(
        scanner,
        "controls_from_visual_targets",
        lambda *_args: (
            [],
            {},
            [
                {
                    "visual_target": "Current contact avatar",
                    "error": "no distinct UIA node",
                    "observation": target,
                }
            ],
        ),
    )

    result = scanner.scan_window(
        "Example",
        "https://api.example/v1",
        "secret",
        "vision-model",
        "0.6.1",
        root=tmp_path,
    )

    directory = knowledge.app_dir("example-app", tmp_path)
    records = knowledge.search_controls(directory, "contact avatar")
    assert result["visual_only_observations"] == 1
    assert records[0]["status"] == "observed"
    assert records[0]["entity_type"] == "contact-avatar"
    assert records[0]["dynamic_context"] is True
    assert records[0]["actions"] == []


def test_failed_actionable_control_is_quarantined(monkeypatch, tmp_path) -> None:
    button = FakeControl("Save", "ButtonControl", (40, 30, 160, 90))
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [button])
    monkeypatch.setattr(scanner, "_activate_window", lambda _window: None)
    monkeypatch.setattr(scanner, "_find_window", lambda _title: window)
    monkeypatch.setattr(scanner, "_process_name", lambda _pid: "example-app")
    monkeypatch.setattr(scanner, "_screenshot_window", lambda _rect: _test_screen())
    monkeypatch.setattr(
        scanner,
        "segment_interface",
        lambda *_args: scanner._validate_segments({}, 400, 300),
    )
    monkeypatch.setattr(
        scanner,
        "get_control_xpath",
        lambda control: [{"ControlType": control.ControlTypeName, "Name": control.Name}],
    )
    monkeypatch.setattr(
        scanner,
        "_verify_location",
        lambda _location, _rect, *_args: (False, {}, "LOCATION did not resolve"),
    )
    monkeypatch.setattr(
        scanner,
        "_template_verification",
        lambda *_args: {
            "ok": False,
            "score": 0.5,
            "second_score": 0.49,
            "detail": "not unique",
        },
    )
    monkeypatch.setattr(
        scanner,
        "analyze_control_semantics",
        lambda _image, candidates, *_args: _semantic_result(candidates),
    )

    result = scanner.scan_window(
        "Example",
        "https://api.example/v1",
        "secret",
        "vision-model",
        "1.0.0",
        strategy="full-uia",
        root=tmp_path,
    )

    directory = knowledge.app_dir("example-app", tmp_path)
    assert result["status_counts"]["quarantined"] == 1
    assert not knowledge.available_commands(directory)


def test_low_confidence_semantics_are_quarantined(monkeypatch, tmp_path) -> None:
    button = FakeControl("", "ButtonControl", (40, 30, 160, 90))
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [button])
    monkeypatch.setattr(scanner, "_activate_window", lambda _window: None)
    monkeypatch.setattr(scanner, "_find_window", lambda _title: window)
    monkeypatch.setattr(scanner, "_process_name", lambda _pid: "example-app")
    monkeypatch.setattr(scanner, "_screenshot_window", lambda _rect: _test_screen())
    monkeypatch.setattr(
        scanner,
        "segment_interface",
        lambda *_args: scanner._validate_segments({}, 400, 300),
    )
    monkeypatch.setattr(
        scanner,
        "analyze_control_semantics",
        lambda _image, candidates, *_args: _semantic_result(candidates, confidence=0.4),
    )
    monkeypatch.setattr(
        scanner,
        "get_control_xpath",
        lambda control: [{"ControlType": control.ControlTypeName, "Name": control.Name}],
    )
    monkeypatch.setattr(
        scanner,
        "_verify_location",
        lambda _location, rect, *_args: (True, rect, "bounds IoU=1.000"),
    )

    result = scanner.scan_window(
        "Example",
        "https://api.example/v1",
        "secret",
        "vision-model",
        "1.0.0",
        strategy="full-uia",
        root=tmp_path,
    )

    directory = knowledge.app_dir("example-app", tmp_path)
    controls = knowledge.load_index(directory)["controls"]
    button_record = next(control for control in controls if control["supported_actions"])
    assert result["semantic_counts"]["uncertain"] == 1
    assert button_record["status"] == "quarantined"
    assert button_record["semantic_ambiguity"] == "Icon meaning is unclear."
    assert not knowledge.available_commands(directory)
