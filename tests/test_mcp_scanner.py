"""Tests for application scanning and control-image verification."""

from __future__ import annotations

from types import SimpleNamespace

from PIL import Image, ImageDraw

from easy_uiauto.mcp import knowledge, scanner


class FakeControl:
    def __init__(
        self,
        name: str,
        control_type: str,
        rect: tuple[int, int, int, int],
        children: list[FakeControl] | None = None,
        automation_id: str = "",
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

    def GetChildren(self):
        return self._children


def _test_screen() -> Image.Image:
    image = Image.new("RGB", (400, 300), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((40, 30, 160, 90), fill="blue")
    draw.line((45, 35, 155, 85), fill="white", width=4)
    return image


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
        lambda _location, rect: (True, rect, "bounds IoU=1.000"),
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
    assert list((directory / "images" / "controls").glob("*.png"))
    assert (directory / "operations" / "UI-CLI.md").is_file()


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
        lambda _location, _rect: (False, {}, "LOCATION did not resolve"),
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
        lambda _location, rect: (True, rect, "bounds IoU=1.000"),
    )

    result = scanner.scan_window(
        "Example",
        "https://api.example/v1",
        "secret",
        "vision-model",
        "1.0.0",
        root=tmp_path,
    )

    directory = knowledge.app_dir("example-app", tmp_path)
    controls = knowledge.load_index(directory)["controls"]
    button_record = next(control for control in controls if control["supported_actions"])
    assert result["semantic_counts"]["uncertain"] == 1
    assert button_record["status"] == "quarantined"
    assert button_record["semantic_ambiguity"] == "Icon meaning is unclear."
    assert not knowledge.available_commands(directory)
