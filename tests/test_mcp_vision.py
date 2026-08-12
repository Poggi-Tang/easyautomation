"""Tests for image-template and OCR fallback helpers."""

from __future__ import annotations

import json
from types import SimpleNamespace

from easy_uiauto.mcp import configuration, server


def test_locate_image_returns_bounds(monkeypatch, tmp_path) -> None:
    template = tmp_path / "button.png"
    template.write_bytes(b"placeholder")
    received: dict = {}

    def fake_locate(path, **kwargs):
        received["path"] = path
        received.update(kwargs)
        return SimpleNamespace(left=10, top=20, width=40, height=30)

    monkeypatch.setattr(server.pyautogui, "locateOnScreen", fake_locate)

    match = server._locate_image(str(template), 0.9, True, (1, 2, 300, 200))

    assert match == {
        "left": 10,
        "top": 20,
        "width": 40,
        "height": 30,
        "center_x": 30,
        "center_y": 35,
    }
    assert received["confidence"] == 0.9
    assert received["grayscale"] is True
    assert received["region"] == (1, 2, 300, 200)


def test_find_text_on_screen_applies_region_offset(monkeypatch) -> None:
    fake_tesseract = SimpleNamespace(
        Output=SimpleNamespace(DICT="dict"),
        image_to_data=lambda *_args, **_kwargs: {
            "text": ["Save"],
            "conf": ["88.5"],
            "left": [10],
            "top": [20],
            "width": [40],
            "height": [30],
        },
    )
    monkeypatch.setattr(server, "_require_pytesseract", lambda: fake_tesseract)
    monkeypatch.setattr(server.pyautogui, "screenshot", lambda **_kwargs: object())

    match = server._find_text_on_screen("save", "eng", "exact", 60, (100, 200, 400, 300))

    assert match == {
        "text": "Save",
        "confidence": 88.5,
        "left": 110,
        "top": 220,
        "width": 40,
        "height": 30,
        "center_x": 130,
        "center_y": 235,
    }


def test_vision_response_parses_json_fence() -> None:
    assert server._vision_response_json('```json\n{"found": false}\n```') == {"found": False}


def test_server_reads_vision_settings_dynamically(monkeypatch) -> None:
    values = {
        configuration.VISION_API_URL: "https://api.example/v1/chat/completions",
        configuration.VISION_API_KEY: "secret",
        configuration.VISION_MODEL: "vision-model",
    }
    monkeypatch.setattr(
        configuration,
        "_existing_vision_value",
        lambda name: values.get(name, ""),
    )

    assert server._vision_api_settings("") == (
        "https://api.example/v1/chat/completions",
        "secret",
        "vision-model",
    )


def test_learning_readiness_returns_no_secret(monkeypatch, tmp_path) -> None:
    monkeypatch.setattr(
        configuration,
        "vision_configuration_status",
        lambda: {
            "ready": True,
            "api_url_configured": True,
            "api_key_configured": True,
            "model_configured": True,
            "model": "vision-model",
            "detail": "ready",
        },
    )
    monkeypatch.setattr(server.knowledge, "vault_root", lambda: tmp_path)
    monkeypatch.setattr(server.knowledge, "list_apps", lambda: [])

    result = json.loads(server.get_ui_learning_readiness())

    assert result["ready"] is True
    assert result["knowledge_vault"] == str(tmp_path)
    assert "secret" not in json.dumps(result).casefold()
