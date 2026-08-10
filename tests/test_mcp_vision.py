"""Tests for image-template and OCR fallback helpers."""

from __future__ import annotations

from types import SimpleNamespace

from easy_uiauto.mcp import server


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
