"""Tests for class-agnostic rectangle detection in coarse UIA canvases."""

from __future__ import annotations

import json

from PIL import Image, ImageDraw

from easy_uiauto.mcp import server, visual_segmentation


def _canvas() -> Image.Image:
    image = Image.new("RGB", (320, 160), (250, 250, 250))
    draw = ImageDraw.Draw(image)
    draw.rectangle((18, 20, 57, 59), fill=(110, 150, 210))
    draw.rectangle((76, 20, 118, 33), fill=(80, 80, 80))
    draw.rounded_rectangle((70, 48, 285, 105), radius=8, fill=(238, 238, 240))
    draw.rectangle((84, 65, 250, 78), fill=(55, 55, 55))
    return image


def _matching_box(boxes: list[dict], expected: tuple[int, int, int, int], tolerance: int = 4):
    left, top, right, bottom = expected
    for box in boxes:
        rectangle = box["rect_in_input"]
        if max(
            abs(rectangle["left"] - left),
            abs(rectangle["top"] - top),
            abs(rectangle["right"] - right),
            abs(rectangle["bottom"] - bottom),
        ) <= tolerance:
            return box
    return None


def test_normalize_screen_rect_accepts_control_bounds_wrapper() -> None:
    assert visual_segmentation.normalize_screen_rect(
        {"Name": "Canvas", "bounds": {"left": 100, "top": 200, "right": 500, "bottom": 600}}
    ) == {
        "left": 100,
        "top": 200,
        "right": 500,
        "bottom": 600,
        "width": 400,
        "height": 400,
    }


def test_detect_rectangles_finds_canvas_elements_and_nested_content() -> None:
    result = visual_segmentation.detect_rectangles(
        _canvas(),
        {"left": 100, "top": 200, "width": 320, "height": 160},
    )

    avatar = _matching_box(result["boxes"], (18, 20, 58, 60))
    name = _matching_box(result["boxes"], (76, 20, 119, 34))
    bubble = _matching_box(result["boxes"], (70, 48, 286, 106))
    content = _matching_box(result["boxes"], (84, 65, 251, 79))

    assert avatar is not None
    assert name is not None
    assert bubble is not None
    assert content is not None
    assert content["parent_id"] == bubble["id"]
    assert bubble["kind"] == "container"
    assert bubble["child_count"] >= 1
    assert content["kind"] == "element"
    assert 0.5 <= content["confidence"] <= 1.0
    assert bubble["rect"]["left"] == 170
    assert bubble["rect"]["top"] == 248
    assert bubble["rect_relative_to_input"]["left"] == 0.21875
    assert result["semantic_analysis"] is False
    assert result["detector"] == "local-class-agnostic-rectangles-v1"


def test_detect_visual_elements_tool_uses_only_supplied_rect(monkeypatch) -> None:
    received = {}

    def detect(rect, **options):
        received["rect"] = rect
        received.update(options)
        return {"ok": True, "box_count": 2, "boxes": []}

    monkeypatch.setattr(server.visual_segmentation, "detect_screen_rectangles", detect)

    result = json.loads(
        server.detect_visual_elements(
            {"bounds": {"left": 10, "top": 20, "right": 110, "bottom": 80}},
            min_width=4,
            min_height=5,
            max_boxes=20,
            save_annotated=False,
        )
    )

    assert result["ok"] is True
    assert received == {
        "rect": {"bounds": {"left": 10, "top": 20, "right": 110, "bottom": 80}},
        "min_width": 4,
        "min_height": 5,
        "max_boxes": 20,
        "save_annotated": False,
    }
