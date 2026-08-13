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
    assert result["detector"] == "local-class-agnostic-rectangles-v2"
    assert result["layout_model"] == "size-independent-spatial-relations-v1"
    assert any(relation["type"] == "right-of" for relation in result["relations"])
    assert any(relation["type"] == "below" for relation in result["relations"])
    assert result["groups"][0]["size_policy"] == "variable"


def _layout_for_variable_content(name_width: int, message_height: int) -> tuple[list, list]:
    boxes = [
        {
            "left": 20,
            "top": 10,
            "right": 56,
            "bottom": 46,
            "width": 36,
            "height": 36,
            "sources": ["background-difference"],
        },
        {
            "left": 70,
            "top": 10,
            "right": 70 + name_width,
            "bottom": 24,
            "width": name_width,
            "height": 14,
            "sources": ["grouped-edge"],
        },
        {
            "left": 70,
            "top": 40,
            "right": 250,
            "bottom": 40 + message_height,
            "width": 180,
            "height": message_height,
            "sources": ["background-difference"],
        },
    ]
    visual_segmentation._assign_hierarchy(boxes)
    relations = visual_segmentation._layout_graph(boxes)
    groups = visual_segmentation._layout_groups(boxes, relations, (400, 240))
    return relations, groups


def test_layout_signature_ignores_variable_name_width_and_message_height() -> None:
    short_relations, short_groups = _layout_for_variable_content(30, 30)
    long_relations, long_groups = _layout_for_variable_content(100, 100)

    assert [relation["type"] for relation in short_relations] == [
        relation["type"] for relation in long_relations
    ]
    assert short_groups[0]["layout_signature"] == long_groups[0]["layout_signature"]
    assert short_groups[0]["rect_in_input"]["height"] != long_groups[0]["rect_in_input"]["height"]
    assert short_groups[0]["size_policy"] == "variable"


def test_vertical_repeated_items_are_not_merged_into_one_group() -> None:
    boxes = [
        {
            "left": 10,
            "top": top,
            "right": 210,
            "bottom": top + height,
            "width": 200,
            "height": height,
            "sources": ["background-difference"],
        }
        for top, height in ((10, 20), (40, 35), (85, 50))
    ]
    visual_segmentation._assign_hierarchy(boxes)
    relations = visual_segmentation._layout_graph(boxes)
    groups = visual_segmentation._layout_groups(boxes, relations, (300, 200))

    assert relations
    assert all(relation["type"] == "below" for relation in relations)
    assert groups == []


def test_table_rows_remain_separate_despite_vertical_column_relations() -> None:
    boxes = []
    for top, height in ((10, 20), (42, 34)):
        for left, width in ((10, 30), (50, 60), (120, 45)):
            boxes.append(
                {
                    "left": left,
                    "top": top,
                    "right": left + width,
                    "bottom": top + height,
                    "width": width,
                    "height": height,
                    "sources": ["background-difference"],
                }
            )
    visual_segmentation._assign_hierarchy(boxes)
    relations = visual_segmentation._layout_graph(boxes)
    groups = visual_segmentation._layout_groups(boxes, relations, (220, 120))

    assert len(groups) == 2
    assert groups[0]["member_ids"] == [1, 2, 3]
    assert groups[1]["member_ids"] == [4, 5, 6]
    assert groups[0]["layout_signature"] == groups[1]["layout_signature"]


def test_repeated_composite_rows_attach_variable_content_to_nearest_row() -> None:
    boxes = []
    for row_top, message_height in ((10, 35), (80, 70)):
        boxes.extend(
            [
                {
                    "left": 20,
                    "top": row_top,
                    "right": 56,
                    "bottom": row_top + 36,
                    "width": 36,
                    "height": 36,
                    "sources": ["background-difference"],
                },
                {
                    "left": 70,
                    "top": row_top,
                    "right": 105,
                    "bottom": row_top + 14,
                    "width": 35,
                    "height": 14,
                    "sources": ["grouped-edge"],
                },
                {
                    "left": 70,
                    "top": row_top + 20,
                    "right": 260,
                    "bottom": row_top + 20 + message_height,
                    "width": 190,
                    "height": message_height,
                    "sources": ["background-difference"],
                },
            ]
        )
    visual_segmentation._assign_hierarchy(boxes)
    relations = visual_segmentation._layout_graph(boxes)
    groups = visual_segmentation._layout_groups(boxes, relations, (320, 220))

    assert len(groups) == 2
    assert groups[0]["member_ids"] == [1, 2, 3]
    assert groups[1]["member_ids"] == [4, 5, 6]
    assert groups[0]["layout_signature"] == groups[1]["layout_signature"]


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
