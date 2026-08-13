"""Tests for the public non-MCP visual segmentation API."""

from PIL import Image, ImageDraw

from easy_uiauto import segment_image
from easy_uiauto.visual import segment_screen_rect


def test_segment_image_uses_image_dimensions_by_default() -> None:
    image = Image.new("RGB", (160, 90), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((20, 20, 60, 60), fill="black")

    result = segment_image(image)

    assert result["ok"] is True
    assert result["input_rect"] == {
        "left": 0,
        "top": 0,
        "right": 160,
        "bottom": 90,
        "width": 160,
        "height": 90,
    }
    assert result["box_count"] >= 1


def test_segment_screen_rect_is_public_without_mcp_server(monkeypatch) -> None:
    from easy_uiauto import visual

    expected = {"ok": True, "boxes": []}
    monkeypatch.setattr(visual, "detect_screen_rectangles", lambda *_args, **_kwargs: expected)

    rect = {"bounds": {"left": 1, "top": 2, "width": 30, "height": 40}}
    assert segment_screen_rect(rect) is expected
