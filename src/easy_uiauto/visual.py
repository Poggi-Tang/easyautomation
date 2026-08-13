"""Public local visual-segmentation API for ordinary Python applications."""

from __future__ import annotations

from .mcp.visual_segmentation import (
    detect_rectangles,
    detect_screen_rectangles,
    normalize_screen_rect,
)


def segment_screen_rect(
    rect: dict,
    min_width: int = 6,
    min_height: int = 6,
    max_boxes: int = 300,
    save_annotated: bool = False,
) -> dict:
    """Capture and segment one supplied screen rectangle locally.

    ``rect`` may be a direct ``left/top/right/bottom`` mapping or a result
    containing a ``bounds``, ``rect``, or ``rectangle`` mapping. No MCP server,
    remote API, OCR, or semantic model is used.
    """
    return detect_screen_rectangles(
        rect,
        min_width=min_width,
        min_height=min_height,
        max_boxes=max_boxes,
        save_annotated=save_annotated,
    )


def segment_image(
    image,
    screen_rect: dict | None = None,
    min_width: int = 6,
    min_height: int = 6,
    max_boxes: int = 300,
) -> dict:
    """Segment an existing Pillow image without capturing the desktop."""
    if screen_rect is None:
        screen_rect = {
            "left": 0,
            "top": 0,
            "width": image.width,
            "height": image.height,
        }
    return detect_rectangles(
        image,
        screen_rect,
        min_width=min_width,
        min_height=min_height,
        max_boxes=max_boxes,
    )


__all__ = ["normalize_screen_rect", "segment_image", "segment_screen_rect"]
