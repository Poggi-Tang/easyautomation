"""Conservative pixel-matrix cropping of a rectangular window inside a screenshot."""

from __future__ import annotations

from .visual_segmentation import _candidate_masks, _contour_candidates, normalize_screen_rect


def _frame_stroke_bounds(rgb, rectangle):
    """Snap a strong one-pixel outer stroke without depending on its colour."""
    import numpy

    height, width = rgb.shape[:2]
    found = []
    for side in ("left", "top", "right", "bottom"):
        horizontal = side in {"top", "bottom"}
        seed = rectangle[side] - (1 if side in {"right", "bottom"} else 0)
        limit = height if horizontal else width
        start = rectangle["left"] + 5 if horizontal else rectangle["top"] + 5
        end = rectangle["right"] - 5 if horizontal else rectangle["bottom"] - 5
        strokes = []
        for position in range(max(1, seed - 2), min(limit - 1, seed + 3)):
            line = rgb[position, start:end] if horizontal else rgb[start:end, position]
            before = rgb[position - 1, start:end] if horizontal else rgb[start:end, position - 1]
            after = rgb[position + 1, start:end] if horizontal else rgb[start:end, position + 1]
            colour = numpy.median(line, axis=0)
            uniform = (numpy.abs(line - colour).max(axis=1) <= 4).mean()
            inside, outside = (after, before) if side in {"left", "top"} else (before, after)
            inner_contrast = numpy.abs(line - inside).max(axis=1)
            outer_contrast = numpy.abs(line - outside).max(axis=1)
            # A coloured shadow can closely match the frame. Require the strong
            # contrast on the client side while allowing a weaker exterior edge.
            if (uniform >= 0.95 and float((inner_contrast >= 20).mean()) >= 0.95
                    and float((outer_contrast >= 6).mean()) >= 0.95):
                strokes.append((float(numpy.median(inner_contrast + outer_contrast)), position))
        if not strokes:
            return None
        found.append(max(strokes)[1] + (1 if side in {"right", "bottom"} else 0))
    left, top, right, bottom = found
    return dict(left=left, top=top, right=right, bottom=bottom)


def crop_window_image(image, screen_rect: dict | None = None, *, max_margin: int = 64,
                      padding: int = 0) -> dict:
    """Find a near-full window frame, including when the outside background is busy.

    A small control or a text bounding box is never a window-crop candidate.
    A busy exterior requires a continuous frame, balanced acquisition margins
    and a broad title-bar transition, rather than treating exterior pixels as
    expendable whitespace. Quiet exteriors retain the stricter content check.
    padding adds optional exterior pixels after the visible frame is determined.
    Ambiguous images are returned unchanged. All coordinates use exclusive
    right/bottom bounds; the result includes both local and screen offsets.
    """
    import cv2
    import numpy

    if max_margin < 0 or padding < 0:
        raise ValueError("max_margin and padding must be non-negative")
    width, height = image.size
    origin = normalize_screen_rect(screen_rect or dict(left=0, top=0, width=width, height=height))
    if (origin["width"], origin["height"]) != image.size:
        raise ValueError("image size must match screen_rect")
    full = dict(left=0, top=0, right=width, bottom=height, width=width, height=height)

    def result(rect, confidence=0.0, evidence="no-reliable-rectangle"):
        changed = rect != full
        absolute = {**rect, "left": origin["left"] + rect["left"],
                    "top": origin["top"] + rect["top"],
                    "right": origin["left"] + rect["right"],
                    "bottom": origin["top"] + rect["bottom"]}
        return {"image": image.crop((rect["left"], rect["top"], rect["right"], rect["bottom"]))
                if changed else image,
                "cropped": changed, "rect_in_input": rect, "screen_rect": absolute,
                "offset": {"x": rect["left"], "y": rect["top"]},
                "confidence": round(confidence, 4), "evidence": evidence,
                "method": "rectangle-edge" if changed else "unchanged"}

    if width < 32 or height < 32 or max_margin == 0:
        return result(full)
    masks = _candidate_masks(image)
    edges = masks[1][0] != 0
    rgb = numpy.array(image.convert("RGB"), dtype="int16")
    # A robust corner colour estimates the external background without treating
    # the application's own white client area as disposable content.
    corner = max(2, min(8, width // 8, height // 8))
    background = numpy.median(numpy.concatenate([
        rgb[:corner, :corner].reshape(-1, 3), rgb[:corner, -corner:].reshape(-1, 3),
        rgb[-corner:, :corner].reshape(-1, 3), rgb[-corner:, -corner:].reshape(-1, 3),
    ]), axis=0)
    distance = numpy.max(numpy.abs(rgb - background), axis=2)
    candidates = []
    for mask, source in masks[:2]:
        candidates.extend(_contour_candidates(
            mask, source, image.size, max(24, width - max_margin * 2),
            max(24, height - max_margin * 2),
        ))
    accepted = []
    for candidate in candidates:
        left, top, right, bottom = (candidate[key] for key in ("left", "top", "right", "bottom"))
        margins = (left, top, width - right, height - bottom)
        if min(margins) < 2 or max(margins) > max_margin or candidate["area_ratio"] < 0.6:
            continue
        if candidate.get("fill_ratio", 0) < 0.9:
            continue
        # Each of the four sides must contain a mostly continuous edge. Rounded
        # corners and one-pixel antialiasing are tolerated by a narrow band.
        supports = [
            edges[max(0, top - 2):min(height, top + 3), left:right].any(axis=0).mean(),
            edges[max(0, bottom - 3):min(height, bottom + 2), left:right].any(axis=0).mean(),
            edges[top:bottom, max(0, left - 2):min(width, left + 3)].any(axis=1).mean(),
            edges[top:bottom, max(0, right - 3):min(width, right + 2)].any(axis=1).mean(),
        ]
        if min(supports) < 0.8:
            continue
        outside = numpy.ones((height, width), dtype=bool)
        outside[max(0, top - 3):min(height, bottom + 3),
                max(0, left - 3):min(width, right + 3)] = False
        outside[:2] = outside[-2:] = False
        outside[:, :2] = outside[:, -2:] = False
        quiet = float((distance[outside] <= 32).mean()) if outside.sum() >= 64 else 0.0
        quiet_border = quiet >= 0.98 and float(edges[outside].mean()) <= 0.01
        # Even a short window title can occupy very few pixels. Reject connected
        # strokes outside the candidate instead of deleting them as "mostly blank".
        if quiet_border:
            _count, _labels, stats, _centres = cv2.connectedComponentsWithStats(
                (edges & outside).astype("uint8"), connectivity=8,
            )
            quiet_border = len(stats) == 1 or int(stats[1:, cv2.CC_STAT_AREA].max()) < 8

        # A captured window can be surrounded by a desktop, text or another app.
        # Select its outer frame rather than requiring those exterior pixels to
        # be blank. A title-bar boundary must span most of the window width.
        start = top + 10
        stop = min(bottom - 8, top + 80, top + candidate["height"] // 4)
        caption_support = 0.0
        if stop > start and right - left > 16:
            # Compare bands, not adjacent scanlines: a thin content-panel border
            # must not masquerade as a filled title bar transitioning to a client area.
            above = numpy.median(numpy.stack([
                rgb[start - offset:stop - offset, left + 5:right - 5]
                for offset in (1, 2, 3)
            ]), axis=0)
            below = numpy.median(numpy.stack([
                rgb[start + offset:stop + offset, left + 5:right - 5]
                for offset in (0, 1, 2)
            ]), axis=0)
            transition = numpy.max(numpy.abs(below - above), axis=2)
            caption_support = float((transition >= 6).mean(axis=1).max())
        balanced = max(margins) - min(margins) <= max(4, min(margins) * 0.4)
        window_frame = (candidate["area_ratio"] >= 0.75
                        and candidate["fill_ratio"] >= 0.95
                        and min(supports) >= 0.95 and balanced
                        and caption_support >= 0.85)
        if not (quiet_border or window_frame):
            continue
        # Canny's contour can fall just inside the visible stroke. Preserve a
        # one-pixel guard when it cannot be snapped to all four actual frame lines.
        frame = _frame_stroke_bounds(rgb, candidate) if window_frame else None
        if frame is None:
            frame = dict(left=left - 1, top=top - 1, right=right + 1, bottom=bottom + 1)
        rectangle = dict(left=max(0, frame["left"] - padding), top=max(0, frame["top"] - padding),
                         right=min(width, frame["right"] + padding), bottom=min(height, frame["bottom"] + padding))
        rectangle.update(width=rectangle["right"] - rectangle["left"],
                         height=rectangle["bottom"] - rectangle["top"])
        accepted.append((rectangle["width"] * rectangle["height"], rectangle,
                         min(0.99, 0.5 + float(numpy.mean(supports)) * 0.4
                             + max(quiet if quiet_border else 0, caption_support) * 0.1),
                         "window-frame-and-titlebar" if window_frame else "quiet-outer-border"))
    if not accepted:
        return result(full)
    _area, rectangle, confidence, evidence = max(accepted, key=lambda value: value[0])
    return result(rectangle, confidence, evidence)
