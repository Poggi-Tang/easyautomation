"""Fit existing graphic observations to nearby closed rectangular contours."""

from __future__ import annotations

import copy
import time
from pathlib import Path


def _iou(first, second):
    overlap = max(0, min(first["right"], second["right"]) - max(first["left"], second["left"])) * max(
        0, min(first["bottom"], second["bottom"]) - max(first["top"], second["top"]))
    union = first["width"] * first["height"] + second["width"] * second["height"] - overlap
    return overlap / union if union > 0 else 0.0


def snap_rectangles(image, elements: list[dict], *, refine_ids=None, box_padding=1):
    """Apply the supplied threshold/IoU/growth rule without changing text or IDs.

    The contour must actually enclose a mostly rectangular area: a bounding box
    around disconnected text strokes is not evidence of a control border.
    Raw detections and unmatched observations are retained for review.
    """
    import cv2
    import numpy as np

    from . import diagram_geometry, image_analysis

    started = time.perf_counter()
    gray = cv2.cvtColor(np.asarray(image.convert("RGB")), cv2.COLOR_RGB2GRAY)
    _value, foreground = cv2.threshold(gray, 210, 255, cv2.THRESH_BINARY_INV)
    contours, _hierarchy = cv2.findContours(foreground, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    bounds = set()
    for contour in contours:
        x, y, width, height = cv2.boundingRect(contour)
        if not (30 <= width <= 600 and 12 <= height <= 60):
            continue
        if cv2.contourArea(contour) / max(1, (width - 1) * (height - 1)) < 0.7:
            continue
        bounds.add((x, y, width, height))
    candidates = [dict(left=x, top=y, right=x+w, bottom=y+h, width=w, height=h)
                  for x, y, w, h in sorted(bounds, key=lambda r: (r[1], r[0], r[2], r[3]))]
    selected = set(refine_ids) if refine_ids is not None else None
    result, updates = copy.deepcopy(elements), []
    for element in result:
        if diagram_geometry.is_text(element) or (selected is not None and element["id"] not in selected):
            continue
        original = dict(element["rect"])
        original.setdefault("width", original["right"] - original["left"])
        original.setdefault("height", original["bottom"] - original["top"])
        best = max(candidates, key=lambda c: _iou(original, c), default=None)
        overlap = _iou(original, best) if best else 0.0
        detail = "No rectangular contour reaches IoU 0.35"
        matched = best is not None and overlap >= 0.35
        if matched and (best["width"] > original["width"] * 1.6
                        or best["height"] > original["height"] * 1.6):
            matched = False
            detail = "Candidate exceeds the 1.6x width/height growth limit"
        if matched:
            detail = "Snap to the highest-IoU closed rectangle within the growth limit"
            element.setdefault("detected_rect", dict(element["rect"]))
            element["corrected_rect"] = dict(best)
            element["geometry_padding"] = box_padding
            element["geometry_kind"] = "yolo-snapped" if element.get("source") == "yolo" else "rectangle-snapped"
            element["snap_iou"] = round(overlap, 4)
            element["geometry_correction"] = {"method": "rectangle-contour", "detail": detail,
                                               "candidate_rect": dict(best), "iou": round(overlap, 4)}
            if element.get("ambiguity") == "YOLO region has no fitted component outline":
                element["pre_snap_ambiguity"] = element["ambiguity"]
                element["ambiguity"] = "Rectangular border fitted; control identity needs review"
        updates.append(dict(id=element["id"], annotation_label=element.get("annotation_label"),
                            method="rectangle-contour", before=original,
                            candidate_rect=dict(best) if best else None,
                            iou=round(overlap, 4), matched=matched, applied=False, detail=detail))
    # Recompute bbox, centre and normalized geometry from corrected_rect together.
    recomputed = image_analysis.complete_element_bounds(result, *image.size, box_padding)
    matched_ids = {u["id"] for u in updates if u["matched"]}
    result = [new if old["id"] in matched_ids else old
              for old, new in zip(result, recomputed, strict=True)]
    by_id = {element["id"]: element for element in result}
    for update in updates:
        update["after"] = dict(by_id[update["id"]]["rect"])
        update["applied"] = update["matched"] and update["before"] != update["after"]
    return result, {"enabled": True, "method": "rectangle-contour",
                    "candidate_count": len(candidates), "candidates": candidates,
                    "parameters": dict(threshold=210, min_width=30, max_width=600,
                                       min_height=12, max_height=60, min_iou=0.35,
                                       max_growth=1.6, min_rectangularity=0.7,
                                       box_padding=box_padding),
                    "updates": updates,
                    "matched_count": sum(u["matched"] for u in updates),
                    "applied_count": sum(u["applied"] for u in updates),
                    "skipped_count": sum(not u["matched"] for u in updates),
                    "elapsed_seconds": round(time.perf_counter() - started, 4)}


def save_snap_comparison(image, before: list[dict], report: dict, directory: Path) -> str:
    """Red original frames, green fitted frames, blue unchanged frames; 1px strokes."""
    from PIL import ImageDraw

    from . import visualization

    canvas = image.convert("RGB").copy()
    draw = ImageDraw.Draw(canvas)
    changed = {u["id"]: u for u in report["updates"]}
    markers = []
    for element in before:
        update = changed.get(element["id"])
        if update is None:
            continue
        if update["applied"]:
            old = update["before"]
            draw.rectangle((old["left"], old["top"], old["right"]-1, old["bottom"]-1),
                           outline="#dc0000", width=1)
        markers.append({"index": element["id"], "rect": update["after"],
                        "annotation_label": element.get("annotation_label", str(element["id"])),
                        "color": "#00a000" if update["applied"] else "#1565c0"})
    annotated = visualization.annotate_screenshot(
        canvas, markers, avoid_label_overlap=True, outline_width=1, compact_labels=True)
    path = directory / "snap-compare.png"
    annotated.save(path)
    return str(path)
