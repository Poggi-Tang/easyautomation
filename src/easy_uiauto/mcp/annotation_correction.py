"""Correct saved annotations locally, preserving model evidence and element IDs."""

from __future__ import annotations

import copy
import hashlib
import io
import json
import math
import time
from pathlib import Path
from uuid import uuid4

from . import image_analysis


def _rect(left, top, right, bottom):
    return {"left": int(left), "top": int(top), "right": int(right), "bottom": int(bottom),
            "width": int(right - left), "height": int(bottom - top)}


def _manual_rect(value, width, height):
    if not isinstance(value, dict):
        raise ValueError("Each correction needs a rect object")
    keys = ("left", "top", "right", "bottom")
    if any(type(value.get(key)) is not int for key in keys):
        raise ValueError("rect needs integer left, top, right, bottom (exclusive)")
    left, top, right, bottom = (value[key] for key in keys)
    if not (0 <= left < right <= width and 0 <= top < bottom <= height):
        raise ValueError("Correction rectangle must stay inside the original image")
    return _rect(left, top, right, bottom)


def _seed_rect(item):
    if item.get("geometry_correction", {}).get("method") == "explicit":
        return item["corrected_rect"]
    return item.get("detected_rect", item["rect"])


def _pixel_bounds(image, item, elements, margin):
    """Fit isolated icons on light backgrounds; decline connected/ambiguous crops."""
    import cv2
    import numpy as np

    seed = _seed_rect(item)
    width, height = image.size
    left, top = max(0, seed["left"] - margin), max(0, seed["top"] - margin)
    right, bottom = min(width, seed["right"] + margin), min(height, seed["bottom"] + margin)
    cx, cy = (seed["left"] + seed["right"]) / 2, (seed["top"] + seed["bottom"]) / 2
    family = {item["id"], item.get("parent_element_id"),
              *item.get("component_element_ids", [])}
    # Existing identities define ownership boundaries. A small gap alone never
    # merges independently numbered controls.
    for other in elements:
        if other["id"] in family:
            continue
        rect = _seed_rect(other)
        ox, oy = (rect["left"] + rect["right"]) / 2, (rect["top"] + rect["bottom"]) / 2
        if min(seed["bottom"], rect["bottom"]) > max(seed["top"], rect["top"]):
            if ox < cx:
                boundary = ((rect["right"] + seed["left"]) / 2
                            if rect["right"] <= seed["left"] else (ox + cx) / 2)
                left = max(left, math.floor(boundary))
            elif ox > cx:
                boundary = ((rect["left"] + seed["right"]) / 2
                            if rect["left"] >= seed["right"] else (ox + cx) / 2)
                right = min(right, math.ceil(boundary))
        if min(seed["right"], rect["right"]) > max(seed["left"], rect["left"]):
            if oy < cy:
                top = max(top, math.floor((rect["bottom"] + seed["top"]) / 2))
            elif oy > cy:
                bottom = min(bottom, math.ceil((rect["top"] + seed["bottom"]) / 2))
    if right <= left or bottom <= top:
        return None, "Conflicting neighboring bounds; supply an explicit rectangle"
    if (right - left) * (bottom - top) > 512 * 512:
        return None, "Search region is too large for isolated-icon fitting; use explicit bounds"
    rgb = np.asarray(image.crop((left, top, right, bottom)).convert("RGB")).astype("int16")
    border = np.concatenate((rgb[0], rgb[-1], rgb[:, 0], rgb[:, -1]))
    background = np.median(border, axis=0)
    if background.min() < 200:
        return None, "Automatic fitting currently requires a light, plain background"
    distance = np.max(np.abs(rgb - background), axis=2)
    strong = distance >= 96
    weak = distance >= 6
    # Remove pale toolbar frames without removing dark/colored icon strokes.
    frame = weak & ~strong
    broad = cv2.morphologyEx(frame.astype("uint8"), cv2.MORPH_OPEN,
                            np.ones((3, 3), dtype="uint8")).astype(bool)
    thin = frame & ~broad
    weak[thin.sum(axis=1) >= 0.65 * rgb.shape[1], :] = False
    weak[:, thin.sum(axis=0) >= 0.65 * rgb.shape[0]] = False
    weak |= strong
    sx1, sy1 = max(0, seed["left"] - left), max(0, seed["top"] - top)
    sx2, sy2 = min(right - left, seed["right"] - left), min(bottom - top, seed["bottom"] - top)
    if sx1 >= sx2 or sy1 >= sy2:
        return None, "No seed pixels remain inside the search region"
    if item.get("parent_element_id") is not None:
        # A badge can touch its parent. Restrict it to the seed's color instead
        # of absorbing the parent's connected artwork.
        seed_pixels = rgb[sy1:sy2, sx1:sx2][strong[sy1:sy2, sx1:sx2]]
        if len(seed_pixels) < 2:
            return None, "No distinct badge-color seed; supply an explicit rectangle"
        color = np.median(seed_pixels, axis=0)
        weak &= np.max(np.abs(rgb - color), axis=2) <= 80
        strong &= weak
    if not strong[sy1:sy2, sx1:sx2].any():
        return None, "No reliable foreground seed; supply an explicit rectangle"
    _count, labels, stats, _centroids = cv2.connectedComponentsWithStats(
        weak.astype("uint8"), connectivity=8)
    selected = []
    for label, (x, y, w, h, area) in enumerate(stats[1:], start=1):
        if area < 2 or not strong[labels == label].any():
            continue
        overlap = int((labels[sy1:sy2, sx1:sx2] == label).sum())
        selected.append((label, overlap, (int(x), int(y), int(x + w), int(y + h))))
    if item.get("parent_element_id") is not None:
        selected = sorted(selected, key=lambda value: value[1], reverse=True)[:1]
    if not selected:
        return None, "No isolated icon component found"
    chosen = np.isin(labels, [value[0] for value in selected])
    solid = chosen & strong
    if ((left > 0 and solid[:, 0].any()) or (right < width and solid[:, -1].any())
            or (top > 0 and solid[0].any()) or (bottom < height and solid[-1].any())):
        return None, "Foreground reaches the crop boundary; use explicit bounds or a larger margin"
    ys, xs = np.where(chosen)
    return _rect(left + xs.min(), top + ys.min(), left + xs.max() + 1,
                 top + ys.max() + 1), "Fit visible foreground near the existing target"


def correct_annotations(coordinates_json: str, refine_ids: list[int] | None = None,
                        corrections: list[dict] | None = None, output_dir: str = "",
                        search_margin: int = 8, box_padding: int = 1,
                        profile: str = "icons", discover_missing: bool = False) -> dict:
    """Create a new PNG/JSON pair without a vision request or desktop access."""
    from PIL import Image, ImageOps

    started = time.perf_counter()
    if profile not in {"icons", "circuit"}:
        raise ValueError("profile must be icons or circuit")
    if discover_missing and profile != "circuit":
        raise ValueError("discover_missing is supported only with profile=circuit")
    if type(search_margin) is not int or not 1 <= search_margin <= 64:
        raise ValueError("search_margin must be an integer from 1 to 64")
    if type(box_padding) is not int or not 0 <= box_padding <= 20:
        raise ValueError("box_padding must be an integer from 0 to 20")
    path = Path(coordinates_json).expanduser().resolve(strict=True)
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or data.get("ok") is not True:
        raise ValueError("Supply the coordinates_json from a successful image analysis")
    items = data.get("elements")
    if not isinstance(items, list) or any(not isinstance(item, dict) for item in items):
        raise ValueError("elements must be an array of annotation objects")
    ids = [item.get("id") for item in items]
    if any(type(value) is not int for value in ids) or len(ids) != len(set(ids)):
        raise ValueError("Element IDs must be unique integers")
    refine_ids, corrections = refine_ids or [], corrections or []
    if not refine_ids and not corrections and not discover_missing:
        raise ValueError("Specify refine_ids for pixel fitting or corrections for exact bounds")
    manual_ids = [item.get("id") if isinstance(item, dict) else None for item in corrections]
    requested = [*refine_ids, *manual_ids]
    if any(type(value) is not int or value not in ids for value in requested):
        raise ValueError("Every requested ID must identify an existing annotation")
    if len(requested) != len(set(requested)):
        raise ValueError("Specify each ID once, using either pixel fitting or exact bounds")
    source = Path(data["image_path"]).expanduser().resolve(strict=True)
    source_bytes = source.read_bytes()
    source_hash = hashlib.sha256(source_bytes).hexdigest()
    if data.get("source_image_sha256") and data["source_image_sha256"] != source_hash:
        raise ValueError("Source image changed since analysis; analyze it again first")
    with Image.open(io.BytesIO(source_bytes)) as original:
        image = ImageOps.exif_transpose(original).convert("RGBA")
        image = Image.alpha_composite(Image.new("RGBA", image.size, "white"), image).convert("RGB")
    width, height = image.size
    if (width, height) != (data["image_width"], data["image_height"]):
        raise ValueError("Source image dimensions no longer match the annotation coordinates")
    explicit = {item["id"]: _manual_rect(item.get("rect"), width, height)
                for item in corrections}
    elements, updates = copy.deepcopy(items), []
    circuit_report = {}
    if profile == "circuit":
        from . import circuit_geometry

        elements, updates, circuit_report = circuit_geometry.correct(
            image, elements, refine_ids, box_padding, discover_missing)
    for item in elements:
        if item["id"] not in requested or (profile == "circuit" and item["id"] not in explicit):
            continue
        before = dict(item["rect"])
        if item["id"] in explicit:
            fitted, detail, method = explicit[item["id"]], "Explicit content bounds", "explicit"
        else:
            fitted, detail = _pixel_bounds(image, item, items, search_margin)
            method = "pixel-foreground"
        update = {"id": item["id"], "method": method, "before": before, "detail": detail,
                  "applied": fitted is not None}
        if fitted is not None:
            item.setdefault("detected_rect", dict(item["rect"]))
            item["corrected_rect"] = fitted
            item["geometry_padding"] = 0 if method == "explicit" else box_padding
            item["geometry_correction"] = {"method": method, "detail": detail}
        updates.append(update)
    elements = image_analysis.complete_element_bounds(
        elements, width, height, data.get("box_padding", 2))
    by_id = {item["id"]: item for item in elements}
    for update in updates:
        update["after"] = by_id[update.get("canonical_id", update["id"])]["rect"]
    base = Path(output_dir).expanduser() if output_dir else path.parent
    destination = (base / f"image-correction-{uuid4().hex[:12]}").resolve()
    destination.mkdir(parents=True, exist_ok=False)
    for item in elements:
        item.update(status="observed", uia_verified=False, executable=False)
    result = {**data, "elements": elements, "element_count": len(elements),
              "source_image_sha256": source_hash,
              "uia_verified": False, "commands_generated": 0,
              **image_analysis.save_image_annotations(image, elements, destination),
              "coordinates_json": str(destination / "elements.json"),
              "correction": {"source_coordinates_json": str(path), "updates": updates,
                             "profile": profile, "input_element_count": len(items),
                             **circuit_report,
                             "applied_count": sum(item["applied"] and item["before"] is not None
                                                  for item in updates),
                             "skipped_count": sum(not item["applied"] for item in updates),
                             "needs_review_ids": [item["id"] for item in updates
                                                  if not item["applied"]],
                             "source_validation": ("sha256" if data.get("source_image_sha256")
                                                   else "dimensions-only"),
                             "api_requests": 0,
                             "elapsed_seconds": round(time.perf_counter() - started, 3)}}
    Path(result["coordinates_json"]).write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result
