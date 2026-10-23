"""Local circuit-symbol geometry; no reference annotations or remote requests."""

from __future__ import annotations

import copy
import math
from collections import Counter


def rect(x, y, width, height):
    return {"left": int(x), "top": int(y), "right": int(x + width),
            "bottom": int(y + height), "width": int(width), "height": int(height)}


def expand(box, amount, size):
    left, top = max(0, box["left"] - amount), max(0, box["top"] - amount)
    right, bottom = min(size[0], box["right"] + amount), min(size[1], box["bottom"] + amount)
    return rect(left, top, right - left, bottom - top)


def iou(a, b):
    overlap = (max(0, min(a["right"], b["right"]) - max(a["left"], b["left"]))
               * max(0, min(a["bottom"], b["bottom"]) - max(a["top"], b["top"])))
    return overlap / max(1, a["width"] * a["height"] + b["width"] * b["height"] - overlap)


def kind(item):
    from . import diagram_geometry

    if diagram_geometry.is_text(item):
        return ""
    text = str(item.get("semantic_name", "")).lower()
    diagram_kind = diagram_geometry.kind(item)
    # PowerPort model blocks and small electrical terminals share a name, but
    # have different outlines. An explicit terminal takes precedence.
    if "powerport" in text and diagram_kind != "terminal":
        return "block"
    if diagram_kind:
        return diagram_kind
    for name in (str(item.get("entity_type", "")).lower(),
                 str(item.get("semantic_name", "")).lower()):
        for token, category in (("powerport", "block"), ("transformer", "transformer"),
                                ("变压器", "transformer"), ("load", "load"), ("负荷", "load"),
                                ("负载", "load"), ("line", "line"), ("线路", "line"),
                                ("busbar", "busbar"), ("母线", "busbar")):
            if token in name:
                return category
    return ""


def _contains(outer, inner):
    return (outer["left"] <= inner["left"] and outer["right"] >= inner["right"]
            and outer["top"] <= inner["top"] and outer["bottom"] >= inner["bottom"])


def _wire_thickness(mask, length):
    import cv2
    import numpy as np

    widths = []
    for kernel, transpose in (((1, length), True), ((length, 1), False)):
        opened = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones(kernel, dtype="uint8")) > 0
        rows = opened.T if transpose else opened
        changes = np.diff(np.pad(rows, ((0, 0), (1, 1))).ravel().astype("int8"))
        runs = np.where(changes == -1)[0] - np.where(changes == 1)[0]
        widths.extend(runs[(runs > 0) & (runs <= 12)].tolist())
    return int(np.bincount(widths).argmax()) if widths else 1


def _rectangle_holes(contours, minimum, maximum, minimum_fill=0.8, epsilon=0.015):
    import cv2

    found = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if min(w, h) < minimum or max(w, h) > maximum:
            continue
        polygon = cv2.approxPolyDP(contour, epsilon * cv2.arcLength(contour, True), True)
        if len(polygon) != 4 or cv2.contourArea(contour) / (w * h) < minimum_fill:
            continue
        points = polygon[:, 0, :]
        if any(min(abs(int(p[0] - q[0])), abs(int(p[1] - q[1]))) > 2
               for p, q in zip(points, [*points[1:], points[0]], strict=True)):
            continue
        box = rect(x, y, w, h)
        if not any(iou(box, other) > 0.95 for other in found):
            found.append(box)
    return found


def _circles(mask, contours):
    import cv2
    import numpy as np

    distance = cv2.distanceTransform(255 - mask, cv2.DIST_L2, 3)
    angles = np.linspace(0, 2 * np.pi, 180, endpoint=False)
    cosine, sine = np.cos(angles), np.sin(angles)

    def best(cx, cy, radius, x_offsets):
        xs, ys, rs = np.meshgrid(x_offsets, np.arange(-2, 3), np.arange(-1, 3), indexing="ij")
        xx, yy, rr = cx + xs.ravel(), cy + ys.ravel(), radius + rs.ravel()
        sample_x = np.rint(xx[:, None] + rr[:, None] * cosine).astype(int)
        sample_y = np.rint(yy[:, None] + rr[:, None] * sine).astype(int)
        sample_x = np.clip(sample_x, 0, mask.shape[1] - 1)
        sample_y = np.clip(sample_y, 0, mask.shape[0] - 1)
        errors = distance[sample_y, sample_x]
        scores = np.exp(-errors ** 2 / 1.28).mean(axis=1)
        index = int(scores.argmax())
        return xx[index], yy[index], rr[index], float(scores[index])

    found = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        fill = cv2.contourArea(contour) / max(1, w * h)
        if not (min(w, h) >= 20 and max(w, h) <= 180 and 0.65 <= w / h <= 1.4
                and 0.5 <= fill <= 0.85):
            continue
        (cx, cy), radius = cv2.minEnclosingCircle(contour)
        cx, cy, radius, score = best(cx, cy, radius, np.arange(-2, 3))
        if score < 0.75:
            continue
        offsets = np.concatenate((np.arange(-1.8 * radius, -0.75 * radius),
                                  np.arange(0.75 * radius, 1.8 * radius)))
        bx, by, br, partner_score = best(cx, cy, radius, offsets)
        if partner_score < 0.75:
            continue
        left, top = math.floor(min(cx - radius, bx - br)), math.floor(min(cy - radius, by - br))
        right, bottom = math.ceil(max(cx + radius, bx + br)), math.ceil(max(cy + radius, by + br))
        box = rect(left - 1, top - 1, right - left + 2, bottom - top + 2)
        if not any(iou(box, item["rect"]) > 0.75 for item in found):
            found.append({"kind": "transformer", "rect": box,
                          "evidence": "Two overlapping circles supported by original dark pixels",
                          "pixel_support": round(min(score, partner_score), 4)})
    return found


def _load(mask, resistor, markers):
    """Follow a resistor toward three diminishing ground bars on either side."""
    import numpy as np

    x, y, w, h = (resistor[key] for key in ("left", "top", "width", "height"))
    vertical = h > w
    length, short = max(w, h), min(w, h)
    axis, transverse = (y, x + w / 2) if vertical else (x, y + h / 2)
    axis_size, transverse_size = (mask.shape[0], mask.shape[1]) if vertical else mask.shape[::-1]
    low, high = max(0, axis - 4 * length), min(axis_size, axis + 5 * length)
    t1, t2 = max(0, int(transverse - 2 * short)), min(transverse_size, int(transverse + 2 * short))
    band = mask[low:high, t1:t2] if vertical else mask[t1:t2, low:high].T
    counts = (band > 0).sum(axis=1)
    active = counts >= max(3, short * 0.3)
    changes = np.diff(np.r_[False, active, False].astype(int))
    starts, ends = np.where(changes == 1)[0], np.where(changes == -1)[0]
    runs = [(int(a), int(b), int(counts[a:b].max())) for a, b in zip(starts, ends, strict=True)
            if b - a <= max(3, short * 0.4)]
    grounds = []
    for a, b, c in zip(runs, runs[1:], runs[2:], strict=False):
        if not (a[2] >= max(6, short * 0.7) and 0.25 < b[2] / a[2] < 0.85
                and 0.2 < c[2] / b[2] < 0.85 and b[0] - a[1] <= short
                and c[0] - b[1] <= short):
            continue
        first, last = low + a[0], low + c[1]
        if last < axis or first > axis + length:
            grounds.append((min(abs(first - axis), abs(last - axis - length)), first, last))
    if not grounds:
        return None
    _distance, first, last = min(grounds)
    start, end = min(axis - 1, first), max(axis + length + 1, last)
    visible = band[max(0, start - low):min(high - low, end - low)] > 0
    yy, xx = np.where(visible)
    if not len(xx):
        return None
    leading = first > axis + length
    if leading:
        start -= 3
    else:
        end += 3
    transverse_low, transverse_high = t1 + int(xx.min()), t1 + int(xx.max()) + 1
    if vertical:
        box = rect(transverse_low, start, transverse_high - transverse_low, end - start)
    else:
        box = rect(start, transverse_low, end - start, transverse_high - transverse_low)
    for marker in markers:
        mx, my = (marker["left"] + marker["right"]) / 2, (marker["top"] + marker["bottom"]) / 2
        along, across = (my, mx) if vertical else (mx, my)
        leading_distance = axis - along if leading else along - axis - length
        if 0 < leading_distance < length * 0.7 and abs(across - transverse) < short * 1.5:
            left, top = min(box["left"], marker["left"]), min(box["top"], marker["top"])
            right, bottom = max(box["right"], marker["right"]), max(box["bottom"], marker["bottom"])
            box = rect(left, top, right - left, bottom - top)
    return box


def candidates(image, block_diagram=False):
    """Find supported frames, twin circles, thick busbars and grounded RLC branches."""
    import cv2
    import numpy as np

    from . import diagram_geometry

    rgb = np.asarray(image.convert("RGB"))
    maximum, minimum = rgb.max(axis=2), rgb.min(axis=2)
    mask = ((maximum < 170) & (maximum.astype("int16") - minimum < 80)).astype("uint8") * 255
    contours, _hierarchy = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    min_size = max(16, round(min(image.size) * 0.035))
    boxes = _rectangle_holes(contours, min_size, min(image.size) * 0.3)
    outer = [box for box in boxes if not any(
        other != box and _contains(other, box) for other in boxes)]
    found = []
    for box in outer:
        nested = any(other != box and _contains(box, other) for other in boxes)
        found.append({"kind": "block" if nested else "line",
                      "rect": expand(box, 2, image.size),
                      "evidence": "Closed rectangular frame in original pixels"})
    outlines = diagram_geometry.shapes(image)
    if block_diagram:
        found = list(outlines)
    resistors = [box for box in _rectangle_holes(
        contours, 5, max(40, min_size * 2.2), minimum_fill=0.65, epsilon=0.02)
                 if min(box["width"], box["height"]) <= max(14, min_size * 0.8)
                 and max(box["width"], box["height"]) / min(box["width"], box["height"]) > 1.8]
    # Thick strokes inside symbol frames/resistors are not busbars.
    length = max(18, round(min(image.size) * 0.04))
    wire = _wire_thickness(mask, length)
    exclusions = [expand(box, 3 * wire, image.size)
                  for box in [*outer, *resistors, *[c["rect"] for c in outlines]]]
    for kernel in ((2 * wire, length), (length, 2 * wire)):
        opened = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones(kernel, dtype="uint8"))
        components, _ = cv2.findContours(opened, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for contour in components:
            x, y, w, h = cv2.boundingRect(contour)
            if max(w, h) / min(w, h) < 10:
                continue
            box = rect(x, y, w, h)
            if any(_contains(other, box) for other in exclusions):
                continue
            thickness = min(w, h)
            if h > w:
                box = rect(x - thickness, y - 2, w + 2 * thickness, h + 3)
            else:
                box = rect(x - 3, y - thickness, w + 5, h + 2 * thickness)
            box = expand(box, 0, image.size)
            found.append({"kind": "busbar", "rect": box,
                          "evidence": "Long thick stroke outside component frames"})
    if block_diagram:
        found.extend(diagram_geometry.grounds(image, [c["rect"] for c in outlines]))
        found.extend(diagram_geometry.terminals(image, [c["rect"] for c in outlines]))
    else:
        found.extend(_circles(mask, contours))
    blue = ((rgb[:, :, 2] > 160) & (rgb[:, :, 0] < 100)).astype("uint8") * 255
    blue_contours, _ = cv2.findContours(blue, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    markers = []
    for contour in blue_contours:
        x, y, w, h = cv2.boundingRect(contour)
        if 3 <= w <= 8 and 3 <= h <= 8:
            patch = blue[y:y + h, x:x + w] > 0
            if patch.sum(0).max() >= h * 0.7 and patch.sum(1).max() >= w * 0.7:
                markers.append(rect(x, y, w, h))
    for resistor in ([] if block_diagram else resistors):
        if any(_contains(box, resistor) for box in outer):
            continue
        box = _load(mask, resistor, markers)
        if box:
            found.append({"kind": "load", "rect": expand(box, 0, image.size),
                          "evidence": "Resistor followed by three diminishing ground bars"})
    found.sort(key=lambda item: (item["kind"], item["rect"]["top"], item["rect"]["left"]))
    for index, item in enumerate(found, 1):
        item["candidate_id"] = index
    return found


def match(item, choices):
    """Choose nearby same-kind geometry; decline distant or equally plausible matches."""
    category = kind(item)
    seed = item.get("detected_rect", item["rect"])
    cx, cy = (seed["left"] + seed["right"]) / 2, (seed["top"] + seed["bottom"]) / 2
    ranked = []
    for candidate in choices:
        box = candidate["rect"]
        bx, by = (box["left"] + box["right"]) / 2, (box["top"] + box["bottom"]) / 2
        distance = math.hypot(cx - bx, cy - by)
        diagonal = math.hypot(box["width"], box["height"])
        overlap = iou(seed, box)
        same_kind = candidate["kind"] == category
        closed_kinds = {"frame", "tag", "ellipse", "terminal"}
        if not same_kind and not (category in closed_kinds
                                  and candidate["kind"] in closed_kinds and overlap >= 0.25):
            continue
        if category in {"frame", "tag", "ellipse", "terminal"}:
            ratio = box["width"] * box["height"] / max(1, seed["width"] * seed["height"])
            nearby = distance <= max(12, math.hypot(seed["width"], seed["height"]) * 0.65)
            if not 0.1 <= ratio <= 10 or (overlap < 0.12 and not nearby):
                continue
        radius = max(64 if category == "ground" else 24, diagonal * 1.25)
        if distance > radius:
            continue
        score = overlap + 0.5 * max(0, 1 - distance / radius) + (0.1 if same_kind else 0)
        ranked.append((score, candidate))
    ranked.sort(key=lambda pair: pair[0], reverse=True)
    if not ranked or (len(ranked) > 1 and ranked[0][0] - ranked[1][0] < 0.02):
        return None
    return ranked[0][1]


def correct(image, elements, refine_ids, padding, discover_missing=False):
    """Snap selected observations to symbols; merge only shared geometry candidates."""
    from . import diagram_geometry

    block_diagram = any(diagram_geometry.kind(item) in {"frame", "tag", "ellipse"}
                        for item in elements)
    choices = candidates(image, block_diagram=block_diagram)
    selected = set(refine_ids)
    assignments = {item["id"]: match(item, choices) for item in elements}
    canonical, kept, updates, merged_ids = {}, [], [], []
    for item in elements:
        if item["id"] not in selected:
            kept.append(item)
            continue
        candidate = assignments[item["id"]]
        update = {"id": item["id"], "method": "circuit-geometry", "before": dict(item["rect"]),
                  "applied": candidate is not None}
        if candidate is None:
            update["detail"] = ("Text control retained; circuit outlines do not validate "
                                "text bounds"
                                if diagram_geometry.is_text(item) else
                                "No unambiguous supported circuit-symbol geometry nearby")
            update["geometry_validation"] = {"status": "unresolved", "detail": update["detail"]}
            kept.append(item)
        else:
            key = candidate["candidate_id"]
            if key not in canonical:
                fitted = copy.deepcopy(item)
                fitted.update(corrected_rect=candidate["rect"], geometry_padding=padding,
                              geometry_validation={"status": "supported", "kind": candidate["kind"],
                                                   "evidence": candidate["evidence"]},
                              merged_source_element_ids=item.get("merged_source_element_ids",
                                                                 [item["id"]]),
                              geometry_correction={"method": "circuit-geometry",
                                                   "candidate_id": key,
                                                   "evidence": candidate["evidence"]})
                fitted.setdefault("detected_rect", dict(item["rect"]))
                canonical[key] = fitted
                kept.append(fitted)
            else:
                fitted = canonical[key]
                merged_ids.append(item["id"])
                fitted["merged_source_element_ids"] = sorted(set(
                    fitted["merged_source_element_ids"]
                    + item.get("merged_source_element_ids", [item["id"]])))
                fitted.setdefault("merged_observations", []).append({
                    "id": item["id"], "detected_rect": item.get("detected_rect", item["rect"]),
                    "semantic_name": item.get("semantic_name"),
                    "observed_value": item.get("observed_value"),
                    "source_tiles": item.get("source_tiles", [])})
                fitted["source_tiles"] = sorted(set(fitted.get("source_tiles", []))
                                                | set(item.get("source_tiles", [])))
                values = list(dict.fromkeys(value for value in [
                    *fitted.get("observed_value_variants", []), fitted.get("observed_value"),
                    item.get("observed_value")] if value))
                if len(values) > 1:
                    fitted["observed_value_variants"] = values
                    fitted["caption_ambiguity"] = "Same pixel geometry; model captions differ"
            update.update(canonical_id=canonical[key]["id"], candidate_id=key,
                          detail=candidate["evidence"])
        updates.append(update)
    added_ids = []
    if discover_missing:
        assigned = {candidate["candidate_id"] for candidate in assignments.values() if candidate}
        next_id = max((item["id"] for item in elements), default=0) + 1
        for candidate in choices:
            if candidate["kind"] != "busbar" or candidate["candidate_id"] in assigned:
                continue
            item = {"id": next_id, "semantic_name": "Busbar component", "entity_type": "busbar",
                    "intent": f"observe-busbar-{next_id}", "role": "read-only", "confidence": 0.0,
                    "ambiguity": "Local geometric observation; electrical identity not verified",
                    "source": "local-circuit-geometry", "suggested_actions": [],
                    "rect": candidate["rect"], "detected_rect": candidate["rect"],
                    "detected_rect_provenance": "Local pixels, not a model prediction",
                    "corrected_rect": candidate["rect"], "geometry_padding": padding,
                    "geometry_correction": {"method": "circuit-geometry-discovery",
                                            "candidate_id": candidate["candidate_id"],
                                            "evidence": candidate["evidence"]}}
            item["geometry_validation"] = {"status": "supported", "kind": "busbar",
                                           "evidence": candidate["evidence"]}
            kept.append(item)
            added_ids.append(next_id)
            updates.append({"id": next_id, "method": "circuit-geometry-discovery", "before": None,
                            "applied": True, "candidate_id": candidate["candidate_id"],
                            "detail": candidate["evidence"]})
            next_id += 1
    return kept, updates, {"geometry_source": "original-image-pixels",
                           "text_control_ids": [item["id"] for item in kept
                                                if diagram_geometry.is_text(item)],
                           "diagram_family": "block" if block_diagram else "electrical",
                           "candidate_counts": dict(Counter(item["kind"] for item in choices)),
                           "geometry_candidates": choices, "merged_element_ids": merged_ids,
                           "merged_count": len(merged_ids), "added_element_ids": added_ids,
                           "added_count": len(added_ids)}
