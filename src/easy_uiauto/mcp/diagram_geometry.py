"""Closed component outlines in connected block diagrams; no OCR or reference boxes."""

from __future__ import annotations

import re


def is_text(item):
    """Text is an independent control, even when it overlaps its graphic owner."""
    category = str(item.get("entity_type", "")).lower().replace("_", "-")
    if category in {"text", "text-control", "text-label", "label", "caption", "text-line",
                    "port-label", "value-text", "numeric-text"}:
        return True
    words = set(re.findall(r"[a-z]+", str(item.get("semantic_name", "")).lower()))
    return category in {"", "ui-element", "block"} and bool(words & {"label", "caption"})


def kind(item):
    """Geometry families, independent of model roles and incidental captions."""
    text = " ".join(str(item.get(k, "")).lower()
                    for k in ("entity_type", "semantic_name"))
    words = set(re.findall(r"[a-z]+", text))
    if is_text(item):
        return ""
    if "ground" in words:
        return "ground"
    if ("terminal" in words and words & {"electrical", "port"}) or (
            words & {"terminal", "port"} and "powerport" in text):
        return "terminal"
    if words & {"machine", "ammeter", "voltmeter", "meter"}:
        return "ellipse"
    if words & {"from", "goto"}:
        return "tag"
    if words & {"subsystem", "controller", "constant", "sum", "mux", "selector", "busselector",
                "memory", "inport", "terminator"} or "userdefineblock" in text:
        return "frame"
    return ""


def _box(x, y, w, h, padding, size):
    left, top = max(0, x - padding), max(0, y - padding)
    right, bottom = min(size[0], x + w + padding), min(size[1], y + h + padding)
    return dict(left=left, top=top, right=right, bottom=bottom,
                width=right - left, height=bottom - top)


def _inside(a, b):
    return all(a[k] <= b[k] for k in ("left", "top")) and all(
        a[k] >= b[k] for k in ("right", "bottom"))


def shapes(image):
    """Find frames, pointed tags and ellipses from enclosed original pixels.

    Wires attached to an outline do not change its enclosed contour. Internal
    divider lines can split a bus selector's hole: join those pieces only when
    the original vertical borders continue across the gap.
    """
    import cv2
    import numpy as np

    rgb = np.asarray(image.convert("RGB"))
    mask = ((rgb.max(2) < 150) & (rgb.max(2).astype("int16") - rgb.min(2) < 80))
    contours, hierarchy = cv2.findContours(mask.astype("uint8"), cv2.RETR_TREE,
                                          cv2.CHAIN_APPROX_SIMPLE)
    found = []
    for i, contour in enumerate(contours):
        if hierarchy[0][i][3] < 0:
            continue
        x, y, w, h = cv2.boundingRect(contour)
        if min(w, h) < 2 or max(w, h) < 10 or w * h > image.width * image.height * 0.6:
            continue
        area = cv2.contourArea(contour) / max(1, (w - 1) * (h - 1))
        polygon = cv2.approxPolyDP(contour, min(0.012 * cv2.arcLength(contour, True),
                                               min(w, h) * 0.15), True)
        points = polygon[:, 0, :]
        axis_aligned = len(points) == 4 and all(
            min(abs(int(p[0] - q[0])), abs(int(p[1] - q[1]))) <= 2
            for p, q in zip(points, [*points[1:], points[0]], strict=True))
        if (axis_aligned and area > 0.85) or area > 0.96:
            category = "frame"
        elif 5 <= len(points) <= 6 and min(w, h) >= 12 and 0.65 < area < 0.98:
            # A tag has two parallel horizontal edges and a pointed end.
            horizontal = sum(abs(int(p[1] - q[1])) <= 2 and abs(int(p[0] - q[0])) > w * 0.4
                             for p, q in zip(points, [*points[1:], points[0]], strict=True))
            if horizontal < 2:
                continue
            category = "tag"
        elif len(contour) >= 8 and min(w, h) >= 14 and 0.65 < area < 0.87:
            xy = contour[:, 0, :].astype(float)
            radial = ((xy[:, 0] - (x + (w - 1) / 2)) / ((w - 1) / 2)) ** 2 + (
                (xy[:, 1] - (y + (h - 1) / 2)) / ((h - 1) / 2)) ** 2
            if np.quantile(abs(radial - 1), 0.9) > 0.2:
                continue
            category = "ellipse"
        else:
            continue
        found.append({"kind": category, "rect": _box(x, y, w, h, 1, image.size),
                      "evidence": f"Closed {category} outline in original pixels"})

    # Retain the exterior frame, including a second border where present.
    frames = [c["rect"] for c in found if c["kind"] == "frame"]
    for a in list(frames):
        for b in list(frames):
            if a == b or abs(a["left"] - b["left"]) > 2 or abs(a["right"] - b["right"]) > 2:
                continue
            gap = b["top"] - a["bottom"]
            if not (a["top"] < b["top"] and -4 <= gap <= max(
                    20, min(a["width"], b["width"]) * 0.4)):
                continue
            sides = []
            for x in (a["left"], a["right"] - 1):
                band = mask[min(a["bottom"] - 1, b["top"]):max(a["bottom"], b["top"] + 1),
                            max(0, x - 1):x + 2]
                sides.append(band.any(axis=1).mean() if band.size else 0)
            if min(sides) > 0.85:
                box = _box(min(a["left"], b["left"]), a["top"],
                           max(a["right"], b["right"]) - min(a["left"], b["left"]),
                           b["bottom"] - a["top"], 0, image.size)
                found.append({"kind": "frame", "rect": box,
                              "evidence": "Continuous exterior frame around internal divider"})
    found = [c for c in found if not any(
        c != other and _inside(_box(other["rect"]["left"], other["rect"]["top"],
                                   other["rect"]["width"], other["rect"]["height"],
                                   2, image.size), c["rect"])
        and other["rect"]["width"] * other["rect"]["height"]
        > c["rect"]["width"] * c["rect"]["height"] * 1.001
        and other["rect"] != c["rect"] for other in found)]
    unique = {}
    for c in found:
        unique[(c["kind"], *[c["rect"][k] for k in ("left", "top", "right", "bottom")])] = c
    return list(unique.values())


def grounds(image, exclusions):
    """Three aligned diminishing bars outside component frames."""
    import numpy as np

    rgb = np.asarray(image.convert("RGB"))
    mask = ((rgb.max(2) < 150) & (rgb.max(2).astype("int16") - rgb.min(2) < 80))
    # Row runs retain even a one-pixel final bar, which morphological opening
    # would erase. Collapse identical adjacent runs to avoid counting thickness
    # as several separate bars.
    ys, xs = np.where(np.diff(np.pad(mask.astype("int8"), ((0, 0), (1, 1))), axis=1) != 0)
    active, bars = {}, []
    for y, start, end in zip(ys[::2], xs[::2], xs[1::2], strict=True):
        key = (int(start), int(end))
        if end - start > 36:
            continue
        previous = active.get(key)
        if previous is not None and previous[1] + previous[3] == y:
            previous[3] += 1
        else:
            previous = [int(start), int(y), int(end - start), 1]
            active[key] = previous
            bars.append(previous)
    bars = sorted([b for b in bars if b[3] <= 4], key=lambda r: r[1])
    found = []
    for a in bars:
        if a[2] < 7:
            continue
        center = round(a[0] + (a[2] - 1) / 2)
        stem = mask[max(0, a[1] - 6):a[1], center]
        if len(stem) < 3 or stem.mean() < 0.8:
            continue
        following = [b for b in bars if 0 < b[1] - a[1] <= max(12, a[2] * 1.2)
                     and abs(b[0] + b[2] / 2 - a[0] - a[2] / 2) <= 2]
        for b in following:
            for c in following:
                if not (a[2] >= b[2] + 2 and b[2] >= c[2] + 1
                        and 2 <= b[1] - a[1] <= max(6, a[2] * 0.4)
                        and 2 <= c[1] - b[1] <= max(6, a[2] * 0.4)):
                    continue
                box = _box(a[0], a[1] - 2, a[2], c[1] + c[3] - a[1] + 2, 0, image.size)
                if not any(_inside(e, box) for e in exclusions):
                    found.append({"kind": "ground", "rect": box,
                                  "evidence": "Three aligned diminishing ground bars"})
    return found


def terminals(image, exclusions=()):
    """Small closed rings, including colored electrical terminal symbols."""
    import cv2
    import numpy as np

    rgb = np.asarray(image.convert("RGB"))
    ink = rgb.min(2) < 160
    contours, hierarchy = cv2.findContours(ink.astype("uint8"), cv2.RETR_TREE,
                                          cv2.CHAIN_APPROX_SIMPLE)
    found = []
    angles = np.linspace(0, 2 * np.pi, 48, endpoint=False)
    for i, contour in enumerate(contours):
        x, y, w, h = cv2.boundingRect(contour)
        if hierarchy[0][i][3] < 0 or not (5 <= min(w, h) <= max(w, h) <= 32
                                                        and 0.8 <= w / h <= 1.25):
            continue
        if any(_inside(e, _box(x, y, w, h, 0, image.size)) for e in exclusions):
            continue
        if not 0.55 < cv2.contourArea(contour) / max(1, (w - 1) * (h - 1)) < 0.95:
            continue
        cx, cy = x + (w - 1) / 2, y + (h - 1) / 2
        dot = ink[max(0, round(cy) - 1):round(cy) + 2,
                  max(0, round(cx) - 1):round(cx) + 2]
        if dot.mean() < 0.5:
            continue
        radius = min(w, h) / 2
        supported = []
        for r in np.arange(max(2, radius - 1), radius + 5, 0.5):
            xx = np.clip(np.rint(cx + r * np.cos(angles)).astype(int), 0, image.width - 1)
            yy = np.clip(np.rint(cy + r * np.sin(angles)).astype(int), 0, image.height - 1)
            if ink[yy, xx].mean() > 0.7:
                supported.append(r)
        if not supported:
            continue
        r = max(supported) + 1
        box = _box(int(np.floor(cx - r)), int(np.floor(cy - r)),
                   int(np.ceil(2 * r + 1)), int(np.ceil(2 * r + 1)), 0, image.size)
        if min(box["width"], box["height"]) < 10:
            continue
        duplicate = next((c for c in found if abs(
            c["rect"]["left"] + c["rect"]["right"] - box["left"] - box["right"]) <= 4
            and abs(c["rect"]["top"] + c["rect"]["bottom"]
                    - box["top"] - box["bottom"]) <= 4), None)
        if duplicate:
            if (box["width"] * box["height"]
                    > duplicate["rect"]["width"] * duplicate["rect"]["height"]):
                duplicate["rect"] = box
            continue
        found.append({"kind": "terminal", "rect": box,
                      "evidence": "Small closed ring supported around its circumference"})
    return found
