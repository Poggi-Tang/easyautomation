"""Color foreground and proposal evidence for saved diagram images.

Colors are selected by chroma, never by a named hue or a reference diagram.
Neutral wires do not join neighboring colored bodies. OCR remains a separate
observation layer and only prevents colored lettering becoming graphic boxes.
"""

from __future__ import annotations


def _rect(x, y, width, height):
    return dict(left=int(x), top=int(y), right=int(x + width), bottom=int(y + height),
                width=int(width), height=int(height))


def _overlap(a, b):
    return max(0, min(a["right"], b["right"]) - max(a["left"], b["left"])) * max(
        0, min(a["bottom"], b["bottom"]) - max(a["top"], b["top"]))


def _ports(rgb):
    """Small closed neutral squares survive while open connecting wires do not."""
    import cv2
    import numpy as np

    high, low = rgb.max(2), rgb.min(2)
    ink = ((high < 180) & (high.astype("int16") - low < 80)).astype("uint8")
    contours, hierarchy = cv2.findContours(ink, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    found = []
    for i, contour in enumerate(contours):
        if hierarchy[0][i][3] < 0:
            continue
        x, y, w, h = cv2.boundingRect(contour)
        if not (2 <= min(w, h) <= max(w, h) <= 9 and 0.6 <= w / h <= 1.6):
            continue
        if cv2.contourArea(contour) / max(1, (w - 1) * (h - 1)) < 0.7:
            continue
        left, top = max(0, x - 1), max(0, y - 1)
        right, bottom = min(rgb.shape[1], x + w + 1), min(rgb.shape[0], y + h + 1)
        found.append(_rect(left, top, right - left, bottom - top))
    # Antialiasing can open one corner of a tiny ring. Close that gap locally,
    # then remove thin wires; only small square masses can extend a color body.
    kernel = np.ones((3, 3), dtype="uint8")
    squares = cv2.morphologyEx(cv2.morphologyEx(ink, cv2.MORPH_CLOSE, kernel),
                               cv2.MORPH_OPEN, kernel)
    _count, _labels, stats, _centers = cv2.connectedComponentsWithStats(squares)
    for x, y, w, h, area in stats[1:]:
        if 3 <= min(w, h) <= max(w, h) <= 9 and 0.6 <= w / h <= 1.6 and area / (w * h) >= 0.8:
            found.append(_rect(x, y, w, h))
    return found


def solid_bars(image, exclusions=()):
    """Filled neutral separators/mux bars, without one-pixel connecting wires."""
    import cv2
    import numpy as np

    rgb = np.asarray(image.convert("RGB"))
    ink = ((rgb.max(2) < 150) & (rgb.max(2).astype("int16") - rgb.min(2) < 80))
    opened = cv2.morphologyEx(ink.astype("uint8"), cv2.MORPH_OPEN,
                              np.ones((2, 2), dtype="uint8"), anchor=(0, 0))
    contours, _ = cv2.findContours(opened, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    found = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if min(w, h) < 2 or max(w, h) < 16 or not 3 <= max(w, h) / min(w, h) <= 30:
            continue
        box = _rect(x, y, w, h)
        if any(_overlap(box, other) / (w * h) >= 0.8 for other in exclusions):
            continue
        if ink[y:y + h, x:x + w].mean() < 0.8:
            continue
        found.append(dict(kind="solid-bar", rect=box, provider="neutral-foreground",
                          evidence="Filled neutral bar, separated from thin connecting wires"))
    return found


def color_candidates(image, text_rects=()):
    """Connected multicolor bodies, including filled polygons and thin symbols.

    Only bridge narrow *dark* dividers inside a body. Closing white gaps would
    join a body to a nearby caption. Touching hues already share a component.
    Small closed ports adjacent to the body are included without following their
    wires. The unexpanded color rectangle is retained as evidence.
    """
    import cv2
    import numpy as np

    rgb = np.asarray(image.convert("RGB"))
    high, low = rgb.max(2), rgb.min(2)
    chroma = high.astype("int16") - low
    mask = ((chroma >= 45) & (chroma >= high * 0.22)).astype("uint8")
    closed = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((3, 3), dtype="uint8"))
    joined = mask | (closed & (high < 190).astype("uint8"))
    count, labels, stats, _centers = cv2.connectedComponentsWithStats(joined)
    ports, found = _ports(rgb), []
    for index in range(1, count):
        x, y, w, h, area = (int(v) for v in stats[index])
        if area < 48 or min(w, h) < 5 or max(w, h) < 12 or w * h > image.width * image.height * 0.6:
            continue
        seed = _rect(x, y, w, h)
        if any(_overlap(seed, t) / (w * h) >= 0.7 for t in text_rects):
            continue
        component = (labels[y:y + h, x:x + w] == index).astype("uint8")
        contours, _ = cv2.findContours(component, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contour = max(contours, key=cv2.contourArea)
        hull_area = cv2.contourArea(cv2.convexHull(contour))
        # Solid bodies and closed outlines qualify; sparse colored glyphs do not.
        silhouette = cv2.contourArea(contour)
        if hull_area <= 0 or silhouette / hull_area < 0.72 or silhouette / (w * h) < 0.3:
            continue
        attached = []
        reach = max(5, min(12, min(w, h) * 0.6))
        for port in ports:
            dx = max(0, x - port["right"], port["left"] - x - w)
            dy = max(0, y - port["bottom"], port["top"] - y - h)
            if max(dx, dy) > reach:
                continue
            cx, cy = (port["left"] + port["right"]) / 2, (port["top"] + port["bottom"]) / 2
            # Ports inside the body are artwork, not additions to its bounds.
            if x < cx < x + w and y < cy < y + h:
                continue
            if any(_overlap(port, t) >= port["width"] * port["height"] * 0.7 for t in text_rects):
                continue
            attached.append(port)
        left = min([x, *[p["left"] for p in attached]])
        top = min([y, *[p["top"] for p in attached]])
        right = max([x + w, *[p["right"] for p in attached]])
        bottom = max([y + h, *[p["bottom"] for p in attached]])
        found.append(dict(
            kind="color-body", rect=_rect(left, top, right - left, bottom - top),
            color_rect=seed, attached_port_rects=attached, provider="color-foreground",
            color_pixel_count=int(mask[y:y + h, x:x + w][component > 0].sum()),
            silhouette_fraction=round(silhouette / (w * h), 4),
            evidence="Connected color foreground with a compact exterior; adjacent closed ports"))
    return found


def proposal_evidence(image, box, text_rects=()):
    """Require visible non-text area before retaining an unmatched YOLO region.

    Closed/filled arbitrary contours are allowed; a wire, caption, or largely
    empty canvas is not sufficient. This gate does not assign a symbol family.
    """
    import cv2
    import numpy as np

    rgb = np.asarray(image.convert("RGB").crop(tuple(box[k] for k in (
        "left", "top", "right", "bottom"))))
    ink = (rgb.min(2) < 210).astype("uint8")
    for text in text_rects:
        left, top = max(0, text["left"] - box["left"]), max(0, text["top"] - box["top"])
        right = min(box["width"], text["right"] - box["left"])
        bottom = min(box["height"], text["bottom"] - box["top"])
        if right > left and bottom > top:
            ink[top:bottom, left:right] = 0
    contours, _ = cv2.findContours(ink, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    total = box["width"] * box["height"]
    best = max((cv2.contourArea(c) for c in contours), default=0)
    density = float(ink.mean()) if ink.size else 0
    accepted = best >= max(24, total * 0.12) and density >= 0.03
    return dict(accepted=bool(accepted), non_text_ink_fraction=round(density, 4),
                largest_contour_fraction=round(best / max(1, total), 4),
                reason="Non-text graphic pixels" if accepted else
                "No substantial closed or filled graphic pixels after excluding OCR text")
