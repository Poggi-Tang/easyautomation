"""Conservative page matching: window identity, layout, then foreground pixels."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

from . import knowledge


def layout_fingerprint(segmentation: dict, size: tuple[int, int]) -> str:
    """Use already detected entities; never perform another UIA tree traversal."""
    width, height = size
    entries = []
    for kind in ("regions", "controls"):
        for item in segmentation.get(kind, []):
            rect = item.get("relative_rect") or item.get("rect") or {}
            geometry = [round(float(rect.get(k, 0)) / scale, 2)
                        for k, scale in (("left", width), ("top", height),
                                         ("right", width), ("bottom", height))]
            entries.append([kind, item.get("role", ""), item.get("entity_type", ""),
                            item.get("intent") or item.get("semantic_name") or item.get("name", ""),
                            geometry])
    if not entries:
        return ""
    return hashlib.sha256(json.dumps(sorted(entries, key=str), ensure_ascii=False,
                                     sort_keys=True).encode()).hexdigest()


def visual_match(first, second) -> tuple[bool, float]:
    """Compare equal-size images using edges, so white backgrounds cannot dominate."""
    import cv2
    import numpy as np

    if first.size != second.size:
        return False, 0.0
    a, b = np.array(first.convert("RGB")), np.array(second.convert("RGB"))
    if np.array_equal(a, b):
        return True, 1.0
    gray_a, gray_b = (cv2.cvtColor(im, cv2.COLOR_RGB2GRAY) for im in (a, b))
    edges_a, edges_b = (cv2.Canny(im, 60, 120) > 0 for im in (gray_a, gray_b))
    union = edges_a | edges_b
    if not union.any():
        return False, 0.0
    score = float((edges_a & edges_b).sum() / union.sum())
    # Same outlines with different filled panels are also different states.
    changed = np.max(np.abs(a.astype('int16') - b), axis=2) > 16
    return score >= .97 and float(changed.mean()) <= .01, round(score, 4)


def select(directory: Path, screenshot, page: dict) -> dict:
    """Reuse only one corroborated page; an AI-generated ID is not identity proof."""
    from PIL import Image

    records = list(knowledge.iter_records(directory, "page"))
    identity = page.get("window_identity", {})
    names = {identity.get("title"), identity.get("uia_name")} - {None, ""}
    matches = []
    for _path, previous in records:
        old_identity = previous.get("window_identity")
        # Legacy UIA aliases such as dialog_center_widget can describe many native
        # dialogs. Without recorded native identity, only an exact caption is evidence.
        same_window = (old_identity == identity if old_identity else
                       bool(identity.get("title")) and previous.get("window_name") == identity["title"])
        if not names or not same_window:
            continue
        image_name = previous.get("image")
        if not image_name:
            continue
        image_path = (directory / image_name).resolve()
        if not image_path.is_relative_to(directory.resolve()) or not image_path.is_file():
            continue
        try:
            with Image.open(image_path) as previous_image:
                exact = (previous_image.size == screenshot.size
                         and previous_image.convert("RGB").tobytes() == screenshot.convert("RGB").tobytes())
                same_layout = (old_identity and page.get("layout_fingerprint")
                               and previous.get("layout_fingerprint") == page["layout_fingerprint"])
                valid, score = visual_match(previous_image, screenshot) if exact or same_layout else (False, 0.)
        except OSError:
            continue
        if valid:
            matches.append((previous, score))
    # If old duplicates exist, prefer the corroborated requested ID, otherwise refuse ambiguity.
    preferred = [match for match in matches if match[0]["id"] == page["id"]]
    if len(preferred) == 1:
        matches = preferred
    if len(matches) == 1:
        previous, score = matches[0]
        return {**page, "id": previous["id"], "name": previous.get("name") or page.get("name"),
                "previous_ai_id": page["id"], "identity_similarity": score,
                "identity_decision": "matched-window-and-page"}
    occupied = {knowledge.slugify(r["id"]) for _p, r in records}
    proposed = knowledge.slugify(page["id"], "page")
    if proposed not in occupied:
        return {**page, "id": proposed, "identity_decision": "new-page"}
    digest = hashlib.sha256(json.dumps(identity, sort_keys=True).encode()
                            + screenshot.convert("RGB").tobytes()).hexdigest()[:12]
    suffix, counter = f"-{digest}", 2
    identifier = proposed[:80-len(suffix)] + suffix
    while identifier in occupied:
        suffix = f"-{digest}-{counter}"
        identifier = proposed[:80-len(suffix)] + suffix
        counter += 1
    return {**page, "id": identifier, "previous_ai_id": page["id"],
            "identity_decision": "new-page-id-collision"}
