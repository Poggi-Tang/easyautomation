"""Analyze saved UI images without inspecting or controlling the desktop."""

from __future__ import annotations

import hashlib
import io
import json
import math
import time
from collections.abc import Callable
from pathlib import Path
from tempfile import gettempdir
from uuid import uuid4

from . import scanner, visualization


def complete_element_bounds(elements: list[dict], width: int, height: int,
                            padding: int = 2) -> list[dict]:
    """Include known attached parts and a small margin without losing raw predictions."""
    if not 0 <= padding <= 20:
        raise ValueError("box_padding must be between 0 and 20 original-image pixels")
    completed = [{**item, "detected_rect": dict(item.get("detected_rect", item["rect"]))}
                 for item in elements]
    by_id = {item["id"]: item for item in completed}
    parents = {}
    for item in completed:
        parent = item.get("parent_element_id")
        if parent is None:
            intents = [rel.get("target_intent") for rel in item.get("relationships", [])
                       if rel.get("type") == "attached-to" and rel.get("target_intent")]
            candidates = [other["id"] for other in completed
                          if other["id"] != item["id"] and other.get("intent") in intents]
            parent = candidates[0] if len(candidates) == 1 else None
        if parent in by_id and parent != item["id"]:
            parents[item["id"]] = parent
    # Ignore cyclic attachment claims rather than inventing a component hierarchy.
    for child in list(parents):
        visited, current = set(), child
        while current in parents:
            if current in visited:
                parents.pop(child, None)
                break
            visited.add(current)
            current = parents[current]
    envelopes = {item["id"]: dict(item.get("corrected_rect", item["detected_rect"]))
                 for item in completed}
    for item in completed:
        current, rect = item["id"], item.get("corrected_rect", item["detected_rect"])
        while current in parents:
            current = parents[current]
            envelope = envelopes[current]
            for side in ("left", "top"):
                envelope[side] = min(envelope[side], rect[side])
            for side in ("right", "bottom"):
                envelope[side] = max(envelope[side], rect[side])
    for item in completed:
        rect = envelopes[item["id"]]
        item_padding = item.get("geometry_padding", padding)
        if not 0 <= item_padding <= 20:
            raise ValueError("geometry_padding must be between 0 and 20")
        left, top = max(0, rect["left"] - item_padding), max(0, rect["top"] - item_padding)
        right, bottom = (min(width, rect["right"] + item_padding),
                         min(height, rect["bottom"] + item_padding))
        item.update(
            rect={"left": left, "top": top, "right": right, "bottom": bottom,
                  "width": right - left, "height": bottom - top},
            bbox={"x": left, "y": top, "width": right - left, "height": bottom - top},
            center={"x": (left + right) / 2, "y": (top + bottom) / 2},
            normalized_rect={"left": left / width, "top": top / height,
                             "right": right / width, "bottom": bottom / height},
            box_padding=item_padding,
        )
        if item["id"] in parents:
            item["parent_element_id"] = parents[item["id"]]
        children = [child for child, parent in parents.items() if parent == item["id"]]
        if children:
            item["component_element_ids"] = children
    return completed


def save_image_annotations(image, elements: list[dict], directory: Path, *,
                           filename: str = "annotated.png") -> dict:
    """Render typed thin outlines and adjacent IDs; preserve numeric element IDs."""
    from PIL import Image

    from . import diagram_geometry

    if Path(filename).name != filename or Path(filename).suffix.lower() != ".png":
        raise ValueError("Annotation filename must be a PNG basename")
    kinds = ["text" if diagram_geometry.is_text(item) else "graphic" for item in elements]
    counters = dict.fromkeys(visualization.IMAGE_ANNOTATION_STYLES, 0)
    used = set()
    for item, kind in zip(elements, kinds, strict=True):
        prefix = visualization.IMAGE_ANNOTATION_STYLES[kind]["prefix"]
        label = str(item.get("annotation_label", ""))
        if (label.startswith(prefix) and label[1:].isdigit() and int(label[1:]) > 0
                and label not in used):
            counters[kind] = max(counters[kind], int(label[1:]))
            used.add(label)
        else:
            item.pop("annotation_label", None)
    for item, kind in zip(elements, kinds, strict=True):
        style = visualization.IMAGE_ANNOTATION_STYLES[kind]
        if "annotation_label" not in item:
            counters[kind] += 1
            item["annotation_label"] = f"{style['prefix']}{counters[kind]}"
        item["annotation_color"] = style["color"]

    def render(canvas, scale_x=1, scale_y=1):
        markers = [{"index": item["id"], "color": item["annotation_color"],
                    "annotation_label": item["annotation_label"],
                    "text_control": diagram_geometry.is_text(item),
                    "label_parent_index": item.get("related_element_id"),
                    "rect": {key: round(value * (scale_x if key in {"left", "right", "width"}
                                                  else scale_y))
                             for key, value in item["rect"].items()}}
                   for item in elements]
        return visualization.annotate_screenshot(
            canvas, markers, avoid_label_overlap=True, outline_width=1, compact_labels=True)

    annotated_path = directory / filename
    annotated = render(image)
    annotated.save(annotated_path)
    layout = annotated.info["annotation_layout"]
    scale = max(1, min(4, 960 / max(image.size)))
    preview_path = annotated_path
    preview_layout = layout
    if scale > 1:
        size = (round(image.width * scale), round(image.height * scale))
        preview_path = directory / f"{Path(filename).stem}.preview.png"
        preview = render(image.resize(size, Image.Resampling.LANCZOS),
                         size[0] / image.width, size[1] / image.height)
        preview.save(preview_path)
        preview_layout = preview.info["annotation_layout"]
    return {"annotated_image": str(annotated_path),
            "annotated_preview_image": str(preview_path),
            "annotation_legend": visualization.IMAGE_ANNOTATION_STYLES,
            "annotation_layout": layout,
            "annotation_preview_layout": preview_layout}


def analyze_image(
    image_path: str,
    api_url: str,
    api_key: str,
    model: str,
    version: str,
    output_dir: str = "",
    max_elements: int = 1000,
    progress: Callable[[str], None] | None = None,
    analysis_scale: float = 0,
    box_padding: int = 2,
    analysis_context: str = "",
) -> dict:
    """Return numbered observations and coordinates in the displayed input image."""
    from PIL import Image, ImageOps

    started = time.perf_counter()
    if not 1 <= max_elements <= 3000:
        raise ValueError("max_elements must be between 1 and 3000")
    if not math.isfinite(analysis_scale) or (analysis_scale != 0 and not 1 <= analysis_scale <= 4):
        raise ValueError("analysis_scale must be 0 (automatic) or between 1 and 4")
    if not 0 <= box_padding <= 20:
        raise ValueError("box_padding must be between 0 and 20 original-image pixels")
    source = Path(image_path).expanduser().resolve(strict=True)
    if not source.is_file():
        raise ValueError("image_path must identify an image file")
    source_bytes = source.read_bytes()
    with Image.open(io.BytesIO(source_bytes)) as original:
        if getattr(original, "n_frames", 1) != 1:
            raise ValueError("Supply a single-frame image, not an animation or multipage file")
        orientation = original.getexif().get(274, 1)
        oriented = ImageOps.exif_transpose(original).convert("RGBA")
        # Transparent screenshots should look the same against a white viewer background.
        background = Image.new("RGBA", oriented.size, "white")
        image = Image.alpha_composite(background, oriented).convert("RGB")
    width, height = image.size
    scale = analysis_scale or max(1, min(4, 960 / max(width, height)))
    analysis_size = (round(width * scale), round(height * scale))
    analysis_image = image.resize(analysis_size, Image.Resampling.LANCZOS)
    scale_x, scale_y = analysis_size[0] / width, analysis_size[1] / height
    segmentation = scanner.segment_interface(
        analysis_image, api_url, api_key, model, version, progress,
        *((analysis_context,) if analysis_context else ()))

    def original_rect(rect: dict) -> dict:
        left, top = max(0, math.floor(rect["left"] / scale_x)), max(
            0, math.floor(rect["top"] / scale_y))
        right, bottom = min(width, math.ceil(rect["right"] / scale_x)), min(
            height, math.ceil(rect["bottom"] / scale_y))
        return {"left": left, "top": top, "right": right, "bottom": bottom,
                "width": right - left, "height": bottom - top}

    segmentation["regions"] = [{**region, "rect": original_rect(region["rect"])}
                               for region in segmentation["regions"]]
    targets = segmentation["controls"]
    elements = []
    for index, target in enumerate(targets[:max_elements], start=1):
        rect = original_rect(target["relative_rect"])
        elements.append({
            **{key: value for key, value in target.items()
               if key not in {"id", "actions", "relative_rect"} and not key.startswith("_")},
            "id": index,
            "status": "observed",
            "uia_verified": False,
            "executable": False,
            "suggested_actions": target.get("actions", []),
            "rect": rect,
            "bbox": {"x": rect["left"], "y": rect["top"],
                     "width": rect["width"], "height": rect["height"]},
            "center": {"x": (rect["left"] + rect["right"]) / 2,
                       "y": (rect["top"] + rect["bottom"]) / 2},
            "normalized_rect": {"left": rect["left"] / width, "top": rect["top"] / height,
                                "right": rect["right"] / width,
                                "bottom": rect["bottom"] / height},
        })
    elements = complete_element_bounds(elements, width, height, box_padding)

    # A separate run folder prevents overwriting either the input or previous results.
    base = Path(output_dir).expanduser() if output_dir else Path(gettempdir()) / "easy_uiauto"
    destination = (base / f"image-analysis-{uuid4().hex[:12]}").resolve()
    destination.mkdir(parents=True, exist_ok=False)
    coordinates_path = destination / "elements.json"
    annotation_paths = save_image_annotations(image, elements, destination)
    result = {
        "ok": True,
        "image_path": str(source),
        "source_image_sha256": hashlib.sha256(source_bytes).hexdigest(),
        "image_width": width,
        "image_height": height,
        "analysis_image_width": analysis_size[0],
        "analysis_image_height": analysis_size[1],
        "analysis_scale": {"x": scale_x, "y": scale_y},
        "exif_orientation_applied": orientation not in (None, 1),
        "coordinate_system": {
            "origin": "top-left", "units": "pixels", "reference": "display-oriented-input",
            "bounds": "left/top inclusive, right/bottom exclusive",
            "screen_coordinates": False,
        },
        "page": segmentation["page"],
        "regions": segmentation["regions"],
        "elements": elements,
        "element_count": len(elements),
        "detected_element_count": len(targets),
        "truncated": len(targets) > max_elements,
        "coverage": segmentation.get("coverage", {}),
        **annotation_paths,
        "coordinates_json": str(coordinates_path),
        "uia_verified": False,
        "commands_generated": 0,
        "model": model,
        "analysis_context": analysis_context,
        "box_padding": box_padding,
        "elapsed_seconds": round(time.perf_counter() - started, 2),
    }
    coordinates_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result
