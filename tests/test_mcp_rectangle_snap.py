"""Rectangle snapping preserves full control borders, raw evidence and text IDs."""

import json
from pathlib import Path

from PIL import Image, ImageDraw

from easy_uiauto.mcp import annotation_correction, image_analysis, rectangle_snap


def _sample():
    image = Image.new("RGB", (260, 140), "white")
    ImageDraw.Draw(image).rectangle((40, 40, 200, 70), outline=(180, 180, 180))
    elements = image_analysis.complete_element_bounds([
        {"id": 9, "annotation_label": "G2", "entity_type": "graphic-control", "source": "yolo",
         "rect": dict(left=50, top=45, right=190, bottom=66, width=140, height=21)},
        {"id": 41, "annotation_label": "T5", "entity_type": "text-label", "observed_value": "Value",
         "rect": dict(left=60, top=50, right=95, bottom=62, width=35, height=12)},
    ], *image.size, 1)
    return image, elements


def test_partial_prediction_snaps_to_the_complete_visible_border():
    image, elements = _sample()
    result, report = rectangle_snap.snap_rectangles(image, elements)
    graphic, text = result
    assert graphic["rect"] == dict(left=39, top=39, right=202, bottom=72, width=163, height=33)
    assert graphic["bbox"] == dict(x=39, y=39, width=163, height=33)
    assert graphic["center"] == {"x": 120.5, "y": 55.5}
    assert graphic["normalized_rect"]["right"] == 202 / 260
    assert graphic["detected_rect"] == elements[0]["detected_rect"]
    assert graphic["geometry_kind"] == "yolo-snapped"
    assert graphic["id"] == 9 and graphic["annotation_label"] == "G2"
    assert text == elements[1]
    assert report["applied_count"] == 1 and report["matched_count"] == 1
    assert elements[0]["rect"]["left"] == 49  # The input record is not mutated.


def test_text_and_unmatched_boxes_are_unchanged_even_with_different_padding():
    image = Image.new("RGB", (260, 140), "white")
    _other, elements = _sample()
    result, report = rectangle_snap.snap_rectangles(image, elements, box_padding=4)
    assert result == elements
    assert report["applied_count"] == report["matched_count"] == 0


def test_growth_limit_rejects_an_enclosing_panel():
    image = Image.new("RGB", (260, 140), "white")
    ImageDraw.Draw(image).rectangle((30, 40, 149, 59), outline="black")
    elements = image_analysis.complete_element_bounds([
        {"id": 1, "entity_type": "graphic-control", "rect": {
            "left": 60, "top": 40, "right": 120, "bottom": 60, "width": 60, "height": 20}}
    ], *image.size, 0)
    result, report = rectangle_snap.snap_rectangles(image, elements, box_padding=0)
    assert result == elements
    assert "1.6x" in report["updates"][0]["detail"]


def test_open_strokes_are_not_rectangle_candidates():
    image = Image.new("RGB", (260, 140), "white")
    ImageDraw.Draw(image).line([(40, 40), (40, 70), (200, 70)], fill="black")
    _other, elements = _sample()
    result, report = rectangle_snap.snap_rectangles(image, elements)
    assert result == elements and report["candidate_count"] == 0


def test_saved_result_can_be_corrected_without_reindexing_or_losing_metadata(tmp_path):
    image, elements = _sample()
    source = tmp_path / "input.png"
    image.save(source)
    path = tmp_path / "elements.json"
    path.write_text(json.dumps(dict(ok=True, image_path=str(source), image_width=260,
        image_height=140, box_padding=1, elements=elements, model="original",
        coordinate_system={"original_image_offset": {"x": 16, "y": 16}})), encoding="utf-8")
    result = annotation_correction.correct_annotations(str(path), profile="rectangles")
    assert result["model"] == "original"
    assert result["coordinate_system"]["original_image_offset"] == {"x": 16, "y": 16}
    assert [(e["id"], e["annotation_label"]) for e in result["elements"]] == [(9, "G2"), (41, "T5")]
    assert all(result["elements"][1][key] == value for key, value in elements[1].items())
    assert result["correction"]["applied_count"] == 1
    assert result["correction"]["api_requests"] == 0
    assert Path(result["annotated_image"]).is_file()
    assert json.loads(Path(result["coordinates_json"]).read_text(encoding="utf-8")) == result
