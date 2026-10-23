"""Diagram foreground regressions use generated shapes, not private screenshots."""

import pytest
from PIL import Image, ImageDraw

from easy_uiauto import image_detection, image_geometry
from easy_uiauto.mcp import circuit_geometry, diagram_geometry


def _empty_ocr():
    return dict(boxes=[], texts=[], scores=[])


def _ocr(bounds, text="label"):
    x, y, right, bottom = bounds
    return dict(boxes=[[[x, y], [right, y], [right, bottom], [x, bottom]]],
                texts=[text], scores=[0.99])


@pytest.mark.parametrize("color", ["blue", "red", "lime", "yellow", "magenta", "cyan", "orange"])
def test_arbitrary_color_triangle_is_a_full_control_without_yolo(color):
    image = Image.new("RGB", (220, 170), "white")
    draw = ImageDraw.Draw(image)
    draw.polygon([(60, 40), (140, 40), (100, 110)], fill=color)
    draw.line((100, 5, 100, 40), fill="black")
    result = image_detection.combine_detections(image, dict(boxes=[], scores=[]), _empty_ocr())
    assert result["graphic_count"] == result["color_candidate_count"] == 1
    graphic = result["elements"][0]
    assert graphic["geometry_kind"] == "color-body"
    assert graphic["rect"]["left"] <= 60 and graphic["rect"]["right"] >= 141
    assert graphic["rect"]["top"] >= 35 and graphic["rect"]["bottom"] >= 111
    assert graphic["source"] == "local-color-geometry"
    assert not graphic["uia_verified"] and not graphic["executable"]


def test_touching_different_hues_merge_without_following_a_neutral_wire():
    image = Image.new("RGB", (240, 150), "white")
    draw = ImageDraw.Draw(image)
    draw.line((20, 60, 210, 60), fill="black")
    draw.rectangle((30, 40, 60, 80), fill="orange")
    draw.rectangle((30, 40, 60, 49), fill="red")
    draw.rectangle((160, 40, 190, 80), fill="cyan")
    choices = image_geometry.color_candidates(image)
    assert len(choices) == 2
    assert sorted(c["color_rect"]["width"] for c in choices) == [31, 31]
    assert all(c["rect"]["width"] < 40 for c in choices)


def test_color_body_includes_tiny_ports_but_not_a_nearby_caption():
    image = Image.new("RGB", (220, 160), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((70, 40, 85, 100), fill="blue")
    for x in (65, 77, 89):
        draw.line((x, 5, x, 135), fill="black")
        for y in (35, 101):
            draw.rectangle((x - 2, y, x + 2, y + 4), outline="black")
    draw.text((89, 72), "caption", fill="blue")
    choices = image_geometry.color_candidates(image, [dict(
        left=89, top=72, right=135, bottom=83, width=46, height=11)])
    assert len(choices) == 1
    box = choices[0]["rect"]
    assert box["left"] <= 63 and box["right"] >= 92
    assert 30 <= box["top"] <= 35 and 106 <= box["bottom"] < 115
    assert box["right"] < 100
    assert choices[0]["attached_port_rects"]


def test_colored_text_stays_text_and_does_not_expand_its_graphic():
    image = Image.new("RGB", (220, 150), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((30, 30, 100, 70), fill="cyan", outline="black")
    draw.text((104, 45), "LABEL", fill="blue")
    ocr = _ocr((104, 45, 140, 57), "LABEL")
    result = image_detection.combine_detections(image, dict(boxes=[], scores=[]), ocr)
    assert result["graphic_count"] == result["text_count"] == 1
    graphic, text = result["elements"]
    assert graphic["rect"]["right"] <= text["rect"]["left"]
    assert text["observed_value"] == "LABEL"


def test_color_and_neutral_frame_are_one_graphic_with_the_complete_border():
    image = Image.new("RGB", (240, 180), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((40, 40, 150, 120), fill="lime", outline="black", width=2)
    yolo = dict(boxes=[[42, 42, 149, 119]], scores=[0.8])
    result = image_detection.combine_detections(image, yolo, _empty_ocr())
    assert result["graphic_count"] == result["color_supported_graphics"] == 1
    graphic = result["elements"][0]
    assert graphic["rect"]["left"] <= 40 and graphic["rect"]["right"] >= 151
    assert graphic["color_evidence"] and graphic["yolo_proposal_ids"] == [1]


def test_unknown_outline_yolo_proposal_survives_but_empty_canvas_and_wires_do_not():
    image = Image.new("RGB", (240, 180), "white")
    draw = ImageDraw.Draw(image)
    draw.polygon([(30, 25), (70, 35), (85, 65), (40, 100)], outline="black", width=2)
    draw.line((140, 30, 210, 30), fill="black")
    yolo = dict(boxes=[[25, 20, 90, 105], [130, 20, 220, 40], [120, 110, 220, 160]],
                scores=[0.4, 0.4, 0.4])
    result = image_detection.combine_detections(image, yolo, _empty_ocr())
    assert result["graphic_count"] == result["unmatched_yolo_graphics"] == 1
    assert result["elements"][0]["geometry_kind"] == "yolo-region"
    assert result["elements"][0]["pixel_evidence"]["accepted"]
    assert {p["id"] for p in result["unassigned_yolo_proposals"]} == {2, 3}


def test_partial_yolo_prediction_does_not_create_a_second_graphic():
    image = Image.new("RGB", (240, 180), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((40, 40, 150, 120), fill="lime", outline="black")
    result = image_detection.combine_detections(image, dict(
        boxes=[[40, 110, 151, 132]], scores=[0.3]), _ocr((45, 125, 146, 136)))
    assert result["graphic_count"] == 1


def test_busbar_detail_is_not_ground_but_separate_ground_is_retained(monkeypatch):
    def box(x, y, w, h):
        return circuit_geometry.rect(x, y, w, h)

    candidates = [dict(kind="busbar", rect=box(30, 20, 15, 100), evidence="bar"),
                  dict(kind="ground", rect=box(31, 25, 10, 8), evidence="bar detail"),
                  dict(kind="ground", rect=box(90, 120, 14, 12), evidence="ground")]
    monkeypatch.setattr(circuit_geometry, "candidates", lambda *_a, **_kw: candidates)
    result = image_detection.combine_detections(Image.new("RGB", (200, 160), "white"),
                                                 dict(boxes=[], scores=[]), _empty_ocr())
    assert [g["geometry_kind"] for g in result["elements"]] == ["busbar", "ground"]
    assert candidates[1]["fusion_status"] == "rejected"


def test_tag_with_text_touching_its_border_keeps_the_complete_exterior():
    image = Image.new("RGB", (220, 150), "white")
    draw = ImageDraw.Draw(image)
    draw.polygon([(40, 35), (110, 35), (135, 60), (110, 85), (40, 85)], outline="black")
    # A glyph touches the left border and makes an inward notch in the white hole.
    draw.line([(40, 58), (51, 73), (55, 51), (60, 68)], fill="black", width=2)
    draw.line((135, 60, 190, 60), fill="black")
    choices = [c for c in diagram_geometry.shapes(image) if c["kind"] == "tag"]
    assert len(choices) == 1
    assert choices[0]["rect"]["left"] <= 40 and choices[0]["rect"]["right"] >= 136


def test_solid_bar_is_retained_without_treating_attached_wires_as_its_bounds():
    image = Image.new("RGB", (220, 150), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((90, 40, 96, 80), fill="black")
    draw.line((30, 60, 180, 60), fill="black")
    result = image_detection.combine_detections(image, dict(boxes=[], scores=[]), _empty_ocr())
    assert result["graphic_count"] == 1
    graphic = result["elements"][0]
    assert graphic["geometry_kind"] == "solid-bar"
    assert graphic["rect"]["width"] <= 12 and graphic["rect"]["height"] < 50
