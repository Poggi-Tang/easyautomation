"""Block diagrams with attached wires, shadows, internal dividers and small symbols."""

import hashlib
import json

import pytest
from PIL import Image, ImageDraw

from easy_uiauto.mcp import annotation_correction, circuit_geometry, image_analysis


def diagram(tmp_path, scale=1, offset=(0, 0)):
    im = Image.new("RGB", (700 + offset[0], 450 + offset[1]), "white")
    d = ImageDraw.Draw(im)
    ox, oy = offset

    def box(bounds, **kwargs):
        a, b, c, e = bounds
        d.rectangle((a + ox, b + oy, c + ox, e + oy), **kwargs)

    def line(points, **kwargs):
        d.line([(x + ox, y + oy) for x, y in points], **kwargs)

    box((62, 62, 222, 212), fill="#777777")
    box((60, 60, 220, 210), fill="#f1f6fa", outline="black")
    box((64, 64, 216, 206), outline="black")
    line([(0, 100), (60, 100)], fill="black")
    line([(220, 130), (300, 130)], fill="black")
    box((302, 82, 332, 202), fill="#777777")
    box((300, 80, 330, 200), fill="white", outline="black")
    line([(300, 140), (330, 140)], fill="black")
    line([(330, 100), (365, 100)], fill="black")
    d.polygon([(365 + ox, 85 + oy), (415 + ox, 85 + oy), (430 + ox, 100 + oy),
               (415 + ox, 115 + oy), (365 + ox, 115 + oy)], fill="white", outline="black")
    line([(430, 100), (500, 100)], fill="black")
    d.ellipse((500 + ox, 60 + oy, 555 + ox, 140 + oy), outline="black", width=2)
    line([(555, 100), (620, 100)], fill="black")
    line([(600, 60), (600, 145)], fill="black", width=4)
    line([(470, 290), (560, 290)], fill="black", width=3)
    line([(515, 260), (515, 320)], fill="black")
    line([(600, 290), (600, 340)], fill="black")
    for y, half in ((340, 6), (344, 3), (348, 0)):
        line([(600 - half, y), (600 + half, y)], fill="black")
    d.ellipse((636 + ox, 92 + oy, 652 + ox, 108 + oy), outline="#006699", width=2)
    d.ellipse((642 + ox, 98 + oy, 646 + ox, 102 + oy), fill="#0099ff")
    if scale != 1:
        im = im.resize((im.width * scale, im.height * scale), Image.Resampling.NEAREST)
    source = tmp_path / "diagram.png"
    im.save(source)
    seeds = [
        ("Controller subsystem", (70, 65, 208, 175)),
        ("Controller block", (65, 65, 215, 200)),
        ("BusSelector block", (302, 85, 328, 137)),
        ("From tag", (368, 88, 420, 112)),
        ("Synchronous machine", (505, 65, 547, 117)),
        ("Busbar", (597, 85, 604, 130)),
        ("Ground block", (592, 331, 610, 358)),
        ("PowerPort_3ph terminal", (635, 90, 655, 111)),
    ]
    elements = []
    for i, (name, bounds) in enumerate(seeds, 1):
        x1, y1, x2, y2 = bounds
        x1, x2 = (x1 + ox) * scale, (x2 + ox) * scale
        y1, y2 = (y1 + oy) * scale, (y2 + oy) * scale
        elements.append(dict(id=i, semantic_name=name, observed_value=f"value-{i}",
                             entity_type="ui-element", status="observed", uia_verified=False,
                             executable=False, rect=circuit_geometry.rect(
                                 x1, y1, x2 - x1, y2 - y1)))
    data = dict(ok=True, image_path=str(source), image_width=im.width, image_height=im.height,
                source_image_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
                elements=image_analysis.complete_element_bounds(elements, *im.size), box_padding=2)
    path = tmp_path / "original.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    return path, data


@pytest.mark.parametrize("scale,offset", [(1, (0, 0)), (1, (27, 19)), (2, (13, 17))])
def test_block_geometry_completes_outlines_without_discovering_shadows(tmp_path, scale, offset):
    p, old = diagram(tmp_path, scale, offset)
    result = annotation_correction.correct_annotations(
        str(p), refine_ids=list(range(1, 9)), profile="circuit", discover_missing=True)
    report = result["correction"]
    assert report["diagram_family"] == "block"
    assert report["candidate_counts"]["busbar"] == 2
    assert report["merged_count"] == 1 and report["added_count"] == 1
    assert report["skipped_count"] == 0
    by_id = {e["id"]: e for e in result["elements"]}
    assert by_id[1]["merged_source_element_ids"] == [1, 2]
    assert by_id[1]["detected_rect"] == old["elements"][0]["detected_rect"]
    assert by_id[1]["rect"]["bottom"] >= (210 + offset[1]) * scale
    assert by_id[3]["rect"]["bottom"] >= (200 + offset[1]) * scale
    assert by_id[4]["rect"]["right"] >= (430 + offset[0]) * scale
    assert by_id[5]["rect"]["bottom"] >= (140 + offset[1]) * scale
    assert not any(e["uia_verified"] or e["executable"] for e in by_id.values())


def test_small_wrong_box_outside_large_frame_is_not_absorbed(tmp_path):
    p, data = diagram(tmp_path)
    item = data["elements"][0]
    item["rect"] = item["detected_rect"] = circuit_geometry.rect(230, 145, 15, 10)
    data["elements"] = image_analysis.complete_element_bounds(
        data["elements"], data["image_width"], data["image_height"])
    p.write_text(json.dumps(data), encoding="utf-8")
    result = annotation_correction.correct_annotations(str(p), refine_ids=[1], profile="circuit")
    assert result["correction"]["skipped_count"] == 1
    assert result["elements"][0]["detected_rect"] == data["elements"][0]["detected_rect"]
    assert result["elements"][0]["rect"] == data["elements"][0]["rect"]
    assert result["correction"]["updates"][0]["geometry_validation"]["status"] == "unresolved"
    assert result["correction"]["needs_review_ids"] == [1]


def test_block_correction_is_idempotent(tmp_path):
    p, _ = diagram(tmp_path)
    first = annotation_correction.correct_annotations(
        str(p), refine_ids=list(range(1, 9)), profile="circuit", discover_missing=True)
    second = annotation_correction.correct_annotations(
        first["coordinates_json"], refine_ids=[e["id"] for e in first["elements"]],
        profile="circuit", discover_missing=True)
    assert second["correction"]["added_count"] == second["correction"]["merged_count"] == 0
    assert [(e["id"], e["rect"]) for e in first["elements"]] == [
        (e["id"], e["rect"]) for e in second["elements"]]


def test_text_control_is_retained_and_not_fitted_to_a_graphic_frame(tmp_path):
    p, data = diagram(tmp_path)
    data["elements"][0]["entity_type"] = "text-label"
    p.write_text(json.dumps(data), encoding="utf-8")
    result = annotation_correction.correct_annotations(str(p), refine_ids=[1], profile="circuit")
    assert result["elements"][0] == {**data["elements"][0], "annotation_label": "T1",
                                     "annotation_color": "#c65d00"}
    assert result["correction"]["text_control_ids"] == [1]
    assert "Text control retained" in result["correction"]["updates"][0]["detail"]
    assert circuit_geometry.kind({"entity_type": "text-label", "semantic_name": "Busbar label"}) \
        == ""
    assert circuit_geometry.kind({"entity_type": "text-label", "semantic_name": "PowerPort_3ph"}) \
        == ""
