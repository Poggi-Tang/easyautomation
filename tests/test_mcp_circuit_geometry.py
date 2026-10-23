"""Circuit correction from synthetic pixels, with wires and changed layouts."""

import asyncio
import hashlib
import json
from pathlib import Path

import pytest
from PIL import Image, ImageDraw

from easy_uiauto.mcp import annotation_correction, circuit_geometry, image_analysis, scanner, server


def circuit_image(scale=1, offset=(0, 0)):
    image = Image.new("RGB", (600 + offset[0], 400 + offset[1]), "white")
    draw = ImageDraw.Draw(image)
    ox, oy = offset

    def rectangle(box, **kwargs):
        x1, y1, x2, y2 = box
        draw.rectangle((x1 + ox, y1 + oy, x2 + ox, y2 + oy), **kwargs)

    def line(points, **kwargs):
        draw.line([(x + ox, y + oy) for x, y in points], **kwargs)

    rectangle((30, 30, 120, 80), outline="black", width=1)
    rectangle((35, 35, 115, 75), outline="black", width=1)
    line([(120, 55), (165, 55)], fill="black")
    rectangle((40, 170, 130, 215), outline="black", width=2)
    line([(0, 190), (40, 190)], fill="black")
    line([(130, 190), (165, 190)], fill="black")
    for cx in (205, 238):
        draw.ellipse((cx - 24 + ox, 150 - 24 + oy, cx + 24 + ox, 150 + 24 + oy),
                     outline="black", width=2)
    line([(150, 150), (181, 150)], fill="black")
    line([(262, 150), (295, 150)], fill="black")
    line([(330, 65), (330, 140)], fill="black", width=4)
    line([(280, 100), (380, 100)], fill="black")
    line([(300, 270), (375, 270)], fill="black", width=3)
    line([(335, 245), (335, 295)], fill="black")
    # Vertical RLC branch. The ordinary wire continues far beyond its resistor.
    rectangle((430, 160, 440, 186), outline="black", width=2)
    line([(435, 130), (435, 160)], fill="black")
    line([(435, 186), (438, 190), (432, 194), (438, 198), (435, 203), (435, 230)],
         fill="black")
    for y in (212, 218):
        line([(428, y), (442, y)], fill="black")
    for y, half in ((230, 7), (234, 5), (238, 3)):
        line([(435 - half, y), (435 + half, y)], fill="black")
    # A separate horizontal grounded branch.
    rectangle((75, 294, 101, 306), outline="black", width=2)
    line([(30, 300), (75, 300)], fill="black")
    line([(101, 300), (107, 303), (112, 297), (118, 303), (124, 300), (167, 300)],
         fill="black")
    for x in (147, 153):
        line([(x, 293), (x, 307)], fill="black")
    for x, half in ((167, 7), (171, 5), (175, 3)):
        line([(x, 300 - half), (x, 300 + half)], fill="black")
    draw.text((300 + ox, 305 + oy), "BUS LABEL", fill="#0099ff")
    if scale != 1:
        image = image.resize((image.width * scale, image.height * scale), Image.Resampling.NEAREST)
    return image


def saved_circuit(tmp_path, scale=1, offset=(0, 0)):
    image = circuit_image(scale, offset)
    image_path = tmp_path / "diagram.png"
    image.save(image_path)
    seeds = [
        ("PowerPort_3ph block", (34, 32, 116, 76)),
        ("Boxed pi/RLC line", (35, 166, 140, 220)),
        ("Pi line component", (42, 174, 126, 213)),
        ("Double-circle transformer", (217, 132, 263, 173)),
        ("Serial load branch", (407, 170, 447, 205)),
        ("Load branch", (80, 290, 150, 310)),
        ("Busbar", (325, 80, 337, 125)),
    ]
    elements = []
    for index, (name, bounds) in enumerate(seeds, 1):
        x1, y1, x2, y2 = bounds
        x1, x2 = (x1 + offset[0]) * scale, (x2 + offset[0]) * scale
        y1, y2 = (y1 + offset[1]) * scale, (y2 + offset[1]) * scale
        elements.append({"id": index, "semantic_name": name, "observed_value": f"caption-{index}",
                         "rect": circuit_geometry.rect(x1, y1, x2 - x1, y2 - y1),
                         "status": "observed", "uia_verified": False, "executable": False})
    data = {"ok": True, "image_path": str(image_path), "image_width": image.width,
            "image_height": image.height, "box_padding": 2, "element_count": len(elements),
            "source_image_sha256": hashlib.sha256(image_path.read_bytes()).hexdigest(),
            "elements": image_analysis.complete_element_bounds(elements, *image.size)}
    path = tmp_path / "model-output.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    return path, data


@pytest.mark.parametrize(("scale", "offset"), [(1, (0, 0)), (1, (43, 29)), (2, (11, 17))])
def test_structural_correction_works_after_translation_and_scaling(tmp_path, scale, offset):
    path, old = saved_circuit(tmp_path, scale, offset)
    result = annotation_correction.correct_annotations(
        str(path), refine_ids=list(range(1, 8)), profile="circuit", discover_missing=True)
    report = result["correction"]
    assert report["candidate_counts"] == {
        "block": 1, "busbar": 2, "line": 1, "load": 2, "transformer": 1}
    assert report["applied_count"] == 7 and report["skipped_count"] == 0
    assert report["merged_count"] == 1 and report["added_count"] == 1
    assert result["element_count"] == 7
    line = next(e for e in result["elements"] if e["id"] == 2)
    assert line["merged_source_element_ids"] == [2, 3]
    assert line["detected_rect"] == old["elements"][1]["detected_rect"]
    assert line["observed_value_variants"] == ["caption-2", "caption-3"]
    load = next(e for e in result["elements"] if e["id"] == 5)
    assert load["rect"]["bottom"] >= (239 + offset[1]) * scale
    assert load["rect"]["height"] > load["rect"]["width"] * 2
    transformer = next(e for e in result["elements"] if e["id"] == 4)
    assert transformer["rect"]["left"] <= (181 + offset[0]) * scale
    assert transformer["rect"]["right"] >= (262 + offset[0]) * scale
    assert all(not e["executable"] and not e["uia_verified"] for e in result["elements"])
    added = next(e for e in result["elements"] if e["id"] in report["added_element_ids"])
    assert added["source"] == "local-circuit-geometry"
    assert report["api_requests"] == 0


def test_circuit_repeat_does_not_duplicate_recovered_busbars(tmp_path):
    path, _old = saved_circuit(tmp_path)
    first = annotation_correction.correct_annotations(
        str(path), refine_ids=list(range(1, 8)), profile="circuit", discover_missing=True)
    second = annotation_correction.correct_annotations(
        first["coordinates_json"], refine_ids=[e["id"] for e in first["elements"]],
        profile="circuit", discover_missing=True)
    assert second["correction"]["added_count"] == second["correction"]["merged_count"] == 0
    assert [(e["id"], e["rect"]) for e in first["elements"]] == [
        (e["id"], e["rect"]) for e in second["elements"]]


def test_component_kind_does_not_confuse_a_line_caption_with_its_busbars():
    assert circuit_geometry.kind({"semantic_name": "线路-母线A-母线B"}) == "line"
    assert circuit_geometry.kind({"semantic_name": "Line connection", "entity_type": "busbar"}) \
        == "busbar"


def test_partial_selection_keeps_other_annotations_and_discovery_is_opt_in(tmp_path):
    path, old = saved_circuit(tmp_path)
    result = annotation_correction.correct_annotations(str(path), refine_ids=[4], profile="circuit")
    assert result["element_count"] == 7 and result["correction"]["added_count"] == 0
    assert result["elements"][0] == {**old["elements"][0], "annotation_label": "G1",
                                     "annotation_color": "#1565c0"}
    assert {e["id"] for e in result["elements"]} == set(range(1, 8))


def test_plain_wire_has_no_symbol_candidate_and_no_fabricated_correction(tmp_path):
    path, old = saved_circuit(tmp_path)
    image = Image.new("RGB", (600, 400), "white")
    ImageDraw.Draw(image).line((0, 55, 599, 55), fill="black", width=1)
    image.save(old["image_path"])
    old["source_image_sha256"] = hashlib.sha256(Path(old["image_path"]).read_bytes()).hexdigest()
    path.write_text(json.dumps(old), encoding="utf-8")
    result = annotation_correction.correct_annotations(str(path), refine_ids=[1], profile="circuit")
    assert result["correction"]["skipped_count"] == 1
    assert result["elements"] == [{**e, "annotation_label": f"G{i}",
                                   "annotation_color": "#1565c0"}
                                  for i, e in enumerate(old["elements"], 1)]


def test_mcp_circuit_profile_has_no_api_or_reference_dependency(monkeypatch, tmp_path):
    path, _old = saved_circuit(tmp_path)

    def forbidden(*_args, **_kwargs):
        raise AssertionError("Circuit correction must use original image pixels locally")

    monkeypatch.setattr(scanner.urlrequest, "urlopen", forbidden)
    monkeypatch.setattr(server, "_vision_api_settings", forbidden)
    result = json.loads(asyncio.run(server.correct_ui_image_annotations(
        str(path), refine_ids=list(range(1, 8)), profile="circuit", discover_missing=True)))
    assert result["ok"] and result["element_count"] == 7
    schema = next(t.inputSchema for t in asyncio.run(server.mcp.list_tools())
                  if t.name == "correct_ui_image_annotations")
    assert "profile" in schema["properties"] and "discover_missing" in schema["properties"]
