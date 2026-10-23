"""Correction behavior on clipped icons, attached badges and ambiguous geometry."""

import asyncio
import hashlib
import json
from pathlib import Path

import pytest
from PIL import Image, ImageDraw

from easy_uiauto.mcp import annotation_correction, image_analysis, scanner, server


def saved_annotations(tmp_path, image, boxes):
    source = tmp_path / "input.png"
    image.save(source)
    items = []
    for index, box in enumerate(boxes, 1):
        left, top, right, bottom = box
        items.append({"id": index, "semantic_name": f"Icon {index}", "intent": f"icon-{index}",
                      "rect": {"left": left, "top": top, "right": right, "bottom": bottom,
                               "width": right - left, "height": bottom - top},
                      "status": "observed", "uia_verified": False, "executable": False})
    data = {"ok": True, "image_path": str(source), "image_width": image.width,
            "image_height": image.height, "box_padding": 2, "model": "original-model",
            "source_image_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "elements": image_analysis.complete_element_bounds(items, *image.size),
            "element_count": len(items)}
    path = tmp_path / "input-elements.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    return path, data


def test_pixel_fit_encloses_clipped_artwork_without_absorbing_neighbor(monkeypatch, tmp_path):
    image = Image.new("RGB", (100, 60), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((10, 15, 25, 40), fill="blue")
    draw.rectangle((30, 15, 45, 40), fill="red")
    path, old = saved_annotations(tmp_path, image, [(14, 20, 23, 35), (30, 15, 46, 41)])
    before = path.read_bytes()

    def forbidden(*_args, **_kwargs):
        raise AssertionError("Correction must not call vision or inspect the desktop")

    monkeypatch.setattr(scanner.urlrequest, "urlopen", forbidden)
    monkeypatch.setattr(scanner, "_screenshot_window", forbidden)
    result = annotation_correction.correct_annotations(str(path), refine_ids=[1])
    first, second = result["elements"]
    assert first["rect"] == {"left": 9, "top": 14, "right": 27, "bottom": 42,
                              "width": 18, "height": 28}
    assert first["detected_rect"] == old["elements"][0]["detected_rect"]
    assert second == {**old["elements"][1], "annotation_label": "G2",
                      "annotation_color": "#1565c0"}
    assert path.read_bytes() == before
    assert result["correction"]["api_requests"] == 0
    assert result["correction"]["applied_count"] == 1
    repeated = annotation_correction.correct_annotations(result["coordinates_json"], refine_ids=[1])
    assert [e["rect"] for e in repeated["elements"]] == [e["rect"] for e in result["elements"]]
    assert repeated["coordinates_json"] != result["coordinates_json"]


def test_disconnected_parts_inside_one_known_icon_remain_one_id(tmp_path):
    image = Image.new("RGB", (100, 60), "white")
    draw = ImageDraw.Draw(image)
    draw.ellipse((10, 15, 14, 19), fill="blue")
    draw.rectangle((19, 23, 22, 28), fill="blue")
    draw.rectangle((32, 15, 45, 30), fill="red")
    path, _old = saved_annotations(tmp_path, image, [(9, 14, 18, 25), (32, 15, 46, 31)])
    result = annotation_correction.correct_annotations(str(path), refine_ids=[1], search_margin=10)
    rect = result["elements"][0]["rect"]
    assert rect["left"] < 10 and rect["right"] > 22 and rect["bottom"] > 28
    assert rect["right"] < 32
    assert result["element_count"] == 2


def test_faint_avatar_outline_is_retained_while_long_toolbar_frame_is_excluded(tmp_path):
    image = Image.new("RGB", (100, 70), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((5, 8, 90, 55), outline="#bdcfe5")
    draw.ellipse((20, 15, 42, 38), fill="#f3f3f3")
    draw.rectangle((25, 22, 35, 30), fill="blue")
    path, _old = saved_annotations(tmp_path, image, [(25, 22, 36, 31)])
    result = annotation_correction.correct_annotations(str(path), refine_ids=[1])
    assert result["elements"][0]["rect"] == {
        "left": 19, "top": 14, "right": 44, "bottom": 40, "width": 25, "height": 26}


def test_touching_badge_is_fitted_separately_and_parent_still_encloses_it(tmp_path):
    image = Image.new("RGB", (100, 70), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((20, 30, 39, 49), fill="#9966cc")
    draw.ellipse((35, 23, 43, 31), fill="#ff5533")
    path, old = saved_annotations(tmp_path, image, [(22, 32, 38, 48), (38, 24, 42, 28)])
    old["elements"][1]["relationships"] = [{"type": "attached-to", "target_intent": "icon-1"}]
    old["elements"] = image_analysis.complete_element_bounds(old["elements"], *image.size)
    path.write_text(json.dumps(old), encoding="utf-8")
    result = annotation_correction.correct_annotations(str(path), refine_ids=[1, 2])
    parent, badge = result["elements"]
    assert badge["parent_element_id"] == 1
    assert badge["rect"]["top"] <= 23 and badge["rect"]["bottom"] >= 32
    assert badge["rect"]["left"] >= 34 and badge["rect"]["width"] <= 12
    assert parent["rect"]["left"] <= badge["rect"]["left"]
    assert parent["rect"]["right"] >= badge["rect"]["right"]
    assert parent["rect"]["top"] <= badge["rect"]["top"]
    assert result["element_count"] == 2


def test_exact_rectangle_preserves_evidence_and_matching_artifact_coordinates(tmp_path):
    path, old = saved_annotations(tmp_path, Image.new("RGB", (100, 60), "white"),
                                  [(10, 20, 25, 35), (50, 20, 65, 35)])
    rect = {"left": 8, "top": 15, "right": 29, "bottom": 41}
    result = annotation_correction.correct_annotations(
        str(path), corrections=[{"id": 1, "rect": rect}])
    first = result["elements"][0]
    assert all(first["rect"][key] == value for key, value in rect.items())
    assert first["bbox"] == {"x": 8, "y": 15, "width": 21, "height": 26}
    assert first["center"] == {"x": 18.5, "y": 28}
    assert first["detected_rect"] == old["elements"][0]["detected_rect"]
    assert result["elements"][1] == {**old["elements"][1], "annotation_label": "G2",
                                     "annotation_color": "#1565c0"}
    assert json.loads(Path(result["coordinates_json"]).read_text(encoding="utf-8")) == result
    with Image.open(result["annotated_image"]) as annotated:
        assert annotated.getpixel((8, 30)) == (21, 101, 192)
        assert annotated.getpixel((9, 30)) == (255, 255, 255)


def test_connected_wire_at_search_edge_is_skipped_instead_of_inventing_bounds(tmp_path):
    image = Image.new("RGB", (100, 60), "white")
    ImageDraw.Draw(image).line((0, 30, 99, 30), fill="black", width=2)
    path, old = saved_annotations(tmp_path, image, [(40, 25, 60, 35)])
    result = annotation_correction.correct_annotations(str(path), refine_ids=[1])
    assert result["elements"] == [{**old["elements"][0], "annotation_label": "G1",
                                   "annotation_color": "#1565c0"}]
    assert result["correction"]["skipped_count"] == 1
    assert "crop boundary" in result["correction"]["updates"][0]["detail"]


def test_source_changed_at_same_dimensions_is_rejected_before_outputs(tmp_path):
    path, old = saved_annotations(tmp_path, Image.new("RGB", (100, 60), "white"),
                                  [(10, 20, 25, 35)])
    Image.new("RGB", (100, 60), "black").save(old["image_path"])
    with pytest.raises(ValueError, match="changed since analysis"):
        annotation_correction.correct_annotations(str(path), refine_ids=[1])
    assert not list(tmp_path.glob("image-correction-*"))


def test_analysis_hash_binds_pixels_actually_sent_even_if_source_changes(monkeypatch, tmp_path):
    source = tmp_path / "input.png"
    Image.new("RGB", (100, 60), "white").save(source)
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()

    def analyze(image, *_args):
        assert image.getpixel((0, 0)) == (255, 255, 255)
        Image.new("RGB", (100, 60), "black").save(source)
        return scanner._validate_segments({"controls": [
            {"left": 10, "top": 20, "width": 15, "height": 15,
             "semantic_name": "Icon", "confidence": 0.9}]}, *image.size)

    monkeypatch.setattr(scanner, "segment_interface", analyze)
    result = image_analysis.analyze_image(str(source), "url", "key", "model", "v",
                                          str(tmp_path), analysis_scale=1)
    assert result["source_image_sha256"] == original_hash
    with pytest.raises(ValueError, match="changed since analysis"):
        annotation_correction.correct_annotations(result["coordinates_json"], refine_ids=[1])


@pytest.mark.parametrize("kwargs", [
    {"refine_ids": [99]}, {"refine_ids": [1, 1]},
    {"corrections": [{"id": 1, "rect": {"left": -1, "top": 0, "right": 4, "bottom": 5}}]},
    {"refine_ids": [True]}, {"refine_ids": [1], "search_margin": 0},
])
def test_invalid_corrections_do_not_mutate_or_produce_artifacts(tmp_path, kwargs):
    path, _old = saved_annotations(tmp_path, Image.new("RGB", (100, 60), "white"),
                                   [(10, 20, 25, 35)])
    before = path.read_bytes()
    with pytest.raises(ValueError):
        annotation_correction.correct_annotations(str(path), **kwargs)
    assert path.read_bytes() == before
    assert not list(tmp_path.glob("image-correction-*"))


def test_mcp_correction_needs_no_vision_configuration_and_reports_errors(monkeypatch, tmp_path):
    path, _old = saved_annotations(tmp_path, Image.new("RGB", (100, 60), "white"),
                                   [(10, 20, 25, 35)])

    def forbidden(*_args):
        raise AssertionError("Local correction needs no API configuration")

    monkeypatch.setattr(server, "_vision_api_settings", forbidden)
    result = json.loads(asyncio.run(server.correct_ui_image_annotations(
        str(path), corrections=[{"id": 1, "rect": {
            "left": 9, "top": 19, "right": 26, "bottom": 36}}])))
    assert result["ok"] and not result["elements"][0]["executable"]
    assert Path(result["annotated_image"]).is_file()
    failure = json.loads(asyncio.run(server.correct_ui_image_annotations(
        str(path), refine_ids=[99])))
    assert not failure["ok"] and "ValueError" in failure["error"]
    tools = asyncio.run(server.mcp.list_tools())
    schema = next(tool.inputSchema for tool in tools if tool.name == "correct_ui_image_annotations")
    assert schema["required"] == ["coordinates_json"]
