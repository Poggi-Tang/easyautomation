"""Local image CLI resource exports, failure behavior and engine boundaries."""

import csv
import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image, ImageDraw

from easy_uiauto import image_cli, image_detection


@pytest.fixture
def local_engines(monkeypatch):
    yolo = dict(model_id=image_detection.MODEL_ID, model_sha256=image_detection.MODEL_SHA256,
                boxes=[[29, 39, 152, 122]], scores=[0.8], elapsed_seconds=0.01)
    ocr = dict(engine="test OCR", boxes=[[[60, 75], [100, 75], [100, 90], [60, 90]]],
               texts=["中文 A"], scores=[0.95], elapsed_seconds=0.01)
    monkeypatch.setattr(image_detection, "check_runtime", lambda *_a: None)
    monkeypatch.setattr(image_detection, "ensure_weights", lambda path, **_kw: path)
    monkeypatch.setattr(image_detection, "run_yolo", lambda *_a, **_kw: yolo)
    monkeypatch.setattr(image_detection, "run_ocr", lambda *_a, **_kw: ocr)
    return yolo, ocr


def test_cli_exports_all_three_layers_and_matching_resources(
        local_engines, monkeypatch, tmp_path, capsys):
    from easy_uiauto.mcp import scanner

    def forbidden(*_args, **_kwargs):
        raise AssertionError("The image CLI must not use the desktop or a vision API")

    monkeypatch.setattr(scanner, "_screenshot_window", forbidden)
    monkeypatch.setattr(scanner.urlrequest, "urlopen", forbidden)
    source = tmp_path / "图 & 输入.png"
    image = Image.new("RGB", (240, 160), "white")
    ImageDraw.Draw(image).rectangle((30, 40, 150, 120), outline="black", width=1)
    image.save(source)
    original = source.read_bytes()
    output = tmp_path / "输出 资源"
    assert image_cli.main([str(source), str(output), "--offline", "--json"]) == 0
    result = json.loads(capsys.readouterr().out)
    folder = Path(result["output_directory"])
    assert result["graphic_count"] == result["text_count"] == 1
    assert result["element_count"] == 2 and result["api_requests"] == 0
    assert source.read_bytes() == original
    assert folder.parent == output
    assert set(result["resources"]) <= {p.name for p in folder.iterdir()}
    assert {"controls.png", "text.png", "combined.png", "original.png", "elements.json",
            "elements.csv", "raw-yolo.json", "raw-ocr.json", "ocr-recovery.json",
            "geometry-candidates.json",
            "index.html", "summary.json", "run.log"} <= set(result["resources"])
    assert json.loads((folder / "elements.json").read_text(encoding="utf-8")) == result
    with (folder / "elements.csv").open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    assert [r["annotation_label"] for r in rows] == ["G1", "T1"]
    assert rows[1]["observed_value"] == "中文 A"
    assert result["layers"]["text"]["annotation_layout"]["labels"][0]["annotation_label"] == "T1"
    assert "图 &amp; 输入.png" in (folder / "index.html").read_text(encoding="utf-8")
    for name in ("controls.png", "text.png", "combined.png"):
        with Image.open(folder / name) as annotated:
            assert annotated.size == image.size
    again = image_cli.analyze_local_image(str(source), str(output), progress=lambda _s: None)
    assert again["output_directory"] != result["output_directory"]
    assert (folder / "elements.json").is_file()


def test_missing_input_returns_nonzero_without_creating_output(tmp_path, capsys):
    output = tmp_path / "results"
    assert image_cli.main([str(tmp_path / "missing.png"), str(output)]) == 2
    assert "FileNotFoundError" in capsys.readouterr().err
    assert not output.exists()


def test_runtime_failure_leaves_an_explicit_failure_report(monkeypatch, tmp_path):
    source = tmp_path / "image.png"
    Image.new("RGB", (100, 60), "white").save(source)

    def missing(*_a):
        raise RuntimeError("missing OCR package")

    monkeypatch.setattr(image_detection, "check_runtime", missing)
    output = tmp_path / "out"
    with pytest.raises(RuntimeError, match="missing OCR"):
        image_cli.analyze_local_image(source, output, progress=lambda _s: None)
    report = json.loads(next(output.glob("*/summary.json")).read_text(encoding="utf-8"))
    assert report["ok"] is False and "missing OCR" in report["error"]
    assert not list(output.glob("*/combined.png"))


@pytest.mark.parametrize("confidence", [0, 1, float("nan"), float("inf")])
def test_invalid_threshold_fails_before_loading_image_or_models(tmp_path, confidence):
    output = tmp_path / "out"
    with pytest.raises(ValueError, match="confidence"):
        image_cli.analyze_local_image(tmp_path / "missing.png", output, confidence=confidence)
    assert not output.exists()


def test_model_checksum_rejects_corrupt_cache_and_offline_never_downloads(monkeypatch, tmp_path):
    def forbidden(*_a, **_kw):
        raise AssertionError("Must not download in offline mode or overwrite a corrupt model")

    monkeypatch.setattr(image_detection.request, "urlopen", forbidden)
    model = tmp_path / "model.pt"
    with pytest.raises(FileNotFoundError, match="not cached"):
        image_detection.ensure_weights(model, offline=True)
    model.write_bytes(b"corrupt model")
    with pytest.raises(ValueError, match="SHA256"):
        image_detection.ensure_weights(model)
    assert model.read_bytes() == b"corrupt model"


def test_ui_profile_keeps_graphic_owner_but_deduplicates_yolo_text():
    image = Image.new("RGB", (240, 160), "white")
    yolo = dict(boxes=[[10, 10, 200, 140], [20, 50, 80, 65]], scores=[0.9, 0.8])
    ocr = dict(boxes=[[[20, 50], [80, 50], [80, 65], [20, 65]]],
               texts=["literal label"], scores=[0.99])
    result = image_detection.combine_detections(image, yolo, ocr, profile="ui")
    assert result["graphic_count"] == result["text_count"] == 1
    graphic, text = result["elements"]
    assert graphic["rect"]["right"] == 201
    assert text["related_element_id"] == graphic["id"]
    assert text["observed_value"] == "literal label"
    assert text["relationships"][0]["type"] == "text-of"
    assert "parent_element_id" not in text


def test_nms_empty_and_duplicate_proposals():
    assert image_detection.nms([], []) == []
    assert image_detection.nms([[0, 0, 30, 30], [1, 1, 31, 31], [50, 50, 70, 70]],
                               [0.8, 0.9, 0.7]) == [1, 2]


def test_ocr_empty_result_and_package_version_do_not_depend_on_module_version_attr(
        monkeypatch, tmp_path):
    options = {}

    def engine(params):
        options.update(params)
        return lambda *_a, **_kw: SimpleNamespace(boxes=None, txts=None, scores=None)

    module = SimpleNamespace(__file__=str(tmp_path / "__init__.py"), RapidOCR=engine)
    monkeypatch.setattr(image_detection, "_rapidocr_module", lambda _p: module)
    monkeypatch.setattr(image_detection, "package_version", lambda _n: "3.10.0")
    result = image_detection.run_ocr(Image.new("RGB", (80, 50), "white"), recover_small_text=False)
    assert result["boxes"] == result["texts"] == result["scores"] == []
    assert result["version"] == "3.10.0"
    assert options["Global.use_cls"] is False
    assert options["Det.model_path"].endswith("PP-OCRv6_det_small.onnx")


def _small_zero():
    image = Image.new("RGB", (80, 60), "white")
    ImageDraw.Draw(image).rectangle((30, 25, 34, 32), outline="black")
    frames = [{"kind": "frame", "rect": {
        "left": 4, "top": 4, "right": 76, "bottom": 56, "width": 72, "height": 52}}]
    quad = [[28, 23], [37, 23], [37, 35], [28, 35]]
    return image, frames, quad


def test_small_text_recovery_uses_real_ink_inside_a_frame_and_avoids_existing_text():
    image, frames, quad = _small_zero()
    candidates, rejected = image_detection._small_text_candidates(
        image, [quad], [0.4], [], frames)
    assert not rejected and len(candidates) == 1
    assert candidates[0]["ink_rect"] == {
        "left": 30, "top": 25, "right": 35, "bottom": 33, "width": 5, "height": 8}
    duplicated, _ = image_detection._small_text_candidates(
        image, [quad], [0.4], [quad], frames)
    assert duplicated == []
    outside, _ = image_detection._small_text_candidates(image, [quad], [0.4], [], [])
    assert outside == []


def test_small_text_recovery_rejects_a_wire_crossing_the_crop():
    image, frames, quad = _small_zero()
    ImageDraw.Draw(image).line((10, 29, 70, 29), fill="black")
    candidates, rejected = image_detection._small_text_candidates(
        image, [quad], [0.4], [], frames)
    assert candidates == []
    assert "connected wire" in rejected[0]["reason"]


@pytest.mark.parametrize("texts,scores,accepted", [
    (["0", "0"], [0.999, 0.998], True),
    (["0", "O"], [0.999, 0.999], False),
    (["0", "0"], [0.94, 0.93], False),
    (["0", "0"], [0.99, 0.89], False),
    (["", ""], [0.99, 0.99], False),
])
def test_recovered_text_requires_agreement_and_confidence(texts, scores, accepted):
    image, frames, quad = _small_zero()
    candidates, _ = image_detection._small_text_candidates(image, [quad], [0.4], [], frames)

    def recognize(patches):
        assert len(patches) == 2 and patches[1].shape[0] > patches[0].shape[0]
        return SimpleNamespace(txts=texts, scores=scores)

    regions, rejected = image_detection._recognize_small_candidates(
        SimpleNamespace(recognize_txt=recognize), image, candidates)
    assert bool(regions) == accepted
    assert bool(rejected) != accepted
    if accepted:
        assert regions[0]["quad"] == [[30, 25], [35, 25], [35, 33], [30, 33]]
        assert regions[0]["text"] == "0"


def test_recovery_provenance_and_existing_text_are_preserved_when_combining():
    image = Image.new("RGB", (240, 160), "white")
    yolo = dict(boxes=[[10, 10, 200, 140]], scores=[0.9])
    old_quad = [[20, 50], [80, 50], [80, 65], [20, 65]]
    new_quad = [[90, 50], [95, 50], [95, 58], [90, 58]]
    detail = dict(detector_score=0.37, detector_quad=new_quad,
                  text="0", enlarged_text="0", recognition_score=0.99,
                  enlarged_recognition_score=0.98)
    ocr = dict(boxes=[old_quad, new_quad], texts=["old text", "0"], scores=[0.99, 0.98],
               base_text_count=1, region_sources=["full-image", "small-text-recovery"],
               region_details=[{}, detail],
               small_text_recovery={"added_count": 1, "elapsed_seconds": 0.1})
    result = image_detection.combine_detections(image, yolo, ocr, profile="ui")
    assert result["graphic_count"] == 1 and result["text_count"] == 2
    assert result["ocr_base_text_count"] == result["ocr_recovery_count"] == 1
    _graphic, old, new = result["elements"]
    assert old["id"] == 2 and old["observed_value"] == "old text"
    assert old["source"] == "rapidocr-pp-ocrv6" and "ocr_recovery" not in old
    assert new["observed_value"] == "0" and new["source"] == "rapidocr-small-text-recovery"
    assert new["ocr_recovery"] == detail and new["detection_confidence"] == 0.37
    assert old["related_element_id"] == new["related_element_id"] == 1


def test_detector_capture_preserves_default_results_without_a_second_model_pass():
    calls = []

    def postprocess(predictions, shape):
        calls.append((predictions, shape))
        return ["default boxes"], ["default scores"]

    capture = image_detection._CaptureTextDetection(postprocess)
    prediction, shape = object(), (60, 80)
    assert capture(prediction, shape) == (["default boxes"], ["default scores"])
    assert calls == [(prediction, shape)]
    assert capture.predictions is prediction and capture.shape == shape
