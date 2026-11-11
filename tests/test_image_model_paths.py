"""Persisted model choices must reach checkers, downloads and both inference engines."""
import json
from pathlib import Path
import subprocess
import sys
from types import SimpleNamespace

import pytest

from easy_uiauto import image_detection, image_model_paths, image_models


@pytest.fixture(autouse=True)
def isolated_config(tmp_path, monkeypatch):
    monkeypatch.setenv(image_model_paths.CONFIG_ENV, str(tmp_path / "config/model-paths.json"))


@pytest.fixture
def selected(tmp_path):
    return {key: str(tmp_path / "本地模型" / name) for key, name in
            (("yolo", "自选.pt"), ("det", "检测.onnx"), ("rec", "识别.onnx"))}


def test_paths_persist_across_processes_and_override_default_locations(selected, tmp_path, monkeypatch):
    image_model_paths.save_model_paths(selected)
    monkeypatch.setenv("EASY_UIAUTO_MODEL_DIR", str(tmp_path / "environment-default"))
    assert image_detection.default_weights() == Path(selected["yolo"])
    assert image_model_paths.ocr_model_paths(tmp_path / "another-runtime") == {
        key: Path(selected[key]) for key in ("det", "rec")}
    output = subprocess.check_output([sys.executable, "-X", "utf8", "-c",
        "import json; from easy_uiauto.image_models import model_specs; print(json.dumps(model_specs()))"],
        text=True, encoding="utf-8")
    assert {model["id"]: model["path"] for model in json.loads(output)} == selected
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path / "frozen"), raising=False)
    assert {model["id"]: model["path"] for model in image_models.model_specs()} == selected


@pytest.mark.parametrize("bad", ["relative.pt", "C:/bad.onnx", 42])
def test_invalid_path_never_replaces_saved_configuration(selected, bad):
    image_model_paths.save_model_paths(selected)
    with pytest.raises(ValueError):
        image_model_paths.save_model_paths({**selected, "yolo": bad})
    assert image_model_paths.load_model_paths() == selected


def test_duplicate_ocr_files_are_rejected(selected):
    with pytest.raises(ValueError, match="同一个"):
        image_model_paths.save_model_paths({**selected, "rec": selected["det"]})


def test_reset_changes_only_preferences_and_keeps_models(selected):
    for value in selected.values():
        path = Path(value)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"existing model")
    image_model_paths.save_model_paths(selected)
    image_model_paths.save_model_paths({key: "" for key in selected})
    assert image_model_paths.load_model_paths() == {}
    assert all(Path(value).read_bytes() == b"existing model" for value in selected.values())


def test_ocr_preflight_and_inference_use_selected_files(selected, tmp_path, monkeypatch):
    from PIL import Image
    image_model_paths.save_model_paths(selected)
    for key in ("det", "rec"):
        path = Path(selected[key]);path.parent.mkdir(parents=True, exist_ok=True);path.touch()
    params = {}
    def engine(*, params: dict):
        captured.update(params)
        return lambda *a, **kw: SimpleNamespace(boxes=None, txts=None, scores=None)
    captured = params
    module = SimpleNamespace(__file__=str(tmp_path / "empty-rapidocr/__init__.py"), RapidOCR=engine)
    monkeypatch.setattr(image_detection, "_rapidocr_module", lambda *a: module)
    real_import = image_detection.importlib.import_module
    monkeypatch.setattr(image_detection.importlib, "import_module",
                        lambda name: object() if name in {"torch", "onnxruntime", "cv2", "numpy"} else real_import(name))
    assert image_detection.check_runtime() is module
    monkeypatch.setattr(image_detection, "package_version", lambda _: "test")
    image_detection.run_ocr(Image.new("RGB", (80, 50)), recover_small_text=False)
    assert captured["Det.model_path"] == selected["det"]
    assert captured["Rec.model_path"] == selected["rec"]
    Path(selected["det"]).unlink()
    with pytest.raises(FileNotFoundError, match="检测"):
        image_detection.check_runtime()


def test_download_targets_persisted_paths_without_touching_default_models(selected, tmp_path, monkeypatch):
    import hashlib
    image_model_paths.save_model_paths(selected)
    specs = image_models.model_specs()
    source = tmp_path / "known-model"
    source.write_bytes(b"valid-test-model")
    for spec in specs:
        spec.update(url=source.as_uri(), sha256=hashlib.sha256(source.read_bytes()).hexdigest())
    monkeypatch.setattr(image_models, "model_specs", lambda: specs)
    monkeypatch.setattr(image_models, "runtime_issues", lambda: [])
    report = image_models.download_missing_models(progress=lambda _: None)
    assert report["ok"]
    assert all(Path(value).read_bytes() == source.read_bytes() for value in selected.values())
