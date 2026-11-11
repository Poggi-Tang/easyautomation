"""Model status must describe inference paths; incomplete downloads never become weights."""
import hashlib
from pathlib import Path
import sys

import pytest

from easy_uiauto import image_detection, image_models, image_model_paths


@pytest.fixture(autouse=True)
def isolated_path_configuration(tmp_path, monkeypatch):
    monkeypatch.setenv(image_model_paths.CONFIG_ENV, str(tmp_path / "config/model-paths.json"))


@pytest.fixture
def models(tmp_path, monkeypatch):
    specs = []
    for key in ("yolo", "det", "rec"):
        source = tmp_path / (key + ".download")
        source.write_bytes((key * 1000).encode())
        specs.append(dict(id=key, name=key, path=str(tmp_path / "models" / key),
                          url=source.as_uri(), sha256=hashlib.sha256(source.read_bytes()).hexdigest()))
    monkeypatch.setattr(image_models, "model_specs", lambda: specs)
    monkeypatch.setattr(image_models, "runtime_issues", lambda: [])
    return specs


def test_missing_corrupt_and_valid_models_are_distinguished_and_repaired(models):
    Path(models[0]["path"]).parent.mkdir()
    Path(models[0]["path"]).write_bytes(b"broken old weights")
    image_models.download_model(models[1], progress=lambda _: None)
    report = image_models.inspect_models()
    assert [m["state"] for m in report["models"]] == ["invalid", "ready", "missing"]
    assert not report["ok"]
    report = image_models.download_missing_models(progress=lambda _: None)
    assert report["ok"]
    assert all(m["state"] == "ready" for m in report["models"])


def test_valid_models_are_not_downloaded_again(models, monkeypatch):
    image_models.download_missing_models(progress=lambda _: None)
    monkeypatch.setattr(image_models.request, "urlopen", lambda *a, **kw: pytest.fail("unnecessary download"))
    assert image_models.download_missing_models(progress=lambda _: None)["ok"]


@pytest.mark.parametrize("failure", ["checksum", "cancel", "network"])
def test_failed_download_preserves_old_file_and_removes_partial(models, monkeypatch, failure):
    spec = models[0].copy()
    target = Path(spec["path"])
    target.parent.mkdir()
    target.write_bytes(b"old model")
    if failure == "checksum":
        spec["sha256"] = "wrong digest"
    if failure == "network":
        def offline(*args, **kwargs):
            raise OSError("network unavailable")
        monkeypatch.setattr(image_models.request, "urlopen", offline)

    def progress(message):
        if failure == "cancel" and message.startswith("正在保存"):
            raise InterruptedError("cancelled")

    with pytest.raises((ValueError, OSError)):
        image_models.download_model(spec, progress=progress)
    assert target.read_bytes() == b"old model"
    assert list(target.parent.iterdir()) == [target]


def test_runtime_failure_is_separate_from_three_valid_models(models, monkeypatch):
    image_models.download_missing_models(progress=lambda _: None)
    monkeypatch.setattr(image_models, "runtime_issues", lambda: ["缺少运行依赖：torch"])
    report = image_models.inspect_models()
    assert not report["ok"] and report["runtime_issues"]
    assert all(m["state"] == "ready" for m in report["models"])


def test_frozen_checker_uses_the_bundled_inference_paths(tmp_path, monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
    monkeypatch.delenv("EASY_UIAUTO_MODEL_DIR", raising=False)
    specs = image_models.model_specs()
    runtime = tmp_path / "image-runtime"
    assert Path(specs[0]["path"]) == runtime / "models/omniparser/icon_detect_v3/model.pt"
    assert Path(specs[1]["path"]) == runtime / "Lib/site-packages/rapidocr/models/PP-OCRv6_det_small.onnx"
    assert Path(specs[2]["path"]) == runtime / "Lib/site-packages/rapidocr/models/PP-OCRv6_rec_small.onnx"
    assert any("Python" in issue for issue in image_models.runtime_issues())
    monkeypatch.setenv("EASY_UIAUTO_MODEL_DIR", str(tmp_path / "custom"))
    assert image_models.model_specs()[0]["path"] == str(image_detection.default_weights().resolve())
    assert Path(image_models.model_specs()[0]["path"]).is_relative_to(tmp_path / "custom")


def test_source_checker_uses_the_actual_rapidocr_package():
    spec = __import__("importlib.util", fromlist=["find_spec"]).find_spec("rapidocr")
    if spec is None:
        pytest.skip("optional image runtime unavailable")
    assert image_models.rapidocr_directory() == Path(spec.origin).parent


def test_download_worker_reports_ready_without_connecting_to_a_window(models, tmp_path):
    import json
    from easy_uiauto import library_tasks
    folder = tmp_path / "job"
    folder.mkdir()
    library_tasks._write(folder / "request.json", dict(operation="download-image-models"))
    library_tasks._worker(folder)
    state = json.loads((folder / "state.json").read_text(encoding="utf-8"))
    assert state["state"] == "completed" and state["result"]["ok"]


def test_cancelled_download_worker_does_not_publish_any_models(models, tmp_path):
    import json
    from easy_uiauto import library_tasks
    folder = tmp_path / "job"
    folder.mkdir()
    library_tasks._write(folder / "request.json", dict(operation="download-image-models"))
    (folder / "cancel").touch()
    library_tasks._worker(folder)
    state = json.loads((folder / "state.json").read_text(encoding="utf-8"))
    assert state["state"] == "cancelled"
    assert all(not Path(spec["path"]).exists() for spec in models)


def test_forced_stop_cleans_only_its_recorded_model_temporary_file(models, tmp_path, monkeypatch):
    import json
    from uuid import uuid4
    from easy_uiauto import library_tasks
    target = Path(models[0]["path"])
    target.parent.mkdir()
    target.write_bytes(b"original model")
    temporary = target.with_name(target.name + "." + uuid4().hex + ".partial")
    temporary.write_bytes(b"unfinished")
    unrelated = target.with_name("another-model." + uuid4().hex + ".partial")
    unrelated.write_bytes(b"keep")
    task = object.__new__(library_tasks.LibraryTask)
    task.operation = "download-image-models"
    task.base = tmp_path / "jobs"
    task.folder = task.base / uuid4().hex
    task.folder.mkdir(parents=True)
    (task.folder / "download-files.json").write_text(json.dumps([str(temporary), str(target), str(unrelated)]))
    task._remove_folder()
    assert not temporary.exists() and not task.folder.exists()
    assert target.read_bytes() == b"original model"
    assert unrelated.read_bytes() == b"keep"
