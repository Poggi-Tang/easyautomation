"""Deployment integrity and offline installation boundaries without network access."""

import hashlib
import importlib.util
import json
from pathlib import Path
from types import SimpleNamespace
import zipfile

import pytest

ROOT = Path(__file__).parents[1]
spec = importlib.util.spec_from_file_location("standalone_deployment", ROOT / "src/easy_uiauto/deployment.py")
deployment = importlib.util.module_from_spec(spec)
spec.loader.exec_module(deployment)


def wheel(directory, name, version="1.0", *, extra="", exe=False):
    path = directory / f"{name}-{version}-py3-none-any.whl"
    with zipfile.ZipFile(path, "w") as archive:
        metadata = f"Metadata-Version: 2.1\nName: {name}\nVersion: {version}\n"
        if extra:
            metadata += f"Provides-Extra: {extra}\n"
        archive.writestr(f"{name}-{version}.dist-info/METADATA", metadata)
        if exe:
            archive.writestr("easy_uiauto/bin/easy-uiauto-config.exe", b"MZ")
    return path


def bundle(tmp_path, monkeypatch):
    root = tmp_path / "bundle"
    for name in ("payload/python.zip", "payload/VC_redist.x64.exe", "requirements.lock",
                 "models/omniparser/icon_detect_v3/model.pt", "tesseract/tesseract.exe",
                 "tesseract/tessdata/eng.traineddata", "tesseract/tessdata/chi_sim.traineddata",
                 "deployment.py", "wheelhouse/example.whl"):
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(name.encode())
    monkeypatch.setattr(deployment, "PYTHON_SHA256", deployment.sha256(root / "payload/python.zip"))
    monkeypatch.setitem(deployment.ASSETS, "model.pt", ("unused", deployment.sha256(root / deployment.MODEL_RELATIVE)))
    deployment.write_manifest(root, "1.2.3", {})
    return root


def test_full_extra_covers_all_functional_extras():
    import tomllib

    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    extras = data["project"]["optional-dependencies"]
    assert set(extras["full"]) == set(extras["mcp"] + extras["image"] + extras["vision"])
    assert data["project"]["scripts"]["easy_uiauto_setup"] == "easy_uiauto.deployment:main"


def test_installation_fixture_matches_the_public_example():
    source = ROOT / "docs/examples/IEEE9/IEEE9.png"
    fixture = ROOT / "src/easy_uiauto/resources/IEEE9.png"
    assert fixture.read_bytes() == source.read_bytes()


@pytest.mark.parametrize("name", ["../escape", "/absolute", "C:/absolute", "nested/../../escape", "nested\\escape", "."])
def test_manifest_paths_cannot_escape(tmp_path, name):
    with pytest.raises(ValueError):
        deployment.contained(tmp_path, name)


def test_archive_traversal_rejected_before_any_extraction(tmp_path):
    path = tmp_path / "bad.zip"
    target = tmp_path / "target"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("safe.txt", "safe")
        archive.writestr("../escape.txt", "bad")
    with pytest.raises(ValueError):
        deployment.extract_zip(path, target)
    assert not target.exists()
    assert not (tmp_path / "escape.txt").exists()


def test_wheel_lock_hashes_built_and_downloaded_packages(tmp_path):
    files = [wheel(tmp_path, name) for name in ("easy_uiauto", "torch", "rapidocr", "pyautogui")]
    lock = deployment.wheel_lock(tmp_path)
    for path in files:
        assert "--hash=sha256:" + hashlib.sha256(path.read_bytes()).hexdigest() in lock
    assert "pyautogui==1.0" in lock
    wheel(tmp_path, "torch", "2.0")
    with pytest.raises(ValueError, match="Multiple wheels"):
        deployment.wheel_lock(tmp_path)


def test_incomplete_wheelhouse_is_not_an_offline_bundle(tmp_path):
    wheel(tmp_path, "easy_uiauto")
    with pytest.raises(ValueError, match="Incomplete wheelhouse"):
        deployment.wheel_lock(tmp_path)


def test_vendored_metadata_does_not_confuse_wheel_identity(tmp_path):
    for name in ("easy_uiauto", "torch", "rapidocr"):
        wheel(tmp_path, name)
    path = wheel(tmp_path, "setuptools", "84.2.1")
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("setuptools/_vendor/other-1.0.dist-info/METADATA", "Name: other\nVersion: 1.0\n")
    assert "setuptools==84.2.1" in deployment.wheel_lock(tmp_path)


def test_project_wheel_requires_full_extra_and_exe(tmp_path):
    path = wheel(tmp_path, "easy_uiauto")
    with pytest.raises(ValueError, match="full extra"):
        deployment.project_wheel("1.0", tmp_path, path)
    path = wheel(tmp_path, "easy_uiauto", extra="full")
    with pytest.raises(ValueError, match="EXE"):
        deployment.project_wheel("1.0", tmp_path, path)
    path = wheel(tmp_path, "easy_uiauto", extra="full", exe=True)
    assert deployment.project_wheel("1.0", tmp_path, path) == path
    with pytest.raises(ValueError):
        deployment.project_wheel("2.0", tmp_path, path)


def test_tampered_bundle_fails_before_target_changes(tmp_path, monkeypatch):
    root = bundle(tmp_path, monkeypatch)
    assert deployment.verify_bundle(root)["version"] == "1.2.3"
    (root / "wheelhouse/example.whl").write_bytes(b"changed")
    monkeypatch.setattr(deployment, "require_windows", lambda: None)
    target = tmp_path / "installed"
    with pytest.raises(ValueError, match="SHA256 mismatch"):
        deployment.install_bundle(root, target, no_integrate=True)
    assert not target.exists()


def test_missing_and_unlisted_files_are_rejected(tmp_path, monkeypatch):
    root = bundle(tmp_path, monkeypatch)
    extra = root / "wheelhouse/unlisted.whl"
    extra.write_bytes(b"extra")
    with pytest.raises(ValueError, match="file list"):
        deployment.verify_bundle(root)
    extra.unlink()
    (root / "tesseract/tessdata/chi_sim.traineddata").unlink()
    with pytest.raises(ValueError, match="file list"):
        deployment.verify_bundle(root)


def test_existing_user_directory_is_not_overwritten(tmp_path, monkeypatch):
    root = bundle(tmp_path, monkeypatch)
    target = tmp_path / "personal"
    target.mkdir()
    (target / "user.txt").write_text("keep")
    monkeypatch.setattr(deployment, "require_windows", lambda: None)
    with pytest.raises(ValueError, match="empty target"):
        deployment.install_bundle(root, target)
    assert (target / "user.txt").read_text() == "keep"


def test_offline_install_uses_only_hash_locked_local_wheels(tmp_path, monkeypatch):
    root = bundle(tmp_path, monkeypatch)
    target = tmp_path / "installed"
    calls = []
    monkeypatch.setattr(deployment, "require_windows", lambda: None)
    monkeypatch.setattr(deployment, "ensure_vc_runtime", lambda *_args: None)
    monkeypatch.setattr(deployment, "extract_zip", lambda *_args: None)
    monkeypatch.setattr(deployment, "verify_runtime", lambda _target: {"ok": True})
    monkeypatch.setattr(deployment, "run", lambda args, **kwargs: calls.append((args, kwargs)))
    monkeypatch.setattr(deployment.urllib.request, "urlopen", lambda *_args, **_kwargs: pytest.fail("Offline installation accessed the network"))
    deployment.install_bundle(root, target, no_integrate=True)
    pip = next((args, options) for args, options in calls if "install" in args)
    assert {"--isolated", "--no-index", "--no-deps", "--require-hashes"}.issubset(pip[0])
    assert pip[1]["env"]["PIP_NO_INDEX"] == "1"
    assert root / "wheelhouse" in pip[0]
    assert json.loads((target / "deployment.json").read_text())["status"] == "ready"


def test_failed_self_check_never_marks_installation_ready(tmp_path, monkeypatch):
    root = bundle(tmp_path, monkeypatch)
    target = tmp_path / "installed"
    monkeypatch.setattr(deployment, "require_windows", lambda: None)
    monkeypatch.setattr(deployment, "ensure_vc_runtime", lambda *_args: None)
    monkeypatch.setattr(deployment, "extract_zip", lambda *_args: None)
    monkeypatch.setattr(deployment, "run", lambda *_args, **_kwargs: "")
    monkeypatch.setattr(deployment, "verify_runtime", lambda _target: (_ for _ in ()).throw(RuntimeError("OCR failed")))
    with pytest.raises(RuntimeError, match="OCR failed"):
        deployment.install_bundle(root, target, no_integrate=True)
    assert json.loads((target / "deployment.json").read_text())["status"] == "installing"
    assert not (target / "verification.json").exists()


def test_bad_download_never_replaces_a_verified_cache_file(tmp_path, monkeypatch):
    class Response:
        def __enter__(self):
            import io
            return io.BytesIO(b"wrong")

        def __exit__(self, *_args):
            pass

    path = tmp_path / "asset"
    path.write_bytes(b"old")
    monkeypatch.setattr(deployment.urllib.request, "urlopen", lambda *_args, **_kwargs: Response())
    with pytest.raises(RuntimeError, match="SHA256 mismatch"):
        deployment.download("https://example.invalid", path, "0" * 64)
    assert path.read_bytes() == b"old"
    assert not list(tmp_path.glob("*.partial"))


def test_bootstrap_is_standalone_and_offline_entry_has_no_downloads():
    assert "Invoke-WebRequest" not in deployment.OFFLINE_PS
    assert "Get-FileHash" in deployment.OFFLINE_PS
    assert "Expand-Archive" in deployment.OFFLINE_PS


def test_python_children_force_utf8_even_in_isolated_mode(monkeypatch, tmp_path):
    commands = []
    def capture(command, **kwargs):
        commands.append(command)
        return SimpleNamespace(returncode=0, stdout="中文", stderr="")
    monkeypatch.setattr(deployment.subprocess, "run", capture)
    result = deployment.run(["C:/runtime/python.exe", "-I", "-m", "pip", "check"], log=tmp_path / "setup.log")
    assert commands[0][:4] == ["C:/runtime/python.exe", "-X", "utf8", "-I"]
    assert result == "中文"


def test_model_directory_can_be_configured(monkeypatch, tmp_path):
    from easy_uiauto.image_detection import default_weights

    monkeypatch.setenv("EASY_UIAUTO_MODEL_DIR", str(tmp_path))
    assert default_weights() == tmp_path / "omniparser/icon_detect_v3/model.pt"


def test_detector_can_load_non_ascii_windows_paths(monkeypatch, tmp_path):
    from easy_uiauto import image_detection
    import sys

    weights = tmp_path / "模型目录" / "model.pt"
    weights.parent.mkdir()
    weights.write_bytes(b"serialized model")
    loaded = []
    model = SimpleNamespace(eval=lambda: "ready")

    def load(stream, **kwargs):
        assert not isinstance(stream, str), "Native fopen must not receive a non-ASCII path"
        loaded.append(stream.read())
        return model

    monkeypatch.setitem(sys.modules, "torch", SimpleNamespace(
        set_num_threads=lambda _threads: None, jit=SimpleNamespace(load=load)))
    monkeypatch.setattr(image_detection, "_predict_yolo", lambda *_args, **_kwargs: dict(
        boxes=[[1, 2, 20, 30]], scores=[0.9], forward_seconds=0.01, tensor_shape=[1, 3, 1280, 1280]))
    assert image_detection.run_yolo(None, weights)["boxes"] == [[1, 2, 20, 30]]
    assert loaded == [b"serialized model"]
