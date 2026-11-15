"""GUI inference must probe and invoke the same runtime without visible windows."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from types import SimpleNamespace

import pytest

from easy_uiauto import image_cli, image_model_paths, local_image_worker as worker


def test_probe_can_load_user_site_installation(tmp_path):
    """A real child interpreter must see the same pip --user packages as the GUI."""
    python = getattr(sys, "_base_executable", sys.executable)
    environment = os.environ.copy()
    environment.pop("PYTHONPATH", None)
    environment.pop("PYTHONNOUSERSITE", None)
    environment["PYTHONUSERBASE"] = str(tmp_path / "user-base")
    environment["PYTHONIOENCODING"] = "utf-8"
    environment[image_model_paths.CONFIG_ENV] = str(tmp_path / "models.json")
    environment[image_model_paths.DEVICE_ENV] = "cpu"
    options = dict(env=environment, cwd=tmp_path, capture_output=True, text=True,
                   encoding="utf-8", timeout=30,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    result = subprocess.run([python, "-P", "-X", "utf8", "-c",
        "import site; print(site.getusersitepackages())"], check=True, **options)
    user_site = Path(result.stdout.strip())
    assert user_site.is_relative_to(tmp_path)
    package = user_site / "easy_uiauto"
    package.mkdir(parents=True)
    source = Path(worker.__file__).parent
    for filename in ("__init__.py", "local_image_worker.py", "inference_device.py", "image_model_paths.py"):
        shutil.copyfile(source / filename, package / filename)
    (user_site / "torch.py").write_text(
        "from types import SimpleNamespace\n__version__ = 'user-site-test'\n"
        "version = SimpleNamespace(cuda=None)\n", encoding="utf-8")
    result = subprocess.run([python, "-P", "-X", "utf8", "-c",
        "import json, easy_uiauto; from easy_uiauto.local_image_worker import inspect_yolo_device; "
        "print(json.dumps(dict(package=easy_uiauto.__file__, report=inspect_yolo_device())))"], **options)
    assert result.returncode == 0, result.stderr
    response = json.loads(result.stdout)
    assert Path(response["package"]).parent == package
    assert response["report"]["ok"] and response["report"]["device"] == "cpu"
    assert response["report"]["torch_version"] == "user-site-test"


@pytest.mark.parametrize("frozen", [False, True])
def test_probe_uses_exact_runtime_and_shared_config(tmp_path, monkeypatch, frozen):
    monkeypatch.setenv(image_model_paths.CONFIG_ENV, str(tmp_path / "custom.json"))
    monkeypatch.setenv(image_model_paths.DEVICE_ENV, "cuda:1")
    monkeypatch.setattr(sys, "frozen", frozen, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
    dll_calls = []
    if frozen:
        import ctypes
        monkeypatch.setattr(ctypes, "windll", SimpleNamespace(kernel32=SimpleNamespace(
            SetDllDirectoryW=dll_calls.append)), raising=False)
    calls = []
    def run(args, **options):
        calls.append((args, options))
        return SimpleNamespace(returncode=0, stdout=json.dumps(dict(ok=True, device="cuda:1")))
    monkeypatch.setattr(worker.subprocess, "run", run)
    assert worker.inspect_yolo_device()["device"] == "cuda:1"
    args, options = calls[0]
    assert Path(args[0]) == (tmp_path / "image-runtime/python.exe" if frozen else Path(sys.executable))
    assert ("-I" in args) == (frozen or bool(sys.flags.isolated))
    assert options["env"][image_model_paths.CONFIG_ENV] == str(tmp_path / "custom.json")
    assert options["env"][image_model_paths.DEVICE_ENV] == "cuda:1"
    assert options["stdin"] == worker.subprocess.DEVNULL and options["capture_output"]
    if sys.platform == "win32":
        assert options["creationflags"] & worker.subprocess.CREATE_NO_WINDOW
    assert dll_calls == ([None, str(tmp_path)] if frozen else [])


@pytest.mark.parametrize("error", [
    "ModuleNotFoundError: No module named 'easy_uiauto'",
    "OSError: [WinError 126] Error loading torch/lib/c10.dll",
])
def test_failed_probe_reports_import_or_dll_error(monkeypatch, error):
    monkeypatch.setattr(sys, "frozen", False, raising=False)
    monkeypatch.setattr(worker.subprocess, "run", lambda *a, **kw:
        SimpleNamespace(returncode=1, stdout="", stderr="Traceback (most recent call last):\n" + error + "\n"))
    with pytest.raises(RuntimeError) as failure:
        worker.inspect_yolo_device()
    assert error in str(failure.value)
    assert "退出码 1" in str(failure.value)


def test_worker_request_passes_device_to_analysis(tmp_path, monkeypatch):
    request = tmp_path / "request.json"
    request.write_text(json.dumps(dict(source="in.png", output="out", device="cuda:1")))
    monkeypatch.setattr(sys, "argv", ["worker", str(request)])
    calls = []
    def analyze(*args, **options):
        calls.append(options)
        return {"ok": True}
    monkeypatch.setattr(image_cli, "analyze_local_image", analyze)
    worker.main()
    assert calls[0]["device"] == "cuda:1"
    assert json.loads((tmp_path / "result.json").read_text())["ok"]
