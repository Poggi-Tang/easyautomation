"""GUI inference must probe and invoke the same runtime without visible windows."""
import json
from pathlib import Path
import sys
from types import SimpleNamespace

import pytest

from easy_uiauto import image_cli, image_model_paths, local_image_worker as worker


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
    assert options["env"][image_model_paths.CONFIG_ENV] == str(tmp_path / "custom.json")
    assert options["env"][image_model_paths.DEVICE_ENV] == "cuda:1"
    assert options["stdin"] == worker.subprocess.DEVNULL and options["capture_output"]
    if sys.platform == "win32":
        assert options["creationflags"] & worker.subprocess.CREATE_NO_WINDOW
    assert dll_calls == ([None, str(tmp_path)] if frozen else [])


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
