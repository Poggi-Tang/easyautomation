"""Optional GPU installation is transactional and relocatable; no downloads."""
import importlib.util
from pathlib import Path
import shutil
import subprocess
import sys

import pytest

spec = importlib.util.spec_from_file_location("install_gpu", Path(__file__).resolve().parents[1] / "scripts/install_gpu.py")
installer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(installer)


@pytest.fixture
def installation(tmp_path, monkeypatch):
    site = tmp_path / "软件" / "site-packages"
    site.mkdir(parents=True)
    cpu = site / "torch" / "__init__.py"
    cpu.parent.mkdir()
    cpu.write_text("device = 'cpu'")
    monkeypatch.setattr(installer.sys, "platform", "win32")
    monkeypatch.setattr(installer.sysconfig, "get_path", lambda _: str(site))
    monkeypatch.setattr(installer, "pip_command", lambda cache: ["-m", "pip"])
    calls = []
    def run(args, **kwargs):
        calls.append(args)
        if "download" in args:
            destination = Path(args[args.index("--dest") + 1])
            destination.mkdir()
            (destination / "torch-test.whl").touch()
        else:
            stage = Path(args[args.index("--target") + 1])
            (stage / "torch").mkdir()
            (stage / "torch/__init__.py").write_text("device = 'cuda'")
    monkeypatch.setattr(installer, "run", run)
    return site, calls


def test_failed_verification_preserves_existing_runtime(installation, monkeypatch):
    site, _ = installation
    marker = site / "easy_uiauto_gpu.pth"
    marker.write_bytes(b"previous configuration")
    def reject(stage):
        assert marker.read_bytes() == b"previous configuration"
        raise RuntimeError("incompatible driver")
    monkeypatch.setattr(installer, "verify_stage", reject)
    with pytest.raises(RuntimeError, match="driver"):
        installer.install()
    assert marker.read_bytes() == b"previous configuration"
    assert (site / "torch/__init__.py").read_text() == "device = 'cpu'"
    assert not list(site.glob("easy_uiauto_gpu_*"))


def test_activation_after_verification_survives_move_and_disable(installation, monkeypatch):
    site, calls = installation
    verified = []
    def verify(stage):
        assert not (site / "easy_uiauto_gpu.pth").exists()
        verified.append(stage)
    monkeypatch.setattr(installer, "verify_stage", verify)
    installer.install()
    assert len(verified) == 1
    assert "https://download.pytorch.org/whl/cu128" in calls[0]
    assert calls[1][-1].endswith("torch-test.whl")
    moved = site.parent / "移动后"
    shutil.move(str(site), moved)
    code = "import site, sys; site.addsitedir(sys.argv[1]); import torch; print(torch.device)"
    result = subprocess.check_output([sys.executable, "-I", "-S", "-c", code, str(moved)], text=True)
    assert result.strip() == "cuda"
    monkeypatch.setattr(installer.sysconfig, "get_path", lambda _: str(moved))
    installer.main(["--disable"])
    result = subprocess.check_output([sys.executable, "-I", "-S", "-c", code, str(moved)], text=True)
    assert result.strip() == "cpu"
    assert (moved / verified[0].name / "torch/__init__.py").is_file()
