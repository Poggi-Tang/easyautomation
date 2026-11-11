"""Exercise distribution and setup decisions without launching any user interface."""
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import sys
import zipfile

import pytest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("package_gui", ROOT / "scripts/package_gui.py")
packager = importlib.util.module_from_spec(spec)
spec.loader.exec_module(packager)


def fake_distribution(tmp_path, monkeypatch):
    source = tmp_path / "built"
    files = ["easy-uiauto-library.exe", "_internal/python313.dll", "_internal/PySide6/Qt6Widgets.dll",
             "_internal/PySide6/Qt6Svg.dll", "_internal/PySide6/plugins/platforms/qwindows.dll",
             "_internal/image-runtime/python.exe", "_internal/image-runtime/python313.dll",
             "_internal/image-runtime/runtime-manifest.json", "THIRD_PARTY_NOTICES.txt",
             "third_party/qt/licenses/LGPL-3.0-only.txt"]
    files += [f"_internal/image-runtime/Lib/site-packages/{name}/__init__.py"
              for name in ("torch", "onnxruntime", "rapidocr", "cv2", "numpy", "PIL")]
    files += ["_internal/image-runtime/models/omniparser/icon_detect_v3/model.pt"]
    files += [f"_internal/image-runtime/Lib/site-packages/rapidocr/models/{name}"
              for name in ("PP-OCRv6_det_small.onnx", "PP-OCRv6_rec_small.onnx")]
    for name in files:
        path = source / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(name.encode())
        if path.suffix in {".pt", ".onnx"}:
            monkeypatch.setitem(packager.deployment.ASSETS, path.name, ("unused", packager.deployment.sha256(path)))
    manifest = dict(version="1.0", frontend_only=False,
                    exe_sha256=packager.deployment.sha256(source / "easy-uiauto-library.exe"),
                    sources={"src/easy_uiauto/qt_gui.py": packager.deployment.sha256(ROOT / "src/easy_uiauto/qt_gui.py")})
    (source / "gui-build.json").write_text(json.dumps(manifest), encoding="utf-8")
    return source


def test_package_contains_complete_launch_path_and_no_old_exes(tmp_path, monkeypatch):
    source = fake_distribution(tmp_path, monkeypatch)
    (source / "old-debug.exe").write_bytes(b"old")
    installer = tmp_path / "VC_redist.x64.exe"
    installer.write_bytes(b"runtime installer")
    monkeypatch.setitem(packager.deployment.ASSETS, installer.name, ("unused", packager.deployment.sha256(installer)))
    archive = packager.package_directory(source, tmp_path / "delivery", vc_installer=installer)
    with zipfile.ZipFile(archive) as zipped:
        names = zipped.namelist()
        prefix = "easy-uiauto-gui-win-x64/"
        for name in ("setup.bat", "scripts/gui_setup.ps1", "gui-distribution.json", "使用说明.txt",
                     "prerequisites/VC_redist.x64.exe", "_internal/image-runtime/python.exe"):
            assert prefix + name in names
        assert not any("old-debug" in name for name in names)
        assert zipped.testzip() is None
    assert archive.with_suffix(".zip.sha256").read_text().startswith(packager.deployment.sha256(archive))


@pytest.mark.parametrize("damage", ["frontend_only", "model", "exe", "source"])
def test_partial_or_stale_build_is_rejected_before_packaging(tmp_path, monkeypatch, damage):
    source = fake_distribution(tmp_path, monkeypatch)
    manifest_path = source / "gui-build.json"
    manifest = json.loads(manifest_path.read_text())
    if damage == "frontend_only":
        manifest["frontend_only"] = True
    elif damage == "source":
        manifest["sources"]["src/easy_uiauto/qt_gui.py"] = "0" * 64
    else:
        path = source / ("easy-uiauto-library.exe" if damage == "exe" else
                         "_internal/image-runtime/models/omniparser/icon_detect_v3/model.pt")
        path.write_bytes(b"changed")
    manifest_path.write_text(json.dumps(manifest))
    output = tmp_path / "delivery"
    with pytest.raises(ValueError):
        packager.package_directory(source, output)
    assert not output.exists()


@pytest.mark.skipif(sys.platform != "win32", reason="Windows PowerShell setup")
def test_source_setup_installs_once_and_updates_same_version_source(tmp_path):
    # A tiny fake installer records calls; no network, GUI, registry or shortcuts.
    source = tmp_path / "源码 with spaces"
    scripts = source / "scripts"
    scripts.mkdir(parents=True)
    module = source / "src/easy_uiauto/qt_gui.py"
    module.parent.mkdir(parents=True)
    module.write_text("first")
    (source / "pyproject.toml").write_text('version = "1.0"')
    (source / "MANIFEST.in").write_text("include setup.bat")
    shutil.copy2(ROOT / "scripts/gui_setup.ps1", scripts / "gui_setup.ps1")
    shutil.copy2(ROOT / "setup.bat", source / "setup.bat")
    (scripts / "full_setup.ps1").write_text('''param([string]$Target, [switch]$NoIntegrate)
if (-not $NoIntegrate) { throw "GUI must not change MCP integration" }
New-Item -ItemType Directory -Path (Join-Path $Target "python/Lib/site-packages/easy_uiauto/qt_ui") -Force > $null
Set-Content -LiteralPath (Join-Path $Target "python/pythonw.exe") -Value "fake"
Set-Content -LiteralPath (Join-Path $Target "python/Lib/site-packages/easy_uiauto/qt_ui/window.py") -Value "fake"
Add-Content -LiteralPath (Join-Path $Target "calls.txt") -Value "install"
$global:LASTEXITCODE = 0
''', encoding="utf-8-sig")
    target = tmp_path / "installed gui"
    command = ["cmd.exe", "/d", "/c", "setup.bat",
               "-Target", str(target), "-NoLaunch", "-NoShortcut"]
    for _ in range(2):
        result = subprocess.run(command, cwd=source, stdin=subprocess.DEVNULL, capture_output=True, timeout=30)
        assert result.returncode == 0, result.stderr.decode(errors="replace")
    assert len((target / "calls.txt").read_text().splitlines()) == 1
    module.write_text("second")
    result = subprocess.run(command, cwd=source, stdin=subprocess.DEVNULL, capture_output=True, timeout=30)
    assert result.returncode == 0, result.stderr.decode(errors="replace")
    assert len((target / "calls.txt").read_text().splitlines()) == 2


@pytest.mark.skipif(sys.platform != "win32", reason="Windows PowerShell setup")
def test_portable_setup_prepares_without_python_install_and_rejects_changed_exe(tmp_path, monkeypatch):
    source = fake_distribution(tmp_path, monkeypatch)
    (source / "scripts").mkdir()
    shutil.copy2(ROOT / "scripts/gui_setup.ps1", source / "scripts/gui_setup.ps1")
    (source / "gui-distribution.json").write_text(json.dumps(dict(format=1,
        exe_sha256=packager.deployment.sha256(source / "easy-uiauto-library.exe"))))
    harness = source / "check.ps1"
    harness.write_text('''function Get-ItemProperty { [pscustomobject]@{ Installed = 1 } }
function Start-Process { throw "No processes should be launched during this check" }
& (Join-Path $PSScriptRoot "scripts/gui_setup.ps1") -NoLaunch -NoShortcut
''', encoding="utf-8-sig")
    command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(harness)]
    result = subprocess.run(command, capture_output=True, timeout=30)
    assert result.returncode == 0, result.stderr.decode(errors="replace")
    assert b"GUI is ready" in result.stdout
    (source / "easy-uiauto-library.exe").write_bytes(b"changed")
    result = subprocess.run(command, capture_output=True, timeout=30)
    assert result.returncode != 0
    assert b"checksum mismatch" in result.stderr
