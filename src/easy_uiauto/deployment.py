"""Install the complete Windows runtime, or prepare a self-contained offline ZIP.

This module deliberately uses only the standard library and also runs as a file
before easy-uiauto or any of its dependencies have been installed.
"""

from __future__ import annotations

import argparse
import ast
import ctypes
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
import uuid
import zipfile
from email.parser import BytesParser
from pathlib import Path, PurePosixPath

PYTHON_VERSION = "3.13.17"
PYTHON_URL = f"https://www.python.org/ftp/python/{PYTHON_VERSION}/python-{PYTHON_VERSION}-amd64.zip"
PYTHON_SHA256 = "b26ca4b4ee51e157836049934d2b8eabd55939be296f3130254251b6df96ec2f"
TESSDATA_COMMIT = "87416418657359cb625c412a48b6e1d6d41c29bd"
ASSETS = {
    "python.zip": (PYTHON_URL, PYTHON_SHA256),
    "VC_redist.x64.exe": (
        "https://download.visualstudio.microsoft.com/download/pr/ebdab8e5-1d7b-4d9f-a11b-cbb1720c3b12/"
        "843068991DAAA1F73AD9F6239BCE4D0F6A07A51F18C37EA2A867E9BECA71295C/VC_redist.x64.exe",
        "843068991daaa1f73ad9f6239bce4d0f6a07a51f18c37ea2a867e9beca71295c"),
    "7zr.exe": (
        "https://github.com/ip7z/7zip/releases/download/25.01/7zr.exe",
        "27cbe3d5804ad09e90bbcaa916da0d5c3b0be9462d0e0fb6cb54be5ed9030875"),
    "7zip.exe": (
        "https://github.com/ip7z/7zip/releases/download/25.01/7z2501-x64.exe",
        "78afa2a1c773caf3cf7edf62f857d2a8a5da55fb0fff5da416074c0d28b2b55f"),
    "tesseract-installer.exe": (
        "https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/"
        "tesseract-ocr-w64-setup-5.4.0.20240606.exe",
        "c885fff6998e0608ba4bb8ab51436e1c6775c2bafc2559a19b423e18678b60c9"),
    "eng.traineddata": (
        f"https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/{TESSDATA_COMMIT}/eng.traineddata",
        "7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2"),
    "chi_sim.traineddata": (
        f"https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/{TESSDATA_COMMIT}/chi_sim.traineddata",
        "a5fcb6f0db1e1d6d8522f39db4e848f05984669172e584e8d76b6b3141e1f730"),
    "model.pt": (
        "https://huggingface.co/microsoft/OmniParser-v2.0/resolve/main/icon_detect_v3/model.pt",
        "11c6cbb77f22569fab22d86c76407a83ec81ab89dbfe28279854822d6e3fb00c"),
}
MODEL_RELATIVE = Path("models/omniparser/icon_detect_v3/model.pt")
TORCH_INDEX = "https://download.pytorch.org/whl/cpu"


def package_version() -> str:
    source = Path(__file__).with_name("__init__.py")
    if source.is_file():
        for node in ast.parse(source.read_text(encoding="utf-8")).body:
            if isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id == "__version__" for t in node.targets
            ):
                return ast.literal_eval(node.value)
    # Standalone copies obtain their release version from the bundle manifest.
    manifest = Path(__file__).with_name("manifest.json")
    if manifest.is_file():
        return json.loads(manifest.read_text(encoding="utf-8"))["version"]
    raise RuntimeError("Use --version when running a standalone deployment.py")


def default_home() -> Path:
    return Path(os.environ.get("LOCALAPPDATA", Path.home() / ".local")) / "easy-uiauto"


def require_windows() -> None:
    if os.name != "nt" or platform.machine().lower() not in ("amd64", "x86_64"):
        raise RuntimeError("Full deployment supports Windows x64 only")


def sha256(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def download(url: str, destination: Path, expected: str) -> Path:
    """Resume through verified cache entries; never publish a partial file."""
    if destination.is_file() and sha256(destination) == expected:
        return destination
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(destination.name + f".{uuid.uuid4().hex}.partial")
    try:
        print(f"Downloading {destination.name}", flush=True)
        request = urllib.request.Request(url, headers={"User-Agent": "easy-uiauto-setup"})
        with urllib.request.urlopen(request, timeout=90) as response, temporary.open("xb") as out:
            shutil.copyfileobj(response, out, length=1024 * 1024)
        if sha256(temporary) != expected:
            raise RuntimeError(f"SHA256 mismatch: {destination.name}")
        temporary.replace(destination)
    finally:
        temporary.unlink(missing_ok=True)
    return destination


def contained(root: Path, name: str) -> Path:
    """Validate portable manifest/archive paths before reading or writing them."""
    parts = PurePosixPath(name)
    if not name or "\\" in name or ":" in name or parts.is_absolute() or ".." in parts.parts:
        raise ValueError(f"Unsafe relative path: {name}")
    target = (root / name).resolve()
    if not target.is_relative_to(root.resolve()) or target == root.resolve():
        raise ValueError(f"Path escapes its root: {name}")
    return target


def extract_zip(archive: Path, destination: Path) -> None:
    with zipfile.ZipFile(archive) as source:
        for item in source.infolist():
            contained(destination, item.filename.rstrip("/"))
            if (item.external_attr >> 16) & 0o170000 == 0o120000:
                raise ValueError("ZIP symbolic links are not supported")
        source.extractall(destination)


def run(command: list[str], *, log: Path, env=None, timeout=1800, accepted=(0,)) -> str:
    log.parent.mkdir(parents=True, exist_ok=True)
    command = [str(x) for x in command]
    if Path(command[0]).name.lower() in {"python", "python.exe", "python3", "python3.exe"}:
        command[1:1] = ["-X", "utf8"]
    result = subprocess.run(
        command, stdin=subprocess.DEVNULL, capture_output=True,
        text=True, encoding="utf-8", errors="replace", timeout=timeout, env=env,
    )
    output = result.stdout + result.stderr
    with log.open("a", encoding="utf-8") as stream:
        stream.write(output + "\n")
    if result.returncode not in accepted:
        raise RuntimeError(f"Command failed ({result.returncode}); see {log}\n{output[-4000:]}")
    return result.stdout


def wheel_lock(wheelhouse: Path) -> str:
    """Lock the complete resolved closure, including locally built pure wheels."""
    records = {}
    for path in sorted(wheelhouse.glob("*.whl")):
        with zipfile.ZipFile(path) as source:
            names = [n for n in source.namelist() if n.count("/") == 1 and n.endswith(".dist-info/METADATA")]
            if len(names) != 1:
                raise ValueError(f"Invalid wheel metadata: {path.name}")
            metadata = BytesParser().parsebytes(source.read(names[0]))
        name, version = metadata["Name"], metadata["Version"]
        key = name.lower().replace("_", "-").replace(".", "-")
        if key in records:
            raise ValueError(f"Multiple wheels for {name}; start with a clean wheelhouse")
        records[key] = f"{name}=={version} --hash=sha256:{sha256(path)}"
    if "easy-uiauto" not in records or "torch" not in records or "rapidocr" not in records:
        raise ValueError("Incomplete wheelhouse: easy-uiauto, torch and RapidOCR are required")
    return "\n".join(records[name] for name in sorted(records)) + "\n"


def project_wheel(version: str, cache: Path, wheel: Path | None) -> Path:
    if wheel:
        path = wheel.expanduser().resolve(strict=True)
    else:
        with urllib.request.urlopen(f"https://pypi.org/pypi/easy-uiauto/{version}/json", timeout=60) as response:
            data = json.load(response)
        selected = [x for x in data["urls"] if x["filename"].endswith("-win_amd64.whl")]
        if len(selected) != 1:
            raise RuntimeError(f"Windows x64 wheel unavailable for easy-uiauto {version}")
        item = selected[0]
        path = download(item["url"], cache / item["filename"], item["digests"]["sha256"])
    with zipfile.ZipFile(path) as source:
        meta = next(n for n in source.namelist() if n.count("/") == 1 and n.endswith(".dist-info/METADATA"))
        metadata = BytesParser().parsebytes(source.read(meta))
        if metadata["Version"] != version or "full" not in metadata.get_all("Provides-Extra", []):
            raise ValueError("Project wheel must match --version and provide the full extra")
        if "easy_uiauto/bin/easy-uiauto-config.exe" not in source.namelist():
            raise ValueError("Project wheel is missing the bundled configuration EXE")
    return path


OFFLINE_PS = r'''param([string]$Target = "", [ValidateSet("auto","codex","claude","none")][string]$Client = "auto", [switch]$NoIntegrate)
$ErrorActionPreference = "Stop"
$OutputEncoding = [Text.UTF8Encoding]::new()
[Console]::OutputEncoding = $OutputEncoding
$bundleRoot = $PSScriptRoot
$manifest = Get-Content -LiteralPath (Join-Path $bundleRoot "manifest.json") -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($entry in $manifest.files.PSObject.Properties) {
    $item = [IO.Path]::GetFullPath((Join-Path $bundleRoot $entry.Name))
    if (-not $item.StartsWith($bundleRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw "Invalid bundle path" }
    if ((Get-FileHash -LiteralPath $item -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.Value) { throw "Checksum mismatch: $($entry.Name)" }
}
$bootstrapRoot = Join-Path ([IO.Path]::GetTempPath()) ("easy-uiauto-" + [Guid]::NewGuid().ToString("N"))
Expand-Archive -LiteralPath (Join-Path $bundleRoot "payload/python.zip") -DestinationPath $bootstrapRoot
$arguments = @("-I", "-X", "utf8", (Join-Path $bundleRoot "deployment.py"), "install-offline", $bundleRoot, "--client", $Client)
if ($Target) { $arguments += @("--target", $Target) }
if ($NoIntegrate) { $arguments += "--no-integrate" }
try {
    & (Join-Path $bootstrapRoot "python.exe") @arguments
    if ($LASTEXITCODE -ne 0) { throw "Installation failed ($LASTEXITCODE)" }
} finally {
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    $resolvedBootstrap = [IO.Path]::GetFullPath($bootstrapRoot)
    if (-not $resolvedBootstrap.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) { throw "Invalid temporary path" }
    if (-not ([IO.Path]::GetFileName($resolvedBootstrap)).StartsWith("easy-uiauto-")) { throw "Invalid temporary directory" }
    Remove-Item -LiteralPath $resolvedBootstrap -Recurse -Force
}
'''


def write_manifest(root: Path, version: str, sources: dict) -> dict:
    files = {p.relative_to(root).as_posix(): sha256(p)
             for p in sorted(root.rglob("*")) if p.is_file() and p != root / "manifest.json"}
    data = dict(format=1, version=version, platform="win_amd64", python=PYTHON_VERSION,
                files=files, sources=sources)
    (root / "manifest.json").write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
    return data


def verify_bundle(root: Path) -> dict:
    data = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
    if data.get("format") != 1 or data.get("platform") != "win_amd64" or data.get("python") != PYTHON_VERSION:
        raise ValueError("Unsupported offline bundle format/platform/Python")
    required = {"payload/python.zip", "payload/VC_redist.x64.exe", "requirements.lock", MODEL_RELATIVE.as_posix(),
                "tesseract/tesseract.exe", "tesseract/tessdata/eng.traineddata",
                "tesseract/tessdata/chi_sim.traineddata", "deployment.py"}
    if not required.issubset(data["files"]) or not any(n.startswith("wheelhouse/") for n in data["files"]):
        raise ValueError("Offline bundle is incomplete")
    actual = {p.relative_to(root).as_posix() for p in root.rglob("*")
              if p.is_file() and p.relative_to(root).as_posix() != "manifest.json"}
    if actual != set(data["files"]):
        raise ValueError("Offline bundle file list differs from its manifest")
    for name, expected in data["files"].items():
        if sha256(contained(root, name)) != expected:
            raise ValueError(f"SHA256 mismatch: {name}")
    if data["files"]["payload/python.zip"] != PYTHON_SHA256 or data["files"][MODEL_RELATIVE.as_posix()] != ASSETS["model.pt"][1]:
        raise ValueError("Runtime/model do not match the pinned release")
    return data


def download_bundle(output: Path, version: str, *, wheel: Path | None = None,
                    cache: Path | None = None, archive=True) -> Path:
    require_windows()
    cache = (cache or default_home() / "setup-cache").expanduser().resolve()
    cache.mkdir(parents=True, exist_ok=True)
    output = output.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    name = f"easy-uiauto-{version}-offline-win-x64-{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}"
    root = output / name
    root.mkdir()
    log = cache / "download.log"
    sources = {name: dict(url=url, sha256=digest) for name, (url, digest) in ASSETS.items()}
    print("[1/6] Downloading pinned runtime and OCR resources", flush=True)
    for name in ASSETS:
        # Prefer a verified existing detector cache without copying local images/settings.
        if name == "model.pt" and not (cache / name).exists():
            old = [default_home() / "models/omniparser/icon_detect_v3/model.pt",
                   Path.home() / ".codex/models/omniparser/icon_detect_v3/model.pt"]
            for candidate in old:
                if candidate.is_file() and sha256(candidate) == ASSETS[name][1]:
                    shutil.copy2(candidate, cache / name)
                    break
        download(ASSETS[name][0], cache / name, ASSETS[name][1])
    payload = root / "payload"
    payload.mkdir()
    shutil.copy2(cache / "python.zip", payload / "python.zip")
    shutil.copy2(cache / "VC_redist.x64.exe", payload / "VC_redist.x64.exe")
    with tempfile.TemporaryDirectory(prefix="prepare-", dir=cache) as scratch_text:
        scratch = Path(scratch_text)
        print("[2/6] Preparing the Windows Python runtime", flush=True)
        runtime = scratch / "python"
        extract_zip(cache / "python.zip", runtime)
        python = runtime / "python.exe"
        run([python, "-I", "-m", "ensurepip"], log=log)
        print("[3/6] Building the complete dependency wheelhouse (CPU PyTorch)", flush=True)
        wheels = root / "wheelhouse"
        wheels.mkdir()
        package = project_wheel(version, cache, wheel)
        run([python, "-I", "-m", "pip", "--isolated", "download", "--only-binary=:all:", "--no-deps",
             "--index-url", TORCH_INDEX, "--dest", wheels, "torch>=2.0"], log=log)
        torch = next(wheels.glob("torch-*.whl"))
        run([python, "-I", "-m", "pip", "--isolated", "wheel", "--wheel-dir", wheels, "--prefer-binary",
             f"{package}[full]", torch, "pip", "setuptools", "wheel"], log=log)
        (root / "requirements.lock").write_text(wheel_lock(wheels), encoding="utf-8")
        print("[4/6] Extracting portable Tesseract (no installer/network needed offline)", flush=True)
        seven = scratch / "7zip"
        run([cache / "7zr.exe", "x", cache / "7zip.exe", f"-o{seven}", "-y"], log=log)
        extracted = scratch / "tesseract"
        run([seven / "7z.exe", "x", cache / "tesseract-installer.exe", f"-o{extracted}", "-y"], log=log)
        # NSIS metadata/plugins are not part of the portable application.
        tesseract = root / "tesseract"
        shutil.copytree(extracted, tesseract, ignore=shutil.ignore_patterns("$*", "uninstall*", "*-uninstall.exe"))
        if not (tesseract / "tesseract.exe").is_file():
            raise RuntimeError("Could not extract the Tesseract runtime")
    for language in ("eng", "chi_sim"):
        (tesseract / "tessdata").mkdir(exist_ok=True)
        shutil.copy2(cache / f"{language}.traineddata", tesseract / "tessdata" / f"{language}.traineddata")
    model = root / MODEL_RELATIVE
    model.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(cache / "model.pt", model)
    print("[5/6] Writing offline installer and checksums", flush=True)
    shutil.copy2(Path(__file__), root / "deployment.py")
    (root / "install-offline.ps1").write_text(OFFLINE_PS, encoding="utf-8-sig")
    (root / "install-offline.cmd").write_text(
        '@echo off\r\npowershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-offline.ps1" %*\r\nexit /b %errorlevel%\r\n', encoding="ascii")
    (root / "README.txt").write_text(
        "easy-uiauto complete offline deployment / 完整离线部署\n\n"
        "Windows x64. Unzip the entire archive, then double-click install-offline.cmd.\n"
        "解压整个压缩包，双击 install-offline.cmd，无需联网或预先安装 Python。\n"
        "Custom location / 指定安装位置:\n"
        '  install-offline.cmd -Target "D:\\easy-uiauto"\n'
        "Default / 默认: %LOCALAPPDATA%\\easy-uiauto\\full\n"
        "Run the generated easy_uiauto_image.cmd, easy_uiauto_config.cmd or activate.cmd.\n"
        "安装完成后可运行上述图片标注、配置窗口或环境激活入口。\n"
        "Existing MCP clients are registered automatically; API settings stay in the configuration window.\n"
        "自动接入已安装的 MCP 客户端；远程 API 参数可在配置窗口中设置。\n",
        encoding="utf-8-sig")
    write_manifest(root, version, sources)
    verify_bundle(root)
    if archive:
        print("[6/6] Compressing the offline deployment ZIP", flush=True)
        path = Path(shutil.make_archive(str(root), "zip", root.parent, root.name))
        path.with_suffix(".zip.sha256").write_text(f"{sha256(path)}  {path.name}\n", encoding="ascii")
        print(f"Offline ZIP: {path}", flush=True)
    return root


SITECUSTOMIZE = '''# easy-uiauto managed runtime
import os
import sys
from pathlib import Path
_app = Path(sys.prefix).parent
os.environ["PATH"] = str(_app / "tesseract") + os.pathsep + os.environ.get("PATH", "")
os.environ["TESSDATA_PREFIX"] = str(_app / "tesseract" / "tessdata")
os.environ.setdefault("EASY_UIAUTO_MODEL_DIR", str(_app / "models"))
'''


def write_launchers(target: Path) -> None:
    modules = {"easy_uiauto": "easy_uiauto.mcp.server", "easy_uiauto_image": "easy_uiauto.image_cli",
               "easy_uiauto_ui": "easy_uiauto.mcp.ui_cli_command", "easy_uiauto_setup": "easy_uiauto.deployment",
               "easy_uiauto_config": "easy_uiauto.mcp.config_launcher"}
    for name, module in modules.items():
        (target / f"{name}.cmd").write_text(
            f'@echo off\r\n"%~dp0python\\python.exe" -I -X utf8 -m {module} %*\r\nexit /b %errorlevel%\r\n', encoding="ascii")
    (target / "activate.cmd").write_text(
        '@echo off\r\nset "PATH=%~dp0;%~dp0python;%~dp0python\\Scripts;%~dp0tesseract;%PATH%"\r\n', encoding="ascii")


def integrate(target: Path, client: str, log: Path) -> None:
    from importlib.util import module_from_spec, spec_from_file_location

    # Reuse the existing registry/PATH helper without importing the Windows GUI API
    # into the bootstrap interpreter.
    helper = target / "python/Lib/site-packages/easy_uiauto/mcp/diagnostics.py"
    spec = spec_from_file_location("deployment_diagnostics", helper)
    diagnostics = module_from_spec(spec)
    spec.loader.exec_module(diagnostics)
    diagnostics._add_to_user_path(target)
    python = target / "python/python.exe"
    run([python, "-m", "easy_uiauto.mcp.server", "--install-codex-skills"], log=log)
    clients = [x for x in ("codex", "claude") if shutil.which(x) or shutil.which(x + ".cmd")] if client == "auto" else [client]
    for name in clients:
        if name == "none":
            continue
        if name == "codex":
            command = (
                "from easy_uiauto.mcp.configuration import replace_codex_preserving_settings; "
                "replace_codex_preserving_settings()"
            )
        else:
            command = (
                "from easy_uiauto.mcp import configuration as c\n"
                "try: c.uninstall_claude_code()\n"
                "except RuntimeError: pass\n"
                "c.install_claude_code()"
            )
        run([python, "-I", "-c", command], log=log, timeout=120)


def ensure_vc_runtime(installer: Path, log: Path) -> None:
    try:
        for name in ("msvcp140.dll", "vcruntime140.dll", "vcruntime140_1.dll"):
            ctypes.WinDLL(name)
        return
    except OSError:
        pass
    print("Installing Microsoft Visual C++ Runtime; Windows may request elevation", flush=True)
    env = dict(os.environ, EASY_UIAUTO_VC_INSTALLER=str(installer))
    command = (
        "$p=Start-Process -FilePath $env:EASY_UIAUTO_VC_INSTALLER "
        "-ArgumentList '/install','/quiet','/norestart' -Verb RunAs -WindowStyle Hidden -Wait -PassThru; "
        "exit $p.ExitCode"
    )
    run(["powershell.exe", "-NoProfile", "-Command", command], env=env, log=log,
        timeout=600, accepted=(0, 1638, 3010))


def verify_runtime(target: Path) -> dict:
    python = target / "python/python.exe"
    log = target / "install.log"
    run([python, "-I", "-m", "pip", "check"], log=log, timeout=120)
    probe = r'''
import importlib, json, os, pathlib, tempfile
import sys
def deny_network(event, args):
    if event in ("socket.connect", "socket.getaddrinfo"):
        raise RuntimeError("Installation self-check must work without network access")
sys.addaudithook(deny_network)
from PIL import Image, ImageDraw, ImageFont
for name in ("tkinter", "pyautogui", "pynput", "uiautomation", "mcp", "cv2", "torch", "onnxruntime", "rapidocr"):
    importlib.import_module(name)
import tkinter
tkinter.Tcl().eval("info patchlevel")
import easy_uiauto, pytesseract
from easy_uiauto import _highlight_worker
from easy_uiauto import image_detection
from easy_uiauto.mcp.diagnostics import _ocr_languages
image_detection.check_runtime()
weights = image_detection.ensure_weights(image_detection.default_weights(), offline=True)
with tempfile.TemporaryDirectory() as scratch:
    image = Image.new("RGB", (1440, 900), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((20, 25, 200, 95), outline="black", width=1)
    draw.text((40, 45), "Easy 123", fill="black", font=ImageFont.truetype("arial.ttf", 20))
    languages, language_error = _ocr_languages(20)
    assert language_error is None, language_error
    assert {"eng", "chi_sim"}.issubset(languages), languages
    text = pytesseract.image_to_string(image, lang="eng+chi_sim", config="--psm 6")
    assert "123" in text, text
    sample = pathlib.Path(easy_uiauto.__file__).parent / "resources/IEEE9.png"
    with Image.open(sample) as source:
        example = source.convert("RGB")
    ocr = image_detection.run_ocr(example, threads=1, recover_small_text=False)
    assert "BUS_2" in ocr["texts"], ocr["texts"]
    detection = image_detection.run_yolo(example, weights, threads=2)
    assert detection["boxes"], "YOLO returned no graphics for the installation example"
gui = pathlib.Path(easy_uiauto.__file__).parent / "bin/easy-uiauto-config.exe"
assert gui.is_file() and gui.read_bytes()[:2] == b"MZ", "Configuration EXE missing"
print(json.dumps(dict(ok=True, version=easy_uiauto.__version__, python=__import__("sys").version.split()[0],
    tesseract_languages=sorted(languages), rapidocr=True, yolo=True, configuration_exe=True)))
'''
    output = run([python, "-I", "-c", probe], log=log, timeout=300)
    result = json.loads(output.strip().splitlines()[-1])
    (target / "verification.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    return result


def install_bundle(root: Path, target: Path, *, client="auto", no_integrate=False) -> dict:
    require_windows()
    root, target = root.expanduser().resolve(strict=True), target.expanduser().resolve()
    print("[1/5] Checking every offline resource", flush=True)
    manifest = verify_bundle(root)
    if target == root or target.is_relative_to(root) or root.is_relative_to(target):
        raise ValueError("Installation directory and offline bundle must be separate")
    marker = target / "deployment.json"
    if target.exists() and any(target.iterdir()) and not marker.is_file():
        raise ValueError("Choose an empty target or an existing easy-uiauto managed installation")
    if marker.is_file() and json.loads(marker.read_text(encoding="utf-8")).get("python") != PYTHON_VERSION:
        raise ValueError("Use a new target for a different managed Python version")
    target.mkdir(parents=True, exist_ok=True)
    marker.write_text(json.dumps(dict(python=PYTHON_VERSION, version=manifest["version"], status="installing")), encoding="utf-8")
    (target / "verification.json").unlink(missing_ok=True)
    log = target / "install.log"
    print("[2/5] Installing Python, portable Tesseract and cached models", flush=True)
    ensure_vc_runtime(root / "payload/VC_redist.x64.exe", log)
    runtime = target / "python"
    runtime_marker = runtime / ".easy-uiauto-runtime"
    if not runtime_marker.is_file() or runtime_marker.read_text(encoding="ascii") != PYTHON_SHA256:
        extract_zip(root / "payload/python.zip", runtime)
        runtime.mkdir(parents=True, exist_ok=True)
        runtime_marker.write_text(PYTHON_SHA256, encoding="ascii")
    for name in ("tesseract", "models"):
        shutil.copytree(root / name, target / name, dirs_exist_ok=True)
    site = runtime / "Lib/site-packages"
    site.mkdir(parents=True, exist_ok=True)
    (site / "sitecustomize.py").write_text(SITECUSTOMIZE, encoding="utf-8")
    python = runtime / "python.exe"
    run([python, "-I", "-m", "ensurepip"], log=log)
    print("[3/5] Installing all Python dependencies from local wheels", flush=True)
    env = dict(os.environ, PIP_NO_INDEX="1", PIP_DISABLE_PIP_VERSION_CHECK="1")
    run([python, "-I", "-m", "pip", "--isolated", "install", "--no-index", "--no-deps", "--no-compile",
         "--progress-bar", "off", "--no-warn-script-location",
         "--require-hashes", "--find-links", root / "wheelhouse", "-r", root / "requirements.lock"],
        env=env, log=log)
    write_launchers(target)
    print("[4/5] Testing imports, OCR, YOLO inference and configuration resources", flush=True)
    result = verify_runtime(target)
    print("[5/5] Configuring launchers and MCP clients", flush=True)
    if not no_integrate:
        integrate(target, client, log)
    marker.write_text(json.dumps(dict(python=PYTHON_VERSION, version=manifest["version"], status="ready")), encoding="utf-8")
    print(f"Complete installation verified: {target}", flush=True)
    return result


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    actions = parser.add_subparsers(dest="action", required=True)
    for name in ("install", "download"):
        action = actions.add_parser(name)
        action.add_argument("--version", help="Package version (defaults to this release)")
        action.add_argument("--wheel", type=Path, help="Build from a local release wheel")
        action.add_argument("--cache", type=Path)
        if name == "download":
            action.add_argument("output", type=Path, help="Directory for offline ZIP and SHA256")
    for name in ("install", "install-offline", "verify"):
        action = actions.choices.get(name) or actions.add_parser(name)
        action.add_argument("--target", type=Path, default=default_home() / "full")
        if name != "verify":
            action.add_argument("--client", choices=("auto", "codex", "claude", "none"), default="auto")
            action.add_argument("--no-integrate", action="store_true", help="Skip user PATH, skills and MCP registration")
        if name == "install-offline":
            action.add_argument("bundle", type=Path)
    args = parser.parse_args(argv)
    try:
        require_windows()
        if args.action == "verify":
            print(json.dumps(verify_runtime(args.target), indent=2))
        elif args.action == "install-offline":
            install_bundle(args.bundle, args.target, client=args.client, no_integrate=args.no_integrate)
        else:
            version = args.version or package_version()
            root = download_bundle(args.output if args.action == "download" else default_home() / "setup-cache/bundles",
                                   version, wheel=args.wheel, cache=args.cache, archive=args.action == "download")
            if args.action == "install":
                install_bundle(root, args.target, client=args.client, no_integrate=args.no_integrate)
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile, subprocess.SubprocessError) as error:
        print(f"Setup failed: {error}", file=sys.stderr, flush=True)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
