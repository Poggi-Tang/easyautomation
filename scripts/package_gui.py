"""Create the complete, offline GUI ZIP from a built directory distribution."""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
from pathlib import Path
import shutil
import zipfile

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("gui_package_deployment", ROOT / "src/easy_uiauto/deployment.py")
deployment = importlib.util.module_from_spec(spec)
spec.loader.exec_module(deployment)


def link_or_copy(source, destination):
    try:
        os.link(source, destination)
    except OSError:
        shutil.copy2(source, destination)
    return str(destination)


def validate_distribution(source: Path):
    required = ["easy-uiauto-library.exe", "_internal/python313.dll", "_internal/PySide6/Qt6Widgets.dll",
                "_internal/PySide6/Qt6Svg.dll", "_internal/PySide6/plugins/platforms/qwindows.dll",
                "_internal/image-runtime/python.exe", "_internal/image-runtime/python313.dll",
                "_internal/image-runtime/runtime-manifest.json", "THIRD_PARTY_NOTICES.txt",
                "third_party/qt/licenses/LGPL-3.0-only.txt", "gui-build.json"]
    for name in required:
        if not (source / name).is_file():
            raise ValueError(f"Incomplete GUI distribution: {name}")
    for name in ("torch", "onnxruntime", "rapidocr", "cv2", "numpy", "PIL"):
        path = source / "_internal/image-runtime/Lib/site-packages" / name / "__init__.py"
        if not path.is_file():
            raise ValueError(f"Missing inference dependency: {name}")
    build = json.loads((source / "gui-build.json").read_text(encoding="utf-8"))
    if build.get("frontend_only"):
        raise ValueError("A frontend-only build cannot be distributed as a complete GUI")
    if deployment.sha256(source / "easy-uiauto-library.exe") != build["exe_sha256"]:
        raise ValueError("GUI executable differs from its build manifest")
    for relative, digest in build["sources"].items():
        if deployment.sha256(ROOT / relative) != digest:
            raise ValueError(f"Source changed after the GUI build: {relative}; rebuild first")
    models = {
        "model.pt": source / "_internal/image-runtime/models/omniparser/icon_detect_v3/model.pt",
        **{name: source / "_internal/image-runtime/Lib/site-packages/rapidocr/models" / name
           for name in ("PP-OCRv6_det_small.onnx", "PP-OCRv6_rec_small.onnx")},
    }
    for name, path in models.items():
        if not path.is_file() or deployment.sha256(path) != deployment.ASSETS[name][1]:
            raise ValueError(f"Missing or invalid bundled model: {name}")
    return build


def package_directory(source: Path, output: Path, *, vc_installer: Path | None = None):
    source, output = source.resolve(strict=True), output.resolve()
    if output == source or output.is_relative_to(source) or source.is_relative_to(output):
        raise ValueError("Package output must be separate from the build directory")
    build = validate_distribution(source)
    output.mkdir(parents=True, exist_ok=True)
    target = output / "easy-uiauto-gui-win-x64"
    archive = output / (target.name + ".zip")
    if target.exists() or archive.exists():
        raise ValueError("Choose a new output directory; an existing package will not be overwritten")
    installer = vc_installer or deployment.default_home() / "setup-cache/VC_redist.x64.exe"
    url, digest = deployment.ASSETS["VC_redist.x64.exe"]
    if vc_installer:
        if deployment.sha256(installer) != digest:
            raise ValueError("Visual C++ installer checksum mismatch")
    else:
        deployment.download(url, installer, digest)
    target.mkdir()
    for name in ("_internal", "third_party"):
        shutil.copytree(source / name, target / name, copy_function=link_or_copy,
                        ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "*.log"))
    for name in ("easy-uiauto-library.exe", "gui-build.json", "THIRD_PARTY_NOTICES.txt"):
        link_or_copy(source / name, target / name)
    for path in source.glob("LICENSE*"):
        if path.is_file():
            link_or_copy(path, target / path.name)
    (target / "scripts").mkdir()
    (target / "prerequisites").mkdir()
    shutil.copy2(ROOT / "setup.bat", target / "setup.bat")
    shutil.copy2(ROOT / "scripts/gui_setup.ps1", target / "scripts/gui_setup.ps1")
    shutil.copy2(installer, target / "prerequisites/VC_redist.x64.exe")
    manifest = dict(format=1, version=build["version"], exe_sha256=build["exe_sha256"],
                    vc_redist_sha256=digest, models_included=True)
    (target / "gui-distribution.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    (target / "使用说明.txt").write_text(
        "easy-uiauto 知识库 GUI（Windows x64）\n\n"
        "1. 把 ZIP 完整解压到固定目录，不要在压缩包内直接运行。\n"
        "2. 双击 setup.bat，按 Windows 提示授权；将创建桌面快捷方式并打开 GUI。\n"
        "3. 以后双击桌面的 easy-uiauto GUI，或目录内的 EXE 即可。\n\n"
        "Python、界面依赖、图形检测和文字识别模型已随包附带；本地功能无需联网。\n"
        "初次使用可直接选择软件窗口、添加区域并学习。知识库默认保存到用户目录下的 easy_uiauto_vault。\n"
        "需要导入参考资料、生成说明时，再到“设置”填写 API。\n"
        "移动整个目录后重新运行 setup.bat 更新快捷方式；不要只复制 EXE。\n"
        "此入口不注册 MCP，也不改动 Codex、Claude 或用户 PATH。\n"
        "启动错误日志：%TEMP%\\easy_uiauto\\library-startup.log。\n",
        encoding="utf-8-sig")
    print(f"Compressing complete GUI: {target}", flush=True)
    temporary = archive.with_suffix(".zip.partial")
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED, compresslevel=3) as stream:
            for path in sorted(target.rglob("*")):
                if path.is_file():
                    stream.write(path, path.relative_to(output).as_posix())
        temporary.replace(archive)
    finally:
        temporary.unlink(missing_ok=True)
    (archive.with_suffix(".zip.sha256")).write_text(
        f"{deployment.sha256(archive)}  {archive.name}\n", encoding="ascii")
    print(f"Ready to share: {archive}", flush=True)
    return archive


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--vc-installer", type=Path)
    args = parser.parse_args()
    package_directory(args.source, args.output_dir, vc_installer=args.vc_installer)
