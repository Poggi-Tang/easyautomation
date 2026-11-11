"""Build a directory-distributed Qt manager with its isolated UIA workers."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
from pathlib import Path
import subprocess
import sys
import tomllib


def copy_runtime_file(source, destination):
    # A hard link is a normal distributable file, with no dependency on its old
    # pathname. Avoid recopying hundreds of MB on the same build volume.
    if Path(destination).exists():
        if os.path.samefile(source,destination):return str(destination)
        Path(destination).unlink()
    try:os.link(source,destination)
    except OSError:shutil.copy2(source,destination)
    return str(destination)


def bundle_image_runtime(python: Path, target: Path, root: Path, weights: Path):
    """Keep ML outside the GUI archive and copy only declared distributions.

    The source must be a redistributable CPython install, not a frozen executable.
    Package RECORDs include native DLLs, OCR models and license files.
    """
    probe = '''import json,sys
from importlib.metadata import distribution, PackageNotFoundError
from packaging.requirements import Requirement
pending=['torch','rapidocr','onnxruntime','Pillow','numpy','uiautomation','pyautogui','pynput','pyperclip','build']
try:distribution('opencv-python-headless');pending.append('opencv-python-headless')
except PackageNotFoundError:pending.append('opencv-python')
seen=set();files=[]
while pending:
    name=pending.pop();dist=distribution(name);key=dist.metadata['Name'].lower().replace('_','-')
    if key in seen:continue
    seen.add(key)
    for value in dist.requires or []:
        req=Requirement(value)
        if not req.marker or req.marker.evaluate({'extra':''}):pending.append(req.name)
    for f in dist.files or []:
        if '..' in f.parts or '__pycache__' in f.parts:continue
        # CPU inference needs the runtime DLLs, not C++ extension build headers,
        # CMake metadata, import libraries or debug symbols.
        if f.parts[:2] in {('torch','include'),('torch','share')} or f.suffix.lower() in {'.lib','.pdb'}:continue
        files.append(str(dist.locate_file(f)))
print(json.dumps(dict(base=sys.base_prefix,site=str(dist.locate_file('')),files=sorted(set(files)),packages=sorted(seen))))
'''
    manifest = json.loads(subprocess.check_output([str(python),"-I","-X","utf8","-c",probe],text=True,encoding="utf-8"))
    source, site = Path(manifest["base"]), Path(manifest["site"])
    if not (source/"python.exe").is_file():raise ValueError("Image runtime requires a complete Windows CPython directory")
    target.mkdir(parents=True,exist_ok=True)
    for pattern in ("python.exe","pythonw.exe","python3.dll","python313.dll","vcruntime*.dll","LICENSE*.txt"):
        for path in source.glob(pattern):copy_runtime_file(path,target/path.name)
    shutil.copytree(source/"DLLs",target/"DLLs",dirs_exist_ok=True,ignore=shutil.ignore_patterns("__pycache__"),copy_function=copy_runtime_file)
    shutil.copytree(source/"Lib",target/"Lib",dirs_exist_ok=True,
        ignore=shutil.ignore_patterns("site-packages","__pycache__","test","tests","idlelib","turtledemo","ensurepip"),copy_function=copy_runtime_file)
    destination = target/"Lib/site-packages"
    for name in manifest["files"]:
        path=Path(name).resolve()
        if not path.is_file() or not path.is_relative_to(site.resolve()):continue
        out=destination/path.relative_to(site.resolve());out.parent.mkdir(parents=True,exist_ok=True)
        copy_runtime_file(path,out)
    shutil.copytree(root/"src/easy_uiauto",destination/"easy_uiauto",dirs_exist_ok=True,
                    ignore=shutil.ignore_patterns("__pycache__"))
    model=target/"models/omniparser/icon_detect_v3/model.pt";model.parent.mkdir(parents=True,exist_ok=True)
    copy_runtime_file(weights,model)
    # RapidOCR wheels can omit downloaded weights from RECORD. Include the exact
    # two models explicitly, even when packaging on a clean build machine.
    from easy_uiauto import deployment
    ocr_models = destination / "rapidocr/models"
    for name in ("PP-OCRv6_det_small.onnx", "PP-OCRv6_rec_small.onnx"):
        cached = site / "rapidocr/models" / name
        model = ocr_models / name
        url, digest = deployment.ASSETS[name]
        if cached.is_file() and deployment.sha256(cached) == digest:
            model.parent.mkdir(parents=True, exist_ok=True)
            copy_runtime_file(cached, model)
        else:
            deployment.download(url, model, digest)
    (target/"runtime-manifest.json").write_text(json.dumps({"packages":manifest["packages"],"detector":"microsoft/OmniParser-v2.0/icon_detect_v3"},indent=2),encoding="utf-8")


def write_notices(root: Path, target: Path):
    shutil.copytree(root/"third_party/qt",target/"third_party/qt",dirs_exist_ok=True)
    for license_file in root.glob("LICENSE*"):
        if license_file.is_file():shutil.copy2(license_file,target/license_file.name)
    (target/"THIRD_PARTY_NOTICES.txt").write_text(
        "easy-uiauto knowledge manager\n\nQt for Python (PySide6 Essentials, Shiboken6)\n"
        "Qt components retain their original licenses and copyright.\n"
        "License texts: third_party/qt/licenses, plus the distribution metadata under _internal.\n"
        "Qt source and licensing: https://www.qt.io/qt-for-python ; https://code.qt.io/\n"
        "This directory distribution keeps the Qt shared libraries separately replaceable.\n\n"
        "Local image runtime: _internal/image-runtime\n"
        "CPython license: _internal/image-runtime/LICENSE.txt\n"
        "Runtime package list: _internal/image-runtime/runtime-manifest.json\n"
        "Package licenses and notices: _internal/image-runtime/Lib/site-packages/*.dist-info\n"
        "YOLO weights: microsoft/OmniParser-v2.0/icon_detect_v3\n"
        "Model source: https://huggingface.co/microsoft/OmniParser-v2.0\n",encoding="utf-8")


def main():
    root=Path(__file__).resolve().parents[1]
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir",type=Path,default=root/"dist")
    parser.add_argument("--work-dir",type=Path,default=root/"build/library-exe")
    parser.add_argument("--incremental",action="store_true",help="Reuse PyInstaller's validated build cache")
    parser.add_argument("--image-python",type=Path,help="Redistributable Windows Python with image dependencies installed")
    parser.add_argument("--image-weights",type=Path,help="Pinned OmniParser model.pt (uses the standard model cache by default)")
    parser.add_argument("--frontend-only",action="store_true",help="Build just the GUI; deploy its EXE alongside an existing compatible _internal directory")
    parser.add_argument("--package-dir",type=Path,help="Also write a complete GUI ZIP with setup.bat to a new directory")
    args=parser.parse_args()
    if os.name!="nt":parser.error("Build the Windows app on Windows")
    if sys.version_info[:2] != (3,13):parser.error("Build the GUI with Python 3.13")
    if args.frontend_only and args.package_dir:parser.error("A shareable ZIP requires the complete image runtime")
    if not args.frontend_only:
        if not args.image_python:parser.error("--image-python is required unless --frontend-only is specified")
        from easy_uiauto import image_detection
        weights=image_detection.ensure_weights(args.image_weights or image_detection.default_weights(),offline=True)
    work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    # Absolute import lets PyInstaller discover the package from this launcher.
    entry=work/"launch_library.py"
    launch="from easy_uiauto.library_launcher import main\nif __name__ == '__main__':\n    main()\n"
    if not entry.is_file() or entry.read_text(encoding="utf-8")!=launch:entry.write_text(launch,encoding="utf-8")
    resources=root/"src/easy_uiauto/resources"
    command=[sys.executable,"-m","PyInstaller","--noconfirm","--onedir","--windowed","--noupx","--uac-admin",
        "--name","easy-uiauto-library","--icon",str(resources/"easy-uiauto.ico"),
        "--paths",str(root/"src"),"--distpath",str(args.output_dir.resolve()),
        "--workpath",str(work/"work"),"--specpath",str(work),
        "--add-data",f"{resources}{os.pathsep}easy_uiauto/resources",
        "--collect-all","uiautomation","--collect-all","pypdfium2_raw","--collect-submodules","comtypes.gen",
        "--copy-metadata","pypdf","--copy-metadata","pypdfium2",
        "--copy-metadata","PySide6-Essentials","--copy-metadata","shiboken6"]
    if not args.incremental:command.append("--clean")
    # ML runs in image-runtime only when learning, keeping the GUI startup lazy.
    for module in ("torch","torchvision","torchaudio","rapidocr","onnxruntime","onnx",
                   "tensorflow","scipy","pandas","matplotlib","IPython","pytest","PyQt5","PyQt6",
                   "mcp","easy_uiauto.mcp.server","easy_uiauto.knowledge_gui",
                   "easy_uiauto.usage_gui","easy_uiauto.workflow_gui"):
        command.extend(["--exclude-module",module])
    # Resolve DLLs against the build runtime and Windows, not unrelated tools
    # inherited from PATH (e.g. Poppler's ICU exports differ from Windows ICU).
    environment=os.environ.copy();windows=Path(os.environ.get("SystemRoot",r"C:\Windows"))
    environment["PATH"]=os.pathsep.join(str(p) for p in (Path(sys.executable).parent,Path(sys.base_prefix),
        Path(sys.base_prefix)/"DLLs",windows/"System32",windows,windows/"System32/Wbem"))
    environment["PYTHONPATH"]=str(root/"src")
    environment["PYTHONHASHSEED"]="0"
    for name in ("QT_PLUGIN_PATH","QML2_IMPORT_PATH","QML_IMPORT_PATH"):environment.pop(name,None)
    command.append(str(entry));subprocess.run(command,check=True,cwd=root,env=environment)
    target=args.output_dir.resolve()/"easy-uiauto-library"
    if not args.frontend_only:
        print("Bundling the separate local image runtime",flush=True)
        bundle_image_runtime(args.image_python,target/"_internal/image-runtime",root,weights)
    write_notices(root,target)
    def digest(path):
        with path.open('rb') as stream:return hashlib.file_digest(stream,'sha256').hexdigest()
    sources={path.relative_to(root).as_posix():digest(path)
             for path in (root/'src/easy_uiauto').rglob('*')
             if path.is_file() and path.suffix in {'.py','.svg','.ico','.png','.json'} and '__pycache__' not in path.parts}
    (target/'gui-build.json').write_text(json.dumps(dict(
        version=tomllib.loads((root/'pyproject.toml').read_text(encoding='utf-8'))['project']['version'],
        frontend_only=args.frontend_only,exe_sha256=digest(target/'easy-uiauto-library.exe'),sources=sources),indent=2),encoding='utf-8')
    if args.package_dir:
        from package_gui import package_directory
        package_directory(target,args.package_dir)
    print(target)


if __name__=="__main__":main()
