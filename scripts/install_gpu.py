"""Install an optional CUDA runtime alongside CPU packages; activate only after validation."""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import sysconfig
import tempfile
from urllib.request import urlopen
from uuid import uuid4


def run(arguments, **kwargs):
    return subprocess.run([sys.executable, "-I", "-X", "utf8", *arguments], check=True, **kwargs)


def pip_command(cache):
    if importlib.util.find_spec("pip"):
        return ["-m", "pip"]
    # The portable GUI intentionally omits pip. Run a checksum-verified pip wheel
    # directly from its ZIP, without replacing the bundled Python installation.
    with urlopen("https://pypi.org/pypi/pip/json", timeout=30) as response:
        metadata = json.load(response)
    wheel = next(item for item in metadata["urls"] if item["filename"].endswith("-py3-none-any.whl"))
    if not wheel["url"].startswith("https://files.pythonhosted.org/"):
        raise RuntimeError("Unexpected pip download host")
    destination = cache / Path(wheel["filename"]).name
    with urlopen(wheel["url"], timeout=60) as response, destination.open("wb") as output:
        shutil.copyfileobj(response, output)
    if hashlib.sha256(destination.read_bytes()).hexdigest() != wheel["digests"]["sha256"]:
        raise RuntimeError("pip checksum mismatch")
    return [str(destination / "pip")]


def verify_stage(stage):
    # Run the real, pinned YOLO model, not just torch.cuda.is_available(). The
    # child uses staged packages; the existing runtime is unchanged on failure.
    code = """
import json, sys
sys.path.insert(0, sys.argv[1])
from PIL import Image
from easy_uiauto.image_detection import default_weights, ensure_weights, run_yolo
result = run_yolo(Image.new('RGB', (128, 128), 'white'),
                  ensure_weights(default_weights(), offline=True), device='cuda', threads=2)
print(json.dumps({key: result[key] for key in ('device', 'gpu_name', 'cuda_version')}))
"""
    environment = os.environ.copy()
    runtime = Path(sys.executable).parent
    for models in (runtime / "models", runtime.parent / "models"):
        if (models / "omniparser/icon_detect_v3/model.pt").is_file():
            environment.setdefault("EASY_UIAUTO_MODEL_DIR", str(models))
            break
    run(["-c", code, str(stage)], env=environment)


def install(cuda_index="cu128"):
    if sys.platform != "win32":
        raise RuntimeError("GPU setup supports Windows x64 only")
    if not re.fullmatch(r"cu[0-9]{3}", cuda_index):
        raise ValueError("CUDA index must look like cu128 or cu130")
    site = Path(sysconfig.get_path("purelib")).resolve()
    site.mkdir(parents=True, exist_ok=True)
    marker = site / "easy_uiauto_gpu.pth"
    stage = site / ("easy_uiauto_gpu_" + uuid4().hex)
    stage.mkdir()
    active = False
    try:
        with tempfile.TemporaryDirectory(prefix="easy-uiauto-gpu-download-") as temporary:
            cache = Path(temporary)
            pip = pip_command(cache)
            wheels = cache / "wheels"
            print("Downloading CUDA PyTorch from the official index. This may take several GB.", flush=True)
            run([*pip, "--isolated", "download", "--only-binary=:all:", "--no-deps",
                 "--index-url", "https://download.pytorch.org/whl/" + cuda_index,
                 "--dest", str(wheels), "torch"])
            wheel, = wheels.glob("torch-*.whl")
            # Install that exact CUDA wheel, preventing the dependency resolver
            # from choosing a newer CPU torch on another index.
            run([*pip, "--isolated", "install", "--only-binary=:all:", "--ignore-installed",
                 "--target", str(stage), "--index-url", "https://pypi.org/simple", str(wheel)])
        print("Verifying CUDA and a real YOLO inference before activation...", flush=True)
        verify_stage(stage)
        temporary_marker = marker.with_suffix(".tmp")
        try:
            # Resolve relative to this site directory so portable installations
            # still work after moving the complete GUI folder.
            temporary_marker.write_text(
                "import os, sys; sys.path.insert(0, os.path.join(sitedir, " + repr(stage.name) + "))\n",
                encoding="ascii")
            os.replace(temporary_marker, marker)
            active = True
        finally:
            temporary_marker.unlink(missing_ok=True)
        print("GPU runtime ready. Restart the GUI/MCP and select Auto or NVIDIA GPU.", flush=True)
    finally:
        if not active:
            # This is the exact directory created by this invocation, never a
            # path supplied by the user or taken from an existing marker.
            shutil.rmtree(stage)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cuda-index", default="cu128", help="Official PyTorch CUDA index (default: cu128)")
    parser.add_argument("--disable", action="store_true", help="Use the original CPU runtime on the next launch")
    args = parser.parse_args(argv)
    if args.disable:
        (Path(sysconfig.get_path("purelib")) / "easy_uiauto_gpu.pth").unlink(missing_ok=True)
        print("GPU runtime disabled. Restart the GUI/MCP to use the original CPU runtime.")
    else:
        install(args.cuda_index)


if __name__ == "__main__":
    main()
