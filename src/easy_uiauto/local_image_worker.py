"""Lazy image inference; frozen managers use a separate bundled Python runtime."""
from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time


def _ocr(source):
    from PIL import Image
    from .image_detection import run_ocr
    with Image.open(source) as image:
        result=run_ocr(image.convert("RGB"),recover_small_text=False)
    return {"text":"\n".join(result["texts"])}


def read_text(source, *, progress):
    """OCR scanned reference pages with the existing isolated local runtime."""
    return analyze(source,"",progress=progress,_operation="ocr")["text"]


def analyze(source, output, *, progress, _operation="analyze"):
    if not getattr(sys, "frozen", False):
        if _operation=="ocr":return _ocr(source)
        from .image_cli import analyze_local_image
        return analyze_local_image(source, output, auto_crop=False, profile="ui", progress=progress)
    runtime = Path(sys._MEIPASS) / "image-runtime"
    python = runtime / "python.exe"
    if not python.is_file():
        raise RuntimeError("缺少本地标注运行环境，请使用包含 image-runtime 的完整 GUI 目录")
    with tempfile.TemporaryDirectory(prefix="easy-uiauto-inference-") as temporary:
        job = Path(temporary)
        request = job / "request.json"
        request.write_text(json.dumps(dict(source=str(source), output=str(output),operation=_operation)), encoding="utf-8")
        environment = os.environ.copy()
        environment.setdefault("EASY_UIAUTO_MODEL_DIR", str(runtime / "models"))
        from .image_model_paths import CONFIG_ENV, configuration_path
        environment[CONFIG_ENV] = str(configuration_path())
        # A frozen parent's DLL directory must not shadow the child's runtime DLLs.
        import ctypes
        ctypes.windll.kernel32.SetDllDirectoryW(None)
        try:
            with (job/"stderr.log").open("wb") as log:
                process = subprocess.Popen([str(python), "-I", "-X", "utf8", "-m", __name__, str(request)],
                    env=environment, stdin=subprocess.DEVNULL, stdout=log, stderr=log,
                    creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0))
                try:
                    while process.poll() is None:
                        state = job / "progress.json"
                        try: message = json.loads(state.read_text(encoding="utf-8"))["message"]
                        except (OSError, ValueError, KeyError): message = "正在启动本地 YOLO＋RapidOCR"
                        progress(message)  # Also checks cancellation while inference is running.
                        time.sleep(.25)
                    if process.returncode:
                        raise RuntimeError("本地标注失败："+(job/"stderr.log").read_text(encoding="utf-8",errors="replace")[-4000:])
                    return json.loads((job/"result.json").read_text(encoding="utf-8"))
                finally:
                    if process.poll() is None:
                        process.terminate(); process.wait(timeout=5)
        finally:
            ctypes.windll.kernel32.SetDllDirectoryW(sys._MEIPASS)


def main():
    request = Path(sys.argv[1])
    data = json.loads(request.read_text(encoding="utf-8"))
    from .image_cli import analyze_local_image
    from .library_tasks import _write
    def progress(message):
        _write(request.parent / "progress.json",dict(message=message))
    result = (_ocr(data["source"]) if data.get("operation")=="ocr" else
              analyze_local_image(data["source"],data["output"],auto_crop=False,profile="ui",progress=progress))
    (request.parent/"result.json").write_text(json.dumps(result,ensure_ascii=False),encoding="utf-8")


if __name__ == "__main__": main()
