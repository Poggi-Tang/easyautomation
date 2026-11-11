"""Inspect and repair the exact local models used by image inference, without ML imports."""
from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import time
from urllib import request
from uuid import uuid4

from .image_detection import MODEL_SHA256, MODEL_URL, default_weights, file_sha256
from .image_model_paths import load_model_paths, ocr_model_paths, rapidocr_directory, save_model_paths

OCR_MODELS = (
    ("det", "OCR 文字检测", "PP-OCRv6_det_small.onnx",
     "090f04abcd9d9a7498bc4ebf677e4cb9bdce1fe4197ddb7e529f1ef44e1ff94f"),
    ("rec", "OCR 文字识别", "PP-OCRv6_rec_small.onnx",
     "6f327246b50388f3c176ae304bd95767ea6dc0c9ae92153ef8cbe210b3c14884"),
)


def model_specs():
    models = [dict(id="yolo", name="YOLO 图形检测", path=str(default_weights().resolve()),
                   url=MODEL_URL, sha256=MODEL_SHA256)]
    paths = ocr_model_paths()
    for kind, name, filename, digest in OCR_MODELS:
        models.append(dict(id=kind, name=name, path=str(paths[kind].resolve()),
                           url=f"https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.10.0/onnx/PP-OCRv6/{kind}/{filename}",
                           sha256=digest))
    configured = load_model_paths()
    return [{**model, "configured_path": configured.get(model["id"], "")} for model in models]


def runtime_issues():
    required = ("torch", "onnxruntime", "cv2", "numpy", "PIL")
    if getattr(sys, "frozen", False):
        runtime = Path(sys._MEIPASS) / "image-runtime"
        issues = [] if (runtime / "python.exe").is_file() else ["缺少本地标注 Python 运行环境"]
        for name in required:
            if not (runtime / "Lib/site-packages" / name / "__init__.py").is_file():
                issues.append(f"缺少运行依赖：{name}")
    else:
        issues = [f"缺少运行依赖：{name}" for name in required
                  if importlib.util.find_spec(name) is None]
    if not (rapidocr_directory() / "__init__.py").is_file():
        issues.append("缺少 RapidOCR 运行依赖")
    return issues


def inspect_models(*, progress=lambda _message: None):
    models = []
    for spec in model_specs():
        progress("正在检查 " + spec["name"])
        path = Path(spec["path"])
        state, detail = "missing", "未下载"
        try:
            if path.is_file():
                state = "ready" if file_sha256(path) == spec["sha256"] else "invalid"
                detail = "已就绪（校验通过）" if state == "ready" else "文件损坏或版本不匹配"
        except OSError as error:
            state, detail = "unreadable", "无法读取：" + str(error)
        models.append({**spec, "state": state, "detail": detail})
    issues = runtime_issues()
    ready = sum(model["state"] == "ready" for model in models)
    message = (f"标注模型已就绪（{ready}/{len(models)}）" if ready == len(models)
               else f"标注模型待补齐（{ready}/{len(models)} 已就绪）")
    if issues:
        message += "；运行环境不完整"
    return dict(ok=ready == len(models) and not issues, models=models,
                runtime_issues=issues, message=message)


def download_model(spec, *, progress):
    """Publish only verified files; cancellation and failed transfers preserve old models."""
    path = Path(spec["path"])
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f"{path.name}.{uuid4().hex}.partial")
    # LibraryTask owns this one path, so a forced worker stop can clean it too.
    register = getattr(progress, "register_download", None)
    if register:
        register(temporary)
    try:
        progress("正在连接下载源：" + spec["name"])
        req = request.Request(spec["url"], headers={"User-Agent": "easy-uiauto-image"})
        with request.urlopen(req, timeout=20) as response, temporary.open("xb") as stream:
            total = int(response.headers.get("Content-Length", 0))
            downloaded, last_report = 0, 0
            while True:
                # The worker's progress callback also observes cancellation.
                amount = f"{downloaded / 1_000_000:.1f} MB"
                if total:
                    amount += f" / {total / 1_000_000:.1f} MB（{downloaded / total:.0%}）"
                if time.monotonic() - last_report >= .2:
                    progress(f"正在下载 {spec['name']}：{amount}")
                    last_report = time.monotonic()
                block = response.read(256 * 1024)
                if not block:
                    break
                stream.write(block)
                downloaded += len(block)
        progress("正在校验 " + spec["name"])
        if file_sha256(temporary) != spec["sha256"]:
            raise ValueError(f"{spec['name']} 下载文件校验失败，请重试")
        progress("正在保存 " + spec["name"])
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def download_missing_models(*, progress):
    report = inspect_models(progress=progress)
    for spec in report["models"]:
        if spec["state"] != "ready":
            try:
                download_model(spec, progress=progress)
            except InterruptedError:
                raise
            except Exception as error:
                raise RuntimeError(f"{spec['name']} 下载失败：{error}。可点击“一键下载 / 修复”重试。") from error
    return inspect_models(progress=progress)
