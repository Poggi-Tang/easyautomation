"""Shared, persistent model file choices for the GUI and isolated inference runtime."""
from __future__ import annotations

import importlib.util
import json
import os
import re
from pathlib import Path
import sys
from uuid import uuid4

MODEL_FILES = {"yolo": "model.pt", "det": "PP-OCRv6_det_small.onnx",
               "rec": "PP-OCRv6_rec_small.onnx"}
CONFIG_ENV = "EASY_UIAUTO_MODEL_PATHS_FILE"
DEVICE_ENV = "EASY_UIAUTO_YOLO_DEVICE"


def configuration_path():
    if os.environ.get(CONFIG_ENV):
        return Path(os.environ[CONFIG_ENV]).expanduser().resolve()
    cache = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".cache"))
    return cache / "easy-uiauto/model-paths.json"


def _validate_paths(paths):
    if not isinstance(paths, dict) or set(paths) - MODEL_FILES.keys():
        raise ValueError("模型路径配置格式不正确")
    result = {}
    for key, value in paths.items():
        if not isinstance(value, str):
            raise ValueError("模型路径必须为文本")
        value = value.strip()
        if not value:
            continue
        path = Path(os.path.expandvars(value)).expanduser()
        if not path.is_absolute():
            raise ValueError("模型路径须为绝对文件路径")
        if path.suffix.lower() != Path(MODEL_FILES[key]).suffix:
            raise ValueError(f"{key} 模型须选择 {Path(MODEL_FILES[key]).suffix} 文件")
        path = path.resolve()
        if path.is_dir():
            raise ValueError("请选择模型文件，不能使用文件夹路径")
        if any(path == Path(existing) for existing in result.values()):
            raise ValueError("不同模型不能使用同一个文件路径")
        result[key] = str(path)
    return result


def _load_configuration():
    file = configuration_path()
    try:
        data = json.loads(file.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {"version": 1, "paths": {}}
    except (OSError, ValueError) as error:
        raise ValueError(f"无法读取模型路径配置：{file}（{error}）") from error
    if not isinstance(data, dict) or data.get("version") != 1:
        raise ValueError(f"模型路径配置格式不正确：{file}")
    _validate_paths(data.get("paths"))
    validate_yolo_device(data.get("yolo_device", "auto"))
    return data


def load_model_paths():
    return _validate_paths(_load_configuration()["paths"])


def validate_yolo_device(value):
    if not isinstance(value, str):
        raise ValueError("YOLO device must be auto, cpu, cuda or cuda:N")
    value = value.strip().lower()
    if value not in {"auto", "cpu", "cuda"} and not re.fullmatch(r"cuda:[0-9]+", value):
        raise ValueError("YOLO device must be auto, cpu, cuda or cuda:N")
    return value


def configured_yolo_device():
    return validate_yolo_device(_load_configuration().get("yolo_device", "auto"))


def yolo_device(value=""):
    """Per-call selection wins over environment, then GUI preferences, then auto."""
    return validate_yolo_device(value or os.environ.get(DEVICE_ENV) or configured_yolo_device())


def save_yolo_device(value):
    value = validate_yolo_device(value)
    data = _load_configuration()
    data["yolo_device"] = value
    _save_configuration(data)
    return value


def save_model_paths(paths):
    paths = _validate_paths(paths)
    data = _load_configuration()
    data["paths"] = paths
    _save_configuration(data)
    return paths


def _save_configuration(data):
    file = configuration_path()
    file.parent.mkdir(parents=True, exist_ok=True)
    temporary = file.with_name(f"{file.name}.{uuid4().hex}.tmp")
    try:
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                             encoding="utf-8")
        temporary.replace(file)
    finally:
        temporary.unlink(missing_ok=True)


def default_yolo_path():
    configured = os.environ.get("EASY_UIAUTO_MODEL_DIR")
    if configured:
        return Path(configured).expanduser() / "omniparser/icon_detect_v3/model.pt"
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "image-runtime/models/omniparser/icon_detect_v3/model.pt"
    cache = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".cache")) / "easy-uiauto"
    preferred = cache / "models/omniparser/icon_detect_v3/model.pt"
    previous = Path.home() / ".codex/models/omniparser/icon_detect_v3/model.pt"
    return preferred if preferred.is_file() or not previous.is_file() else previous


def rapidocr_directory():
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "image-runtime/Lib/site-packages/rapidocr"
    spec = importlib.util.find_spec("rapidocr")
    if spec and spec.origin:
        return Path(spec.origin).parent
    return Path.home() / ".codex/models/runtimes/rapidocr-trial/rapidocr"


def yolo_path():
    configured = load_model_paths().get("yolo")
    return Path(configured) if configured else default_yolo_path()


def ocr_model_paths(package_directory=None):
    configured = load_model_paths()
    directory = Path(package_directory) if package_directory is not None else rapidocr_directory()
    return {key: Path(configured[key]) if key in configured else directory / "models" / MODEL_FILES[key]
            for key in ("det", "rec")}
