"""Local OmniParser detection, OCR and original-pixel diagram geometry.

The engines are optional and imported only when used. No vision API, desktop
inspection, reference annotations or manually supplied geometry is involved.
"""

from __future__ import annotations

import hashlib
import importlib
import math
import os
import sys
import time
from importlib.metadata import version as package_version
from pathlib import Path
from urllib import request
from uuid import uuid4

MODEL_ID = "microsoft/OmniParser-v2.0/icon_detect_v3"
MODEL_URL = "https://huggingface.co/microsoft/OmniParser-v2.0/resolve/main/icon_detect_v3/model.pt"
MODEL_SHA256 = "11c6cbb77f22569fab22d86c76407a83ec81ab89dbfe28279854822d6e3fb00c"
MODEL_SIZE = 1280
DEFAULT_CONFIDENCE = 0.1


def file_sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def default_weights() -> Path:
    cache = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".cache")) / "easy-uiauto"
    preferred = cache / "models" / "omniparser" / "icon_detect_v3" / "model.pt"
    # Reuse the model already downloaded by the local experiment.
    previous = Path.home() / ".codex" / "models" / "omniparser" / "icon_detect_v3" / "model.pt"
    return preferred if preferred.is_file() or not previous.is_file() else previous


def ensure_weights(path: Path, *, offline=False, progress=print) -> Path:
    """Verify the pinned official model before loading; download atomically once."""
    path = Path(path).expanduser().resolve()
    if path.is_file():
        if file_sha256(path) != MODEL_SHA256:
            raise ValueError(f"YOLO model SHA256 mismatch: {path}")
        return path
    if offline:
        raise FileNotFoundError(f"YOLO weights not cached: {path}. Run once without --offline.")
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f"{path.name}.{uuid4().hex}.partial")
    try:
        progress("Downloading official YOLO model (~281 MB); later runs reuse the cache.")
        downloaded, last_report = 0, time.monotonic()
        req = request.Request(MODEL_URL, headers={"User-Agent": "easy-uiauto-image"})
        with request.urlopen(req, timeout=60) as response, temporary.open("xb") as stream:
            for block in iter(lambda: response.read(1024 * 1024), b""):
                stream.write(block)
                downloaded += len(block)
                if time.monotonic() - last_report >= 5:
                    progress(f"Model download: {downloaded / 1_000_000:.1f} MB")
                    last_report = time.monotonic()
        if file_sha256(temporary) != MODEL_SHA256:
            raise ValueError("Downloaded YOLO model failed SHA256 verification")
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
    return path


def _rapidocr_module(runtime_dir=""):
    if runtime_dir:
        runtime = Path(runtime_dir).expanduser().resolve(strict=True)
        sys.path.insert(0, str(runtime))
        importlib.invalidate_caches()
    elif importlib.util.find_spec("rapidocr") is None:
        previous = Path.home() / ".codex" / "models" / "runtimes" / "rapidocr-trial"
        if (previous / "rapidocr" / "__init__.py").is_file():
            sys.path.insert(0, str(previous))
            importlib.invalidate_caches()
    try:
        return importlib.import_module("rapidocr")
    except ImportError as error:
        raise RuntimeError(
            'OCR dependencies unavailable. Install with: python -m pip install ".[image]"'
        ) from error


def check_runtime(runtime_dir=""):
    """Fail before YOLO inference if the optional OCR/runtime dependencies are absent."""
    try:
        for name in ("torch", "onnxruntime", "cv2", "numpy"):
            importlib.import_module(name)
        rapidocr = _rapidocr_module(runtime_dir)
    except ImportError as error:
        raise RuntimeError(
            'Local image dependencies unavailable. Install: python -m pip install ".[image]"'
        ) from error
    models = Path(rapidocr.__file__).parent / "models"
    for name in ("PP-OCRv6_det_small.onnx", "PP-OCRv6_rec_small.onnx"):
        if not (models / name).is_file():
            raise FileNotFoundError(
                f"Bundled OCR model missing: {models / name}. Install rapidocr==3.10.0."
            )
    return rapidocr


def nms(boxes, scores, threshold=0.45):
    """Single-class NMS, avoiding a torchvision runtime dependency."""
    import numpy as np

    boxes, scores = np.asarray(boxes), np.asarray(scores)
    if not len(boxes):
        return []
    order, keep = np.argsort(-scores, kind="stable"), []
    areas = np.maximum(0, boxes[:, 2] - boxes[:, 0]) * np.maximum(
        0, boxes[:, 3] - boxes[:, 1])
    while len(order):
        index, remaining = int(order[0]), order[1:]
        keep.append(index)
        if not len(remaining):
            break
        sizes = np.maximum(0, np.minimum(boxes[index, 2:], boxes[remaining, 2:])
                           - np.maximum(boxes[index, :2], boxes[remaining, :2]))
        overlap = sizes[:, 0] * sizes[:, 1]
        union = np.maximum(1e-9, areas[index] + areas[remaining] - overlap)
        order = remaining[overlap / union <= threshold]
    return keep


def _yolo_input_size(image_size, *, rectangular=False, content_bottom=None):
    """Keep the detector scale while removing unused stride-aligned padding."""
    width, height = image_size
    scale = min(MODEL_SIZE / width, MODEL_SIZE / height)
    resized = (max(1, round(width * scale)), max(1, round(height * scale)))
    if not rectangular:
        return resized, (MODEL_SIZE, MODEL_SIZE)
    used_height = resized[1]
    if content_bottom is not None:
        used_height = min(used_height, max(1, math.ceil(content_bottom * resized[1] / height)))
    return resized, (math.ceil(resized[0] / 32) * 32, math.ceil(used_height / 32) * 32)


def _predict_yolo(image, model, confidence, *, rectangular=False, content_bottom=None):
    import numpy as np
    import torch
    from PIL import Image

    started = time.perf_counter()
    # Track rounded resize ratios separately to preserve original pixel coordinates.
    size, canvas_size = _yolo_input_size(image.size, rectangular=rectangular,
                                        content_bottom=content_bottom)
    canvas = Image.new("RGB", canvas_size, (114, 114, 114))
    canvas.paste(image.resize(size, Image.Resampling.BILINEAR), (0, 0))
    tensor = torch.from_numpy(np.array(canvas)).permute(2, 0, 1).float()[None] / 255
    forward_start = time.perf_counter()
    with torch.inference_mode():
        outputs = model(tensor)
    forward_seconds = time.perf_counter() - forward_start
    boxes, scores = [], []
    for i, stride in enumerate((8, 16, 32)):
        classes = outputs[2 * i].sigmoid()[0, 0]
        left, top, right, bottom = outputs[2 * i + 1][0]
        gy, gx = torch.meshgrid(torch.arange(classes.shape[0], dtype=torch.float32),
                               torch.arange(classes.shape[1], dtype=torch.float32), indexing="ij")
        cx, cy = gx + 0.5, gy + 0.5
        decoded = torch.stack(((cx - left) * stride, (cy - top) * stride,
                               (cx + right) * stride, (cy + bottom) * stride), -1)
        boxes.append(decoded.reshape(-1, 4))
        scores.append(classes.reshape(-1))
    boxes, scores = torch.cat(boxes), torch.cat(scores)
    selected = scores > confidence
    boxes, scores = boxes[selected].numpy(), scores[selected].numpy()
    keep = nms(boxes, scores)
    boxes, scores = boxes[keep], scores[keep]
    boxes[:, 0::2] /= size[0] / image.width
    boxes[:, 1::2] /= size[1] / image.height
    boxes[:, 0::2] = np.clip(boxes[:, 0::2], 0, image.width)
    boxes[:, 1::2] = np.clip(boxes[:, 1::2], 0, image.height)
    valid = ((boxes[:, 2] - boxes[:, 0]) > 2) & ((boxes[:, 3] - boxes[:, 1]) > 2)
    return dict(forward_seconds=forward_seconds, tensor_shape=list(tensor.shape),
                resized_image_size=list(size),
                elapsed_seconds=round(time.perf_counter() - started, 3),
                boxes=boxes[valid].tolist(), scores=scores[valid].tolist())


def run_yolo(image, weights, *, threads=4, confidence=DEFAULT_CONFIDENCE,
             recover_details=False, progress=None, rectangular=False):
    """Keep the whole-image pass and optionally inspect native-scale UI patches."""
    import torch

    torch.set_num_threads(threads)
    started=time.perf_counter()
    model=torch.jit.load(str(weights),map_location='cpu').eval()
    loaded=time.perf_counter()
    whole=_predict_yolo(image,model,confidence,rectangular=rectangular)
    boxes,scores=whole['boxes'],whole['scores']
    forward=whole['forward_seconds']
    detail=dict(enabled=False,reason='Native-scale UI detail recovery disabled',added_count=0)
    passes=1
    input_shapes=[whole.get('tensor_shape')]
    if recover_details:
        from . import image_details

        planning=time.perf_counter()
        detail=image_details.plan(image,MODEL_SIZE)
        detail['planning_seconds']=round(time.perf_counter()-planning,4)
        if detail['enabled']:
            if progress:
                progress(f"YOLO detail recovery: {len(detail['regions'])} native-scale content regions in one extra pass")
            atlas=image_details.render_atlas(image,detail)
            bottom=max(r['atlas_rect'][3] for r in detail['regions'])+detail.get(
                'parameters',{}).get('gutter',16)
            extra=_predict_yolo(atlas,model,confidence,rectangular=rectangular,
                                content_bottom=bottom)
            boxes,scores=image_details.merge_predictions(boxes,scores,extra['boxes'],extra['scores'],detail)
            detail['inference_seconds']=extra['elapsed_seconds']
            detail['tensor_shape']=extra.get('tensor_shape')
            input_shapes.append(extra.get('tensor_shape'))
            forward+=extra['forward_seconds']
            passes+=1
    return dict(model_id=MODEL_ID,architecture='YOLOv9-E',model_sha256=MODEL_SHA256,
                device='cpu',threads=threads,input_size=MODEL_SIZE,
                confidence_threshold=confidence,nms_iou_threshold=0.45,
                load_seconds=round(loaded-started,3),forward_seconds=round(forward,3),
                elapsed_seconds=round(time.perf_counter()-started,3),
                rectangular_inference=rectangular,input_shapes=input_shapes,
                inference_passes=passes,whole_image_count=len(whole['boxes']),
                detail_recovery=detail,boxes=boxes,scores=scores)


def run_detectors(image, weights, *, threads=4, confidence=DEFAULT_CONFIDENCE,
                  runtime_dir='', recover_details=False, recover_small_text=True,
                  rectangular=False, parallel=True, progress=None):
    """Share one CPU thread budget; deliver all progress on the caller thread."""
    from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
    from queue import Empty, SimpleQueue

    parallel = bool(parallel and threads > 1)
    yolo_threads, ocr_threads = (threads - 1, 1) if parallel else (threads, threads)
    messages = SimpleQueue()
    report = messages.put if parallel else progress
    yolo_options = dict(threads=yolo_threads, confidence=confidence,
                        recover_details=recover_details, rectangular=rectangular, progress=report)
    ocr_options = dict(threads=ocr_threads, runtime_dir=runtime_dir,
                       recover_small_text=recover_small_text, progress=report)
    if progress:
        progress(f"Inference schedule: {'parallel' if parallel else 'sequential'}; "
                 f"YOLO {yolo_threads} threads, OCR {ocr_threads} threads")
    started = time.perf_counter()
    if parallel:
        results = {}
        with ThreadPoolExecutor(max_workers=2, thread_name_prefix='image-inference') as pool:
            tasks = {pool.submit(run_yolo, image, weights, **yolo_options): 'yolo',
                     pool.submit(run_ocr, image, **ocr_options): 'ocr'}
            pending = set(tasks)
            while pending:
                done, pending = wait(pending, timeout=0.1, return_when=FIRST_COMPLETED)
                while True:
                    try:
                        message = messages.get_nowait()
                    except Empty:
                        break
                    if progress:
                        progress(message)
                for task in done:
                    results[tasks[task]] = task.result()
                    if progress:
                        progress(f"{tasks[task].upper()} inference completed")
        yolo, ocr = results['yolo'], results['ocr']
    else:
        yolo = run_yolo(image, weights, **yolo_options)
        ocr = run_ocr(image, **ocr_options)
    return yolo, ocr, dict(parallel=parallel, thread_budget=threads,
                           yolo_threads=yolo_threads, ocr_threads=ocr_threads,
                           rectangular_inference=rectangular,
                           wall_seconds=round(time.perf_counter()-started,3))


class _CaptureTextDetection:
    """Retain the existing detector's probability map without another forward pass."""

    def __init__(self, postprocess):
        self.postprocess = postprocess
        self.predictions = None
        self.shape = None

    def __call__(self, predictions, shape):
        self.predictions, self.shape = predictions, shape
        return self.postprocess(predictions, shape)


def _small_text_candidates(image, boxes, scores, full_boxes, frames):
    """Limit relaxed candidates to isolated, previously uncovered ink inside frames."""
    import numpy as np

    candidates, rejected = [], []
    rgb = np.asarray(image)
    existing = [_rect((min(p[0] for p in q), min(p[1] for p in q),
                       max(p[0] for p in q), max(p[1] for p in q)), image.size)
                for q in full_boxes]
    for quad, detection_score in zip(boxes, scores, strict=True):
        box = _rect((min(p[0] for p in quad), min(p[1] for p in quad),
                     max(p[0] for p in quad), max(p[1] for p in quad)), image.size)
        if box is None:
            continue
        if any(old and _overlap(box, old) / min(_area(box), _area(old)) >= 0.3
               for old in existing):
            continue
        owners = [frame["rect"] for frame in frames
                  if frame["rect"]["left"] + 4 <= box["left"]
                  and frame["rect"]["top"] + 4 <= box["top"]
                  and frame["rect"]["right"] - 4 >= box["right"]
                  and frame["rect"]["bottom"] - 4 >= box["bottom"]]
        if not owners or box["height"] > 32 or box["width"] > 128:
            continue
        owner = min(owners, key=_area)
        patch = dict(left=max(owner["left"] + 4, box["left"] - 3),
                     top=max(owner["top"] + 4, box["top"] - 3),
                     right=min(owner["right"] - 4, box["right"] + 3),
                     bottom=min(owner["bottom"] - 4, box["bottom"] + 3))
        pixels = rgb[patch["top"]:patch["bottom"], patch["left"]:patch["right"]]
        mask = ((pixels.max(2) < 170)
                & (pixels.max(2).astype("int16") - pixels.min(2) < 80))
        yy, xx = np.where(mask)
        reason = ""
        if len(xx) < 3:
            reason = "Insufficient neutral dark text pixels"
        elif mask[0].any() or mask[-1].any() or mask[:, 0].any() or mask[:, -1].any():
            reason = "Ink touches crop boundary; may be a connected wire or frame stroke"
        else:
            ink_height, ink_width = int(yy.max() - yy.min() + 1), int(xx.max() - xx.min() + 1)
            if ink_height < 3 or ink_width / ink_height > 12 or ink_height / ink_width > 10:
                reason = "Stroke dimensions do not support a small text region"
        evidence = dict(detector_quad=np.asarray(quad).tolist(),
                        detector_score=float(detection_score),
                        frame_rect=owner, patch_rect=patch)
        if reason:
            rejected.append({**evidence, "reason": reason})
            continue
        ink = _rect((patch["left"] + int(xx.min()), patch["top"] + int(yy.min()),
                     patch["left"] + int(xx.max()) + 1, patch["top"] + int(yy.max()) + 1),
                    image.size)
        if any(_overlap(ink, old["ink_rect"]) / min(_area(ink), _area(old["ink_rect"])) >= 0.5
               for old in candidates):
            continue
        candidates.append({**evidence, "ink_rect": ink})
    return candidates, rejected


def _recognize_small_candidates(engine, image, candidates):
    """Require agreement between the original crop and an enlarged version."""
    import cv2
    import numpy as np

    bgr = np.asarray(image)[:, :, ::-1].copy()
    patches, recovered, rejected = [], [], []
    for candidate in candidates:
        patch = candidate["patch_rect"]
        crop = bgr[patch["top"]:patch["bottom"], patch["left"]:patch["right"]]
        scale = max(2, min(4, 32 / candidate["ink_rect"]["height"]))
        patches.extend((crop, cv2.resize(crop, None, fx=scale, fy=scale,
                                        interpolation=cv2.INTER_CUBIC)))
    if not patches:
        return recovered, rejected
    recognition = engine.recognize_txt(patches)
    for i, candidate in enumerate(candidates):
        text, enlarged_text = (str(recognition.txts[j]).strip() for j in (2 * i, 2 * i + 1))
        score, enlarged_score = (float(recognition.scores[j]) for j in (2 * i, 2 * i + 1))
        evidence = {**candidate, "text": text, "recognition_score": score,
                    "enlarged_text": enlarged_text,
                    "enlarged_recognition_score": enlarged_score}
        if (not text or text != enlarged_text or max(score, enlarged_score) < 0.95
                or min(score, enlarged_score) < 0.9 or len(text) > 12):
            rejected.append({**evidence, "reason": "Recognition confidence/agreement"})
            continue
        r = candidate["ink_rect"]
        quad = [[r["left"], r["top"]], [r["right"], r["top"]],
                [r["right"], r["bottom"]], [r["left"], r["bottom"]]]
        recovered.append(dict(quad=quad, text=text, score=min(score, enlarged_score),
                              evidence=evidence))
    return recovered, rejected


def _recover_small_text(engine, image, captured, full_boxes, *, progress=None):
    """Recover high-confidence small text from relaxed detections inside real frames."""
    import numpy as np
    from rapidocr.utils.process_img import apply_vertical_padding, map_boxes_to_original

    from .mcp import diagram_geometry

    started = time.perf_counter()
    report = dict(enabled=True, full_box_threshold=0.5, local_box_threshold=0.2,
                  minimum_best_recognition_score=0.95, minimum_each_recognition_score=0.9,
                  candidate_count=0, accepted=[], rejected=[], added_count=0)
    if captured is None or captured.predictions is None:
        report["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        return [], report
    frames = [c for c in diagram_geometry.shapes(image) if c["kind"] == "frame"]
    postprocess = captured.postprocess
    original_threshold = postprocess.box_thresh
    report["full_box_threshold"] = original_threshold
    try:
        # Only use these extra proposals inside supported frames. The whole-image
        # detections and their original threshold/results are left untouched.
        postprocess.box_thresh = report["local_box_threshold"]
        boxes, scores = postprocess(captured.predictions, captured.shape)
    finally:
        postprocess.box_thresh = original_threshold
    if not len(boxes) or not frames:
        report["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        return [], report
    bgr = np.asarray(image)[:, :, ::-1].copy()
    processed, operations = engine.preprocess_img(bgr)
    if engine.cfg.Global.use_vertical_padding:
        _padded, operations = apply_vertical_padding(
            processed, operations, engine.width_height_ratio, engine.min_height)
    boxes = map_boxes_to_original(boxes.copy(), operations, image.height, image.width)
    candidates, rejected = _small_text_candidates(
        image, boxes, scores, full_boxes, frames)
    report["candidate_count"], report["rejected"] = len(candidates), rejected
    if progress:
        progress(f"Small-text recovery: checking {len(candidates)} isolated regions inside frames")
    recovered, rejected = _recognize_small_candidates(engine, image, candidates)
    report["accepted"] = [region["evidence"] for region in recovered]
    report["rejected"].extend(rejected)
    report["added_count"] = len(recovered)
    report["elapsed_seconds"] = round(time.perf_counter() - started, 3)
    return recovered, report


def run_ocr(image, *, threads=4, runtime_dir="", recover_small_text=True, progress=None):
    import numpy as np

    rapidocr = _rapidocr_module(runtime_dir)
    models = Path(rapidocr.__file__).parent / "models"
    started = time.perf_counter()
    engine = rapidocr.RapidOCR(params={
        "Global.max_side_len": max(4096, max(image.size)), "Global.use_cls": False,
        "Global.log_level": "warning", "Det.limit_side_len": 1280, "Det.limit_type": "min",
        "Det.model_path": str(models / "PP-OCRv6_det_small.onnx"),
        "Rec.model_path": str(models / "PP-OCRv6_rec_small.onnx"),
        "EngineConfig.onnxruntime.intra_op_num_threads": threads,
        "EngineConfig.onnxruntime.inter_op_num_threads": 1,
    })
    loaded = time.perf_counter()
    captured = None
    if recover_small_text:
        detector = engine._load_det_model()
        captured = _CaptureTextDetection(detector.postprocess_op)
        detector.postprocess_op = captured
    # RapidOCR accepts ndarray images in OpenCV BGR order.
    try:
        result = engine(np.asarray(image)[:, :, ::-1].copy(), use_cls=False)
    finally:
        if captured is not None:
            detector.postprocess_op = captured.postprocess
    boxes = [] if result.boxes is None else np.asarray(result.boxes).tolist()
    texts = [] if result.txts is None else list(result.txts)
    scores = [] if result.scores is None else [float(s) for s in result.scores]
    base_count = len(boxes)
    recovery_report = dict(enabled=False, added_count=0, elapsed_seconds=0,
                           accepted=[], rejected=[])
    recovered = []
    if recover_small_text:
        recovered, recovery_report = _recover_small_text(engine, image, captured, boxes,
                                                        progress=progress)
    for region in recovered:
        boxes.append(region["quad"])
        texts.append(region["text"])
        scores.append(region["score"])
    return dict(engine="RapidOCR PP-OCRv6 small ONNX CPU", version=package_version("rapidocr"),
                load_seconds=round(loaded - started, 3),
                elapsed_seconds=round(time.perf_counter() - started, 3),
                boxes=boxes, texts=texts, scores=scores, base_text_count=base_count,
                region_sources=(["full-image"] * base_count
                                + ["small-text-recovery"] * len(recovered)),
                region_details=[{} for _ in range(base_count)] + [r["evidence"] for r in recovered],
                small_text_recovery=recovery_report)


def _rect(bounds, image_size):
    from .mcp.component_geometry import rect

    left = max(0, math.floor(bounds[0]))
    top = max(0, math.floor(bounds[1]))
    right = min(image_size[0], math.ceil(bounds[2]))
    bottom = min(image_size[1], math.ceil(bounds[3]))
    if right <= left or bottom <= top:
        return None
    return rect(left, top, right - left, bottom - top)


def _area(box):
    return max(1, box["width"] * box["height"])


def _overlap(a, b):
    return max(0, min(a["right"], b["right"]) - max(a["left"], b["left"])) * max(
        0, min(a["bottom"], b["bottom"]) - max(a["top"], b["top"]))


def _caption_distance(graphic, text):
    fraction = max(0, min(graphic["right"], text["right"]) - max(
        graphic["left"], text["left"])) / max(1, min(graphic["width"], text["width"]))
    gap = text["top"] - graphic["bottom"]
    return max(0, gap) if fraction >= 0.5 and -2 <= gap <= max(8, 1.8 * text["height"]) else None


def combine_detections(image, yolo, ocr, *, profile="block-diagram"):
    """Fuse independent geometry, color and YOLO evidence; keep text as controls."""
    from . import image_geometry
    from .mcp import component_geometry, image_analysis

    started = time.perf_counter()
    if profile not in {"block-diagram", "ui"}:
        raise ValueError(f"Unknown image profile: {profile}")
    texts, rejected_text = [], []
    for index, (points, text, score) in enumerate(zip(
            ocr["boxes"], ocr["texts"], ocr["scores"], strict=True), 1):
        xs, ys = [p[0] for p in points], [p[1] for p in points]
        box = _rect((min(xs), min(ys), max(xs), max(ys)), image.size)
        if box is None:
            rejected_text.append(dict(id=index, reason="OCR box outside the image", quad=points))
            continue
        sources = ocr.get("region_sources")
        details = ocr.get("region_details")
        texts.append(dict(rect=box, text=text, score=score, quad=points, raw_id=index,
                          recovery=bool(sources and sources[index - 1] == "small-text-recovery"),
                          recovery_evidence=details[index - 1] if details else {}))
    proposals = [(index, _rect(b, image.size), score) for index, (b, score) in enumerate(
        zip(yolo["boxes"], yolo["scores"], strict=True), 1)]
    proposals = [(i, box, score) for i, box, score in proposals if box is not None]
    ui_metadata, recovered_icons, ui_report = {}, [], {"enabled": False}
    if profile == "ui":
        from .ui_image_fusion import prepare_ui

        proposals, texts, ui_metadata, recovered_icons, ui_report = prepare_ui(image, proposals, texts)
        rejected_text.extend(ui_report["rejected_ocr"])
    geometry = ([] if profile == "ui" else component_geometry.candidates(image, block_diagram=True))
    text_rects = [t["rect"] for t in texts]
    colors = [] if profile == "ui" else image_geometry.color_candidates(image, text_rects)
    for color in colors:
        duplicate = next((c for c in geometry if component_geometry.iou(
            color["rect"], c["rect"]) >= 0.65), None)
        if duplicate:
            # Keep the enclosing neutral border when it already fits the body.
            duplicate["color_evidence"] = color
            a, b = duplicate["rect"], color["rect"]
            duplicate["rect"] = _rect((min(a["left"], b["left"]), min(a["top"], b["top"]),
                                        max(a["right"], b["right"]), max(a["bottom"], b["bottom"])),
                                       image.size)
        else:
            geometry.append(color)
    if profile != "ui":
        geometry.extend(image_geometry.solid_bars(
            image, [c["rect"] for c in geometry] + text_rects))
    for index, candidate in enumerate(geometry, 1):
        candidate["candidate_id"] = index
    busbars = [c["rect"] for c in geometry if c["kind"] == "busbar"]
    elements, accepted = [], set()
    for candidate in sorted(geometry, key=lambda c: (c["rect"]["top"], c["rect"]["left"])):
        box = candidate["rect"]
        if candidate["kind"] == "ground" and any(
                _overlap(box, bus) / _area(box) >= 0.5 for bus in busbars):
            candidate.update(fusion_status="rejected", rejection_reason="Detail of a busbar")
            continue
        supporting = [(i, score) for i, p, score in proposals
                      if component_geometry.iou(box, p) >= 0.25
                      or (_overlap(box, p) / _area(box) >= 0.9 and _area(p) <= _area(box) * 12)]
        inside = [t["raw_id"] for t in texts if _overlap(box, t["rect"]) / _area(t["rect"]) >= 0.8]
        captions = [t["raw_id"] for t in texts if _caption_distance(box, t["rect"]) is not None]
        primitive = candidate["kind"] in {"busbar", "ground", "terminal", "solid-bar"}
        color_supported = candidate["kind"] == "color-body" or "color_evidence" in candidate
        # A wire loop can resemble a pointed tag and happen to sit above a label.
        # A tag needs its own internal text or YOLO evidence, not just that caption.
        if candidate["kind"] == "tag" and not (supporting or inside):
            candidate.update(fusion_status="rejected",
                             rejection_reason="Unconfirmed pointed outline")
            continue
        if not (supporting or inside or captions or primitive or color_supported):
            candidate.update(fusion_status="rejected",
                             rejection_reason="No independent control evidence")
            continue
        candidate["fusion_status"] = "accepted"
        accepted.update(i for i, _score in supporting)
        score = max((score for _i, score in supporting), default=None)
        provider = "color-geometry" if color_supported else "pixel-geometry"
        elements.append(dict(
            id=len(elements) + 1, entity_type="graphic-control", geometry_kind=candidate["kind"],
            semantic_name=f"{candidate['kind']} control", observed_value="", rect=box,
            source=f"yolo-and-{provider}" if supporting else f"local-{provider}",
            detection_confidence=score, confidence=score if score is not None else 0,
            confidence_kind="yolo-detection" if supporting else "unscored-local-geometry",
            evidence=[candidate["evidence"]], yolo_proposal_ids=[i for i, _score in supporting],
            color_evidence=candidate.get("color_evidence", candidate if color_supported else None),
            ocr_support_ids=inside, caption_support_ids=captions,
            status="observed", uia_verified=False, executable=False, suggested_actions=[]))
    unmatched = list(ui_report.get("rejected_yolo", []))
    for index, box, score in proposals:
        if index in accepted:
            continue
        if profile != "ui":
            # Do not emit a second box for a partial prediction of an existing body.
            duplicate = next((g for g in elements if _overlap(g["rect"], box) / min(
                _area(g["rect"]), _area(box)) >= (
                    0.5 if _area(g["rect"]) >= _area(box) * 2 else 0.8)
                and _area(box) <= _area(g["rect"]) * 4), None)
            if duplicate:
                unmatched.append(dict(id=index, rect=box, confidence=score,
                                      reason="Partial or enclosing duplicate of a fitted graphic",
                                      duplicate_of_graphic_id=duplicate["id"]))
                continue
        text_duplicate = any(component_geometry.iou(box, t["rect"]) >= 0.5
                             or (_overlap(box, t["rect"]) / _area(box) >= 0.8
                                 and _area(box) <= _area(t["rect"]) * 1.8) for t in texts)
        frame = ui_metadata.get(index, {}).get("compact_frame")
        icon = ui_metadata.get(index, {}).get("compact_icon")
        if text_duplicate and not (frame or icon):
            unmatched.append(dict(id=index, rect=box, confidence=score,
                                  reason="Handled as an independent text control"))
            continue
        evidence = (ui_metadata[index]["pixel_evidence"] if profile == "ui" else
                    image_geometry.proposal_evidence(image, box, text_rects))
        if profile != "ui" and not evidence["accepted"]:
            unmatched.append(dict(id=index, rect=box, confidence=score, **evidence))
            continue
        accepted.add(index)
        fitted = frame or icon
        local_fit=ui_metadata.get(index,{}).get('local_geometry')
        frame_fields = ({"detected_rect": box, "corrected_rect": fitted, "geometry_padding": 1}
                        if fitted else {})
        elements.append(dict(
            id=len(elements) + 1, entity_type="graphic-control", geometry_kind=(
                local_fit['geometry_kind'] if local_fit else
                "compact-frame" if frame else "compact-icon" if icon else "yolo-region"),
            semantic_name="Graphic control candidate", observed_value="", rect=fitted or box,
            source="yolo", detection_confidence=score, confidence=score,
            confidence_kind="yolo-detection", yolo_proposal_ids=[index, *ui_report.get("merged_yolo_ids", {}).get(index, [])],
            icon_row_evidence=ui_metadata.get(index, {}).get("icon_row"),
            geometry_evidence=local_fit,
            pixel_evidence=evidence, ambiguity="Compact border observed; control semantics need review" if frame else
            "Compact icon pixels fitted; control semantics need review" if icon else
            "YOLO region has no fitted component outline",
            status="observed", uia_verified=False, executable=False, suggested_actions=[], **frame_fields))
    for recovered in recovered_icons:
        elements.append(dict(id=len(elements)+1, entity_type="graphic-control",
                             geometry_kind="repeated-icon", semantic_name="Repeated graphic candidate",
                             observed_value="", rect=recovered["rect"], source="local-template-recovery",
                             yolo_proposal_ids=[], confidence=recovered["match_score"],
                             confidence_kind="template-correlation", detection_confidence=None,
                             template_evidence=recovered,
                             ambiguity="Repeated pixels observed; control semantics need review",
                             status="observed", uia_verified=False, executable=False, suggested_actions=[]))
    for addition in ui_report.get("geometry_additions", []):
        elements.append(dict(id=len(elements)+1, entity_type="graphic-control",
                             geometry_kind=addition['geometry_kind'], semantic_name="Local UI graphic candidate",
                             observed_value="", rect=addition['rect'], corrected_rect=addition['rect'],
                             source="local-ui-geometry", yolo_proposal_ids=addition['supporting_yolo_ids'],
                             confidence=0, confidence_kind="unscored-local-geometry", detection_confidence=None,
                             geometry_evidence=addition, geometry_padding=1,
                             ambiguity="Pixel shape observed; control identity needs review",
                             status="observed", uia_verified=False, executable=False, suggested_actions=[]))
    if profile == "ui":
        elements.sort(key=lambda e: (e['rect']['top'], e['rect']['left']))
        for index, element in enumerate(elements, 1):
            element['id'] = index
    graphics = list(elements)
    for t in texts:
        owners = [g for g in graphics if _overlap(g["rect"], t["rect"]) / _area(t["rect"]) >= 0.65]
        relation = "text-of"
        if owners:
            owner = min(owners, key=lambda g: _area(g["rect"]))
        else:
            nearby = [(_caption_distance(g["rect"], t["rect"]), g) for g in graphics]
            nearby = [(distance, g) for distance, g in nearby if distance is not None]
            owner = min(nearby, key=lambda p: (p[0], _area(p[1]["rect"])))[1] if nearby else None
            relation = "label-of"
        symbol_overlap = owner and owner["geometry_kind"] in {"busbar", "ground", "terminal"}
        ambiguity = "OCR may have read a graphic stroke as text" if symbol_overlap else (
            "Low OCR recognition score; review literal content" if t["score"] < 0.9 else "")
        item = dict(id=len(elements) + 1, entity_type="text-label", semantic_name="Text control",
                    observed_value=t["text"], rect=t["rect"],
                    source="rapidocr-small-text-recovery" if t["recovery"] else "rapidocr-pp-ocrv6",
                    raw_ocr_id=t["raw_id"], text_quadrilateral=t["quad"], confidence=t["score"],
                    confidence_kind="ocr-recognition", recognition_confidence=t["score"],
                    detection_confidence=t["recovery_evidence"].get("detector_score"),
                    ambiguity=ambiguity,
                    status="observed", uia_verified=False, executable=False, suggested_actions=[])
        if owner:
            item["related_element_id"] = owner["id"]
            item["relationships"] = [dict(type=relation, target_element_id=owner["id"],
                                          evidence="Spatial containment or adjacent caption")]
        if t["recovery"]:
            item["ocr_recovery"] = t["recovery_evidence"]
        if t.get("layout_correction"):
            item["ocr_layout_correction"] = t["layout_correction"]
        elements.append(item)
    return dict(
        profile=profile, elements=image_analysis.complete_element_bounds(elements, *image.size, 1),
        ui_fusion=ui_report,
        template_recovered_graphics=len(recovered_icons),
        blank_rejected_graphics=sum(not r.get("accepted", True) for r in unmatched),
        element_count=len(elements), graphic_count=len(graphics), text_count=len(texts),
        yolo_supported_graphics=sum(bool(e["yolo_proposal_ids"]) for e in graphics),
        geometry_recovered_graphics=sum(not e["yolo_proposal_ids"] for e in graphics),
        yolo_raw_count=len(yolo["boxes"]), geometry_candidates=geometry,
        color_candidate_count=len(colors),
        color_supported_graphics=sum(bool(e.get("color_evidence")) for e in graphics),
        unmatched_yolo_graphics=sum(e["geometry_kind"] == "yolo-region" for e in graphics),
        fusion_seconds=round(time.perf_counter() - started, 3),
        ocr_base_text_count=ocr.get("base_text_count", len(texts)),
        ocr_recovery_count=ocr.get("small_text_recovery", {}).get("added_count", 0),
        ocr_recovery_seconds=ocr.get("small_text_recovery", {}).get("elapsed_seconds", 0),
        rejected_ocr_regions=rejected_text,
        needs_review_ids=[e["id"] for e in elements if e.get("ambiguity")],
        unassigned_yolo_proposals=unmatched)
