"""Command-line export of local graphic and text image annotations."""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import io
import json
import math
import sys
import time
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from . import image_detection


def _write_json(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_image(path):
    from PIL import Image, ImageOps

    path = Path(path).expanduser().resolve(strict=True)
    data = path.read_bytes()
    with Image.open(io.BytesIO(data)) as original:
        if getattr(original, "n_frames", 1) != 1:
            raise ValueError("Supply one image, not an animation or multipage file")
        orientation = original.getexif().get(274, 1)
        oriented = ImageOps.exif_transpose(original).convert("RGBA")
        image = Image.alpha_composite(Image.new("RGBA", oriented.size, "white"), oriented)
    return path, image.convert("RGB"), hashlib.sha256(data).hexdigest(), orientation


def _write_csv(path, elements):
    columns = ("id", "annotation_label", "entity_type", "observed_value", "semantic_name",
               "left", "top", "right", "bottom", "center_x", "center_y", "source",
               "confidence_kind", "confidence", "related_element_id", "ambiguity")
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for element in elements:
            writer.writerow({**element, **element["rect"], "center_x": element["center"]["x"],
                             "center_y": element["center"]["y"]})


def _write_viewer(directory, result):
    cards = []
    for title, name in (("文字与图形合并标注", "combined"), ("文字标注", "text"),
                        ("图形控件标注", "controls")):
        cards.append(f'<section><h2>{title}</h2><a href="{name}.png">'
                     f'<img src="{name}.png" alt="{title}" loading="lazy"></a></section>')
    resources = "".join(f'<li><a href="{html.escape(name, quote=True)}">'
                        f'{html.escape(name)}</a></li>'
                        for name in result["resources"] if name != "index.html")
    document = f"""<!doctype html>
<html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>图片标注结果</title><style>
body{{font:16px system-ui,sans-serif;margin:24px;color:#243047;background:#f5f7fa}}
main{{max-width:1500px;margin:auto}}section{{background:white;padding:16px;margin:20px 0}}
h1,h2{{font-weight:600}}h2{{font-size:18px}}img{{display:block;width:100%;height:auto}}
.graphic{{color:#1565c0}}.text{{color:#c65d00}}a{{color:#1565c0}}ul{{line-height:1.9}}
</style><main><h1>图片标注结果</h1>
<p>{html.escape(Path(result['image_path']).name)}
 · {result['image_width']} × {result['image_height']}
 · 图形 {result['graphic_count']} · 文字 {result['text_count']}</p>
<p><span class="graphic">蓝色 G：图形控件</span>　<span class="text">橙色 T：文字控件</span>
 · 标签位于各自框外侧。点击图片可查看原尺寸。</p>
{''.join(cards)}<section><h2>资源文件</h2><ul>{resources}</ul></section></main></html>"""
    (directory / "index.html").write_text(document, encoding="utf-8")


def analyze_local_image(image_path, output_dir, *, weights=None, profile="block-diagram",
                        threads=4, confidence=0.05, offline=False, ocr_runtime="",
                        small_text_recovery=True, progress=print):
    """Save a self-contained run folder, preserving previous runs and input pixels."""
    from .mcp import diagram_geometry, image_analysis

    started = time.perf_counter()
    if not 1 <= threads <= 64:
        raise ValueError("threads must be between 1 and 64")
    if not math.isfinite(confidence) or not 0 < confidence < 1:
        raise ValueError("confidence must be between 0 and 1")
    if profile not in {"block-diagram", "circuit", "ui"}:
        raise ValueError(f"Unknown image profile: {profile}")
    source, image, source_hash, orientation = _load_image(image_path)
    base = Path(output_dir).expanduser().resolve()
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    # Keep run folder names short even when input basenames approach Windows limits.
    directory = base / f"{source.stem[:60]}-{stamp}-{uuid4().hex[:6]}"
    directory.mkdir(parents=True, exist_ok=False)
    log_path = directory / "run.log"

    def report(message):
        line = f"[{time.perf_counter() - started:6.2f}s] {message}"
        with log_path.open("a", encoding="utf-8") as stream:
            stream.write(line + "\n")
        progress(line)

    try:
        report(f"[1/6] Input {image.width}x{image.height}; checking local runtime")
        image.save(directory / "original.png")
        image_detection.check_runtime(ocr_runtime)
        model = image_detection.ensure_weights(
            Path(weights) if weights else image_detection.default_weights(),
            offline=offline, progress=report)
        report("[2/6] YOLO detection on CPU")
        yolo = image_detection.run_yolo(image, model, threads=threads, confidence=confidence)
        _write_json(directory / "raw-yolo.json", yolo)
        report(f"[3/6] OCR on CPU; YOLO proposed {len(yolo['boxes'])} regions")
        ocr = image_detection.run_ocr(image, threads=threads, runtime_dir=ocr_runtime,
                                      recover_small_text=small_text_recovery, progress=report)
        _write_json(directory / "raw-ocr.json", ocr)
        _write_json(directory / "ocr-recovery.json", ocr.get("small_text_recovery", {
            "enabled": False, "added_count": 0}))
        report(f"[4/6] Combining graphics and {len(ocr['texts'])} OCR regions")
        result = image_detection.combine_detections(image, yolo, ocr, profile=profile)
        _write_json(directory / "geometry-candidates.json", result.pop("geometry_candidates"))
        elements = result["elements"]
        report("[5/6] Rendering combined, text and control annotation images")
        layers = {"combined": image_analysis.save_image_annotations(
            image, elements, directory, filename="combined.png")}
        for name, is_text in (("text", True), ("controls", False)):
            layer = [e for e in elements if diagram_geometry.is_text(e) == is_text]
            layers[name] = image_analysis.save_image_annotations(
                image, layer, directory, filename=f"{name}.png")
        coordinates = directory / "elements.json"
        result.update(
            ok=True, image_path=str(source), source_image_sha256=source_hash,
            image_width=image.width, image_height=image.height,
            exif_orientation_applied=orientation not in (None, 1),
            coordinate_system={"origin": "top-left", "units": "pixels",
                               "bounds": "left/top inclusive, right/bottom exclusive",
                               "reference": "display-oriented-input", "screen_coordinates": False},
            output_directory=str(directory), coordinates_json=str(coordinates),
            **layers["combined"], layers=layers,
            text_annotated_image=layers["text"]["annotated_image"],
            control_annotated_image=layers["controls"]["annotated_image"],
            model_id=yolo["model_id"], model_sha256=yolo["model_sha256"],
            ocr_engine=ocr["engine"], yolo_seconds=yolo["elapsed_seconds"],
            ocr_seconds=ocr["elapsed_seconds"], local_only=True, api_requests=0,
            manual_geometry=False, reference_annotations=False,
            needs_review_labels=[e["annotation_label"] for e in elements if e.get("ambiguity")],
            uia_verified=False, commands_generated=0,
            notes=["All boxes are observations; detection and OCR may miss or misread content.",
                   "OCR recognition scores do not validate bounding boxes or control functions.",
                   "Diagram geometry supports known outline families, not every circuit symbol."],
        )
        _write_csv(directory / "elements.csv", elements)
        report("[6/6] Writing coordinates, report and HTML viewer")
        result["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        result["resources"] = sorted({p.name for p in directory.iterdir()} | {
            "index.html", "elements.json", "summary.json"})
        _write_viewer(directory, result)
        _write_json(coordinates, result)
        summary = {key: result[key] for key in (
            "ok", "image_path", "output_directory", "image_width", "image_height", "profile",
            "element_count", "graphic_count", "text_count", "yolo_raw_count",
            "yolo_supported_graphics", "geometry_recovered_graphics", "needs_review_labels",
            "ocr_base_text_count", "ocr_recovery_count", "ocr_recovery_seconds",
            "model_id", "model_sha256", "ocr_engine", "yolo_seconds", "ocr_seconds",
            "elapsed_seconds", "resources", "notes")}
        summary["viewer"] = str(directory / "index.html")
        _write_json(directory / "summary.json", summary)
        report(f"Completed: {result['graphic_count']} graphic + "
               f"{result['text_count']} text controls")
        return result
    except Exception as error:
        failure = dict(ok=False, error=f"{type(error).__name__}: {error}",
                       image_path=str(source), output_directory=str(directory),
                       elapsed_seconds=round(time.perf_counter() - started, 3))
        _write_json(directory / "summary.json", failure)
        report(f"FAILED: {failure['error']}")
        raise


def build_parser():
    parser = argparse.ArgumentParser(
        prog="easy_uiauto_image",
        description="Annotate an image locally with YOLO + OCR; export three PNGs and coordinates.")
    parser.add_argument("image_path", help="Input image path (BMP, PNG, JPEG, WebP, ...)")
    parser.add_argument("output_dir", help="Parent folder for a new self-contained result folder")
    parser.add_argument("--profile", choices=("block-diagram", "circuit", "ui"),
                        default="block-diagram", help="Geometry profile; default: block-diagram")
    parser.add_argument("--weights", help="Path to official OmniParser icon_detect_v3/model.pt")
    parser.add_argument("--ocr-runtime", default="",
                        help="Optional isolated RapidOCR package folder")
    parser.add_argument("--threads", type=int, default=4, help="CPU threads (default: 4)")
    parser.add_argument("--confidence", type=float, default=0.05,
                        help="YOLO threshold (default: 0.05)")
    parser.add_argument("--offline", action="store_true", help="Require cached models; no download")
    parser.add_argument("--no-small-text-recovery", action="store_true",
                        help="Disable automatic recovery of small text inside frames")
    parser.add_argument("--json", action="store_true", help="Print the complete result JSON")
    return parser


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    args = build_parser().parse_args(argv)
    try:
        result = analyze_local_image(
            args.image_path, args.output_dir, weights=args.weights, profile=args.profile,
            threads=args.threads, confidence=args.confidence, offline=args.offline,
            ocr_runtime=args.ocr_runtime, small_text_recovery=not args.no_small_text_recovery,
            progress=lambda s: print(s, file=sys.stderr, flush=True))
        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            print(f"Graphics: {result['graphic_count']} | Text: {result['text_count']}")
            print(f"Elapsed: {result['elapsed_seconds']:.2f} seconds")
            print(f"Output: {result['output_directory']}")
            print(f"Open: {Path(result['output_directory']) / 'index.html'}")
        return 0
    except Exception as error:
        print(f"ERROR: {type(error).__name__}: {error}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("Cancelled.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
