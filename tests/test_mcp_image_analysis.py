"""Saved-image coordinates and MCP behavior without desktop access or network calls."""

import asyncio
import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from easy_uiauto.mcp import image_analysis, scanner, server, visualization


def segments():
    return scanner._validate_segments({
        "controls": [
            {"semantic_name": "Input", "left": 10, "top": 20, "width": 30, "height": 20,
             "confidence": 0.95, "actions": ["click"]},
            {"semantic_name": "Badge", "left": 42, "top": 18, "width": 6, "height": 6,
             "confidence": 0.8, "role": "status"},
        ],
    }, 100, 80)


def test_saved_image_coordinates_and_annotations_do_not_access_desktop(monkeypatch, tmp_path):
    source = tmp_path / "input with spaces.png"
    Image.new("RGB", (100, 80), "white").save(source)
    before = source.read_bytes()
    calls = []

    def analyze(image, *args):
        calls.append(image.size)
        return segments()

    def desktop_forbidden(*_args, **_kwargs):
        raise AssertionError("Saved image analysis must not inspect the desktop")

    monkeypatch.setattr(scanner, "segment_interface", analyze)
    monkeypatch.setattr(scanner, "_find_window", desktop_forbidden)
    monkeypatch.setattr(scanner, "_screenshot_window", desktop_forbidden)
    result = image_analysis.analyze_image(str(source), "url", "secret", "model", "version",
                                          str(tmp_path / "results"), analysis_scale=1)
    assert calls == [(100, 80)]
    assert result["element_count"] == 2
    assert result["truncated"] is False
    assert source.read_bytes() == before
    first, second = result["elements"]
    assert first["id"] == 1 and second["id"] == 2
    assert first["bbox"] == {"x": 8, "y": 18, "width": 34, "height": 24}
    assert first["detected_rect"]["left"] == 10
    assert first["center"] == {"x": 25, "y": 30}
    assert first["normalized_rect"] == {"left": 0.08, "top": 0.225,
                                         "right": 0.42, "bottom": 0.525}
    assert first["status"] == "observed" and not first["uia_verified"]
    assert not first["executable"] and "actions" not in first
    assert first["suggested_actions"] == ["click"]
    assert result["commands_generated"] == 0
    assert json.loads(Path(result["coordinates_json"]).read_text(encoding="utf-8")) == result
    with Image.open(result["annotated_image"]) as annotated:
        assert annotated.size == (100, 80)
        assert annotated.getpixel((8, 30)) == (21, 101, 192)
        assert annotated.getpixel((9, 30)) == (255, 255, 255)
    with Image.open(result["annotated_preview_image"]) as preview:
        assert preview.size == (400, 320)
        assert preview.getpixel((32, 120)) == (21, 101, 192)
        assert preview.getpixel((33, 120)) == (255, 255, 255)


def test_truncation_is_explicit_and_each_run_has_separate_outputs(monkeypatch, tmp_path):
    source = tmp_path / "input.png"
    Image.new("RGB", (100, 80), "white").save(source)
    monkeypatch.setattr(scanner, "segment_interface", lambda *_: segments())
    args = (str(source), "url", "secret", "model", "version", str(tmp_path))
    first = image_analysis.analyze_image(*args, max_elements=1, analysis_scale=1)
    second = image_analysis.analyze_image(*args, max_elements=1, analysis_scale=1)
    assert first["truncated"] and first["detected_element_count"] == 2
    assert first["element_count"] == 1
    assert first["annotated_image"] != second["annotated_image"]
    assert Path(first["annotated_image"]).is_file()


def test_dense_small_controls_get_a_readable_preview_with_original_coordinates(tmp_path):
    image = Image.new('RGB',(1000,700),'white')
    elements = image_analysis.complete_element_bounds([
        dict(id=i+1, entity_type='graphic-control', rect=dict(left=30+(i%8)*100,
             top=30+(i//8)*60, right=46+(i%8)*100, bottom=46+(i//8)*60))
        for i in range(40)], *image.size, 1)
    original = [dict(e['rect']) for e in elements]
    result = image_analysis.save_image_annotations(image,elements,tmp_path)
    with Image.open(result['annotated_preview_image']) as preview:
        origin = result['annotation_preview_layout']['image_origin']
        assert preview.width >= 2000 and preview.height >= 1400
        assert preview.getpixel((elements[0]['rect']['left']*2+origin['x'],
                                 70+origin['y']))==(21,101,192)
    assert [e['rect'] for e in elements]==original


def test_exif_orientation_and_transparency_use_displayed_image_coordinates(monkeypatch, tmp_path):
    source = tmp_path / "rotated.png"
    image = Image.new("RGBA", (100, 80), (0, 0, 0, 0))
    exif = image.getexif()
    exif[274] = 6
    image.save(source, exif=exif)

    def analyze(image, *_):
        assert image.size == (80, 100)
        assert image.getpixel((0, 0)) == (255, 255, 255)
        return scanner._validate_segments({}, *image.size)

    monkeypatch.setattr(scanner, "segment_interface", analyze)
    result = image_analysis.analyze_image(str(source), "url", "secret", "model", "v",
                                          str(tmp_path), analysis_scale=1)
    assert result["exif_orientation_applied"]
    assert result["image_width"] == 80 and result["image_height"] == 100
    assert result["element_count"] == 0


@pytest.mark.parametrize("kind", ["missing", "corrupt", "animation", "limit"])
def test_invalid_input_fails_before_remote_analysis(monkeypatch, tmp_path, kind):
    source = tmp_path / "input.png"
    if kind == "corrupt":
        source.write_text("not an image", encoding="utf-8")
    elif kind == "animation":
        source = tmp_path / "input.gif"
        Image.new("RGB", (20, 20), "white").save(
            source, save_all=True, append_images=[Image.new("RGB", (20, 20), "black")])

    def forbidden(*_):
        raise AssertionError("Invalid inputs must not reach remote analysis")

    monkeypatch.setattr(scanner, "segment_interface", forbidden)
    with pytest.raises((ValueError, OSError)):
        image_analysis.analyze_image(str(source), "url", "key", "model", "v",
                                      str(tmp_path), max_elements=0 if kind == "limit" else 10)


def test_mcp_image_tool_returns_artifact_paths_and_structured_errors(monkeypatch, tmp_path):
    source = tmp_path / "input.png"
    Image.new("RGB", (100, 80), "white").save(source)
    monkeypatch.setattr(server, "_vision_api_settings", lambda _: ("url", "key", "model"))
    monkeypatch.setattr(scanner, "segment_interface", lambda *_: segments())
    progress = []

    async def report(*args):
        progress.append(args)

    context = SimpleNamespace(report_progress=report)
    result = json.loads(asyncio.run(server.analyze_ui_image(
        str(source), context, str(tmp_path), analysis_scale=1, engine="api")))
    assert result["ok"] and Path(result["annotated_image"]).is_file()
    assert progress[0][0] == 0 and progress[-1][0] == 1
    failure = json.loads(asyncio.run(server.analyze_ui_image(
        str(source)+"missing", context, engine="api")))
    assert failure["ok"] is False and "FileNotFoundError" in failure["error"]


def test_mcp_advertises_image_file_input():
    tools = asyncio.run(server.mcp.list_tools())
    tool = next(item for item in tools if item.name == "analyze_ui_image")
    assert tool.inputSchema["required"] == ["image_path"]
    assert "output_dir" in tool.inputSchema["properties"]
    assert tool.inputSchema["properties"]["engine"]["default"] == "local"
    assert tool.inputSchema["properties"]["auto_crop"]["default"] is True
    assert tool.inputSchema["properties"]["confidence"]["default"] == 0.1
    assert tool.inputSchema["properties"]["rectangle_snap"]["default"] is True
    assert visualization.STATUS_COLORS["observed"] == "#00b8d4"


def test_image_profile_schema_and_removed_profile_never_reach_a_model(monkeypatch, tmp_path):
    def forbidden(*_a, **_kw):
        raise AssertionError('Unsupported profiles must fail before model or file access')

    tools = asyncio.run(server.mcp.list_tools())
    analysis = next(t.inputSchema for t in tools if t.name=='analyze_ui_image')
    correction = next(t.inputSchema for t in tools if t.name=='correct_ui_image_annotations')
    assert set(analysis['properties']['profile']['enum'])=={'ui','block-diagram'}
    assert set(correction['properties']['profile']['enum'])=={'icons','block-diagram','rectangles'}
    monkeypatch.setattr(server, '_vision_api_settings', forbidden)
    monkeypatch.setattr(server.image_cli, 'analyze_local_image', forbidden)
    for engine in ('local','api'):
        result=json.loads(asyncio.run(server.analyze_ui_image('missing.png',engine=engine,profile='circuit')))
        assert not result['ok'] and 'profile' in result['error']
    result=json.loads(asyncio.run(server.correct_ui_image_annotations('missing.json',profile='circuit')))
    assert not result['ok'] and 'profile' in result['error']
    assert not list(tmp_path.iterdir())


def test_progress_disconnect_does_not_discard_completed_image_result(monkeypatch, tmp_path):
    source = tmp_path / "input.png"
    Image.new("RGB", (100, 80), "white").save(source)
    monkeypatch.setattr(server, "_vision_api_settings", lambda _: ("url", "key", "model"))
    monkeypatch.setattr(scanner, "segment_interface", lambda *_: segments())

    async def broken_progress(*_):
        raise OSError("Progress transport closed")

    result = json.loads(asyncio.run(server.analyze_ui_image(
        str(source), SimpleNamespace(report_progress=broken_progress), str(tmp_path),
        analysis_scale=1, engine="api")))
    assert result["ok"] and Path(result["coordinates_json"]).is_file()


def test_mcp_defaults_to_local_without_resolving_api_credentials(monkeypatch, tmp_path):
    calls = []

    def remote_forbidden(*_args, **_kwargs):
        raise AssertionError("Default image annotation must not resolve or use an API")

    def local(path, output, **options):
        calls.append((path, output, options))
        return {"ok": True, "engine": "local", "api_requests": 0, "elements": []}

    monkeypatch.setattr(server, "_vision_api_settings", remote_forbidden)
    monkeypatch.setattr(server.image_analysis, "analyze_image", remote_forbidden)
    monkeypatch.setattr(server.image_cli, "analyze_local_image", local)
    result = json.loads(asyncio.run(server.analyze_ui_image("input.png", output_dir=str(tmp_path))))
    assert result["ok"] and result["engine"] == "local" and result["api_requests"] == 0
    path, output, options = calls[0]
    assert (path, output) == ("input.png", str(tmp_path))
    assert options["profile"] == "ui" and options["auto_crop"] is True
    assert options["max_elements"] == 1000 and options["box_padding"] == 1
    assert options["confidence"] == 0.1 and options["rectangle_snap"] is True


def test_missing_local_dependencies_never_trigger_a_remote_fallback(monkeypatch):
    def remote_forbidden(*_args, **_kwargs):
        raise AssertionError("A local failure must not upload the user's image")

    def unavailable(*_args, **_kwargs):
        raise RuntimeError('Install local dependencies with: pip install "easy-uiauto[image]"')

    monkeypatch.setattr(server, "_vision_api_settings", remote_forbidden)
    monkeypatch.setattr(server.image_cli, "analyze_local_image", unavailable)
    result = json.loads(asyncio.run(server.analyze_ui_image("input.png")))
    assert result["ok"] is False and "easy-uiauto[image]" in result["error"]


@pytest.mark.parametrize("options", [{"engine": "unknown"}, {"model": "remote-model"}])
def test_invalid_engine_or_implicit_remote_options_do_not_upload(monkeypatch, options):
    def forbidden(*_args, **_kwargs):
        raise AssertionError("Invalid routing must fail before any analysis")

    monkeypatch.setattr(server, "_vision_api_settings", forbidden)
    monkeypatch.setattr(server.image_cli, "analyze_local_image", forbidden)
    result = json.loads(asyncio.run(server.analyze_ui_image("input.png", **options)))
    assert result["ok"] is False and "ValueError" in result["error"]


def test_small_image_enlargement_maps_coordinates_back_to_original(monkeypatch, tmp_path):
    source = tmp_path / "tiny.png"
    Image.new("RGB", (100, 80), "white").save(source)

    def analyze(image, *_):
        assert image.size == (400, 320)
        return scanner._validate_segments({"controls": [
            {"left": 41, "top": 81, "width": 118, "height": 78,
             "semantic_name": "Tiny icon", "confidence": 0.9}]}, *image.size)

    monkeypatch.setattr(scanner, "segment_interface", analyze)
    result = image_analysis.analyze_image(str(source), "url", "key", "model", "v",
                                          str(tmp_path))
    assert result["image_width"] == 100 and result["analysis_image_width"] == 400
    assert result["elements"][0]["bbox"] == {"x": 8, "y": 18, "width": 34, "height": 24}
    assert result["elements"][0]["detected_rect"]["left"] == 10
    assert result["regions"][0]["rect"]["width"] == 100
    with Image.open(result["annotated_image"]) as annotated:
        assert annotated.size == (100, 80)


def test_attached_badge_label_does_not_obscure_parent_label_or_icon():
    markers = [
        {"index": 6, "rect": {"left": 182, "top": 63, "right": 200, "bottom": 83}},
        {"index": 7, "rect": {"left": 193, "top": 59, "right": 201, "bottom": 67}},
    ]
    labels = visualization._annotation_label_rects(markers, (320, 167))
    obstacles = [*labels, *(tuple(marker["rect"].values()) for marker in markers)]
    for index, label in enumerate(labels):
        assert 0 <= label[0] < label[2] <= 320
        assert 0 <= label[1] < label[3] <= 167
        for obstacle_index, other in enumerate(obstacles):
            if index != obstacle_index:
                assert not (label[0] < other[2] and label[2] > other[0]
                            and label[1] < other[3] and label[3] > other[1])


def test_parent_encloses_attached_badge_and_padding_is_idempotent():
    items = [
        {"id": 6, "intent": "open-theme", "rect": {
            "left": 10, "top": 20, "right": 30, "bottom": 40, "width": 20, "height": 20}},
        {"id": 7, "relationships": [{"type": "attached-to", "target_intent": "open-theme"}],
         "rect": {"left": 28, "top": 17, "right": 35, "bottom": 24, "width": 7, "height": 7}},
    ]
    result = image_analysis.complete_element_bounds(items, 100, 80)
    assert result[0]["bbox"] == {"x": 8, "y": 15, "width": 29, "height": 27}
    assert result[0]["component_element_ids"] == [7]
    assert result[1]["parent_element_id"] == 6
    assert result[0]["detected_rect"] == items[0]["rect"]
    assert image_analysis.complete_element_bounds(result, 100, 80) == result
    edge = image_analysis.complete_element_bounds([
        {"id": 1, "rect": {"left": 0, "top": 0, "right": 100, "bottom": 80,
                            "width": 100, "height": 80}}], 100, 80)
    assert edge[0]["bbox"] == {"x": 0, "y": 0, "width": 100, "height": 80}


def test_analysis_scope_is_forwarded_without_desktop_access(monkeypatch, tmp_path):
    source = tmp_path / "circuit.png"
    Image.new("RGB", (100, 80), "white").save(source)

    def analyze(image, *_args):
        assert _args[-1] == "Complete diagram components"
        return segments()

    monkeypatch.setattr(scanner, "segment_interface", analyze)
    result = image_analysis.analyze_image(str(source), "url", "key", "model", "v",
                                          str(tmp_path), analysis_scale=1, box_padding=0,
                                          analysis_context="Complete diagram components")
    assert result["elements"][0]["bbox"]["x"] == 10


def test_graphics_and_text_have_distinct_stable_display_ids_across_layers(tmp_path):
    image = Image.new("RGB", (240, 160), "white")
    graphic = {"id": 4, "entity_type": "graphic-control", "rect": {
        "left": 30, "top": 30, "right": 210, "bottom": 140}}
    text = {"id": 49, "entity_type": "text-label", "rect": {
        "left": 60, "top": 80, "right": 100, "bottom": 95}, "related_element_id": 4}
    second_text = {"id": 50, "entity_type": "text-label", "rect": {
        "left": 110, "top": 80, "right": 150, "bottom": 95}}
    result = image_analysis.save_image_annotations(image, [graphic, text, second_text], tmp_path)
    assert [e["id"] for e in [graphic, text, second_text]] == [4, 49, 50]
    assert [e["annotation_label"] for e in [graphic, text, second_text]] == ["G1", "T1", "T2"]
    assert graphic["annotation_color"] != text["annotation_color"]
    with Image.open(result["annotated_image"]) as annotated:
        assert annotated.getpixel((30, 110)) == (21, 101, 192)
        assert annotated.getpixel((80, 94)) == (198, 93, 0)
    layer = tmp_path / "text-layer"
    layer.mkdir()
    text_result = image_analysis.save_image_annotations(image, [second_text], layer)
    assert text_result["annotation_layout"]["labels"][0]["annotation_label"] == "T2"
    assert second_text["id"] == 50
