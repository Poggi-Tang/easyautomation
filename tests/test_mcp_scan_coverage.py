"""Regressions for sparse desktops, coarse UIA providers and tiled screenshots."""

import json
from types import SimpleNamespace

import pytest
from PIL import Image

from easy_uiauto.mcp import knowledge, scanner
from test_mcp_scanner import FakeControl


def target(left, top=20, width=60, height=40, **values):
    return {
        "semantic_name": "Open item", "intent": "open-item", "role": "action",
        "actions": ["click"], "confidence": 0.97, "risk": "safe",
        "region_id": "content", "source": "ai-vision-target",
        "relative_rect": {
            "left": left, "top": top, "right": left + width, "bottom": top + height,
            "width": width, "height": height,
        },
        **values,
    }


@pytest.mark.parametrize("body, reason", [
    ({"choices": [{"message": {"content": None}}]}, "no completion content"),
    ({"error": {"message": "private upstream detail"}}, "error response"),
    ({"choices": [{"finish_reason": "length", "message": {"content": "{"}}]},
     "truncated"),
])
def test_invalid_vision_completions_have_actionable_errors(body, reason):
    response = SimpleNamespace(read=lambda: json.dumps(body).encode())
    with pytest.raises(RuntimeError, match=reason):
        scanner._read_completion_content(response)


def test_image_segmentation_requests_validated_json_and_keeps_all_items(monkeypatch):
    items = [{"semantic_name": "App", "left": 10, "top": 10, "width": 40, "height": 40},
             {"semantic_name": "App", "left": 60, "top": 10, "width": 40, "height": 40}]
    response = SimpleNamespace(read=lambda: json.dumps({"choices": [{"message": {
        "content": json.dumps({"page": {"id": "desktop"}, "regions": [], "controls": items})
    }}]}).encode())
    class Context:
        def __enter__(self):
            return response

        def __exit__(self, *_):
            pass

    def open_request(request, **_):
        payload = json.loads(request.data)
        assert payload["response_format"] == {"type": "json_object"}
        assert payload["stream"] is False
        return Context()

    monkeypatch.setattr(scanner.urlrequest, "urlopen", open_request)
    result = scanner._segment_image(Image.new("RGB", (200, 100)), "https://example.test",
                                    "key", "model", "0.8.0")
    assert len(result["controls"]) == 2
    assert [item["relative_rect"]["left"] for item in result["controls"]] == [10, 60]


def test_small_visual_target_rejects_named_full_desktop_container():
    pane = FakeControl("File Explorer", "PaneControl", (0, 0, 4480, 1440),
                       automation_id="desktop")
    window = FakeControl("Program Manager", "WindowControl", (0, 0, 4480, 1440), [pane])
    assert scanner.control_from_visual_target(
        window, scanner._rect(window), target(100), point_lookup=lambda *_: pane
    ) is None


def test_disjoint_controls_with_identical_xpath_do_not_merge(monkeypatch):
    first = FakeControl("Item", "ButtonControl", (20, 20, 80, 60))
    second = FakeControl("Item", "ButtonControl", (120, 20, 180, 60))
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [first, second])
    monkeypatch.setattr(scanner, "get_control_xpath", lambda _: [
        {"ControlType": "WindowControl", "Name": "Example"},
        {"ControlType": "ButtonControl", "Name": "Item"},
    ])
    controls, semantics, failures = scanner.controls_from_visual_targets(
        window, scanner._rect(window),
        [target(20, _uia_control=first), target(120, _uia_control=second)], 20,
    )
    assert [control for control, _ in controls] == [first, second]
    assert len(semantics) == 2
    assert failures == []


def test_duplicate_detection_of_same_visual_item_is_deduplicated(monkeypatch):
    button = FakeControl("Save", "ButtonControl", (20, 20, 80, 60))
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [button])
    monkeypatch.setattr(scanner, "get_control_xpath", lambda _: [
        {"ControlType": "ButtonControl", "Name": "Save"}
    ])
    controls, _, failures = scanner.controls_from_visual_targets(
        window, scanner._rect(window),
        [target(20, _uia_control=button, semantic_name="Save", intent="save-file"),
         target(21, _uia_control=button, semantic_name="Save", intent="save-file")], 20,
    )
    assert len(controls) == 1
    assert failures == []


def test_coarse_canvas_keeps_two_independent_non_executable_records(monkeypatch, tmp_path):
    canvas = FakeControl("Canvas", "PaneControl", (0, 0, 400, 300))
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [canvas])
    monkeypatch.setattr(scanner, "_find_window", lambda _: window)
    monkeypatch.setattr(scanner, "_activate_window", lambda _: None)
    monkeypatch.setattr(scanner, "_process_name", lambda _: "example")
    monkeypatch.setattr(scanner, "_screenshot_window", lambda _: Image.new("RGB", (400, 300)))
    monkeypatch.setattr(scanner.uiautomation, "ControlFromPoint", lambda *_: canvas)
    monkeypatch.setattr(scanner, "segment_interface", lambda *_: {
        "page": {"id": "main", "name": "Main"},
        "regions": scanner._validate_segments({}, 400, 300)["regions"],
        "controls": [target(20), target(200)],
    })
    result = scanner.scan_window("Example", "unused", "unused", "unused", "0.8.0",
                                 root=tmp_path)
    records = knowledge.load_index(knowledge.app_dir("example", tmp_path))["controls"]
    assert len(records) == 2
    assert len({record["id"] for record in records}) == 2
    assert all(record["status"] == "observed" and not record["actions"] for record in records)
    assert result["annotated_controls"] == 2
    assert result["commands"] == 0


def test_native_desktop_coverage_includes_omitted_shortcut_on_second_monitor(monkeypatch):
    icon = FakeControl("SimuLab5.2.0.release.1", "ListItemControl", (2560, 670, 2634, 730))
    hidden = FakeControl("Hidden", "ListItemControl", (2700, 670, 2774, 730))
    hidden.IsOffscreen = True
    listing = FakeControl("Desktop", "ListControl", (2560, 0, 4480, 1440), [icon, hidden])
    listing.ClassName = "SysListView32"
    worker = FakeControl("", "PaneControl", (2560, 0, 4480, 1440), [listing])
    worker.ClassName = "WorkerW"
    worker.NativeWindowHandle = 2
    window = FakeControl("Program Manager", "PaneControl", (0, 0, 4480, 1440))
    window.ClassName = "Progman"
    unrelated = FakeControl("Other app", "WindowControl", (0, 0, 500, 500))
    unrelated.GetChildren = lambda: (_ for _ in ()).throw(AssertionError("Other app traversed"))
    monkeypatch.setattr(scanner.native_windows, "enumerate_windows", lambda **_kw: [
        {"handle": 1, "class_name": "Progman"},
        {"handle": 2, "class_name": "WorkerW"},
        {"handle": 3, "class_name": "Other app"}])
    monkeypatch.setattr(scanner.uiautomation, "GetRootControl", lambda: (_ for _ in ()).throw(
        AssertionError("Desktop UIA root must not be enumerated")))
    monkeypatch.setattr(scanner.uiautomation, "ControlFromHandle", lambda handle: worker
                        if handle == 2 else (_ for _ in ()).throw(
                            AssertionError("Unrelated application was bound to UIA")))
    monkeypatch.setattr(scanner.uiautomation, "ControlFromPoint", lambda *_: icon)
    segmentation = scanner._validate_segments({}, 4480, 1440)
    scanner._supplement_desktop_targets(window, scanner._rect(window), segmentation, 100)
    assert [item["observed_value"] for item in segmentation["controls"]] == [icon.Name]
    assert segmentation["controls"][0]["relative_rect"]["left"] == 2560
    assert segmentation["coverage"]["desktop_items_added"] == 1


def test_native_desktop_coverage_excludes_occluded_icons(monkeypatch):
    icon = FakeControl("Covered", "ListItemControl", (20, 20, 80, 60))
    window = FakeControl("Program Manager", "PaneControl", (0, 0, 400, 300), [icon])
    window.ClassName = "Progman"
    foreground = FakeControl("Foreground app", "PaneControl", (0, 0, 400, 300))
    monkeypatch.setattr(scanner, "_desktop_shell_roots", lambda _: [window])
    monkeypatch.setattr(scanner.uiautomation, "ControlFromPoint", lambda *_: foreground)
    segmentation = scanner._validate_segments({}, 400, 300)
    scanner._supplement_desktop_targets(window, scanner._rect(window), segmentation, 100)
    assert segmentation["controls"] == []
    assert segmentation["coverage"]["desktop_items_occluded"] == 1


def test_native_desktop_visibility_accepts_icon_text_child(monkeypatch):
    label = FakeControl("Item", "TextControl", (20, 40, 80, 60))
    icon = FakeControl("Item", "ListItemControl", (20, 20, 80, 60), [label])
    monkeypatch.setattr(scanner.uiautomation, "ControlFromPoint", lambda *_: label)
    assert scanner._desktop_item_is_visible(icon)


def test_desktop_vision_only_processes_taskbars_and_translates_coordinates(monkeypatch):
    window = FakeControl("Program Manager", "PaneControl", (0, 0, 2000, 1200))
    window.ClassName = "Progman"
    bar = FakeControl("Taskbar", "PaneControl", (1000, 950, 2000, 1000))
    bar.ClassName = "Shell_SecondaryTrayWnd"
    monkeypatch.setattr(scanner, "_desktop_shell_roots", lambda _: [window, bar])
    calls = []

    def segment(image, *_):
        calls.append(image.size)
        result = scanner._validate_segments({}, *image.size)
        result["controls"] = [target(20, top=5, width=30, height=30)]
        return result

    monkeypatch.setattr(scanner, "_segment_image", segment)
    result = scanner._segment_desktop(Image.new("RGB", (2000, 1200)), window,
                                      scanner._rect(window), "unused", "key", "model",
                                      "0.8.0", 100, lambda _: None)
    assert calls == [(1000, 50)]
    assert result["controls"][0]["relative_rect"]["left"] == 1020
    assert result["controls"][0]["relative_rect"]["top"] == 955
    assert result["coverage"]["desktop_strategy"] == "native-items-and-visual-taskbars"


def test_tiles_translate_coordinates_and_keep_adjacent_identical_items(monkeypatch):
    screen = Image.new("RGB", (2000, 300))
    screen.putpixel((720, 0), (1, 0, 0))

    def segment(image, *_):
        second = image.getpixel((0, 0)) == (1, 0, 0)
        result = scanner._validate_segments({}, *image.size)
        result["controls"] = ([target(80, observed_value="App"),
                                target(200, observed_value="App")]
                               if second else [target(800, observed_value="App")])
        return result

    monkeypatch.setattr(scanner, "_segment_image", segment)
    updates = []
    result = scanner.segment_interface(screen, "unused", "unused", "unused", "0.8.0",
                                       updates.append)
    assert sorted(item["relative_rect"]["left"] for item in result["controls"]) == [800, 920]
    assert result["coverage"]["visual_tiles"] == 2
    assert len(updates) == 2


def test_tile_deduplication_prefers_complete_component_over_clipped_fragment(monkeypatch):
    screen = Image.new("RGB", (2000, 300))
    screen.putpixel((720, 0), (1, 0, 0))

    def segment(image, *_):
        second = image.getpixel((0, 0)) == (1, 0, 0)
        result = scanner._validate_segments({}, *image.size)
        result["controls"] = [target(480 if second else 1200,
                                     width=110 if second else 80, observed_value="Component")]
        return result

    monkeypatch.setattr(scanner, "_segment_image", segment)
    result = scanner.segment_interface(screen, "unused", "key", "model", "v")
    assert len(result["controls"]) == 1
    assert result["controls"][0]["relative_rect"]["left"] == 1200
    assert result["controls"][0]["relative_rect"]["right"] == 1310
    assert result["coverage"]["edge_clipped_targets"] == 0


def test_cross_tile_caption_conflict_does_not_duplicate_same_pixel_component():
    controls = []
    first = target(40, role="component", semantic_name="Busbar component",
                   observed_value="Bus 8", source_tiles=[1])
    second = target(40, role="component", semantic_name="Busbar component",
                    observed_value="Bus B", source_tiles=[2])
    scanner._merge_tile_target(controls, first)
    scanner._merge_tile_target(controls, second)
    assert len(controls) == 1
    assert controls[0]["observed_value_variants"] == ["Bus 8", "Bus B"]
    assert controls[0]["source_tiles"] == [1, 2]
    # Separate visual entities within one tile keep their own identities.
    scanner._merge_tile_target(controls, target(
        40, role="component", semantic_name="Busbar component",
        observed_value="Bus C", source_tiles=[2]))
    assert len(controls) == 2


def test_clipped_fragment_with_caption_difference_uses_complete_bounds():
    controls = []
    fragment = target(1219, top=531, width=61, height=62, role="component",
                      semantic_name="Transformer component", source_tiles=[1],
                      observed_value="", _tile_clipped=True)
    complete = target(1219, top=531, width=82, height=48, role="component",
                      semantic_name="Transformer component", source_tiles=[2],
                      observed_value="Transformer 1", _tile_clipped=False)
    scanner._merge_tile_target(controls, fragment)
    scanner._merge_tile_target(controls, complete)
    assert len(controls) == 1
    assert controls[0]["relative_rect"] == complete["relative_rect"]


def test_ancestor_automation_id_never_resolves_as_leaf(monkeypatch):
    icon = FakeControl("App", "ListItemControl", (20, 20, 80, 60))
    listing = FakeControl("Desktop", "ListControl", (0, 0, 400, 300), [icon], "1")
    window = FakeControl("Example", "WindowControl", (0, 0, 400, 300), [listing])
    monkeypatch.setattr(scanner.uiautomation, "Control", lambda **_: (
        _ for _ in ()).throw(AssertionError("An ancestor ID must not short-circuit the leaf")))
    location = {
        "WindowName": "Example", "Name": "App", "ControlType": "ListItemControl",
        "foundIndex": 64, "Xpath": [
            {"ControlType": "WindowControl", "Name": "Example", "searchDepth": 1},
            {"ControlType": "ListControl", "AutomationId": "1", "searchDepth": 2},
            {"ControlType": "ListItemControl", "Name": "App", "foundIndex": 64,
             "searchDepth": 3},
        ],
    }
    assert scanner.resolve_location(location, window=window) is icon


@pytest.mark.parametrize("length,size", [(1281, 1280), (2000, 1280), (2449, 1280),
                                         (2560, 1280), (4480, 1280), (1288, 960)])
def test_tile_origins_cover_axis_without_near_identical_trailing_tiles(length, size):
    origins = scanner._tile_origins(length, size)
    assert origins[0] == 0 and origins[-1] + size == length
    assert all(0 < b - a <= size - 80 for a, b in zip(origins, origins[1:], strict=False))
    assert all(b - a >= (length - size) / len(origins)
               for a, b in zip(origins, origins[1:], strict=False))
    if length == 2449:
        assert len(origins) == 2


def test_cross_tile_identity_is_not_blocked_by_generic_model_types_and_roles():
    controls = []
    first = target(200, entity_type="subsystem", role="entity", source_tiles=[1],
                   semantic_name="Controller", observed_value="GA", confidence=0.9)
    second = target(201, entity_type="ui-element", role="readonly", source_tiles=[2],
                    semantic_name="PGV block", observed_value="PGV", confidence=0.9)
    scanner._merge_tile_target(controls, first)
    scanner._merge_tile_target(controls, second)
    assert len(controls) == 1
    assert controls[0]["relative_rect"] in (first["relative_rect"], second["relative_rect"])
    assert controls[0]["source_tiles"] == [1, 2]
    assert controls[0]["observed_value_variants"] == ["GA", "PGV"]
    assert len(controls[0]["tile_observations"]) == 2


def test_cross_tile_geometry_keeps_badge_and_neighboring_controls_separate():
    controls = []
    for item in [target(100, source_tiles=[1]), target(158, source_tiles=[2]),
                 target(100, source_tiles=[2], entity_type="badge", relationships=[
                     {"type": "attached-to", "target_intent": "open-item"}])]:
        scanner._merge_tile_target(controls, item)
    assert len(controls) == 3


def test_text_and_graphic_can_overlap_without_suppressing_each_other():
    controls = []
    scanner._merge_tile_target(controls, target(100, entity_type="constant", source_tiles=[1]))
    scanner._merge_tile_target(controls, target(100, entity_type="text-label", source_tiles=[2],
                                               observed_value="1"))
    scanner._merge_tile_target(controls, target(101, entity_type="text-label", source_tiles=[3],
                                               observed_value="1"))
    assert len(controls) == 2
    assert controls[0]["entity_type"] == "constant"
    assert controls[1]["source_tiles"] == [2, 3]
