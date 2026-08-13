"""Tests for learned-control page annotations and live overlay preparation."""

from __future__ import annotations

from types import SimpleNamespace

from PIL import Image

from easy_uiauto.mcp import knowledge, scanner, server, visualization


def _rect(left: int, top: int, right: int, bottom: int) -> dict:
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "width": right - left,
        "height": bottom - top,
    }


def _control(control_id: str, status: str = "verified") -> dict:
    return {
        "id": control_id,
        "app_id": "example",
        "app_name": "Example",
        "page_id": "main",
        "region_id": "content",
        "semantic_name": control_id.title(),
        "intent": control_id,
        "semantic_status": "verified",
        "command": f"main.content.{control_id}",
        "actions": ["click"],
        "status": status,
        "normalized_rect": {
            "left": 0.025,
            "top": 0.066667,
            "right": 0.275,
            "bottom": 0.2,
        },
        "geometry_role": "visual-hint-only",
        "location": {"WindowName": "Example", "Xpath": []},
    }


def test_control_markers_follow_window_move_and_resize() -> None:
    markers = visualization.control_markers(
        [_control("save")],
        _rect(100, 200, 500, 500),
        _rect(1000, 500, 1800, 1100),
        {"save"},
    )

    assert markers[0]["rect"] == _rect(1020, 540, 1220, 620)
    assert markers[0]["target"] is True
    assert markers[0]["index"] == 1


def test_save_scan_annotation_writes_numbered_image(tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    screenshot = Image.new("RGB", (400, 300), "white")

    path, markers = visualization.save_scan_annotation(
        directory,
        "main",
        screenshot,
        _rect(100, 200, 500, 500),
        [_control("save")],
    )

    assert path.is_file()
    assert len(markers) == 1
    with Image.open(path) as annotated:
        assert annotated.getpixel((10, 20)) != (255, 255, 255)


def test_show_page_controls_lists_only_executable_controls(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    verified = _control("save")
    quarantined = _control("delete", "quarantined")
    index = {
        "application": {"id": "example", "name": "Example"},
        "pages": [{"id": "main", "name": "Main"}],
        "controls": [verified, quarantined],
    }
    monkeypatch.setattr(
        visualization,
        "_match_current_page",
        lambda *_args: {
            "index": index,
            "page": index["pages"][0],
            "similarity": 0.97,
            "window": object(),
            "window_rect": _rect(200, 300, 600, 600),
            "screenshot": Image.new("RGB", (400, 300)),
        },
    )
    monkeypatch.setattr(
        visualization,
        "_runtime_control_markers",
        lambda *_args: (
            [
                {
                    "index": 1,
                    "control_id": "save",
                    "label": "Save",
                    "status": "verified",
                    "rect": _rect(210, 320, 310, 360),
                    "resolution_method": "location",
                }
            ],
            [],
        ),
    )
    monkeypatch.setattr(
        knowledge,
        "available_commands",
        lambda *_args: [
            {
                "control_id": "save",
                "command": "main.content.save.click",
            }
        ],
    )
    shown = []
    monkeypatch.setattr(
        visualization,
        "show_markers",
        lambda markers, duration: shown.extend(markers) or {"shown": True},
    )

    result = visualization.show_page_controls(directory)

    assert result["page_id"] == "main"
    assert result["page_similarity"] == 0.97
    assert [item["control_id"] for item in result["controls"]] == ["save"]
    assert result["controls"][0]["commands"] == ["main.content.save.click"]
    assert len(shown) == 1


def test_runtime_markers_use_current_location_not_saved_geometry(monkeypatch, tmp_path) -> None:
    directory = knowledge.initialize_app("example", "Example", tmp_path)
    current = SimpleNamespace(
        BoundingRectangle=SimpleNamespace(left=520, top=330, right=680, bottom=380)
    )
    monkeypatch.setattr(scanner, "resolve_location", lambda *_args, **_kwargs: current)
    record = _control("save")
    window_rect = _rect(500, 300, 900, 600)

    markers, unresolved = visualization._runtime_control_markers(
        directory,
        [record],
        object(),
        window_rect,
        Image.new("RGB", (400, 300), "white"),
    )

    assert unresolved == []
    assert markers[0]["rect"] == _rect(520, 330, 680, 380)
    assert markers[0]["resolution_method"] == "location"


def test_mcp_show_ui_controls_returns_numbered_legend(monkeypatch) -> None:
    monkeypatch.setattr(
        server.visualization,
        "show_page_controls",
        lambda *_args: {
            "ok": True,
            "page_id": "main",
            "controls": [{"index": 1, "semantic_name": "Save"}],
        },
    )

    result = server.show_ui_controls("example")

    assert '"page_id": "main"' in result
    assert '"semantic_name": "Save"' in result
