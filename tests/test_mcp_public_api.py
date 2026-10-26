"""All registered tools have importable Python APIs, without an MCP context."""

from __future__ import annotations

import asyncio
import importlib
import importlib.util
import inspect
import json
from pathlib import Path

import pytest

from easy_uiauto import Controller
from easy_uiauto import mcp as api
from easy_uiauto.mcp import get_control_tree, get_mode, list_windows, server


def test_every_registered_tool_is_exported_with_its_original_signature():
    tools = asyncio.run(server.mcp.list_tools())
    assert set(api.__all__) == {tool.name for tool in tools}
    for tool in tools:
        function = getattr(api, tool.name)
        assert function is getattr(server, tool.name)
        assert inspect.signature(function) == inspect.signature(getattr(server, tool.name))
        assert tool.name in dir(api)
    assert get_control_tree is server.get_control_tree
    assert list_windows is server.list_windows


def test_imported_tool_runs_without_starting_a_transport_or_server_loop(monkeypatch):
    def forbidden(*_args, **_kwargs):
        pytest.fail("Python calls must not start an MCP server")

    monkeypatch.setattr(server.mcp, "run", forbidden)
    assert json.loads(get_mode())["mode"] in {"operate", "learn", "mixed"}
    assert callable(Controller.input_text)  # Existing controller API stays available.


def test_async_context_is_optional_in_python_and_hidden_from_mcp_schema():
    tools = {tool.name: tool for tool in asyncio.run(server.mcp.list_tools())}
    for name in ("scan_window_knowledge", "explore_ui_workflows", "analyze_ui_image"):
        function = getattr(api, name)
        assert inspect.iscoroutinefunction(function)
        assert inspect.signature(function).parameters["ctx"].default is None
        assert "ctx" not in tools[name].inputSchema["properties"]


def test_async_image_tool_can_be_awaited_without_context(monkeypatch, tmp_path):
    from easy_uiauto.mcp import analyze_ui_image

    monkeypatch.setattr(server, "_vision_api_settings", lambda _model: ("url", "key", "model"))
    monkeypatch.setattr(
        server.image_analysis,
        "analyze_image",
        lambda *_args, **_kwargs: {
            "ok": True,
            "elements": [],
            "annotated_image": str(tmp_path / "sample.png"),
        },
    )
    outcome = json.loads(asyncio.run(analyze_ui_image("sample.png")))
    assert outcome["ok"] is True


def test_package_load_and_unknown_attributes_do_not_import_optional_server(monkeypatch):
    path = Path(api.__file__)
    spec = importlib.util.spec_from_file_location(
        "_lazy_mcp_api_test", path, submodule_search_locations=[str(path.parent)]
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    imported = []
    monkeypatch.setattr(module, "import_module", lambda *args: imported.append(args))
    assert "get_control_tree" in dir(module)
    with pytest.raises(AttributeError):
        _value = module.nonexistent_tool
    assert imported == []


def test_missing_mcp_dependency_has_an_install_instruction(monkeypatch):
    path = Path(api.__file__)
    spec = importlib.util.spec_from_file_location(
        "_missing_mcp_api_test", path, submodule_search_locations=[str(path.parent)]
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    def missing(*_args):
        raise ModuleNotFoundError("No module named mcp", name="mcp")

    monkeypatch.setattr(module, "import_module", missing)
    with pytest.raises(ImportError, match="easy-uiauto\\[mcp\\]"):
        module.__getattr__("list_windows")


def test_type_imports_cover_all_exports_and_marker_is_packaged():
    import ast

    import tomllib

    tree = ast.parse(Path(api.__file__).read_text(encoding="utf-8"))
    imports = {
        alias.name
        for node in ast.walk(tree)
        if isinstance(node, ast.ImportFrom) and node.module == "server"
        for alias in node.names
    }
    assert imports == set(api.__all__)
    root = Path(__file__).resolve().parents[1]
    assert (root / "src/easy_uiauto/py.typed").is_file()
    config = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))
    assert "py.typed" in config["tool"]["setuptools"]["package-data"]["easy_uiauto"]


def test_drag_omitted_xpaths_are_valid_empty_location_arrays(monkeypatch):
    from unittest.mock import Mock

    drag = Mock(return_value="Success")
    monkeypatch.setattr(server.Controller, "drag_control", drag)
    monkeypatch.setattr(server, "CURRENT_MODE", "operate")
    value = json.loads(
        server.drag_control("Editor", source_type="EditControl", target_type="EditControl")
    )
    assert value["ok"] is True
    assert drag.call_args.kwargs["Xpath"] == []
    assert drag.call_args.kwargs["PARAMETERS"]["目的控件Xpath"] == []
