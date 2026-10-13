"""Exercise strict lookup with fake UIA trees, without loading Windows modules."""

import ast
from contextlib import contextmanager
from contextvars import ContextVar
import importlib.util
from pathlib import Path
import sys
from types import ModuleType, SimpleNamespace

import pytest


SOURCE = Path(__file__).resolve().parents[1] / "src" / "easy_uiauto"


@pytest.fixture
def resolver(monkeypatch):
    # A private package keeps portable tests independent of the production package
    # __init__, which intentionally imports Windows desktop dependencies.
    package = "_control_resolution_test"
    for name, path in ((package, SOURCE), (f"{package}.mcp", SOURCE / "mcp")):
        module = ModuleType(name)
        module.__path__ = [str(path)]
        monkeypatch.setitem(sys.modules, name, module)
    for name in ("results", "protocol"):
        fullname = f"{package}.mcp.{name}"
        spec = importlib.util.spec_from_file_location(fullname, SOURCE / "mcp" / f"{name}.py")
        module = importlib.util.module_from_spec(spec)
        monkeypatch.setitem(sys.modules, fullname, module)
        spec.loader.exec_module(module)
    selected = {
        "strict_control_resolution", "_observe_match", "observe_control", "find_control",
    }
    tree = ast.parse((SOURCE / "utils.py").read_text(encoding="utf-8"))
    nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in selected]
    namespace = {
        "__package__": package,
        "contextmanager": contextmanager,
        "_STRICT_CONTROL_RESOLUTION": ContextVar("test_strict", default=False),
        "_ACTIVATE_RESOLVED_WINDOW": ContextVar("test_activate", default=False),
        "CONTROL_TYPE_IDS": {"WindowControl": 1, "ButtonControl": 2, "PaneControl": 3},
        "uiautomation": SimpleNamespace(),
        "timeit": lambda function: function,
    }
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(SOURCE / "utils.py"), "exec"), namespace)
    namespace["ActionError"] = sys.modules[f"{package}.mcp.results"].ActionError
    return namespace


class Control:
    def __init__(self, name="", control_type=2, children=(), automation_id="", class_name=""):
        self.Name = name
        self.ControlType = control_type
        self.AutomationId = automation_id
        self.ClassName = class_name
        self.children = list(children)
        self.lookups = []

    def Control(self, foundIndex, searchDepth, **conditions):
        self.lookups.append((foundIndex, searchDepth, conditions))
        candidates = []

        def visit(parent, depth):
            for child in parent.children:
                if all(getattr(child, key) == value for key, value in conditions.items()):
                    candidates.append(child)
                if depth > 1:
                    visit(child, depth - 1)

        visit(self, searchDepth)
        if foundIndex <= len(candidates):
            return candidates[foundIndex - 1]
        return SimpleNamespace(Refind=lambda **kwargs: False)

    def Refind(self, **kwargs):
        assert kwargs == {"maxSearchSeconds": 0, "raiseException": False}
        return True


def desktop(resolver, children=()):
    window = Control("Editor", 1, children)
    root = Control("Desktop", 0, [window])
    resolver["uiautomation"].GetRootControl = lambda: root
    return root, window


def error_code(resolver, location):
    with pytest.raises(resolver["ActionError"]) as raised:
        resolver["observe_control"](location)
    return raised.value.code


def test_absence_and_invalid_window_are_distinct(resolver):
    desktop(resolver)
    assert error_code(resolver, {"WindowName": "Missing", "Name": "Save"}) == "window_invalid"
    assert error_code(resolver, {"WindowName": "Editor", "Name": "Save"}) == "control_not_found"


def test_ambiguous_selector_requires_explicit_index(resolver):
    first, second = Control("Save"), Control("Save")
    desktop(resolver, [first, second])
    location = {"WindowName": "Editor", "Name": "Save"}
    assert error_code(resolver, location) == "ambiguous_match"
    assert resolver["observe_control"]({**location, "foundIndex": 2}) is second
    assert error_code(resolver, {**location, "foundIndex": 3}) == "control_not_found"


def test_duplicate_windows_are_ambiguous(resolver):
    root, _ = desktop(resolver)
    root.children.append(Control("Editor", 1))
    assert error_code(resolver, {"WindowName": "Editor"}) == "ambiguous_match"


def test_stable_id_ignores_changed_caption_without_activation(resolver):
    button = Control("New caption", automation_id="save")
    desktop(resolver, [button])
    assert resolver["observe_control"]({
        "WindowName": "Editor", "Name": "Old caption", "AutomationId": "save",
    }) is button


def test_xpath_disambiguates_duplicate_ids(resolver):
    first, second = Control("Save", automation_id="save"), Control("Save", automation_id="save")
    desktop(resolver, [Control("Left", 3, [first]), Control("Right", 3, [second])])
    assert resolver["observe_control"]({
        "WindowName": "Editor", "AutomationId": "save",
        "Xpath": [
            {"ControlType": "WindowControl", "Name": "Editor", "searchDepth": 1},
            {"ControlType": "PaneControl", "Name": "Right", "searchDepth": 2},
            {"ControlType": "ButtonControl", "AutomationId": "save", "searchDepth": 3},
        ],
    }) is second


def test_backend_failure_never_becomes_absence(resolver):
    _, window = desktop(resolver)

    def broken(**kwargs):
        raise RuntimeError("COM disconnected")

    window.Control = broken
    assert error_code(resolver, {"WindowName": "Editor", "Name": "Save"}) == "action_failed"
    resolver["uiautomation"].GetRootControl = lambda: None
    assert error_code(resolver, {"WindowName": "Editor"}) == "action_failed"


@pytest.mark.parametrize("location", [
    {}, {"Xpath": "not json"}, {"Name": "Save", "ControlType": "Unknown"},
    {"Name": "Save", "foundIndex": "bad"}, {"Name": "Save", "foundIndex": -1},
])
def test_invalid_locations_are_classified(resolver, location):
    desktop(resolver)
    assert error_code(resolver, location) == "invalid_argument"


def test_context_routes_legacy_lookup_and_resets_after_failure(resolver):
    button = Control("Save")
    desktop(resolver, [button])
    location = {"WindowName": "Editor", "Name": "Save"}
    assert resolver["find_control"](location, strict=True) is button
    with resolver["strict_control_resolution"]():
        assert resolver["find_control"](location) is button
    assert resolver["_STRICT_CONTROL_RESOLUTION"].get() is False
    with pytest.raises(resolver["ActionError"]):
        with resolver["strict_control_resolution"]():
            resolver["find_control"]({"WindowName": "Missing"})
    assert resolver["_STRICT_CONTROL_RESOLUTION"].get() is False


def test_prepare_preserves_typed_failure_without_retry(resolver):
    tree = ast.parse((SOURCE / "ctrl.py").read_text(encoding="utf-8"))
    node = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "locate_and_prepare")
    exec(compile(ast.Module(body=[node], type_ignores=[]), str(SOURCE / "ctrl.py"), "exec"), resolver)
    calls = []

    def lookup(location):
        calls.append(location)
        raise resolver["ActionError"]("ambiguous_match", "Two candidates")

    resolver["find_control"] = lookup
    resolver["time"] = SimpleNamespace(time=lambda: 0, sleep=lambda _: pytest.fail("Must not retry"))
    with pytest.raises(resolver["ActionError"]) as raised:
        resolver["locate_and_prepare"]({"Name": "Save"}, {})
    assert raised.value.code == "ambiguous_match"
    assert len(calls) == 1


def test_only_mutation_scope_activates_resolved_window(resolver):
    button = Control("Save")
    desktop(resolver, [button])
    calls = []
    resolver['set_top_window'] = lambda title: calls.append(title) or True
    location = {'WindowName': 'Editor', 'Name': 'Save'}
    assert resolver['find_control'](location, strict=True) is button
    assert calls == []
    with resolver['strict_control_resolution'](activate=True):
        assert resolver['find_control'](location) is button
    assert calls == ['Editor']
    assert resolver['_ACTIVATE_RESOLVED_WINDOW'].get() is False
    resolver['set_top_window'] = lambda title: False
    with pytest.raises(resolver['ActionError']) as raised:
        with resolver['strict_control_resolution'](activate=True):
            resolver['find_control'](location)
    assert raised.value.code == 'window_invalid'
    assert resolver['_ACTIVATE_RESOLVED_WINDOW'].get() is False


def test_missing_control_is_never_activated(resolver):
    desktop(resolver)
    resolver['set_top_window'] = lambda title: pytest.fail('Missing control must not activate')
    with pytest.raises(resolver['ActionError']):
        with resolver['strict_control_resolution'](activate=True):
            resolver['find_control']({'WindowName': 'Editor', 'Name': 'Save'})
