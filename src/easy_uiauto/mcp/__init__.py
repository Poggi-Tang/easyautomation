"""Import MCP tools directly, without starting a transport or server loop.

Example::

    import json
    from easy_uiauto.mcp import get_control_tree, list_windows

    windows = json.loads(list_windows())
    tree = json.loads(get_control_tree(windows[0]["title"]))

Tools retain their MCP return types (usually JSON strings). Async tools must be
awaited; their optional ``ctx`` argument can be omitted in ordinary Python code.
"""

from __future__ import annotations

from importlib import import_module
from typing import TYPE_CHECKING

__all__ = [
    "activate_window",
    "analyze_ui_image",
    "capture_control_record_at_position",
    "capture_control_record_on_ctrl",
    "cleanup_ui_knowledge_images",
    "click",
    "click_at_position",
    "click_by_image",
    "click_by_vision",
    "click_text_on_screen",
    "correct_ui_image_annotations",
    "detect_visual_elements",
    "drag_control",
    "explore_ui_workflows",
    "find_control",
    "find_control_by_image",
    "find_control_by_vision",
    "find_text_on_screen",
    "get_capability_readiness",
    "get_control_at_position",
    "get_control_tree",
    "get_mode",
    "get_ui_learning_readiness",
    "get_ui_learning_status",
    "get_window_state",
    "hotkey",
    "learn_ui_command_effect",
    "list_ui_commands",
    "list_ui_interactions",
    "list_ui_knowledge_apps",
    "list_windows",
    "mouse_scroll",
    "move_mouse",
    "press_key",
    "rebuild_ui_knowledge_index",
    "revalidate_ui_control",
    "run_action",
    "run_actions",
    "run_ui_command",
    "run_ui_commands",
    "scan_window_knowledge",
    "scroll",
    "search_ui_knowledge",
    "search_ui_knowledge_batch",
    "set_mode",
    "set_text",
    "show_ui_controls",
    "take_screenshot",
    "take_window_screenshot",
    "teach_ui_control",
    "type_text",
    "upsert_control_record",
    "wait_for_control",
    "wait_for_text",
    "wait_for_window",
]

# Explicit type imports give editors the real signatures, including async tools.
# Runtime loading stays lazy so local helpers do not require the MCP extra.
if TYPE_CHECKING:
    from .server import (
        activate_window,
        analyze_ui_image,
        capture_control_record_at_position,
        capture_control_record_on_ctrl,
        cleanup_ui_knowledge_images,
        click,
        click_at_position,
        click_by_image,
        click_by_vision,
        click_text_on_screen,
        correct_ui_image_annotations,
        detect_visual_elements,
        drag_control,
        explore_ui_workflows,
        find_control,
        find_control_by_image,
        find_control_by_vision,
        find_text_on_screen,
        get_capability_readiness,
        get_control_at_position,
        get_control_tree,
        get_mode,
        get_ui_learning_readiness,
        get_ui_learning_status,
        get_window_state,
        hotkey,
        learn_ui_command_effect,
        list_ui_commands,
        list_ui_interactions,
        list_ui_knowledge_apps,
        list_windows,
        mouse_scroll,
        move_mouse,
        press_key,
        rebuild_ui_knowledge_index,
        revalidate_ui_control,
        run_action,
        run_actions,
        run_ui_command,
        run_ui_commands,
        scan_window_knowledge,
        scroll,
        search_ui_knowledge,
        search_ui_knowledge_batch,
        set_mode,
        set_text,
        show_ui_controls,
        take_screenshot,
        take_window_screenshot,
        teach_ui_control,
        type_text,
        upsert_control_record,
        wait_for_control,
        wait_for_text,
        wait_for_window,
    )


def __getattr__(name: str):
    if name not in __all__:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    try:
        server = import_module(".server", __name__)
    except ModuleNotFoundError as error:
        if error.name == "mcp" or (error.name or "").startswith("mcp."):
            raise ImportError(
                'MCP tools require: python -m pip install "easy-uiauto[mcp]"'
            ) from error
        raise
    value = getattr(server, name)
    globals()[name] = value
    return value


def __dir__():
    return sorted(set(globals()) | set(__all__))
