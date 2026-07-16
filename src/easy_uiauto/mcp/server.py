from __future__ import annotations

import atexit
import logging
from typing import Any

from .backend import UIAutomationBackend
from .executor import UIAutomationExecutor
from .models import ToolResult
from .policy import MCPPolicy

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as exc:  # pragma: no cover - exercised by installed entry point
    raise RuntimeError("Install easy-uiauto[mcp] to use the MCP server") from exc


mcp = FastMCP("easy-uiauto")
executor = UIAutomationExecutor(
    lambda: UIAutomationBackend(policy=MCPPolicy.from_environment())
)
atexit.register(executor.close)


def _call(method: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
    try:
        result: ToolResult = executor.call(method, *args, **kwargs)
    except Exception as exc:
        result = ToolResult.error(f"UIAutomation worker error: {exc}")
    return result.to_dict()


@mcp.tool()
def list_windows(limit: int = 100) -> dict[str, Any]:
    """List top-level Windows UI Automation windows as JSON snapshots."""
    return _call("list_windows", limit)


@mcp.tool()
def find_control(location: dict[str, Any]) -> dict[str, Any]:
    """Find a control using an easy-uiauto LOCATION object and return a short-lived ref."""
    return _call("find", location)


@mcp.tool()
def inspect_control(reference: str) -> dict[str, Any]:
    """Refresh a short-lived control reference and return its current JSON snapshot."""
    return _call("inspect", reference)


@mcp.tool()
def list_children(reference: str, offset: int = 0, limit: int = 50) -> dict[str, Any]:
    """List direct children of a control reference with bounded pagination."""
    return _call("list_children", reference, offset, limit)


@mcp.tool()
def get_control_tree(
    reference: str,
    max_depth: int = 3,
    max_nodes: int = 200,
) -> dict[str, Any]:
    """Serialize a bounded UI Automation subtree without returning COM objects."""
    return _call("get_control_tree", reference, max_depth, max_nodes)


@mcp.tool()
def cache_stats() -> dict[str, Any]:
    """Return core cache and MCP reference statistics."""
    return _call("cache_stats")


@mcp.tool()
def clear_caches() -> dict[str, Any]:
    """Clear cached UI Automation controls and short-lived MCP references."""
    return _call("clear_caches")


@mcp.tool()
def invalidate_control(reference: str) -> dict[str, Any]:
    """Explicitly invalidate one MCP control reference."""
    return _call("invalidate_reference", reference)


@mcp.tool()
def start_recording() -> dict[str, Any]:
    """Start a global input recording session when recording is explicitly enabled."""
    return _call("start_recording", timeout=15)


@mcp.tool()
def recording_status(
    session_id: str, include_actions: bool = True
) -> dict[str, Any]:
    """Return recording session state and optionally its pure-JSON actions."""
    return _call("recording_status", session_id, include_actions)


@mcp.tool()
def stop_recording(session_id: str) -> dict[str, Any]:
    """Stop a recording session and return its final actions and cleanup status."""
    return _call("stop_recording", session_id, timeout=15)


@mcp.tool()
def start_highlight(
    reference: str,
    color: str = "#FF0000",
    line_width: int = 2,
    alpha: float = 1.0,
) -> dict[str, Any]:
    """Start a click-through highlight session for a control reference."""
    return _call("start_highlight", reference, color, line_width, alpha, timeout=15)


@mcp.tool()
def update_highlight(
    session_id: str,
    reference: str | None = None,
    color: str | None = None,
    line_width: int | None = None,
    alpha: float | None = None,
) -> dict[str, Any]:
    """Move or restyle an existing highlight session."""
    return _call(
        "update_highlight",
        session_id,
        reference=reference,
        color=color,
        line_width=line_width,
        alpha=alpha,
    )


@mcp.tool()
def stop_highlight(session_id: str) -> dict[str, Any]:
    """Stop and destroy a highlight session."""
    return _call("stop_highlight", session_id, timeout=15)


@mcp.tool()
def perform_action(
    action: str,
    location: dict[str, Any] | None = None,
    reference: str | None = None,
    parameters: dict[str, Any] | None = None,
    observe_location: dict[str, Any] | None = None,
    observe_reference: str | None = None,
    expected: dict[str, Any] | None = None,
    dry_run: bool = False,
    confirm_high_impact: bool = False,
) -> dict[str, Any]:
    """Execute an easy-uiauto action on a LOCATION object or control reference."""
    return _call(
        "perform_action",
        action,
        location=location,
        reference=reference,
        parameters=parameters,
        observe_location=observe_location,
        observe_reference=observe_reference,
        expected=expected,
        dry_run=dry_run,
        confirm_high_impact=confirm_high_impact,
    )


def main() -> None:
    logging.getLogger("mcp").setLevel(logging.WARNING)
    try:
        mcp.run(transport="stdio")
    finally:
        executor.close()
