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
def perform_action(
    action: str,
    location: dict[str, Any] | None = None,
    reference: str | None = None,
    parameters: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Execute an easy-uiauto action on a LOCATION object or control reference."""
    return _call(
        "perform_action",
        action,
        location=location,
        reference=reference,
        parameters=parameters,
    )


def main() -> None:
    logging.getLogger("mcp").setLevel(logging.WARNING)
    try:
        mcp.run(transport="stdio")
    finally:
        executor.close()
