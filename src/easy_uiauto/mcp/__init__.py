"""MCP server support for easy-uiauto."""

from .backend import UIAutomationBackend
from .executor import UIAutomationExecutor
from .models import ToolResult
from .policy import MCPPolicy

__all__ = ["MCPPolicy", "ToolResult", "UIAutomationBackend", "UIAutomationExecutor"]
