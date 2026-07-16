from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any

from .models import ToolResult

HIGH_RISK_ACTIONS = {
    "按下鼠标左键",
    "释放鼠标左键",
    "按下鼠标右键",
    "释放鼠标右键",
    "拖拽",
    "键盘按下",
    "键盘释放",
    "组合键",
}


def _enabled(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True, slots=True)
class MCPPolicy:
    allow_high_risk: bool = False
    allow_image_paths: bool = False

    @classmethod
    def from_environment(cls) -> MCPPolicy:
        return cls(
            allow_high_risk=_enabled("EASY_UIAUTO_MCP_ALLOW_HIGH_RISK"),
            allow_image_paths=_enabled("EASY_UIAUTO_MCP_ALLOW_IMAGE_PATHS"),
        )

    def validate_action(self, action: str, location: dict[str, Any]) -> ToolResult | None:
        if action in HIGH_RISK_ACTIONS and not self.allow_high_risk:
            return ToolResult.error(
                f"动作 {action!r} 默认被 MCP 策略阻止；"
                "设置 EASY_UIAUTO_MCP_ALLOW_HIGH_RISK=1 后才允许"
            )
        if location.get("Img") and not self.allow_image_paths:
            return ToolResult.error(
                "MCP 默认禁止读取图片路径；"
                "设置 EASY_UIAUTO_MCP_ALLOW_IMAGE_PATHS=1 后才允许"
            )
        return None
