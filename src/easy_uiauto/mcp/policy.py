from __future__ import annotations

import os
import re
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
HIGH_IMPACT_ENGLISH_TERMS = {
    "run",
    "simulate",
    "compile",
    "build",
    "delete",
    "clear",
    "close",
    "exit",
    "import",
}
HIGH_IMPACT_CHINESE_TERMS = {
    "运行",
    "仿真",
    "编译",
    "构建",
    "删除",
    "清空",
    "关闭",
    "退出",
    "导入",
}


def _enabled(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True, slots=True)
class MCPPolicy:
    allow_high_risk: bool = False
    allow_image_paths: bool = False
    allow_high_impact: bool = False
    allow_recording: bool = False

    @classmethod
    def from_environment(cls) -> MCPPolicy:
        return cls(
            allow_high_risk=_enabled("EASY_UIAUTO_MCP_ALLOW_HIGH_RISK"),
            allow_image_paths=_enabled("EASY_UIAUTO_MCP_ALLOW_IMAGE_PATHS"),
            allow_high_impact=_enabled("EASY_UIAUTO_MCP_ALLOW_HIGH_IMPACT"),
            allow_recording=_enabled("EASY_UIAUTO_MCP_ALLOW_RECORDING"),
        )

    def validate_action(
        self,
        action: str,
        location: dict[str, Any],
        *,
        confirm_high_impact: bool = False,
    ) -> ToolResult | None:
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
        if (
            not self.allow_high_impact
            and not confirm_high_impact
            and self.is_high_impact_target(location)
        ):
            return ToolResult.error(
                "目标看起来可能触发运行、删除、清空、关闭或导入等高影响操作；"
                "请在确认后传入 confirm_high_impact=true，或设置 "
                "EASY_UIAUTO_MCP_ALLOW_HIGH_IMPACT=1"
            )
        return None

    @staticmethod
    def is_high_impact_target(location: dict[str, Any]) -> bool:
        target_text = " ".join(
            str(location.get(field, "")) for field in ("Name", "AutomationId", "ClassName")
        )
        tokenized = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", target_text)
        english_tokens = set(re.findall(r"[a-z]+", tokenized.lower()))
        return bool(
            english_tokens.intersection(HIGH_IMPACT_ENGLISH_TERMS)
            or any(term in target_text for term in HIGH_IMPACT_CHINESE_TERMS)
        )
