from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass(slots=True)
class ToolResult:
    ok: bool
    status: str
    message: str = ""
    data: Any = None
    warnings: list[str] = field(default_factory=list)

    @classmethod
    def success(cls, data: Any = None, message: str = "") -> ToolResult:
        return cls(ok=True, status="ok", message=message, data=data)

    @classmethod
    def warning(cls, data: Any = None, message: str = "") -> ToolResult:
        return cls(ok=True, status="warning", message=message, data=data)

    @classmethod
    def error(cls, message: str, data: Any = None) -> ToolResult:
        return cls(ok=False, status="error", message=message, data=data)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
