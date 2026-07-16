from __future__ import annotations

import time
import uuid
from dataclasses import dataclass
from typing import Any

import uiautomation

from easy_uiauto.ctrl import get_message_type, run_action
from easy_uiauto.utils import find_control, get_control_xpath

from .models import ToolResult
from .policy import MCPPolicy


@dataclass(slots=True)
class _ControlReference:
    control: Any
    xpath: list[dict[str, Any]]
    fingerprint: tuple[Any, ...]
    expires_at: float


class UIAutomationBackend:
    def __init__(self, policy: MCPPolicy | None = None, *, reference_ttl: float = 30):
        self.policy = policy or MCPPolicy.from_environment()
        self.reference_ttl = max(1.0, float(reference_ttl))
        self._references: dict[str, _ControlReference] = {}

    @staticmethod
    def _fingerprint(control: Any) -> tuple[Any, ...]:
        return tuple(
            getattr(control, name, None)
            for name in (
                "ControlTypeName",
                "Name",
                "ClassName",
                "AutomationId",
                "NativeWindowHandle",
                "FrameworkId",
            )
        )

    @staticmethod
    def _rect(control: Any) -> dict[str, int] | None:
        try:
            rect = control.BoundingRectangle
            return {
                "left": int(rect.left),
                "top": int(rect.top),
                "right": int(rect.right),
                "bottom": int(rect.bottom),
            }
        except Exception:
            return None

    @staticmethod
    def _pattern_state(control: Any) -> dict[str, Any]:
        state: dict[str, Any] = {}
        readers = (
            (
                "value",
                uiautomation.PatternId.ValuePattern,
                lambda pattern: {
                    "value": pattern.Value,
                    "is_read_only": bool(pattern.IsReadOnly),
                },
            ),
            (
                "toggle",
                uiautomation.PatternId.TogglePattern,
                lambda pattern: {"state": int(pattern.ToggleState)},
            ),
            (
                "selection_item",
                uiautomation.PatternId.SelectionItemPattern,
                lambda pattern: {"is_selected": bool(pattern.IsSelected)},
            ),
            (
                "expand_collapse",
                uiautomation.PatternId.ExpandCollapsePattern,
                lambda pattern: {"state": int(pattern.ExpandCollapseState)},
            ),
            (
                "grid",
                uiautomation.PatternId.GridPattern,
                lambda pattern: {
                    "row_count": int(pattern.RowCount),
                    "column_count": int(pattern.ColumnCount),
                },
            ),
        )
        for name, pattern_id, reader in readers:
            try:
                pattern = control.GetPattern(pattern_id)
                if pattern is not None:
                    state[name] = reader(pattern)
            except Exception:
                continue
        return state

    def _store_reference(self, control: Any, xpath: list[dict[str, Any]]) -> str:
        now = time.monotonic()
        self._references = {
            key: value for key, value in self._references.items() if value.expires_at >= now
        }
        if len(self._references) >= 1000:
            oldest = min(self._references, key=lambda key: self._references[key].expires_at)
            self._references.pop(oldest, None)
        reference = uuid.uuid4().hex
        self._references[reference] = _ControlReference(
            control=control,
            xpath=xpath,
            fingerprint=self._fingerprint(control),
            expires_at=now + self.reference_ttl,
        )
        return reference

    def _resolve_reference(self, reference: str) -> _ControlReference:
        record = self._references.get(reference)
        if record is None:
            raise KeyError("unknown control reference")
        if time.monotonic() > record.expires_at:
            self._references.pop(reference, None)
            raise KeyError("expired control reference")
        try:
            valid = bool(record.control.Exists(0))
        except Exception:
            valid = False
        if not valid or self._fingerprint(record.control) != record.fingerprint:
            self._references.pop(reference, None)
            raise KeyError("stale control reference")
        record.expires_at = time.monotonic() + self.reference_ttl
        return record

    def _snapshot(self, control: Any, *, include_xpath: bool = True) -> dict[str, Any]:
        xpath = get_control_xpath(control) if include_xpath else []
        reference = self._store_reference(control, xpath)
        return {
            "ref": reference,
            "name": getattr(control, "Name", ""),
            "class_name": getattr(control, "ClassName", ""),
            "control_type": getattr(control, "ControlTypeName", ""),
            "automation_id": getattr(control, "AutomationId", ""),
            "framework_id": getattr(control, "FrameworkId", ""),
            "native_window_handle": int(getattr(control, "NativeWindowHandle", 0) or 0),
            "is_enabled": bool(getattr(control, "IsEnabled", True)),
            "rect": self._rect(control),
            "xpath": xpath,
            "patterns": self._pattern_state(control),
        }

    def list_windows(self, limit: int = 100) -> ToolResult:
        root = uiautomation.GetRootControl()
        windows = []
        for control in root.GetChildren():
            if getattr(control, "ControlTypeName", "") not in {"WindowControl", "PaneControl"}:
                continue
            if not int(getattr(control, "NativeWindowHandle", 0) or 0):
                continue
            windows.append(self._snapshot(control))
            if len(windows) >= max(1, min(int(limit), 500)):
                break
        return ToolResult.success(windows, f"found {len(windows)} top-level windows")

    def find(self, location: dict[str, Any]) -> ToolResult:
        if not isinstance(location, dict):
            return ToolResult.error("location must be an object")
        control = find_control(location)
        if control is None:
            return ToolResult.error("control not found")
        return ToolResult.success(self._snapshot(control), "control found")

    def inspect(self, reference: str) -> ToolResult:
        try:
            record = self._resolve_reference(reference)
        except KeyError as exc:
            return ToolResult.error(str(exc))
        return ToolResult.success(self._snapshot(record.control), "control inspected")

    @staticmethod
    def _location_from_reference(record: _ControlReference) -> dict[str, Any]:
        target = record.xpath[-1] if record.xpath else {}
        window = record.xpath[0] if record.xpath else {}
        return {
            "WindowName": window.get("Name", ""),
            "Name": target.get("Name", ""),
            "ClassName": target.get("ClassName", ""),
            "ControlType": target.get("ControlType", ""),
            "foundIndex": target.get("foundIndex", 1),
            "AutomationId": target.get("AutomationId", ""),
            "Xpath": record.xpath,
            "Img": "",
            "PARAMETERS": {},
        }

    def perform_action(
        self,
        action: str,
        *,
        location: dict[str, Any] | None = None,
        reference: str | None = None,
        parameters: dict[str, Any] | None = None,
    ) -> ToolResult:
        if reference:
            try:
                location = self._location_from_reference(self._resolve_reference(reference))
            except KeyError as exc:
                return ToolResult.error(str(exc))
        if not isinstance(location, dict):
            return ToolResult.error("location or reference is required")
        normalized_location = dict(location)
        normalized_location["PARAMETERS"] = dict(parameters or location.get("PARAMETERS") or {})
        violation = self.policy.validate_action(action, normalized_location)
        if violation is not None:
            return violation
        message = run_action({"ACTION": action, "LOCATION": normalized_location})
        message_type = get_message_type()
        if message_type == 0:
            return ToolResult.error(message)
        if message_type == 1:
            return ToolResult.warning(message=message)
        return ToolResult.success(message=message)

    def close(self) -> None:
        self._references.clear()
