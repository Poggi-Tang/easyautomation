from __future__ import annotations

import time
import uuid
from dataclasses import dataclass
from typing import Any

import uiautomation

from easy_uiauto.ctrl import (
    _coerce_optional_bool,
    get_action_mechanism,
    get_message_type,
    run_action,
)
from easy_uiauto.draw import get_visible_rect_map_by_control
from easy_uiauto.utils import (
    clear_control_cache,
    find_control,
    get_control_cache_stats,
    get_control_xpath,
)

from .models import ToolResult
from .policy import MCPPolicy
from .sessions import HighlightSession, RecordingSession


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
        self._recordings: dict[str, RecordingSession] = {}
        self._highlights: dict[str, HighlightSession] = {}

    @staticmethod
    def _safe_attr(control: Any, name: str, default: Any = "") -> Any:
        try:
            return getattr(control, name, default)
        except Exception:
            return default

    @classmethod
    def _fingerprint(cls, control: Any) -> tuple[Any, ...]:
        try:
            runtime_id = tuple(control.GetRuntimeId() or ())
        except Exception:
            runtime_id = ()
        stable = (
            cls._safe_attr(control, "ProcessId", 0),
            runtime_id,
            cls._safe_attr(control, "NativeWindowHandle", 0),
            cls._safe_attr(control, "FrameworkId", ""),
            cls._safe_attr(control, "ControlTypeName", ""),
        )
        if runtime_id or stable[2]:
            return stable
        return stable + (
            cls._safe_attr(control, "Name", ""),
            cls._safe_attr(control, "ClassName", ""),
            cls._safe_attr(control, "AutomationId", ""),
        )

    @classmethod
    def _reference_matches_control(cls, record: _ControlReference, control: Any) -> bool:
        if record.control is control:
            return True
        fingerprint = cls._fingerprint(control)
        if fingerprint != record.fingerprint:
            return False
        runtime_id = record.fingerprint[1]
        native_handle = record.fingerprint[2]
        control_type = record.fingerprint[4]
        return bool(runtime_id or (native_handle and control_type in {"WindowControl", "PaneControl"}))

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
                "invoke",
                uiautomation.PatternId.InvokePattern,
                lambda _pattern: {"supported": True},
            ),
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
            (
                "selection",
                uiautomation.PatternId.SelectionPattern,
                lambda pattern: {
                    "can_select_multiple": bool(pattern.CanSelectMultiple),
                    "is_selection_required": bool(pattern.IsSelectionRequired),
                    "selected_count": len(pattern.GetSelection()),
                },
            ),
            (
                "scroll",
                uiautomation.PatternId.ScrollPattern,
                lambda pattern: {
                    "horizontally_scrollable": bool(pattern.HorizontallyScrollable),
                    "vertically_scrollable": bool(pattern.VerticallyScrollable),
                    "horizontal_percent": float(pattern.HorizontalScrollPercent),
                    "vertical_percent": float(pattern.VerticalScrollPercent),
                },
            ),
            (
                "range_value",
                uiautomation.PatternId.RangeValuePattern,
                lambda pattern: {
                    "value": float(pattern.Value),
                    "minimum": float(pattern.Minimum),
                    "maximum": float(pattern.Maximum),
                    "is_read_only": bool(pattern.IsReadOnly),
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

    def _snapshot(
        self,
        control: Any,
        *,
        include_xpath: bool = True,
        reference: str | None = None,
    ) -> dict[str, Any]:
        try:
            xpath = get_control_xpath(control) if include_xpath else []
        except Exception:
            xpath = []
        if reference:
            record = self._references.get(reference)
            if record is None or not self._reference_matches_control(record, control):
                reference = None
            elif xpath:
                record.xpath = xpath
        reference = reference or self._store_reference(control, xpath)
        return {
            "ref": reference,
            "name": self._safe_attr(control, "Name", ""),
            "class_name": self._safe_attr(control, "ClassName", ""),
            "control_type": self._safe_attr(control, "ControlTypeName", ""),
            "automation_id": self._safe_attr(control, "AutomationId", ""),
            "framework_id": self._safe_attr(control, "FrameworkId", ""),
            "process_id": int(self._safe_attr(control, "ProcessId", 0) or 0),
            "native_window_handle": int(
                self._safe_attr(control, "NativeWindowHandle", 0) or 0
            ),
            "is_enabled": bool(self._safe_attr(control, "IsEnabled", True)),
            "is_offscreen": bool(self._safe_attr(control, "IsOffscreen", False)),
            "has_keyboard_focus": bool(
                self._safe_attr(control, "HasKeyboardFocus", False)
            ),
            "is_keyboard_focusable": bool(
                self._safe_attr(control, "IsKeyboardFocusable", False)
            ),
            "rect": self._rect(control),
            "xpath": xpath,
            "patterns": self._pattern_state(control),
        }

    def list_windows(self, limit: int = 100) -> ToolResult:
        try:
            root = uiautomation.GetRootControl()
            controls = root.GetChildren()
        except Exception as exc:
            return ToolResult.error(f"top-level window enumeration failed: {exc}")
        windows = []
        warnings = []
        for control in controls:
            if self._safe_attr(control, "ControlTypeName", "") not in {
                "WindowControl",
                "PaneControl",
            }:
                continue
            if not int(self._safe_attr(control, "NativeWindowHandle", 0) or 0):
                continue
            try:
                windows.append(self._snapshot(control, include_xpath=False))
            except Exception as exc:
                warnings.append(str(exc))
            if len(windows) >= max(1, min(int(limit), 500)):
                break
        result = ToolResult.success(windows, f"found {len(windows)} top-level windows")
        result.warnings.extend(warnings)
        return result

    def find(self, location: dict[str, Any]) -> ToolResult:
        if not isinstance(location, dict):
            return ToolResult.error("location must be an object")
        try:
            control = find_control(location)
        except Exception as exc:
            return ToolResult.error(f"control lookup failed: {exc}")
        if control is None:
            return ToolResult.error("control not found")
        try:
            snapshot = self._snapshot(control)
        except Exception as exc:
            return ToolResult.error(f"control snapshot failed: {exc}")
        return ToolResult.success(snapshot, "control found")

    def inspect(self, reference: str) -> ToolResult:
        try:
            record = self._resolve_reference(reference)
        except KeyError as exc:
            return ToolResult.error(str(exc))
        return ToolResult.success(
            self._snapshot(record.control, reference=reference), "control inspected"
        )

    def list_children(self, reference: str, offset: int = 0, limit: int = 50) -> ToolResult:
        try:
            record = self._resolve_reference(reference)
            children = record.control.GetChildren()
        except Exception as exc:
            return ToolResult.error(str(exc))
        start = max(0, int(offset))
        page_size = max(1, min(int(limit), 200))
        snapshots = []
        warnings = []
        for child in children[start:start + page_size]:
            try:
                snapshots.append(self._snapshot(child, include_xpath=False))
            except Exception as exc:
                warnings.append(str(exc))
        result = ToolResult.success(
            {
                "items": snapshots,
                "offset": start,
                "limit": page_size,
                "total": len(children),
                "has_more": start + page_size < len(children),
            },
            f"returned {len(snapshots)} children",
        )
        result.warnings.extend(warnings)
        return result

    def get_control_tree(
        self,
        reference: str,
        max_depth: int = 3,
        max_nodes: int = 200,
    ) -> ToolResult:
        try:
            root = self._resolve_reference(reference).control
        except KeyError as exc:
            return ToolResult.error(str(exc))
        depth_limit = max(0, min(int(max_depth), 10))
        node_limit = max(1, min(int(max_nodes), 1000))
        count = 0
        truncated = False
        warnings = []

        def build(control: Any, depth: int) -> dict[str, Any] | None:
            nonlocal count, truncated
            if count >= node_limit:
                truncated = True
                return None
            try:
                node = self._snapshot(control, include_xpath=False)
            except Exception as exc:
                warnings.append(str(exc))
                return None
            count += 1
            node["children"] = []
            if depth >= depth_limit:
                return node
            try:
                children = control.GetChildren()
            except Exception as exc:
                warnings.append(str(exc))
                return node
            for child in children:
                child_node = build(child, depth + 1)
                if child_node is not None:
                    node["children"].append(child_node)
                if count >= node_limit:
                    truncated = True
                    break
            return node

        tree = build(root, 0)
        result = ToolResult.success(
            {"root": tree, "node_count": count, "truncated": truncated},
            f"serialized {count} controls",
        )
        result.warnings.extend(warnings)
        return result

    def cache_stats(self) -> ToolResult:
        data = get_control_cache_stats()
        data["mcp_references"] = len(self._references)
        return ToolResult.success(data)

    def clear_caches(self) -> ToolResult:
        control_entries = clear_control_cache()
        reference_entries = len(self._references)
        self._references.clear()
        return ToolResult.success(
            {
                "control_cache_entries_removed": control_entries,
                "mcp_references_removed": reference_entries,
            },
            "caches cleared",
        )

    def invalidate_reference(self, reference: str) -> ToolResult:
        removed = self._references.pop(reference, None) is not None
        if not removed:
            return ToolResult.error("unknown control reference")
        return ToolResult.success(message="control reference invalidated")

    def start_recording(self) -> ToolResult:
        if not self.policy.allow_recording:
            return ToolResult.error(
                "MCP 录制默认禁用；设置 EASY_UIAUTO_MCP_ALLOW_RECORDING=1 后启用"
            )
        try:
            session = RecordingSession.start()
        except Exception as exc:
            return ToolResult.error(f"recording failed to start: {exc}")
        self._recordings[session.session_id] = session
        return ToolResult.success(session.status(include_actions=False), "recording started")

    def recording_status(
        self, session_id: str, include_actions: bool = True
    ) -> ToolResult:
        session = self._recordings.get(session_id)
        if session is None:
            return ToolResult.error("unknown recording session")
        return ToolResult.success(session.status(include_actions=include_actions))

    def stop_recording(self, session_id: str) -> ToolResult:
        session = self._recordings.pop(session_id, None)
        if session is None:
            return ToolResult.error("unknown recording session")
        status = session.stop()
        if status["cleanup_errors"]:
            return ToolResult.warning(status, "recording stopped with cleanup warnings")
        return ToolResult.success(status, "recording stopped")

    def start_highlight(
        self,
        reference: str,
        color: str = "#FF0000",
        line_width: int = 2,
        alpha: float = 1.0,
    ) -> ToolResult:
        try:
            control = self._resolve_reference(reference).control
            rect = get_visible_rect_map_by_control(control)
            if not rect or rect.get("width", 0) <= 0 or rect.get("height", 0) <= 0:
                return ToolResult.error("control has no visible highlight rectangle")
            session = HighlightSession(
                rect,
                color=color,
                line_width=line_width,
                alpha=alpha,
            )
            session.start()
        except Exception as exc:
            return ToolResult.error(f"highlight failed to start: {exc}")
        self._highlights[session.session_id] = session
        return ToolResult.success(session.status(), "highlight started")

    def update_highlight(
        self,
        session_id: str,
        *,
        reference: str | None = None,
        color: str | None = None,
        line_width: int | None = None,
        alpha: float | None = None,
    ) -> ToolResult:
        session = self._highlights.get(session_id)
        if session is None:
            return ToolResult.error("unknown highlight session")
        try:
            if reference:
                control = self._resolve_reference(reference).control
                rect = get_visible_rect_map_by_control(control)
                if not rect:
                    return ToolResult.error("control has no visible highlight rectangle")
                session.update_rect(rect)
            if any(value is not None for value in (color, line_width, alpha)):
                session.update_style(color=color, line_width=line_width, alpha=alpha)
        except Exception as exc:
            return ToolResult.error(f"highlight update failed: {exc}")
        return ToolResult.success(session.status(), "highlight updated")

    def stop_highlight(self, session_id: str) -> ToolResult:
        session = self._highlights.pop(session_id, None)
        if session is None:
            return ToolResult.error("unknown highlight session")
        status = session.stop()
        if status["running"] or status.get("forced_termination") or status.get("error"):
            return ToolResult.warning(status, "highlight stopped with cleanup warnings")
        return ToolResult.success(status, "highlight stopped")

    @staticmethod
    def _location_from_reference(record: _ControlReference) -> dict[str, Any]:
        if not record.xpath:
            try:
                record.xpath = get_control_xpath(record.control)
            except Exception as exc:
                raise ValueError("control reference XPath could not be resolved") from exc
        if not record.xpath:
            raise ValueError("control reference has no actionable XPath")
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

    @staticmethod
    def _nested_value(data: dict[str, Any] | None, path: str) -> Any:
        current: Any = data
        for part in path.split("."):
            if not isinstance(current, dict) or part not in current:
                raise KeyError(path)
            current = current[part]
        return current

    @classmethod
    def _verify_action(
        cls,
        action: str,
        parameters: dict[str, Any],
        target_after: dict[str, Any] | None,
        observed_before: dict[str, Any] | None,
        observed_after: dict[str, Any] | None,
        expected: dict[str, Any] | None,
    ) -> tuple[bool | None, str | None]:
        verification_target = observed_after or target_after
        if expected:
            path = str(expected.get("path", ""))
            if not path:
                return False, "expected.path is required"
            try:
                actual = cls._nested_value(verification_target, path)
            except KeyError:
                return False, f"verification path not found: {path}"
            wanted = expected.get("equals")
            return actual == wanted, None if actual == wanted else f"{path}={actual!r}"
        try:
            if action in {"输入文本", "设置文本"}:
                key = "输入文本" if action == "输入文本" else "设置文本"
                actual = cls._nested_value(target_after, "patterns.value.value")
                wanted = str(parameters[key])
                return actual == wanted, None if actual == wanted else f"value={actual!r}"
            if action == "切换状态" and "选中" in parameters:
                actual = int(cls._nested_value(target_after, "patterns.toggle.state"))
                desired = _coerce_optional_bool(parameters["选中"], "选中")
                if desired is None:
                    return None, "toggle action did not specify a target state"
                wanted = 1 if desired else 0
                return actual == wanted, None if actual == wanted else f"toggle={actual}"
            if action == "展开折叠":
                actual = int(
                    cls._nested_value(target_after, "patterns.expand_collapse.state")
                )
                desired = _coerce_optional_bool(parameters.get("展开", True), "展开")
                wanted = 1 if desired is not False else 0
                return actual == wanted, None if actual == wanted else f"expand={actual}"
            if action == "选择":
                if parameters.get("选择项") is not None:
                    actual = cls._nested_value(target_after, "patterns.value.value")
                    wanted = str(parameters["选择项"])
                    return actual == wanted, None if actual == wanted else f"value={actual!r}"
                actual = bool(
                    cls._nested_value(target_after, "patterns.selection_item.is_selected")
                )
                return actual, None if actual else "selection item is not selected"
        except (KeyError, TypeError, ValueError):
            pass
        if observed_before is not None and observed_after is not None:
            before = {
                "name": observed_before.get("name"),
                "patterns": observed_before.get("patterns"),
                "rect": observed_before.get("rect"),
            }
            after = {
                "name": observed_after.get("name"),
                "patterns": observed_after.get("patterns"),
                "rect": observed_after.get("rect"),
            }
            changed = before != after
            return changed, None if changed else "observed control state did not change"
        return None, "no postcondition was supplied for this action"

    @staticmethod
    def _mechanism(action: str, before: dict[str, Any] | None) -> str:
        patterns = (before or {}).get("patterns", {})
        if action in {"输入文本", "设置文本"} and "value" in patterns:
            return "value_pattern_or_fallback"
        if action == "切换状态" and "toggle" in patterns:
            return "toggle_pattern"
        if action == "选择":
            return "selection_pattern"
        if action == "展开折叠" and "expand_collapse" in patterns:
            return "expand_collapse_pattern"
        if action == "点击" and "invoke" in patterns:
            return "invoke_pattern"
        return "provider_or_physical_fallback"

    def perform_action(
        self,
        action: str,
        *,
        location: dict[str, Any] | None = None,
        reference: str | None = None,
        parameters: dict[str, Any] | None = None,
        observe_location: dict[str, Any] | None = None,
        observe_reference: str | None = None,
        expected: dict[str, Any] | None = None,
        dry_run: bool = False,
        confirm_high_impact: bool = False,
    ) -> ToolResult:
        if reference:
            try:
                location = self._location_from_reference(self._resolve_reference(reference))
            except (KeyError, ValueError) as exc:
                return ToolResult.error(str(exc))
        if not isinstance(location, dict):
            return ToolResult.error("location or reference is required")
        normalized_location = dict(location)
        normalized_location["PARAMETERS"] = dict(parameters or location.get("PARAMETERS") or {})
        requires_high_impact_confirmation = (
            not self.policy.allow_high_impact
            and self.policy.is_high_impact_target(normalized_location)
        )
        violation = self.policy.validate_action(
            action,
            normalized_location,
            confirm_high_impact=confirm_high_impact or dry_run,
        )
        if violation is not None:
            return violation
        try:
            target_control = find_control(normalized_location)
        except Exception as exc:
            return ToolResult.error(f"pre-action target lookup failed: {exc}")
        target_before = (
            self._snapshot(target_control, reference=reference)
            if target_control is not None
            else None
        )
        observed_control = None
        if observe_reference:
            try:
                observed_control = self._resolve_reference(observe_reference).control
            except KeyError as exc:
                return ToolResult.error(str(exc))
        elif observe_location:
            try:
                observed_control = find_control(observe_location)
            except Exception as exc:
                return ToolResult.error(f"pre-action observed control lookup failed: {exc}")
        observed_before = (
            self._snapshot(observed_control, reference=observe_reference)
            if observed_control is not None
            else None
        )
        if dry_run:
            return ToolResult.success(
                {
                    "action": action,
                    "location": normalized_location,
                    "target": target_before,
                    "observed": observed_before,
                    "would_require_high_impact_confirmation": (
                        requires_high_impact_confirmation
                    ),
                },
                "dry run; no action executed",
            )
        try:
            message = run_action({"ACTION": action, "LOCATION": normalized_location})
            message_type = get_message_type()
            action_mechanism = get_action_mechanism() or self._mechanism(
                action, target_before
            )
        except Exception as exc:
            return ToolResult.error(
                f"action execution failed: {exc}",
                {
                    "before": target_before,
                    "mechanism": self._mechanism(action, target_before),
                    "verified": False,
                },
            )
        if message_type == 0:
            return ToolResult.error(
                message,
                {
                    "before": target_before,
                    "mechanism": action_mechanism,
                    "verified": False,
                },
            )
        post_action_warnings = []
        try:
            target_after_control = find_control(normalized_location)
        except Exception as exc:
            target_after_control = None
            post_action_warnings.append(f"post-action target lookup failed: {exc}")
        target_after = (
            self._snapshot(target_after_control, reference=reference)
            if target_after_control is not None
            else None
        )
        if observe_reference:
            try:
                observed_after_control = self._resolve_reference(observe_reference).control
            except KeyError as exc:
                observed_after_control = None
                post_action_warnings.append(f"post-action observed reference failed: {exc}")
        elif observe_location:
            try:
                observed_after_control = find_control(observe_location)
            except Exception as exc:
                observed_after_control = None
                post_action_warnings.append(
                    f"post-action observed control lookup failed: {exc}"
                )
        else:
            observed_after_control = None
        observed_after = (
            self._snapshot(observed_after_control, reference=observe_reference)
            if observed_after_control is not None
            else None
        )
        verified, verification_error = self._verify_action(
            action,
            normalized_location["PARAMETERS"],
            target_after,
            observed_before,
            observed_after,
            expected,
        )
        data = {
            "before": target_before,
            "after": target_after,
            "observed_before": observed_before,
            "observed_after": observed_after,
            "mechanism": action_mechanism,
            "verified": verified,
            "verification_error": verification_error,
        }
        if verified is False:
            result = ToolResult.error(
                f"{message}; postcondition failed: {verification_error}", data
            )
        elif message_type == 1 or verified is None or post_action_warnings:
            result = ToolResult.warning(data=data, message=message)
        else:
            result = ToolResult.success(data=data, message=message)
        result.warnings.extend(post_action_warnings)
        return result

    def close(self) -> None:
        for recording_session in list(self._recordings.values()):
            try:
                recording_session.stop()
            except Exception:
                pass
        self._recordings.clear()
        for highlight_session in list(self._highlights.values()):
            try:
                highlight_session.stop()
            except Exception:
                pass
        self._highlights.clear()
        self._references.clear()
