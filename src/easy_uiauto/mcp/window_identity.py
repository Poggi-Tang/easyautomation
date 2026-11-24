"""Separate native captions, recorded UIA selectors and live observation handles."""

from __future__ import annotations

import uiautomation

from . import native_windows
from .results import ActionError


def describe(window, hint: str = "") -> dict:
    handle = int(getattr(window, "NativeWindowHandle", 0) or 0)
    matches = native_windows.matching_windows("", handle=handle) if handle else []
    native = matches[0] if len(matches) == 1 else {}
    return dict(title=native.get("title", hint), class_name=native.get("class_name", ""),
                uia_name=str(getattr(window, "Name", "") or ""),
                uia_class=str(getattr(window, "ClassName", "") or ""),
                automation_id=str(getattr(window, "AutomationId", "") or ""))


def session(window, hint: str = "") -> dict:
    identity = describe(window, hint)
    return dict(window_name=identity["title"] or identity["uia_name"], window_identity=identity,
                target_handle=int(getattr(window, "NativeWindowHandle", 0) or 0),
                target_process_id=int(getattr(window, "ProcessId", 0) or 0),
                target_native_class=identity["class_name"])


def resolve(target):
    """A recorded selector cannot be overridden by a coincidentally matching caption."""
    from ..utils import _native_window_scope

    if isinstance(target, dict):
        scope = target.get("window_scope", {})
        owner = target.get("owner_window_identity")
        if scope.get("kind") == "owned" and owner:
            from ..library_actions import owned_windows
            from .scanner import _matches_step
            main = _native_window_scope(owner.get("title", ""), {
                key: value for key, value in dict(Name=owner.get("uia_name"),
                    ClassName=owner.get("uia_class"), AutomationId=owner.get("automation_id")).items() if value})
            expected = dict(handle=int(main.NativeWindowHandle), process_id=int(main.ProcessId))
            candidates = [w for w in owned_windows(expected) if w["class_name"] == scope.get("class_name")
                          and (not scope.get("title") or w["title"] == scope["title"])]
            if not candidates:
                raise ActionError("control_not_found", "The learned popup is not open", retryable=True)
            path = target.get("location", {}).get("Xpath", [])
            bound = [uiautomation.ControlFromHandle(w["handle"]) for w in candidates]
            matches = [w for w in bound if w is not None and path and _matches_step(w, path[0])]
            if len(matches) != 1:
                raise ActionError("ambiguous_match" if len(matches)>1 else "window_invalid",
                                  "The owned popup does not uniquely match its recorded UIA identity")
            return matches[0]
        handle = int(target.get("target_handle", 0) or 0)
        if handle:
            matches = native_windows.matching_windows("", handle=handle)
            if not matches or not matches[0]["visible"]:
                raise ActionError("window_invalid", "Original observation window closed or became hidden",
                                  reason="session_window_missing")
            native = matches[0]
            if (target.get("target_process_id") and native.get("process_id") != target["target_process_id"]
                    or target.get("target_native_class") and native.get("class_name") != target["target_native_class"]):
                raise ActionError("window_invalid", "Original observation window was replaced",
                                  reason="session_window_replaced")
            control = uiautomation.ControlFromHandle(handle)
            if control is None:
                raise ActionError("action_failed", "Could not bind the original observation window")
            return control
        location = target.get("location") or {}
        xpath = location.get("Xpath") or []
        first = xpath[0] if xpath and xpath[0].get("ControlType") in {
            "WindowControl", "MenuControl", "PaneControl", "ListControl"} else {}
        name = location.get("WindowName") or first.get("Name")
        if first:
            # A native caption narrows candidates without replacing the recorded UIA Name.
            hint = target.get("app_name") if first.get("Name") else name
            return _native_window_scope(hint or name or "", first)
        target = target.get("window_name") or target.get("app_name") or name or ""
    title = str(target).strip()
    if not title:
        raise ValueError("Window title must not be empty")
    candidates = [entry for entry in native_windows.matching_windows(title, partial=True) if entry["visible"]]
    if not candidates:
        # Compatibility for old visual-only records containing the provider's Name.
        return _native_window_scope(title, {})
    if len(candidates) != 1:
        captions = [dict(title=w.get("title", ""), handle=w.get("handle", 0), class_name=w.get("class_name", ""),
                         process_id=w.get("process_id", 0)) for w in candidates[:10]]
        raise ActionError("ambiguous_match", f"Multiple visible windows match: {title}; candidates={captions}")
    control = uiautomation.ControlFromHandle(candidates[0]["handle"])
    if control is None:
        raise ActionError("action_failed", f"Could not bind UI Automation to window: {title}")
    return control
